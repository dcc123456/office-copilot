/**
 * 主进程侧的 `.xlsx` 出入口（spec 阶段 2 · 提交 6）。
 *
 * 与 docx 那条路同一个口径：`sourceXml` 只在主进程这一侧存在，渲染层拿到的是
 * `WorkbookPayload`（结构化格子），保存时**重新从磁盘打开**、按 `partPath` 配对写回，
 * 再走 `写 .tmp → rename`（`AGENTS.md` §2.4）。原地写回先留一份 `.bak`（§8 备份语义）。
 */
import { dialog, type BrowserWindow, type IpcMain } from "electron";
import fs from "node:fs/promises";
import path from "node:path";
import {
  SheetFile,
  applySheetPayload,
  auditXlsxPackage,
  workbookPayload,
  type CompatIssue,
  type ResolvedChart,
  type SheetPayload,
  type WorkbookPayload,
} from "@office-copilot/sheet-io";
import { fileExists } from "./fsutil.js";

/** 渲染层据此决定走表格那条路；`.xlsm` 由 `SheetFile` 明确拒绝，不在可打开的扩展名里。 */
export const SHEET_EXT = /\.xlsx$/i;

export type SheetOpenResult =
  | { ok: true; path: string; workbook: WorkbookPayload }
  | { ok: false; canceled: true }
  | { ok: false; message: string };

export type SheetSaveResult =
  | { ok: true; path: string }
  | { ok: false; canceled: true }
  | { ok: false; message: string };

/** 一张待插入的图表：落在哪张表（按 `partPath` 配对）+ 引擎已解析好的图表数据。 */
export interface SheetChartRequest {
  partPath: string;
  chart: ResolvedChart;
}

/** `addChart` 的错误码 → 给人看的那半句原因；码由引擎给，文案在这里组装（`AGENTS.md` §5.7）。 */
const CHART_ADD_FAILURE: Record<"sheet-missing" | "drawing-exists" | "no-worksheet-root", string> = {
  "sheet-missing": "这张表在文件里已经找不到了",
  "drawing-exists": "它已经有一个图表或图片，一张表只能挂一个",
  "no-worksheet-root": "表部件的结构读不懂",
};

/** 渲染层送回来的保存请求：一张表只带跨界的那部分，原文由主进程自己重读。 */
export interface SheetSavePayload {
  /** 当前文档路径；null 表示还没落过盘。 */
  path: string | null;
  /** 所有表的当前内容，顺序与打开时一致。 */
  sheets: SheetPayload[];
  /** 本次要新插入的原生图表；已写进文件的那些不在这里（它们随底包原样保留）。 */
  charts?: SheetChartRequest[];
  /** 另存为时对话框用的缺省文件名。 */
  defaultName?: string;
}

/**
 * 注册表格相关的通道。
 *
 * @param ipc 主进程的 `ipcMain`。
 * @param getWindow 当前窗口，另存为对话框要挂在它上面。
 */
export function registerSheetHandlers(ipc: IpcMain, getWindow: () => BrowserWindow | null): void {
  // 「打开」由 fileHandlers 的单一对话框分派（docx / xlsx 都走那里），这里不再开第二个。
  ipc.handle("sheet:read", async (_e, filePath: string): Promise<SheetOpenResult> => openSheetDoc(filePath));

  ipc.handle("sheet:save", async (_e, payload: SheetSavePayload): Promise<SheetSaveResult> => {
    if (!payload.path) return { ok: false, message: "没有目标路径。" };
    return saveSheetDoc(payload.path, payload.path, payload.sheets, payload.charts);
  });

  ipc.handle("sheet:save-as", async (_e, payload: SheetSavePayload): Promise<SheetSaveResult> => {
    const win = getWindow();
    const result = await dialog.showSaveDialog(win!, {
      title: "另存为 Excel 表格",
      defaultPath: payload.defaultName ?? "工作簿.xlsx",
      filters: [{ name: "Excel 表格", extensions: ["xlsx"] }],
    });
    if (result.canceled || !result.filePath) return { ok: false, canceled: true };
    // 未建模部件（图表、透视表、外部链接…）只存在于打开的那份文件里，
    // 所以另存为必须以它为底，否则存出去的包会少掉用户没让我们碰的东西。
    const base = payload.path ?? result.filePath;
    return saveSheetDoc(result.filePath, base, payload.sheets, payload.charts);
  });
}

/**
 * 打开一份 `.xlsx` 并压成跨界载荷。
 *
 * @param filePath 磁盘路径。
 * @returns ok:true + 工作簿载荷（含 warnings 与未建模部件清单）；读不到或含宏时 ok:false + 可读消息。
 */
export async function openSheetDoc(filePath: string): Promise<SheetOpenResult> {
  try {
    const bytes = await fs.readFile(filePath);
    const file = await SheetFile.fromBytes(bytes, { fileName: path.basename(filePath) });
    const workbook = await file.toWorkbook();
    return {
      ok: true,
      path: filePath,
      workbook: workbookPayload(workbook.fileName, workbook.sheets, workbook.warnings, workbook.unsupported),
    };
  } catch (err) {
    return { ok: false, message: `打开失败：${err instanceof Error ? err.message : String(err)}` };
  }
}

/**
 * 把渲染层的内容写回磁盘：重读底包 → 按 `partPath` 逐表落地 → 插图表 → 原子替换。
 *
 * @param targetPath 要写的路径。
 * @param basePath 未建模部件的来源路径（通常是当前文档；另存为时是原文件）。
 * @param sheets 渲染层带回的表内容。
 * @param charts 本次要新插入的原生图表（AI 出图那条，spec 决策 D5）；省略即不碰图表部件。
 * @returns ok:true + 落盘路径；表配不上、图表插不进（那张表已有画图部件）、写出会新引入结构问题，
 *   都整批不写并回一句中文（§2.4 的事务语义）。
 */
export async function saveSheetDoc(
  targetPath: string,
  basePath: string,
  sheets: SheetPayload[],
  charts?: SheetChartRequest[],
): Promise<SheetSaveResult> {
  try {
    const baseExists = await fileExists(basePath);
    if (!baseExists) {
      return { ok: false, message: "新建表格文件还不支持：请先打开一份 .xlsx 再另存为。" };
    }
    const baseBytes = await fs.readFile(basePath);
    const file = await SheetFile.fromBytes(baseBytes, { fileName: path.basename(basePath) });
    for (const payload of sheets) {
      const target = await sheetForPartPath(file, payload.partPath);
      if (!target) return { ok: false, message: `表 ${payload.name} 在文件里找不到了，本次保存已取消。` };
      applySheetPayload(target, payload);
      file.writeSheet(target);
    }
    for (const request of charts ?? []) {
      const added = await file.addChart(request.partPath, request.chart);
      if (!added.ok) {
        return { ok: false, message: `图表插不进这张表（${CHART_ADD_FAILURE[added.code]}），本次保存已取消。` };
      }
    }
    const out = await file.toBytes();
    // 保存前过一次结构审计：写坏了用户是打不开的，而这时磁盘上的文件还没动（§2.4 的整批放弃）。
    const outIssues = await auditQuietly(out);
    // 只在"输出有问题"时才去审底包——正常保存这条路只多一次解压，不翻倍。
    const introduced = outIssues.length === 0 ? [] : newlyIntroducedIssues(await auditQuietly(baseBytes), outIssues);
    if (introduced.length > 0) {
      return {
        ok: false,
        message: `保存已取消：这次改动会新引入 ${introduced.length} 处让办公软件判损坏的结构问题（${[...new Set(introduced.map((issue) => issue.part))].join("、")}），磁盘上的文件一个字都没动。`,
      };
    }
    if (await fileExists(targetPath)) {
      await fs.copyFile(targetPath, `${targetPath}.bak`);
    }
    const tmp = `${targetPath}.tmp`;
    await fs.writeFile(tmp, out);
    await fs.rename(tmp, targetPath);
    return { ok: true, path: targetPath };
  } catch (err) {
    return { ok: false, message: `保存失败：${err instanceof Error ? err.message : String(err)}` };
  }
}

/**
 * 出字节后过一次结构审计，**审计自己读不懂就回空**。
 *
 * @param bytes 准备写盘（或底包里原有）的 `.xlsx` 字节。
 * @returns 发现的问题；`@office-copilot/docx-io` 的 XML 解析器认不出的陌生部件会整包判读失败，
 *   那时按"判不了"处理而不是"坏了"——判不出不许把用户正在编辑的东西挡在保存门外。
 */
async function auditQuietly(bytes: Uint8Array): Promise<CompatIssue[]> {
  try {
    return await auditXlsxPackage(bytes);
  } catch {
    return [];
  }
}

/**
 * 只留"这次保存新引入的"结构问题。
 *
 * @param baseIssues 底包（磁盘上原来那份）审计出来的问题。
 * @param outIssues 准备写出的字节审计出来的问题。
 * @returns 底包里没有的那些。底包本来就不合规（第三方导出器留下的）不该让用户没法保存，
 *   这里只对**本次事务**负责（`AGENTS.md` §2.4）。
 */
export function newlyIntroducedIssues(baseIssues: CompatIssue[], outIssues: CompatIssue[]): CompatIssue[] {
  const seen = new Set(baseIssues.map(issueKey));
  return outIssues.filter((issue) => !seen.has(issueKey(issue)));
}

/** 一条问题的身份键：码 + 部件 + 细节，同一个包审两遍会得到同一个键。 */
function issueKey(issue: CompatIssue): string {
  return `${issue.code}@${issue.part}@${issue.detail}`;
}

/**
 * 按部件路径取回主进程这次打开的那张表。
 *
 * @param file 刚从磁盘解析出来的工作簿。
 * @param partPath 载荷里带的部件路径，如 `xl/worksheets/sheet1.xml`。
 * @returns 表模型（带 `sourceXml`）；包里已经没有这个部件时返回 undefined。
 */
async function sheetForPartPath(file: SheetFile, partPath: string) {
  for (let i = 0; i < file.sheetCount; i += 1) {
    if (file.sheetRefAt(i)?.partPath !== partPath) continue;
    return await file.sheetAt(i);
  }
  return undefined;
}
