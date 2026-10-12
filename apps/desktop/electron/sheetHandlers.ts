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
  workbookPayload,
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

/** 渲染层送回来的保存请求：一张表只带跨界的那部分，原文由主进程自己重读。 */
export interface SheetSavePayload {
  /** 当前文档路径；null 表示还没落过盘。 */
  path: string | null;
  /** 所有表的当前内容，顺序与打开时一致。 */
  sheets: SheetPayload[];
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
    return saveSheetDoc(payload.path, payload.path, payload.sheets);
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
    return saveSheetDoc(result.filePath, base, payload.sheets);
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
 * 把渲染层的内容写回磁盘：重读底包 → 按 `partPath` 逐表落地 → 原子替换。
 *
 * @param targetPath 要写的路径。
 * @param basePath 未建模部件的来源路径（通常是当前文档；另存为时是原文件）。
 * @param sheets 渲染层带回的表内容。
 * @returns ok:true + 落盘路径；任何一张表配不上就整批不写（§2.4 的事务语义）。
 */
export async function saveSheetDoc(
  targetPath: string,
  basePath: string,
  sheets: SheetPayload[],
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
    const out = await file.toBytes();
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
