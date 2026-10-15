/**
 * 主进程侧的 `.pdf` 出入口（`docs/office-parity-spec.md` 阶段 7 的第一扇门）。
 *
 * 与 docx / xlsx / pptx 同一条纪律：内容流原文只留在主进程，渲染层拿到的是 `PdfReadModel`
 * （页号 + 定位键 + 解出来的文字），要改也只送**定位键 + 新文字**，改写由引擎在主进程现做
 * （`AGENTS.md` §3.3：界面只做装配）。
 *
 * 与那三种最大的不同是**落地方式**：`.pdf` 不许全量重写（那会把 163 205 字节的原文件重排成
 * 790 601 字节，等于"改一个字动整份文件"，spec D8），所以写盘写的是引擎给出的
 * "原文件逐字节 + 追加的 delta"（ISO 32000-1 `7.5.6` 增量更新），原子替换走 `fsutil.writeFileAtomically`（§2.4），
 * 原地写回先留一份 `.bak`（§8 备份语义）。
 */
import { dialog, type BrowserWindow, type IpcMain } from "electron";
import fs from "node:fs/promises";
import {
  applyPdfTextEdits,
  readPdfModel,
  revertPdfTextEdits,
  type PdfAppliedEdit,
  type PdfEditFailure,
  type PdfReadModel,
  type PdfTextEdit,
} from "@office-copilot/pdf-io";
import { fileExists, writeFileAtomically } from "./fsutil.js";

/** 渲染层据此决定走 pdf 那条路。 */
export const PDF_EXT = /\.pdf$/i;

/** 打开一份 `.pdf` 的回包。 */
export type PdfOpenResult =
  | { ok: true; path: string; document: PdfReadModel }
  | { ok: false; canceled: true }
  | { ok: false; message: string };

/** 一次改写或撤销的回包：引擎的拒绝是**结构化错误码**，中文文案由渲染层组装（§5.7）。 */
export type PdfEditResult =
  | {
      ok: true;
      path: string;
      /** 这次真正落地的改动（含旧文字），撤销要原样送回 `pdf:revert`（§2.5）。 */
      applied: PdfAppliedEdit[];
      /** 追加的那段 delta 的字节数；原文件前缀一字未动。 */
      deltaBytes: number;
      document: PdfReadModel;
    }
  | { ok: false; failures: PdfEditFailure[] }
  | { ok: false; canceled: true }
  | { ok: false; message: string };

/** 一次改写请求。 */
export interface PdfEditPayload {
  path: string;
  edits: PdfTextEdit[];
}

/** 一次撤销请求：`applied` 取自那次成功的回包，顺序原样送回即可。 */
export interface PdfRevertPayload {
  path: string;
  applied: PdfAppliedEdit[];
}

/**
 * 注册 pdf 相关的通道。
 *
 * @param ipc 主进程的 `ipcMain`。
 * @param getWindow 当前窗口，另存为对话框要挂在它上面。
 */
export function registerPdfHandlers(ipc: IpcMain, getWindow: () => BrowserWindow | null): void {
  // 「打开」由 fileHandlers 的单一对话框分派（docx / xlsx / pptx / pdf 都走那里），这里不再开第二个。
  ipc.handle("pdf:read", async (_e, filePath: string): Promise<PdfOpenResult> => openPdfDoc(filePath));

  ipc.handle("pdf:edit", async (_e, payload: PdfEditPayload): Promise<PdfEditResult> => {
    if (!payload.path) return { ok: false, message: "没有目标路径。" };
    return editPdfDoc(payload.path, payload.path, payload.edits);
  });

  ipc.handle("pdf:revert", async (_e, payload: PdfRevertPayload): Promise<PdfEditResult> =>
    revertPdfDoc(payload.path, payload.path, payload.applied),
  );

  ipc.handle("pdf:save-as", async (_e, payload: PdfEditPayload): Promise<PdfEditResult> => {
    const result = await dialog.showSaveDialog(getWindow()!, {
      title: "另存为 PDF",
      defaultPath: payload.path ? payload.path.split(/[\\/]/).pop() : "文档.pdf",
      filters: [{ name: "PDF", extensions: ["pdf"] }],
    });
    if (result.canceled || !result.filePath) return { ok: false, canceled: true };
    // 增量更新必须对着**打开时的那份**追加，所以另存为以原文件为底、写到新路径。
    return editPdfDoc(result.filePath, payload.path, payload.edits);
  });
}

/**
 * 打开一份 `.pdf` 并压成跨界载荷。
 *
 * @param filePath 磁盘路径。
 * @returns ok:true + 读模型（页数与全部文字 run）；不是 PDF 或读不动时 ok:false + 可读消息。
 */
export async function openPdfDoc(filePath: string): Promise<PdfOpenResult> {
  try {
    const bytes = await fs.readFile(filePath);
    return { ok: true, path: filePath, document: await readPdfModel(bytes) };
  } catch (err) {
    return { ok: false, message: `打开失败：${err instanceof Error ? err.message : String(err)}` };
  }
}

/**
 * 把一批文字改写落到磁盘：重读底包 → 引擎整批现做 → 原子替换。
 *
 * @param targetPath 要写的路径。
 * @param basePath 增量更新的底包路径（原地保存时与 `targetPath` 相同；另存为时是原文件）。
 * @param edits 渲染层带回的改动，定位键取自上一次读模型。
 * @returns ok:true + 新字节写盘后的读模型；任一条改不动时**整批不写**并原样带回错误码（§2.4）。
 */
export async function editPdfDoc(targetPath: string, basePath: string, edits: PdfTextEdit[]): Promise<PdfEditResult> {
  return await applyBatch(targetPath, basePath, "保存", async (bytes) => await applyPdfTextEdits(bytes, edits));
}

/**
 * 撤销一批已落地的改写：把每条的旧文字换回去，落地流程与 {@link editPdfDoc} 共用同一段（§3.2）。
 *
 * @param targetPath 要写的路径。
 * @param basePath 底包路径（撤销总是对着磁盘上当前那份重做，所以两者相同）。
 * @param applied 之前那批改动回包里的 `applied`。
 * @returns 与 {@link editPdfDoc} 同形；那条 run 后来又被改过时引擎给 `text-drifted`，整批不写（§2.5）。
 */
export async function revertPdfDoc(targetPath: string, basePath: string, applied: PdfAppliedEdit[]): Promise<PdfEditResult> {
  return await applyBatch(targetPath, basePath, "撤销", async (bytes) => await revertPdfTextEdits(bytes, applied));
}

/**
 * 一次改写的公共落地流程：读底包 → 交给引擎整批做 → 原子替换 → 重新读一遍模型。
 *
 * @param targetPath 要写的路径。
 * @param basePath 底包路径。
 * @param verb 回给人听的那个动作名（"保存" / "撤销"），失败文案要跟着当时的动作走（§5.7）。
 * @param run 在这份底包字节上完成本次改写，成功时给出"原字节 + delta"的完整结果，失败时给出错误码清单。
 * @returns ok:true + 落盘路径与刷新后的读模型；底包不存在或任一条改不动时整批不写。
 */
async function applyBatch(
  targetPath: string,
  basePath: string,
  verb: string,
  run: (bytes: Uint8Array) => Promise<{ ok: true; bytes: Uint8Array; deltaBytes: number; applied: PdfAppliedEdit[] } | { ok: false; failures: PdfEditFailure[] }>,
): Promise<PdfEditResult> {
  try {
    if (!(await fileExists(basePath))) {
      return { ok: false, message: `${verb}已取消：找不到要改的那份文件（新建 PDF 还不支持，请先打开一份 .pdf）。` };
    }
    const baseBytes = await fs.readFile(basePath);
    const made = await run(new Uint8Array(baseBytes));
    if (!made.ok) return made;
    if (await fileExists(targetPath)) {
      await fs.copyFile(targetPath, `${targetPath}.bak`);
    }
    await writeFileAtomically(targetPath, made.bytes);
    // 读模型是打开时算好的，改过之后要重新解析一次才拿得到新文字（渲染层据此刷新列表）。
    return { ok: true, path: targetPath, applied: made.applied, deltaBytes: made.deltaBytes, document: await readPdfModel(made.bytes) };
  } catch (err) {
    return { ok: false, message: `${verb}失败：${err instanceof Error ? err.message : String(err)}` };
  }
}
