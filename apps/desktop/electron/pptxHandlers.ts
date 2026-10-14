/**
 * 主进程侧的 `.pptx` 出入口（`docs/office-parity-spec.md` 阶段 3 的第一扇门）。
 *
 * 与 docx / xlsx 同一条纪律：幻灯片 XML 原文只留在主进程，渲染层拿到的是 `PresentationModel`
 * （结构化形状与文字），要改也只送**定位坐标**（部件路径 + 形状号 + 段落/run 序号），
 * op 由引擎在主进程现做（`AGENTS.md` §3.3：界面只做装配），落地走 `写 .tmp → rename`（§2.4），
 * 原地写回先留一份 `.bak`（§8 备份语义）。
 */
import { dialog, type BrowserWindow, type IpcMain } from "electron";
import fs from "node:fs/promises";
import path from "node:path";
import {
  PptxFile,
  applySlideEdit,
  auditPptxPackage,
  revertSlideEdits,
  setRunColor,
  setRunFont,
  setShapeGeometry,
  setShapeRunText,
  type PptxAuditIssue,
  type PresentationModel,
  type PptxEditPayload,
  type PptxEditRequest,
  type PptxRevertPayload,
  type SlideEditOp,
} from "@office-copilot/pptx-io";
import { newlyIntroducedIssues } from "@office-copilot/opc";
import { fileExists, writeFileAtomically } from "./fsutil.js";

/** 渲染层据此决定走幻灯片那条路；老的 `.ppt` 二进制格式不在可打开的扩展名里。 */
export const PPTX_EXT = /\.pptx$/i;

export type PptxOpenResult =
  | { ok: true; path: string; presentation: PresentationModel }
  | { ok: false; canceled: true }
  | { ok: false; message: string };

export type PptxEditResult =
  | {
      ok: true;
      path: string;
      /** 本次真正动过的幻灯片部件（去重排序），验收与"改了哪几页"的提示用它。 */
      changedParts: string[];
      /**
       * 这批改动的 op（含前/后像）。撤销必须由渲染层把它们原样送回 `ppt:revert`——
       * op 只在主进程现做，渲染层拿不到，所以成功时不送回来就等于没有撤销（§2.5）。
       */
      ops: SlideEditOp[];
      /**
       * 撤销时没落地的条目理由（正向保存不会有：那批是一条不落地就整批放弃）。
       * 非空代表"部分撤销"，渲染层据此回用户 x/y 而不是假装全回去了（§2.5）。
       */
      skipped?: string[];
      presentation: PresentationModel;
    }
  | { ok: false; canceled: true }
  | { ok: false; message: string };

/**
 * 注册幻灯片相关的通道。
 *
 * @param ipc 主进程的 `ipcMain`。
 * @param getWindow 当前窗口，另存为对话框要挂在它上面。
 */
export function registerPptxHandlers(ipc: IpcMain, getWindow: () => BrowserWindow | null): void {
  // 「打开」由 fileHandlers 的单一对话框分派（docx / xlsx / pptx 都走那里），这里不再开第二个。
  ipc.handle("ppt:read", async (_e, filePath: string): Promise<PptxOpenResult> => openPptxDoc(filePath));

  ipc.handle("ppt:edit", async (_e, payload: PptxEditPayload): Promise<PptxEditResult> => {
    if (!payload.path) return { ok: false, message: "没有目标路径。" };
    return editPptxDoc(payload.path, payload.path, payload.edits);
  });

  ipc.handle("ppt:revert", async (_e, payload: PptxRevertPayload): Promise<PptxEditResult> =>
    revertPptxDoc(payload.path, payload.path, payload.ops),
  );

  ipc.handle("ppt:save-as", async (_e, payload: PptxEditPayload): Promise<PptxEditResult> => {
    const result = await dialog.showSaveDialog(getWindow()!, {
      title: "另存为 PowerPoint 演示文稿",
      defaultPath: payload.defaultName ?? "演示文稿.pptx",
      filters: [{ name: "PowerPoint 演示文稿", extensions: ["pptx"] }],
    });
    if (result.canceled || !result.filePath) return { ok: false, canceled: true };
    // 未建模部件（主题、母版、备注、图片…）只存在于打开的那份文件里，
    // 所以另存为必须以它为底，否则存出去的包会少掉用户没让我们碰的东西。
    const base = payload.path ?? result.filePath;
    return editPptxDoc(result.filePath, base, payload.edits);
  });
}

/**
 * 打开一份 `.pptx` 并压成跨界载荷。
 *
 * @param filePath 磁盘路径。
 * @returns ok:true + 演示文稿模型（幻灯片顺序、EMU 页面尺寸、形状树）；读不到时 ok:false + 可读消息。
 */
export async function openPptxDoc(filePath: string): Promise<PptxOpenResult> {
  try {
    const bytes = await fs.readFile(filePath);
    const file = await PptxFile.open(bytes);
    return { ok: true, path: filePath, presentation: file.model };
  } catch (err) {
    return { ok: false, message: `打开失败：${err instanceof Error ? err.message : String(err)}` };
  }
}

/**
 * 把一批改动落到磁盘：重读底包 → 逐条现做 op 并落地 → 出字节 → 过审计 → 原子替换。
 *
 * @param targetPath 要写的路径。
 * @param basePath 未建模部件的来源路径（通常是当前文档；另存为时是原文件）。
 * @param edits 渲染层带回的改动，**按数组顺序**应用，后面的那条看得到前面已落地的结果。
 * @returns ok:true + 落盘路径与本次改动的部件清单和刷新后的模型；任一条定位失败、或写出会新引入
 *   结构问题，都整批不写并回一句中文（§2.4 的事务语义）。
 */
export async function editPptxDoc(
  targetPath: string,
  basePath: string,
  edits: PptxEditRequest[],
): Promise<PptxEditResult> {
  return await applyBatch(targetPath, basePath, "保存", async (file) => {
    const ops: SlideEditOp[] = [];
    for (const edit of edits) {
      const made = await makeEdit(file, edit);
      if (!made.ok) return made;
      applySlideEdit(file, made.op);
      ops.push(made.op);
    }
    return { ok: true, ops };
  });
}

/**
 * 一批幻灯片改动的公共落地流程：重读底包 → 交给 `run` 现做并应用 op → 出字节 → 过审计 → 原子替换。
 *
 * 正向编辑与撤销共用这一段（`AGENTS.md` §3.2：同一条事务流程不许长两份），差别只在 `run` 里
 * 写的是后像还是前像。
 *
 * @param targetPath 要写的路径。
 * @param basePath 底包路径（未建模部件的来源；原地保存时与 `targetPath` 相同）。
 * @param verb 回给人听的那个动作名（"保存" / "撤销"），失败文案要跟着当时的动作走（§5.7）。
 * @param run 在一个刚打开的读写器上完成本次改动，返回这批改动的 op（失败时给引擎的理由）；
 *   撤销那条路还会带回没落地的条目理由（`skipped`），正向保存没有这一项。
 * @returns ok:true + 落盘路径、本次真正动过的部件清单（去重）与刷新后的模型；
 *   `run` 失败、底包不存在、或写出会新引入结构问题时整批不写并回一句中文（§2.4）。
 */
async function applyBatch(
  targetPath: string,
  basePath: string,
  verb: string,
  run: (
    file: PptxFile,
  ) => Promise<{ ok: true; ops: SlideEditOp[]; skipped?: string[] } | { ok: false; reason: string }>,
): Promise<PptxEditResult> {
  try {
    if (!(await fileExists(basePath))) {
      return { ok: false, message: `${verb}已取消：找不到要改的那份文件（新建演示文稿还不支持，请先打开一份 .pptx）。` };
    }
    const baseBytes = await fs.readFile(basePath);
    const file = await PptxFile.open(baseBytes);
    const made = await run(file);
    if (!made.ok) return { ok: false, message: `${verb}已取消（磁盘上的文件一个字都没动）：${made.reason}` };
    const out = await file.toBytes();
    // 写盘前过一次结构审计：写坏了用户是打不开的，而这时磁盘上的文件还没动（§2.4 的整批放弃）。
    const outIssues = await auditQuietly(out);
    // 只在"输出有问题"时才去审底包——正常保存这条路只多一次解压，不翻倍。
    const introduced = outIssues.length === 0 ? [] : newlyIntroducedIssues(await auditQuietly(baseBytes), outIssues);
    if (introduced.length > 0) {
      return {
        ok: false,
        message: `${verb}已取消：这次改动会新引入 ${introduced.length} 处让办公软件判损坏的结构问题（${[...new Set(introduced.map((issue) => issue.part))].join("、")}），磁盘上的文件一个字都没动。`,
      };
    }
    if (await fileExists(targetPath)) {
      await fs.copyFile(targetPath, `${targetPath}.bak`);
    }
    await writeFileAtomically(targetPath, out);
    // `PptxFile.model` 是打开时算好的，改过之后要重新解析一次才拿得到新文字（渲染层据此刷新幻灯片）。
    const refreshed = await PptxFile.open(out);
    return {
      ok: true,
      path: targetPath,
      changedParts: [...new Set(made.ops.map((op) => op.partPath))].sort(),
      ops: made.ops,
      ...(made.skipped?.length ? { skipped: made.skipped } : {}),
      presentation: refreshed.model,
    };
  } catch (err) {
    return { ok: false, message: `${verb}失败：${err instanceof Error ? err.message : String(err)}` };
  }
}

/**
 * 按 `kind` 把一条跨界请求翻成引擎的 op 生产者调用。
 *
 * @param file 已打开的读写器。
 * @param edit 渲染层送来的一条改动。
 * @returns `{ ok, op }`；引擎给的理由原样带回，由调用方组装中文（§5.7）。
 */
async function makeEdit(file: PptxFile, edit: PptxEditRequest): Promise<{ ok: true; op: SlideEditOp } | { ok: false; reason: string }> {
  switch (edit.kind) {
    case "text":
      return await setShapeRunText(file, edit.target, edit.text);
    case "font":
      return await setRunFont(file, edit.target, edit.format);
    case "color":
      return await setRunColor(file, edit.target, edit.hex);
    case "geometry":
      return await setShapeGeometry(file, edit.target, edit.box);
  }
}

/**
 * 撤销一批已落地的改动：把每条 op 的前像写回部件，其余同 {@link editPptxDoc}。
 *
 * @param targetPath 要写的路径。
 * @param basePath 底包路径（撤销总是对着磁盘上当前那份重做，所以两者相同）。
 * @param ops 之前那批改动落地的 op，**按逆序**传入即可撤销整批。
 * @returns 与 {@link editPptxDoc} 同形，另带没落地的条目理由（`skipped`）；一条都落不了时整批不写，
 *   op 指向的部件已不在包里或那页后来又被改过都算落不了（§2.5 不许牵连用户后来的编辑）。
 */
export async function revertPptxDoc(targetPath: string, basePath: string, ops: SlideEditOp[]): Promise<PptxEditResult> {
  return await applyBatch(targetPath, basePath, "撤销", async (file) => {
    const { applied, skipped } = await revertSlideEdits(file, ops);
    if (!applied.length) return { ok: false, reason: skipped.map((item) => item.reason).join("；") };
    return { ok: true, ops: applied, skipped: skipped.map((item) => item.reason) };
  });
}

/**
 * 出字节后过一次结构审计，**审计自己读不懂就回空**。
 *
 * @param bytes 准备写盘（或底包里原有）的 `.pptx` 字节。
 * @returns 发现的问题；解析器认不出的陌生部件会整包判读失败，那时按"判不了"处理而不是"坏了"——
 *   判不出不许把用户正在编辑的东西挡在保存门外（与 `sheetHandlers.auditQuietly` 同口径）。
 */
async function auditQuietly(bytes: Uint8Array): Promise<PptxAuditIssue[]> {
  try {
    return await auditPptxPackage(bytes);
  } catch {
    return [];
  }
}
