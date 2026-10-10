import { dialog, type IpcMain, type BrowserWindow } from "electron";
import fs from "node:fs/promises";
import {
  DocxFile,
  type FootnoteBodies,
  type FootnoteKind,
  type FootnoteStore,
  type LinkStyle,
  type PaneStyle,
  type ResolvedStyle,
  type SectionGeometry,
  type StyleEdit,
} from "@office-copilot/docx-io";
import type { PMDoc } from "@office-copilot/docx-io/src/pm-schema.js";

/** 两种部件的固定遍历顺序：脚注在前，与 Word 窗格里的排布一致。 */
const FOOTNOTE_KINDS: readonly FootnoteKind[] = ["footnote", "endnote"];

/** Default geometry for a brand-new document (A4, 1-inch margins). */
export const DEFAULT_SECTION: SectionGeometry = {
  pageWidth: 793.7,
  pageHeight: 1122.5,
  marginTop: 96,
  marginRight: 96,
  marginBottom: 96,
  marginLeft: 96,
  header: 48,
  footer: 48,
  contentWidth: 601.7,
  contentHeight: 930.5,
};

/** Plain-JS shape of a header/footer payload carried over IPC. */
export interface HeaderFooterPayload {
  header?: PMDoc | null;
  footer?: PMDoc | null;
  section?: SectionGeometry | null;
  link?: LinkStyle | null;
  /**
   * 样式窗格里"新建样式 / 更新样式以匹配所选内容"攒下的编辑，保存时一次性落进
   * styles.xml。空数组与不带这个字段都表示"这次保存不动样式表"。
   */
  styles?: StyleEdit[];
  /**
   * 脚注窗格里的正文。只带 id 与 PM 正文：部件的原始 XML、`separator` 那两条版式件
   * 与未建模属性都不跨界，渲染层改得动正文，也不该看见其余那些。
   */
  footnotes?: FootnoteBodies;
}

/** 渲染层经 preload 送进来的保存请求：正文 + 需要落地的各项文档级设置。 */
export type SavePayload = HeaderFooterPayload & {
  doc: PMDoc;
  path?: string | null;
  defaultName?: string;
};

/**
 * 样式窗格的预览：把待落地的样式编辑投影成窗格清单与样式链，**不碰磁盘**。
 *
 * 走的是保存时的同一个 `DocxFile.applyStyleEdits`，因此窗格里看到的名字、
 * 加粗与颜色，就是保存后 Word 里的那份格式；渲染层不再自己拼第二套样式解析。
 */
export type StylePreview =
  | { ok: true; styles: PaneStyle[]; styleFormats: ResolvedStyle[] }
  | { ok: false; message: string };

export type OpenResult =
  | {
      ok: true;
      path: string;
      doc: PMDoc;
      warnings: string[];
      header: PMDoc | null;
      footer: PMDoc | null;
      section: SectionGeometry;
      /**
       * 样式窗格的数据源：本文档 styles.xml 里声明的段落样式，按 uiPriority 排好。
       * 空数组 = 这份文档没有具名样式（新建文档就是这种），不是读取失败。
       */
      styles: PaneStyle[];
      /**
       * 样式链按 Word 的优先级合并后的可见格式，渲染层据此注入文档作用域的样式表。
       * 空数组 = 这份文档没有 styles.xml（新建文档就是这种），渲染层因此什么都不加。
       */
      styleFormats: ResolvedStyle[];
      /**
       * 脚注与尾注的正文，按部件分组、按部件里的条目顺序排。
       * 空对象 = 这份文档没有脚注；正文里没有引用点的条目（版式件）不会出现在这里。
       */
      footnotes: FootnoteBodies;
    }
  | { ok: false; canceled: true }
  | { ok: false; message: string };

export function registerFileHandlers(
  ipc: IpcMain,
  getWindow: () => BrowserWindow | null,
): void {
  ipc.handle("file:new", async () => {
    return {
      ok: true as const,
      path: null,
      doc: {
        type: "doc",
        content: [
          {
            type: "heading",
            attrs: { level: 1 },
            content: [{ type: "text", text: "无标题" }],
          },
          { type: "paragraph" },
        ],
      } as PMDoc,
      header: null,
      footer: null,
      section: DEFAULT_SECTION,
      warnings: [],
      // 新建文档的空白模板里没有 styles.xml，样式窗格因此是空的。
      styles: [],
      styleFormats: [],
      // 空白模板也没有脚注部件。
      footnotes: {},
    };
  });

  ipc.handle("file:open", async (): Promise<OpenResult> => {
    const win = getWindow();
    const result = await dialog.showOpenDialog(win!, {
      title: "打开 Word 文档",
      properties: ["openFile"],
      filters: [
        { name: "Word 文档", extensions: ["docx"] },
        { name: "所有文件", extensions: ["*"] },
      ],
    });
    if (result.canceled || result.filePaths.length === 0) {
      return { ok: false, canceled: true };
    }
    const filePath = result.filePaths[0]!;
    return openDocx(filePath);
  });

  ipc.handle(
    "file:read-docx",
    async (_e, filePath: string): Promise<OpenResult> => openDocx(filePath),
  );

  ipc.handle(
    "file:save",
    async (
      _e,
      payload: SavePayload,
    ): Promise<{ ok: true; path: string } | { ok: false; message: string }> => {
      if (!payload.path) return { ok: false, message: "没有目标路径。" };
      return saveDocx(payload.path, payload.doc, payload);
    },
  );

  ipc.handle(
    "file:save-as",
    async (
      _e,
      payload: SavePayload,
    ): Promise<
      | { ok: true; path: string }
      | { ok: false; canceled: true }
      | { ok: false; message: string }
    > => {
      const win = getWindow();
      const result = await dialog.showSaveDialog(win!, {
        title: "另存为 Word 文档",
        defaultPath: payload.defaultName ?? "document.docx",
        filters: [{ name: "Word 文档", extensions: ["docx"] }],
      });
      if (result.canceled || !result.filePath) {
        return { ok: false, canceled: true };
      }
      const r = await saveDocx(result.filePath, payload.doc, payload);
      if (!r.ok) return r;
      return { ok: true, path: result.filePath };
    },
  );

  ipc.handle(
    "file:write-docx",
    async (_e, payload: SavePayload & { path: string }) =>
      saveDocx(payload.path, payload.doc, payload),
  );

  ipc.handle(
    "styles:preview",
    async (
      _e,
      payload: { path: string | null; styles: StyleEdit[] },
    ): Promise<StylePreview> => previewStyles(payload.path, payload.styles),
  );

  ipc.handle("file:pick-image", async () => {
    const win = getWindow();
    const result = await dialog.showOpenDialog(win!, {
      title: "选择图片",
      properties: ["openFile"],
      filters: [
        {
          name: "图片",
          extensions: ["png", "jpg", "jpeg", "gif", "webp"],
        },
      ],
    });
    if (result.canceled || result.filePaths.length === 0) {
      return { ok: false as const, canceled: true as const };
    }
    const p = result.filePaths[0]!;
    const buf = await fs.readFile(p);
    const ext = p.split(".").pop()?.toLowerCase() ?? "png";
    const mime =
      ext === "jpg" || ext === "jpeg"
        ? "image/jpeg"
        : ext === "gif"
          ? "image/gif"
          : ext === "webp"
            ? "image/webp"
            : "image/png";
    const dataUrl = `data:${mime};base64,${buf.toString("base64")}`;
    return { ok: true as const, path: p, dataUrl };
  });
}

async function openDocx(filePath: string): Promise<OpenResult> {
  try {
    const bytes = await fs.readFile(filePath);
    const docx = await DocxFile.fromBytes(bytes);
    const { doc, warnings, headers, section, styles, styleFormats, footnotes } =
      await docx.toProseMirror();
    return {
      ok: true,
      path: filePath,
      doc,
      warnings,
      header: headers?.entries.header?.pm ?? null,
      footer: headers?.entries.footer?.pm ?? null,
      section,
      styles,
      styleFormats,
      footnotes: footnoteBodiesOf(footnotes),
    };
  } catch (err) {
    return {
      ok: false,
      message: `打开失败：${err instanceof Error ? err.message : String(err)}`,
    };
  }
}

/**
 * 把脚注 store 收成跨界用的形状：只留正常条目的 id 与正文。
 *
 * `separator` / `continuationSeparator` 是脚注区的版式件，用户改不了也不该看见；
 * `unmodeled` 的部件由主进程整段透传，所以这里的条目缺了就缺了，写回时不会被覆盖。
 *
 * @param store `toProseMirror()` 带回来的脚注与尾注条目。
 * @returns 按部件分组的正文；某一种部件没有正常条目时这个键不出现。
 */
function footnoteBodiesOf(store: FootnoteStore): FootnoteBodies {
  const out: FootnoteBodies = {};
  for (const kind of FOOTNOTE_KINDS) {
    const bodies = store[kind]
      .filter((entry) => !entry.role)
      .map((entry) => ({ id: entry.id, doc: entry.pm }));
    if (bodies.length) out[kind] = bodies;
  }
  return out;
}

async function saveDocx(
  filePath: string,
  pmDoc: PMDoc,
  hf: HeaderFooterPayload,
): Promise<{ ok: true; path: string } | { ok: false; message: string }> {
  try {
    const docx = await loadForEdit(filePath);
    if (hf.header) docx.setHeaderFooter("header", hf.header);
    if (hf.footer) docx.setHeaderFooter("footer", hf.footer);
    if (hf.section) docx.setSectionGeometry(hf.section);
    if (hf.link) docx.setDocStyle({ link: hf.link });
    for (const kind of FOOTNOTE_KINDS) {
      const bodies = hf.footnotes?.[kind];
      if (bodies)
        docx.setFootnoteBodies(kind, bodies);
    }
    // 样式表要先落地：正文里的 w:pStyle 引用的是 styles.xml 声明的 id，
    // 换了新 id 却还没写进样式表，Word 会当成未声明样式退回默认格式。
    if (hf.styles?.length) await docx.applyStyleEdits(hf.styles);
    await docx.replaceProseMirror(pmDoc);
    const out = await docx.toBytes();
    const tmp = filePath + ".tmp";
    await fs.writeFile(tmp, out);
    await fs.rename(tmp, filePath);
    return { ok: true, path: filePath };
  } catch (err) {
    return {
      ok: false,
      message: `保存失败：${err instanceof Error ? err.message : String(err)}`,
    };
  }
}

/**
 * 样式窗格的预览：在内存里把待落地的编辑合并进样式表，取回投影结果，**不写文件**。
 *
 * 与 `saveDocx` 同一个 `applyStyleEdits`，所以"窗格预览 == 保存结果"不是巧合，
 * 而是同一段代码跑两遍；渲染层因此不需要自己的样式合并/链解析。
 *
 * @param filePath 当前文档路径；null 表示还没保存过的新文档，用空白模板打底。
 * @param edits 攒下的样式编辑，按创建先后排列（同 id 的后续编辑会覆盖前一条）。
 * @returns ok:true + 窗格清单与样式链；读文件失败时 ok:false + 说明（窗格保持原样）。
 */
async function previewStyles(
  filePath: string | null,
  edits: StyleEdit[],
): Promise<StylePreview> {
  try {
    const docx = await loadForEdit(filePath);
    const { styles, styleFormats } = await docx.applyStyleEdits(edits);
    return { ok: true, styles, styleFormats };
  } catch (err) {
    return {
      ok: false,
      message: `样式预览失败：${err instanceof Error ? err.message : String(err)}`,
    };
  }
}

/**
 * 取一份可以就地改写的文档：路径上有文件就读它，否则用空白模板新建（不落盘）。
 *
 * @param filePath 目标 .docx 路径；null / 空串表示尚无文件（新文档）。
 * @returns 内存中的 `DocxFile`，调用方改完再决定要不要写回。
 */
async function loadForEdit(filePath: string | null): Promise<DocxFile> {
  if (!filePath || !(await exists(filePath))) {
    return DocxFile.fromBytes(await minimalDocxBytes());
  }
  const bytes = await fs.readFile(filePath);
  return DocxFile.fromBytes(bytes);
}

async function exists(p: string): Promise<boolean> {
  try {
    await fs.access(p);
    return true;
  } catch {
    return false;
  }
}

async function minimalDocxBytes(): Promise<Uint8Array> {
  const { default: JSZip } = await import("jszip");
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Default Extension="jpeg" ContentType="image/jpeg"/>
  <Default Extension="jpg" ContentType="image/jpeg"/>
  <Default Extension="gif" ContentType="image/gif"/>
  <Default Extension="webp" ContentType="image/webp"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`,
  );
  zip.file(
    "_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`,
  );
  zip.file(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>`,
  );
  zip.file(
    "word/_rels/document.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`,
  );
  return zip.generateAsync({ type: "uint8array" });
}
