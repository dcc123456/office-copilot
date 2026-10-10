import { dialog, type IpcMain, type BrowserWindow } from "electron";
import fs from "node:fs/promises";
import path from "node:path";
import {
  DocxFile,
  type LinkStyle,
  type PaneStyle,
  type ResolvedStyle,
  type SectionGeometry,
} from "@office-copilot/docx-io";
import type { PMDoc } from "@office-copilot/docx-io/src/pm-schema.js";

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
}

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
      payload: {
        path: string | null;
        doc: PMDoc;
        header?: PMDoc | null;
        footer?: PMDoc | null;
        section?: SectionGeometry | null;
        link?: LinkStyle | null;
      },
    ): Promise<{ ok: true; path: string } | { ok: false; message: string }> => {
      if (!payload.path) return { ok: false, message: "没有目标路径。" };
      return saveDocx(payload.path, payload.doc, {
        header: payload.header ?? null,
        footer: payload.footer ?? null,
        section: payload.section ?? null,
        link: payload.link ?? null,
      });
    },
  );

  ipc.handle(
    "file:save-as",
    async (
      _e,
      payload: {
        doc: PMDoc;
        defaultName?: string;
        header?: PMDoc | null;
        footer?: PMDoc | null;
        section?: SectionGeometry | null;
        link?: LinkStyle | null;
      },
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
      const r = await saveDocx(result.filePath, payload.doc, {
        header: payload.header ?? null,
        footer: payload.footer ?? null,
        section: payload.section ?? null,
        link: payload.link ?? null,
      });
      if (!r.ok) return r;
      return { ok: true, path: result.filePath };
    },
  );

  ipc.handle(
    "file:write-docx",
    async (
      _e,
      payload: {
        path: string;
        doc: PMDoc;
        header?: PMDoc | null;
        footer?: PMDoc | null;
      },
    ) =>
      saveDocx(payload.path, payload.doc, {
        header: payload.header ?? null,
        footer: payload.footer ?? null,
      }),
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
    const { doc, warnings, headers, section, styles, styleFormats } =
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
    };
  } catch (err) {
    return {
      ok: false,
      message: `打开失败：${err instanceof Error ? err.message : String(err)}`,
    };
  }
}
async function saveDocx(
  filePath: string,
  pmDoc: PMDoc,
  hf: HeaderFooterPayload,
): Promise<{ ok: true; path: string } | { ok: false; message: string }> {
  try {
    let docx: DocxFile;
    let isNew = false;
    if (await exists(filePath)) {
      const bytes = await fs.readFile(filePath);
      docx = await DocxFile.fromBytes(bytes);
    } else {
      docx = await DocxFile.fromBytes(await minimalDocxBytes());
      isNew = true;
    }
    if (hf.header) docx.setHeaderFooter("header", hf.header);
    if (hf.footer) docx.setHeaderFooter("footer", hf.footer);
    if (hf.section) docx.setSectionGeometry(hf.section);
    if (hf.link) docx.setDocStyle({ link: hf.link });
    const warnings = await docx.replaceProseMirror(pmDoc);
    const out = await docx.toBytes();
    const tmp = filePath + ".tmp";
    await fs.writeFile(tmp, out);
    await fs.rename(tmp, filePath);
    if (isNew) {
      void 0;
    }
    void warnings;
    return { ok: true, path: filePath } as { ok: true; path: string } & {
      warnings?: string[];
    };
  } catch (err) {
    return {
      ok: false,
      message: `保存失败：${err instanceof Error ? err.message : String(err)}`,
    };
  }
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
  void path;
  return zip.generateAsync({ type: "uint8array" });
}
