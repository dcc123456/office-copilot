"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || (function () {
    var ownKeys = function(o) {
        ownKeys = Object.getOwnPropertyNames || function (o) {
            var ar = [];
            for (var k in o) if (Object.prototype.hasOwnProperty.call(o, k)) ar[ar.length] = k;
            return ar;
        };
        return ownKeys(o);
    };
    return function (mod) {
        if (mod && mod.__esModule) return mod;
        var result = {};
        if (mod != null) for (var k = ownKeys(mod), i = 0; i < k.length; i++) if (k[i] !== "default") __createBinding(result, mod, k[i]);
        __setModuleDefault(result, mod);
        return result;
    };
})();
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.DEFAULT_SECTION = void 0;
exports.registerFileHandlers = registerFileHandlers;
const electron_1 = require("electron");
const promises_1 = __importDefault(require("node:fs/promises"));
const node_path_1 = __importDefault(require("node:path"));
const docx_io_1 = require("@office-copilot/docx-io");
/** Default geometry for a brand-new document (A4, 1-inch margins). */
exports.DEFAULT_SECTION = {
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
function registerFileHandlers(ipc, getWindow) {
    ipc.handle("file:new", async () => {
        return {
            ok: true,
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
            },
            header: null,
            footer: null,
            section: exports.DEFAULT_SECTION,
            warnings: [],
        };
    });
    ipc.handle("file:open", async () => {
        const win = getWindow();
        const result = await electron_1.dialog.showOpenDialog(win, {
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
        const filePath = result.filePaths[0];
        return openDocx(filePath);
    });
    ipc.handle("file:read-docx", async (_e, filePath) => openDocx(filePath));
    ipc.handle("file:save", async (_e, payload) => {
        if (!payload.path)
            return { ok: false, message: "没有目标路径。" };
        return saveDocx(payload.path, payload.doc, {
            header: payload.header ?? null,
            footer: payload.footer ?? null,
            section: payload.section ?? null,
            link: payload.link ?? null,
        });
    });
    ipc.handle("file:save-as", async (_e, payload) => {
        const win = getWindow();
        const result = await electron_1.dialog.showSaveDialog(win, {
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
        if (!r.ok)
            return r;
        return { ok: true, path: result.filePath };
    });
    ipc.handle("file:write-docx", async (_e, payload) => saveDocx(payload.path, payload.doc, {
        header: payload.header ?? null,
        footer: payload.footer ?? null,
    }));
    ipc.handle("file:pick-image", async () => {
        const win = getWindow();
        const result = await electron_1.dialog.showOpenDialog(win, {
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
            return { ok: false, canceled: true };
        }
        const p = result.filePaths[0];
        const buf = await promises_1.default.readFile(p);
        const ext = p.split(".").pop()?.toLowerCase() ?? "png";
        const mime = ext === "jpg" || ext === "jpeg"
            ? "image/jpeg"
            : ext === "gif"
                ? "image/gif"
                : ext === "webp"
                    ? "image/webp"
                    : "image/png";
        const dataUrl = `data:${mime};base64,${buf.toString("base64")}`;
        return { ok: true, path: p, dataUrl };
    });
}
async function openDocx(filePath) {
    try {
        const bytes = await promises_1.default.readFile(filePath);
        const docx = await docx_io_1.DocxFile.fromBytes(bytes);
        const { doc, warnings, headers, section } = await docx.toProseMirror();
        return {
            ok: true,
            path: filePath,
            doc,
            warnings,
            header: headers?.entries.header?.pm ?? null,
            footer: headers?.entries.footer?.pm ?? null,
            section,
        };
    }
    catch (err) {
        return {
            ok: false,
            message: `打开失败：${err instanceof Error ? err.message : String(err)}`,
        };
    }
}
async function saveDocx(filePath, pmDoc, hf) {
    try {
        let docx;
        let isNew = false;
        if (await exists(filePath)) {
            const bytes = await promises_1.default.readFile(filePath);
            docx = await docx_io_1.DocxFile.fromBytes(bytes);
        }
        else {
            docx = await docx_io_1.DocxFile.fromBytes(await minimalDocxBytes());
            isNew = true;
        }
        if (hf.header)
            docx.setHeaderFooter("header", hf.header);
        if (hf.footer)
            docx.setHeaderFooter("footer", hf.footer);
        if (hf.section)
            docx.setSectionGeometry(hf.section);
        if (hf.link)
            docx.setDocStyle({ link: hf.link });
        const warnings = await docx.replaceProseMirror(pmDoc);
        const out = await docx.toBytes();
        const tmp = filePath + ".tmp";
        await promises_1.default.writeFile(tmp, out);
        await promises_1.default.rename(tmp, filePath);
        if (isNew) {
            void 0;
        }
        void warnings;
        return { ok: true, path: filePath };
    }
    catch (err) {
        return {
            ok: false,
            message: `保存失败：${err instanceof Error ? err.message : String(err)}`,
        };
    }
}
async function exists(p) {
    try {
        await promises_1.default.access(p);
        return true;
    }
    catch {
        return false;
    }
}
async function minimalDocxBytes() {
    const { default: JSZip } = await Promise.resolve().then(() => __importStar(require("jszip")));
    const zip = new JSZip();
    zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Default Extension="jpeg" ContentType="image/jpeg"/>
  <Default Extension="jpg" ContentType="image/jpeg"/>
  <Default Extension="gif" ContentType="image/gif"/>
  <Default Extension="webp" ContentType="image/webp"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`);
    zip.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`);
    zip.file("word/document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>`);
    zip.file("word/_rels/document.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`);
    void node_path_1.default;
    return zip.generateAsync({ type: "uint8array" });
}
//# sourceMappingURL=fileHandlers.js.map