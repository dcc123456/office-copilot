"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.registerPrintHandlers = registerPrintHandlers;
exports.userDataPath = userDataPath;
const electron_1 = require("electron");
const promises_1 = __importDefault(require("node:fs/promises"));
const node_path_1 = __importDefault(require("node:path"));
/**
 * Small helper that opens a hidden BrowserWindow, loads the supplied
 * HTML, then either shows the print dialog or exports a PDF. We use
 * a dedicated window so the toolbar / sidebars don't show up on
 * paper. The renderer is responsible for building print-ready HTML
 * (A4 page boxes, margins, header/footer).
 */
function registerPrintHandlers(ipc) {
    ipc.handle("print:pdf", async (e, payload) => {
        const win = e.sender;
        const result = await electron_1.dialog.showSaveDialog(win, {
            title: "导出 PDF",
            defaultPath: payload.defaultName ?? "document.pdf",
            filters: [{ name: "PDF", extensions: ["pdf"] }],
        });
        if (result.canceled || !result.filePath) {
            return { ok: false, canceled: true, message: "已取消" };
        }
        try {
            const pdfBytes = await renderPdf(payload.html);
            const tmp = result.filePath + ".tmp";
            await promises_1.default.writeFile(tmp, pdfBytes);
            await promises_1.default.rename(tmp, result.filePath);
            return { ok: true, path: result.filePath };
        }
        catch (err) {
            return {
                ok: false,
                message: `导出 PDF 失败：${err instanceof Error ? err.message : String(err)}`,
            };
        }
    });
    ipc.handle("print:dialog", async (_e, payload) => {
        try {
            await showPrintDialog(payload.html);
            return { ok: true };
        }
        catch (err) {
            return {
                ok: false,
                message: err instanceof Error ? err.message : String(err),
            };
        }
    });
}
async function makeWindow(html) {
    const child = new electron_1.BrowserWindow({
        show: false,
        webPreferences: { offscreen: false, sandbox: true },
        parent: undefined,
    });
    child.webContents.on("did-finish-load", () => {
        child.webContents.setZoomFactor(1);
    });
    await child.loadURL("data:text/html;charset=utf-8," + encodeURIComponent(html));
    return child;
}
async function renderPdf(html) {
    const child = await makeWindow(html);
    try {
        return await child.webContents.printToPDF({
            pageSize: "A4",
            printBackground: true,
            // We control margins inside the HTML (each .page has its own
            // margin), so tell Chromium to use zero outer margins.
            margins: { top: 0, bottom: 0, left: 0, right: 0 },
        });
    }
    finally {
        child.destroy();
    }
}
async function showPrintDialog(html) {
    const child = await makeWindow(html);
    return new Promise((resolve, reject) => {
        child.webContents.print({ silent: false, printBackground: true, pageSize: "A4" }, (success, err) => {
            child.destroy();
            if (success)
                resolve();
            else
                reject(new Error(err ?? "打印失败"));
        });
    });
}
/** Resolve a path under userData; callers can use it for caches/drafts. */
function userDataPath(...parts) {
    return node_path_1.default.join(electron_1.app.getPath("userData"), ...parts);
}
//# sourceMappingURL=printHandlers.js.map