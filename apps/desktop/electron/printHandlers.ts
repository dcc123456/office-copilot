import { app, BrowserWindow, dialog, type IpcMain } from "electron";
import fs from "node:fs/promises";
import path from "node:path";

/**
 * Small helper that opens a hidden BrowserWindow, loads the supplied
 * HTML, then either shows the print dialog or exports a PDF. We use
 * a dedicated window so the toolbar / sidebars don't show up on
 * paper. The renderer is responsible for building print-ready HTML
 * (A4 page boxes, margins, header/footer).
 */

export function registerPrintHandlers(ipc: IpcMain): void {
  ipc.handle(
    "print:pdf",
    async (
      e,
      payload: { html: string; defaultName?: string },
    ): Promise<{ ok: true; path: string } | { ok: false; message: string; canceled?: boolean }> => {
      const win = e.sender as unknown as BrowserWindow;
      const result = await dialog.showSaveDialog(win, {
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
        await fs.writeFile(tmp, pdfBytes);
        await fs.rename(tmp, result.filePath);
        return { ok: true, path: result.filePath };
      } catch (err) {
        return {
          ok: false,
          message: `导出 PDF 失败：${err instanceof Error ? err.message : String(err)}`,
        };
      }
    },
  );

  ipc.handle("print:dialog", async (_e, payload: { html: string }) => {
    try {
      await showPrintDialog(payload.html);
      return { ok: true };
    } catch (err) {
      return {
        ok: false,
        message: err instanceof Error ? err.message : String(err),
      };
    }
  });
}

async function makeWindow(html: string): Promise<BrowserWindow> {
  const child = new BrowserWindow({
    show: false,
    webPreferences: { offscreen: false, sandbox: true },
    parent: undefined,
  });
  child.webContents.on("did-finish-load", () => {
    child.webContents.setZoomFactor(1);
  });
  await child.loadURL(
    "data:text/html;charset=utf-8," + encodeURIComponent(html),
  );
  return child;
}

async function renderPdf(html: string): Promise<Uint8Array> {
  const child = await makeWindow(html);
  try {
    return await child.webContents.printToPDF({
      pageSize: "A4",
      printBackground: true,
      // We control margins inside the HTML (each .page has its own
      // margin), so tell Chromium to use zero outer margins.
      margins: { top: 0, bottom: 0, left: 0, right: 0 },
    });
  } finally {
    child.destroy();
  }
}

async function showPrintDialog(html: string): Promise<void> {
  const child = await makeWindow(html);
  return new Promise((resolve, reject) => {
    child.webContents.print(
      { silent: false, printBackground: true, pageSize: "A4" },
      (success, err) => {
        child.destroy();
        if (success) resolve();
        else reject(new Error(err ?? "打印失败"));
      },
    );
  });
}

/** Resolve a path under userData; callers can use it for caches/drafts. */
export function userDataPath(...parts: string[]): string {
  return path.join(app.getPath("userData"), ...parts);
}
