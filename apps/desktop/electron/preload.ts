import { contextBridge, ipcRenderer, webUtils } from "electron";

interface AITransformPayload {
  action: string;
  text: string;
  instruction?: string;
  system?: string;
}

interface AIResult {
  ok: boolean;
  text?: string;
  message?: string;
}

interface OpenResult {
  ok: boolean;
  canceled?: boolean;
  path?: string;
  doc?: unknown;
  header?: unknown;
  footer?: unknown;
  section?: unknown;
  /** styles.xml 里声明的段落样式清单，见 @office-copilot/docx-io 的 PaneStyle。 */
  styles?: unknown;
  warnings?: string[];
  message?: string;
}

interface SaveResult {
  ok: boolean;
  canceled?: boolean;
  path?: string;
  message?: string;
}

interface SavePayload {
  path: string | null;
  doc: unknown;
  header?: unknown;
  footer?: unknown;
  section?: unknown;
  link?: unknown;
}

/**
 * The renderer only talks to the main process through this typed surface.
 */
const api = {
  newFile: (): Promise<OpenResult> => ipcRenderer.invoke("file:new"),
  openFile: (): Promise<OpenResult> => ipcRenderer.invoke("file:open"),
  saveFile: (payload: SavePayload): Promise<SaveResult> =>
    ipcRenderer.invoke("file:save", payload),
  saveFileAs: (
    payload: SavePayload & { defaultName?: string },
  ): Promise<SaveResult> => ipcRenderer.invoke("file:save-as", payload),
  readDocx: (filePath: string): Promise<OpenResult> =>
    ipcRenderer.invoke("file:read-docx", filePath),
  /** Electron 32: File.path is removed; resolve a dropped File to its disk path. */
  pathForFile: (file: File): string => webUtils.getPathForFile(file),
  writeDocx: (payload: { path: string; doc: unknown }): Promise<SaveResult> =>
    ipcRenderer.invoke("file:write-docx", payload),
  pickImage: (): Promise<
    | { ok: true; path: string; dataUrl: string }
    | { ok: false; canceled: true }
  > => ipcRenderer.invoke("file:pick-image"),
  /** 右键菜单的粘贴来源：主进程读系统剪贴板，渲染层不直接碰 Node 能力。 */
  readClipboard: (): Promise<{ text: string; html: string }> =>
    ipcRenderer.invoke("clipboard:read"),
  aiTransform: (payload: AITransformPayload): Promise<AIResult> =>
    ipcRenderer.invoke("ai:transform", payload),
  aiStreamStart: (
    payload: { requestId: string } & AITransformPayload & {
      providerId?: string;
    },
  ): Promise<AIResult & { ok: boolean }> =>
    ipcRenderer.invoke("ai:stream-start", payload),
  aiStreamCancel: (requestId: string): Promise<{ ok: true }> =>
    ipcRenderer.invoke("ai:stream-cancel", requestId),
  onAIStream: (
    cb: (ev: {
      requestId: string;
      kind: "delta" | "done" | "error";
      text?: string;
      message?: string;
      needsSetup?: boolean;
    }) => void,
  ): (() => void) => {
    const listener = (
      _e: unknown,
      ev: {
        requestId: string;
        kind: "delta" | "done" | "error";
        text?: string;
        message?: string;
        needsSetup?: boolean;
      },
    ) => cb(ev);
    ipcRenderer.on("ai:stream", listener);
    return () => ipcRenderer.off("ai:stream", listener);
  },
  getSettings: (): Promise<unknown> => ipcRenderer.invoke("settings:get"),
  saveSettings: (settings: unknown): Promise<{ ok: true }> =>
    ipcRenderer.invoke("settings:save", settings),
  testProvider: (
    provider: unknown,
  ): Promise<{ ok: true } | { ok: false; message: string }> =>
    ipcRenderer.invoke("settings:test-provider", provider),
  saveSnapshot: (payload: {
    originalPath: string | null;
    bytes: ArrayBuffer | Uint8Array;
    reason?: string;
    ext?: string;
  }): Promise<{ ok: true; path: string } | { ok: false; message: string }> =>
    ipcRenderer.invoke("snapshot:save", payload),
  listSnapshots: (payload: {
    originalPath: string | null;
  }): Promise<Array<{ path: string; mtime: number; size: number }>> =>
    ipcRenderer.invoke("snapshot:list", payload),
  readSnapshot: (
    snapshotPath: string,
  ): Promise<{ ok: true; bytes: ArrayBuffer } | { ok: false; message: string }> =>
    ipcRenderer.invoke("snapshot:read", { path: snapshotPath }),

  // AI modification log (diff + per-change undo).
  listModifications: (docKey: string): Promise<unknown[]> =>
    ipcRenderer.invoke("modlog:list", docKey),
  appendModification: (payload: {
    docKey: string;
    record: unknown;
  }): Promise<{ ok: boolean; message?: string }> =>
    ipcRenderer.invoke("modlog:append", payload),
  patchModification: (payload: {
    docKey: string;
    id: string;
    patch: Record<string, unknown>;
  }): Promise<{ ok: boolean; message?: string }> =>
    ipcRenderer.invoke("modlog:patch", payload),
  removeModification: (payload: {
    docKey: string;
    id?: string;
  }): Promise<{ ok: boolean }> => ipcRenderer.invoke("modlog:remove", payload),

  // Print / export PDF.
  printToPdf: (html: string, defaultName?: string) =>
    ipcRenderer.invoke("print:pdf", { html, defaultName }),
  printDialog: (html: string) =>
    ipcRenderer.invoke("print:dialog", { html }),

  // Recent files.
  listRecent: (): Promise<string[]> => ipcRenderer.invoke("recent:list"),
  clearRecent: (): Promise<string[]> => ipcRenderer.invoke("recent:clear"),
  addRecent: (p: string) => ipcRenderer.send("recent:add", p),
  onRecentChanged: (cb: (paths: string[]) => void): (() => void) => {
    const listener = (_e: unknown, paths: string[]) => cb(paths);
    ipcRenderer.on("recent:changed", listener);
    return () => ipcRenderer.off("recent:changed", listener);
  },
  onOpenRecent: (cb: (path: string) => void): (() => void) => {
    const listener = (_e: unknown, p: string) => cb(p);
    ipcRenderer.on("menu:open-recent", listener);
    return () => ipcRenderer.off("menu:open-recent", listener);
  },

  onMenu: (channel: string, cb: () => void): (() => void) => {
    const listener = () => cb();
    ipcRenderer.on(channel, listener);
    return () => ipcRenderer.off(channel, listener);
  },

  getWindowBounds: (): Promise<{
    x: number;
    y: number;
    width: number;
    height: number;
  }> => ipcRenderer.invoke("window:get-bounds"),
  dragResizeWindow: (payload: {
    base: { x: number; y: number; width: number; height: number };
    dx: number;
    dy: number;
    edge: string;
  }): void => ipcRenderer.send("window:drag-resize", payload),
};

contextBridge.exposeInMainWorld("officeCopilot", api);

export type OfficeCopilotAPI = typeof api;
