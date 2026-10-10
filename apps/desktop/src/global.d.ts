export interface ProviderConfig {
  id: string;
  name: string;
  baseUrl: string;
  apiKey: string;
  model: string;
  authHeader?: string;
  extraBody?: Record<string, unknown>;
  temperature?: number;
  enabled: boolean;
}

export interface AppSettings {
  version: 1;
  providers: ProviderConfig[];
  defaultProviderId: string | null;
  proxy?: string;
}

interface OfficeCopilotAPI {
  newFile: () => Promise<unknown>;
  openFile: () => Promise<{
    ok: boolean;
    canceled?: boolean;
    message?: string;
    path?: string | null;
    doc?: unknown;
    header?: unknown;
    footer?: unknown;
    section?: unknown;
    /** styles.xml 声明的段落样式清单，结构见 docx-io 的 PaneStyle。 */
    styles?: unknown;
    warnings?: string[];
  }>;
  saveFile: (payload: {
    path: string | null;
    doc: unknown;
    header?: unknown;
    footer?: unknown;
    section?: unknown;
    link?: unknown;
  }) => Promise<any>;
  saveFileAs: (payload: {
    doc: unknown;
    defaultName?: string;
    header?: unknown;
    footer?: unknown;
    section?: unknown;
    link?: unknown;
  }) => Promise<any>;
  readDocx: (filePath: string) => Promise<any>;
  pathForFile: (file: File) => string;
  writeDocx: (payload: {
    path: string;
    doc: unknown;
    header?: unknown;
    footer?: unknown;
  }) => Promise<any>;
  pickImage: () => Promise<
    | { ok: true; path: string; dataUrl: string }
    | { ok: false; canceled: true }
  >;
  aiTransform: (payload: {
    action: string;
    text: string;
    instruction?: string;
    system?: string;
    providerId?: string;
  }) => Promise<{ ok: true; text: string } | { ok: false; message: string; needsSetup?: boolean }>;
  aiStreamStart: (payload: {
    requestId: string;
    action: string;
    text: string;
    instruction?: string;
    system?: string;
    providerId?: string;
  }) => Promise<{ ok: true } | { ok: false; message: string; needsSetup?: boolean }>;
  aiStreamCancel: (requestId: string) => Promise<{ ok: true }>;
  onAIStream: (
    cb: (ev: {
      requestId: string;
      kind: "delta" | "done" | "error";
      text?: string;
      message?: string;
      needsSetup?: boolean;
    }) => void,
  ) => () => void;
  getSettings: () => Promise<AppSettings>;
  saveSettings: (settings: AppSettings) => Promise<{ ok: true }>;
  testProvider: (
    provider: ProviderConfig,
  ) => Promise<{ ok: true } | { ok: false; message: string }>;
  saveSnapshot: (payload: {
    originalPath: string | null;
    bytes: ArrayBuffer | Uint8Array;
    reason?: string;
    ext?: string;
  }) => Promise<{ ok: true; path: string } | { ok: false; message: string }>;
  listSnapshots: (payload: {
    originalPath: string | null;
  }) => Promise<Array<{ path: string; mtime: number; size: number }>>;
  readSnapshot: (
    snapshotPath: string,
  ) => Promise<{ ok: true; bytes: ArrayBuffer } | { ok: false; message: string }>;
  listModifications: (docKey: string) => Promise<unknown[]>;
  appendModification: (payload: {
    docKey: string;
    record: unknown;
  }) => Promise<{ ok: boolean; message?: string }>;
  patchModification: (payload: {
    docKey: string;
    id: string;
    patch: Record<string, unknown>;
  }) => Promise<{ ok: boolean; message?: string }>;
  removeModification: (payload: { docKey: string; id?: string }) => Promise<{ ok: boolean }>;
  printToPdf: (
    html: string,
    defaultName?: string,
  ) => Promise<{ ok: true; path: string } | { ok: false; message: string; canceled?: boolean }>;
  printDialog: (
    html: string,
  ) => Promise<{ ok: true } | { ok: false; message: string }>;
  listRecent: () => Promise<string[]>;
  readClipboard: () => Promise<{ text: string; html: string }>;
  clearRecent: () => Promise<string[]>;
  addRecent: (path: string) => void;
  onRecentChanged: (cb: (paths: string[]) => void) => () => void;
  onOpenRecent: (cb: (path: string) => void) => () => void;
  onMenu: (channel: string, cb: () => void) => () => void;
  getWindowBounds: () => Promise<{
    x: number;
    y: number;
    width: number;
    height: number;
  }>;
  dragResizeWindow: (payload: {
    base: { x: number; y: number; width: number; height: number };
    dx: number;
    dy: number;
    edge: "n" | "s" | "e" | "w" | "ne" | "nw" | "se" | "sw";
  }) => void;
}

declare global {
  interface Window {
    officeCopilot: OfficeCopilotAPI;
  }
}

export {};
