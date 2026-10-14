import type { ResolvedChart, SheetPayload, WorkbookPayload } from "@office-copilot/sheet-io";
import type { PptxEditPayload, PptxRevertPayload, PresentationModel, SlideEditOp } from "@office-copilot/pptx-io";

/** 打开 `.xlsx` 的结果：跨界只有结构化格子，工作表 XML 原文留在主进程。 */
export interface SheetOpenResult {
  ok: boolean;
  canceled?: boolean;
  message?: string;
  path?: string;
  workbook?: WorkbookPayload;
}

/** 保存 `.xlsx` 的请求：表按打开时的顺序带回，`partPath` 用来在主进程重开的包里配对。 */
export interface SheetSaveRequest {
  path: string | null;
  sheets: SheetPayload[];
  /** 本次要新插入的原生图表（AI 出图，spec 决策 D5）；已随文件保存过的那些不用再送。 */
  charts?: { partPath: string; chart: ResolvedChart }[];
  defaultName?: string;
}

/** 保存结果：`ok:false + canceled` 表示用户在另存为对话框里取消了。 */
export interface SheetSaveResult {
  ok: boolean;
  canceled?: boolean;
  message?: string;
  path?: string;
}

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
  /** 批注与修订的署名，写进 `w:author`；空着时 docx-io 用它的占位作者名。 */
  authorName?: string;
}

/** 会话里的一轮文字（用户输入或 AI 回复），续聊时按顺序回填给模型。 */
export interface SessionTurn {
  role: "user" | "assistant";
  text: string;
}

/** 会话列表上的一行，结构见 electron/sessionStore.ts 的 SessionSummary。 */
export interface SessionSummary {
  id: string;
  title: string;
  /** 最后一次改动的时间戳（毫秒）。 */
  updatedAt: number;
  /** 会话里的轮数。 */
  turnCount: number;
}

/** 磁盘上的完整会话，结构见 electron/sessionStore.ts 的 SessionRecord。 */
export interface SessionRecord {
  id: string;
  title: string;
  createdAt: number;
  updatedAt: number;
  turns: SessionTurn[];
}

/**
 * 工具回路里的一条消息，结构见 electron/aiHandlers.ts 的 ToolTurnMessage。
 * `role:"tool"` 的回填必须带 `tool_call_id`，告诉模型这条结果对应哪一次调用。
 */
export interface ToolTurnMessage {
  role: "system" | "user" | "assistant" | "tool";
  content?: string;
  tool_call_id?: string;
  tool_calls?: Array<{
    id: string;
    type: "function";
    function: { name: string; arguments: string };
  }>;
}

/** 模型要求的一次工具调用；`arguments` 是原文 JSON 串，落地前必须过 zod 校验（AGENTS §2.8）。 */
export interface ToolCall {
  id: string;
  name: string;
  arguments: string;
}

/** 一轮工具回路的请求：上下文与工具定义都在渲染层攒，主进程只转发给服务商。 */
export interface ToolTurnRequest {
  messages: ToolTurnMessage[];
  tools: unknown[];
  providerId?: string;
}

/** 一轮的回应：模型只说话时 `toolCalls` 为空数组；`needsSetup` 表示还没配服务商。 */
export type ToolTurnResult =
  | { ok: true; content: string; toolCalls: ToolCall[] }
  | { ok: false; message: string; needsSetup?: boolean };

/** 打开 `.pptx` 的结果：跨界只有结构化形状与文字，幻灯片 XML 原文留在主进程。 */
export interface PptxOpenOutcome {
  ok: boolean;
  canceled?: boolean;
  message?: string;
  path?: string;
  presentation?: PresentationModel;
}

/**
 * 幻灯片保存 / 撤销的结果。
 *
 * `ops` 是主进程现做出的那批前/后像：渲染层拿不到幻灯片原文，所以撤销要把它们原样送回 `ppt:revert`（§2.5）。
 * `presentation` 是主进程写盘后重读的那份，视图直接用它刷新，不必再发一次读请求。
 */
export interface PptxEditOutcome {
  ok: boolean;
  canceled?: boolean;
  message?: string;
  path?: string;
  /** 本次真正动过的幻灯片部件路径。 */
  changedParts?: string[];
  ops?: SlideEditOp[];
  presentation?: PresentationModel;
}

/** 打开对话框的选择结果：只有路径，读 docx 还是 xlsx 由渲染层按扩展名分派。 */
export interface OpenPickResult {
  ok: boolean;
  canceled?: boolean;
  message?: string;
  path?: string;
}

interface OfficeCopilotAPI {
  newFile: () => Promise<unknown>;
  openFile: () => Promise<OpenPickResult>;
  readSheet: (filePath: string) => Promise<SheetOpenResult>;
  readPptx: (filePath: string) => Promise<PptxOpenOutcome>;
  editPptx: (payload: PptxEditPayload) => Promise<PptxEditOutcome>;
  revertPptx: (payload: PptxRevertPayload) => Promise<PptxEditOutcome>;
  savePptxAs: (payload: PptxEditPayload) => Promise<PptxEditOutcome>;
  saveSheet: (payload: SheetSaveRequest) => Promise<SheetSaveResult>;
  saveSheetAs: (payload: SheetSaveRequest) => Promise<SheetSaveResult>;
  saveFile: (payload: {
    path: string | null;
    doc: unknown;
    header?: unknown;
    footer?: unknown;
    section?: unknown;
    link?: unknown;
    /** 待落进 styles.xml 的样式编辑，结构见 docx-io 的 StyleEdit。 */
    styles?: unknown;
    /** 脚注窗格里改过的正文，结构见 docx-io 的 FootnoteBodies。 */
    footnotes?: unknown;
    /** 批注侧栏里改过的正文与署名，结构见 docx-io 的 CommentBody[]。 */
    comments?: unknown;
  }) => Promise<any>;
  saveFileAs: (payload: {
    doc: unknown;
    defaultName?: string;
    header?: unknown;
    footer?: unknown;
    section?: unknown;
    link?: unknown;
    styles?: unknown;
    footnotes?: unknown;
    comments?: unknown;
  }) => Promise<any>;
  readDocx: (filePath: string) => Promise<any>;
  pathForFile: (file: File) => string;
  writeDocx: (payload: {
    path: string;
    doc: unknown;
    header?: unknown;
    footer?: unknown;
  }) => Promise<any>;
  /**
   * 样式窗格预览：主进程内存里合并样式编辑，返回窗格清单与样式链，不写磁盘。
   * styles / styleFormats 的结构分别是 docx-io 的 PaneStyle[] / ResolvedStyle[]。
   */
  previewStyles: (payload: {
    path: string | null;
    styles: unknown;
  }) => Promise<{
    ok: boolean;
    styles?: unknown;
    styleFormats?: unknown;
    message?: string;
  }>;
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
    /** 续聊时带上的历史轮次，主进程截到最近若干轮再上送。 */
    history?: SessionTurn[];
  }) => Promise<{ ok: true; text: string } | { ok: false; message: string; needsSetup?: boolean }>;
  /** 工具回路的一轮：见 electron/aiHandlers.ts 的 `ai:tool-turn`。上下文由调用方攒。 */
  aiToolTurn: (payload: ToolTurnRequest) => Promise<ToolTurnResult>;
  aiStreamStart: (payload: {
    requestId: string;
    action: string;
    text: string;
    instruction?: string;
    system?: string;
    providerId?: string;
    history?: SessionTurn[];
  }) => Promise<{ ok: true } | { ok: false; message: string; needsSetup?: boolean }>;
  aiStreamCancel: (requestId: string) => Promise<{ ok: true }>;
  onAIStream: (
    cb: (ev: {
      requestId: string;
      kind: "delta" | "done" | "error";
      text?: string;
      message?: string;
      needsSetup?: boolean;
      /** `kind:"done"` 时为 true 表示用户按了暂停，不是模型写完了。 */
      stopped?: boolean;
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
  /** 会话历史列表（userData/sessions 下的文件，按更新时间从新到旧）。 */
  listSessions: () => Promise<SessionSummary[]>;
  /** 读一个会话的完整正文；id 不合法或文件不在时返回 null。 */
  readSession: (id: string) => Promise<SessionRecord | null>;
  /** 覆盖保存一个会话（不存在就新建）；内容为空时返回 ok:false。 */
  saveSession: (payload: {
    id: string;
    title?: string;
    turns?: SessionTurn[];
  }) => Promise<{ ok: boolean; message?: string }>;
  removeSession: (id: string) => Promise<{ ok: boolean }>;
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
