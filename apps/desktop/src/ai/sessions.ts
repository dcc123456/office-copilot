import type { SessionRecord, SessionSummary, SessionTurn } from "../global.js";

/**
 * AI 会话的渲染层门面：id、标题、历史回填这几件纯事实在这里，跨界读写只经
 * `window.officeCopilot`（`electron/sessionStore.ts` 那四个 handler）。
 * 界面文案（卡片状态怎么说）留在面板里组装，这里只搬结构化轮次（AGENTS §5.7）。
 */

/** 标题取第一句用户说的话，太长就截断——列表里一眼能认出是哪个会话就够了。 */
const TITLE_CHARS = 24;

/** 一个全新的会话编号：与磁盘上的文件名一一对应。 */
export function newSessionId(): string {
  return crypto.randomUUID();
}

/**
 * 从若干轮对话里给会话起个标题。
 *
 * @param turns 会话正文（用户与 AI 交替的文字）。
 * @returns 首条用户文字的截断；一条用户文字都没有时返回"未命名会话"。
 */
export function titleOf(turns: SessionTurn[]): string {
  const firstAsk = turns.find((t) => t.role === "user");
  const text = firstAsk?.text.replace(/\s+/g, " ").trim() ?? "";
  if (!text) return "未命名会话";
  return text.length > TITLE_CHARS ? `${text.slice(0, TITLE_CHARS)}…` : text;
}

/** 会话列表：按更新时间从新到旧，读不到就空列表（首次用没有目录是正常情况）。 */
export async function listSessions(): Promise<SessionSummary[]> {
  return await window.officeCopilot.listSessions();
}

/**
 * 打开一个会话。
 *
 * @param id 会话编号。
 * @returns 完整会话；文件不存在或编号不合法时为 null。
 */
export async function openSession(id: string): Promise<SessionRecord | null> {
  return await window.officeCopilot.readSession(id);
}

/**
 * 覆盖保存一个会话（不存在即新建）。
 *
 * @param opts.id 会话编号。
 * @param opts.turns 要留下的全部轮次，标题按它现算。
 * @returns 主进程的落盘结果；会话空不空由主进程判（它是唯一的口径），空会话返回 `ok:false`。
 */
export async function saveSession(opts: {
  id: string;
  turns: SessionTurn[];
}): Promise<{ ok: boolean; message?: string }> {
  return await window.officeCopilot.saveSession({
    id: opts.id,
    title: titleOf(opts.turns),
    turns: opts.turns,
  });
}

/** 删掉一个会话；文件本来就不在也算成功。 */
export async function removeSession(id: string): Promise<boolean> {
  const r = await window.officeCopilot.removeSession(id);
  return r.ok;
}
