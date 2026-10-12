import { app, type IpcMain } from "electron";
import path from "node:path";
import fs from "node:fs/promises";

/**
 * AI 会话存储：一个会话一个 JSON 文件，落在 userData/sessions 下。
 *
 * 写法与修改历史（modlog.ts）一致——原子落盘、越界即剪、读到的内容一律当不可信校验一遍；
 * 只是这里按会话切文件，列表面板要显示标题和时间，所以打开目录逐个读。渲染层只经
 * preload 的白名单访问这些 handler，自己不碰文件系统（AGENTS §5.2）。
 */

/** 磁盘上最多留多少个会话，超出的按更新时间剪掉最旧的。 */
const CAP_SESSIONS = 50;
/** 单个会话最多留多少轮，超出的从头部剪。 */
const CAP_TURNS = 200;
/** 单轮文字的上限（字符数）：会话是用户自己打的字，只要防"整篇文档粘进来"这种意外。 */
const CAP_TURN_CHARS = 4000;

export interface SessionTurn {
  role: "user" | "assistant";
  text: string;
}

export interface SessionRecord {
  id: string;
  title: string;
  createdAt: number;
  updatedAt: number;
  turns: SessionTurn[];
}

export interface SessionSummary {
  id: string;
  title: string;
  updatedAt: number;
  /** 会话里的轮数，列表上显示"多少条对话"。 */
  turnCount: number;
}

/** 会话 id 由渲染层用 crypto.randomUUID() 生成；读盘前必须验一遍，防路径穿越。 */
const SESSION_ID = /^[A-Za-z0-9._-]{1,64}$/;

function root(): string {
  return path.join(app.getPath("userData"), "sessions");
}

function fileFor(id: string): string {
  return path.join(root(), `${id}.json`);
}

/**
 * 把磁盘上读到的一份 JSON 收成可用的会话记录。
 *
 * @param raw 文件里 parse 出来的任意值（旧版本、被手改过的、写坏的都算）。
 * @param id 文件名给出的会话 id，作为最终口径，不信任文件里那份。
 * @returns 校验好的 SessionRecord；`raw` 完全不像会话时返回 null。
 */
export function normaliseRecord(
  raw: unknown,
  id: string,
): SessionRecord | null {
  if (!raw || typeof raw !== "object") return null;
  const o = raw as Record<string, unknown>;
  if (!Array.isArray(o.turns)) return null;
  const turns: SessionTurn[] = [];
  for (const t of o.turns) {
    if (!t || typeof t !== "object") continue;
    const role = (t as Record<string, unknown>).role;
    const text = (t as Record<string, unknown>).text;
    if (typeof text !== "string" || !text.trim()) continue;
    if (role !== "user" && role !== "assistant") continue;
    turns.push({ role, text: text.slice(0, CAP_TURN_CHARS) });
  }
  const now = Date.now();
  const createdAt = typeof o.createdAt === "number" ? o.createdAt : now;
  return {
    id,
    title: typeof o.title === "string" && o.title.trim() ? o.title.trim().slice(0, 60) : "未命名会话",
    createdAt,
    updatedAt: typeof o.updatedAt === "number" ? o.updatedAt : createdAt,
    turns: turns.slice(-CAP_TURNS),
  };
}

/**
 * 列出磁盘上的会话，按更新时间从新到旧。
 *
 * @returns 会话摘要清单；目录还不存在返回空列表，某个文件读坏了跳过它继续列。
 */
async function listSummaries(): Promise<SessionSummary[]> {
  const names = await fs.readdir(root()).catch(() => [] as string[]);
  const out: SessionSummary[] = [];
  for (const name of names) {
    if (!name.endsWith(".json")) continue;
    const id = name.slice(0, -".json".length);
    if (!SESSION_ID.test(id)) continue;
    const rec = await readRecord(id);
    if (rec) {
      out.push({ id: rec.id, title: rec.title, updatedAt: rec.updatedAt, turnCount: rec.turns.length });
    }
  }
  return out.sort((a, b) => b.updatedAt - a.updatedAt);
}

/**
 * 读一个会话文件。
 *
 * @param id 会话编号（已经过 `SESSION_ID` 校验）。
 * @returns 会话记录；文件不存在、读不出、JSON 坏了、不像会话时都是 null。
 */
async function readRecord(id: string): Promise<SessionRecord | null> {
  try {
    const raw = await fs.readFile(fileFor(id), "utf8");
    return normaliseRecord(JSON.parse(raw), id);
  } catch {
    return null;
  }
}

/**
 * 原子写入一个会话：先写 .tmp 再 rename，中途断电也不会留下半份 JSON。
 *
 * @param record 已经过 `normaliseRecord` 口径的会话。
 */
async function writeRecord(record: SessionRecord): Promise<void> {
  const dir = root();
  await fs.mkdir(dir, { recursive: true });
  const file = fileFor(record.id);
  const tmp = `${file}.tmp`;
  await fs.writeFile(tmp, JSON.stringify(record), "utf8");
  await fs.rename(tmp, file);
}

/**
 * 会话总数超过上限时，按更新时间剪掉最旧的几个。
 *
 * @param keep 要保留的会话数上限。
 */
async function prune(keep: number): Promise<void> {
  const all = await listSummaries();
  for (const stale of all.slice(keep)) {
    await fs.unlink(fileFor(stale.id)).catch(() => {});
  }
}

export function registerSessionStoreHandlers(ipc: IpcMain): void {
  ipc.handle("session:list", async (): Promise<SessionSummary[]> => await listSummaries());

  ipc.handle(
    "session:read",
    async (_e, id: string): Promise<SessionRecord | null> => {
      if (!SESSION_ID.test(String(id ?? ""))) return null;
      return await readRecord(String(id));
    },
  );

  ipc.handle(
    "session:save",
    async (
      _e,
      payload: { id: string; title?: string; turns?: unknown },
    ): Promise<{ ok: boolean; message?: string }> => {
      const id = String(payload?.id ?? "");
      if (!SESSION_ID.test(id)) return { ok: false, message: "会话编号不合法。" };
      try {
        const prev = await readRecord(id);
        const now = Date.now();
        const draft = normaliseRecord(
          {
            title: payload?.title,
            createdAt: prev?.createdAt ?? now,
            updatedAt: now,
            turns: payload?.turns,
          },
          id,
        );
        // 空会话不留文件：列表里多一条没内容的记录，比不存更让人困惑。
        if (!draft || !draft.turns.length) {
          return { ok: false, message: "会话内容是空的，没有保存。" };
        }
        await writeRecord(draft);
        await prune(CAP_SESSIONS);
        return { ok: true };
      } catch (err) {
        return { ok: false, message: err instanceof Error ? err.message : String(err) };
      }
    },
  );

  ipc.handle(
    "session:remove",
    async (_e, id: string): Promise<{ ok: boolean }> => {
      if (!SESSION_ID.test(String(id ?? ""))) return { ok: false };
      await fs.unlink(fileFor(String(id))).catch(() => {});
      return { ok: true };
    },
  );
}
