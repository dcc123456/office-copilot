import { app, type IpcMain } from "electron";
import path from "node:path";
import fs from "node:fs/promises";

/**
 * Modification log: one JSON file per document under userData/modlog, holding
 * every AI edit (what it changed, how to roll it back, and the block text before
 * and after so the history panel can render a diff without re-reading .docx).
 *
 * Records are capped; the oldest are pruned. Undo is the renderer's job — the
 * main process only stores opaque, already-validated JSON.
 */

const CAP = 60;

export interface ModLogRecord {
  id: string;
  ts: number;
  label: string;
  summary: string;
  [key: string]: unknown;
}

function slug(key: string): string {
  const s = key
    .replace(/[\\/]/g, "_")
    .replace(/[^a-zA-Z0-9._一-龥-]/g, "_")
    .slice(-120);
  return s || "untitled";
}

async function root(): Promise<string> {
  const dir = path.join(app.getPath("userData"), "modlog");
  await fs.mkdir(dir, { recursive: true });
  return dir;
}

async function fileFor(docKey: string): Promise<string> {
  return path.join(await root(), `${slug(docKey)}.json`);
}

async function read(docKey: string): Promise<ModLogRecord[]> {
  const raw = await fs.readFile(await fileFor(docKey), "utf8").catch(() => "[]");
  try {
    const parsed: unknown = JSON.parse(raw);
    return Array.isArray(parsed) ? (parsed as ModLogRecord[]) : [];
  } catch {
    return [];
  }
}

async function write(docKey: string, records: ModLogRecord[]): Promise<void> {
  const file = await fileFor(docKey);
  const tmp = `${file}.tmp`;
  await fs.writeFile(tmp, JSON.stringify(records), "utf8");
  await fs.rename(tmp, file);
}

export function registerModlogHandlers(ipc: IpcMain): void {
  ipc.handle("modlog:list", async (_e, docKey: string): Promise<ModLogRecord[]> => {
    const records = await read(String(docKey ?? ""));
    return records.sort((a, b) => b.ts - a.ts);
  });

  ipc.handle(
    "modlog:append",
    async (_e, payload: { docKey: string; record: ModLogRecord }): Promise<{ ok: boolean; message?: string }> => {
      try {
        const docKey = String(payload.docKey ?? "");
        const records = await read(docKey);
        records.push(payload.record);
        await write(docKey, records.slice(-CAP));
        return { ok: true };
      } catch (err) {
        return { ok: false, message: err instanceof Error ? err.message : String(err) };
      }
    },
  );

  ipc.handle(
    "modlog:patch",
    async (
      _e,
      payload: { docKey: string; id: string; patch: Record<string, unknown> },
    ): Promise<{ ok: boolean; message?: string }> => {
      try {
        const docKey = String(payload.docKey ?? "");
        const records = await read(docKey);
        const i = records.findIndex((r) => r.id === payload.id);
        if (i < 0) return { ok: false, message: "记录不存在。" };
        const existing = records[i]!;
        records[i] = { ...existing, ...payload.patch, id: existing.id, ts: existing.ts, label: existing.label, summary: existing.summary };
        await write(docKey, records);
        return { ok: true };
      } catch (err) {
        return { ok: false, message: err instanceof Error ? err.message : String(err) };
      }
    },
  );

  ipc.handle(
    "modlog:remove",
    async (_e, payload: { docKey: string; id?: string }): Promise<{ ok: boolean }> => {
      const docKey = String(payload.docKey ?? "");
      if (!payload.id) {
        await fs.unlink(await fileFor(docKey)).catch(() => {});
        return { ok: true };
      }
      const records = await read(docKey);
      await write(docKey, records.filter((r) => r.id !== payload.id));
      return { ok: true };
    },
  );
}
