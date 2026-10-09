import { app, type IpcMain } from "electron";
import path from "node:path";
import fs from "node:fs/promises";

/**
 * Automatic version snapshots. Before any AI-rewrite or "AI 排版" operation,
 * the renderer asks the main process to save a copy of the current .docx
 * under userData/versions/. The user can browse these from the menu later.
 *
 * Snapshots are capped at 50 per document; oldest are pruned. We never delete
 * the original file.
 */

async function versionsRoot(): Promise<string> {
  const root = path.join(app.getPath("userData"), "versions");
  await fs.mkdir(root, { recursive: true });
  return root;
}

function safeSlug(p: string): string {
  return p
    .replace(/[\\/]/g, "_")
    .replace(/[^a-zA-Z0-9._\u4e00-\u9fa5-]/g, "_")
    .slice(-120);
}

function timestamp(): string {
  const d = new Date();
  const pad = (n: number) => String(n).padStart(2, "0");
  return (
    `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}` +
    `-${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`
  );
}

export function registerSnapshotHandlers(ipc: IpcMain): void {
  ipc.handle(
    "snapshot:save",
    async (
      _e,
      payload: { originalPath: string | null; bytes: ArrayBuffer | Uint8Array; reason?: string; ext?: string },
    ): Promise<{ ok: true; path: string } | { ok: false; message: string }> => {
      try {
        const root = await versionsRoot();
        const slug = payload.originalPath
          ? safeSlug(path.basename(payload.originalPath))
          : "未命名";
        const dir = path.join(root, slug);
        await fs.mkdir(dir, { recursive: true });

        // Keep only the 50 newest snapshots.
        const entries = await fs.readdir(dir);
        if (entries.length >= 50) {
          const sorted = entries.sort();
          for (let i = 0; i < entries.length - 49; i++) {
            await fs.unlink(path.join(dir, sorted[i]!)).catch(() => {});
          }
        }

        const ext = payload.ext ?? "docx";
        const file = path.join(
          dir,
          `${timestamp()}${payload.reason ? "-" + safeSlug(payload.reason) : ""}.${ext}`,
        );
        const buf =
          payload.bytes instanceof ArrayBuffer
            ? Buffer.from(payload.bytes)
            : Buffer.from(payload.bytes as Uint8Array);
        await fs.writeFile(file, buf);
        return { ok: true, path: file };
      } catch (err) {
        return {
          ok: false,
          message: err instanceof Error ? err.message : String(err),
        };
      }
    },
  );

  ipc.handle(
    "snapshot:read",
    async (
      _e,
      payload: { path: string },
    ): Promise<
      | { ok: true; bytes: ArrayBuffer }
      | { ok: false; message: string }
    > => {
      try {
        // Basic safety: only allow reading from inside the versions root.
        const root = await versionsRoot();
        const resolved = path.resolve(payload.path);
        if (!resolved.startsWith(path.resolve(root))) {
          return { ok: false, message: "路径不在快照目录内。" };
        }
        const buf = await fs.readFile(resolved);
        // Copy into an ArrayBuffer so it crosses the IPC structured-clone
        // boundary cleanly (Buffer is also fine, but this is explicit).
        return { ok: true, bytes: buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) as ArrayBuffer };
      } catch (err) {
        return {
          ok: false,
          message: err instanceof Error ? err.message : String(err),
        };
      }
    },
  );

  ipc.handle(
    "snapshot:list",
    async (
      _e,
      payload: { originalPath: string | null },
    ): Promise<Array<{ path: string; mtime: number; size: number }>> => {
      try {
        const root = await versionsRoot();
        const slug = payload.originalPath
          ? safeSlug(path.basename(payload.originalPath))
          : "未命名";
        const dir = path.join(root, slug);
        const entries = await fs.readdir(dir).catch(() => [] as string[]);
        const out: Array<{ path: string; mtime: number; size: number }> = [];
        for (const name of entries) {
          if (!/\.(docx|json)$/i.test(name)) continue;
          const p = path.join(dir, name);
          const st = await fs.stat(p).catch(() => null);
          if (st) out.push({ path: p, mtime: st.mtimeMs, size: st.size });
        }
        return out.sort((a, b) => b.mtime - a.mtime);
      } catch {
        return [];
      }
    },
  );
}
