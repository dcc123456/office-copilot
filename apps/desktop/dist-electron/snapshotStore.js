"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.registerSnapshotHandlers = registerSnapshotHandlers;
const electron_1 = require("electron");
const node_path_1 = __importDefault(require("node:path"));
const promises_1 = __importDefault(require("node:fs/promises"));
/**
 * Automatic version snapshots. Before any AI-rewrite or "AI 排版" operation,
 * the renderer asks the main process to save a copy of the current .docx
 * under userData/versions/. The user can browse these from the menu later.
 *
 * Snapshots are capped at 50 per document; oldest are pruned. We never delete
 * the original file.
 */
async function versionsRoot() {
    const root = node_path_1.default.join(electron_1.app.getPath("userData"), "versions");
    await promises_1.default.mkdir(root, { recursive: true });
    return root;
}
function safeSlug(p) {
    return p
        .replace(/[\\/]/g, "_")
        .replace(/[^a-zA-Z0-9._\u4e00-\u9fa5-]/g, "_")
        .slice(-120);
}
function timestamp() {
    const d = new Date();
    const pad = (n) => String(n).padStart(2, "0");
    return (`${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}` +
        `-${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`);
}
function registerSnapshotHandlers(ipc) {
    ipc.handle("snapshot:save", async (_e, payload) => {
        try {
            const root = await versionsRoot();
            const slug = payload.originalPath
                ? safeSlug(node_path_1.default.basename(payload.originalPath))
                : "未命名";
            const dir = node_path_1.default.join(root, slug);
            await promises_1.default.mkdir(dir, { recursive: true });
            // Keep only the 50 newest snapshots.
            const entries = await promises_1.default.readdir(dir);
            if (entries.length >= 50) {
                const sorted = entries.sort();
                for (let i = 0; i < entries.length - 49; i++) {
                    await promises_1.default.unlink(node_path_1.default.join(dir, sorted[i])).catch(() => { });
                }
            }
            const ext = payload.ext ?? "docx";
            const file = node_path_1.default.join(dir, `${timestamp()}${payload.reason ? "-" + safeSlug(payload.reason) : ""}.${ext}`);
            const buf = payload.bytes instanceof ArrayBuffer
                ? Buffer.from(payload.bytes)
                : Buffer.from(payload.bytes);
            await promises_1.default.writeFile(file, buf);
            return { ok: true, path: file };
        }
        catch (err) {
            return {
                ok: false,
                message: err instanceof Error ? err.message : String(err),
            };
        }
    });
    ipc.handle("snapshot:read", async (_e, payload) => {
        try {
            // Basic safety: only allow reading from inside the versions root.
            const root = await versionsRoot();
            const resolved = node_path_1.default.resolve(payload.path);
            if (!resolved.startsWith(node_path_1.default.resolve(root))) {
                return { ok: false, message: "路径不在快照目录内。" };
            }
            const buf = await promises_1.default.readFile(resolved);
            // Copy into an ArrayBuffer so it crosses the IPC structured-clone
            // boundary cleanly (Buffer is also fine, but this is explicit).
            return { ok: true, bytes: buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) };
        }
        catch (err) {
            return {
                ok: false,
                message: err instanceof Error ? err.message : String(err),
            };
        }
    });
    ipc.handle("snapshot:list", async (_e, payload) => {
        try {
            const root = await versionsRoot();
            const slug = payload.originalPath
                ? safeSlug(node_path_1.default.basename(payload.originalPath))
                : "未命名";
            const dir = node_path_1.default.join(root, slug);
            const entries = await promises_1.default.readdir(dir).catch(() => []);
            const out = [];
            for (const name of entries) {
                if (!/\.(docx|json)$/i.test(name))
                    continue;
                const p = node_path_1.default.join(dir, name);
                const st = await promises_1.default.stat(p).catch(() => null);
                if (st)
                    out.push({ path: p, mtime: st.mtimeMs, size: st.size });
            }
            return out.sort((a, b) => b.mtime - a.mtime);
        }
        catch {
            return [];
        }
    });
}
//# sourceMappingURL=snapshotStore.js.map