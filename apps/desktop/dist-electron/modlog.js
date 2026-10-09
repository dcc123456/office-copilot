"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.registerModlogHandlers = registerModlogHandlers;
const electron_1 = require("electron");
const node_path_1 = __importDefault(require("node:path"));
const promises_1 = __importDefault(require("node:fs/promises"));
/**
 * Modification log: one JSON file per document under userData/modlog, holding
 * every AI edit (what it changed, how to roll it back, and the block text before
 * and after so the history panel can render a diff without re-reading .docx).
 *
 * Records are capped; the oldest are pruned. Undo is the renderer's job — the
 * main process only stores opaque, already-validated JSON.
 */
const CAP = 60;
function slug(key) {
    const s = key
        .replace(/[\\/]/g, "_")
        .replace(/[^a-zA-Z0-9._一-龥-]/g, "_")
        .slice(-120);
    return s || "untitled";
}
async function root() {
    const dir = node_path_1.default.join(electron_1.app.getPath("userData"), "modlog");
    await promises_1.default.mkdir(dir, { recursive: true });
    return dir;
}
async function fileFor(docKey) {
    return node_path_1.default.join(await root(), `${slug(docKey)}.json`);
}
async function read(docKey) {
    const raw = await promises_1.default.readFile(await fileFor(docKey), "utf8").catch(() => "[]");
    try {
        const parsed = JSON.parse(raw);
        return Array.isArray(parsed) ? parsed : [];
    }
    catch {
        return [];
    }
}
async function write(docKey, records) {
    const file = await fileFor(docKey);
    const tmp = `${file}.tmp`;
    await promises_1.default.writeFile(tmp, JSON.stringify(records), "utf8");
    await promises_1.default.rename(tmp, file);
}
function registerModlogHandlers(ipc) {
    ipc.handle("modlog:list", async (_e, docKey) => {
        const records = await read(String(docKey ?? ""));
        return records.sort((a, b) => b.ts - a.ts);
    });
    ipc.handle("modlog:append", async (_e, payload) => {
        try {
            const docKey = String(payload.docKey ?? "");
            const records = await read(docKey);
            records.push(payload.record);
            await write(docKey, records.slice(-CAP));
            return { ok: true };
        }
        catch (err) {
            return { ok: false, message: err instanceof Error ? err.message : String(err) };
        }
    });
    ipc.handle("modlog:patch", async (_e, payload) => {
        try {
            const docKey = String(payload.docKey ?? "");
            const records = await read(docKey);
            const i = records.findIndex((r) => r.id === payload.id);
            if (i < 0)
                return { ok: false, message: "记录不存在。" };
            const existing = records[i];
            records[i] = { ...existing, ...payload.patch, id: existing.id, ts: existing.ts, label: existing.label, summary: existing.summary };
            await write(docKey, records);
            return { ok: true };
        }
        catch (err) {
            return { ok: false, message: err instanceof Error ? err.message : String(err) };
        }
    });
    ipc.handle("modlog:remove", async (_e, payload) => {
        const docKey = String(payload.docKey ?? "");
        if (!payload.id) {
            await promises_1.default.unlink(await fileFor(docKey)).catch(() => { });
            return { ok: true };
        }
        const records = await read(docKey);
        await write(docKey, records.filter((r) => r.id !== payload.id));
        return { ok: true };
    });
}
//# sourceMappingURL=modlog.js.map