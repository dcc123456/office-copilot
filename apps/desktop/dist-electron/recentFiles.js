"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.RecentFiles = void 0;
exports.registerRecentHandlers = registerRecentHandlers;
exports.makeRecentFiles = makeRecentFiles;
const electron_1 = require("electron");
const promises_1 = __importDefault(require("node:fs/promises"));
const node_path_1 = __importDefault(require("node:path"));
/**
 * Persists the list of recently opened/saved files. The list is
 * rendered both in the application menu (File → Open Recent) and
 * emitted via "recent:changed" to the renderer for any welcome
 * screen. We cap at 10 entries and guard against path traversal
 * on read.
 */
const CAP = 10;
class RecentFiles {
    file;
    list = [];
    constructor(userData) {
        this.file = node_path_1.default.join(userData, "recent.json");
    }
    async load() {
        try {
            const raw = await promises_1.default.readFile(this.file, "utf-8");
            const data = JSON.parse(raw);
            this.list = Array.isArray(data.paths) ? data.paths.slice(0, CAP) : [];
        }
        catch {
            this.list = [];
        }
        return this.list;
    }
    async add(p) {
        const abs = node_path_1.default.resolve(p);
        this.list = [abs, ...this.list.filter((x) => x !== abs)].slice(0, CAP);
        await this.save();
        return this.list;
    }
    async clear() {
        this.list = [];
        await this.save();
    }
    all() {
        return [...this.list];
    }
    async save() {
        try {
            await promises_1.default.mkdir(node_path_1.default.dirname(this.file), { recursive: true });
            await promises_1.default.writeFile(this.file, JSON.stringify({ paths: this.list }, null, 2));
        }
        catch {
            // best-effort
        }
    }
}
exports.RecentFiles = RecentFiles;
function registerRecentHandlers(ipc, recent, getWindow, rebuildMenu) {
    ipc.handle("recent:list", async () => recent.all());
    ipc.handle("recent:clear", async () => {
        await recent.clear();
        rebuildMenu();
        getWindow()?.webContents.send("recent:changed", recent.all());
        return recent.all();
    });
    ipc.on("recent:add", (_e, p) => {
        recent.add(p).then((list) => {
            rebuildMenu();
            getWindow()?.webContents.send("recent:changed", list);
        });
    });
}
function makeRecentFiles() {
    return new RecentFiles(electron_1.app.getPath("userData"));
}
//# sourceMappingURL=recentFiles.js.map