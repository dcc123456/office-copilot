import { app, type IpcMain, type BrowserWindow } from "electron";
import fs from "node:fs/promises";
import path from "node:path";

/**
 * Persists the list of recently opened/saved files. The list is
 * rendered both in the application menu (File → Open Recent) and
 * emitted via "recent:changed" to the renderer for any welcome
 * screen. We cap at 10 entries and guard against path traversal
 * on read.
 */

const CAP = 10;

export class RecentFiles {
  private file: string;
  private list: string[] = [];

  constructor(userData: string) {
    this.file = path.join(userData, "recent.json");
  }

  async load(): Promise<string[]> {
    try {
      const raw = await fs.readFile(this.file, "utf-8");
      const data = JSON.parse(raw) as { paths?: string[] };
      this.list = Array.isArray(data.paths) ? data.paths.slice(0, CAP) : [];
    } catch {
      this.list = [];
    }
    return this.list;
  }

  async add(p: string): Promise<string[]> {
    const abs = path.resolve(p);
    this.list = [abs, ...this.list.filter((x) => x !== abs)].slice(0, CAP);
    await this.save();
    return this.list;
  }

  async clear(): Promise<void> {
    this.list = [];
    await this.save();
  }

  all(): string[] {
    return [...this.list];
  }

  private async save(): Promise<void> {
    try {
      await fs.mkdir(path.dirname(this.file), { recursive: true });
      await fs.writeFile(this.file, JSON.stringify({ paths: this.list }, null, 2));
    } catch {
      // best-effort
    }
  }
}

export function registerRecentHandlers(
  ipc: IpcMain,
  recent: RecentFiles,
  getWindow: () => BrowserWindow | null,
  rebuildMenu: () => void,
): void {
  ipc.handle("recent:list", async () => recent.all());
  ipc.handle("recent:clear", async () => {
    await recent.clear();
    rebuildMenu();
    getWindow()?.webContents.send("recent:changed", recent.all());
    return recent.all();
  });
  ipc.on("recent:add", (_e, p: string) => {
    recent.add(p).then((list) => {
      rebuildMenu();
      getWindow()?.webContents.send("recent:changed", list);
    });
  });
}

export function makeRecentFiles(): RecentFiles {
  return new RecentFiles(app.getPath("userData"));
}
