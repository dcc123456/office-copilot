import { app, BrowserWindow, Menu, shell, ipcMain, clipboard } from "electron";
import path from "node:path";
import { registerFileHandlers } from "./fileHandlers.js";
import { registerSheetHandlers } from "./sheetHandlers.js";
import { registerPptxHandlers } from "./pptxHandlers.js";
import { registerAIHandlers } from "./aiHandlers.js";
import { registerSettingsHandlers } from "./settingsStore.js";
import { registerSnapshotHandlers } from "./snapshotStore.js";
import { registerModlogHandlers } from "./modlog.js";
import { registerSessionStoreHandlers } from "./sessionStore.js";
import { registerPrintHandlers } from "./printHandlers.js";
import { makeRecentFiles, registerRecentHandlers } from "./recentFiles.js";

function loadDotenv(): void {
  try {
    // eslint-disable-next-line @typescript-eslint/no-var-requires
    const dotenv = require("dotenv") as typeof import("dotenv");
    const candidates = [
      path.resolve(process.cwd(), ".env"),
      path.resolve(__dirname, "../../..", ".env"),
      path.resolve(path.dirname(app.getPath("exe")), "..", ".env"),
    ];
    for (const p of candidates) dotenv.config({ path: p });
  } catch {
    // optional
  }
}
loadDotenv();

const isDev = !!process.env.VITE_DEV_SERVER_URL;
let mainWindow: BrowserWindow | null = null;
const recent = makeRecentFiles();

// ---- 从 Finder「打开方式」/ 双击 / 拖到 Dock 打开文档 ----
// macOS 走 open-file 事件，Windows/Linux 把路径塞在 argv 里。
// 冷启动时事件可能早于 whenReady，所以先暂存，等窗口就绪后由渲染进程来取。
// 渲染进程按扩展名分派：docx 走 fileHandlers，xlsx 走 sheetHandlers。
const OPENABLE_EXT = /\.(docx|xlsx)$/i;
let pendingOpenPath: string | null = null;

function openableFromArgv(argv: string[]): string | null {
  return argv.slice(1).find((a) => !a.startsWith("-") && OPENABLE_EXT.test(a)) ?? null;
}

function queueOpenPath(p: string | null | undefined): void {
  if (!p || !OPENABLE_EXT.test(p)) return;
  pendingOpenPath = p;
  flushOpenPath();
}

/** 把暂存的路径推给渲染进程；窗口还没加载完就等 did-finish-load。 */
function flushOpenPath(): void {
  if (!pendingOpenPath || !mainWindow) return;
  const p = pendingOpenPath;
  const wc = mainWindow.webContents;
  const send = (): void => {
    pendingOpenPath = null;
    if (!mainWindow) return;
    if (mainWindow.isMinimized()) mainWindow.restore();
    mainWindow.show();
    mainWindow.focus();
    wc.send("menu:open-recent", p);
  };
  if (wc.isLoading()) wc.once("did-finish-load", send);
  else send();
}

app.on("open-file", (event, filePath) => {
  event.preventDefault();
  queueOpenPath(filePath);
});

// 单实例：已经开着的时候再打开一份文档，复用现有窗口而不是再起一个。
if (!app.requestSingleInstanceLock()) {
  app.quit();
} else {
  app.on("second-instance", (_event, argv) => {
    const p = openableFromArgv(argv);
    if (p) {
      queueOpenPath(p);
      return;
    }
    if (mainWindow) {
      if (mainWindow.isMinimized()) mainWindow.restore();
      mainWindow.show();
      mainWindow.focus();
    }
  });
}

function createWindow(): BrowserWindow {
  const win = new BrowserWindow({
    width: 1280,
    height: 820,
    minWidth: 960,
    minHeight: 600,
    title: "Office Copilot",
    titleBarStyle: "hiddenInset",
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      sandbox: false,
      contextIsolation: true,
      nodeIntegration: false,
    },
  });

  if (isDev && process.env.VITE_DEV_SERVER_URL) {
    win.loadURL(process.env.VITE_DEV_SERVER_URL);
  } else {
    win.loadFile(path.join(__dirname, "../dist/index.html"));
  }

  win.webContents.setWindowOpenHandler(({ url }) => {
    shell.openExternal(url);
    return { action: "deny" };
  });

  return win;
}

function buildMenu(): Menu {
  const isMac = process.platform === "darwin";
  const recentPaths = recent.all();
  const recentMenu: Electron.MenuItemConstructorOptions[] =
    recentPaths.length > 0
      ? [
          { type: "separator" },
          { label: "最近打开", enabled: false },
          ...recentPaths.slice(0, 10).map(
            (p): Electron.MenuItemConstructorOptions => ({
              label: p.split(/[\\/]/).pop() ?? p,
              click: () => mainWindow?.webContents.send("menu:open-recent", p),
            }),
          ),
          {
            label: "清除最近记录",
            click: () => {
              recent.clear().then(() => {
                rebuildMenu();
                mainWindow?.webContents.send("recent:changed", recent.all());
              });
            },
          },
        ]
      : [];

  const template: Electron.MenuItemConstructorOptions[] = [
    ...(isMac
      ? ([
          {
            label: "Office Copilot",
            submenu: [
              { role: "about", label: "关于 Office Copilot" },
              { type: "separator" },
              { role: "hide", label: "隐藏 Office Copilot" },
              { role: "hideOthers", label: "隐藏其他" },
              { role: "unhide", label: "全部显示" },
              { type: "separator" },
              { role: "quit", label: "退出 Office Copilot" },
            ],
          },
        ] as Electron.MenuItemConstructorOptions[])
      : []),
    {
      label: "文件",
      submenu: [
        {
          label: "新建",
          accelerator: "CmdOrCtrl+N",
          click: () => mainWindow?.webContents.send("menu:new"),
        },
        {
          label: "打开…",
          accelerator: "CmdOrCtrl+O",
          click: () => mainWindow?.webContents.send("menu:open"),
        },
        ...recentMenu,
        { type: "separator" },
        {
          label: "保存",
          accelerator: "CmdOrCtrl+S",
          click: () => mainWindow?.webContents.send("menu:save"),
        },
        {
          label: "另存为…",
          accelerator: "CmdOrCtrl+Shift+S",
          click: () => mainWindow?.webContents.send("menu:save-as"),
        },
        { type: "separator" },
        {
          label: "打印 / 导出 PDF…",
          accelerator: "CmdOrCtrl+P",
          click: () => mainWindow?.webContents.send("menu:print"),
        },
        { type: "separator" },
        isMac ? { role: "close" } : { role: "quit" },
      ],
    },
    {
      label: "编辑",
      submenu: [
        { role: "undo", label: "撤销" },
        { role: "redo", label: "重做" },
        { type: "separator" },
        { role: "cut", label: "剪切" },
        { role: "copy", label: "复制" },
        { role: "paste", label: "粘贴" },
        { role: "selectAll", label: "全选" },
        { type: "separator" },
        {
          label: "查找和替换…",
          accelerator: "CmdOrCtrl+F",
          click: () => mainWindow?.webContents.send("menu:find"),
        },
      ],
    },
    {
      label: "AI",
      submenu: [
        {
          label: "扩写（选中/全文）",
          accelerator: "Alt+CmdOrCtrl+J",
          click: () => mainWindow?.webContents.send("ai:expand"),
        },
        {
          label: "缩写（选中/全文）",
          accelerator: "Alt+CmdOrCtrl+U",
          click: () => mainWindow?.webContents.send("ai:shorten"),
        },
        {
          label: "润色（选中/全文）",
          accelerator: "Alt+CmdOrCtrl+I",
          click: () => mainWindow?.webContents.send("ai:polish"),
        },
        { type: "separator" },
        {
          label: "AI 排版全文",
          accelerator: "Alt+CmdOrCtrl+Shift+F",
          click: () => mainWindow?.webContents.send("ai:format-document"),
        },
        {
          label: "美化排版…",
          accelerator: "Alt+CmdOrCtrl+Shift+B",
          click: () => mainWindow?.webContents.send("ai:beautify-document"),
        },
        { type: "separator" },
        {
          label: "切换 AI 侧边栏",
          accelerator: "Alt+CmdOrCtrl+L",
          click: () => mainWindow?.webContents.send("ai:toggle-panel"),
        },
      ],
    },
    {
      label: "视图",
      submenu: [
        {
          label: "切换大纲侧栏",
          accelerator: "Alt+CmdOrCtrl+K",
          click: () => mainWindow?.webContents.send("outline:toggle"),
        },
        { type: "separator" },
        { role: "reload", label: "重新载入" },
        { role: "toggleDevTools", label: "开发者工具" },
        { type: "separator" },
        { role: "resetZoom", label: "实际大小" },
        { role: "zoomIn", label: "放大" },
        { role: "zoomOut", label: "缩小" },
        { type: "separator" },
        { role: "togglefullscreen", label: "进入全屏幕" },
      ],
    },
  ];
  return Menu.buildFromTemplate(template);
}

function rebuildMenu(): void {
  Menu.setApplicationMenu(buildMenu());
}

app.whenReady().then(async () => {
  await recent.load();
  rebuildMenu();
  mainWindow = createWindow();
  registerFileHandlers(ipcMain, () => mainWindow);
  registerSheetHandlers(ipcMain, () => mainWindow);
  registerPptxHandlers(ipcMain, () => mainWindow);
  registerAIHandlers(ipcMain);
  registerSettingsHandlers(ipcMain);
  registerSnapshotHandlers(ipcMain);
  registerModlogHandlers(ipcMain);
  registerSessionStoreHandlers(ipcMain);
  registerPrintHandlers(ipcMain);
  registerRecentHandlers(ipcMain, recent, () => mainWindow, rebuildMenu);

  // 渲染进程挂载后会来取冷启动时暂存的待打开文件（避免 push 的时序竞争）。
  ipcMain.handle("file:take-pending-open", () => {
    const p = pendingOpenPath;
    pendingOpenPath = null;
    return p;
  });

  // 右键菜单要按"剪贴板里有没有东西"决定粘贴是否置灰，并支持三种粘贴方式。
  // 剪贴板属于系统能力，一律在主进程读，渲染层只经 preload 的白名单访问。
  ipcMain.handle("clipboard:read", () => ({
    text: clipboard.readText(),
    html: clipboard.readHTML(),
  }));

  // 窗口边缘拖拽缩放：渲染层热区条发「起始 bounds + 累计位移 + 边标识」，
  // 边语义（哪侧被拖动、原点是否跟随）在这里统一处理，并夹到最小尺寸。
  ipcMain.handle("window:get-bounds", (e) => {
    const win = BrowserWindow.fromWebContents(e.sender);
    if (!win) return { x: 0, y: 0, width: 1280, height: 820 };
    const b = win.getBounds();
    return { x: b.x, y: b.y, width: b.width, height: b.height };
  });
  ipcMain.on(
    "window:drag-resize",
    (
      e,
      payload: {
        base: { x: number; y: number; width: number; height: number };
        dx: number;
        dy: number;
        edge: string;
      },
    ) => {
      const win = BrowserWindow.fromWebContents(e.sender);
      if (!win || win.isMaximized() || win.isSimpleFullScreen()) return;
      const [minW = 0, minH = 0] = win.getMinimumSize();
      const { base, dx, dy } = payload;
      const right = payload.edge.includes("e");
      const left = payload.edge.includes("w");
      const down = payload.edge.includes("s");
      const up = payload.edge.includes("n");
      let width = base.width + (right ? dx : 0) - (left ? dx : 0);
      let height = base.height + (down ? dy : 0) - (up ? dy : 0);
      let x = base.x + (left ? dx : 0);
      let y = base.y + (up ? dy : 0);
      if (width < minW) {
        if (left) x = base.x + base.width - minW;
        width = minW;
      }
      if (height < minH) {
        if (up) y = base.y + base.height - minH;
        height = minH;
      }
      win.setBounds({ x: Math.round(x), y: Math.round(y), width: Math.round(width), height: Math.round(height) });
    },
  );

  // 冷启动：Windows/Linux 双击文件时路径在 argv 里；macOS 走上面的 open-file。
  queueOpenPath(openableFromArgv(process.argv));
  flushOpenPath();

  // Record every successful open/save into Recent.
  ipcMain.on("file:opened", (_e, p: string) => {
    if (typeof p === "string" && p) {
      recent.add(p).then((list) => {
        rebuildMenu();
        mainWindow?.webContents.send("recent:changed", list);
      });
    }
  });

  app.on("activate", () => {
    if (BrowserWindow.getAllWindows().length === 0) mainWindow = createWindow();
  });
});

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") app.quit();
});

process.on("unhandledRejection", (reason) => {
  // eslint-disable-next-line no-console
  console.error("Unhandled rejection:", reason);
});
