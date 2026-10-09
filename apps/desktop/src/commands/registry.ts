import type { LucideIcon } from "lucide-react";
import {
  FilePlus2,
  FolderOpen,
  Save,
  Download,
  Printer,
  Settings2,
  Search,
  PanelLeft,
  Sparkles,
  Palette,
  History,
  Wand2,
  SpellCheck,
  Table2,
  Workflow,
} from "lucide-react";
import type { ThemeId } from "../uiStore.js";

export interface CommandContext {
  handleNew: () => void;
  handleOpen: () => void;
  handleSave: () => void;
  handleSaveAs: () => void;
  handlePrint: () => void;
  openPageSetup: () => void;
  showFindBar: () => void;
  toggleOutline: () => void;
  toggleAI: () => void;
  openHistory: () => void;
  openSettings: () => void;
  runAI: (action: string) => void;
  formatDocument: () => void;
  openTemplatePicker: () => void;
  setTheme: (t: ThemeId) => void;
}

export interface Command {
  id: string;
  title: string;
  /** Extra match text (English names, aliases). */
  keywords?: string;
  icon?: LucideIcon;
  group: string;
  shortcut?: string;
  run: (ctx: CommandContext) => void;
}

const THEME_COMMANDS: { id: ThemeId; label: string }[] = [
  { id: "light", label: "素白" },
  { id: "dark", label: "夜墨" },
  { id: "spring", label: "春·青野" },
  { id: "summer", label: "夏·陶阳" },
  { id: "autumn", label: "秋·琥珀" },
  { id: "winter", label: "冬·霜蓝" },
];

export const COMMANDS: Command[] = [
  { id: "file.new", title: "新建文档", keywords: "new document", icon: FilePlus2, group: "文件", shortcut: "⌘N", run: (c) => c.handleNew() },
  { id: "file.open", title: "打开文档…", keywords: "open browse", icon: FolderOpen, group: "文件", shortcut: "⌘O", run: (c) => c.handleOpen() },
  { id: "file.save", title: "保存", keywords: "save", icon: Save, group: "文件", shortcut: "⌘S", run: (c) => c.handleSave() },
  { id: "file.save-as", title: "另存为…", keywords: "save as", icon: Download, group: "文件", shortcut: "⌘⇧S", run: (c) => c.handleSaveAs() },
  { id: "file.print", title: "打印 / 导出 PDF", keywords: "print pdf export", icon: Printer, group: "文件", shortcut: "⌘P", run: (c) => c.handlePrint() },
  { id: "file.page-setup", title: "页面设置…", keywords: "page setup margin", icon: Settings2, group: "文件", run: (c) => c.openPageSetup() },

  { id: "edit.find", title: "查找与替换", keywords: "find replace search", icon: Search, group: "编辑", shortcut: "⌘F", run: (c) => c.showFindBar() },

  { id: "view.outline", title: "显示/隐藏大纲", keywords: "outline sidebar toc", icon: PanelLeft, group: "视图", shortcut: "⌘B", run: (c) => c.toggleOutline() },
  { id: "view.ai", title: "显示/隐藏 AI 助手", keywords: "ai assistant panel", icon: Sparkles, group: "视图", shortcut: "⌥⌘L", run: (c) => c.toggleAI() },
  { id: "view.history", title: "历史版本…", keywords: "history snapshot versions", icon: History, group: "视图", run: (c) => c.openHistory() },
  { id: "view.settings", title: "模型设置…", keywords: "settings provider model api key", icon: Settings2, group: "视图", run: (c) => c.openSettings() },

  { id: "ai.expand", title: "AI：扩写选中文字", keywords: "expand elaborate", icon: Sparkles, group: "AI", run: (c) => c.runAI("扩写") },
  { id: "ai.shorten", title: "AI：缩写选中文字", keywords: "shorten condense", icon: Sparkles, group: "AI", run: (c) => c.runAI("缩写") },
  { id: "ai.polish", title: "AI：润色选中文字", keywords: "polish refine", icon: Sparkles, group: "AI", run: (c) => c.runAI("润色") },
  { id: "ai.typos", title: "AI：纠正全文错别字", keywords: "typo proofread spell check", icon: SpellCheck, group: "AI", run: (c) => c.runAI("纠正错别字") },
  { id: "ai.table", title: "AI：插入汇总表格", keywords: "insert table", icon: Table2, group: "AI", run: (c) => c.runAI("插入表格") },
  { id: "ai.diagram", title: "AI：插入示意图", keywords: "insert diagram flowchart canvas", icon: Workflow, group: "AI", run: (c) => c.runAI("插入示意图") },
  { id: "ai.format", title: "AI：排版全文", keywords: "format document typography", icon: Wand2, group: "AI", shortcut: "⌥⌘⇧F", run: (c) => c.formatDocument() },
  { id: "ai.beautify", title: "美化排版…（选模板）", keywords: "beautify template typography 排版 模板 公文 论文", icon: Wand2, group: "AI", shortcut: "⌥⌘⇧B", run: (c) => c.openTemplatePicker() },

  ...THEME_COMMANDS.map((t) => ({
    id: `theme.${t.id}`,
    title: `切换主题：${t.label}`,
    keywords: `theme ${t.id}`,
    icon: Palette,
    group: "主题",
    run: (c: CommandContext) => c.setTheme(t.id),
  })),
];

/** Case-insensitive subsequence match; higher = better. null = no match. */
export function fuzzyScore(query: string, text: string): number | null {
  const q = query.toLowerCase();
  if (!q) return 0;
  const t = text.toLowerCase();
  let ti = 0;
  let score = 0;
  let streak = 0;
  let matched = 0;
  for (const ch of q) {
    const idx = t.indexOf(ch, ti);
    if (idx === -1) return null;
    matched += 1;
    if (idx === ti) {
      streak += 1;
      score += 2 + streak;
    } else {
      streak = 0;
      score += 1;
      // word-boundary bonus
      if (idx === 0 || /[\s:：·/_-]/.test(t[idx - 1] ?? " ")) score += 2;
    }
    ti = idx + 1;
  }
  if (t.startsWith(q)) score += 10;
  return score - (matched < q.length ? 100 : 0);
}

export function rankCommands(
  commands: Command[],
  query: string,
  usage: Record<string, number>,
): Command[] {
  const scored: Array<{ cmd: Command; score: number }> = [];
  for (const cmd of commands) {
    const s = query
      ? Math.max(
          fuzzyScore(query, cmd.title) ?? -Infinity,
          cmd.keywords ? fuzzyScore(query, cmd.keywords) ?? -Infinity : -Infinity,
        )
      : 0;
    if (query && !isFinite(s)) continue;
    scored.push({ cmd, score: s + Math.min(usage[cmd.id] ?? 0, 10) });
  }
  if (query) {
    scored.sort((a, b) => b.score - a.score);
  } else {
    scored.sort(
      (a, b) => (usage[b.cmd.id] ?? 0) - (usage[a.cmd.id] ?? 0),
    );
  }
  return scored.map((x) => x.cmd);
}
