import type { Editor } from "@tiptap/react";
import { useCallback, useEffect, useRef, useState } from "react";
import {
  Settings2,
  ChevronDown,
  PanelRightClose,
  PanelRightOpen,
  Sparkles,
  Info,
  Copy,
  Check,
  X,
  History,
  Wand2,
  Square,
  Loader2,
  PencilLine,
  Table2,
  Workflow,
  SpellCheck,
  MessageSquare,
  MessageSquarePlus,
  MessagesSquare,
  Trash2,
  Zap,
  Play,
} from "lucide-react";
import { SettingsModal } from "../components/SettingsModal.js";
import { DiffView, InlineDiff } from "./DiffView.js";
import {
  alignRewrite,
  describeCoreOp,
  rewritableBlocks,
  rewriteResume,
  type Block,
  type CoreOp,
  type Proposal,
} from "./docEdit.js";
import { blocksOf } from "../editor/docBlocks.js";
import {
  aiStreamTransform,
  applyProposal,
  commitOps,
  commitSelectionRewrite,
  requestProposal,
  rewriteInstruction,
  simulateProposal,
  streamedParagraphs,
  streamedSummary,
} from "./agent.js";
import { toast } from "../components/ui/Toast.js";
import { useAnchoredMenu } from "../hooks/useAnchoredMenu.js";
import { runSheetAgent } from "./sheetAgent.js";
import { formatA1 } from "@office-copilot/sheet-io";
import type { ChartKind, ResolvedChart, SheetCoreOp, SheetOpsResult, SheetPayload } from "@office-copilot/sheet-io";
import type { SheetEditT } from "@office-copilot/sheet-tools";
import type {
  AppSettings,
  ProviderConfig,
  SessionSummary,
  SessionTurn,
  ToolTurnMessage,
} from "../global.js";
import {
  listSessions,
  newSessionId,
  openSession,
  removeSession,
  saveSession,
} from "./sessions.js";
import "../styles/aiPanel.css";

/* ------------------------------------------------------------------ */
/* Thread model                                                         */
/* ------------------------------------------------------------------ */

type CardStatus = "streaming" | "pending" | "applying" | "applied" | "failed" | "dismissed";

/** 聊天：只给建议；半自动：改动需确认；全自动：直接写入文档。 */
type AiMode = "chat" | "semi" | "auto";

interface PlanCard {
  summary: string;
  ops: CoreOp[];
  before: Block[];
  after: Block[];
  inverses: CoreOp[];
  notes: string[];
  status: CardStatus;
  message?: string;
  /** 流式阶段模型已经吐出的文字，卡片边收边滚动显示。 */
  raw?: string;
  /** 用户按了暂停：卡片给「继续」，不当失败处理。 */
  stopped?: boolean;
  /** 续写要送回原指令，以及上一次收到的改动（CoreOp 回不到 AiOp，只能留这份）。 */
  instruction?: string;
  proposal?: Proposal;
}

interface RewriteCard {
  /** selection：只改选区那段；document：没有选区时改整篇。 */
  scope: "selection" | "document";
  /** 暂停后续写要用同一条内置语气指令。 */
  action: string;
  from: number;
  to: number;
  beforeText: string;
  rewritten: string;
  status: CardStatus;
  message?: string;
  stopped?: boolean;
}

/**
 * 表格提案卡：模型交出的改动条目 + 引擎算好的可落地 op。
 *
 * 与文档那张的差别只在"没有逐字流式"：工具回合是一整轮一整轮回来的，
 * 中途能展示的只有"正在核对表格"，所以 `status` 停在 streaming 时不显示正文。
 */
interface SheetPlanCard {
  summary: string;
  /** 模型提的条目原样留着，卡片按它列改动清单（op 是引擎的表示，给人看的是这个）。 */
  edits: SheetEditT[];
  /** 能落地的部分；空数组意味着这份提案落不了地，卡片不会挂出来。 */
  ops: SheetCoreOp[];
  /** 没能变成 op 的条目原因，逐条给人看。 */
  notes: string[];
  /**
   * `sheet_create_chart` 解析过的图表；给了它这张卡就是图表卡（`ops` 必为空）。
   *
   * 图表不是 op：它随保存写进 `.xlsx`，撤销栈管不到，卡片要如实说这句（§2.5 的边界）。
   */
  chart?: ResolvedChart;
  status: CardStatus;
  message?: string;
  stopped?: boolean;
  /** 续跑要用同一条诉求，以及暂停时攒下的 transcript（spec 验收 B14）。 */
  instruction: string;
  messages: ToolTurnMessage[];
}

type Item =
  | {
      id: string;
      t: "msg";
      role: "user" | "assistant";
      text: string;
      pending?: boolean;
      /** 暂停留下的半截：气泡上给「继续」，并记住它是从哪句话生成的。 */
      stopped?: boolean;
      prompt?: string;
    }
  | { id: string; t: "plan"; label: string; card: PlanCard }
  | { id: string; t: "rewrite"; label: string; card: RewriteCard }
  | { id: string; t: "sheet"; label: string; card: SheetPlanCard };

/**
 * 提案条目 → 卡片上的一行中文说明。
 *
 * 引擎只输出结构化 op 与错误码，文案在渲染层组装（`AGENTS.md` §5.7）。
 *
 * @param edits `sheet_submit_proposal` 校验过的条目。
 * @returns 与条目同序的中文行；长文本截到 40 字，清单要能一眼扫完。
 */
function sheetEditLines(edits: readonly SheetEditT[]): string[] {
  return edits.map((edit) => {
    switch (edit.type) {
      case "set_cell": {
        const value = edit.text.length > 40 ? `${edit.text.slice(0, 40)}…` : edit.text;
        return `${edit.address.toUpperCase()} ← ${value || "（清空）"}`;
      }
      case "set_column_width":
        return `列 ${edit.column} 的宽度改为 ${Math.round(edit.widthPx)}px`;
      case "insert_rows":
        return `在第 ${edit.row} 行之前插入 ${edit.count} 行`;
      case "delete_rows":
        return `删掉第 ${[...edit.rows].sort((left, right) => left - right).join("、")} 行`;
    }
  });
}

/** 图表类型 → 卡片上那句中文。引擎只给 `bar/line/pie` 三个码（`AGENTS.md` §5.7）。 */
const CHART_KIND_LABEL: Record<ChartKind, string> = { bar: "柱状图", line: "折线图", pie: "饼图" };

/** 引擎解析过的图表 → 卡片上的一串中文行。
 *
 * @param chart `resolveChartSpec` 的产物：引用与数值都已从现场表读好。
 * @returns 首行说类型与类别轴，中间每条系列一行，末行给落点与尺寸（px，非 pt，§2.1）。
 */
function sheetChartLines(chart: ResolvedChart): string[] {
  return [
    `${CHART_KIND_LABEL[chart.kind]}：类别 ${chart.categoryRef}（${chart.categories.length} 项）`,
    ...chart.series.map((item) => `系列「${item.name}」← ${item.valueRef}（${item.values.length} 个数值）`),
    `落点 ${formatA1(chart.anchor)}，${Math.round(chart.widthPx)}×${Math.round(chart.heightPx)}px`,
    `标题：${chart.title || "（无）"}`,
  ];
}

const uid = () => crypto.randomUUID();

/** 每个新会话开头的那句引导语：不进会话正文，存盘时要跳过。 */
const WELCOME_ID = "welcome";

const WELCOME_ITEM: Item = {
  id: WELCOME_ID,
  t: "msg",
  role: "assistant",
  text:
    "选中文字后点快捷按钮就改那几段，不选中就是改整篇；也可以直接写下你的要求（比如「把第三节改成表格」「全文纠正错别字」）。默认是「半自动」：我先给 diff 预览，你确认后才写入文档；「全自动」直接改文档，「聊天」只给建议。每次改动都能在修改历史里看 diff 并单独撤销。",
};

/** 打开 .xlsx 时的引导语：说的是这张表能做什么，不是文档的选区改写。 */
const SHEET_WELCOME_ITEM: Item = {
  id: WELCOME_ID,
  t: "msg",
  role: "assistant",
  text:
    "直接说要对这张表做什么，比如「各区域销售额汇总」「找出销售额最高的三个月」「删掉重复的行」「把这一列拉宽」。数字由表格引擎算，不是我估的；能改的只有写格子、列宽、插行删行四类，每次改动都要你点「应用到表格」才落地，落地后能在修改历史里单独撤销。",
};

/** 会话历史浮层宽度（px），须与 `styles/aiPanel.css` 的 `.ai-session-pop` 一致。 */
const SESSION_POP_WIDTH = 250;

/**
 * 把面板上的一串会话记录折成"存盘的轮次"——会话历史只存文字，卡片存不下也不该存。
 *
 * 卡片按它当时的状态折成一句 assistant 的话：重开这个会话时看得懂当时发生了什么，
 * 但不会给一张点了会改坏当前文档的旧卡片（文档早就变了）。
 *
 * @param items 面板线程里的全部条目。
 * @returns 可持久化的轮次；引导语、还在流式的气泡和空文字都不算。
 */
function turnsOfItems(items: readonly Item[]): SessionTurn[] {
  const turns: SessionTurn[] = [];
  for (const it of items) {
    if (it.id === WELCOME_ID) continue;
    if (it.t === "msg") {
      if (it.pending || !it.text.trim()) continue;
      turns.push({ role: it.role, text: it.text });
      continue;
    }
    if (it.t === "plan") {
      const c = it.card;
      if (c.status === "streaming") continue;
      const result =
        c.status === "applied"
          ? "已写入文档"
          : c.status === "dismissed"
            ? "用户放弃了这个方案"
            : c.status === "failed"
              ? `未写入（${c.stopped ? "已暂停" : c.message ?? "失败"}）`
              : `${c.ops.length} 处改动待确认`;
      turns.push({ role: "assistant", text: `【${it.label}】${c.summary || "（无摘要）"} · ${result}` });
      continue;
    }
    if (it.t === "sheet") {
      const c = it.card;
      if (c.status === "streaming") continue;
      const result =
        c.status === "applied"
          ? "已写入表格"
          : c.status === "dismissed"
            ? "用户放弃了这份提案"
            : c.status === "failed"
              ? `未写入（${c.stopped ? "已暂停" : c.message ?? "失败"}）`
              : `${c.ops.length} 处改动待确认`;
      const lines = sheetEditLines(c.edits);
      turns.push({
        role: "assistant",
        text: `【${it.label}】${c.summary || "（无摘要）"} · ${result}${lines.length ? `\n· ${lines.join("\n· ")}` : ""}`,
      });
      continue;
    }
    const c = it.card;
    if (c.status === "streaming") continue;
    const result =
      c.status === "applied"
        ? "已写入文档"
        : c.status === "dismissed"
          ? "用户放弃了这次改写"
          : c.status === "failed"
            ? `未写入（${c.stopped ? "已暂停" : c.message ?? "失败"}）`
            : "改写结果待确认";
    turns.push({ role: "assistant", text: `【${it.label}】${result}\n${c.rewritten}` });
  }
  return turns;
}

const QUICK_ACTIONS = [
  { label: "扩写", action: "扩写" },
  { label: "缩写", action: "缩写" },
  { label: "润色", action: "润色" },
  { label: "更正式", action: "换语气:正式商务" },
  { label: "更口语", action: "换语气:口语化" },
  { label: "译成英文", action: "翻译:英文" },
];

const DOC_ACTION_CHIPS = [
  {
    label: "纠正错别字",
    icon: SpellCheck,
    instruction:
      "通读全文，纠正错别字、明显的标点误用和多字漏字。只修改确实有错的段落，其它文字一律保持不变。",
  },
  {
    label: "插入表格",
    icon: Table2,
    instruction:
      "根据文档内容，在最合适的位置插入一张表格来归纳要点（insertTable，第一行作表头）。内容要来自文档本身，不要编造数据；如果文档不足以整理成表格，就返回空的 ops 并说明原因。",
  },
  {
    label: "插入示意图",
    icon: Workflow,
    instruction:
      "为文档中最核心的一段说明画一张示意图（insertDiagram，flow-v 或 flow-h），节点用短句，放在该段之后，用来增强说明。若文档没有可画的流程或结构，就返回空的 ops 并说明原因。",
  },
];

const DOC_BY_LABEL = new Map(DOC_ACTION_CHIPS.map((d) => [d.label, d]));

/**
 * 表格模式下三种模式各说一句人话：这里没有逐字 diff，只有提案卡与"部分写入"。
 *
 * @see MODES —— 文档那套提示词，形状一样，说的是另一件事。
 */
const SHEET_MODE_HINTS: Record<AiMode, string> = {
  chat: "聊天模式：AI 只回答与给建议，不会改动表格。",
  semi: "半自动：AI 给出提案卡，你点「应用到表格」才写入，落地后能在修改历史里单独撤销。",
  auto: "全自动：AI 给出提案就直接写入表格，改动记录在修改历史里，可单独撤销。",
};

const MODES: Array<{ id: AiMode; label: string; icon: typeof MessageSquare; title: string; hint: string }> = [
  {
    id: "chat",
    label: "聊天",
    icon: MessageSquare,
    title: "只给修改建议和草稿，不会改动文档",
    hint: "聊天模式：AI 只提供建议，不会改动文档。",
  },
  {
    id: "semi",
    label: "半自动",
    icon: PencilLine,
    title: "AI 给出修改方案，你点「应用到文档」才写入",
    hint: "半自动：先给 diff 预览，确认后写入文档。",
  },
  {
    id: "auto",
    label: "全自动",
    icon: Zap,
    title: "AI 直接改文档，不需要逐次确认（每次改动仍可单独撤销）",
    hint: "全自动：AI 直接写入文档，改动记录在修改历史里，可单独撤销。",
  },
];

/* ------------------------------------------------------------------ */
/* Panel                                                               */
/* ------------------------------------------------------------------ */

/** 表格面板要接的东西：当前这张表的载荷（只读）与唯一的落地口。 */
export interface SheetPanelBridge {
  /** 当前工作表的载荷；代理只读它，写入一律经 `onCommitOps`（spec 决策 D6）。 */
  payload: SheetPayload;
  /**
   * 提案落地的回调（`App.tsx` 的 `commitSheetOps`）：核对现场、换进工作簿、记一条修改历史。
   *
   * @param ops 提案里能被现有 op 表达的部分。
   * @param summary 一句话说明，进修改历史。
   * @returns 落地结果，用于在卡上如实报"部分写入"；当前表已不在时返回 null。
   */
  onCommitOps: (ops: SheetCoreOp[], summary: string) => SheetOpsResult | null;
  /**
   * 图表落地的回调（`App.tsx` 的 `commitSheetChart`）：排进待写清单并置脏，真正的部件在保存时写入。
   *
   * @param chart 引擎解析过的图表，引用与数值都已核对。
   * @returns 当前这张表还在并成功排队时 true；表已经关了返回 false，卡片据此报"没排进去"。
   */
  onCommitChart: (chart: ResolvedChart) => boolean;
}

export function AIPanel({
  editor,
  docKey,
  sheet,
  collapsed,
  onToggle,
  onOpenHistory,
  onChanged,
}: {
  editor: Editor;
  docKey: string;
  /** 打开 .xlsx 时给；没有它就是文档模式，走原来那套选区改写与方案卡。 */
  sheet?: SheetPanelBridge;
  collapsed: boolean;
  onToggle: () => void;
  onOpenHistory?: () => void;
  onChanged: () => void;
}) {
  const [items, setItems] = useState<Item[]>([sheet ? SHEET_WELCOME_ITEM : WELCOME_ITEM]);
  const [input, setInput] = useState("");
  const [mode, setMode] = useState<AiMode>("semi");
  const modeRef = useRef<AiMode>("semi");
  const [chipsOpen, setChipsOpen] = useState(false);
  const [modeTabsOpen, setModeTabsOpen] = useState(false);
  const [modeHintOpen, setModeHintOpen] = useState(false);
  const [busy, setBusy] = useState(false);
  const [settingsOpen, setSettingsOpen] = useState(false);
  const [settings, setSettings] = useState<AppSettings | null>(null);
  const [providerMenuOpen, setProviderMenuOpen] = useState(false);
  const [activeProviderId, setActiveProviderId] = useState<string | null>(null);
  const [copiedId, setCopiedId] = useState<string | null>(null);
  const [sessions, setSessions] = useState<SessionSummary[]>([]);
  /** 会话历史浮层：面板有 `overflow: hidden`，只能用视口坐标挂浮层（AGENTS §5.5）。 */
  const {
    anchorRef: sessionsAnchorRef,
    rect: sessionsRect,
    open: sessionsOpen,
    toggle: toggleSessionsMenu,
    hide: hideSessionsMenu,
  } = useAnchoredMenu(SESSION_POP_WIDTH, "right");
  const listRef = useRef<HTMLDivElement>(null);
  const activeRequestRef = useRef<string | null>(null);
  const busyRef = useRef(false);
  /**
   * 暂停标志：表格的工具回合是整轮整轮回来的，没有可取消的流式请求，
   * 于是「暂停」只置这个位，代理在两轮请求之间读到它就停下（spec 验收 B14）。
   */
  const pausedRef = useRef(false);
  /** 当前会话的编号，同时是 userData/sessions 下的文件名。 */
  const sessionIdRef = useRef<string>(newSessionId());
  /** 最近一次存盘的会话内容：一样就不重复写盘，免得效果循环。 */
  const savedTurnsRef = useRef<string>("[]");
  /** `items` 的镜像：发请求要接上文、按条目 id 排除自己，闭包里的 items 都是旧的。 */
  const itemsRef = useRef<Item[]>(items);

  const setBusyBoth = useCallback((v: boolean) => {
    busyRef.current = v;
    setBusy(v);
  }, []);

  /** 新会话开头那句引导语：表格与文档说的不是一回事。 */
  const welcomeItem = sheet ? SHEET_WELCOME_ITEM : WELCOME_ITEM;

  const refreshSettings = useCallback(async () => {
    const s = (await window.officeCopilot.getSettings()) as AppSettings;
    setSettings(s);
    setActiveProviderId(
      s.defaultProviderId ?? s.providers.find((p) => p.enabled)?.id ?? null,
    );
  }, []);

  useEffect(() => {
    void refreshSettings();
  }, [refreshSettings]);

  useEffect(() => {
    listRef.current?.scrollTo({ top: listRef.current.scrollHeight });
  }, [items]);

  /** `items` 的镜像：发请求时要接上文、按条目 id 排除自己，闭包里的 `items` 是旧的。 */
  useEffect(() => {
    itemsRef.current = items;
  }, [items]);

  /**
   * 会话内容一变就落盘（去抖 600ms）：切会话、关窗口都不丢，
   * 又不会每来一个字就写一次文件。与磁盘上那份一致时直接跳过，避免自己触发自己。
   */
  useEffect(() => {
    const turns = turnsOfItems(items);
    const snapshot = JSON.stringify(turns);
    if (!turns.length || snapshot === savedTurnsRef.current) return;
    const timer = setTimeout(() => {
      savedTurnsRef.current = snapshot;
      void saveSession({ id: sessionIdRef.current, turns });
    }, 600);
    return () => clearTimeout(timer);
  }, [items]);

  const setModeBoth = useCallback((next: AiMode) => {
    modeRef.current = next;
    setMode(next);
  }, []);

  const push = useCallback((it: Item) => setItems((prev) => [...prev, it]), []);

  const patchCard = useCallback(
    (id: string, fn: (c: PlanCard | RewriteCard) => PlanCard | RewriteCard) => {
      setItems((prev) =>
        prev.map((it) =>
          it.id === id && (it.t === "plan" || it.t === "rewrite")
            ? ({ ...it, card: fn(it.card) } as Item)
            : it,
        ),
      );
    },
    [],
  );

  const note = useCallback(
    (text: string) => push({ id: uid(), t: "msg", role: "assistant", text }),
    [push],
  );

  /** 拉开会话列表时现读一次盘：省掉开面板就去撞磁盘的开销。 */
  const toggleSessions = useCallback(() => {
    toggleSessionsMenu();
    void listSessions().then(setSessions);
  }, [toggleSessionsMenu]);

  /**
   * 起一个新会话：当前这份先存好，线程清回引导语。
   *
   * 正在流式输出时不给点（`busy`），否则半截内容会存成上一个会话的尾巴。
   */
  const startSession = useCallback(async () => {
    if (busyRef.current) return;
    const turns = turnsOfItems(items);
    if (turns.length) await saveSession({ id: sessionIdRef.current, turns });
    sessionIdRef.current = newSessionId();
    savedTurnsRef.current = "[]";
    setItems([welcomeItem]);
    setSessions(await listSessions());
  }, [items, welcomeItem]);

  /**
   * 打开历史里的某个会话，接着聊。
   *
   * @param id 会话编号。
   */
  const loadSession = useCallback(async (id: string) => {
    if (busyRef.current) return;
    const record = await openSession(id);
    if (!record) {
      note("这个会话已经不在了，列表刷新一下。");
      setSessions(await listSessions());
      return;
    }
    sessionIdRef.current = record.id;
    savedTurnsRef.current = JSON.stringify(record.turns);
    setItems(
      record.turns.map<Item>((t) => ({ id: uid(), t: "msg", role: t.role, text: t.text })),
    );
    hideSessionsMenu();
  }, [note, hideSessionsMenu]);

  /** 删掉一个历史会话；删的正是当前会话时就另起一份新的。 */
  const dropSession = useCallback(async (id: string) => {
    await removeSession(id);
    if (id === sessionIdRef.current) {
      sessionIdRef.current = newSessionId();
      savedTurnsRef.current = "[]";
      setItems([welcomeItem]);
    }
    setSessions(await listSessions());
  }, [welcomeItem]);

  const selectionText = useCallback(() => {
    const { from, to } = editor.state.selection;
    return editor.state.doc.textBetween(from, to, "\n");
  }, [editor]);

  /* ---------------- agent: propose a document edit ---------------- */

  const applyPlan = useCallback(
    async (id: string, card: PlanCard) => {
      patchCard(id, (c) => ({ ...c, status: "applying" }));
      const r = await commitOps(editor, {
        ops: card.ops,
        before: card.before,
        after: card.after,
        inverses: card.inverses,
        label: "AI 修改",
        summary: card.summary,
        docKey,
      });
      if (!r.ok) {
        patchCard(id, (c) => ({ ...c, status: "failed", message: r.message }));
        return;
      }
      patchCard(id, (c) => ({ ...c, status: "applied", message: r.message }));
      onChanged();
      if (r.message) toast(r.message, "error");
    },
    [editor, docKey, onChanged, patchCard],
  );

  /** 从会话里摘掉一张卡片（用户放弃、或模型其实没要改任何东西）。 */
  const dismiss = useCallback(
    (id: string) => setItems((prev) => prev.filter((it) => it.id !== id)),
    [],
  );

  /**
   * 跑一次"改文档"的请求，结果落到指定或新建的方案卡上。
   *
   * @param opts.instruction 用户原话；续写时仍是原话，"接着上次写"由引擎再拼约束。
   * @param opts.label 卡片标题。
   * @param opts.cardId 复用已有卡片（点「继续」时），不给就新挂一张流式卡。
   * @param opts.continueFrom 上一次暂停留下的半截改动与原始输出。
   */
  const runAgent = useCallback(
    async (opts: {
      instruction: string;
      label: string;
      cardId?: string;
      continueFrom?: { ops: Proposal["ops"]; rawTail: string };
    }) => {
      const id = opts.cardId ?? uid();
      if (opts.cardId) {
        patchCard(id, (c) => ({ ...c, status: "streaming", raw: "", message: undefined, stopped: false }));
      } else {
        // 卡片先挂出来：流式阶段逐字显示模型的输出，收完才换成改动清单与 diff。
        push({
          id,
          t: "plan",
          label: opts.label,
          card: { summary: "", ops: [], before: [], after: [], inverses: [], notes: [], status: "streaming", raw: "" },
        });
      }
      setBusyBoth(true);
      const requestId = uid();
      activeRequestRef.current = requestId;
      try {
        const out = await requestProposal({
          editor,
          instruction: opts.instruction,
          providerId: activeProviderId,
          selection: selectionText(),
          requestId,
          continueFrom: opts.continueFrom,
          onDelta: (chunk) =>
            patchCard(id, (c) => {
              const plan = c as PlanCard;
              return { ...plan, raw: (plan.raw ?? "") + chunk };
            }),
        });
        if (!out.ok || !out.proposal) {
          if (out.needsSetup) setSettingsOpen(true);
          patchCard(id, (c) => ({
            ...c,
            status: "failed",
            stopped: out.stopped,
            instruction: opts.instruction,
            message: out.message ?? "AI 没有给出修改方案。",
          }));
          return;
        }
        if (!out.proposal.ops.length) {
          dismiss(id);
          note(`检查过了，没有需要改的地方：${out.proposal.summary}`);
          return;
        }
        const plan = simulateProposal(editor, out.proposal);
        if (!plan.ok) {
          patchCard(id, (c) => ({
            ...c,
            status: "failed",
            summary: out.proposal!.summary,
            message: `方案无法应用：${plan.message ?? "请换个说法再试。"}`,
          }));
          return;
        }
        const card: PlanCard = {
          summary: out.proposal.summary,
          ops: plan.ops!,
          before: plan.before!,
          after: plan.after!,
          inverses: plan.inverses!,
          notes: [...out.warnings, ...plan.skipped],
          status: "pending",
          stopped: out.stopped,
          instruction: opts.instruction,
          proposal: out.proposal,
        };
        patchCard(id, () => card);
        // 暂停留下的半截方案不完整，不替用户直接写进文档，等他点「继续」或自己确认。
        if (modeRef.current === "auto" && !out.stopped) {
          toast("全自动：已直接写入文档，可在修改历史里撤销这次修改。", "success");
          await applyPlan(id, card);
        }
      } catch (e) {
        const err = e as Error & { needsSetup?: boolean };
        if (err.needsSetup) setSettingsOpen(true);
        patchCard(id, (c) => ({
          ...c,
          status: "failed",
          instruction: opts.instruction,
          message: "调用失败：" + (err.message ?? String(e)),
        }));
      } finally {
        activeRequestRef.current = null;
        setBusyBoth(false);
      }
    },
    [
      activeProviderId,
      applyPlan,
      dismiss,
      editor,
      note,
      patchCard,
      push,
      selectionText,
      setBusyBoth,
    ],
  );

  const askAgent = useCallback(
    async (instruction: string, label = "AI 修改") => {
      const trimmed = instruction.trim();
      if (!trimmed || busyRef.current) return;
      if (modeRef.current === "chat") {
        note("聊天模式只给建议、不会改动文档。需要我直接改文档时，切到「半自动」或「全自动」。");
        return;
      }
      push({ id: uid(), t: "msg", role: "user", text: trimmed });
      await runAgent({ instruction: trimmed, label });
    },
    [note, push, runAgent],
  );

  /** 点「继续」：已收到的改动留着，只让模型把剩下的补完。 */
  const continuePlan = useCallback(
    async (id: string, card: PlanCard, label: string) => {
      if (busyRef.current || !card.instruction) return;
      await runAgent({
        instruction: card.instruction,
        label,
        cardId: id,
        continueFrom: { ops: card.proposal?.ops ?? [], rawTail: card.raw ?? "" },
      });
    },
    [runAgent],
  );

  /* ------------- sheet: run the tool loop, land on confirm ------- */

  const patchSheetCard = useCallback((id: string, fn: (c: SheetPlanCard) => SheetPlanCard) => {
    setItems((prev) =>
      prev.map((it) => (it.id === id && it.t === "sheet" ? { ...it, card: fn(it.card) } : it)),
    );
  }, []);

  /**
   * 确认落地一张表格提案：op 交给 app 的唯一写入口，它自己核对现场并记修改历史。
   *
   * 图表卡走另一条：它没有 op，只能排进待写清单、随下一次保存写进 `.xlsx`，
   * 所以这里如实告诉用户"撤销栈管不到它"（§2.5 的边界）。
   *
   * @param id 卡片条目 id。
   * @param card 待落地的提案。
   */
  const applySheetPlan = useCallback(
    async (id: string, card: SheetPlanCard) => {
      patchSheetCard(id, (c) => ({ ...c, status: "applying" }));
      if (card.chart) {
        if (!sheet?.onCommitChart(card.chart)) {
          patchSheetCard(id, (c) => ({ ...c, status: "failed", message: "这张表已经关了，图表没有排进去。" }));
          return;
        }
        patchSheetCard(id, (c) => ({
          ...c,
          status: "applied",
          message: "已排进待写清单，保存后写进文件。撤销栈管不到图表，要回退请用同名 .bak。",
        }));
        onChanged();
        return;
      }
      const result = sheet?.onCommitOps(card.ops, card.summary);
      if (!result || !result.ops.length) {
        patchSheetCard(id, (c) => ({
          ...c,
          status: "failed",
          message:
            result?.skipped.map((item) => item.reason).join("；") || "表格内容已经变化，这次没有写入。",
        }));
        return;
      }
      // 落地条数与提案条数不一致时必须说出来：剩下那些没生效（AGENTS.md §2.5）。
      const partial = result.skipped.length
        ? `部分写入（${result.ops.length}/${card.ops.length}）：${result.skipped.map((item) => item.reason).join("；")}`
        : undefined;
      patchSheetCard(id, (c) => ({ ...c, status: "applied", message: partial }));
      onChanged();
      if (partial) toast(partial, "error");
    },
    [onChanged, patchSheetCard, sheet],
  );

  /**
   * 跑一次表格任务：转工具回合，结果落到指定或新建的提案卡上。
   *
   * @param opts.instruction 用户原话；点「继续」时还是这一句。
   * @param opts.label 卡片标题。
   * @param opts.cardId 复用已有卡片（点「继续」时），不给就新挂一张。
   * @param opts.messages 暂停时攒下的 transcript；给了就从这一串往后接着跑，不让模型重读整表。
   */
  const runSheetAsk = useCallback(
    async (opts: {
      instruction: string;
      label: string;
      cardId?: string;
      messages?: ToolTurnMessage[];
    }) => {
      const payload = sheet?.payload;
      if (!payload) return;
      const id = opts.cardId ?? uid();
      if (opts.cardId) {
        patchSheetCard(id, (c) => ({ ...c, status: "streaming", message: undefined, stopped: false }));
      } else {
        push({
          id,
          t: "sheet",
          label: opts.label,
          card: {
            summary: "",
            edits: [],
            ops: [],
            notes: [],
            status: "streaming",
            instruction: opts.instruction,
            messages: [],
          },
        });
      }
      setBusyBoth(true);
      pausedRef.current = false;
      try {
        const outcome = await runSheetAgent({
          sheet: payload,
          instruction: opts.instruction,
          providerId: activeProviderId ?? undefined,
          messages: opts.messages,
          isPaused: () => pausedRef.current,
        });
        if (outcome.kind === "paused") {
          patchSheetCard(id, (c) => ({
            ...c,
            status: "failed",
            stopped: true,
            messages: outcome.messages,
            message: "已暂停，工具回合停在两轮之间。点「继续」接着跑。",
          }));
          return;
        }
        if (outcome.kind === "failed") {
          if (outcome.needsSetup) setSettingsOpen(true);
          patchSheetCard(id, (c) => ({
            ...c,
            status: "failed",
            messages: outcome.messages,
            message: outcome.message,
          }));
          return;
        }
        // 纯分析问题（汇总、找极值）不出提案：这一轮的正文就是答复。
        if (outcome.kind === "answer") {
          dismiss(id);
          note(outcome.text);
          return;
        }
        // 出图那条：引用与数值都已由引擎从现场表读好，这里只挂一张要用户确认的图表卡。
        if (outcome.kind === "chart") {
          const chartLines = sheetChartLines(outcome.chart);
          if (modeRef.current === "chat") {
            // 聊天模式不动表：把要画的图念出来，不挂一张点了会排进待写清单的卡（§2.8）。
            dismiss(id);
            note(`按你说的，这张表可以画这么一张图（聊天模式没有动它）：\n· ${chartLines.join("\n· ")}\n要落地请切到「半自动」或「全自动」。`);
            return;
          }
          const chartCard: SheetPlanCard = {
            summary: `插入一张${CHART_KIND_LABEL[outcome.chart.kind]}`,
            edits: [],
            ops: [],
            notes: [],
            chart: outcome.chart,
            status: "pending",
            instruction: opts.instruction,
            messages: outcome.messages,
          };
          patchSheetCard(id, () => chartCard);
          if (modeRef.current === "auto") {
            toast("全自动：图表已排进待写清单，保存后写进文件。", "success");
            await applySheetPlan(id, chartCard);
          }
          return;
        }
        if (!outcome.ops.length) {
          dismiss(id);
          note(`这份提案里没有能落地的改动：${outcome.notes.join("；") || outcome.summary}`);
          return;
        }
        const card: SheetPlanCard = {
          summary: outcome.summary,
          edits: outcome.edits,
          ops: outcome.ops,
          notes: outcome.notes,
          status: "pending",
          instruction: opts.instruction,
          messages: outcome.messages,
        };
        if (modeRef.current === "chat") {
          // 聊天模式不动表：把想改的地方念出来，不挂一张点了会落地的卡（§2.8）。
          dismiss(id);
          const lines = sheetEditLines(card.edits);
          note(
            `按你说的，这张表可以这么改（聊天模式没有动它）：${card.summary}\n· ${lines.join("\n· ")}${
              card.notes.length ? `\n${card.notes.map((n) => `· ${n}`).join("\n")}` : ""
            }\n要落地请切到「半自动」或「全自动」。`,
          );
          return;
        }
        patchSheetCard(id, () => card);
        // 半自动要用户点「应用到表格」；全自动是用户显式选的例外，仍然逐条记进修改历史可撤销（§2.8）。
        if (modeRef.current === "auto") {
          toast("全自动：已直接写入表格，可在修改历史里撤销这次改动。", "success");
          await applySheetPlan(id, card);
        }
      } catch (e) {
        const err = e as Error & { needsSetup?: boolean };
        if (err.needsSetup) setSettingsOpen(true);
        patchSheetCard(id, (c) => ({ ...c, status: "failed", message: "调用失败：" + (err.message ?? String(e)) }));
      } finally {
        setBusyBoth(false);
      }
    },
    [activeProviderId, applySheetPlan, dismiss, note, patchSheetCard, push, setBusyBoth, sheet],
  );

  /**
   * 用户那句诉求 → 表格代理。
   *
   * @param instruction 原话（已 trim）。
   * @param label 卡片标题。
   */
  const askSheet = useCallback(
    async (instruction: string, label = "AI 表格提案") => {
      const trimmed = instruction.trim();
      if (!trimmed || busyRef.current || !sheet) return;
      push({ id: uid(), t: "msg", role: "user", text: trimmed });
      await runSheetAsk({ instruction: trimmed, label });
    },
    [push, runSheetAsk, sheet],
  );

  /** 点「继续」：把暂停时攒下的 transcript 原样送回，工具回合从这一轮往后跑。 */
  const continueSheetPlan = useCallback(
    async (id: string, card: SheetPlanCard, label: string) => {
      if (busyRef.current || !card.messages.length) return;
      await runSheetAsk({ instruction: card.instruction, label, cardId: id, messages: card.messages });
    },
    [runSheetAsk],
  );

  /* ------------- selection rewrite: stream then apply -------------- */

  const applyRewrite = useCallback(
    async (id: string, card: RewriteCard, label: string) => {
      if (card.scope === "document") {
        const blocks = blocksOf(editor.state.doc);
        const current = rewritableBlocks(blocks)
          .map((b) => b.text)
          .join("\n");
        if (current !== card.beforeText) {
          patchCard(id, (c) => ({
            ...c,
            status: "failed",
            message: "文档内容已经变化，请重新点一次。",
          }));
          return;
        }
        // 全文改写落成逐段 setText：表格、图片、空段一律不动，每段自己可撤销。
        const { ops, errors } = alignRewrite(blocks, card.rewritten);
        if (!ops.length) {
          patchCard(id, (c) => ({ ...c, status: "failed", message: errors.join("；") }));
          return;
        }
        patchCard(id, (c) => ({ ...c, status: "applying" }));
        const r = await applyProposal(editor, { summary: label, ops }, docKey);
        if (!r.ok) {
          patchCard(id, (c) => ({ ...c, status: "failed", message: r.message }));
          return;
        }
        patchCard(id, (c) => ({ ...c, status: "applied", message: r.message }));
        onChanged();
        if (r.message) toast(r.message, "error");
        return;
      }

      const current = editor.state.doc.textBetween(card.from, card.to, "\n");
      if (current !== card.beforeText) {
        patchCard(id, (c) => ({
          ...c,
          status: "failed",
          message: "选区已经变化，请重新选中后再试。",
        }));
        return;
      }
      patchCard(id, (c) => ({ ...c, status: "applying" }));
      const r = await commitSelectionRewrite(editor, {
        from: card.from,
        to: card.to,
        rewritten: card.rewritten,
        label: "AI 改写",
        docKey,
      });
      if (!r.ok) {
        patchCard(id, (c) => ({ ...c, status: "failed", message: r.message }));
        return;
      }
      patchCard(id, (c) => ({ ...c, status: "applied", message: r.message }));
      onChanged();
    },
    [editor, docKey, onChanged, patchCard],
  );

  /**
   * 跑一次纯文字改写（选区或全文），增量流进指定卡片。
   *
   * @param opts.cardId 复用已有卡片（点「继续」时），不给就新挂一张流式卡。
   * @param opts.continuePartial 暂停时已经收到的半截改写；给了就只补剩下的段落，
   *   断在段中间的那半截由 `rewriteResume` 丢掉重写。
   */
  const runRewrite = useCallback(
    async (opts: {
      action: string;
      label: string;
      scope: RewriteCard["scope"];
      from: number;
      to: number;
      source: string;
      cardId?: string;
      continuePartial?: string;
    }) => {
      const id = opts.cardId ?? uid();
      if (opts.cardId) {
        patchCard(id, (c) => ({ ...c, status: "streaming", message: undefined, stopped: false }));
      } else {
        push({
          id,
          t: "rewrite",
          label: opts.label,
          card: {
            scope: opts.scope,
            action: opts.action,
            from: opts.from,
            to: opts.to,
            beforeText: opts.source,
            rewritten: "",
            status: "streaming",
          },
        });
      }
      const partial = opts.continuePartial;
      const resume = partial ? rewriteResume(opts.source, partial) : undefined;
      // 半截已经收满全篇段数：没有要补的段落了，再问一次模型只会多出一段，
      // 所以把卡片还原成可写入的样子让用户自己点。
      if (resume && resume.remaining === 0) {
        patchCard(id, (c) => ({
          ...c,
          status: "pending",
          stopped: false,
          message: "上一次已经收到完整段落，直接点写入即可。",
        }));
        return;
      }
      setBusyBoth(true);
      const requestId = uid();
      activeRequestRef.current = requestId;
      // 续写时留下的段落先占住卡片，模型补的那几段接在后面（换行分段，落地时按段对齐）。
      const seed = resume?.kept.length ? `${resume.kept.join("\n")}\n` : "";
      let acc = seed;
      try {
        const { stopped } = await aiStreamTransform({
          action: opts.action,
          text: opts.source,
          instruction: rewriteInstruction(opts.scope, resume),
          providerId: activeProviderId,
          requestId,
          onDelta: (chunk) => {
            acc += chunk;
            const shown = acc;
            setItems((prev) =>
              prev.map((it) =>
                it.id === id && it.t === "rewrite"
                  ? { ...it, card: { ...it.card, rewritten: shown } }
                  : it,
              ),
            );
          },
        });
        if (!acc.trim()) {
          patchCard(id, (c) => ({ ...c, status: "failed", message: "没有收到内容。" }));
          return;
        }
        // 续写一个字都没补上：卡片上那半截原样留着，「继续」还能再点，别假装收完了。
        if (seed && acc.trim() === seed.trim()) {
          patchCard(id, (c) => ({
            ...c,
            status: "pending",
            stopped: true,
            message: "这一次没有补出新的段落，可以再点一次「继续」。",
          }));
          return;
        }
        const card: RewriteCard = {
          scope: opts.scope,
          action: opts.action,
          from: opts.from,
          to: opts.to,
          beforeText: opts.source,
          rewritten: acc,
          status: "pending",
          stopped,
        };
        patchCard(id, () => card);
        if (modeRef.current === "auto" && !stopped) {
          toast("全自动：已直接写入文档，可在修改历史里撤销这次修改。", "success");
          await applyRewrite(id, card, opts.label);
        }
      } catch (e) {
        const err = e as Error & { needsSetup?: boolean };
        if (err.needsSetup) setSettingsOpen(true);
        patchCard(id, (c) => ({
          ...c,
          status: "failed",
          message: err.message ?? String(e),
        }));
      } finally {
        activeRequestRef.current = null;
        setBusyBoth(false);
      }
    },
    [activeProviderId, applyRewrite, patchCard, push, setBusyBoth],
  );

  /**
   * 跑一条快捷指令：有选区就改选区，没有选区就改整篇（而不是报错让用户先选）。
   *
   * @param action 主进程内置的语气键（`扩写`/`缩写`/`润色`/`换语气:…`/`翻译:…`）。
   */
  const quickRewrite = useCallback(
    async (action: string) => {
      if (busyRef.current) return;
      if (modeRef.current === "chat") {
        note("聊天模式只给建议、不会改动文档。需要改写文字时，切到「半自动」或「全自动」。");
        return;
      }
      const { from, to } = editor.state.selection;
      const selected = editor.state.doc.textBetween(from, to, "\n");
      const scope: RewriteCard["scope"] = selected.trim() ? "selection" : "document";
      const source =
        scope === "selection"
          ? selected
          : rewritableBlocks(blocksOf(editor.state.doc))
              .map((b) => b.text)
              .join("\n");
      if (!source.trim()) {
        note("文档里还没有可以改写的文字。");
        return;
      }
      const label = `${scope === "document" ? "全文" : ""}${action.split(":")[0] ?? action}`;
      await runRewrite({ action, label, scope, from, to, source });
    },
    [editor, note, runRewrite],
  );

  /** 改写卡上的「继续」：在半截结果后面把剩下的写完。 */
  const continueRewrite = useCallback(
    async (id: string, card: RewriteCard, label: string) => {
      if (busyRef.current || !card.rewritten.trim()) return;
      await runRewrite({
        action: card.action,
        label,
        scope: card.scope,
        from: card.from,
        to: card.to,
        source: card.beforeText,
        cardId: id,
        continuePartial: card.rewritten,
      });
    },
    [runRewrite],
  );

  /* --------------------------- chat ------------------------------- */

  /**
   * 跑一次聊天，增量流进指定气泡；给了 `partial` 就是在半截回答后面继续写。
   *
   * @param opts.prompt 用户这一轮说的话（续写时原样再送一次）。
   * @param opts.itemId 复用已有气泡（点「继续」时），不给就新起一问一答。
   * @param opts.partial 暂停留下的半截回答。
   */
  const runChat = useCallback(
    async (opts: { prompt: string; itemId?: string; partial?: string }) => {
      const id = opts.itemId ?? uid();
      if (!opts.itemId) {
        push({ id: uid(), t: "msg", role: "user", text: opts.prompt });
        push({ id, t: "msg", role: "assistant", text: "", pending: true, prompt: opts.prompt });
      } else {
        setItems((prev) =>
          prev.map((it) =>
            it.id === id && it.t === "msg" ? { ...it, pending: true, stopped: false } : it,
          ),
        );
      }
      setBusyBoth(true);
      const requestId = uid();
      activeRequestRef.current = requestId;
      let acc = opts.partial ?? "";
      const text = opts.partial
        ? `${opts.prompt}\n\n接着刚才写到的地方继续写完，不要重复已经写好的内容。已经写好的是：\n${opts.partial}`
        : opts.prompt;
      // 接上文：把会话里已经说过的话按顺序送回模型。「继续」那半截已经在 `text` 里，
      // 不能再作为历史重复一遍。上送多少轮由主进程截（AGENTS §2.9）。
      const history = turnsOfItems(
        opts.itemId ? itemsRef.current.filter((it) => it.id !== opts.itemId) : itemsRef.current,
      );
      try {
        const { stopped } = await aiStreamTransform({
          action: "chat",
          text,
          providerId: activeProviderId,
          requestId,
          history,
          onDelta: (chunk) => {
            acc += chunk;
            const shown = acc;
            setItems((prev) =>
              prev.map((it) => (it.id === id && it.t === "msg" ? { ...it, text: shown } : it)),
            );
          },
        });
        // 暂停留下的半截留在原地等「继续」；没收到字才算真的停空了。
        const canResume = stopped && Boolean(acc.trim());
        setItems((prev) =>
          prev.map((it) =>
            it.id === id && it.t === "msg"
              ? {
                  ...it,
                  text: acc || "（已停止，未收到内容）",
                  pending: false,
                  stopped: canResume || undefined,
                }
              : it,
          ),
        );
      } catch (e) {
        const err = e as Error & { needsSetup?: boolean };
        if (err.needsSetup) setSettingsOpen(true);
        setItems((prev) =>
          prev.map((it) =>
            it.id === id && it.t === "msg"
              ? { ...it, text: "调用失败：" + (err.message ?? String(e)), pending: false }
              : it,
          ),
        );
      } finally {
        activeRequestRef.current = null;
        setBusyBoth(false);
      }
    },
    [activeProviderId, push, setBusyBoth],
  );

  const sendChat = useCallback(
    async (raw: string) => {
      const text = raw.trim();
      if (!text || busyRef.current) return;
      // 表格没有"选区改写"这条路：分析、整理、计算都交给代理决定出不出提案。
      if (sheet) {
        await askSheet(text);
        return;
      }
      if (modeRef.current !== "chat") {
        await askAgent(text);
        return;
      }
      await runChat({ prompt: text });
    },
    [askAgent, askSheet, runChat, sheet],
  );

  const stopStream = useCallback(() => {
    // 表格那条没有流式请求可取消，靠这个位在两轮之间停下；文档那条两种都要。
    pausedRef.current = true;
    const id = activeRequestRef.current;
    if (id) window.officeCopilot.aiStreamCancel(id);
  }, []);

  useEffect(
    () => () => {
      if (activeRequestRef.current) {
        window.officeCopilot.aiStreamCancel(activeRequestRef.current);
      }
    },
    [],
  );

  const copyText = useCallback(async (id: string, text: string) => {
    await navigator.clipboard.writeText(text);
    setCopiedId(id);
    setTimeout(() => setCopiedId(null), 1500);
  }, []);

  /* ------------- menu / palette accelerators ------------------------ */

  useEffect(() => {
    if (!window.officeCopilot) return;
    const onRun = (e: Event) => {
      const action = (e as CustomEvent<string>).detail;
      const preset = DOC_BY_LABEL.get(action);
      if (preset) void askAgent(preset.instruction, preset.label);
      else void quickRewrite(action);
    };
    const subs = [
      window.officeCopilot.onMenu("ai:expand", () => void quickRewrite("扩写")),
      window.officeCopilot.onMenu("ai:shorten", () => void quickRewrite("缩写")),
      window.officeCopilot.onMenu("ai:polish", () => void quickRewrite("润色")),
    ];
    window.addEventListener("ai:run", onRun);
    return () => {
      subs.forEach((u) => u());
      window.removeEventListener("ai:run", onRun);
    };
  }, [askAgent, quickRewrite]);

  const activeProvider: ProviderConfig | null =
    settings?.providers.find((p) => p.id === activeProviderId) ?? null;

  if (collapsed) {
    return (
      <>
        <aside
          className="ai-panel ai-panel-collapsed"
          onClick={onToggle}
          title="展开 AI 助手（⌥⌘L）"
        >
          <div className="ai-collapsed-rail">
            <Sparkles size={20} />
            <span className="ai-collapsed-label">AI</span>
            <button
              className="icon"
              title="展开 AI 助手"
              onClick={(e) => {
                e.stopPropagation();
                onToggle();
              }}
            >
              <PanelRightOpen size={16} />
            </button>
          </div>
        </aside>
        <SettingsModal
          open={settingsOpen}
          onClose={() => setSettingsOpen(false)}
          onSaved={refreshSettings}
        />
      </>
    );
  }

  return (
    <>
      <aside className="ai-panel">
        <header>
          <span className="ai-header-title">AI 助手</span>
          <div className="ai-session-wrap" ref={sessionsAnchorRef}>
            <button
              type="button"
              className={"icon" + (sessionsOpen ? " open" : "")}
              title="会话历史：新建会话、接着聊"
              aria-expanded={sessionsOpen}
              onClick={toggleSessions}
            >
              <MessagesSquare size={14} />
            </button>
            {sessionsOpen && sessionsRect && (
              <>
                <div className="ai-pop-overlay" onClick={hideSessionsMenu} />
                <div
                  className="ai-session-pop"
                  style={{ left: sessionsRect.left, top: sessionsRect.top }}
                >
                  <div className="ai-session-head">
                    <span>会话历史</span>
                    <button
                      type="button"
                      className="ai-session-new"
                      title="把当前这份存进历史，另起一个干净会话"
                      disabled={busy}
                      onClick={() => {
                        hideSessionsMenu();
                        void startSession();
                      }}
                    >
                      <MessageSquarePlus size={12} />
                      新建会话
                    </button>
                  </div>
                  {!sessions.length && (
                    <div className="ai-session-empty">还没有存过的会话。</div>
                  )}
                  {sessions.map((session) => (
                    <div key={session.id} className="ai-session-row">
                      <button
                        type="button"
                        className="ai-session-open"
                        title="打开这个会话，接着聊"
                        onClick={() => void loadSession(session.id)}
                      >
                        <span className="ai-session-title">{session.title}</span>
                        <span className="ai-session-meta">
                          {new Date(session.updatedAt).toLocaleString()} · {session.turnCount} 条
                        </span>
                      </button>
                      <button
                        type="button"
                        className="icon ai-session-del"
                        title="删除这个会话"
                        onClick={() => void dropSession(session.id)}
                      >
                        <Trash2 size={12} />
                      </button>
                    </div>
                  ))}
                </div>
              </>
            )}
          </div>
          <button className="icon" title="修改历史" onClick={onOpenHistory}>
            <History size={14} />
          </button>
          <div className="provider-menu-wrap">
            <button
              onClick={() => setProviderMenuOpen((v) => !v)}
              className="provider-pill"
              title={
                activeProvider
                  ? `${activeProvider.name} · ${activeProvider.model}`
                  : "未配置模型"
              }
            >
              <span className="provider-pill-name">
                {activeProvider ? activeProvider.name : "未配置"}
              </span>
              <ChevronDown size={12} className="provider-pill-caret" />
            </button>
            {providerMenuOpen && (
              <>
                <div
                  className="menu-overlay"
                  onClick={() => setProviderMenuOpen(false)}
                />
                <div className="provider-menu">
                  {settings?.providers.filter((p) => p.enabled).length === 0 && (
                    <div className="provider-menu-empty">还没有模型</div>
                  )}
                  {settings?.providers
                    .filter((p) => p.enabled)
                    .map((p) => (
                      <div
                        key={p.id}
                        className={
                          "provider-menu-item" +
                          (p.id === activeProviderId ? " active" : "")
                        }
                        onClick={() => {
                          setActiveProviderId(p.id);
                          setProviderMenuOpen(false);
                        }}
                      >
                        <div>{p.name}</div>
                        <div className="provider-menu-sub">
                          {p.model || "(未填 model)"}
                        </div>
                      </div>
                    ))}
                  <div
                    className="provider-menu-manage"
                    onClick={() => {
                      setProviderMenuOpen(false);
                      setSettingsOpen(true);
                    }}
                  >
                    管理模型…
                  </div>
                </div>
              </>
            )}
          </div>
          <button
            className="icon"
            title="模型设置"
            onClick={() => setSettingsOpen(true)}
          >
            <Settings2 size={14} />
          </button>
          <button
            className="icon"
            title="收起侧边栏（⌥⌘L）"
            onClick={onToggle}
          >
            <PanelRightClose size={14} />
          </button>
        </header>

        <div className="messages" ref={listRef}>
          {items.map((it) => {
            if (it.t === "msg") {
              return (
                <div
                  key={it.id}
                  className={`bubble ${it.role}${it.pending ? " pending" : ""}`}
                >
                  <div className="msg-text">{it.text || (it.pending ? "思考中…" : "")}</div>
                  {it.role === "assistant" && it.text && !it.pending && (
                    <div className="msg-actions">
                      <button
                        type="button"
                        className="msg-btn"
                        title="复制"
                        onClick={() => void copyText(it.id, it.text)}
                      >
                        {copiedId === it.id ? (
                          <>
                            <Check size={11} /> 已复制
                          </>
                        ) : (
                          <>
                            <Copy size={11} /> 复制
                          </>
                        )}
                      </button>
                      {it.stopped && it.prompt && (
                        <button
                          type="button"
                          className="msg-btn"
                          title="接着刚才停下的地方继续写"
                          onClick={() => void runChat({ prompt: it.prompt!, itemId: it.id, partial: it.text })}
                        >
                          <Play size={11} /> 继续
                        </button>
                      )}
                    </div>
                  )}
                </div>
              );
            }

            if (it.t === "plan") {
              const c = it.card;
              // 流式阶段每来一个增量都重算一次预览：收到哪段就显示哪段。
              const streamLines = c.status === "streaming" ? streamedParagraphs(c.raw ?? "") : null;
              return (
                <div key={it.id} className={`ai-card ${c.status}`}>
                  <div className="ai-card-head">
                    <Sparkles size={13} />
                    <span className="ai-card-title">{it.label}方案</span>
                    {c.status === "streaming" ? (
                      <span className="ai-card-streaming">
                        <Loader2 size={12} className="spin" />
                        {streamLines?.length ? `正在写第 ${streamLines.length} 段…` : "正在输出…"}
                      </span>
                    ) : (
                      <span className="ai-card-count">{c.ops.length} 处改动</span>
                    )}
                  </div>
                  {c.status === "streaming" ? (
                    <StreamPreview
                      summary={streamedSummary(c.raw ?? "")}
                      lines={streamLines ?? []}
                    />
                  ) : (
                    <>
                      <div className="ai-card-summary">{c.summary}</div>
                      <ul className="op-list">
                        {c.ops.map((op, i) => (
                          <li key={i}>{describeCoreOp(op)}</li>
                        ))}
                      </ul>
                      <DiffView before={c.before} after={c.after} />
                      {c.notes.length > 0 && (
                        <div className="ai-card-notes">
                          {c.notes.map((n, i) => (
                            <div key={i}>· {n}</div>
                          ))}
                        </div>
                      )}
                    </>
                  )}
                  {c.message && <div className="ai-card-message">{c.message}</div>}
                  <div className="ai-card-actions">
                    {c.status === "pending" && (
                      <>
                        <button
                          className="primary"
                          onClick={() => void applyPlan(it.id, c)}
                        >
                          <Check size={12} /> 应用到文档
                        </button>
                        <button className="ghost" onClick={() => dismiss(it.id)}>
                          <X size={12} /> 放弃
                        </button>
                      </>
                    )}
                    {c.status === "applying" && (
                      <span className="ai-card-status">
                        <Loader2 size={12} className="spin" /> 正在写入…
                      </span>
                    )}
                    {c.status === "applied" && (
                      <span className="ai-card-status ok">
                        <Check size={12} /> 已写入文档
                        {onOpenHistory && (
                          <button className="link-btn" onClick={onOpenHistory}>
                            查看 diff
                          </button>
                        )}
                      </span>
                    )}
                    {c.status === "failed" && (
                      <span className="ai-card-status bad">
                        {c.stopped ? "已暂停" : "未写入"}
                      </span>
                    )}
                    {c.stopped && c.instruction && (
                      <button
                        className="primary"
                        title="接着上一次停下的地方继续写"
                        onClick={() => void continuePlan(it.id, c, it.label)}
                      >
                        <Play size={12} /> 继续
                      </button>
                    )}
                    {(c.status === "pending" || c.status === "failed") && (
                      <button className="ghost" onClick={() => dismiss(it.id)}>
                        忽略
                      </button>
                    )}
                  </div>
                </div>
              );
            }

            if (it.t === "sheet") {
              const c = it.card;
              // 图表卡列的是引用与落点，提案卡列的是改动条目；两样都在渲染层组装（§5.7）。
              const lines = c.chart ? sheetChartLines(c.chart) : sheetEditLines(c.edits);
              // 卡片沿用文档方案卡的那套类名（`.ai-card*` / `.op-list`），只换标题图标与文案。
              return (
                <div key={it.id} className={`ai-card ${c.status}`}>
                  <div className="ai-card-head">
                    <Table2 size={13} />
                    <span className="ai-card-title">{it.label}</span>
                    {c.status === "streaming" ? (
                      <span className="ai-card-streaming">
                        <Loader2 size={12} className="spin" />
                        正在核对表格…
                      </span>
                    ) : (
                      <span className="ai-card-count">{c.chart ? "1 张图表" : `${c.ops.length} 处改动`}</span>
                    )}
                  </div>
                  {c.summary && <div className="ai-card-summary">{c.summary}</div>}
                  {lines.length > 0 && (
                    <ul className="op-list">
                      {lines.map((line, i) => (
                        <li key={i}>{line}</li>
                      ))}
                    </ul>
                  )}
                  {c.notes.length > 0 && (
                    <div className="ai-card-notes">
                      {c.notes.map((n, i) => (
                        <div key={i}>· {n}</div>
                      ))}
                    </div>
                  )}
                  {c.message && <div className="ai-card-message">{c.message}</div>}
                  <div className="ai-card-actions">
                    {c.status === "pending" && (
                      <>
                        <button className="primary" onClick={() => void applySheetPlan(it.id, c)}>
                          <Check size={12} /> {c.chart ? "插入图表" : "应用到表格"}
                        </button>
                        <button className="ghost" onClick={() => dismiss(it.id)}>
                          <X size={12} /> 放弃
                        </button>
                      </>
                    )}
                    {c.status === "applying" && (
                      <span className="ai-card-status">
                        <Loader2 size={12} className="spin" /> 正在写入…
                      </span>
                    )}
                    {c.status === "applied" && (
                      <span className="ai-card-status ok">
                        <Check size={12} /> {c.chart ? "已排进待写清单" : "已写入表格"}
                        {onOpenHistory && (
                          <button className="link-btn" onClick={onOpenHistory}>
                            查看修改历史
                          </button>
                        )}
                      </span>
                    )}
                    {c.status === "failed" && (
                      <span className="ai-card-status bad">{c.stopped ? "已暂停" : "未写入"}</span>
                    )}
                    {c.stopped && c.messages.length > 0 && (
                      <button
                        className="primary"
                        title="接着上一次停下的地方继续跑工具回合"
                        onClick={() => void continueSheetPlan(it.id, c, it.label)}
                      >
                        <Play size={12} /> 继续
                      </button>
                    )}
                    {(c.status === "pending" || c.status === "failed") && (
                      <button className="ghost" onClick={() => dismiss(it.id)}>
                        忽略
                      </button>
                    )}
                  </div>
                </div>
              );
            }

            const c = it.card;
            return (
              <div key={it.id} className={`ai-card ${c.status}`}>
                <div className="ai-card-head">
                  <PencilLine size={13} />
                  <span className="ai-card-title">{it.label}</span>
                  {c.status === "streaming" && (
                    <span className="ai-card-status">
                      <Loader2 size={12} className="spin" /> 生成中
                    </span>
                  )}
                </div>
                <InlineDiff before={c.beforeText} after={c.rewritten} />
                {c.message && <div className="ai-card-message">{c.message}</div>}
                <div className="ai-card-actions">
                  {c.status === "pending" && (
                    <>
                      <button
                        className="primary"
                        onClick={() => void applyRewrite(it.id, c, it.label)}
                      >
                        <Check size={12} /> {c.scope === "document" ? "写入全文" : "替换选中"}
                      </button>
                      <button className="ghost" onClick={() => dismiss(it.id)}>
                        <X size={12} /> 放弃
                      </button>
                    </>
                  )}
                  {c.status === "applying" && (
                    <span className="ai-card-status">
                      <Loader2 size={12} className="spin" /> 正在写入…
                    </span>
                  )}
                  {c.status === "applied" && (
                    <span className="ai-card-status ok">
                      <Check size={12} /> 已写入文档
                    </span>
                  )}
                  {c.status === "failed" && (
                    <span className="ai-card-status bad">{c.stopped ? "已暂停" : "未写入"}</span>
                  )}
                  {(c.status === "pending" || c.status === "failed") && c.stopped && (
                    <button
                      className="primary"
                      title="接着上一次停下的地方继续写"
                      onClick={() => void continueRewrite(it.id, c, it.label)}
                    >
                      <Play size={12} /> 继续
                    </button>
                  )}
                </div>
              </div>
            );
          })}
        </div>

        <div className="composer">
          <div className="ai-mode-bar">
            <div className="ai-pop-wrap">
              <button
                type="button"
                className="ai-mode-toggle"
                onClick={() => setModeTabsOpen((v) => !v)}
                title="AI 工作模式"
              >
                <ChevronDown
                  size={12}
                  className={"ai-chips-caret" + (modeTabsOpen ? " open" : "")}
                />
                模式 · {MODES.find((m) => m.id === mode)?.label}
              </button>
              {modeTabsOpen && (
                <>
                  <div
                    className="ai-pop-overlay"
                    onClick={() => setModeTabsOpen(false)}
                  />
                  <div className="ai-pop">
                    <div className="ai-mode-tabs">
                      {MODES.map((m) => (
                        <button
                          key={m.id}
                          type="button"
                          className={
                            "ai-mode-tab" + (mode === m.id ? " active" : "")
                          }
                          onClick={() => {
                            setModeBoth(m.id);
                            setModeTabsOpen(false);
                          }}
                          title={m.title}
                        >
                          <m.icon size={12} /> {m.label}
                        </button>
                      ))}
                    </div>
                  </div>
                </>
              )}
            </div>
            {/* 快捷指令全是文档动作（改写、排版），表格模式下不给出来。 */}
            {!sheet && mode !== "chat" && (
              <div className="ai-pop-wrap">
                <button
                  type="button"
                  className="ai-chips-toggle"
                  onClick={() => setChipsOpen((v) => !v)}
                  title="快捷指令"
                >
                  <ChevronDown
                    size={12}
                    className={"ai-chips-caret" + (chipsOpen ? " open" : "")}
                  />
                  快捷指令
                </button>
                {chipsOpen && (
                  <>
                    <div
                      className="ai-pop-overlay"
                      onClick={() => setChipsOpen(false)}
                    />
                    <div className="ai-pop">
                      <div className="ai-chips">
                        {QUICK_ACTIONS.map((a) => (
                          <span
                            key={a.label}
                            className="chip"
                            title="改写选中文字；没有选中时改写整篇文档"
                            onClick={() => void quickRewrite(a.action)}
                          >
                            {a.label}
                          </span>
                        ))}
                        {DOC_ACTION_CHIPS.map((d) => (
                          <span
                            key={d.label}
                            className="chip chip-doc"
                            title="对整份文档生效"
                            onClick={() => void askAgent(d.instruction, d.label)}
                          >
                            <d.icon size={11} /> {d.label}
                          </span>
                        ))}
                        <span
                          className="chip chip-accent"
                          onClick={() =>
                            window.dispatchEvent(
                              new CustomEvent("ai:format-document"),
                            )
                          }
                          title="离线规范全文标题/正文字号（⌥⌘⇧F），同样记录到修改历史"
                        >
                          <Wand2 size={11} /> 自动排版
                        </span>
                      </div>
                    </div>
                  </>
                )}
              </div>
            )}
            <button
              type="button"
              className={"ai-mode-info" + (modeHintOpen ? " open" : "")}
              title="说明"
              onClick={() => setModeHintOpen((v) => !v)}
            >
              <Info size={13} />
            </button>
          </div>
          {modeHintOpen && (
            <div className="ai-mode-hint">
              {sheet ? SHEET_MODE_HINTS[mode] : MODES.find((m) => m.id === mode)?.hint}
            </div>
          )}
          <textarea
            value={input}
            onChange={(e) => setInput(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === "Enter" && (e.metaKey || e.ctrlKey)) {
                e.preventDefault();
                void sendChat(input);
                setInput("");
              }
            }}
            placeholder={
              sheet
                ? "说要对这张表做什么，⌘/Ctrl+Enter 发送。例如：按区域汇总销售额，再删掉重复的行。"
                : mode === "chat"
                  ? "和 AI 聊聊这份文档，⌘/Ctrl+Enter 发送。"
                  : "描述要怎么改这份文档，⌘/Ctrl+Enter 发送。例如：在第三节后加一张要点汇总表。"
            }
          />
          <div className="composer-row">
            {busy ? (
              <button className="primary btn-stop" onClick={stopStream} title="暂停生成（已经生成的内容会留着，可以点「继续」接着写）">
                <Square size={11} /> 暂停
              </button>
            ) : (
              <button
                className="primary"
                onClick={() => {
                  void sendChat(input);
                  setInput("");
                }}
                disabled={!input.trim()}
              >
                {sheet
                  ? mode === "chat"
                    ? "提问"
                    : mode === "auto"
                      ? "直接改表"
                      : "让 AI 处理"
                  : mode === "chat"
                    ? "发送"
                    : mode === "auto"
                      ? "直接修改"
                      : "让 AI 修改"}
              </button>
            )}
          </div>
        </div>
      </aside>
      <SettingsModal
        open={settingsOpen}
        onClose={() => setSettingsOpen(false)}
        onSaved={refreshSettings}
      />
    </>
  );
}

/**
 * 方案卡的流式预览：把模型正在写的正文逐段排出来，自动滚到最新一段。
 *
 * @param summary 提案说明，`"summary"` 的值收到才有，没收到时是空串。
 * @param lines 已经出现的段落文字（最后一段可能只写了一半）。
 */
function StreamPreview({ summary, lines }: { summary: string; lines: string[] }) {
  const boxRef = useRef<HTMLDivElement>(null);
  const lastLine = lines[lines.length - 1] ?? "";
  useEffect(() => {
    const box = boxRef.current;
    if (box) box.scrollTop = box.scrollHeight;
  }, [lastLine]);

  if (!lines.length) {
    return <div className="ai-card-stream empty" ref={boxRef}>{summary || "正在等模型回话…"}</div>;
  }
  return (
    <div className="ai-card-stream" ref={boxRef}>
      {summary && <div className="ai-stream-summary">{summary}</div>}
      {lines.map((text, i) => (
        <p key={i} className="ai-stream-p">
          {text}
          {i === lines.length - 1 && <span className="ai-stream-caret" />}
        </p>
      ))}
    </div>
  );
}
