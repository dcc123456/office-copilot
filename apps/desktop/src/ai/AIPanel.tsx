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
  Zap,
} from "lucide-react";
import { SettingsModal } from "../components/SettingsModal.js";
import { DiffView, InlineDiff } from "./DiffView.js";
import { describeCoreOp, type Block, type CoreOp } from "./docEdit.js";
import {
  aiStreamTransform,
  commitOps,
  commitSelectionRewrite,
  requestProposal,
  simulateProposal,
} from "./agent.js";
import { toast } from "../components/ui/Toast.js";
import type { AppSettings, ProviderConfig } from "../global.js";
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
}

interface RewriteCard {
  from: number;
  to: number;
  beforeText: string;
  rewritten: string;
  status: CardStatus;
  message?: string;
}

type Item =
  | {
      id: string;
      t: "msg";
      role: "user" | "assistant";
      text: string;
      pending?: boolean;
    }
  | { id: string; t: "plan"; label: string; card: PlanCard }
  | { id: string; t: "rewrite"; label: string; card: RewriteCard };

const uid = () => crypto.randomUUID();

const SELECTION_ACTIONS = [
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

export function AIPanel({
  editor,
  docKey,
  collapsed,
  onToggle,
  onOpenHistory,
  onChanged,
}: {
  editor: Editor;
  docKey: string;
  collapsed: boolean;
  onToggle: () => void;
  onOpenHistory?: () => void;
  onChanged: () => void;
}) {
  const [items, setItems] = useState<Item[]>([
    {
      id: "welcome",
      t: "msg",
      role: "assistant",
      text:
        "选中文字后点快捷按钮，或直接写下你的要求（比如「把第三节改成表格」「全文纠正错别字」）。默认是「半自动」：我先给 diff 预览，你确认后才写入文档；「全自动」直接改文档，「聊天」只给建议。每次改动都能在修改历史里看 diff 并单独撤销。",
    },
  ]);
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
  const listRef = useRef<HTMLDivElement>(null);
  const activeRequestRef = useRef<string | null>(null);
  const busyRef = useRef(false);

  const setBusyBoth = useCallback((v: boolean) => {
    busyRef.current = v;
    setBusy(v);
  }, []);

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

  const askAgent = useCallback(
    async (instruction: string, label = "AI 修改") => {
      const trimmed = instruction.trim();
      if (!trimmed || busyRef.current) return;
      if (modeRef.current === "chat") {
        note("聊天模式只给建议、不会改动文档。需要我直接改文档时，切到「半自动」或「全自动」。");
        return;
      }
      setBusyBoth(true);
      push({ id: uid(), t: "msg", role: "user", text: trimmed });
      // 卡片先挂出来：流式阶段逐字显示模型的输出，收完才换成改动清单与 diff。
      const id = uid();
      push({
        id,
        t: "plan",
        label,
        card: { summary: "", ops: [], before: [], after: [], inverses: [], notes: [], status: "streaming", raw: "" },
      });
      const requestId = uid();
      activeRequestRef.current = requestId;
      try {
        const out = await requestProposal({
          editor,
          instruction: trimmed,
          providerId: activeProviderId,
          selection: selectionText(),
          requestId,
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
        };
        patchCard(id, () => card);
        if (modeRef.current === "auto") {
          toast("全自动：已直接写入文档，可在修改历史里撤销这次修改。", "success");
          await applyPlan(id, card);
        }
      } catch (e) {
        const err = e as Error & { needsSetup?: boolean };
        if (err.needsSetup) setSettingsOpen(true);
        patchCard(id, (c) => ({
          ...c,
          status: "failed",
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

  /* ------------- selection rewrite: stream then apply -------------- */

  const applyRewrite = useCallback(
    async (id: string, card: RewriteCard) => {
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

  const rewriteSelection = useCallback(
    async (action: string) => {
      if (busyRef.current) return;
      if (modeRef.current === "chat") {
        note("聊天模式只给建议、不会改动文档。需要改写选中的文字时，切到「半自动」或「全自动」。");
        return;
      }
      const { from, to } = editor.state.selection;
      const text = editor.state.doc.textBetween(from, to, "\n");
      if (!text.trim()) {
        note("请先在文档里选中要改写的文字。");
        return;
      }
      const id = uid();
      push({
        id,
        t: "rewrite",
        label: action.split(":")[0] ?? action,
        card: { from, to, beforeText: text, rewritten: "", status: "streaming" },
      });
      setBusyBoth(true);
      const requestId = uid();
      activeRequestRef.current = requestId;
      let acc = "";
      try {
        await aiStreamTransform({
          action,
          text,
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
        patchCard(id, (c) => ({ ...c, status: "pending" }));
        if (modeRef.current === "auto") {
          toast("全自动：已直接写入文档，可在修改历史里撤销这次修改。", "success");
          await applyRewrite(id, { from, to, beforeText: text, rewritten: acc, status: "pending" });
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
    [activeProviderId, applyRewrite, editor, note, patchCard, push, setBusyBoth],
  );

  /* --------------------------- chat ------------------------------- */

  const sendChat = useCallback(
    async (raw: string) => {
      const text = raw.trim();
      if (!text || busyRef.current) return;
      if (modeRef.current !== "chat") {
        await askAgent(text);
        return;
      }
      setBusyBoth(true);
      push({ id: uid(), t: "msg", role: "user", text });
      const id = uid();
      push({ id, t: "msg", role: "assistant", text: "", pending: true });
      const requestId = uid();
      activeRequestRef.current = requestId;
      let acc = "";
      try {
        await aiStreamTransform({
          action: "chat",
          text,
          providerId: activeProviderId,
          requestId,
          onDelta: (chunk) => {
            acc += chunk;
            const shown = acc;
            setItems((prev) =>
              prev.map((it) => (it.id === id && it.t === "msg" ? { ...it, text: shown } : it)),
            );
          },
        });
        setItems((prev) =>
          prev.map((it) =>
            it.id === id && it.t === "msg"
              ? { ...it, text: acc || "（已停止，未收到内容）", pending: false }
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
    [activeProviderId, askAgent, push, setBusyBoth],
  );

  const stopStream = useCallback(() => {
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
      else void rewriteSelection(action);
    };
    const subs = [
      window.officeCopilot.onMenu("ai:expand", () => void rewriteSelection("扩写")),
      window.officeCopilot.onMenu("ai:shorten", () => void rewriteSelection("缩写")),
      window.officeCopilot.onMenu("ai:polish", () => void rewriteSelection("润色")),
    ];
    window.addEventListener("ai:run", onRun);
    return () => {
      subs.forEach((u) => u());
      window.removeEventListener("ai:run", onRun);
    };
  }, [askAgent, rewriteSelection]);

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
                    </div>
                  )}
                </div>
              );
            }

            if (it.t === "plan") {
              const c = it.card;
              return (
                <div key={it.id} className={`ai-card ${c.status}`}>
                  <div className="ai-card-head">
                    <Sparkles size={13} />
                    <span className="ai-card-title">{it.label}方案</span>
                    {c.status === "streaming" ? (
                      <span className="ai-card-streaming">
                        <Loader2 size={12} className="spin" /> 正在输出…
                      </span>
                    ) : (
                      <span className="ai-card-count">{c.ops.length} 处改动</span>
                    )}
                  </div>
                  {c.status === "streaming" ? (
                    <div className="ai-card-stream">{c.raw || "正在等模型回话…"}</div>
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
                      <span className="ai-card-status bad">未写入</span>
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
                        onClick={() => void applyRewrite(it.id, c)}
                      >
                        <Check size={12} /> 替换选中
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
                  {c.status === "failed" && <span className="ai-card-status bad">未写入</span>}
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
            {mode !== "chat" && (
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
                        {SELECTION_ACTIONS.map((a) => (
                          <span
                            key={a.label}
                            className="chip"
                            title="改写选中文字（需要先选中）"
                            onClick={() => void rewriteSelection(a.action)}
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
              {MODES.find((m) => m.id === mode)?.hint}
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
              mode === "chat"
                ? "和 AI 聊聊这份文档，⌘/Ctrl+Enter 发送。"
                : "描述要怎么改这份文档，⌘/Ctrl+Enter 发送。例如：在第三节后加一张要点汇总表。"
            }
          />
          <div className="composer-row">
            {busy ? (
              <button className="primary btn-stop" onClick={stopStream} title="停止生成">
                <Square size={11} /> 停止
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
                {mode === "chat" ? "发送" : mode === "auto" ? "直接修改" : "让 AI 修改"}
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
