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
  | { id: string; t: "rewrite"; label: string; card: RewriteCard };

const uid = () => crypto.randomUUID();

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
        "选中文字后点快捷按钮就改那几段，不选中就是改整篇；也可以直接写下你的要求（比如「把第三节改成表格」「全文纠正错别字」）。默认是「半自动」：我先给 diff 预览，你确认后才写入文档；「全自动」直接改文档，「聊天」只给建议。每次改动都能在修改历史里看 diff 并单独撤销。",
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
      try {
        const { stopped } = await aiStreamTransform({
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
      if (modeRef.current !== "chat") {
        await askAgent(text);
        return;
      }
      await runChat({ prompt: text });
    },
    [askAgent, runChat],
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
