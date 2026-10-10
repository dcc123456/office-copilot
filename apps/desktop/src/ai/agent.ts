import type { Editor } from "@tiptap/react";
import {
  addressOf,
  buildOutline,
  coerceProposal,
  planEdit,
  revert,
  runOps,
  type AiOp,
  type Block,
  type CoreOp,
  type Proposal,
} from "./docEdit.js";
import { describeCoreOp } from "./docEdit.js";
import { diagramResolverFor } from "./diagramCanvas.js";
import { blocksOf, executeOps, locateAll, normalise } from "../editor/docBlocks.js";

/**
 * The document-editing agent: ask the model for a JSON edit plan, validate it
 * against the real outline, apply it as one transaction, and record enough to
 * diff and individually undo it later.
 */

export interface ModificationRecord {
  id: string;
  ts: number;
  label: string;
  summary: string;
  ops: CoreOp[];
  inverses: CoreOp[];
  before: Block[];
  after: Block[];
  undone: boolean;
  undoneAt?: number;
  /** 美化排版记录打标，一键还原据此一次找齐。 */
  kind?: "template";
  /** 套模板前的页边距/页眉页脚/超链接配色；它们不走段落 op，只能靠这份留底复原。 */
  docLevel?: Record<string, unknown>;
}

export const SYSTEM_PROMPT = `你是文档编辑助手，可以直接修改用户的文档。你会看到文档大纲，每行格式为：
[索引] 类型 "文字" {当前样式}
其中类型为 P（正文）、H1…H6（标题）、TABLE、IMAGE。

只输出一个 JSON 对象，不要解释、不要代码块围栏。结构：
{"summary":"一句话说明这次改动做了什么","ops":[ 操作列表 ]}

可用操作：
1. {"op":"setText","index":N,"text":"整段新文字"} — 重写第 N 段。扩写、缩写、润色、改错别字、翻译都用它。text 必须是完整的新段落文字，不是差异片段。
2. {"op":"setStyle","index":N,"style":{...}} — 修改第 N 段的样式/格式。style 支持的字段：
   headingLevel(1~6，null 表示降为正文)、align("left"|"center"|"right"|"justify")、
   bold/italic/underline(true|false)、
   fontSize("16px")、fontFamily("宋体")、fontFamilyWest("Times New Roman")、
   letterSpacing("0.67px")、color("#cc0000")、highlight("#ffff00")、lineHeight("1.5")、
   spacingBefore/spacingAfter(单位磅)、indent/indentFirst(单位磅)。把不需要的字段设成 null 表示清除。
   段内只有部分文字带某种字符样式（如加粗、颜色）时，整段设置那一项会被拒绝，只设其余字段。
3. {"op":"insertParagraph","after":N,"text":"...","style":{...}} — 在第 N 段后插入。after=-1 插到文首。
4. {"op":"insertTable","after":N,"rows":[["列1","列2"],["值1","值2"]],"headerRow":true} — 插入表格，rows 必须是等宽的二维数组。
5. {"op":"insertDiagram","after":N,"alt":"图注","diagram":{"title":"图标题","layout":"flow-v|flow-h|grid","nodes":[{"id":"a","label":"步骤一","shape":"box|round|diamond"}],"edges":[{"from":"a","to":"b","label":"可选条件"}]}} — 生成一张 canvas 示意图（流程图/结构图）插入文档，用来增强说明。
6. {"op":"deleteBlock","index":N} — 删除第 N 段。
7. {"op":"replaceAll","find":"旧词","replace":"新词"} — 全文替换一个词。

规则：
- index/after 只能是大纲里真实存在的编号（插入时 after 用现有段落编号或 -1）。
- 自动排版：给标题设 headingLevel/align/bold，给正文设 fontSize、lineHeight、indentFirst 等；排版任务不要改文字。
- 纠正错别字：只改确实有错的段落，并在 summary 里列出改了哪几个词；没有错误就返回 {"summary":"未发现错别字","ops":[]}。
- 保持文档原有语言，除非用户明确要求翻译。
- 一次改动尽量精简，只做用户要求的事。`;

export function buildUserMessage(instruction: string, blocks: Block[], selection?: string): string {
  const parts = [`文档大纲：\n${buildOutline(blocks)}`];
  if (selection && selection.trim()) {
    parts.push(`用户当前选中的文字：\n${selection.slice(0, 2000)}`);
  }
  parts.push(`指令：${instruction}`);
  return parts.join("\n\n");
}

/** Models wrap JSON in prose or fences often enough that this is worth doing. */
export function extractJson(text: string): unknown | null {
  const withoutFence = text.replace(/```(?:json)?/gi, "").trim();
  const start = withoutFence.indexOf("{");
  const end = withoutFence.lastIndexOf("}");
  if (start < 0 || end <= start) return null;
  try {
    return JSON.parse(withoutFence.slice(start, end + 1));
  } catch {
    return null;
  }
}

export interface ProposalOutcome {
  ok: boolean;
  proposal?: Proposal;
  before: Block[];
  message?: string;
  needsSetup?: boolean;
  warnings: string[];
}

export async function requestProposal(opts: {
  editor: Editor;
  instruction: string;
  providerId?: string | null;
  selection?: string;
}): Promise<ProposalOutcome> {
  const before = blocksOf(opts.editor.state.doc);
  const userMessage = buildUserMessage(opts.instruction, before, opts.selection);
  const r = await window.officeCopilot.aiTransform({
    action: "agent",
    text: userMessage,
    system: SYSTEM_PROMPT,
    providerId: opts.providerId ?? undefined,
  });
  if (!r.ok) {
    return {
      ok: false,
      before,
      message: r.message,
      needsSetup: r.needsSetup,
      warnings: [],
    };
  }
  const raw = extractJson(r.text);
  if (raw === null) {
    return {
      ok: false,
      before,
      message: "模型没有返回可解析的修改方案，请重试或换个说法。",
      warnings: [],
    };
  }
  const coerced = coerceProposal(raw, before.length);
  if (!coerced.proposal) {
    return { ok: false, before, message: coerced.errors.join("；"), warnings: coerced.errors };
  }
  return { ok: true, proposal: coerced.proposal, before, warnings: coerced.errors };
}

export interface CommitResult {
  ok: boolean;
  message?: string;
  record?: ModificationRecord;
}

/**
 * Apply already-resolved ops as one transaction and log the modification so it
 * can be diffed and undone on its own later.
 */
export async function commitOps(
  editor: Editor,
  opts: {
    ops: CoreOp[];
    before: Block[];
    after: Block[];
    inverses: CoreOp[];
    label: string;
    summary: string;
    docKey: string;
    kind?: "template";
    docLevel?: Record<string, unknown>;
  },
): Promise<CommitResult> {
  if (!opts.ops.length) return { ok: false, message: "没有需要写入的修改。" };
  const exec = executeOps(editor, opts.ops, opts.after);
  if (!exec.ok) return { ok: false, message: exec.reason ?? "写入失败。" };

  const record: ModificationRecord = {
    id: crypto.randomUUID(),
    ts: Date.now(),
    label: opts.label,
    summary: opts.summary,
    ops: opts.ops,
    inverses: opts.inverses,
    before: opts.before,
    after: exec.after,
    undone: false,
    ...(opts.kind ? { kind: opts.kind } : null),
    ...(opts.docLevel ? { docLevel: opts.docLevel } : null),
  };
  const saved = await window.officeCopilot.appendModification({ docKey: opts.docKey, record });
  if (!saved.ok) {
    // The edit is in the document; losing the history entry is not worth
    // rolling the user's document back, so surface it instead.
    return { ok: true, record, message: `修改已写入，但历史记录保存失败：${saved.message ?? "未知错误"}` };
  }
  return { ok: true, record };
}

export interface PlanOutcome {
  ok: boolean;
  message?: string;
  ops?: CoreOp[];
  before?: Block[];
  after?: Block[];
  inverses?: CoreOp[];
  skipped: string[];
}

/** Validate + simulate a proposal against the live document without writing. */
export function simulateProposal(editor: Editor, proposal: Proposal): PlanOutcome {
  const before = blocksOf(editor.state.doc);
  const plan = planEdit(before, proposal, (spec, alt) => diagramResolverFor(spec, alt));
  if (!plan.ok) {
    return { ok: false, message: plan.skipped.join("；") || "提案无法执行。", skipped: plan.skipped };
  }
  return {
    ok: true,
    ops: plan.ops,
    before: plan.before,
    after: plan.after,
    inverses: plan.inverses,
    skipped: plan.skipped,
  };
}

export async function applyProposal(
  editor: Editor,
  proposal: Proposal,
  docKey: string,
): Promise<CommitResult> {
  const plan = simulateProposal(editor, proposal);
  if (!plan.ok) return { ok: false, message: plan.message };
  return commitOps(editor, {
    ops: plan.ops!,
    before: plan.before!,
    after: plan.after!,
    inverses: plan.inverses!,
    label: "AI 修改",
    summary: proposal.summary,
    docKey,
  });
}

/**
 * Roll back one recorded modification. Later unrelated edits do not block it:
 * each inverse is matched by content. Ops whose target the user has since
 * changed are reported and skipped rather than guessed at.
 */
export async function undoModification(
  editor: Editor,
  record: ModificationRecord,
  docKey: string,
): Promise<CommitResult & { failed: string[] }> {
  if (record.undone) {
    return { ok: false, message: "这一条修改已经撤销过了。", failed: [] };
  }
  const current = blocksOf(editor.state.doc);
  const r = revert(current, record.inverses);
  if (!r.applied.length) {
    return {
      ok: false,
      message: `无法撤销：这些段落自那次修改后已被改动。\n${r.failed.join("\n")}`,
      failed: r.failed,
    };
  }
  const exec = executeOps(editor, r.applied, r.ok ? r.after : undefined);
  if (!exec.ok) {
    return { ok: false, message: exec.reason ?? "撤销失败。", failed: r.failed };
  }
  await window.officeCopilot.patchModification({
    docKey,
    id: record.id,
    patch: { undone: true, undoneAt: Date.now() },
  });
  const message = r.ok
    ? undefined
    : `部分撤销（${r.applied.length}/${record.inverses.length}）：${r.failed.join("；")}`;
  return { ok: true, message, failed: r.failed };
}

/* ------------------------------------------------------------------ */
/* Non-agent edits that still need to land in the history              */
/* ------------------------------------------------------------------ */

/**
 * Rewrite the current selection with `rewritten` as block-level setText ops, so
 * a bubble rewrite is undoable like any other AI edit.
 */
export async function commitSelectionRewrite(
  editor: Editor,
  opts: { from: number; to: number; rewritten: string; label: string; docKey: string },
): Promise<CommitResult> {
  const before = blocksOf(editor.state.doc);
  const located = locateAll(editor.state.doc);
  const touched = located.filter(
    (l) => l.inner && l.inner.from < opts.to && l.inner.to > opts.from,
  );
  if (!touched.length) return { ok: false, message: "选区不在可改写的段落内。" };

  const clean = normalise(opts.rewritten);
  const ops: CoreOp[] = [];
  touched.forEach((l, i) => {
    const inner = l.inner!;
    const blockIndex = l.block.index;
    const addr = addressOf(before, blockIndex);
    if (!addr) return;
    const prefix = i === 0 ? l.block.text.slice(0, Math.max(0, opts.from - inner.from)) : "";
    const suffix = i === touched.length - 1 ? l.block.text.slice(Math.max(0, opts.to - inner.from)) : "";
    const text = i === 0 ? normalise(prefix + clean + suffix) : normalise(suffix);
    ops.push({ kind: "setText", addr, text });
  });
  if (!ops.length) return { ok: false, message: "选区不在可改写的段落内。" };

  const run = runOps(before, ops);
  return commitOps(editor, {
    ops: run.applied,
    before,
    after: run.after,
    inverses: run.inverses,
    label: opts.label,
    summary: `${opts.label}：${ops.map(describeCoreOp).join("，")}`,
    docKey: opts.docKey,
  });
}

/** Insert assistant text as a real modification instead of a raw paste. */
export async function commitInsertedText(
  editor: Editor,
  opts: { text: string; docKey: string; at: "cursor" | "end"; label?: string },
): Promise<CommitResult> {
  const before = blocksOf(editor.state.doc);
  const paragraphs = opts.text.split(/\n+/).map((s) => s.trim()).filter(Boolean);
  if (!paragraphs.length) return { ok: false, message: "没有可插入的内容。" };

  let anchorIndex = before.length - 1;
  if (opts.at === "cursor") {
    const pos = editor.state.selection.from;
    const hit = locateAll(editor.state.doc).find((l) => l.inner && l.inner.from <= pos && l.inner.to >= pos);
    if (hit) anchorIndex = hit.block.index;
  }
  // Sequential inserts must stack: each one follows the previous index.
  const aiOps: AiOp[] = paragraphs.map((text, i): AiOp => ({
    op: "insertParagraph",
    after: anchorIndex + i,
    text,
  }));
  const proposal: Proposal = {
    summary: opts.label ?? "AI 插入文案",
    ops: aiOps,
  };
  return applyProposal(editor, proposal, opts.docKey);
}
