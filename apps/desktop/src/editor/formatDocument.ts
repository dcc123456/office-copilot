import type { Editor } from "@tiptap/react";
import { blocksOf } from "../editor/docBlocks.js";
import type { Block, Proposal, StylePatch } from "../ai/docEdit.js";
import { styleUnchanged } from "../ai/docEdit.js";
import { commitOps, simulateProposal, type CommitResult } from "../ai/agent.js";
import { completeRoles, ROLE_LABELS, type DocTemplate, type TemplateRole } from "./templates.js";

/**
 * Deterministic "AI 排版" — safe Chinese document conventions without an LLM.
 * Expressed as ordinary edit ops so the result is diffable in the history panel
 * and undoable on its own, exactly like an agent-proposed change.
 */

const BODY_FONT = '"Microsoft YaHei", "PingFang SC", "微软雅黑", sans-serif';

function headingStyle(level: number): StylePatch {
  return {
    align: level === 1 ? "center" : null,
    bold: true,
    fontSize: level === 1 ? "22px" : level === 2 ? "18px" : "16px",
    fontFamily: BODY_FONT,
  };
}

export function typographyProposal(blocks: readonly Block[]): Proposal {
  const ops: Proposal["ops"] = [];
  const summaryBits: string[] = [];

  for (const b of blocks) {
    if (b.type === "heading") {
      const level = Number(b.style.headingLevel ?? 1);
      const want = headingStyle(level);
      const unchanged =
        b.style.bold === want.bold &&
        b.style.fontSize === want.fontSize &&
        b.style.fontFamily === want.fontFamily &&
        (level === 1 ? b.style.align === "center" : !b.style.align);
      if (unchanged) continue;
      ops.push({ op: "setStyle", index: b.index, style: want });
      summaryBits.push(`H${level}`);
    } else if (b.type === "paragraph") {
      // Leave paragraphs the user already formatted by hand alone. Colour is not
      // a reason to skip any more: a block-wide font patch keeps each run's own
      // colour, so a red word survives. Only a hand-set size or font does, and a
      // block where the runs disagree on either.
      const handSet = b.mixed?.includes("fontSize") || b.mixed?.includes("fontFamily");
      if (handSet || b.style.fontSize || b.style.fontFamily) continue;
      if (!b.text.trim()) continue;
      ops.push({
        op: "setStyle",
        index: b.index,
        style: { fontSize: "14px", fontFamily: BODY_FONT },
      });
      summaryBits.push("正文");
    }
  }

  const counted = summaryBits.reduce<Record<string, number>>((acc, k) => {
    acc[k] = (acc[k] ?? 0) + 1;
    return acc;
  }, {});
  const summary = Object.entries(counted)
    .map(([k, n]) => `${k} ×${n}`)
    .join("，");

  return {
    summary: summary ? `统一格式：${summary}` : "格式已符合规范",
    ops,
  };
}

export async function runTypography(
  editor: Editor,
  docKey: string,
): Promise<CommitResult & { touched: number; skipped: string[] }> {
  const proposal = typographyProposal(blocksOf(editor.state.doc));
  if (!proposal.ops.length) {
    return { ok: false, message: "文档格式已经符合规范，无需调整。", touched: 0, skipped: [] };
  }
  const plan = simulateProposal(editor, { ...proposal, summary: `自动排版（${proposal.summary}）` });
  if (!plan.ok) {
    return { ok: false, message: plan.message, touched: 0, skipped: plan.skipped };
  }
  const r = await commitOps(editor, {
    ops: plan.ops!,
    before: plan.before!,
    after: plan.after!,
    inverses: plan.inverses!,
    label: "自动排版",
    summary: `自动排版（${proposal.summary}）`,
    docKey,
  });
  return { ...r, touched: plan.ops!.length, skipped: plan.skipped };
}

/* ------------------------------------------------------------------ */
/* Template-driven 美化排版                                             */
/* ------------------------------------------------------------------ */

const ROLE_ORDER: readonly TemplateRole[] = [
  "title",
  "h1",
  "h2",
  "h3",
  "h4",
  "h5",
  "h6",
  "body",
];

/** 标题归属：哪个块拿模板的 title 待遇，哪些封面行原样不动。 */
interface TitleSlot {
  title: number;
  /** 保持作者原版式的封面行索引。 */
  skip: readonly number[];
}

const NO_TITLE: TitleSlot = { title: -1, skip: [] };

/**
 * 文档开头的连续居中段落构成封面带：第一行就是标题，其余（副标题、许可、联系方式）
 * 没有对应角色，硬塞成正文会把作者排好的版式改小、加缩进。带里没有居中标题时才沿用
 * 老规则（第一个一级标题当标题），所以遇到一级标题就收带，免得它的同级被区别对待。
 */
function titleSlot(blocks: readonly Block[]): TitleSlot {
  const band: number[] = [];
  for (const b of blocks) {
    if (!b.text.trim()) continue;
    if (b.type !== "paragraph" || b.style.align !== "center") break;
    band.push(b.index);
  }
  if (band.length) return { title: band[0]!, skip: band.slice(1) };
  const first = blocks.find(
    (b) => b.type === "heading" && Number(b.style.headingLevel ?? 1) === 1,
  );
  return first ? { title: first.index, skip: [] } : NO_TITLE;
}

/** 取文档标题文字，供带页眉的模板当页眉内容；判不出标题时返回空串。 */
export function documentTitle(editor: Editor): string {
  const blocks = blocksOf(editor.state.doc);
  const i = titleSlot(blocks).title;
  return blocks.find((b) => b.index === i)?.text.trim() ?? "";
}

/** 块对应的模板角色；封面带里除标题以外的行返回 null，即不产生任何 op。 */
function roleOf(b: Block, slot: TitleSlot): TemplateRole | null {
  if (slot.skip.includes(b.index)) return null;
  if (b.index === slot.title) return "title";
  if (b.type === "heading") {
    const level = Number(b.style.headingLevel ?? 1);
    if (level >= 6) return "h6";
    if (level >= 5) return "h5";
    if (level >= 4) return "h4";
    if (level >= 3) return "h3";
    return level === 2 ? "h2" : "h1";
  }
  return b.type === "paragraph" ? "body" : null;
}

/**
 * Drop the run-level keys this block's runs disagree on. `applyCore` refuses
 * the whole op when a patch touches one, so pruning keeps the rest of the
 * template (spacing, indent, alignment) instead of losing the paragraph.
 */
function pruneForBlock(patch: StylePatch, mixed: Block["mixed"]): StylePatch {
  if (!mixed?.length) return patch;
  const out: StylePatch = {};
  for (const [key, value] of Object.entries(patch)) {
    if (mixed.includes(key as never)) continue;
    (out as Record<string, unknown>)[key] = value;
  }
  return out;
}

export function templateProposal(t: DocTemplate, blocks: readonly Block[]): Proposal {
  const ops: Proposal["ops"] = [];
  const counts = new Map<TemplateRole, number>();
  const slot = titleSlot(blocks);
  const roles = completeRoles(t);

  for (const b of blocks) {
    const role = roleOf(b, slot);
    if (!role) continue;
    const want = roles[role];
    if (!want || !b.text.trim()) continue;
    const patch = pruneForBlock(want, b.mixed);
    if (!Object.keys(patch).length) continue;
    if (styleUnchanged(b.style, patch)) continue;
    ops.push({ op: "setStyle", index: b.index, style: patch });
    counts.set(role, (counts.get(role) ?? 0) + 1);
  }

  const summary = ROLE_ORDER.filter((r) => counts.has(r))
    .map((r) => `${ROLE_LABELS[r]} ×${counts.get(r)}`)
    .join("，");

  return { summary: summary ? `统一格式：${summary}` : "格式已符合规范", ops };
}

export async function runTemplate(
  t: DocTemplate,
  editor: Editor,
  docKey: string,
): Promise<CommitResult & { touched: number; skipped: string[] }> {
  const label = `美化排版 · ${t.label}`;
  const proposal = templateProposal(t, blocksOf(editor.state.doc));
  if (!proposal.ops.length) {
    return { ok: false, message: "文档已符合该模板的格式，无需调整。", touched: 0, skipped: [] };
  }
  const plan = simulateProposal(editor, { ...proposal, summary: `${label}（${proposal.summary}）` });
  if (!plan.ok) {
    return { ok: false, message: plan.message, touched: 0, skipped: plan.skipped };
  }
  const r = await commitOps(editor, {
    ops: plan.ops!,
    before: plan.before!,
    after: plan.after!,
    inverses: plan.inverses!,
    label,
    summary: `${label}（${proposal.summary}）`,
    docKey,
  });
  return { ...r, touched: plan.ops!.length, skipped: plan.skipped };
}
