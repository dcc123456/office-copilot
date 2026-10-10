import type { Editor } from "@tiptap/react";
import type { Node as PMNode } from "@tiptap/pm/model";
import type { EditorState, Transaction } from "@tiptap/pm/state";
import {
  addressOf,
  runOps,
  type Anchor,
  type Block,
  type CoreOp,
  type JsonNode,
} from "../ai/docEdit.js";
import { commitOps, type CommitResult } from "../ai/agent.js";
import { blocksOf, jsonContentText, locateAll } from "./docBlocks.js";
import { nextBookmarkId } from "./toc.js";

/**
 * 题注 = 标签文字 + 一条 `SEQ` 域 + 说明文字（ECMA-376 §17.9.4 字段指令，微软
 * 「字段代码：SEQ 字段」：`SEQ 图 \* ARABIC` 按标签各自计数，`\* ARABIC` 要十进制）。
 *
 * Word 在编号外面还套一个 `_Ref…` 书签，「插入交叉引用」列的就是这些书签，所以这里
 * 一并写上（⑤.2 的 `REF`/`PAGEREF` 按名字指向它）。编号是**域的缓存结果**：Word 自己
 * 也要按 F9 才重算，我们则用 {@link renumberSequences} 在每次编辑后现算，于是删掉一张图
 * 之后的题注立刻跟着减号。
 *
 * 书签 id 复用 `toc.ts` 的 `nextBookmarkId`：全篇一个计数器，目录的 `_Toc…` 与题注的
 * `_Ref…` 因此不可能撞号。
 */

/** Word「题注」对话框里"位置"那一项。 */
export type CaptionPosition = "above" | "below";

/** 对话框交给引擎的那三项。 */
export interface CaptionOptions {
  /** 标签，如"图"。带空格的标签在指令里要加引号（Word 就这么写）。 */
  label: string;
  /** 编号后面的说明文字，可空。 */
  captionText: string;
  /** 插在选中项的上方还是下方。 */
  position: CaptionPosition;
}

/** 一条待插入题注算好的结果，界面上要显示的就是它。 */
export interface CaptionTarget {
  /** 选中项在块模型里的索引。 */
  index: number;
  /** 这段落在第几号（同标签的 `SEQ` 域按文档顺序数出来的）。 */
  number: number;
  /** 编号外那对书签的 `w:id` 与名字。 */
  bookmark: { id: string; name: string };
}

/** 生成失败的原因码：界面按它拼中文提示，引擎不给文案（AGENTS §5.7）。 */
export type CaptionReject =
  /** 标签是空的：`SEQ` 指令没有序列名，写出去就是一条算不出来的域。 */
  | "bad-label"
  /** 选区所在的那一段给不出块地址（空文档或未建模的块）。 */
  | "no-target"
  /** op 在块模型上模拟不通，通常意味着锚点找不到了。 */
  | "ops-failed";

export interface CaptionPlan {
  ok: boolean;
  reason?: CaptionReject;
  ops: CoreOp[];
  before: Block[];
  after: Block[];
  /** 逆操作：撤销题注就是删掉那一段（书签在段内，一起走）。 */
  inverses: CoreOp[];
  target?: CaptionTarget;
  skipped: string[];
}

/** 中文 Word「题注」对话框的默认标签；用户可以另填。 */
export const CAPTION_LABELS = ["图", "表", "公式"];

/**
 * `SEQ` 域的指令文本（不含首尾空格，docx-io 落盘时补）。
 *
 * @param label 序列名，也就是题注标签。
 * @returns `SEQ 图 \\* ARABIC`；标签含空白时加引号（`SEQ "图 1" \\* ARABIC`）。
 */
export function seqInstr(label: string): string {
  const name = /\s/.test(label) ? `"${label}"` : label;
  return `SEQ ${name} \\* ARABIC`;
}

/**
 * 从指令里取出序列名，认不出就返回 null（不是 `SEQ` 域，或者写坏了）。
 *
 * @param instr `field` 节点的 `instr` 缓存值，可能带 `\* ARABIC` 等开关。
 * @returns 序列名原样（带引号的去引号，不加引号的原样）。读侧不做单位与颜色那类换算，
 *          标签文字是内容不是格式，因此不改大小写。
 */
export function seqLabelOfInstr(instr: string): string | null {
  const m = /^SEQ\s+(?:"([^"]*)"|(\S+))/i.exec(instr.trim());
  if (!m) return null;
  const raw = m[1] ?? m[2] ?? "";
  return raw.trim() ? raw : null;
}

/**
 * 一段题注的 PM JSON。
 *
 * 节点顺序对着 Word 写盘的样子：标签文字、书签起点、`SEQ` 域、书签终点、说明文字。
 * 书签只裹编号，所以交叉引用取"只有标签和编号"时能整段命中。
 *
 * @param opts     对话框选项。
 * @param number   编号（写进域的缓存结果；真正的号数由 {@link renumberSequences} 保证）。
 * @param bookmark 编号外那对书签的 id 与名字。
 * @returns 段落节点 JSON；`attrs` 故意留空——Word 的题注版式来自 `Caption` 样式，
 *          而文档里往往没有这条样式定义，写 `w:pStyle` 会是悬空引用（同 ③.2 的取舍）。
 */
export function captionJson(
  opts: Pick<CaptionOptions, "label" | "captionText">,
  number: number,
  bookmark: { id: string; name: string },
): JsonNode {
  const content: unknown[] = [
    { type: "text", text: `${opts.label} ` },
    { type: "bookmarkStart", attrs: { id: bookmark.id, name: bookmark.name } },
    {
      type: "field",
      attrs: { kind: "SEQ", instr: seqInstr(opts.label), display: String(number) },
    },
    { type: "bookmarkEnd", attrs: { id: bookmark.id } },
  ];
  if (opts.captionText) content.push({ type: "text", text: ` ${opts.captionText}` });
  return { type: "paragraph", content };
}

/**
 * 按文档顺序给每条 `SEQ` 域算出应有的号数，并把引用这些编号的 `REF` 域一起刷成同一个数。
 *
 * 每个标签各数各的（Word 就是按序列名分开计数），所以一张图一张表交替出现也不会互相打乱。
 * `REF _Ref…` 指向的是题注里那个书签（⑤.2），书签包住的正是这条 `SEQ` 域，所以它的缓存
 * 结果就是这条域的号数——Word 靠按 F9 一起更新，我们在同一次遍历里顺手算完。
 *
 * @param doc 当前文档。
 * @returns 需要改缓存结果的域：绝对位置 + 换过 `display` 后的 attrs；号数没变的域不在其中。
 */
export function captionNumberPatches(doc: PMNode): { pos: number; attrs: Record<string, unknown> }[] {
  const counters: Record<string, number> = {};
  /** `_Ref…` 书签名 → 它包住的 `SEQ` 域号数；交叉引用按这个名字取值，与先后无关。 */
  const numberByBookmark: Record<string, string> = {};
  const refs: { pos: number; attrs: Record<string, unknown>; name: string }[] = [];
  const open = new Map<string, string>();
  const out: { pos: number; attrs: Record<string, unknown> }[] = [];

  doc.descendants((node, pos) => {
    if (node.type.name === "bookmarkStart") {
      open.set(String(node.attrs.id ?? ""), String(node.attrs.name ?? ""));
      return;
    }
    if (node.type.name === "bookmarkEnd") {
      open.delete(String(node.attrs.id ?? ""));
      return;
    }
    if (node.type.name !== "field") return;
    const attrs = node.attrs as Record<string, unknown>;
    const label = seqLabelOfInstr(String(attrs.instr ?? ""));
    if (label !== null) {
      const seen = (counters[label] ?? 0) + 1;
      counters[label] = seen;
      const want = String(seen);
      for (const name of open.values()) if (name.startsWith("_Ref")) numberByBookmark[name] = want;
      if (String(attrs.display ?? "") !== want) out.push({ pos, attrs: { ...attrs, display: want } });
      return;
    }
    const refName = /^REF\s+(\S+)/.exec(String(attrs.instr ?? "").trim())?.[1];
    if (refName) refs.push({ pos, attrs, name: refName });
  });

  // 引用可能在目标之前（"以下"那种），所以必须等编号全算完再回头刷它们。
  for (const ref of refs) {
    const want = numberByBookmark[ref.name];
    if (want === undefined || String(ref.attrs.display ?? "") === want) continue;
    out.push({ pos: ref.pos, attrs: { ...ref.attrs, display: want } });
  }
  return out;
}

/**
 * 把正文里所有 `SEQ` 域、以及引用这些编号的 `REF` 域的缓存结果，刷成当前文档顺序下的号数。
 *
 * 编号是派生显示值，不进撤销栈：Ctrl+Z 要撤的是用户那次编辑，不是后面题注跟着挪号。
 *
 * @param state 编辑器状态。
 * @returns 要派发的事务；没有任何域需要改号时返回 null（绝大多数按键都是这条路径）。
 */
export function renumberSequences(state: EditorState): Transaction | null {
  const patches = captionNumberPatches(state.doc);
  if (!patches.length) return null;
  let tr = state.tr;
  // setNodeMarkup 不改节点尺寸，所以这些绝对位置在同一份 doc 上始终有效。
  for (const patch of patches) tr = tr.setNodeMarkup(patch.pos, undefined, patch.attrs);
  return tr.setMeta("addToHistory", false);
}

/**
 * 选区落在哪一个顶层块上，并算好这段该是第几号。
 *
 * @param doc      当前文档。
 * @param selFrom  选区起点（`Selection.from`）。
 * @param opts     对话框选项；`position` 决定新题注排在选中项之前还是之后，号数因此差一。
 * @returns 目标块索引、号数与书签；空文档或选区落在块模型给不出地址的地方时 null。
 */
export function captionTargetFor(
  doc: PMNode,
  selFrom: number,
  opts: Pick<CaptionOptions, "label" | "position">,
): CaptionTarget | null {
  const label = opts.label.trim();
  if (!label) return null;
  const blocks = blocksOf(doc);
  const target = locateAll(doc).find((l) => l.from <= selFrom && selFrom < l.to);
  if (!target || !addressOf(blocks, target.block.index)) return null;
  // "下方"要跑到选中项之后，所以选中项自己那一段（如果它就是同标签的题注）不计入。
  const boundary = opts.position === "above" ? target.from : target.to;
  const id = nextBookmarkId(doc);
  return {
    index: target.block.index,
    number: sequenceCountBefore(doc, label, boundary) + 1,
    bookmark: { id, name: `_Ref${id}` },
  };
}

/**
 * 插入点之前已经有几条同标签的 `SEQ` 域。
 *
 * @param doc   当前文档。
 * @param label 序列名。
 * @param pos   插入边界（文档坐标）。
 * @returns 之前的条数；这就是新题注的号数减一。
 */
function sequenceCountBefore(doc: PMNode, label: string, pos: number): number {
  let count = 0;
  doc.descendants((node, nodePos) => {
    if (nodePos >= pos) return false;
    if (node.type.name !== "field") return;
    if (seqLabelOfInstr(String(node.attrs.instr ?? "")) === label) count += 1;
  });
  return count;
}

/**
 * 生成一次题注插入的 op，并在块模型上模拟出结果——和 `planTocInsertion` 同一套语义：
 * 实际落地与模拟不一致时整批放弃（AGENTS §2.4）。
 *
 * @param editor 桌面编辑器。
 * @param opts   对话框选项。
 */
export function planCaptionInsertion(editor: Editor, opts: CaptionOptions): CaptionPlan {
  const state = editor.state;
  const before = blocksOf(state.doc);
  const empty: CaptionPlan = {
    ok: false,
    ops: [],
    before,
    after: before,
    inverses: [],
    skipped: [],
  };
  if (!opts.label.trim()) return { ...empty, reason: "bad-label" };

  const target = captionTargetFor(state.doc, state.selection.from, opts);
  if (!target) return { ...empty, reason: "no-target" };

  const anchor: Anchor =
    target.index === 0 && opts.position === "above"
      ? { at: "start" }
      : {
          at: "after",
          after: addressOf(before, opts.position === "above" ? target.index - 1 : target.index)!,
        };

  const json = captionJson(opts, target.number, target.bookmark);
  const ops: CoreOp[] = [
    {
      kind: "insertBlock",
      anchor,
      block: {
        type: "paragraph",
        // 块模型按纯文本寻址，域芯片与书签都不贡献文字。
        text: jsonContentText(json),
        json,
      },
    },
  ];

  const simulated = runOps(before, ops);
  if (simulated.skipped.length) {
    return { ...empty, reason: "ops-failed", target, skipped: simulated.skipped };
  }
  return {
    ok: true,
    ops: simulated.applied,
    before,
    after: simulated.after,
    inverses: simulated.inverses,
    target,
    skipped: [],
  };
}

/**
 * 落地一次题注插入，并写进修改历史（撤销按内容寻址，不牵连用户后来的手工编辑）。
 *
 * @param editor 桌面编辑器。
 * @param docKey 修改历史的文档键，即文件路径。
 * @param opts   对话框选项。
 * @returns 成功时带 modlog 记录；失败时 `reason` 是拒绝码，文案由界面组装。
 */
export async function insertCaption(
  editor: Editor,
  docKey: string,
  opts: CaptionOptions,
): Promise<CommitResult & { number: number; reason?: CaptionReject }> {
  const plan = planCaptionInsertion(editor, opts);
  if (!plan.ok) return { ok: false, number: 0, reason: plan.reason };
  const number = plan.target?.number ?? 0;
  const summary = `插入题注（${opts.label.trim()} ${number}）`;
  const r = await commitOps(editor, {
    ops: plan.ops,
    before: plan.before,
    after: plan.after,
    inverses: plan.inverses,
    label: "插入题注",
    summary,
    docKey,
  });
  return { ...r, number };
}
