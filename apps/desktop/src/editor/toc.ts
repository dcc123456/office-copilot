import { cssPxToPt, nextAnnotationId, type TabStop, type TabStopLeader } from "@office-copilot/docx-io";
import type { Editor } from "@tiptap/react";
import type { Node as PMNode } from "@tiptap/pm/model";
import type { EditorState, Transaction } from "@tiptap/pm/state";
import {
  addressOf,
  runOps,
  type Address,
  type Anchor,
  type Block,
  type CoreOp,
  type JsonNode,
} from "../ai/docEdit.js";
import { commitOps, type CommitResult } from "../ai/agent.js";
import { blocksOf, jsonContentText, locateAll } from "./docBlocks.js";
import { pageNumberOf } from "./extensions/PagedLayout.js";

/**
 * 本地生成 Word 的「目录」：一条 `TOC` 域把条目段夹在中间，每条条目是
 * `标题 ⇥ PAGEREF 页码`，页码按编辑器自己的分页结果缓存。
 *
 * 与 Word 的对应关系（一手依据：ECMA-376 §17.9.7 复合域、微软「字段代码：TOC 字段」）：
 * - 域的 begin+instrText+separate 落在首段开头、end 落在末段之后，中间是普通段落，
 *   这正是 `fieldBegin`/`fieldEnd` 这一对边界节点建模的形态（排名 3 ③.1）。
 * - 页码是段内 `PAGEREF` 子域，指向标题上的 `_Toc…` 书签；Word 自己生成的条目还套
 *   `w:hyperlink`，那一层现在不写（编辑器里目录条目会因此显成普通文字色，
 *   Word 按 F9 重算时会把超链接补上）。
 * - 条目靠右对齐制表位 + 前导符对齐页码，所以排名 5 ① 的制表位数据层是这里的前提。
 */

/**
 * 目录插在哪。Word 的「目录」是在光标处插入的，所以只有这两种：文档开头，
 * 或选区所在那一段之后。{ at: "end" } 不在这张表里——倒序插入靠的是"每次都插在同一
 * 个锚点后面"，文档末尾没有锚点可复用了，硬做要多写一条排序分支，不值。
 */
export type TocPosition = { at: "start" } | { at: "after"; after: Address };

/** Word「目录」对话框里可勾的那几项。 */
export interface TocOptions {
  /** 收录的最高大纲级别（`\o "1-N"`），1–9。 */
  levels: number;
  /** 显示页码：关掉就不写制表符与 `PAGEREF`（Word 的"显示页码"复选框）。 */
  showPageNumbers: boolean;
  /** 前导符，Word 默认点线。 */
  leader: TabStopLeader;
  /** 正文区宽度，CSS px（`SectionGeometry.contentWidth`）；换算是 `cssPxToPt`，不在本文件里。 */
  contentWidthPx: number;
  /** 插入位置；缺省 = 文档最前面。 */
  position?: TocPosition;
}

/** 一条目录条目对应的标题。 */
export interface TocEntry {
  /** blocksOf 的块索引，书签 op 按它取地址。 */
  index: number;
  /** 大纲级别，1 起。 */
  level: number;
  /** 标题文字：条目显示它，`PAGEREF` 按书签名指向标题段。 */
  text: string;
  /** `_Toc…` 书签名。 */
  bookmark: string;
  /** OOXML 的 `w:id`，书签与域都用它配对。 */
  id: string;
  /** 本地分页测出的页码；0 = 分页计划还不存在，缓存结果先留空。 */
  page: number;
}

/** 生成失败的原因码：界面按它拼中文提示，引擎不给文案（AGENTS §5.7）。 */
export type TocReject =
  /** 文档里一个够格的标题都没有。 */
  | "no-headings"
  /** 级别不在 1–9。 */
  | "bad-levels"
  /** op 在块模型上模拟不通，通常意味着锚点找不到了。 */
  | "ops-failed";

export interface TocPlan {
  ok: boolean;
  reason?: TocReject;
  /** 将要落地的 op：先给每个标题打书签，再倒序把条目插到同一个位置之后。 */
  ops: CoreOp[];
  before: Block[];
  after: Block[];
  /** 逆操作，撤销目录时按顺序跑：删条目、撤书签。 */
  inverses: CoreOp[];
  entries: TocEntry[];
  /** op 模拟阶段被挡下的原因；非空说明这次不能写。 */
  skipped: string[];
}

/** 每级条目往右让多少磅，Word 的 TOC 2/3 样式就是这个台阶。 */
const LEVEL_INDENT_PT = 20;

/**
 * 段落该算进目录的第几级：标题按 `level`，正文段按 Word 段落对话框里的
 * 「大纲级别」（`\u` 开关收的就是它，0 起）。
 *
 * @param node 顶层段落或标题节点。
 * @returns 级别（1 起），不属于标题层级时 null。
 */
export function tocLevelOf(node: PMNode): number | null {
  const attrs = node.attrs as Record<string, unknown>;
  if (node.type.name === "heading") {
    const level = Number(attrs.level ?? 1);
    return Number.isFinite(level) && level >= 1 ? level : null;
  }
  if (node.type.name !== "paragraph") return null;
  const outline = attrs.outlineLevel;
  if (typeof outline !== "number" || outline < 0 || outline > 8) return null;
  return outline + 1;
}

/**
 * 下一个可用的 OOXML `w:id`：文档里已有书签的最大 id 加一。
 *
 * 书签名由它拼出来（`_Toc${id}`），所以只要 id 不重复，名字就不可能撞车——
 * 反复插入目录时，Word 自己那条 `_Toc` 也不会被我们改写。
 *
 * @param doc 当前文档。
 * @returns 整数字符串形式的下一个 id；文档里没有书签时是 "1"。
 */
export function nextBookmarkId(doc: PMNode): string {
  const ids: string[] = [];
  doc.descendants((node) => {
    if (node.type.name === "bookmarkStart") ids.push(String(node.attrs.id ?? ""));
    return true;
  });
  return nextAnnotationId(ids, 1);
}

/**
 * 目录条目的来源：文档里够格的标题与带大纲级别的段落，按文档顺序。
 *
 * @param doc     当前文档。
 * @param levels  收录的最高级别（含）。
 * @param pageOf  由文档位置求页码（1 起，测不出来返回 0）。
 * @returns 条目列表，书签名与 id 已分配好。
 */
export function collectTocEntries(
  doc: PMNode,
  levels: number,
  pageOf: (pos: number) => number,
): TocEntry[] {
  let id = Number(nextBookmarkId(doc));
  const entries: TocEntry[] = [];
  for (const target of locateAll(doc)) {
    const node = target.textNode;
    if (!node) continue;
    const level = tocLevelOf(node);
    if (!level || level > levels) continue;
    const text = node.textContent;
    if (!text.trim()) continue;
    const bookmark = `_Toc${id}`;
    entries.push({
      index: target.block.index,
      level,
      text,
      bookmark,
      id: String(id),
      page: pageOf(target.textPos),
    });
    id += 1;
  }
  return entries;
}

/**
 * 域指令，与 Word 自己插入目录时写的那条一致（不含首尾空格，docx-io 落盘时补）。
 *
 * @param levels 收录的最高级别。
 * @returns `TOC \\o "1-N" \\h \\z \\u`。
 */
export function tocFieldInstr(levels: number): string {
  return `TOC \\o "1-${levels}" \\h \\z \\u`;
}

/**
 * 一条条目段落的 PM JSON。
 *
 * @param entry  这条目录条目。
 * @param opts   对话框选项；`contentWidthPx` 在这里换成磅当制表位位置。
 * @param edge   是不是首条/末条：`fieldBegin` 挂在首条开头，`fieldEnd` 挂在末条末尾。
 * @returns 段落节点 JSON，其 `textContent` 即块模型用来寻址的文字。
 */
export function tocEntryJson(
  entry: TocEntry,
  opts: TocOptions,
  edge: { first: boolean; last: boolean },
): JsonNode {
  const attrs: Record<string, unknown> = {};
  // 制表位只为页码而存在：不显示页码时写它，等于给一段没有制表符的文字加格式。
  if (opts.showPageNumbers) {
    const stop: TabStop = {
      positionPt: cssPxToPt(opts.contentWidthPx),
      alignment: "right",
      leader: opts.leader,
    };
    attrs.tabStops = [stop];
  }
  if (entry.level > 1) attrs.indent = (entry.level - 1) * LEVEL_INDENT_PT;

  const children: unknown[] = [];
  if (edge.first) {
    children.push({
      type: "fieldBegin",
      attrs: { kind: "TOC", instr: tocFieldInstr(opts.levels), isSeparated: true },
    });
  }
  // 页码是标题文字后的一项，所以制表符跟着标题；不显示页码时整条就是标题本身。
  const label = opts.showPageNumbers ? `${entry.text}\t` : entry.text;
  children.push({ type: "text", text: label });
  if (opts.showPageNumbers) {
    children.push({
      type: "field",
      attrs: {
        kind: "PAGEREF",
        instr: `PAGEREF ${entry.bookmark} \\h`,
        // 分页还没测出来时留空：Word 打开后重算，本 app 也能在刷新页码时补上。
        display: entry.page > 0 ? String(entry.page) : "",
      },
    });
  }
  if (edge.last) children.push({ type: "fieldEnd" });

  return { type: "paragraph", attrs, content: children };
}

/**
 * 选区所在的那一段，作为「插在这一段之后」的位置——Word 在光标处插入目录。
 *
 * 按顶层块算，所以表格、图片里光标也能当锚点（`addressOf` 认这两种地址）；
 * 选区跨段时取起始那段。
 *
 * @param editor 桌面编辑器。
 * @returns 位置描述；光标落在块模型给不出地址的那一段（例如未建模的顶层节点）时
 *          退回 `{ at: "start" }`，界面因此永远拿得到一个可用位置。
 */
export function tocAnchorForSelection(editor: Editor): TocPosition {
  const { doc, selection } = editor.state;
  let index = -1;
  doc.forEach((node, offset, i) => {
    if (offset <= selection.from && selection.from < offset + node.nodeSize) index = i;
  });
  const addr = index >= 0 ? addressOf(blocksOf(doc), index) : null;
  return addr ? { at: "after", after: addr } : { at: "start" };
}

/**
 * 生成一次目录插入的 op，并在块模型上模拟出结果——和 `planEdit` 同一套语义：
 * 实际落地与模拟不一致时 `executeOps` 会整批放弃（AGENTS §2.4）。
 *
 * @param editor 桌面编辑器。
 * @param opts   目录对话框选项。
 */
export function planTocInsertion(editor: Editor, opts: TocOptions): TocPlan {
  const state = editor.state;
  const before = blocksOf(state.doc);
  const empty: TocPlan = {
    ok: false,
    ops: [],
    before,
    after: before,
    inverses: [],
    entries: [],
    skipped: [],
  };
  if (!Number.isInteger(opts.levels) || opts.levels < 1 || opts.levels > 9) {
    return { ...empty, reason: "bad-levels" };
  }

  const entries = collectTocEntries(state.doc, opts.levels, (pos) => pageNumberOf(state, pos));
  if (!entries.length) return { ...empty, reason: "no-headings" };

  const anchor: Anchor = opts.position ?? { at: "start" };
  const ops: CoreOp[] = entries.map(
    (entry): CoreOp => ({
      kind: "addBookmark",
      addr: addressOf(before, entry.index)!,
      id: entry.id,
      name: entry.bookmark,
    }),
  );
  // 倒序插入：每条都插在同一个锚点之后，最后落地的第一条才在最前面。
  for (let i = entries.length - 1; i >= 0; i--) {
    const entry = entries[i]!;
    const json = tocEntryJson(entry, opts, { first: i === 0, last: i === entries.length - 1 });
    ops.push({
      kind: "insertBlock",
      anchor,
      block: {
        type: "paragraph",
        // 块模型按纯文本寻址，域芯片自己的文字不计入，所以这里只能取 JSON 里的文字。
        text: jsonContentText(json),
        json,
      },
    });
  }

  const simulated = runOps(before, ops);
  // 任何一条 op 在模拟阶段落不下都不写：真实的 `executeOps` 还会按 `after`
  // 复核一遍，但那时已经发过 dispatch 的前提不成立了。
  if (simulated.skipped.length) {
    return { ...empty, reason: "ops-failed", entries, skipped: simulated.skipped };
  }
  return {
    ok: true,
    ops: simulated.applied,
    before,
    after: simulated.after,
    inverses: simulated.inverses,
    entries,
    skipped: [],
  };
}

/**
 * 落地一次目录插入，并写进修改历史（撤销按内容寻址，不会牵连用户后来的手工编辑）。
 *
 * @param editor 桌面编辑器。
 * @param docKey 修改历史的文档键，即文件路径。
 * @param opts   目录对话框选项。
 * @returns 成功时带 modlog 记录与条目数；失败时 `reason` 是拒绝码，文案由界面组装。
 */
export async function insertToc(
  editor: Editor,
  docKey: string,
  opts: TocOptions,
): Promise<CommitResult & { entries: number; reason?: TocReject }> {
  const plan = planTocInsertion(editor, opts);
  if (!plan.ok) {
    return { ok: false, entries: plan.entries.length, reason: plan.reason };
  }
  const summary = `插入目录（${plan.entries.length} 条）`;
  const r = await commitOps(editor, {
    ops: plan.ops,
    before: plan.before,
    after: plan.after,
    inverses: plan.inverses,
    label: "插入目录",
    summary,
    docKey,
  });
  return { ...r, entries: plan.entries.length };
}

/**
 * 把 `PAGEREF` 里缓存的页码刷成分页现在说的数——目录条目与交叉引用都靠它。
 *
 * 目录本身占版面，插进去之后标题会整体往后移，第一遍写的页码因此是旧的；分页重排
 * 完成后跑一次，缓存结果就和 Word 打开后按 F9 的结果一致了。只改 `field` 的
 * `display`（域的缓存值本来就该被重算），不动任何文字。
 *
 * @param state  分页已更新后的编辑器状态。
 * @param pageOf 由文档位置求页码（1 起，测不出来返回 0）；调用方传 PagedLayout 的结果。
 * @returns 要派发的事务，已经把 `addToHistory` 关掉（页码是派生数据，不是用户的编辑）；
 *          没有条目需要改时返回 null。
 */
export function fieldPageNumberRefresh(
  state: EditorState,
  pageOf: (pos: number) => number,
): Transaction | null {
  const pages = bookmarkPages(state.doc, pageOf);
  if (!pages.size) return null;
  let tr = state.tr;
  let changed = 0;
  for (const target of locateAll(state.doc)) {
    const node = target.textNode;
    const inner = target.inner;
    if (!node || !inner) continue;
    // setNodeMarkup 不改节点尺寸，所以这些偏移在同一份 doc 上始终有效。
    for (const patch of pagerefPatchNodes(node, pages)) {
      tr = tr.setNodeMarkup(inner.from + patch.offset, undefined, patch.attrs);
      changed += 1;
    }
  }
  return changed ? tr.setMeta("addToHistory", false) : null;
}

/**
 * 插完目录或交叉引用后盯着分页收敛，把 `PAGEREF` 里缓存的页码刷成分页现在说的数。
 *
 * 不能只刷一次：`PagedLayout` 的测量在 requestAnimationFrame 里跑，字体到位、
 * 上一轮插入的分页符落定之后才出最终计划，第一遍读到的还是旧页码。所以按帧检查，
 * 连续 `settleMs` 没有页码要改就收手，`deadlineMs` 是无条件兜底。
 *
 * @param editor 桌面编辑器。
 * @param opts.settleMs   多久没有再改动就停止观察，默认 600ms。
 * @param opts.deadlineMs 最长观察多久，默认 5000ms。
 * @returns 停止观察的函数；对话框关闭或组件卸载时调用，免得向已销毁的编辑器派发事务。
 */
export function watchFieldPageNumbers(
  editor: Editor,
  opts: { settleMs?: number; deadlineMs?: number } = {},
): () => void {
  const settleMs = opts.settleMs ?? 600;
  const deadlineMs = opts.deadlineMs ?? 5000;
  const startedAt = Date.now();
  let lastChangeAt = startedAt;
  let stopped = false;
  let frame = 0;
  const stop = () => {
    stopped = true;
    cancelAnimationFrame(frame);
  };
  const tick = () => {
    if (stopped) return;
    if (editor.isDestroyed) return stop();
    const now = Date.now();
    const tr = fieldPageNumberRefresh(editor.state, (pos) => pageNumberOf(editor.state, pos));
    if (tr) {
      editor.view.dispatch(tr);
      lastChangeAt = now;
    }
    if (now - lastChangeAt >= settleMs || now - startedAt >= deadlineMs) return stop();
    frame = requestAnimationFrame(tick);
  };
  frame = requestAnimationFrame(tick);
  return stop;
}

/**
 * 书签名 → 它所在段落的页码。
 *
 * 不按前缀过滤：目录条目的 `_Toc…` 与题注的 `_Ref…` 都要能被 `PAGEREF` 查到——
 * Word 的交叉引用同样只对书签名负责，不管这个名字是谁生成的。
 *
 * @param doc    当前文档。
 * @param pageOf 由段落起始位置求页码。
 * @returns 文档里出现过的每个书签名；求不到页码（分页还没测出来）的段整条跳过。
 */
function bookmarkPages(doc: PMNode, pageOf: (pos: number) => number): Map<string, number> {
  const out = new Map<string, number>();
  doc.forEach((block, pos) => {
    const names: string[] = [];
    block.descendants((node) => {
      if (node.type.name === "bookmarkStart") {
        const name = String(node.attrs.name ?? "");
        if (name) names.push(name);
      }
      return true;
    });
    const page = pageOf(pos);
    if (page > 0) for (const name of names) out.set(name, page);
  });
  return out;
}

/**
 * 一段里的 `PAGEREF` 域该改成的新缓存值。
 *
 * @param node  段落节点。
 * @param pages 书签名 → 页码。
 * @returns 偏移（相对段落内容起点）与新 attrs；页码没变的域不在其中。
 */
function pagerefPatchNodes(
  node: PMNode,
  pages: Map<string, number>,
): { offset: number; attrs: Record<string, unknown> }[] {
  const out: { offset: number; attrs: Record<string, unknown> }[] = [];
  node.forEach((child, offset) => {
    if (child.type.name !== "field") return;
    const attrs = child.attrs as Record<string, unknown>;
    const name = /^PAGEREF\s+(\S+)/.exec(String(attrs.instr ?? ""))?.[1];
    if (!name) return;
    const page = pages.get(name);
    if (!page) return;
    const display = String(page);
    if (String(attrs.display ?? "") === display) return;
    out.push({ offset, attrs: { ...attrs, display } });
  });
  return out;
}
