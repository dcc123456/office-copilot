import { Extension } from "@tiptap/core";
import type { CommandProps } from "@tiptap/core";
import { nextAnnotationId } from "@office-copilot/docx-io";
import { isHistoryTransaction } from "@tiptap/pm/history";
import { Fragment, type Mark, type Node as PMNode } from "@tiptap/pm/model";
import { Plugin, PluginKey, TextSelection, type EditorState, type Transaction } from "@tiptap/pm/state";
import { Mapping, ReplaceStep, type StepMap } from "@tiptap/pm/transform";

/**
 * 修订开关（排名 4 ⑥）：开着的时候，"本次编辑"落成 Word 的 `w:ins` / `w:del`。
 *
 * 做法是**在事务落地以后补一笔**（`appendTransaction`），不是拦住输入：ProseMirror 的输入、
 * 粘贴、删除都以事务的形式经过同一条通道，从上游拦截要么错过 IME 的组合输入，要么错过
 * 快捷键删除；社区实现（discuss.prosemirror.net 的 track-changes 那几帖）也都是这个位置，
 * 并且同样提醒过两个坑——**必须跳过撤销/重做产生的事务**（否则撤一次改动会被记成一次新修订），
 * 以及**格式类改动记不了**（`w:rPrChange` / `w:pPrChange` 这一层本仓没建模，见规格排名 4 判据）。
 *
 * 记下来的形状就是 ③④a 那套 `revision` 标记：插入 = 给新打的字补标记，删除 = 把删掉的
 * 文字插回原位并补上 `del` 标记（Word 写 `w:delText` 的那一层在 docx-io 的写侧）。
 */

/** 打开开关时带进来的署名；两者都可选——`CT_TrackChange` 的 `w:author` / `w:date` 不是必填。 */
export interface TrackSignature {
  /** 设置里的作者名（`w:author`）；留空就一个都不编。 */
  author?: string | null;
  /** 到秒的 ISO 时间戳（`w:date`），按下开关那一刻定下来，本会话的新修订共用它。 */
  date?: string | null;
}

/** 插件状态：关着的时候 `appendTransaction` 第一眼就返回。 */
export interface TrackState extends TrackSignature {
  enabled: boolean;
}

/**
 * 我们自己补的那一笔的记号。
 *
 * `appendTransaction` 会再次看到它——不认出来的话，刚插回去的删除文字会被当成一次新插入，
 * 同一段字上同时挂 `ins` 与 `del`。
 */
const OWN_META = "trackChangesApplied";

/** 开关的状态要从外面读与写（工具栏的命令、以及没有 DOM 环境时的单测），所以键是公开的。 */
export const trackKey = new PluginKey<TrackState>("trackChanges");

declare module "@tiptap/core" {
  interface Commands<ReturnType> {
    trackChanges: {
      /** 打开修订开关，并定下这一轮改动的署名。 */
      enableTrackChanges: (signature?: TrackSignature) => ReturnType;
      /** 关掉开关（已经记下的修订留在文档里，交给接受/拒绝）。 */
      disableTrackChanges: () => ReturnType;
    };
  }
}

/** 待补的一次修订：给已经落进文档的新字补标记，或把删掉的文字插回原位。 */
type RevisionJob =
  | { kind: "mark"; at: number; from: number; to: number }
  | { kind: "restore"; at: number; nodes: PMNode[] };

export const TrackChanges = Extension.create({
  name: "trackChanges",

  addProseMirrorPlugins() {
    return [trackChangesPlugin()];
  },

  addCommands() {
    return {
      enableTrackChanges:
        (signature: TrackSignature = {}) =>
        ({ tr, dispatch }: CommandProps) => {
          if (!dispatch) return true;
          tr.setMeta(trackKey, {
            enabled: true,
            author: signature.author?.trim() || null,
            date: signature.date ?? null,
          } satisfies TrackState);
          return true;
        },
      disableTrackChanges:
        () =>
        ({ tr, dispatch }: CommandProps) => {
          if (!dispatch) return true;
          tr.setMeta(trackKey, { enabled: false, author: null, date: null } satisfies TrackState);
          return true;
        },
    };
  },
});

/**
 * 修订开关的 ProseMirror 插件。
 *
 * 单独导出这一层：本仓没有 DOM 测试环境，单测要拿真实 schema 建 `EditorState` 并把插件挂上去，
 * `TrackChanges` 只负责把它交给编辑器。
 *
 * @returns 插件实例；初始状态是关着的。
 */
export function trackChangesPlugin(): Plugin<TrackState> {
  return new Plugin<TrackState>({
    key: trackKey,
    state: {
      init: () => ({ enabled: false, author: null, date: null }),
      apply: (tr, value) => tr.getMeta(trackKey) ?? value,
    },
    // 起点状态用不上：坐标换算要的是"每一步之前那份文档"，那个在事务自己身上（`tr.docs`）。
    appendTransaction: (trs, _oldState, newState) => trackTransaction(trs, newState),
  });
}

/**
 * 把这一批事务里的文字改动改写成一笔修订。
 *
 * @param trs 这一批刚落地的（非空）事务。
 * @param newState 落地以后的状态。
 * @returns 要补的那一笔；这一批里没有可记的改动（或整个开关没开）时返回 null。
 */
function trackTransaction(
  trs: readonly Transaction[],
  newState: EditorState,
): Transaction | null {
  const signature = trackKey.getState(newState);
  if (!signature?.enabled) return null;
  if (trs.some((tr) => tr.getMeta(OWN_META))) return null;
  // 撤销与重做是把过去的状态放回来，不是"本次编辑"：记成修订等于撤销一次改动又生出两条新改动。
  if (trs.some((tr) => isHistoryTransaction(tr))) return null;
  const revisionType = newState.schema.marks.revision;
  if (!revisionType) return null;

  const jobs = revisionJobs(trs, signature);
  if (!jobs.length) return null;

  const tr = newState.tr.setMeta(OWN_META, true);
  let caret = newState.selection.empty ? newState.selection.from : null;
  // 从后往前落地：插回删除的文字会让它后面所有位置右移，倒着走才不用重算坐标。
  // 同一处既有插入又有删除（选中一段字直接打字）时先给新字补标记、再把旧字插到它前面，
  // 出来就是 Word 的 `<w:del>旧</w:del><w:ins>新</w:ins>` 顺序。
  for (const job of jobs.sort(byPositionDesc)) {
    if (job.kind === "mark") {
      const mark = revisionType.create(revisionAttrs("ins", tr.doc, job.from, signature));
      for (const range of textRanges(tr.doc, job.from, job.to)) {
        tr.addMark(range.from, range.to, mark);
      }
      continue;
    }
    const mark = revisionType.create(revisionAttrs("del", tr.doc, job.at, signature));
    const nodes = job.nodes.map((node) => node.mark(withoutRevision(node.marks).concat(mark)));
    tr.insert(job.at, Fragment.from(nodes));
    // 在光标处插回了删除的字：光标留在删除线的右边（Word 也是这个位置），否则下一次删除
    // 会跑到自己刚插回去的那段字前面。
    if (caret !== null && job.at <= caret) caret += nodes.reduce((size, node) => size + node.nodeSize, 0);
  }
  if (caret !== null && caret !== newState.selection.from) {
    tr.setSelection(TextSelection.near(tr.doc.resolve(caret)));
  }
  return tr;
}

/**
 * 这一批事务要补成哪些修订，坐标已经换算到最终文档。
 *
 * @param trs 刚落地的批次。
 * @param signature 当前署名，用来认"这次刚插入的字又被删掉"那种互相抵消的情况。
 * @returns 待补的修订（未排序）。
 */
function revisionJobs(trs: readonly Transaction[], signature: TrackState): RevisionJob[] {
  // 逐步骤走一遍：`tr.docs[i]` 是该步骤之前的文档（删除的文字从那里取），
  // 而后面的步骤会把坐标推走，所以每一步都记住自己在全部映射里的位置。
  const maps: StepMap[] = [];
  const edits: { step: ReplaceStep; before: PMNode; index: number }[] = [];
  for (const tr of trs) {
    tr.steps.forEach((step, i) => {
      // `Transform` 给每一步都存了它之前那份文档，索引必然对得上步骤（这里只是 TS 的下标取值不给这个保证）。
      const before = tr.docs[i] as PMNode;
      // 只认替换步骤：加/摘标记（改粗体、改字号）是格式修订，本仓没建模；
      // 拆分与抬升段落（回车、变列表）改的是段落结构，段落标记的修订同样没建模。
      if (step instanceof ReplaceStep) edits.push({ step, before, index: maps.length });
      maps.push(step.getMap());
    });
  }

  const jobs: RevisionJob[] = [];
  for (const { step, before, index } of edits) {
    const forward = new Mapping(maps.slice(index + 1));
    if (step.slice.size > 0) {
      jobs.push({
        kind: "mark",
        at: forward.map(step.from, -1),
        from: forward.map(step.from, -1),
        to: forward.map(step.from + step.slice.size, -1),
      });
    }
    const removed = removedText(before, step, signature);
    if (removed) jobs.push({ kind: "restore", at: forward.map(step.from, -1), nodes: removed });
  }
  return jobs;
}

/**
 * 这一步删掉的、能还原成删除修订的文字。
 *
 * @param before 这一步之前的文档。
 * @param step 那个替换步骤。
 * @param signature 当前署名。
 * @returns 要插回去的文字节点（带着原来的字符样式）；这一步不该管时返回 null。
 */
function removedText(before: PMNode, step: ReplaceStep, signature: TrackState): PMNode[] | null {
  if (step.to <= step.from) return null;
  const slice = before.slice(step.from, step.to);
  // 跨块的删除（在段首按 Backspace 把两段并起来）要还原的是段落结构而不是几个字，
  // 段落的增删只写在 `w:pPr/w:rPr` 里，本层不建模（规格排名 4 判据）。
  if (slice.openStart > 0 || slice.openEnd > 0) return null;
  const nodes: PMNode[] = [];
  for (let i = 0; i < slice.content.childCount; i++) {
    // 图片、引用点这类原子节点上没有地方放标记，`w:del` 里的图片这一层记不了。
    if (!slice.content.child(i).isText) return null;
    nodes.push(slice.content.child(i));
  }
  if (!nodes.length) return null;
  // 这次刚插入的字又被删掉：两次改动互相抵消，插回去只会留下一条没人接受的删除线。
  // （Word 同规则；别人插入的字被我们删掉时按"外层是本次操作"处理，见下面换成 del。）
  const justInserted = nodes.every((node) =>
    node.marks.some((mark) => mark.type.name === "revision" && isSameRevision(mark, "ins", signature)),
  );
  return justInserted ? null : nodes;
}

/**
 * 这一次修订的标记属性：编号 + 按开关那一刻定下来的署名。
 *
 * @param kind `ins` 还是 `del`。
 * @param doc 当前文档（拿前一个字判断能不能并进同一次修订，也拿它算下一个编号）。
 * @param at 这次改动在文档里的起点。
 * @param signature 当前署名。
 * @returns `revision` 标记的属性（`id` 保持字符串，§2.1）。
 */
function revisionAttrs(
  kind: "ins" | "del",
  doc: PMNode,
  at: number,
  signature: TrackState,
): Record<string, unknown> {
  return {
    kind,
    id: revisionIdAt(kind, doc, at, signature),
    author: signature.author,
    date: signature.date,
  };
}

/**
 * 这一段改动归到哪个 `w:id`。
 *
 * 接着前一个字、同一个作者同一批的同类修订就并进去（Word 也是连续输入合成一条 `w:ins`，
 * 一条条各给 id 会让一次打字碎成几十条修订），否则在本文档的修订编号里取最大值加一。
 *
 * @param kind `ins` 还是 `del`。
 * @param doc 当前文档。
 * @param at 这次改动的起点。
 * @param signature 当前署名。
 * @returns 整数字符串形式的 `w:id`。
 */
function revisionIdAt(kind: "ins" | "del", doc: PMNode, at: number, signature: TrackState): string {
  const caret = doc.resolve(Math.min(Math.max(at, 0), doc.content.size));
  const ahead = caret.nodeBefore?.marks.find((mark) => mark.type.name === "revision");
  // 接着上一次同类改动：并进那一条修订，Word 里就是一次 `w:ins` 包若干个 run。
  if (ahead && isSameRevision(ahead, kind, signature)) return String(ahead.attrs.id);
  return nextAnnotationId(revisionIds(doc), 1);
}

/**
 * 这条修订标记是不是"同一个人、同一批、同一种操作"的再一次改动。
 *
 * 作者与时间戳都是按下开关那一刻定下来的，所以同一个开关会话里的改动天然成一批。
 *
 * @param mark 要判断的 `revision` 标记。
 * @param kind 本次改动的类型。
 * @param signature 当前署名。
 * @returns 能并进同一条修订为 true；`w:id` 缺失的标记不并（宁可多开一条，也不要并到一条没有编号的）。
 */
function isSameRevision(mark: Mark, kind: "ins" | "del", signature: TrackState): boolean {
  return (
    mark.attrs.kind === kind &&
    mark.attrs.author === signature.author &&
    mark.attrs.date === signature.date &&
    typeof mark.attrs.id === "string" &&
    mark.attrs.id.length > 0
  );
}

/**
 * 文档里已经在用的修订编号。
 *
 * @param doc 当前文档。
 * @returns 所有 `revision` 标记上的 `w:id`（字符串原样）。
 */
function revisionIds(doc: PMNode): string[] {
  const ids: string[] = [];
  doc.descendants((node) => {
    for (const mark of node.marks) {
      if (mark.type.name === "revision" && mark.attrs.id) ids.push(String(mark.attrs.id));
    }
    return true;
  });
  return ids;
}

/**
 * 一段范围里逐个文字节点的位置（原子节点与锚点不参与，它们没有地方放标记）。
 *
 * @param doc 当前文档。
 * @param from 起点。
 * @param to 终点。
 * @returns 每个文字节点被这次改动覆盖到的区间。
 */
function textRanges(doc: PMNode, from: number, to: number): { from: number; to: number }[] {
  const ranges: { from: number; to: number }[] = [];
  doc.nodesBetween(from, to, (node, pos) => {
    if (node.isText) ranges.push({ from: Math.max(pos, from), to: Math.min(pos + node.nodeSize, to) });
    return true;
  });
  return ranges;
}

/**
 * 摘掉原有的修订标记。
 *
 * 已经标成插入的字又被删掉时（署名对不上，上面那条抵消规则没走），外层是本次删除——
 * 与读侧"嵌套取外层"同一条规则。
 *
 * @param marks 节点原有的标记。
 * @returns 不含 `revision` 的那一份。
 */
function withoutRevision(marks: readonly Mark[]): Mark[] {
  return marks.filter((mark) => mark.type.name !== "revision");
}

const byPositionDesc = (a: RevisionJob, b: RevisionJob): number =>
  b.at - a.at || (a.kind === b.kind ? 0 : a.kind === "mark" ? -1 : 1);
