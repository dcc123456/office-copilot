import { Mark, mergeAttributes, type CommandProps, type Editor } from "@tiptap/core";
import type { MarkType, Node as PMNode } from "@tiptap/pm/model";
import type { PMMark } from "@office-copilot/docx-io";

/**
 * 修订跟踪（Word 的"修订"）在编辑器里的形状。
 *
 * 文档侧是 `<w:ins>` / `<w:del>` 包住运行（`CT_RunTrackChange`），删除的文字写成 `w:delText`；
 * 这里做成**字符标记**而不是节点，因为一次修订只覆盖段落里的一部分文字，而且被改的字必须
 * 留在原地（原子节点会把文字搬走，选区、撤销、拼写全乱）。属性与 docx-io 那一侧同一个类型，
 * 少一次口径漂移（§2.1）。
 *
 * `inclusive: false`：在一段插入的末尾接着打字，新字属于**下一次**改动，不该继承这条修订的
 * 作者与编号——Word 开着修订时也是另起一条。
 */

/** 修订标记的属性形状，与 `PMMark` 里那条 `revision` 是同一个类型。 */
type RevisionAttrs = Extract<PMMark, { type: "revision" }>["attrs"];

declare module "@tiptap/core" {
  interface Commands<ReturnType> {
    revision: {
      /** 给选区打上修订标记（⑥ 的"开了修订就记成 ins/del"用这条）。 */
      setRevision: (attrs: RevisionAttrs) => ReturnType;
      /** 接受光标/选区碰到的每一次修订。 */
      acceptRevision: () => ReturnType;
      /** 拒绝光标/选区碰到的每一次修订。 */
      rejectRevision: () => ReturnType;
    };
  }
}

/**
 * 一次修订在文档里占的区间。
 *
 * @param doc 当前文档。
 * @param type 修订标记的类型（从 schema 取，避免按名字符串建标记）。
 * @param from 选区起点。
 * @param to 选区终点。
 * @returns 选区碰到的每一条修订及其**全篇**区间（同一次修订在 Word 里常被按段落拆成好几条
 *   `w:ins`，所以按属性归组后要扫全篇）；没碰到修订时为空数组。
 */
function revisionsTouching(
  doc: PMNode,
  type: MarkType,
  from: number,
  to: number,
): { kind: "ins" | "del"; ranges: { from: number; to: number }[] }[] {
  const groups = new Map<string, { kind: "ins" | "del"; ranges: { from: number; to: number }[] }>();
  doc.descendants((node, pos) => {
    if (!node.isText) return true;
    const mark = node.marks.find((m) => m.type === type);
    if (!mark) return true;
    const key = JSON.stringify(mark.attrs);
    const kind = mark.attrs.kind === "del" ? "del" : "ins";
    const group = groups.get(key) ?? { kind, ranges: [] };
    const last = group.ranges[group.ranges.length - 1];
    // 同一次修订里相邻的文字节点并成一段：一次插入通常是若干个节点（各带自己的字符样式）。
    if (last && last.to === pos) last.to = pos + node.nodeSize;
    else group.ranges.push({ from: pos, to: pos + node.nodeSize });
    groups.set(key, group);
    return true;
  });
  // 只留选区碰得到的：光标落在一次修订里就够了，与 Word 的"接受此修订"同语义。
  return [...groups.values()].filter((g) => g.ranges.some((r) => r.from < to && r.to > from));
}

/**
 * 选区（或光标所在处）有没有落在某一次修订里——右键菜单据此决定"接受/拒绝修订"是否可用。
 *
 * @param editor 桌面编辑器实例。
 * @returns 碰到了修订为 true；文档里没有修订、或选区落在修订之外为 false。
 */
export function selectionTouchesRevision(editor: Editor): boolean {
  const type = editor.state.schema.marks.revision;
  if (!type) return false;
  const { from, to } = editor.state.selection;
  return revisionsTouching(editor.state.doc, type, from, to).length > 0;
}

export const Revision = Mark.create({
  name: "revision",

  inclusive: false,

  addOptions() {
    return { HTMLAttributes: {} };
  },

  addAttributes() {
    return {
      /** `w:ins` 还是 `w:del`。 */
      kind: { default: "ins" as "ins" | "del", rendered: false },
      /** `w:id`：ST_DecimalNumber 原样存字符串（文档级编号空间，不做数字转换）。 */
      id: { default: null as string | null, rendered: false },
      /** `w:author`：缺省 null = 文档里没署名，我们不替用户编一个。 */
      author: { default: null as string | null, rendered: false },
      /** `w:date`：ISO-8601 字符串，原样带走。 */
      date: { default: null as string | null, rendered: false },
    };
  },

  // 不从 HTML 认修订：粘进来的内容属于本次编辑，不该带着别人的作者与编号变成一条修订。
  parseHTML() {
    return [];
  },

  renderHTML({ mark, HTMLAttributes }) {
    const attrs = mark.attrs as RevisionAttrs;
    const isDelete = attrs.kind === "del";
    // 只给划线/删除线与颜色，正文颜色留给文档自己的字符样式（Word 也是按作者上色，
    // 但把作者名塞进 color 会把用户的红字改成黑字）。
    return [
      "span",
      mergeAttributes(this.options.HTMLAttributes as Record<string, unknown>, HTMLAttributes, {
        "data-revision": attrs.kind,
        title: `${isDelete ? "删除" : "插入"}${attrs.author ? `：${attrs.author}` : ""}${
          attrs.date ? `（${attrs.date}）` : ""
        }`,
        style: isDelete
          ? "text-decoration: line-through; text-decoration-color: #cf222e;"
          : "text-decoration: underline; text-decoration-color: #2ea043;",
      }),
      0,
    ];
  },

  addCommands() {
    return {
      setRevision:
        (attrs: RevisionAttrs) =>
        ({ state, dispatch }: CommandProps) => {
          const type = state.schema.marks[this.name];
          if (!type || state.selection.empty) return false;
          if (dispatch) dispatch(state.tr.addMark(state.selection.from, state.selection.to, type.create(attrs)));
          return true;
        },
      acceptRevision: () => resolveRevisions(this.name, "accept"),
      rejectRevision: () => resolveRevisions(this.name, "reject"),
    };
  },
});

/**
 * 接受 / 拒绝选区碰到的修订，只发一个事务。
 *
 * 接受 = 改动生效（插入留着、删除的字消失）；拒绝 = 改动作废（插入的字消失、删除的字留下）。
 * 两种都落在"删字"或"摘标记"这两件事上：删哪个、留哪个只由 `kind` 决定。
 *
 * @param name 标记名（命令上下文里的 `this.name`）。
 * @param decision 接受还是拒绝。
 * @returns 选区里有没有修订；没有时 false，调用方（菜单项）据此置灰。
 */
function resolveRevisions(name: string, decision: "accept" | "reject") {
  return ({ state, dispatch }: CommandProps): boolean => {
    const type = state.schema.marks[name];
    if (!type) return false;
    const { from, to } = state.selection;
    const found = revisionsTouching(state.doc, type, from, to);
    if (!found.length) return false;
    let tr = state.tr;
    // 从后往前落地：删掉一段文字会让它后面的位置整体左移，倒着走才不用重新映射坐标。
    const jobs = found
      .flatMap((group) => group.ranges.map((range) => ({ ...range, kind: group.kind })))
      .sort((a, b) => b.from - a.from);
    for (const job of jobs) {
      const drop = decision === "accept" ? job.kind === "del" : job.kind === "ins";
      tr = drop ? tr.delete(job.from, job.to) : tr.removeMark(job.from, job.to, type);
    }
    if (dispatch) dispatch(tr);
    return true;
  };
}
