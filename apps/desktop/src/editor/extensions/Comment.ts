import { Extension, Node, type CommandProps } from "@tiptap/core";
import type { Node as PMNode } from "@tiptap/pm/model";

/**
 * 批注在正文里的三个锚点节点，对应 OOXML 的 `<w:commentRangeStart w:id/>`、
 * `<w:commentRangeEnd w:id/>` 与 `<w:commentReference w:id/>`。
 *
 * 内容与脚注同构：正文只装 id，批注的**文字**住在 `word/comments.xml` 部件里
 * （`docx-io` 的 `CommentStore`）。编辑器不认这三个节点，打开带批注的文档一保存
 * 就把锚点整批丢掉，`docx-io` 随后按锚点对齐还会把部件里的条目一起删掉。
 *
 * 区间端点与引用记号的版面不同，是因为它们在内容模型里位置不同：前两者是 `w:p`
 * 的直接子元素（与 `w:r` 平级），后者必须装在 `w:r` 里面。
 *
 * `CommentCommands` 与 `commentAnchorRanges` 是 ⑤ 的另一半：侧栏的"新增（选区）/删除/
 * 跳转"全靠这两个命令加一份位置索引，正文本身仍然只装 id，批注文字住部件里。
 */

/**
 * 一次批注在正文里占的三个锚点位。
 *
 * 三种锚点各自独立，规范不强制哪一条必须齐全（别的生成器只写区间两端、不写引用记号），
 * 所以缺哪个就是 `undefined`，调用方自己决定怎么兜底。
 */
export interface CommentAnchorPos {
  /** `commentRangeStart` 的位置，PM 位置（不是像素）。 */
  start?: number;
  /** `commentRangeEnd` 的位置。 */
  end?: number;
  /** `commentReference` 的位置。 */
  reference?: number;
}

/** 一条批注的锚点位置，按正文出现顺序排。 */
export interface CommentAnchorRange extends CommentAnchorPos {
  /** OOXML 的 `w:id` 字符串。 */
  id: string;
}

/** 节点名 → 位置字段；不是锚点就是 `undefined`。 */
const EDGE_OF: Record<string, keyof CommentAnchorPos | undefined> = {
  commentRangeStart: "start",
  commentRangeEnd: "end",
  commentReference: "reference",
};

declare module "@tiptap/core" {
  interface Commands<ReturnType> {
    comment: {
      /** 给当前选区套上一对区间锚点并在末尾补引用记号，`id` 是调用方分好的 `w:id`。 */
      insertCommentRange: (id: string) => ReturnType;
      /** 删掉某个批注在正文里的三个锚点；部件里的条目交给保存时的 `syncCommentIds`。 */
      deleteComment: (id: string) => ReturnType;
    };
  }
}

/**
 * 按正文顺序列出每篇批注的锚点位置。
 *
 * 排序用的是"这段区间第一次出现的位置"，与 Word 侧栏的顺序一致（Word 的批注列表也按
 * `commentReference` 在正文里的先后排）。id 是空串的锚点（读不认识的写法）不进表。
 *
 * @param doc 编辑器当前正文（ProseMirror 节点，不是 JSON）。
 * @returns 按文档顺序排的锚点位置表；正文没有批注时是空数组。
 */
export function commentAnchorRanges(doc: PMNode): CommentAnchorRange[] {
  const order: string[] = [];
  const byId = new Map<string, CommentAnchorRange>();
  doc.descendants((node, pos) => {
    const edge = EDGE_OF[node.type.name];
    if (!edge) return;
    const id = String(node.attrs.id ?? "");
    if (!id) return;
    let range = byId.get(id);
    if (!range) {
      range = { id };
      byId.set(id, range);
      order.push(id);
    }
    range[edge] = pos;
  });
  return order.map((id) => byId.get(id)!);
}

/**
 * 把位置挪进允许行内内容的地方。
 *
 * 三个锚点都是行内 atom，而全选（Ctrl+A）的选区起点落在文档开头那个块边界上——那里只
 * 允许块内容，直接插会造出一棵不合法的树。往段内挪一步就够了，正文里的光标位置本来就
 * 都在段落内。
 *
 * @param doc 当前正文，用来判断位置的父节点。
 * @param pos PM 位置。
 * @returns 可以放行内 atom 的位置。
 */
function inlineAnchorPos(doc: PMNode, pos: number): number {
  if (doc.resolve(pos).parent.inlineContent) return pos;
  // 块边界往前一步就回到上一段的末尾（全选时选区终点正是文档末尾这种），否则往后一步进段内。
  const back = Math.max(0, pos - 1);
  return doc.resolve(back).parent.inlineContent ? back : pos + 1;
}

/**
 * 批注的插入与删除命令。
 *
 * 只挂命令、不碰版面，所以是 `Extension` 而不是 `Node`；锚点节点本身在 `CommentAnchors` 里。
 */
export const CommentCommands = Extension.create({
  name: "commentCommands",

  addCommands() {
    return {
      insertCommentRange:
        (id) =>
        ({ state, dispatch }) => {
          const { from, to, empty } = state.selection;
          if (empty || !id) return false;
          const { commentRangeStart, commentRangeEnd, commentReference } = state.schema.nodes;
          // 嵌套编辑器（脚注正文那类）用的 schema 里没有这三个节点，此时不动手而不是抛异常。
          if (!commentRangeStart || !commentRangeEnd || !commentReference) return false;
          // 一次插入会把它之后的位置整体后移，所以先动大位置：引用记号、区间止、区间起。
          const startAt = inlineAnchorPos(state.doc, from);
          const endAt = inlineAnchorPos(state.doc, to);
          const tr = state.tr
            .insert(endAt, commentReference.create({ id }))
            .insert(endAt, commentRangeEnd.create({ id }))
            .insert(startAt, commentRangeStart.create({ id }));
          if (!dispatch) return true;
          dispatch(tr.scrollIntoView());
          return true;
        },

      deleteComment:
        (id) =>
        ({ state, dispatch }) => {
          const positions: number[] = [];
          state.doc.descendants((node, pos) => {
            if (EDGE_OF[node.type.name] && String(node.attrs.id ?? "") === id) positions.push(pos);
          });
          if (positions.length === 0) return false;
          if (!dispatch) return true;
          const tr = state.tr;
          // 倒序删：删掉一个锚点，它之后的位置都要往前挪。
          for (const pos of positions.reverse()) tr.delete(pos, pos + 1);
          dispatch(tr.scrollIntoView());
          return true;
        },
    };
  },
});

/**
 * 造一个批注锚点节点。
 *
 * @param name PM 节点名，必须与 `docx-io` 的 `PMNode` 类型名一致。
 * @param edge 三种锚点之一：`start`/`end` 画区间端点，`reference` 画引用记号芯片。
 * @returns 注册进 `EDITOR_EXTENSIONS` 的 Tiptap 节点。
 */
function commentAnchorNode(
  name: "commentRangeStart" | "commentRangeEnd" | "commentReference",
  edge: "start" | "end" | "reference",
) {
  return Node.create({
    name,
    group: "inline",
    inline: true,
    atom: true,

    addAttributes() {
      return {
        /** OOXML 的 `w:id`：与书签共用文档级编号空间，原样存字符串（换算只在 docx-io 进出口）。 */
        id: {
          default: "",
          parseHTML: (el) => el.getAttribute("data-comment-id") ?? "",
        },
      };
    },

    parseHTML() {
      return [{ tag: `span[data-comment="${edge}"]` }];
    },

    // 只写 data-* 属性：把 HTMLAttributes 原样并进来会连 `id` 一起写到 DOM 上，
    // 那是 ProseMirror 里最不该碰的保留属性（会造出重复的 DOM id）。
    // 版面用行内 style 而不是 styles.css，与 `FootnoteRef.ts` 的芯片同一个做法。
    renderHTML({ node }) {
      const id = String(node.attrs.id ?? "");
      return [
        "span",
        {
          "data-comment": edge,
          "data-comment-id": id,
          title: edge === "reference" ? `批注 ${id}` : `批注区间${edge === "start" ? "起" : "止"} ${id}`,
          class: `pm-comment pm-comment-${edge}`,
          style:
            edge === "reference"
              ? "vertical-align:super;font-size:0.75em;padding:0 2px;color:#8250df;"
              : "color:#8250df;font-size:0.7em;",
        },
        edge === "reference" ? "批" : edge === "start" ? "〔" : "〕",
      ];
    },
  });
}

export const CommentRangeStart = commentAnchorNode("commentRangeStart", "start");

export const CommentRangeEnd = commentAnchorNode("commentRangeEnd", "end");

export const CommentReference = commentAnchorNode("commentReference", "reference");

/** 编辑器要注册的完整一组：三个锚点。 */
export const CommentAnchors = [CommentRangeStart, CommentRangeEnd, CommentReference];
