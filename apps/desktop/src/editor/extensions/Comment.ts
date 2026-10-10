import { Node } from "@tiptap/core";

/**
 * 批注在正文里的三个锚点节点，对应 OOXML 的 `<w:commentRangeStart w:id/>`、
 * `<w:commentRangeEnd w:id/>` 与 `<w:commentReference w:id/>`。
 *
 * 内容与脚注同构：正文只装 id，批注的**文字**住在 `word/comments.xml` 部件里
 * （`docx-io` 的 `CommentStore`）。编辑器不认这三个节点，打开带批注的文档一保存
 * 就把锚点整批丢掉，`docx-io` 随后按锚点对齐还会把部件里的条目一起删掉。
 *
 * 区间端点与引用记号的版面不同，是因为它们在内容模型里位置不同：前两者是 `w:p`
 * 的直接子元素（与 `w:r` 平级），后者必须装在 `w:r` 里面。跨段区间的高亮与编号
 * 属于批注侧栏（spec 排名 4 ⑤），这里只保证锚点活着回来。
 */

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
