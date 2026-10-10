import { Node } from "@tiptap/core";

/**
 * 脚注/尾注的引用点节点，对应 OOXML 的 `<w:footnoteReference w:id/>` 与
 * `<w:endnoteReference w:id/>`（ECMA-376 §17.11.18）。
 *
 * 行内 atom、不带内容：正文里只有这一个记号，脚注的**文字**住在 `word/footnotes.xml`
 * 部件里，靠同一个 id 对应（`docx-io` 的 `FootnoteStore`）。编辑器不认这个节点就得
 * 丢掉引用点，一保存整篇脚注都没了，所以这一层先只做"引用点在原处不动"。
 * 芯片上显示的是"注/尾"而不是 Word 的编号——编号要按正文里的出现顺序现算，
 * `renderHTML` 拿不到位置，那一步和脚注正文的编辑入口一起做。
 *
 * @param kind 脚注还是尾注，决定节点名（与 `docx-io` 的 PM 节点名一致）与内容部件。
 * @returns 注册进 `EDITOR_EXTENSIONS` 的 Tiptap 节点。
 */
function footnoteRefNode(kind: "footnote" | "endnote") {
  return Node.create({
    name: kind === "footnote" ? "footnoteReference" : "endnoteReference",
    group: "inline",
    inline: true,
    atom: true,

    addAttributes() {
      return {
        /** OOXML 的 `w:id`：ST_DecimalNumber 原样存字符串，可以是负数。 */
        id: {
          default: "",
          parseHTML: (el) => el.getAttribute("data-footnote-id") ?? "",
        },
      };
    },

    parseHTML() {
      return [{ tag: `span[data-footnote-ref="${kind}"]` }];
    },

    // 只写 data-* 属性：把 HTMLAttributes 原样并进来会连 `id` 一起写到 DOM 上，
    // 那是 ProseMirror 里最不该碰的保留属性（会造出重复的 DOM id）。
    renderHTML({ node }) {
      const id = String(node.attrs.id ?? "");
      return [
        "span",
        {
          "data-footnote-ref": kind,
          "data-footnote-id": id,
          title: `${kind === "footnote" ? "脚注" : "尾注"} ${id}`,
          class: `pm-footnote-ref pm-footnote-${kind}`,
          style: "vertical-align:super;font-size:0.75em;padding:0 2px;color:#0078d4;",
        },
        kind === "footnote" ? "注" : "尾",
      ];
    },
  });
}

export const FootnoteReference = footnoteRefNode("footnote");

export const EndnoteReference = footnoteRefNode("endnote");
