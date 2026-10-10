import { Node } from "@tiptap/core";

/**
 * 书签的两个锚点节点，对应 OOXML 的 `<w:bookmarkStart w:id w:name/>` 与
 * `<w:bookmarkEnd w:id/>`（ECMA-376 §17.5.2.1）。
 *
 * 两者都是行内 atom、不带文字：`docx-io` 按文档顺序把它们成对收进段落内容，
 * 写出时又放回 `<w:p>` 这一层。起点与终点可以分处不同段落（Word 的书签就是区间），
 * 所以这里不做配对校验——半截书签由写侧的 `dropUnpairedBookmarks` 兜住。
 *
 * @param name 节点名，与 `docx-io` 的 PM 节点名一致。
 * @param isStart 起点带名字（CSS 用 `attr()` 显示成小标签），终点只画一个反向竖标。
 * @returns 注册进 `EDITOR_EXTENSIONS` 的 Tiptap 节点。
 */
function bookmarkNode(name: "bookmarkStart" | "bookmarkEnd", isStart: boolean) {
  return Node.create({
    name,
    group: "inline",
    inline: true,
    atom: true,

    addAttributes() {
      return {
        /** OOXML 的 `w:id`：配对用的整数，原样存字符串（换算只在 docx-io 进出口）。 */
        id: {
          default: "0",
          parseHTML: (el) => el.getAttribute("data-bm-id") ?? "0",
        },
        /**
         * OOXML 的 `w:name`：只有起点带，交叉引用与目录按名字定位。
         * 终点写 `w:name` 是无效属性，所以缺省是 `undefined` 而不是空串。
         */
        name: {
          default: undefined,
          parseHTML: (el) => el.getAttribute("data-bm-name") ?? undefined,
        },
      };
    },

    parseHTML() {
      return [{ tag: `span[data-bookmark="${isStart ? "start" : "end"}"]` }];
    },

    // 只写 data-* 属性：把 HTMLAttributes 原样并进来会连 `id` 一起写到 DOM 上，
    // 那正好是 ProseMirror 里最不该碰的保留属性（会造出重复的 DOM id）。
    renderHTML({ node }) {
      const id = String(node.attrs.id ?? "");
      const label = node.attrs.name ? String(node.attrs.name) : "";
      return [
        "span",
        {
          "data-bookmark": isStart ? "start" : "end",
          "data-bm-id": id,
          ...(label ? { "data-bm-name": label, title: label } : {}),
          class: `pm-bookmark pm-bookmark-${isStart ? "start" : "end"}`,
        },
      ];
    },
  });
}

export const BookmarkStart = bookmarkNode("bookmarkStart", true);

export const BookmarkEnd = bookmarkNode("bookmarkEnd", false);
