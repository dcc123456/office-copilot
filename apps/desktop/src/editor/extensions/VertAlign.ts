import { Mark, mergeAttributes, type CommandProps } from "@tiptap/core";
import type { VerticalAlign } from "@office-copilot/docx-io";

/**
 * 上标 / 下标。文档侧只有一个 `<w:vertAlign w:val>` 元素（ECMA-376 §17.3.2.27 的
 * ST_VerticalAlignRun：baseline / superscript / subscript），所以编辑器里也是**一个标记带
 * 一个枚举值**，而不是两个开关标记——两个标记能同时挂成"又上标又下标"，OOXML 无处表达。
 *
 * `baseline` 在文档里是"显式取消样式继承来的上标"，在编辑器里就是没有这个标记，
 * 因此这里只收另外两档。
 */
export type VertAlignValue = Exclude<VerticalAlign, "baseline">;

declare module "@tiptap/core" {
  interface Commands<ReturnType> {
    vertAlign: {
      setVertAlign: (attrs: { val: VertAlignValue }) => ReturnType;
      toggleVertAlign: (attrs: { val: VertAlignValue }) => ReturnType;
      unsetVertAlign: () => ReturnType;
    };
  }
}

export const VertAlign = Mark.create({
  name: "vertAlign",

  // 与加粗同族：Word 里敲字会一直带着上标，直到用户再点一次取消。
  inclusive: true,

  addOptions() {
    return { HTMLAttributes: {} };
  },

  addAttributes() {
    return {
      /**
       * `w:vertAlign` 的 `w:val`，缺省上标（不指明档位的调用方拿到的就是它）。
       *
       * `rendered: false`：档位已经由 `<sup>`/`<sub>` 标签本身表达，再挂一个 `val="…"`
       * 到 DOM 上是复制/打印出去的非标准属性；落盘走 `editor.getJSON()`，attrs 照样带得出。
       */
      val: { default: "superscript" as VertAlignValue, rendered: false },
    };
  },

  // 档位由标签决定，所以贴回来的 HTML 不用再带 data-*。
  parseHTML() {
    return [
      { tag: "sup", getAttrs: () => ({ val: "superscript" as VertAlignValue }) },
      { tag: "sub", getAttrs: () => ({ val: "subscript" as VertAlignValue }) },
    ];
  },

  renderHTML({ mark, HTMLAttributes }) {
    const val = (mark as { attrs?: { val?: string } } | undefined)?.attrs?.val;
    return [
      val === "subscript" ? "sub" : "sup",
      mergeAttributes(this.options.HTMLAttributes as Record<string, unknown>, HTMLAttributes),
      0,
    ];
  },

  addCommands() {
    return {
      setVertAlign:
        (attrs: { val: VertAlignValue }) =>
        ({ chain }: CommandProps) =>
          chain().setMark(this.name, { val: attrs.val }).run(),
      // 同一档再点一次是取消；另一档再点是换过去（Word 的两个按钮互斥）。
      toggleVertAlign:
        (attrs: { val: VertAlignValue }) =>
        ({ editor, chain }: CommandProps) =>
          editor.isActive(this.name, { val: attrs.val })
            ? chain().unsetMark(this.name).run()
            : chain().setMark(this.name, { val: attrs.val }).run(),
      unsetVertAlign:
        () =>
        ({ chain }: CommandProps) =>
          chain().unsetMark(this.name).run(),
    };
  },
});
