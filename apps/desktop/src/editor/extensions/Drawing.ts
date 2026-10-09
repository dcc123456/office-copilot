import { Node, mergeAttributes } from "@tiptap/core";

/**
 * Inline image node. We use a single node for both inline drawings and
 * floating (wp:anchor) drawings. For anchors we render a zero-width
 * inline "anchor" element + an absolutely-positioned overlay so it does
 * not affect line layout, matching Word's <wp:wrapNone/> behavior.
 *
 * Attributes:
 *   src, alt, title, width, height, ooxml, layout, name, posX, posY, wrap
 */

interface DrawingAttrs {
  src: string;
  alt: string | null;
  title: string | null;
  width: number | null;
  height: number | null;
  ooxml: string | null;
  layout: "inline" | "anchor";
  name: string | null;
  posX: number | null;
  posY: number | null;
  wrap: "none" | "square" | "tight" | "through" | "topAndBottom";
}

export const Drawing = Node.create({
  name: "image",
  group: "inline",
  inline: true,
  atom: true,
  selectable: true,
  draggable: true,

  addAttributes() {
    return {
      src: {},
      alt: { default: null },
      title: { default: null },
      width: { default: null },
      height: { default: null },
      ooxml: { default: null },
      layout: { default: "inline" },
      name: { default: null },
      posX: { default: null },
      posY: { default: null },
      wrap: { default: "none" },
    } as Record<keyof DrawingAttrs, unknown>;
  },

  parseHTML() {
    return [{ tag: "img[data-image]" }, { tag: "span[data-drawing]" }];
  },

  renderHTML({ HTMLAttributes }) {
    const a = HTMLAttributes as Partial<DrawingAttrs>;
    const widthPx = a.width ? `${a.width}px` : "auto";
    const heightPx = a.height ? `${a.height}px` : "auto";

    if (a.layout === "anchor") {
      return [
        "span",
        mergeAttributes(
          {
            "data-drawing": "anchor",
            style:
              "display:inline-block;width:0;height:0;line-height:0;",
          },
          { "data-ooxml": a.ooxml ? "1" : "0" },
        ),
        [
          "img",
          {
            "data-image": "anchor",
            "data-layout": "anchor",
            src: a.src,
            alt: a.alt ?? a.name ?? "",
            title: a.title ?? a.name ?? "",
            style: `position:absolute;left:${a.posX ?? 0}px;top:${a.posY ?? 0}px;width:${widthPx};height:${heightPx};z-index:2;cursor:grab;`,
            draggable: "false",
          },
        ],
      ];
    }

    return [
      "img",
      mergeAttributes(
        {
          "data-image": "inline",
          style: `max-width:100%;height:auto;width:${widthPx};`,
        },
        HTMLAttributes as Record<string, unknown>,
      ),
    ];
  },
});

export const PageBreak = Node.create({
  name: "pageBreak",
  group: "inline",
  inline: true,
  atom: true,
  selectable: true,

  parseHTML() {
    return [{ tag: "span[data-pagebreak]" }];
  },

  renderHTML() {
    return [
      "span",
      {
        "data-pagebreak": "",
        style:
          "display:inline-block;width:100%;height:0;border-top:2px dashed #c8c6c4;margin:16px 0;position:relative;",
        "aria-label": "分页符",
      },
    ];
  },

  addCommands() {
    return {
      insertPageBreak:
        () =>
        // eslint-disable-next-line @typescript-eslint/no-explicit-any

        ({ chain }: any) =>
          chain().insertContent({ type: this.name }).run(),
    } as never;
  },
});
