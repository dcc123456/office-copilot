import { Node, mergeAttributes } from "@tiptap/core";

/**
 * Inline Word field node used for page numbers, dates, etc. in the
 * main document and in header/footer mini-editors. The renderer resolves
 * PAGE/NUMPAGES at render time so users see real numbers; on save the
 * field instruction + current display are serialized back to complex
 * field XML by docx-io.
 *
 * Attrs:
 *   kind    — "page" | "numPages" | "date" | "author" | "fileName" | custom
 *   instr   — raw instruction text (e.g. "PAGE  \\* Arabic  \\* MERGEFORMAT")
 *   display — cached value; live-updated by the Pagination layer. 字面文本，
 *             `\t` / `\n` 对应 OOXML 里的 `<w:tab/>` / `<w:br/>`（目录条目靠它对齐）
 *
 * 三个属性都自带 renderHTML/parseHTML：默认的按属性名写 DOM 属性会生成
 * `kind="page" instr="…"` 这种非 data- 属性，复制粘贴时 PM 读不回来，
 * 粘出去的域就退化成默认的 PAGE 域。
 * 注意属性级 renderHTML 收到的**第一个参数就是节点的 attrs 全体**
 * （getRenderedAttributes 里 `renderHTML(nodeOrMark.attrs)`），不是 { node }。
 */
export const Field = Node.create({
  name: "field",
  group: "inline",
  inline: true,
  atom: true,
  selectable: true,

  addAttributes() {
    return {
      kind: {
        default: "page",
        parseHTML: (element: HTMLElement) => element.getAttribute("data-field") ?? undefined,
        renderHTML: (attrs: { kind?: string }) => ({ "data-field": attrs.kind }),
      },
      instr: {
        default: "PAGE",
        parseHTML: (element: HTMLElement) => element.getAttribute("data-instr") ?? undefined,
        renderHTML: (attrs: { instr?: string }) => ({ "data-instr": attrs.instr }),
      },
      display: {
        default: "1",
        parseHTML: (element: HTMLElement) => element.getAttribute("data-display") ?? undefined,
        renderHTML: (attrs: { display?: string }) => ({ "data-display": attrs.display }),
      },
    };
  },

  parseHTML() {
    return [{ tag: "span[data-field]" }];
  },

  renderHTML({ HTMLAttributes, node }) {
    const display = String(node.attrs.display ?? "");
    return [
      "span",
      mergeAttributes(
        {
          class: "pm-field",
          // 悬停显示指令：域被折叠成一个芯片时，这是唯一能看到 `TOC \o "1-3"` 的地方。
          title: String(node.attrs.instr ?? ""),
          style:
            "display:inline-block;padding:0 4px;border-radius:3px;background:#eaf3fb;color:#0078d4;font-size:0.95em;",
        },
        HTMLAttributes,
      ),
      display || String(node.attrs.kind ?? "field"),
    ];
  },

  addCommands() {
    return {
      insertField:
        (kind: string, instr?: string, display?: string) =>
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        ({ chain }: any) =>
          chain()
            .insertContent({
              type: this.name,
              attrs: {
                kind,
                instr: instr ?? kind.toUpperCase(),
                display:
                  display ?? (kind === "date" ? new Date().toLocaleDateString() : "1"),
              },
            })
            .run(),
    } as never;
  },
});
