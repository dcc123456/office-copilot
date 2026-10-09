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
 *   display — cached value; live-updated by the Pagination layer
 */
export const Field = Node.create({
  name: "field",
  group: "inline",
  inline: true,
  atom: true,
  selectable: true,

  addAttributes() {
    return {
      kind: { default: "page" },
      instr: { default: "PAGE" },
      display: { default: "1" },
    };
  },

  parseHTML() {
    return [{ tag: "span[data-field]" }];
  },

  renderHTML({ HTMLAttributes }) {
    const a = HTMLAttributes as {
      kind: string;
      instr: string;
      display: string;
    };
    return [
      "span",
      mergeAttributes(
        {
          "data-field": a.kind,
          "data-instr": a.instr,
          class: "pm-field",
          style:
            "display:inline-block;padding:0 4px;border-radius:3px;background:#eaf3fb;color:#0078d4;font-size:0.95em;",
        },
        HTMLAttributes as Record<string, unknown>,
      ),
      a.display || a.kind,
    ];
  },

  addCommands() {
    return {
      insertField:
        (kind: string, instr?: string) =>
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        ({ chain }: any) =>
          chain()
            .insertContent({
              type: this.name,
              attrs: {
                kind,
                instr: instr ?? kind.toUpperCase(),
                display: kind === "date" ? new Date().toLocaleDateString() : "1",
              },
            })
            .run(),
    } as never;
  },
});
