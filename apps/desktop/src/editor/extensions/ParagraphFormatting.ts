import { Extension, type CommandProps } from "@tiptap/core";

/**
 * Adds paragraph/heading formatting commands and global attributes:
 *   - lineHeight (CSS string, e.g. "1.5")
 *   - spacingBefore / spacingAfter (points, e.g. 6)
 *   - indent (left indent in points)
 *   - indentFirst (first-line indent in points)
 *
 * These are surfaced to the user via ParagraphPopover and round-tripped
 * by docx-io: the OOXML converter emits <w:spacing w:before/after=...
 * w:line=... w:lineRule="auto"/> and <w:ind w:left=... w:firstLine=.../>
 * when these attrs are present, even if no rawPpr was preserved.
 *
 * We store "points" here for familiarity; 1pt = 20 twips in OOXML.
 */

interface ParagraphFormatAttrs {
  lineHeight?: string;
  spacingBefore?: number;
  spacingAfter?: number;
  indent?: number;
  indentFirst?: number;
}

declare module "@tiptap/core" {
  interface Commands<ReturnType> {
    paragraphFormat: {
      setParagraphFormat: (patch: Partial<ParagraphFormatAttrs>) => ReturnType;
    };
  }
}

export const ParagraphFormatting = Extension.create({
  name: "paragraphFormatting",

  addGlobalAttributes() {
    return [
      {
        types: ["paragraph", "heading"],
        attributes: {
          lineHeight: {
            default: null,
            parseHTML: (el) =>
              (el as HTMLElement).getAttribute("data-line-height") ||
              (el as HTMLElement).style.lineHeight ||
              null,
            renderHTML: (attrs) => {
              if (!attrs.lineHeight) return {};
              const v = String(attrs.lineHeight);
              return {
                "data-line-height": v,
                style: `line-height: ${lineHeightCss(v)}`,
              };
            },
          },
          spacingBefore: {
            default: null,
            parseHTML: (el) =>
              Number((el as HTMLElement).getAttribute("data-spacing-before")) ||
              null,
            renderHTML: (attrs) =>
              attrs.spacingBefore != null
                ? {
                    "data-spacing-before": String(attrs.spacingBefore),
                    style: `margin-top: ${ptToPx(Number(attrs.spacingBefore))}px`,
                  }
                : {},
          },
          spacingAfter: {
            default: null,
            parseHTML: (el) =>
              Number((el as HTMLElement).getAttribute("data-spacing-after")) ||
              null,
            renderHTML: (attrs) =>
              attrs.spacingAfter != null
                ? {
                    "data-spacing-after": String(attrs.spacingAfter),
                    style: `margin-bottom: ${ptToPx(Number(attrs.spacingAfter))}px`,
                  }
                : {},
          },
          indent: {
            default: null,
            parseHTML: (el) =>
              Number((el as HTMLElement).getAttribute("data-indent")) || null,
            renderHTML: (attrs) =>
              attrs.indent != null
                ? {
                    "data-indent": String(attrs.indent),
                    style: `margin-left: ${ptToPx(Number(attrs.indent))}px`,
                  }
                : {},
          },
          indentFirst: {
            default: null,
            parseHTML: (el) =>
              Number((el as HTMLElement).getAttribute("data-indent-first")) ||
              null,
            renderHTML: (attrs) =>
              attrs.indentFirst != null
                ? {
                    "data-indent-first": String(attrs.indentFirst),
                    style: `text-indent: ${ptToPx(Number(attrs.indentFirst))}px`,
                  }
                : {},
          },
        },
      },
    ];
  },

  addCommands() {
    return {
      setParagraphFormat:
        (patch: Partial<ParagraphFormatAttrs>) =>
        ({ commands, state, dispatch }: CommandProps) => {
          // Apply to every block in the current selection.
          const { from, to } = state.selection;
          let success = false;
          state.doc.nodesBetween(from, to, (node, pos) => {
            if (!["paragraph", "heading"].includes(node.type.name)) return;
            const attrs = { ...node.attrs, ...patch };
            if (dispatch) {
              dispatch(state.tr.setNodeMarkup(pos, undefined, attrs));
            }
            success = true;
          });
          return success || commands.updateAttributes("paragraph", patch);
        },
    } as never;
  },
});

function ptToPx(pt: number): number {
  // 1pt = 1/72 inch; at 96 DPI that's 96/72 = 1.3333 px.
  return Math.round(pt * (96 / 72) * 10) / 10;
}

/**
 * "pt:29" is Word's 固定值 29 磅 and needs a concrete CSS length; "atLeast:9"
 * is its 最小值, which only floors the line — a smaller value than the text's
 * natural height must not collapse it, hence max() against 1.2em. A bare
 * number is already a valid CSS multiplier. The attr keeps its original form
 * through data-line-height so re-applying a style stays a no-op.
 */
function lineHeightCss(value: string): string {
  const atLeast = /^atLeast:(\d+(?:\.\d+)?)$/.exec(value);
  if (atLeast?.[1]) return `max(${ptToPx(Number(atLeast[1]))}px, 1.2em)`;
  const m = /^pt:(\d+(?:\.\d+)?)$/.exec(value);
  return m ? `${ptToPx(Number(m[1]))}px` : value;
}
