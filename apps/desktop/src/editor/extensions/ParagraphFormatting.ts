import { Extension, type CommandProps } from "@tiptap/core";

/**
 * Adds paragraph/heading formatting commands and global attributes:
 *   - lineHeight (CSS string, e.g. "1.5")
 *   - spacingBefore / spacingAfter (points, e.g. 6)
 *   - indent / indentRight (left / right indent in points)
 *   - indentFirst (first-line indent in points; negative = hanging)
 *   - keepWithNext / keepLines / widowControl / snapToGrid (Word 的分页开关)
 *   - outlineLevel (raw `w:outlineLvl`，0–8 = 一级～八级，9 = 正文文本)
 *
 * These are surfaced to the user via ParagraphPopover and round-tripped
 * by docx-io: the OOXML converter emits <w:spacing w:before/after=...
 * w:line=... w:lineRule="auto"/>, <w:ind w:left/right=... w:firstLine/hanging=.../>
 * and the <w:keepNext/> family when these attrs are present, even if no
 * rawPpr was preserved.
 *
 * We store "points" here for familiarity; 1pt = 20 twips in OOXML.
 */

interface ParagraphFormatAttrs {
  lineHeight?: string;
  spacingBefore?: number;
  spacingAfter?: number;
  indent?: number;
  indentRight?: number;
  indentFirst?: number;
  keepWithNext?: boolean;
  keepLines?: boolean;
  widowControl?: boolean;
  snapToGrid?: boolean;
  outlineLevel?: number;
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
          indentRight: {
            default: null,
            parseHTML: (el) =>
              Number((el as HTMLElement).getAttribute("data-indent-right")) ||
              null,
            renderHTML: (attrs) =>
              attrs.indentRight != null
                ? {
                    "data-indent-right": String(attrs.indentRight),
                    style: `margin-right: ${ptToPx(Number(attrs.indentRight))}px`,
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
          keepWithNext: {
            default: null,
            parseHTML: (el) =>
              onOffFromHtml(el as HTMLElement, "data-keep-with-next"),
            renderHTML: (attrs) =>
              switchHtml(
                "data-keep-with-next",
                attrs.keepWithNext,
                (on) => (on ? "break-after: avoid" : ""),
              ),
          },
          keepLines: {
            default: null,
            parseHTML: (el) => onOffFromHtml(el as HTMLElement, "data-keep-lines"),
            renderHTML: (attrs) =>
              switchHtml(
                "data-keep-lines",
                attrs.keepLines,
                (on) => (on ? "break-inside: avoid" : ""),
              ),
          },
          widowControl: {
            default: null,
            parseHTML: (el) =>
              onOffFromHtml(el as HTMLElement, "data-widow-control"),
            renderHTML: (attrs) =>
              switchHtml(
                "data-widow-control",
                attrs.widowControl,
                (on) => `widows: ${on ? 2 : 1}; orphans: ${on ? 2 : 1}`,
              ),
          },
          snapToGrid: {
            default: null,
            parseHTML: (el) => onOffFromHtml(el as HTMLElement, "data-snap-to-grid"),
            // Word 的"对齐到网格"是文档网格，浏览器排版里没有对应物，只留 data-*
            // 让 docx-io 读得回来。
            renderHTML: (attrs) =>
              switchHtml("data-snap-to-grid", attrs.snapToGrid, () => ""),
          },
          outlineLevel: {
            default: null,
            parseHTML: (el) => {
              const raw = (el as HTMLElement).getAttribute("data-outline-level");
              if (raw === null || raw === "") return null;
              const n = Number(raw);
              return Number.isInteger(n) && n >= 0 && n <= 9 ? n : null;
            },
            renderHTML: (attrs) =>
              typeof attrs.outlineLevel === "number"
                ? { "data-outline-level": String(attrs.outlineLevel) }
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
 * 读 HTML 上的 Word 开关属性。
 *
 * @param el      被解析的段落元素。
 * @param dataKey data-* 属性名，例如 "data-keep-with-next"。
 * @returns true/false；元素上没有该属性时返回 null，表示"用户没设置过"，
 *          这与 false（Word 明确写了 w:val="0"）必须区分，否则导出会凭空加属性。
 */
function onOffFromHtml(el: HTMLElement, dataKey: string): boolean | null {
  const raw = el.getAttribute(dataKey);
  if (raw === null || raw === "") return null;
  return raw !== "false" && raw !== "0";
}

/**
 * Word 开关属性写回 HTML：data-* 连 false 一起写，重开编辑器才分得清"关"和"没设置"。
 *
 * @param dataKey data-* 属性名。
 * @param value   段落属性值；不是布尔（未设置）时返回空对象，即什么都不渲染。
 * @param cssFor  由开关状态求 CSS 声明，返回空串表示浏览器排版没有等价物。
 */
function switchHtml(
  dataKey: string,
  value: unknown,
  cssFor: (on: boolean) => string,
): Record<string, string> {
  if (typeof value !== "boolean") return {};
  const css = cssFor(value);
  return css
    ? { [dataKey]: String(value), style: css }
    : { [dataKey]: String(value) };
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
