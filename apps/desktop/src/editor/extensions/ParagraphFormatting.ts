import { Extension, type CommandProps } from "@tiptap/core";
import type { Node as PMNode, NodeType } from "@tiptap/pm/model";
import type { TabStop } from "@office-copilot/docx-io";

/**
 * Adds paragraph/heading formatting commands and global attributes:
 *   - lineHeight (CSS string, e.g. "1.5")
 *   - spacingBefore / spacingAfter (points, e.g. 6)
 *   - indent / indentRight (left / right indent in points)
 *   - indentFirst (first-line indent in points; negative = hanging)
 *   - keepWithNext / keepLines / widowControl / snapToGrid (Word 的分页开关)
 *   - outlineLevel (raw `w:outlineLvl`，0–8 = 一级～八级，9 = 正文文本)
 *   - styleId (段落套的具名样式 id，即 `<w:pStyle>` 的原值)
 *   - tabStops (`<w:tabs>` 的制表位清单，位置是 pt)
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
  /** null = 用户显式清掉了样式；缺失 = 没套用也没清除，导出沿用源文档写法。 */
  styleId?: string | null;
  /**
   * 制表位定义，位置是 pt（`<w:tabs>`，换算只在 docx-io）。空数组 = 用户"全部清除"，
   * 缺失 = 这一段没建模过，导出沿用源文档写法。
   */
  tabStops?: TabStop[];
}

declare module "@tiptap/core" {
  interface Commands<ReturnType> {
    paragraphFormat: {
      setParagraphFormat: (patch: Partial<ParagraphFormatAttrs>) => ReturnType;
      applyNamedStyle: (styleId: string | null, headingLevel?: number) => ReturnType;
      setTabStop: (stop: TabStop) => ReturnType;
      clearTabStop: (positionPt: number) => ReturnType;
      clearAllTabStops: () => ReturnType;
    };
  }
}

/**
 * 对选区内每个段落 / 标题块（表格单元里的也算）落地一次改动，只发一个事务。
 *
 * @param state 命令上下文的编辑器状态；选区决定作用范围。
 * @param dispatch 命令上下文的 dispatch；预览（不传 dispatch）时只算不落地。
 * @param nextBlock 给出块节点，返回要写的完整属性，可选换成别的块类型（样式套用要用）。
 * @returns 选区里有没有块；一个都没有时是 false，调用方决定退回单段写法。
 */
function mapFormatBlocks(
  state: CommandProps["state"],
  dispatch: CommandProps["dispatch"],
  nextBlock: (node: PMNode) => { attrs: Record<string, unknown>; type?: NodeType },
): boolean {
  const blocks: { node: PMNode; pos: number }[] = [];
  const { from, to } = state.selection;
  state.doc.nodesBetween(from, to, (node, pos) => {
    if (node.type.name === "paragraph" || node.type.name === "heading") {
      blocks.push({ node, pos });
    }
  });
  if (!blocks.length) return false;
  // 必须攒成一个事务：逐块各发一个基于旧 doc 的 tr，第二块起 PM 抛
  // "Applying a mismatched transaction"，选中多段时只有第一段改上。
  let tr = state.tr;
  for (const { node, pos } of blocks) {
    const next = nextBlock(node);
    tr = tr.setNodeMarkup(pos, next.type ?? node.type, next.attrs);
  }
  if (dispatch) dispatch(tr);
  return true;
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
          styleId: {
            // 默认只能是 undefined，不能像上面几项那样写 null：docx-io 里 null 表示
            // "用户清掉了样式"，拿它当默认值会把导入文档每个段落的 w:pStyle 都删掉。
            default: undefined,
            parseHTML: (el) => {
              const raw = (el as HTMLElement).getAttribute("data-style-id");
              // 属性不在 = 没建模过；空串 = 用户显式清除。
              if (raw === null) return undefined;
              return raw === "" ? null : raw;
            },
            // 只留 data-* 供 docx-io 读回：样式定义本身由 `docStyleSheet` 注入成
            // 文档作用域的样式表，行内 style（直接格式）因此继续在 Word 的优先级位置上赢。
            renderHTML: (attrs) =>
              attrs.styleId === undefined
                ? {}
                : { "data-style-id": String(attrs.styleId ?? "") },
          },
          tabStops: {
            // 与 styleId 同一条理由：默认只能是 undefined。拿 [] 当默认值等于每个段落
            // 都声明"制表位已全部清除"，导入文档里的 w:tabs 会整批消失。
            default: undefined,
            parseHTML: (el) =>
              parseTabStops((el as HTMLElement).getAttribute("data-tab-stops")),
            // 浏览器没有"按制表位对齐"的排版能力，所以只留 data-* 供 docx-io 读回；
            // 编辑器内的 Tab 仍按默认宽度显示，这是规格里写明的差距（排名 5）。
            renderHTML: (attrs) =>
              Array.isArray(attrs.tabStops)
                ? { "data-tab-stops": JSON.stringify(attrs.tabStops) }
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
        ({ commands, state, dispatch }: CommandProps) =>
          // 只补这段落自己缺的属性，已有的属性跟着 node.attrs 一起写回。
          mapFormatBlocks(state, dispatch, (node) => ({
            attrs: { ...node.attrs, ...patch },
          })) || commands.updateAttributes("paragraph", patch),

      setTabStop:
        (stop: TabStop) =>
        ({ state, dispatch }: CommandProps) =>
          // 逐块并入：Word 的「设置」只动这一个位置，别段落在别处的制表位必须留着，
          // 所以不能拿面板里那张清单整份覆盖过去。
          mapFormatBlocks(state, dispatch, (node) => ({
            attrs: { ...node.attrs, tabStops: withTabStop(stopsOf(node), stop) },
          })),

      clearTabStop:
        (positionPt: number) =>
        ({ state, dispatch }: CommandProps) =>
          mapFormatBlocks(state, dispatch, (node) => ({
            attrs: {
              ...node.attrs,
              tabStops: withoutTabStop(stopsOf(node), positionPt),
            },
          })),

      clearAllTabStops:
        () =>
        ({ state, dispatch }: CommandProps) =>
          // 与「清除」不同：这条要的就是空清单，导出因此删掉段落自己的 w:tabs。
          mapFormatBlocks(state, dispatch, (node) => ({
            attrs: { ...node.attrs, tabStops: [] },
          })),

      applyNamedStyle:
        (styleId: string | null, headingLevel?: number) =>
        ({ commands, state, dispatch }: CommandProps) => {
          // Word 套用具名样式只换样式引用（标题样式顺带把段落变成对应级别），直接格式
          // （加粗、字号、荧光笔）永远压过样式——所以这里只动块级属性与节点类型。
          const target =
            headingLevel === undefined
              ? state.schema.nodes.paragraph
              : state.schema.nodes.heading ?? state.schema.nodes.paragraph;
          if (!target) return false;
          return (
            mapFormatBlocks(state, dispatch, (node) => {
              // level 只有标题认，退回正文时 PM 会丢掉它；其余段落属性照 node.attrs 走。
              const attrs: Record<string, unknown> = { ...node.attrs, styleId };
              if (headingLevel !== undefined) attrs.level = headingLevel;
              return { attrs, type: target };
            }) || commands.updateAttributes("paragraph", { styleId })
          );
        },
    } as never;
  },
});

/** pt → CSS px（96 DPI）。编辑器里 pt 与倍数是段落属性的单位，只有显示成 CSS 时换算。 */
export function ptToPx(pt: number): number {
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
 * 这一段现有的制表位清单，命令侧的统一读法。
 *
 * @param node 段落 / 标题节点。
 * @returns 清单（可能是空表）；undefined = 这一段没建模过制表位，导出沿用源文档的 `w:tabs`。
 */
function stopsOf(node: PMNode): TabStop[] | undefined {
  return Array.isArray(node.attrs.tabStops)
    ? (node.attrs.tabStops as TabStop[])
    : undefined;
}

/**
 * 同一个位置就是同一条制表位。
 *
 * @param a 位置，pt。读侧与 `cmToPt` 都给 0.1pt 精度，所以按 0.1pt 取整后逐位比较，
 *   不设容差——差 0.1pt 的两条在 Word 里也是两条。
 * @param b 位置，pt。
 * @returns 是不是同一条。
 */
export function sameTabStopPosition(a: number, b: number): boolean {
  return Math.round(a * 10) === Math.round(b * 10);
}

/**
 * 把一个制表位并进清单：同位置整条替换（含只改对齐或前导符），否则插入，结果按位置升序。
 *
 * @param stops 这一段现有的清单；undefined 从空表开始。
 * @param stop 新的一条，位置是 pt。
 * @returns 新清单，不改动入参。
 */
export function withTabStop(
  stops: TabStop[] | undefined,
  stop: TabStop,
): TabStop[] {
  const others = (stops ?? []).filter(
    (existing) => !sameTabStopPosition(existing.positionPt, stop.positionPt),
  );
  return [...others, stop].sort((a, b) => a.positionPt - b.positionPt);
}

/**
 * 按位置移除一条制表位。
 *
 * @param stops 现有清单；undefined 时原样返回 undefined，因为"这一位本来没有制表位"与
 *   "用户全部清除过"（空表）在导出是两种写法，不能混。
 * @param positionPt 位置，pt。
 * @returns 移除后的清单（清光了就是空表），或 undefined。
 */
export function withoutTabStop(
  stops: TabStop[] | undefined,
  positionPt: number,
): TabStop[] | undefined {
  if (!stops) return undefined;
  return stops.filter((existing) => !sameTabStopPosition(existing.positionPt, positionPt));
}

/**
 * 读 HTML 上的制表位清单（`data-tab-stops`，我们自己的 renderHTML 写的 JSON）。
 *
 * @param raw 属性值；没有这个属性时是 null。
 * @returns 制表位数组（空数组是"全部清除"，必须留着）；读不出形状时返回 undefined，
 *   含义是"这一段没建模过制表位"，导出因此沿用源文档写法。对齐方式与前导符不在这里
 *   校验——docx-io 写 `w:tab` 时才认那张白名单，校验点只留一处。
 */
function parseTabStops(raw: string | null): TabStop[] | undefined {
  if (raw === null || raw === "") return undefined;
  let parsed: unknown;
  try {
    parsed = JSON.parse(raw);
  } catch {
    return undefined;
  }
  if (!Array.isArray(parsed)) return undefined;
  const isStop = (stop: unknown): stop is TabStop => {
    const { positionPt, alignment } = (stop ?? {}) as Partial<TabStop>;
    return (
      typeof positionPt === "number" &&
      Number.isFinite(positionPt) &&
      positionPt >= 0 &&
      typeof alignment === "string"
    );
  };
  return parsed.every(isStop) ? (parsed as TabStop[]) : undefined;
}

/**
 * "pt:29" is Word's 固定值 29 磅 and needs a concrete CSS length; "atLeast:9"
 * is its 最小值, which only floors the line — a smaller value than the text's
 * natural height must not collapse it, hence max() against 1.2em. A bare
 * number is already a valid CSS multiplier. The attr keeps its original form
 * through data-line-height so re-applying a style stays a no-op.
 */
export function lineHeightCss(value: string): string {
  const atLeast = /^atLeast:(\d+(?:\.\d+)?)$/.exec(value);
  if (atLeast?.[1]) return `max(${ptToPx(Number(atLeast[1]))}px, 1.2em)`;
  const m = /^pt:(\d+(?:\.\d+)?)$/.exec(value);
  return m ? `${ptToPx(Number(m[1]))}px` : value;
}
