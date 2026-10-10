import {
  children,
  collectDescendants,
  firstChild,
  isElement,
  parseDocument,
  parseFragment,
  serialize,
  type Element,
} from "./xml.js";
import type { PPrAttrs, PMDoc, PMMark, PMNode, PMTextNode } from "./pm-schema.js";
import { NumberingMap, readListProps, type NumberingKind } from "./numbering.js";
import { StyleMap } from "./styles.js";
import {
  anchorGeom,
  hasPicture,
  shapesFromPict,
  shapesIn,
  spliceTxbxContents,
  type ShapeInfo,
} from "./textbox.js";

/**
 * OOXML (Word .docx) ⇄ ProseMirror JSON converter.
 *
 * Design goals:
 *   - Preserve text, inline formatting, heading levels, lists, tables, images
 *     and hyperlinks with high fidelity.
 *   - Absolutely-positioned text boxes become `floatGroup`/`floatBox` blocks,
 *     because template-driven documents (resumes especially) keep all of their
 *     text inside them.
 *   - Any other block we don't model (SmartArt, charts, section properties,
 *     drawing canvas) becomes an `unmanagedBlock` whose `ooxml` attribute is
 *     the serialized XML; `pmToOoxml` writes it back verbatim.
 *   - The ProseMirror JSON shape matches the TipTap extensions in apps/desktop.
 *
 * The converter is schema-tolerant: missing numbering definitions, malformed
 * nested runs, etc., degrade gracefully instead of throwing.
 */

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
// Namespaces `a:`/`pic:` are referenced in image XML; we spell them inline.
/**
 * Prefixes a round-tripped floating drawing can reference (`mc:AlternateContent`
 * branches, `wps:`/`wpg:` shapes, the `v:`/`o:`/`w10:` VML fallback, `wp14:`
 * edit ids, `w14:`/`w15:` paragraph ids). They must be declared on the document
 * root or Word refuses to open the file we wrote, so we declare them even when
 * a given document does not use all of them.
 */
const DRAWING_NS = {
  mc: "http://schemas.openxmlformats.org/markup-compatibility/2006",
  wpc: "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
  wpg: "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
  wps: "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
  wp14: "http://schemas.microsoft.com/drawingml/2010/wordprocessingDrawing",
  w14: "http://schemas.microsoft.com/office/word/2010/wordml",
  w15: "http://schemas.microsoft.com/office/word/2012/wordml",
  a14: "http://schemas.microsoft.com/office/drawing/2010/main",
  v: "urn:schemas-microsoft-com:vml",
  o: "urn:schemas-microsoft-com:office:office",
  w10: "urn:schemas-microsoft-com:office:word",
} as const;

/**
 * Word stores character sizes in points; the editor renders and stores CSS px,
 * because `style="font-size: 12"` is invalid CSS and was silently ignored.
 * A bare number is treated as points — that is what this converter wrote before
 * the unit became explicit, so existing documents keep their size.
 */
export function ptToCssFontSize(pt: number): string {
  return `${Math.round(pt * (96 / 72) * 10) / 10}px`;
}

/**
 * Letter spacing uses the same on-screen unit as 字号, but at two decimals:
 * a twentieth of a point is the smallest thing OOXML can express, and 0.5 磅
 * rounded to 0.7px converts back to 0.525 磅 — which rounds to 11 twips, not
 * 10, so a saved file would never stop disagreeing with the template.
 */
export function ptToCssLetterSpacing(pt: number): string {
  return `${Math.round(pt * (96 / 72) * 100) / 100}px`;
}

/** Negative tightens. rPr `w:spacing` allows it, so the parser must too. */
export function cssLetterSpacingToPt(value: string): number | null {
  const m = /^\s*(-?\d+(?:\.\d+)?)\s*(px|pt)?\s*$/.exec(value);
  if (!m || !m[1]) return null;
  const n = Number(m[1]);
  if (!Number.isFinite(n)) return null;
  return m[2] === "px" ? (n * 72) / 96 : n;
}

export function cssFontSizeToPt(value: string): number | null {
  const m = /^\s*(\d+(?:\.\d+)?)\s*(px|pt)?\s*$/.exec(value);
  if (!m || !m[1]) return null;
  const n = Number(m[1]);
  if (!Number.isFinite(n) || n <= 0) return null;
  return m[2] === "px" ? (n * 72) / 96 : n;
}

/**
 * OOXML colours are bare hex. Mark attrs are hex for anything this app or a
 * .docx produced, but pasted HTML lands as `rgb(...)` — and a value Word cannot
 * read invalidates the whole run properties, so an unreadable colour is dropped
 * rather than written.
 */
export function cssColorToOoxmlHex(value: unknown): string | null {
  if (typeof value !== "string") return null;
  const v = value.trim();
  const hex = /^#([0-9a-f]{3}|[0-9a-f]{6})$/i.exec(v);
  if (hex?.[1]) {
    const h = hex[1];
    const full = h.length === 3 ? h.split("").map((c) => c + c).join("") : h;
    return /^[0-9a-f]{6}$/.test(full) ? full.toUpperCase() : null;
  }
  const rgb = /^rgba?\(\s*(\d+)[\s,]+(\d+)[\s,]+(\d+)/i.exec(v);
  if (rgb) {
    return rgb
      .slice(1, 4)
      .map((n) => Math.min(255, Math.max(0, Number(n))).toString(16).padStart(2, "0"))
      .join("")
      .toUpperCase();
  }
  return null;
}

const NAMED_CSS_COLORS: Record<string, string> = {
  black: "000000",
  white: "ffffff",
  red: "ff0000",
  green: "008000",
  blue: "0000ff",
  yellow: "ffff00",
  orange: "ffa500",
  purple: "800080",
  gray: "808080",
  grey: "808080",
  silver: "c0c0c0",
  maroon: "800000",
  navy: "000080",
  teal: "008080",
  aqua: "00ffff",
  cyan: "00ffff",
  fuchsia: "ff00ff",
  lime: "00ff00",
  olive: "808000",
};

/**
 * The single canonical form of a colour: lowercase `#rrggbb`.
 *
 * Needed because the editor writes `style="color: #c00000"` and reads it back as
 * `rgb(192, 0, 0)`; comparing those two verbatim makes "does this document
 * already match the template?" answer wrongly, and re-applying a template
 * produces edits that change nothing. Everything that lands in a style patch or
 * a mark attr goes through here first. `auto`, `transparent` and unparseable
 * values become null — meaning "no colour set".
 */
export function canonColor(value: unknown): string | null {
  if (typeof value !== "string") return null;
  const v = value.trim().toLowerCase();
  if (!v || v === "auto" || v === "transparent" || v === "inherit" || v === "currentcolor") {
    return null;
  }
  const named = NAMED_CSS_COLORS[v];
  if (named) return `#${named}`;
  const hex = cssColorToOoxmlHex(v);
  return hex ? `#${hex.toLowerCase()}` : null;
}

const PT_LINE_HEIGHT = /^pt:(\d+(?:\.\d+)?)$/;
const AT_LEAST_LINE_HEIGHT = /^atLeast:(\d+(?:\.\d+)?)$/;
/** OOXML writes colours as bare hex; `canonColor` wants the CSS form. */
function colorVal(v: string | null | undefined): string | null {
  if (!v) return null;
  const t = v.trim();
  if (!t) return null;
  return /^([0-9a-f]{3}|[0-9a-f]{6})$/i.test(t) ? `#${t}` : t;
}
/**
 * Three `lineHeight` forms exist: a bare multiplier ("1.5"), an exact value
 * ("pt:29") which is Word's 固定值 — what 公文 (GB/T 9704) and most 学位论文
 * specs require — and a minimum ("atLeast:9", Word's 最小值). They map to
 * different w:line units: 240ths of a line for lineRule="auto", twentieths of
 * a point for lineRule="exact"/"atLeast".
 */
export function lineHeightSpacingAttrs(value: string | undefined): string[] {
  if (!value) return [];
  const atLeast = AT_LEAST_LINE_HEIGHT.exec(value);
  if (atLeast?.[1]) {
    const pt = Number(atLeast[1]);
    return pt > 0 ? [`w:line="${Math.round(pt * 20)}"`, `w:lineRule="atLeast"`] : [];
  }
  const exact = PT_LINE_HEIGHT.exec(value);
  if (exact?.[1]) {
    const pt = Number(exact[1]);
    return pt > 0 ? [`w:line="${Math.round(pt * 20)}"`, `w:lineRule="exact"`] : [];
  }
  const n = Number(value);
  if (!Number.isFinite(n) || n <= 0) return [];
  return [`w:line="${Math.round(n * 240)}"`, `w:lineRule="auto"`];
}

export function ooxmlToPm(
  xml: string,
  numbering?: NumberingMap,
  styles?: StyleMap,
): PMDoc {
  const root = parseDocument(xml);
  const body = firstChild(root, "w:body");
  if (!body) return { type: "doc", content: [] };
  const raw: PMNode[] = [];
  for (const child of children(body)) {
    // Section properties are emitted by pmToOoxml itself; skip here.
    if (child.name === "w:sectPr") continue;
    raw.push(...blockNodesFromXml(child, numbering, styles));
  }
  return { type: "doc", content: groupListItems(raw) };
}

/**
 * How the document paints hyperlinks. Word keeps this in the `Hyperlink`
 * character style; we write it into each link run instead, because run-level
 * rPr is what every consumer (Word, WPS, our own reader) actually honours.
 */
export interface LinkStyle {
  /** CSS colour, canonical `#rrggbb`; bare/lower-case forms are accepted. */
  color?: string | null;
  underline?: boolean;
}

type WriteCtx = { numbering?: NumberingMap; link?: LinkStyle | null };

export function pmToOoxml(doc: PMDoc, ctx: WriteCtx = {}): string {
  const bodyChildren = doc.content.map((b) => blockToXml(b, ctx)).join("");
  const ns = Object.entries(DRAWING_NS)
    .map(([p, uri]) => ` xmlns:${p}="${uri}"`)
    .join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"${ns}>
<w:body>${bodyChildren}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body>
</w:document>`;
}

// ---------- OOXML → PM ----------

/**
 * One body-level element can yield several blocks: a paragraph that carries
 * floating drawings contributes them as positioned blocks *ahead of itself*,
 * then contributes its own paragraph.
 *
 * The paragraph always survives. Word measures every anchor's offset from the
 * top of its host paragraph, so a paragraph that shows nothing but a page
 * break still holds a line in the flow — drop it and the next page's offsets
 * are measured from the wrong place, and the break ends up ahead of the very
 * boxes that belong on the page it closes. Boxes go first because they take no
 * height, which keeps their top equal to the paragraph's own.
 */
function blockNodesFromXml(
  el: Element,
  numbering?: NumberingMap,
  styles?: StyleMap,
): PMNode[] {
  if (el.name === "w:p") {
    const slots: PMNode[] = [];
    const para = paragraphFromXml(el, numbering, styles, slots);
    if (!slots.length) return [para];
    return [...slots, para];
  }
  const node = blockNodeFromXml(el, numbering, styles);
  return node ? [node] : [];
}

function blockNodeFromXml(
  el: Element,
  numbering?: NumberingMap,
  styles?: StyleMap,
): PMNode | null {
  switch (el.name) {
    case "w:p":
      return paragraphFromXml(el, numbering, styles);
    case "w:tbl":
      return tableFromXml(el, numbering, styles);
    case "w:sectPr":
      return unmanaged(el, "section");
    default:
      return unmanaged(el, el.name.replace(/^w:/, ""));
  }
}

function paragraphFromXml(
  p: Element,
  numbering?: NumberingMap,
  styles?: StyleMap,
  floats?: PMNode[],
): PMNode {
  const pPr = firstChild(p, "w:pPr");
  const style = val(pPr && firstChild(pPr, "w:pStyle"));
  const { numId, ilvl } = readListProps(p);
  const jc = val(pPr && firstChild(pPr, "w:jc"));
  const textAlign = jcToTextAlign(jc);

  const inline = collectInline(p, { floats, numbering, styles });
  const fmt = readParagraphFormat(pPr);

  // Capture the full pPr inner XML so we can write it back verbatim. We
  // keep numPr too: when a list item is regrouped and re-emitted the
  // numPr we prepend in listItemXml wins and the original acts as a
  // fallback that Word tolerates.
  let rawPpr: string | undefined;
  if (pPr) {
    const inner = children(pPr).map((c) => serialize(c)).join("");
    if (inner.trim()) rawPpr = inner;
  }

  if (numId !== undefined) {
    const lvl = numbering?.level(numId, ilvl);
    const ordered = lvl?.kind === "ordered";
    const item: PMNode = {
      type: "listItem",
      content: inline.length ? [{ type: "paragraph", content: inline } as PMNode] : [],
    };
    (item as unknown as { __list?: { ordered: boolean; level: number; numId?: string } }).__list = {
      ordered,
      level: ilvl,
      numId,
    };
    return item;
  }

  const declaredLevel = lvlFromStyles(styles, style);
  if (
    declaredLevel ||
    style?.startsWith("Heading") ||
    style?.startsWith("heading")
  ) {
    const level = clampHeadingLevel(declaredLevel ?? headingLevel(style));
    return {
      type: "heading",
      attrs: { level, ...(textAlign ? { textAlign } : {}), ...fmt },
      content: inline,
    };
  }

  // Some exporters use a bare numeric styleId with no name; treat 1/2/3 as
  // headings only when styles.xml never declared the id — WPS documents give
  // ordinary list paragraphs ids like "1", and guessing there inflates them.
  // Deliberately still capped at 3: nothing about a lone digit says "level 5",
  // and the 4-6 levels above come from a name that actually states them.
  if (style && /^[1-9]$/.test(style) && !styles?.isDeclared(style)) {
    const n = Number(style);
    const level = (n === 1 ? 1 : n === 2 ? 2 : 3) as 1 | 2 | 3;
    return {
      type: "heading",
      attrs: { level, ...(textAlign ? { textAlign } : {}), ...fmt },
      content: inline,
    };
  }

  if (style === "Quote" || style === "IntenseQuote" || style === "BlockText" || style === "Title") {
    return { type: "blockquote", content: [{ type: "paragraph", content: inline }] };
  }

  return {
    type: "paragraph",
    ...(textAlign || rawPpr || Object.keys(fmt).length
      ? {
          attrs: {
            ...(textAlign ? { textAlign } : {}),
            ...(rawPpr ? { rawPpr } : {}),
            ...fmt,
          },
        }
      : {}),
    content: inline,
  };
}

/**
 * Read <w:spacing> and <w:ind> values into PM attrs (in points / CSS
 * ratio). All unit conversions happen here so the UI can stay in
 * familiar points.
 */
function readParagraphFormat(pPr: Element | null | undefined): PPrAttrs {
  const out: PPrAttrs = {};
  if (!pPr) return out;
  const keepNext = onOffValue(firstChild(pPr, "w:keepNext"));
  if (keepNext !== undefined) out.keepWithNext = keepNext;
  const keepLines = onOffValue(firstChild(pPr, "w:keepLines"));
  if (keepLines !== undefined) out.keepLines = keepLines;
  const breakBefore = onOffValue(firstChild(pPr, "w:pageBreakBefore"));
  if (breakBefore !== undefined) out.pageBreakBefore = breakBefore;
  const widow = onOffValue(firstChild(pPr, "w:widowControl"));
  if (widow !== undefined) out.widowControl = widow;
  const grid = onOffValue(firstChild(pPr, "w:snapToGrid"));
  if (grid !== undefined) out.snapToGrid = grid;
  // 大纲级别：OOXML 存 0-9（0 = Word 的"1 级"，9 = 正文文本）。超出范围说明
  // 是别的写入方乱填的，当作没建模过，写回时原样透传。
  const lvl = Number(val(firstChild(pPr, "w:outlineLvl")));
  if (Number.isInteger(lvl) && lvl >= 0 && lvl <= 9) out.outlineLevel = lvl;
  const spacing = firstChild(pPr, "w:spacing");
  if (spacing) {
    // w:before / w:after are in twentieths of a point (1pt = 20).
    const before = spacing.attributes?.["w:before"] as string | undefined;
    const after = spacing.attributes?.["w:after"] as string | undefined;
    if (before !== undefined) {
      const n = Number(before) / 20;
      if (Number.isFinite(n) && n !== 0) out.spacingBefore = Math.round(n * 10) / 10;
    }
    if (after !== undefined) {
      const n = Number(after) / 20;
      if (Number.isFinite(n) && n !== 0) out.spacingAfter = Math.round(n * 10) / 10;
    }
    // w:line with lineRule="auto" is in 240ths of a line (240=single).
    // Without lineRule, "auto" is implied for normal documents.
    // lineRule="exact" is Word's 固定值, "atLeast" its 最小值 — both are
    // twentieths of a point, but 最小值 never squeezes a line below the text.
    const line = spacing.attributes?.["w:line"] as string | undefined;
    const lineRule = spacing.attributes?.["w:lineRule"] as string | undefined;
    if (line !== undefined && (lineRule === "exact" || lineRule === "atLeast")) {
      const pt = Number(line) / 20;
      if (Number.isFinite(pt) && pt > 0) {
        const v = Math.round(pt * 10) / 10;
        out.lineHeight = `${lineRule === "atLeast" ? "atLeast" : "pt"}:${v}`;
      }
    } else if (line !== undefined && (lineRule === "auto" || !lineRule)) {
      const n = Number(line) / 240;
      if (Number.isFinite(n) && n > 0 && n !== 1) {
        out.lineHeight = String(Math.round(n * 100) / 100);
      }
    }
  }
  const ind = firstChild(pPr, "w:ind");
  if (ind) {
    // w:left / w:right / w:firstLine / w:hanging are in twips (20 per point).
    // w:start / w:end are the Office 2010 synonyms of w:left / w:right; reading
    // only the legacy names silently loses indentation in newer documents.
    const left = twipsToPoints(ind, "w:left", "w:start");
    if (left !== undefined) out.indent = left;
    const right = twipsToPoints(ind, "w:right", "w:end");
    if (right !== undefined) out.indentRight = right;
    // firstLine and hanging are mutually exclusive. hanging produces a
    // negative first-line indent (first line starts to the left of the
    // left indent); ignoring it is what makes tightly-set table cells
    // collapse character-by-character and balloon in height.
    const firstLine = twipsToPoints(ind, "w:firstLine");
    const hanging = twipsToPoints(ind, "w:hanging");
    if (firstLine !== undefined) out.indentFirst = firstLine;
    else if (hanging !== undefined) out.indentFirst = -hanging;
  }
  return out;
}

/**
 * 读 OOXML 的开关型段落属性（CT_OnOff，如 `w:keepNext` / `w:widowControl`）。
 *
 * @param el `<w:pPr>` 里的子元素；不存在时传 undefined/null。
 * @returns 裸元素（`<w:keepNext/>`）为 true，`w:val` 是 0/false/off 为 false；
 *   元素缺失或值读不懂时返回 undefined，含义是"这一段没建模过它"，
 *   写回时让 rawPpr 里的原样透传，而不是替 Word 补一个它没写过的值。
 */
function onOffValue(el: Element | null | undefined): boolean | undefined {
  if (!el) return undefined;
  const raw = el.attributes?.["w:val"];
  if (raw === undefined) return true;
  const v = String(raw).toLowerCase();
  if (v === "0" || v === "false" || v === "off") return false;
  return v === "1" || v === "true" || v === "on" ? true : undefined;
}

/**
 * 写 OOXML 的开关型段落属性，是 `onOffValue` 的逆运算。
 *
 * @param tag 元素名（含前缀），如 "w:keepNext"。
 * @param value 段落 attrs 里的建模值；undefined 时不产出元素。
 * @returns 元素字符串；true 用 Word 自己的自闭合写法，false 显式 `w:val="0"`。
 */
function onOffXml(tag: string, value: unknown): string {
  if (value !== true && value !== false) return "";
  return value ? `<${tag}/>` : `<${tag} w:val="0"/>`;
}

/**
 * 读一个 twips 属性并换算成 pt（1pt = 20 twips）。
 *
 * @param names 属性名，按优先级排列；第一个存在的生效，用于 w:left 与 w:start 这类同义写法。
 * @returns pt 值（保留一位小数）；属性缺失、非数字或为 0 时返回 undefined，
 *   因为 0 意味着"这里没建模过"，写回时应保持源文档原样而不是补一个 0。
 */
function twipsToPoints(
  el: Element,
  ...names: string[]
): number | undefined {
  for (const name of names) {
    const raw = el.attributes?.[name] as string | undefined;
    if (raw === undefined) continue;
    const n = Number(raw) / 20;
    if (Number.isFinite(n) && n !== 0) return Math.round(n * 10) / 10;
  }
  return undefined;
}

function jcToTextAlign(jc: string | undefined): string | undefined {
  switch (jc) {
    case "center":
      return "center";
    case "right":
      return "right";
    case "both":
    case "distribute":
      return "justify";
    case "left":
    default:
      return undefined;
  }
}

/** ProseMirror's heading attr, and Word's, top out at 6. */
type HeadingLvl = 1 | 2 | 3 | 4 | 5 | 6;

function clampHeadingLevel(n: number): HeadingLvl {
  return (n < 1 ? 1 : n > 6 ? 6 : Math.round(n)) as HeadingLvl;
}

function headingLevel(style: string | undefined): number {
  if (!style) return 1;
  const m = style.match(/(\d)/);
  if (!m || !m[1]) return 1;
  return Number(m[1]);
}

function lvlFromStyles(styles: StyleMap | undefined, style: string | undefined): number | undefined {
  if (!styles || !style) return undefined;
  return styles.headingLevel(style);
}

function isOrderedList(_numId: string | undefined): boolean {
  // Retained for potential future use; real answer now comes from NumberingMap.
  return false;
}

/** Stand-in for a floating box inside a run's inline array, matched by identity. */
const FLOAT_SLOT_MARK = { type: "floatSlotMark" } as unknown as PMNode;

/**
 * Context threaded through the inline converters. `floats` collects the
 * floating boxes discovered inside a paragraph: they come back as blocks
 * rather than inline nodes, so the paragraph's caller can interleave them with
 * whatever inline content survives.
 */
interface InlineCtx {
  floats?: PMNode[];
  numbering?: NumberingMap;
  styles?: StyleMap;
  /**
   * Inside a <w:hyperlink>. Word paints links with its own blue + underline,
   * and if that decoration is read into the run's style the link disagrees
   * with the sentence around it on `color` — which makes every paragraph
   * containing a link unstyleable. The decoration belongs to the document
   * instead, so it is dropped here and reinjected on write.
   */
  inLink?: boolean;
}

function collectInline(p: Element, ctx: InlineCtx = {}): PMNode[] {
  const out: PMNode[] = [];
  // State for complex fields: fldChar begin/separate/end pairs span
  // multiple runs. We buffer the instrText and the cached result so
  // they become one PM field node.
  let fieldInstr: string | null = null;
  let fieldDisplay: string[] = [];

  const flushField = () => {
    if (fieldInstr === null) return;
    const instr = fieldInstr.trim();
    const display = fieldDisplay.join("").trim() || placeholderFor(instr);
    out.push({
      type: "field",
      attrs: {
        kind: fieldKind(instr),
        instr,
        display,
      },
    });
    fieldInstr = null;
    fieldDisplay = [];
  };

  for (const child of children(p)) {
    if (child.name === "w:r") {
      // Inspect run children for fldChar / instrText so a multi-run
      // complex field collapses into one field node.
      const rChildren = children(child);
      let consumed = false;
      for (const sub of rChildren) {
        if (sub.name === "w:fldChar") {
          const t = sub.attributes?.["w:fldCharType"] as string | undefined;
          if (t === "begin") {
            fieldInstr = "";
            fieldDisplay = [];
          } else if (t === "separate") {
            // Switch from collecting instr to collecting display text.
          } else if (t === "end") {
            flushField();
          }
          consumed = true;
        } else if (sub.name === "w:instrText" && fieldInstr !== null) {
          fieldInstr += textOf(sub);
          consumed = true;
        } else if (sub.name === "w:t" && fieldInstr !== null) {
          // The cached result sits between separate and end.
          fieldDisplay.push(textOf(sub));
          consumed = true;
        }
      }
      if (!consumed) out.push(...runToInlines(child, ctx));
    } else if (child.name === "w:fldSimple") {
      // Simple field: <w:fldSimple w:instr="PAGE"><w:r>...</w:r></w:fldSimple>
      const instr = (child.attributes?.["w:instr"] as string | undefined) ?? "";
      const texts = collectDescendants(child, "w:t").map(textOf).join("");
      out.push({
        type: "field",
        attrs: {
          kind: fieldKind(instr),
          instr,
          display: texts.trim() || placeholderFor(instr),
        },
      });
    } else if (child.name === "w:hyperlink") {
      const href = child.attributes?.["r:id"] as string | undefined;
      const anchor = child.attributes?.["w:anchor"] as string | undefined;
      const url = anchor ? `#${anchor}` : href ? `relationship:${href}` : undefined;
      const inline: PMNode[] = [];
      for (const sub of children(child)) {
        if (sub.name === "w:r") inline.push(...runToInlines(sub, { ...ctx, inLink: true }));
      }
      if (url) addMark(inline, { type: "link", attrs: { href: url } });
      out.push(...inline);
    } else if (child.name === "w:bookmarkStart" || child.name === "w:bookmarkEnd") {
      // bookmarks ignored for now
    } else if (isElement(child)) {
      for (const sub of children(child)) {
        if (sub.name === "w:r") out.push(...runToInlines(sub, ctx));
      }
    }
  }
  // If a field was unterminated, flush what we have so it isn't lost.
  flushField();
  const inlines = mergeAdjacentText(out);
  if (!ctx.floats) return inlines;
  // The marks only say "a box sat here"; the boxes are block-level and were
  // collected into `ctx.floats` in source order.
  return inlines.filter((n) => n !== FLOAT_SLOT_MARK);
}

function fieldKind(instr: string): string {
  const m = instr.trim().match(/^([A-Za-z]+)/);
  return m && m[1] ? m[1].toUpperCase() : "FIELD";
}

function placeholderFor(instr: string): string {
  const k = fieldKind(instr);
  if (k === "PAGE") return "1";
  if (k === "NUMPAGES") return "1";
  if (k === "DATE") return "YYYY/MM/DD";
  if (k === "AUTHOR") return "(作者)";
  if (k === "FILENAME") return "(文件名)";
  return `{${k}}`;
}

function runToInlines(r: Element, ctx: InlineCtx = {}): PMNode[] {
  const rPr = firstChild(r, "w:rPr");
  const marks: PMMark[] = [];
  const textStyle: NonNullable<Extract<PMMark, { type: "textStyle" }>["attrs"]> = {};
  if (rPr) {
    if (flagOn(firstChild(rPr, "w:b"))) marks.push({ type: "bold" });
    if (flagOn(firstChild(rPr, "w:i"))) marks.push({ type: "italic" });
    if (!ctx.inLink && flagOn(firstChild(rPr, "w:u"))) marks.push({ type: "underline" });
    if (flagOn(firstChild(rPr, "w:strike"))) marks.push({ type: "strike" });
    const rFonts = firstChild(rPr, "w:rFonts");
    if (rFonts) {
      const attrs = rFonts.attributes ?? {};
      const eastAsia = attrs["w:eastAsia"] as string | undefined;
      const west =
        (attrs["w:ascii"] as string | undefined) ||
        (attrs["w:hAnsi"] as string | undefined);
      textStyle.fontFamily = eastAsia || west || null;
      // Two different faces really are in use: 中文用宋体、西文用 Times New Roman
      // is the standard pairing in every template we ship.
      if (eastAsia && west && eastAsia !== west) textStyle.fontFamilyWest = west;
    }
    const sz = firstChild(rPr, "w:sz");
    if (sz) {
      const half = Number(val(sz));
      if (half && !Number.isNaN(half)) textStyle.fontSize = ptToCssFontSize(half / 2);
    }
    const spacing = firstChild(rPr, "w:spacing");
    if (spacing) {
      const twips = Number(val(spacing));
      if (Number.isFinite(twips) && twips !== 0) {
        textStyle.letterSpacing = ptToCssLetterSpacing(twips / 20);
      }
    }
    const color = firstChild(rPr, "w:color");
    if (color && !ctx.inLink) {
      const c = canonColor(colorVal(val(color)));
      if (c) textStyle.color = c;
    }
    const highlight = firstChild(rPr, "w:highlight");
    if (highlight) {
      const v = val(highlight);
      const c = v ? canonColor(WORD_HIGHLIGHT[v] ?? colorVal(v)) : null;
      if (c) textStyle.backgroundColor = c;
    }
    const shd = firstChild(rPr, "w:shd");
    if (shd && !textStyle.backgroundColor) {
      const c = canonColor(colorVal(shd.attributes?.["w:fill"] as string | undefined));
      if (c) textStyle.backgroundColor = c;
    }
  }
  if (
    textStyle.fontFamily != null ||
    textStyle.fontFamilyWest != null ||
    textStyle.fontSize != null ||
    textStyle.letterSpacing != null ||
    textStyle.color != null ||
    textStyle.backgroundColor != null
  ) {
    marks.push({ type: "textStyle", attrs: textStyle });
  }

  const out: PMNode[] = [];
  for (const child of children(r)) {
    switch (child.name) {
      case "w:t": {
        const text = textOf(child);
        if (text) out.push({ type: "text", text, ...(marks.length ? { marks: [...marks] } : {}) });
        break;
      }
      case "w:tab":
        out.push({ type: "text", text: "\t", ...(marks.length ? { marks: [...marks] } : {}) });
        break;
      case "w:br": {
        const brType = child.attributes?.["w:type"] as string | undefined;
        if (brType === "page") out.push({ type: "pageBreak" } as PMNode);
        else out.push({ type: "hardBreak" });
        break;
      }
      case "w:drawing":
      case "mc:AlternateContent":
      case "w:pict": {
        const group = floatFromElement(child, ctx);
        if (group) {
          if (ctx.floats) {
            // Runs build their own inline array, so the position inside it is
            // meaningless one level up: leave a mark and let the paragraph
            // resolve where the box actually sat.
            ctx.floats.push(group);
            out.push(FLOAT_SLOT_MARK);
          } else {
            out.push(group);
          }
          break;
        }
        if (hasPicture(child)) {
          const drawing =
            child.name === "w:drawing"
              ? child
              : collectDescendants(child, "w:drawing")[0];
          const img = drawing ? imageFromDrawing(drawing) : null;
          if (img && img.type === "image") {
            // A picture wrapped in `mc:AlternateContent` has to keep the whole
            // branch element, or the legacy VML branch it pairs with is lost.
            if (child.name !== "w:drawing") img.attrs.ooxml = serialize(child);
            out.push(img);
          }
        }
        // Charts and SmartArt have no model here and are dropped.
        break;
      }
      default:
        break;
    }
  }
  return out;
}

function imageFromDrawing(drawing: Element): PMNode | null {
  const blips = collectDescendants(drawing, "a:blip");
  const blip = blips[0];
  const embed = blip?.attributes?.["r:embed"] as string | undefined;
  const inlineEl = firstChild(drawing, "wp:inline");
  const anchorEl = firstChild(drawing, "wp:anchor");
  const host = inlineEl ?? anchorEl;
  const ext = host ? collectDescendants(host, "a:ext")[0] : undefined;
  const cx = ext?.attributes?.cx ? Number(ext.attributes.cx) : undefined;
  const cy = ext?.attributes?.cy ? Number(ext.attributes.cy) : undefined;
  if (!embed) return null;
  const docPr = host ? collectDescendants(host, "wp:docPr")[0] : undefined;
  const name = (docPr?.attributes?.name as string | undefined) ?? null;
  const descr = (docPr?.attributes?.descr as string | undefined) ?? null;

  let posX: number | null = null;
  let posY: number | null = null;
  let wrap: "none" | "square" | "tight" | "through" | "topAndBottom" = "none";
  if (anchorEl) {
    const pH = firstChild(anchorEl, "wp:positionH");
    const pV = firstChild(anchorEl, "wp:positionV");
    const xOff = pH ? firstChild(pH, "wp:posOffset") : undefined;
    const yOff = pV ? firstChild(pV, "wp:posOffset") : undefined;
    // EMU → CSS px (96 dpi): 9525 EMU per px, 914400 per inch.
    if (xOff) posX = Math.round(Number(textOf(xOff)) / 9525);
    if (yOff) posY = Math.round(Number(textOf(yOff)) / 9525);
    if (firstChild(anchorEl, "wp:wrapSquare")) wrap = "square";
    else if (firstChild(anchorEl, "wp:wrapTight")) wrap = "tight";
    else if (firstChild(anchorEl, "wp:wrapThrough")) wrap = "through";
    else if (firstChild(anchorEl, "wp:wrapTopAndBottom")) wrap = "topAndBottom";
    else wrap = "none";
  }

  return {
    type: "image",
    attrs: {
      src: `relationship:${embed}`,
      alt: descr,
      title: name,
      width: cx ? Math.round((cx / 914400) * 96) : null,
      height: cy ? Math.round((cy / 914400) * 96) : null,
      ooxml: serialize(drawing),
      layout: anchorEl ? "anchor" : "inline",
      name,
      posX,
      posY,
      wrap,
    },
  };
}

/**
 * A floating drawing that carries shapes — text boxes, and the decorative
 * fills they sit on. Returns a `floatGroup` positioned at the anchor's
 * absolute coordinates, or null when there is nothing to model (a bare
 * picture, a chart), in which case the caller falls back to the image path.
 */
function floatFromElement(el: Element, ctx: InlineCtx): PMNode | null {
  const anchor = collectDescendants(el, "wp:anchor")[0];
  const boxes: PMNode[] = [];
  let x = 0;
  let y = 0;
  let relH = "paragraph";
  let relV = "paragraph";
  let z = 0;

  if (anchor) {
    const geom = anchorGeom(anchor);
    x = geom.x;
    y = geom.y;
    relH = geom.relH;
    relV = geom.relV;
    z = geom.zOrder;
    for (const data of collectDescendants(anchor, "a:graphicData")) {
      for (const shape of shapesIn(data, { w: geom.w, h: geom.h })) {
        boxes.push(floatBoxNode(shape, ctx));
      }
    }
  } else if (el.name === "w:pict") {
    // Pre-2010 form: positions live in the VML `style` string.
    for (const shape of shapesFromPict(el)) boxes.push(floatBoxNode(shape, ctx));
  } else {
    return null;
  }

  if (!boxes.length) return null;
  return {
    type: "floatGroup",
    attrs: { ooxml: serialize(el), x, y, relH, relV, z },
    content: boxes,
  };
}

function floatBoxNode(shape: ShapeInfo, ctx: InlineCtx): PMNode {
  const content = shape.txbxContentEl
    ? children(shape.txbxContentEl).filter((c) => c.name === "w:p")
    : [];
  return {
    type: "floatBox",
    attrs: {
      x: shape.x,
      y: shape.y,
      w: shape.w,
      h: shape.h,
      fill: shape.fill,
      stroke: shape.stroke,
      strokeWidth: shape.strokeWidth,
      vAlign: shape.vAlign,
      name: shape.name,
      // The source body XML is kept so an untouched box writes back exactly
      // what Word put there, list definitions and run properties included.
      txbx: shape.txbxContentEl
        ? children(shape.txbxContentEl).map((c) => serialize(c)).join("")
        : null,
      dirty: false,
    },
    content: content.length
      ? content.map((c) => paragraphFromXml(c, ctx.numbering, ctx.styles))
      : [{ type: "paragraph" }],
  };
}

function tableFromXml(
  tbl: Element,
  numbering?: NumberingMap,
  styles?: StyleMap,
): PMNode {
  const rows: PMNode[] = [];
  for (const tr of children(tbl).filter((c) => c.name === "w:tr")) {
    const cells: PMNode[] = [];
    for (const tc of children(tr).filter((c) => c.name === "w:tc")) {
      const paras: PMNode[] = [];
      for (const child of children(tc)) {
        if (child.name === "w:p") {
          const p = paragraphFromXml(child, numbering, styles);
          paras.push(p);
        }
      }
      const tcPr = firstChild(tc, "w:tcPr");
      const gridSpan = tcPr && firstChild(tcPr, "w:gridSpan");
      const vMerge = tcPr && firstChild(tcPr, "w:vMerge");
      const isHeader =
        // We can't tell a header row without tblPr/header-to-row mapping.
        // Cells that contain bold-only paragraphs aren't headers; default false.
        false;
      cells.push({
        type: isHeader ? "tableHeader" : "tableCell",
        attrs: {
          ...(gridSpan && val(gridSpan) ? { colspan: Number(val(gridSpan)) } : {}),
          ...(vMerge && val(vMerge) !== "restart" ? { rowspan: 1 } : {}),
        },
        content: paras.length ? paras : [{ type: "paragraph" }],
      });
    }
    rows.push({ type: "tableRow", content: cells });
  }
  return { type: "table", content: rows };
}

function unmanaged(el: Element, kind: string): PMNode {
  return {
    type: "unmanagedBlock",
    attrs: {
      kind,
      ooxml: serialize(el),
      label: kind === "section" ? "（页面设置）" : `（不支持的元素：${kind}）`,
    },
  };
}

// ---------- PM → OOXML ----------

function blockToXml(node: PMNode, ctx: WriteCtx = {}): string {
  switch (node.type) {
    case "paragraph":
      return paragraphXml(node, { link: ctx.link });
    case "heading":
      return headingXml(node, ctx);
    case "blockquote":
      return blockquoteXml(node, ctx);
    case "bulletList":
    case "orderedList":
      return listXml(node, ctx);
    case "listItem":
      return listItemXml(node, "bulletList", 0, undefined, ctx);
    case "codeBlock":
      return codeBlockXml(node);
    case "horizontalRule":
      return `<w:p><w:pPr><w:pStyle w:val="HorizontalRule"/></w:pPr></w:p>`;
    case "table":
      return tableXml(node, ctx);
    case "image":
      return `<w:p><w:r>${imageXml(node)}</w:r></w:p>`;
    case "floatGroup":
      return floatGroupXml(node, ctx);
    case "pageBreak":
      // A standalone pageBreak node becomes its own paragraph for safety.
      return `<w:p><w:r><w:br w:type="page"/></w:r></w:p>`;
    case "field":
      // A field that ended up at block level writes its run sequence
      // directly, without wrapping it in another <w:r>.
      return `<w:p>${fieldXml(node)}</w:p>`;
    case "unmanagedBlock":
      return node.attrs.ooxml;
    case "text":
    case "hardBreak":
      return `<w:p>${runXml([node], ctx.link)}</w:p>`;
    default:
      return "";
  }
}

function paragraphXml(
  node: Extract<PMNode, { type: "paragraph" }> | {
    content?: PMNode[];
    attrs?: PPrAttrs;
  },
  opts: {
    pStyle?: string;
    rawPpr?: string;
    numPr?: string;
    link?: LinkStyle | null;
  } = {},
): string {
  const a = node.attrs ?? {};
  const pPr = paragraphPprXml(a, {
    pStyle: opts.pStyle,
    numPr: opts.numPr,
    rawPpr: opts.rawPpr ?? a.rawPpr,
  });
  const runs = inlineChildrenToXml(node.content ?? [], opts.link);
  return `<w:p>${pPr}${runs}</w:p>`;
}

/**
 * CT_PPr 子元素的合法顺序（ECMA-376 §17.3.1.1）。Word 校验的是**先后**而不只是
 * 元素本身，顺序错乱会弹"文档内容有问题"，所以透传的 rawPpr 与 attrs 生成的元素
 * 都走这一份顺序表，不再由两条路径各自拼接。表里没有的名字排到最后，
 * 维持它们在源文档里的相对次序。
 */
const PPR_CHILD_ORDER = [
  "w:pStyle",
  "w:keepNext",
  "w:keepLines",
  "w:pageBreakBefore",
  "w:framePr",
  "w:widowControl",
  "w:numPr",
  "w:suppressLineNumbers",
  "w:pBdr",
  "w:shd",
  "w:tabs",
  "w:suppressAutoHyphens",
  "w:kinsoku",
  "w:wordWrap",
  "w:overflowPunct",
  "w:topLinePunct",
  "w:autoSpaceDE",
  "w:autoSpaceDN",
  "w:bidi",
  "w:adjustRightInd",
  "w:snapToGrid",
  "w:spacing",
  "w:ind",
  "w:contextualSpacing",
  "w:mirrorIndents",
  "w:suppressOverlap",
  "w:jc",
  "w:textDirection",
  "w:textAlignment",
  "w:textboxTightWrap",
  "w:outlineLvl",
  "w:divId",
  "w:cnfStyle",
  "w:rPr",
  "w:sectPr",
  "w:pPrChange",
];

/** `<w:pPr>` 的一个子元素：名字用于排序，xml 是最终写出的文本。 */
interface PprPart {
  name: string;
  xml: string;
}

/**
 * 按 CT_PPr 顺序稳定排序后拼出 `<w:pPr>`。
 *
 * @param parts 子元素列表，允许乱序；同名元素的相对次序保持传入顺序。
 * @returns 完整的 pPr 字符串；一个子元素都没有时返回空串（段落直接不带 pPr）。
 */
function pprFromParts(parts: PprPart[]): string {
  const rankOf = (name: string): number => {
    const i = PPR_CHILD_ORDER.indexOf(name);
    return i < 0 ? PPR_CHILD_ORDER.length : i;
  };
  const inner = parts
    .map((part, index) => ({ index, rank: rankOf(part.name), xml: part.xml }))
    .sort((a, b) => a.rank - b.rank || a.index - b.index)
    .map((part) => part.xml)
    .join("");
  return inner ? `<w:pPr>${inner}</w:pPr>` : "";
}

/**
 * 段落 attrs 里建模过的那些 pPr 子元素。
 *
 * 每一项同一条规则：attrs 没有值（undefined）就不产出元素，rawPpr 里的原文留着。
 *
 * @param a 段落 attrs（段前/缩进是 pt，lineHeight 是倍数或 pt:N / atLeast:N）。
 * @param rawPpr 源文档 pPr 原文，供 spacing/ind 做属性级合并；undefined 表示新建段落。
 * @returns 待排序写出的子元素。
 */
function modeledPprParts(a: PPrAttrs, rawPpr?: string): PprPart[] {
  const parts: PprPart[] = [];
  const push = (name: string, xml: string) => {
    if (xml) parts.push({ name, xml });
  };
  push("w:pageBreakBefore", onOffXml("w:pageBreakBefore", a.pageBreakBefore));
  push("w:keepNext", onOffXml("w:keepNext", a.keepWithNext));
  push("w:keepLines", onOffXml("w:keepLines", a.keepLines));
  push("w:widowControl", onOffXml("w:widowControl", a.widowControl));
  push("w:snapToGrid", onOffXml("w:snapToGrid", a.snapToGrid));
  push("w:spacing", spacingElementXml(a, rawPpr));
  push("w:ind", indElementXml(a, rawPpr));
  push("w:jc", jcXml(a.textAlign));
  const lvl = a.outlineLevel;
  if (typeof lvl === "number" && Number.isInteger(lvl) && lvl >= 0 && lvl <= 9) {
    push("w:outlineLvl", `<w:outlineLvl w:val="${lvl}"/>`);
  }
  return parts;
}

/**
 * 生成段落的 `<w:pPr>`，两条路径（有 rawPpr 的导入段落 / 编辑器新建段落）共用。
 *
 * 合并规则：attrs 建模过的元素顶替 rawPpr 里的同名元素——spacing/ind 内部还要
 * 逐个属性合并（见 `mergedElementXml`），因为源文档里可能有我们没建模的写法
 * （`w:*Chars`、`w:*Lines`、`w:start`/`w:end`）。其余元素原样保留并按 CT_PPr
 * 顺序排回该在的位置，源文档带 `<w:rPr>`（段落标记的字符属性，Word 写得非常多）
 * 时不会再被挤到 spacing 前面。
 *
 * @param a 段落 attrs。
 * @param opts.pStyle 段落要套用的样式 id（标题、引用）；源文档已有 pStyle 时以源文档为准。
 * @param opts.numPr 列表段落重新分组后由我们给出的 numPr。
 * @param opts.rawPpr 源文档的 pPr 内部 XML。
 * @returns 完整 pPr 字符串；段落没有任何格式属性时返回空串。
 */
function paragraphPprXml(
  a: PPrAttrs,
  opts: { pStyle?: string; numPr?: string; rawPpr?: string } = {},
): string {
  const modeled = modeledPprParts(a, opts.rawPpr);
  const covered = new Set(modeled.map((part) => part.name));
  const rawParts = (opts.rawPpr ? parseFragment(opts.rawPpr) : [])
    .map((el) => ({ name: el.name, xml: serialize(el) }))
    .filter((part) => !covered.has(part.name));
  // 样式引用以源文档为准：同一份 pPr 里出现两个 w:pStyle 时 Word 取第一个，
  // 套用哪个样式就会变得不可预测，所以原文有 pStyle 就不再补我们推断的那个。
  const keepOwnStyle =
    !!opts.pStyle && !rawParts.some((part) => part.name === "w:pStyle");
  const parts: PprPart[] = [];
  if (keepOwnStyle) {
    parts.push({
      name: "w:pStyle",
      xml: `<w:pStyle w:val="${opts.pStyle}"/>`,
    });
  }
  parts.push(...rawParts);
  if (opts.numPr) parts.push({ name: "w:numPr", xml: opts.numPr });
  parts.push(...modeled);
  return pprFromParts(parts);
}

/** textAlign（left/center/right/justify）→ `<w:jc>`；两端对齐在 OOXML 里叫 both。 */
function jcXml(textAlign: string | undefined): string {
  if (!textAlign) return "";
  const v =
    textAlign === "center"
      ? "center"
      : textAlign === "right"
        ? "right"
        : textAlign === "justify"
          ? "both"
          : "left";
  return `<w:jc w:val="${v}"/>`;
}

/** pt → OOXML 的 twips 字符串（1pt = 20 twips）。 */
function twipsValue(pt: number): string {
  return String(Math.round(pt * 20));
}

/**
 * 取 rawPpr 里某个 pPr 子元素（传 "ind" / "spacing"）已有的属性表。
 *
 * 存在的意义：写回做的是**属性级合并**而不是整元素重建。我们没建模的写法
 * （`w:right`、`w:start`/`w:end`、字符单位版 `w:*Chars`、按行版的
 * `w:beforeLines`）以前会被 `cleanRaw` 连元素一起删掉，保存一次就永久丢失。
 */
function rawElementAttrs(
  rawPpr: string | undefined,
  tag: string,
): Record<string, string> {
  if (!rawPpr) return {};
  const match = new RegExp(`<w:${tag}\\b([^>]*?)/?>`, "").exec(rawPpr);
  if (!match) return {};
  const attrs: Record<string, string> = {};
  const inner = match[1] ?? "";
  for (const pair of inner.matchAll(/(w:[\w-]+)="([^"]*)"/g)) {
    const name = pair[1];
    const value = pair[2];
    if (name && value) attrs[name] = value;
  }
  return attrs;
}

/**
 * 合并出 `<w:spacing>` / `<w:ind>`：源文档属性打底，attrs 里建模过的覆盖上去。
 * `suppress` 处理 Word 的优先级陷阱——`w:*Chars`、`w:*Lines`、`w:beforeAuto`
 * 会压过同义的 twips 值，2010 起的 `w:start`/`w:end` 与 `w:left`/`w:right` 互斥，
 * 所以接管一个属性时必须把它的一切同义写法清掉，否则用户在面板里改了不生效。
 *
 * 只在"这一次真的改了值"时才清理同义写法：源文档里两者本来就一致（编辑器里没
 * 动过这段）时把它们留着，重复保存才不会把中文版 Word 写的 `w:leftChars`
 * 一步步磨掉，也才谈得上幂等。
 */
function mergedElementXml(
  tag: string,
  order: string[],
  rawPpr: string | undefined,
  set: Record<string, string>,
  suppress: Record<string, string[]>,
): string {
  const raw = rawElementAttrs(rawPpr, tag);
  const kept: Record<string, string> = { ...raw };
  for (const [name, value] of Object.entries(set)) {
    const changed = raw[name] === undefined || raw[name] !== value;
    kept[name] = value;
    if (changed) for (const drop of suppress[name] ?? []) delete kept[drop];
  }
  const names = [
    ...order.filter((n) => kept[n] !== undefined),
    ...Object.keys(kept).filter((n) => !order.includes(n) && kept[n] !== undefined),
  ];
  if (!names.length) return "";
  return `<w:${tag} ${names.map((n) => `${n}="${kept[n]}"`).join(" ")}/>`;
}

const IND_ORDER = ["w:left", "w:right", "w:firstLine", "w:hanging"];
const IND_SUPPRESS: Record<string, string[]> = {
  "w:left": ["w:start", "w:leftChars"],
  "w:right": ["w:end", "w:rightChars"],
  "w:firstLine": ["w:hanging", "w:firstLineChars"],
  "w:hanging": ["w:firstLine", "w:hangingChars"],
};

const SPACING_ORDER = ["w:before", "w:after", "w:line", "w:lineRule"];
const SPACING_SUPPRESS: Record<string, string[]> = {
  "w:before": ["w:beforeLines", "w:beforeAuto"],
  "w:after": ["w:afterLines", "w:afterAuto"],
  "w:line": ["w:lineRule"],
};

/**
 * 由段落 attrs 生成 `<w:ind>`。attrs 单位是 pt；负的 `indentFirst` 表示
 * 悬挂缩进，写成 `w:hanging`（Word 的"特殊格式：悬挂缩进"就是这个属性）。
 * 传 undefined 的 attrs 项保持源文档原样。
 */
function indElementXml(
  a: PPrAttrs,
  rawPpr?: string,
): string {
  const set: Record<string, string> = {};
  if (typeof a.indent === "number") set["w:left"] = twipsValue(a.indent);
  if (typeof a.indentRight === "number")
    set["w:right"] = twipsValue(a.indentRight);
  if (typeof a.indentFirst === "number") {
    const tw = twipsValue(Math.abs(a.indentFirst));
    if (a.indentFirst >= 0) set["w:firstLine"] = tw;
    else set["w:hanging"] = tw;
  }
  return mergedElementXml("ind", IND_ORDER, rawPpr, set, IND_SUPPRESS);
}

/**
 * 由段落 attrs 生成 `<w:spacing>`。段前/段后是 pt，行距沿用
 * `lineHeightSpacingAttrs`（倍数走 `w:lineRule="auto"`，"pt:N" 是固定值、
 * "atLeast:N" 是最小值，单位都是 twips）。
 */
function spacingElementXml(
  a: PPrAttrs,
  rawPpr?: string,
): string {
  const set: Record<string, string> = {};
  if (typeof a.spacingBefore === "number")
    set["w:before"] = twipsValue(a.spacingBefore);
  if (typeof a.spacingAfter === "number")
    set["w:after"] = twipsValue(a.spacingAfter);
  // 行距先给 w:line 再给 w:lineRule，顺序决定源文档里那条旧的 lineRule 会被覆盖。
  for (const pair of lineHeightSpacingAttrs(a.lineHeight as string | undefined)) {
    const m = /^(w:[\w-]+)="([^"]*)"$/.exec(pair);
    const name = m?.[1];
    const value = m?.[2];
    if (name && value) set[name] = value;
  }
  return mergedElementXml(
    "spacing",
    SPACING_ORDER,
    rawPpr,
    set,
    SPACING_SUPPRESS,
  );
}

function headingXml(
  node: Extract<PMNode, { type: "heading" }>,
  ctx: WriteCtx = {},
): string {
  return paragraphXml(node, {
    pStyle: `Heading${node.attrs.level}`,
    link: ctx.link,
  });
}

function blockquoteXml(
  node: Extract<PMNode, { type: "blockquote" }>,
  ctx: WriteCtx = {},
): string {
  return (node.content ?? [])
    .map((child) => {
      if (child.type === "paragraph") {
        return paragraphXml(child, { pStyle: "Quote", link: ctx.link });
      }
      return blockToXml(child, ctx);
    })
    .join("");
}

function listXml(
  node: Extract<PMNode, { type: "bulletList" | "orderedList" }>,
  ctx: WriteCtx = {},
  level = 0,
): string {
  const ordered = node.type === "orderedList";
  // Prefer the numId that came from the source document; otherwise ask the
  // numbering map for a default of the right kind.
  const existingNumId = node.attrs?.numId;
  let numId: string | undefined = existingNumId;
  if (!numId && ctx.numbering) {
    const found = ctx.numbering.ensureFor(ordered ? "ordered" : "bullet");
    if (found) numId = found.numId;
  }
  if (numId && ctx.numbering) ctx.numbering.markEmitted(numId);
  return (node.content ?? [])
    .map((item, i) =>
      listItemXml(
        item,
        ordered ? "orderedList" : "bulletList",
        i,
        numId,
        ctx,
        level,
      ),
    )
    .join("");
}

function listItemXml(
  item: PMNode,
  kind: "bulletList" | "orderedList",
  index: number,
  numId: string | undefined,
  ctx: WriteCtx = {},
  level = 0,
): string {
  if (item.type !== "listItem") return blockToXml(item, ctx);
  const ilvl = Math.max(0, Math.min(8, level));
  // If we have a real numId from the source / numbering map, emit proper
  // numPr; otherwise fall back to pStyle-based bullet/numbered paragraphs.
  const numPr = numId
    ? `<w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${numId}"/></w:numPr>`
    : "";
  const pStyleFallback = numId
    ? ""
    : `<w:pStyle w:val="${kind === "orderedList" ? "ListNumber" : "ListBullet"}"/>`;
  const blocks = (item.content ?? []).map((child) => {
    if (child.type === "paragraph") {
      // If the list item paragraph came from a source doc with its own pPr
      // (spacing/indent/etc.), preserve it. We prepend our synthesized
      // numPr so it takes precedence over any stale numPr from the source.
      const childRaw = (child.attrs as { rawPpr?: string } | undefined)?.rawPpr;
      if (childRaw) {
        // Strip the original numPr (and any ilvl) to avoid duplicates and
        // ensure our synthesized one with the correct depth wins.
        const stripped = childRaw
          .replace(/<w:numPr>[\s\S]*?<\/w:numPr>/g, "")
          .replace(/<w:ilvl\b[^>]*\/>/g, "");
        return paragraphXml(child, {
          rawPpr: `${numPr}${pStyleFallback}${stripped}`,
          link: ctx.link,
        });
      }
      return paragraphXml(child, {
        rawPpr: `${pStyleFallback}${numPr}`,
        link: ctx.link,
      });
    }
    if (child.type === "bulletList" || child.type === "orderedList") {
      return listXml(child, ctx, level + 1);
    }
    return blockToXml(child, ctx);
  });
  void index;
  return blocks.join("");
}

function codeBlockXml(node: Extract<PMNode, { type: "codeBlock" }>): string {
  const text = (node.content ?? [])
    .map((c) => (c.type === "text" ? c.text ?? "" : ""))
    .join("");
  return `<w:p><w:pPr><w:pStyle w:val="Code"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/><w:sz w:val="20"/></w:rPr><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`;
}

function tableXml(
  node: Extract<PMNode, { type: "table" }>,
  ctx: WriteCtx = {},
): string {
  const rows = (node.content ?? [])
    .filter((c): c is Extract<PMNode, { type: "tableRow" }> => c.type === "tableRow")
    .map((row) => {
      const cells = (row.content ?? [])
        .filter(
          (c): c is Extract<PMNode, { type: "tableCell" | "tableHeader" }> =>
            c.type === "tableCell" || c.type === "tableHeader",
        )
        .map((cell) => tableCellXml(cell, ctx))
        .join("");
      return `<w:tr>${cells}</w:tr>`;
    })
    .join("");
  const tblPr =
    `<w:tblPr><w:tblW w:w="0" w:type="auto"/>` +
    `<w:tblBorders>` +
    `<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>` +
    `<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>` +
    `<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>` +
    `<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>` +
    `<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>` +
    `<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>` +
    `</w:tblBorders></w:tblPr>`;
  return `<w:tbl>${tblPr}${rows}</w:tbl>`;
}

function tableCellXml(
  cell: Extract<PMNode, { type: "tableCell" | "tableHeader" }>,
  ctx: WriteCtx = {},
): string {
  const tcPr: string[] = [];
  if (cell.attrs?.colspan && cell.attrs.colspan > 1) {
    tcPr.push(`<w:gridSpan w:val="${cell.attrs.colspan}"/>`);
  }
  const blocks = (cell.content ?? [])
    .map((c) => {
      if (c.type === "paragraph") {
        if (cell.type === "tableHeader") {
          const withBold: PMNode = {
            ...c,
            content: (c.content ?? []).map((inline) =>
              inline.type === "text"
                ? { ...inline, marks: [...(inline.marks ?? []), { type: "bold" as const }] }
                : inline,
            ),
          };
          return paragraphXml(withBold, { link: ctx.link });
        }
        return paragraphXml(c, { link: ctx.link });
      }
      return blockToXml(c, ctx);
    })
    .join("");
  return `<w:tc>${tcPr.length ? `<w:tcPr>${tcPr.join("")}</w:tcPr>` : ""}${blocks || "<w:p/>"}</w:tc>`;
}

function inlineChildrenToXml(
  nodes: PMNode[],
  link?: LinkStyle | null,
): string {
  // Group into runs by marks so consecutive text with identical marks becomes
  // a single <w:r>.
  const groups = groupByMarks(nodes);
  return groups.map((g) => runXml(g, link)).join("");
}

function groupByMarks(nodes: PMNode[]): PMNode[][] {
  const groups: PMNode[][] = [];
  let current: PMNode[] = [];
  let key = "";
  const flush = () => {
    if (current.length) groups.push(current);
    current = [];
  };
  for (const n of nodes) {
    if (n.type === "hardBreak") {
      current.push(n);
      continue;
    }
    if (n.type === "image" || n.type === "field" || n.type === "pageBreak") {
      flush();
      groups.push([n]);
      key = "";
      continue;
    }
    if (n.type !== "text") continue;
    const k = JSON.stringify(n.marks ?? []);
    if (k !== key) {
      flush();
      key = k;
    }
    current.push(n);
  }
  flush();
  return groups;
}

function runXml(nodes: PMNode[], link?: LinkStyle | null): string {
  // Determine marks from first text node.
  const first = nodes.find((n) => n.type === "text") as PMTextNode | undefined;
  const marks = first?.marks ?? [];
  const rPr = runPropertiesXml(marks, link);
  const linkMark = marks.find((m): m is Extract<PMMark, { type: "link" }> => m.type === "link");

  // Text parts and structural children (br, drawing) must live directly under
  // <w:r>, not inside <w:t>. We split text around structural nodes.
  const parts: string[] = [];
  let buffer = "";
  const flushText = () => {
    if (buffer.length) {
      // Escape newlines as a <w:br/> inside the run.
      const escaped = escapeXml(buffer).replace(
        /\n/g,
        `</w:t><w:br/><w:t xml:space="preserve">`,
      );
      parts.push(`<w:t xml:space="preserve">${escaped}</w:t>`);
      buffer = "";
    }
  };
  for (const n of nodes) {
    if (n.type === "text") {
      buffer += n.text ?? "";
    } else if (n.type === "hardBreak") {
      flushText();
      parts.push("<w:br/>");
    } else if (n.type === "pageBreak") {
      flushText();
      parts.push(`<w:br w:type="page"/>`);
    } else if (n.type === "image") {
      flushText();
      parts.push(imageXml(n));
    } else if (n.type === "field") {
      flushText();
      parts.push(fieldXml(n));
    }
  }
  flushText();

  // A standalone field (no marks, no text around it) serializes to
  // several runs directly; wrapping them in another <w:r> would nest
  // runs and break field parsing in Word/WPS.
  if (
    nodes.length === 1 &&
    nodes[0]!.type === "field" &&
    !rPr &&
    !linkMark
  ) {
    return parts.join("");
  }

  const wrapped = rPr ? `<w:r>${rPr}${parts.join("")}</w:r>` : `<w:r>${parts.join("")}</w:r>`;
  if (linkMark) {
    const href = linkMark.attrs.href;
    if (href.startsWith("relationship:")) {
      return `<w:hyperlink r:id="${href.slice("relationship:".length)}">${wrapped}</w:hyperlink>`;
    }
    return `<w:hyperlink w:anchor="${escapeXml(href)}">${wrapped}</w:hyperlink>`;
  }
  return wrapped;
}

// Word highlight name → CSS color
const WORD_HIGHLIGHT: Record<string, string> = {
  yellow: "#ffff00",
  green: "#00ff00",
  cyan: "#00ffff",
  magenta: "#ff00ff",
  blue: "#0000ff",
  red: "#ff0000",
  darkBlue: "#00008b",
  darkCyan: "#008b8b",
  darkGreen: "#006400",
  darkMagenta: "#8b008b",
  darkRed: "#8b0000",
  darkYellow: "#808000",
  darkGray: "#a9a9a9",
  lightGray: "#d3d3d3",
  black: "#000000",
  white: "#ffffff",
};
const HIGHLIGHT_FROM_CSS: Record<string, string> = Object.fromEntries(
  Object.entries(WORD_HIGHLIGHT).map(([k, v]) => [v, k]),
);

/**
 * CT_RPr's element order. Word accepts any order but always writes this one,
 * and emitting a fixed order is what makes "the same marks produce the same
 * bytes" true — the property the template re-apply check depends on.
 */
const RPR_ORDER = [
  "rFonts",
  "b",
  "i",
  "strike",
  "color",
  "spacing",
  "sz",
  "highlight",
  "u",
  "shd",
] as const;
type RPrKey = (typeof RPR_ORDER)[number];

/** A font name as an XML attribute value. */
function fontAttr(name: string): string {
  return escapeXml(name).replace(/"/g, "");
}

function runPropertiesXml(marks: PMMark[], link?: LinkStyle | null): string {
  // One writer per element name. A duplicated `<w:color>` is not merely
  // untidy: Word honours the first and ignores the rest, which is how a link
  // that also carried the user's own colour used to lose it.
  const el: Partial<Record<RPrKey, string>> = {};
  const styleMark = marks.find(
    (m): m is Extract<PMMark, { type: "textStyle" }> => m.type === "textStyle",
  );
  const ownColor = styleMark?.attrs.color;
  const ownUnderline = marks.some((m) => m.type === "underline");
  // A link's colour and rule are the document-level setting, and only when the
  // run has none of its own.
  const linkColor = cssColorToOoxmlHex(link?.color ?? "") ?? "0563C1";
  const linkUnderline = link?.underline !== false;

  if (marks.some((m) => m.type === "bold")) el.b = "<w:b/>";
  if (marks.some((m) => m.type === "italic")) el.i = "<w:i/>";
  if (ownUnderline) el.u = `<w:u w:val="single"/>`;
  if (marks.some((m) => m.type === "strike")) el.strike = "<w:strike/>";
  if (marks.some((m) => m.type === "code")) {
    el.rFonts = `<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/>`;
  }
  if (marks.some((m) => m.type === "link")) {
    if (!ownColor) el.color = `<w:color w:val="${linkColor}"/>`;
    if (!ownUnderline && linkUnderline) el.u = `<w:u w:val="single"/>`;
  }

  if (styleMark) {
    const { fontFamily, fontFamilyWest, fontSize, letterSpacing, color, backgroundColor } =
      styleMark.attrs;
    const east = fontFamily ? fontAttr(String(fontFamily)) : null;
    const west = fontFamilyWest ? fontAttr(String(fontFamilyWest)) : east;
    if (east || west) {
      // 中文与西文分字体：eastAsia 只作用于汉字，ascii/hAnsi/cs 作用于拉丁
      // 字母。只有一个名字时两者相同，等于旧的整体设置行为。
      const attrs = [
        west ? `w:ascii="${west}" w:hAnsi="${west}" w:cs="${west}"` : "",
        east ? `w:eastAsia="${east}"` : "",
      ]
        .filter((a) => a !== "")
        .join(" ");
      el.rFonts = `<w:rFonts ${attrs}/>`;
    }
    if (fontSize) {
      const pt = cssFontSizeToPt(String(fontSize));
      const n = pt === null ? 0 : Math.round(pt * 2);
      if (n > 0) el.sz = `<w:sz w:val="${n}"/><w:szCs w:val="${n}"/>`;
    }
    if (letterSpacing) {
      // rPr `w:spacing` is character spacing in twentieths of a point, and it
      // may be negative (tightening) — only zero means "inherit".
      const pt = cssLetterSpacingToPt(String(letterSpacing));
      const twips = pt === null ? 0 : Math.round(pt * 20);
      if (twips !== 0) el.spacing = `<w:spacing w:val="${twips}"/>`;
    }
    if (color) {
      const hex = cssColorToOoxmlHex(color);
      if (hex) el.color = `<w:color w:val="${hex}"/>`;
    }
    if (backgroundColor) {
      // Word's 16 highlight names win when the colour is one of them; any
      // other colour goes through shading. Both accept only hex, so an
      // unreadable colour is dropped here rather than written broken.
      const raw = String(backgroundColor).trim().toLowerCase();
      const hex = cssColorToOoxmlHex(raw);
      const named = HIGHLIGHT_FROM_CSS[hex ? `#${hex.toLowerCase()}` : raw];
      if (named) {
        el.highlight = `<w:highlight w:val="${named}"/>`;
      } else if (hex) {
        el.shd = `<w:shd w:val="clear" w:color="auto" w:fill="${hex}"/>`;
      }
    }
  }

  const xml = RPR_ORDER.map((k) => el[k]).filter((v): v is string => Boolean(v)).join("");
  return xml ? `<w:rPr>${xml}</w:rPr>` : "";
}

function fieldXml(node: Extract<PMNode, { type: "field" }>): string {
  const instr = node.attrs.instr ?? node.attrs.kind ?? "PAGE";
  const display = node.attrs.display ?? placeholderFor(instr);
  return (
    `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
    `<w:r><w:instrText xml:space="preserve"> ${escapeXml(instr)} </w:instrText></w:r>` +
    `<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
    `<w:r><w:t xml:space="preserve">${escapeXml(display)}</w:t></w:r>` +
    `<w:r><w:fldChar w:fldCharType="end"/></w:r>`
  );
}

function imageXml(node: Extract<PMNode, { type: "image" }>): string {
  // If the image came from a source .docx and was NOT moved by the user,
  // keep its exact <w:drawing> XML so floating anchors (wp:anchor),
  // wrapping, and extents round-trip perfectly.
  if (node.attrs.ooxml && node.attrs.src.startsWith("relationship:")) {
    return node.attrs.ooxml;
  }
  const rId = node.attrs.src.startsWith("relationship:")
    ? node.attrs.src.slice("relationship:".length)
    : "rIdImg";
  const cx = node.attrs.width ? node.attrs.width * 9525 : 200 * 9525;
  const cy = node.attrs.height
    ? node.attrs.height * 9525
    : node.attrs.width
      ? node.attrs.width * 9525
      : 200 * 9525;
  const name = escapeXml(node.attrs.name ?? "Picture");

  // If the user dragged a floating image (layout=anchor) but the original
  // ooxml was cleared, synthesize a new <wp:anchor> at the new position.
  // posX/posY are CSS px → EMU (×9525). The picture is anchored to the
  // paragraph that contains it (relativeFrom=paragraph for both axes).
  if (node.attrs.layout === "anchor") {
    const xEmu = Math.round((node.attrs.posX ?? 0) * 9525);
    const yEmu = Math.round((node.attrs.posY ?? 0) * 9525);
    const wrap =
      node.attrs.wrap === "square"
        ? `<wp:wrapSquare wrapText="bothSides"/>`
        : node.attrs.wrap === "tight"
          ? `<wp:wrapTight wrapText="bothSides"><wp:wrapPolygon edited="0"><wp:start x="0" y="0"/><wp:lineTo x="${cx}" y="0"/><wp:lineTo x="${cx}" y="${cy}"/><wp:lineTo x="0" y="${cy}"/><wp:lineTo x="0" y="0"/></wp:wrapPolygon></wp:wrapTight>`
          : node.attrs.wrap === "topAndBottom"
            ? `<wp:wrapTopAndBottom/>`
            : `<wp:wrapNone/>`;
    return (
      `<w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
      `<wp:simplePos x="0" y="0"/>` +
      `<wp:positionH relativeFrom="column"><wp:posOffset>${xEmu}</wp:posOffset></wp:positionH>` +
      `<wp:positionV relativeFrom="paragraph"><wp:posOffset>${yEmu}</wp:posOffset></wp:positionV>` +
      `<wp:extent cx="${cx}" cy="${cy}"/>` +
      `<wp:effectExtent l="0" t="0" r="0" b="0"/>` +
      wrap +
      `<wp:docPr id="1" name="${name}"/>` +
      `<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>` +
      `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
      `<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
      `<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
      `<pic:nvPicPr><pic:cNvPr id="1" name="${name}"/><pic:cNvPicPr><a:picLocks noChangeAspect="1"/></pic:cNvPicPr></pic:nvPicPr>` +
      `<pic:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
      `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
      `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
      `</pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing>`
    );
  }

  return (
    `<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">` +
    `<wp:extent cx="${cx}" cy="${cy}"/>` +
    `<wp:effectExtent l="0" t="0" r="0" b="0"/>` +
    `<wp:docPr id="1" name="${name}"/>` +
    `<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>` +
    `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
    `<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
    `<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
    `<pic:nvPicPr><pic:cNvPr id="1" name="${name}"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
    `</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>`
  );
}

/**
 * Write a floating drawing back into the paragraph that anchors it: OOXML
 * requires a `wp:anchor` to live in a run. An untouched group comes back as the
 * exact element we read; once a box has been retyped, only that box's
 * `<w:txbxContent>` body is replaced inside the same template, so fills, group
 * transforms and the `mc:Fallback` branch are not rebuilt from scratch.
 */
function floatGroupXml(
  node: Extract<PMNode, { type: "floatGroup" }>,
  ctx: WriteCtx = {},
): string {
  const inner: string[] = [];
  for (const box of node.content ?? []) {
    if (box.type !== "floatBox") continue;
    // A decorative shape carries no text body, so it takes no slot in the
    // template — listing it would shift every following box by one.
    if (box.attrs.txbx === null) continue;
    inner.push(
      box.attrs.dirty || !box.attrs.txbx
        ? (box.content ?? []).map((b) => blockToXml(b, ctx)).join("")
        : box.attrs.txbx,
    );
  }
  const shape = spliceTxbxContents(node.attrs.ooxml, inner);
  return `<w:p><w:r><w:rPr><w:noProof/></w:rPr>${shape}</w:r></w:p>`;
}

// ---------- helpers ----------

interface ListMarker {
  ordered: boolean;
  level: number;
  numId?: string;
}

function isListItemMarker(node: PMNode): node is PMNode & { __list: ListMarker } {
  return (
    node.type === "listItem" &&
    typeof (node as unknown as { __list?: unknown }).__list === "object" &&
    (node as unknown as { __list: ListMarker }).__list !== null
  );
}

/**
 * Combine consecutive listItem paragraphs (emitted by paragraphFromXml) into
 * proper bulletList/orderedList nodes. Indentation level is represented by
 * nesting lists so ProseMirror/TipTap render it correctly.
 */
function groupListItems(nodes: PMNode[]): PMNode[] {
  const out: PMNode[] = [];
  let i = 0;
  while (i < nodes.length) {
    const node = nodes[i]!;
    if (!isListItemMarker(node)) {
      out.push(stripMarker(node));
      i += 1;
      continue;
    }
    // Gather a run of same-ordered list items.
    const ordered = node.__list.ordered;
    const run: PMNode[] = [];
    while (i < nodes.length) {
      const cur = nodes[i]!;
      if (isListItemMarker(cur) && cur.__list.ordered === ordered) {
        run.push(cur);
        i += 1;
      } else break;
    }
    out.push(buildNestedList(run, ordered, 0, run[0] && isListItemMarker(run[0]) ? (run[0] as unknown as { __list: ListMarker }).__list.numId : undefined));
  }
  return out;
}

function buildNestedList(
  items: PMNode[],
  ordered: boolean,
  level: number,
  numId?: string,
): PMNode {
  const content: PMNode[] = [];
  let i = 0;
  while (i < items.length) {
    const item = items[i]! as PMNode & { __list: ListMarker };
    if (item.__list.level < level) break;
    if (item.__list.level > level) {
      // Collect a sub-run belonging to the previous item.
      const subRun: PMNode[] = [];
      while (
        i < items.length &&
        isListItemMarker(items[i]!) &&
        (items[i] as PMNode & { __list: ListMarker }).__list.level > level
      ) {
        subRun.push(items[i]!);
        i += 1;
      }
      const nested = buildNestedList(subRun, ordered, level + 1, numId);
      const parent = content[content.length - 1];
      if (parent && parent.type === "listItem") {
        parent.content = [...(parent.content ?? []), nested];
      } else {
        content.push({ type: "listItem", content: [nested] });
      }
      continue;
    }
    content.push(stripMarker(item));
    i += 1;
  }
  const listAttrs = numId ? { numId } : undefined;
  return ordered
    ? { type: "orderedList", ...(listAttrs ? { attrs: listAttrs } : {}), content }
    : { type: "bulletList", ...(listAttrs ? { attrs: listAttrs } : {}), content };
}

function stripMarker<T extends PMNode>(node: T): T {
  if ("__list" in (node as object)) {
    const copy = { ...node } as unknown as { __list?: unknown };
    delete copy.__list;
    return copy as T;
  }
  return node;
}

function val(el: Element | undefined): string | undefined {
  if (!el) return undefined;
  return (el.attributes?.["w:val"] as string | undefined) ?? undefined;
}

/**
 * An OOXML on/off property defaults to *on* when the element is merely
 * present, and `<w:b w:val="0"/>` is how Word writes "not bold" — an explicit
 * reset inherited from the paragraph or character style. Reading presence
 * alone therefore turned every one of those resets into styled text.
 * `<w:u>` shares the attribute for its underline type, where `none` is the
 * only value that means off.
 */
function flagOn(el: Element | undefined): boolean {
  if (!el) return false;
  const v = val(el)?.trim().toLowerCase();
  if (v === undefined || v === "") return true;
  return !(v === "0" || v === "false" || v === "off" || v === "none");
}

function textOf(el: Element): string {
  let out = "";
  for (const c of el.children ?? []) {
    if (c.type === "text") out += (c as { text: string }).text;
  }
  return out;
}

function escapeXml(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function addMark(nodes: PMNode[], mark: PMMark): void {
  for (const n of nodes) {
    if (n.type === "text") {
      n.marks = [...(n.marks ?? []), mark];
    }
  }
}

/**
 * Merge adjacent text nodes that have identical marks. This keeps ProseMirror
 * happy and makes the JSON smaller.
 */
function mergeAdjacentText(nodes: PMNode[]): PMNode[] {
  const out: PMNode[] = [];
  for (const n of nodes) {
    if (n.type === "text" && n.text === "") continue;
    const last = out[out.length - 1];
    if (
      last &&
      last.type === "text" &&
      n.type === "text" &&
      sameMarks(last.marks, n.marks)
    ) {
      last.text = (last.text ?? "") + (n.text ?? "");
    } else {
      out.push(n);
    }
  }
  return out;
}

function sameMarks(a: PMMark[] | undefined, b: PMMark[] | undefined): boolean {
  if (!a && !b) return true;
  if (!a || !b) return false;
  if (a.length !== b.length) return false;
  return a.every((m, i) => JSON.stringify(m) === JSON.stringify(b[i]));
}
