import {
  children,
  collectDescendants,
  escapeAttr,
  findChildren,
  firstChild,
  isElement,
  onOffValue,
  parseDocument,
  parseFragment,
  serialize,
  val,
  type Element,
} from "./xml.js";
import type {
  PPrAttrs,
  PMDoc,
  PMMark,
  PMNode,
  PMTextNode,
  TabStop,
  TabStopAlignment,
  TabStopLeader,
} from "./pm-schema.js";
import { NumberingMap, readListProps, type NumberingKind } from "./numbering.js";
import { StyleMap, fitStyleId } from "./styles.js";
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
 * 页面尺寸在本 app 一路是 CSS px（`SectionGeometry`），段落属性（缩进、制表位
 * 位置）在编辑器里是磅。换算同样只允许发生在 docx-io，所以"版心宽度换成磅"
 * 这种一处需求也得从这里走，别在调用方再写一次 72/96。
 *
 * @param px 96 DPI 的 CSS px。
 * @returns 磅值，保留一位小数——与 `readTabStops` 的读侧口径一致，二次换算不会漂移。
 */
export function cssPxToPt(px: number): number {
  return Math.round(((px * 72) / 96) * 10) / 10;
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

type WriteCtx = {
  numbering?: NumberingMap;
  link?: LinkStyle | null;
  /**
   * 这份文档的 styles.xml 视图。段落 attrs 里的 styleId 要先对着它复核"级别/类型
   * 还匹配"才敢原样写回——详见 `fitStyleId`。
   */
  styles?: StyleMap;
};

export function pmToOoxml(doc: PMDoc, ctx: WriteCtx = {}): string {
  const blocks = dropUnpairedBookmarks(doc);
  const bodyChildren = blocks.map((b) => blockToXml(b, ctx)).join("");
  const ns = Object.entries(DRAWING_NS)
    .map(([p, uri]) => ` xmlns:${p}="${uri}"`)
    .join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"${ns}>
<w:body>${bodyChildren}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body>
</w:document>`;
}

/** `w:id`（ST_DecimalNumber）的合法形状：Word 写的是非负整数。读写两侧共用这一份判据。 */
const BOOKMARK_ID = /^\d{1,9}$/;

/** 书签的两个行内节点类型，配对只认 id，与它落在哪一段无关。 */
type BookmarkNode = Extract<PMNode, { type: "bookmarkStart" | "bookmarkEnd" }>;

/**
 * 按文档顺序取出树里的所有书签节点（父节点在前，再进 `content`）。
 *
 * @param nodes 某一层的内容；`undefined` 表示这一层没有子节点。
 * @param found 累加器，递归时复用同一个数组。
 * @returns 文档顺序的书签列表。
 */
function collectBookmarks(nodes: PMNode[] | undefined, found: BookmarkNode[] = []): BookmarkNode[] {
  for (const node of nodes ?? []) {
    if (node.type === "bookmarkStart" || node.type === "bookmarkEnd") found.push(node);
    const content = (node as { content?: PMNode[] }).content;
    if (content) collectBookmarks(content, found);
  }
  return found;
}

/**
 * 写之前丢掉配不上对的书签半截——编辑器里用户可以只删一半，而 ECMA-376 §17.5.2.1 说得很死：
 * "If no bookmarkEnd element exists subsequent to this element in document order with a
 * matching id attribute value, then the document is non-conformant"，Word 打开就弹提示。
 *
 * 配对的判据就是那句话：从前往后扫，起点先入栈，遇到同 id 的终点才算成对；
 * 只有终点的（起点被删了）与到文末还没等到终点的（终点被删了）一律不写。
 *
 * @param doc 编辑器交出来的文档 JSON。
 * @returns 可直接写出的顶层块：没有书签或全部成对时就是原数组，不重建任何节点。
 */
function dropUnpairedBookmarks(doc: PMDoc): PMNode[] {
  const all = collectBookmarks(doc.content);
  if (!all.length) return doc.content;
  const paired = new Set<BookmarkNode>();
  const pending = new Map<string, BookmarkNode>();
  for (const node of all) {
    if (node.type === "bookmarkStart") {
      // 同一个 id 出现两次起点：Word 的 id 是唯一的，第二次不算开新的区间。
      if (!pending.has(node.attrs.id)) pending.set(node.attrs.id, node);
    } else {
      const start = pending.get(node.attrs.id);
      if (start) {
        paired.add(start);
        paired.add(node);
        pending.delete(node.attrs.id);
      }
    }
  }
  if (paired.size === all.length) return doc.content;
  return withoutUnpaired(doc.content, paired);
}

/**
 * 重建内容树，只剔除 `paired` 之外的书签节点。
 *
 * 用展开复制而不是原地改数组：`listItem` 上挂着 `__list`、段落带着 `rawPpr`，
 * 浅拷贝能保住这些字段，不会因为过滤书签而顺手改掉别的属性。
 *
 * @param nodes 某一层的内容。
 * @param paired 配成功，可以写出的书签节点（按对象身份判定）。
 * @returns 过滤后的同层节点数组。
 */
function withoutUnpaired(nodes: PMNode[], paired: Set<BookmarkNode>): PMNode[] {
  const out: PMNode[] = [];
  for (const node of nodes) {
    if ((node.type === "bookmarkStart" || node.type === "bookmarkEnd") && !paired.has(node)) {
      continue;
    }
    const inner = (node as { content?: PMNode[] }).content;
    if (!inner) {
      out.push(node);
      continue;
    }
    // 整体展开再只换 content，`listItem` 的 `__list`、段落的 `rawPpr` 才会跟着走。
    // 这里的断言是必需的：TS 对联合类型做展开会把"给 text 节点加 content"也拼出来。
    out.push({ ...node, content: withoutUnpaired(inner, paired) } as PMNode);
  }
  return out;
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
 * Read a `<w:pPr>` into PM attrs: 具名样式 id、`<w:spacing>`、`<w:ind>`、`<w:tabs>` 与分页开关
 * (points / CSS ratio)。All unit conversions happen here so the UI can stay in
 * familiar points.
 *
 * 样式定义里的 `<w:pPr>` 走同一个函数（`style-format.ts` 解析样式链时调用）：
 * 单位换算只允许发生在 docx-io 的进出口，正文与样式各写一份就会对不上。
 */
export function readParagraphFormat(pPr: Element | null | undefined): PPrAttrs {
  const out: PPrAttrs = {};
  if (!pPr) return out;
  // 样式 id 原样带走，不与样式名互换：id↔名称的对应只有 styles.xml 知道。
  const styleId = val(firstChild(pPr, "w:pStyle"));
  if (styleId) out.styleId = styleId;
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
  const tabs = readTabStops(pPr);
  if (tabs) out.tabStops = tabs;
  return out;
}

/** `<w:tab>` 的 `w:val` 合法值，与 `TabStopAlignment` 一一对应（写侧据此丢掉读不懂的值）。 */
const TAB_ALIGNMENTS = [
  "left",
  "center",
  "right",
  "decimal",
  "bar",
  "list",
  "clear",
] as const satisfies readonly TabStopAlignment[];

/** `<w:tab>` 的 `w:leader` 合法值，与 `TabStopLeader` 一一对应。 */
const TAB_LEADERS = [
  "none",
  "dot",
  "hyphen",
  "dash",
  "line",
  "heavy",
  "middleDot",
] as const satisfies readonly TabStopLeader[];

/**
 * 读 `<w:pPr><w:tabs>` 成制表位清单。
 *
 * @param pPr 段落的 `<w:pPr>` 元素（可缺）。
 * @returns 建模好的清单，可能是空表；这一段没有 `w:tabs`、或其中任何一条读不懂
 *   （没有 `w:pos`、`w:val`/`w:leader` 是表外值）时返回 undefined，意思是"整组没建模过"——
 *   写侧因此把原文原样带走，而不是只留下能认的那几条。
 */
function readTabStops(pPr: Element | null | undefined): TabStop[] | undefined {
  const tabs = pPr ? firstChild(pPr, "w:tabs") : undefined;
  if (!tabs) return undefined;
  const stops: TabStop[] = [];
  for (const el of findChildren(tabs, "w:tab")) {
    const attributes = el.attributes ?? {};
    const pos = Number(attributes["w:pos"]);
    const rawVal = attributes["w:val"] as string | undefined;
    const rawLeader = attributes["w:leader"] as string | undefined;
    const alignment =
      rawVal === undefined
        ? "left"
        : (TAB_ALIGNMENTS as readonly string[]).includes(rawVal)
          ? (rawVal as TabStopAlignment)
          : undefined;
    const leader =
      rawLeader === undefined
        ? undefined
        : (TAB_LEADERS as readonly string[]).includes(rawLeader)
          ? (rawLeader as TabStopLeader)
          : undefined;
    if (!Number.isFinite(pos) || alignment === undefined) return undefined;
    if (rawLeader !== undefined && leader === undefined) return undefined;
    stops.push({
      positionPt: Math.round((pos / 20) * 10) / 10,
      alignment,
      ...(leader ? { leader } : {}),
    });
  }
  return stops;
}

/**
 * 写 `<w:tabs>`，是 `readTabStops` 的逆运算。
 *
 * @param stops 段落 attrs 里的制表位清单，位置是 pt；这里 ×20 换成 twips
 *   （AGENTS §2.1：换算只发生在 docx-io 进出口）。
 * @returns 元素字符串；清单为空时返回空串——调用方仍会把 `w:tabs` 记为"已建模"，
 *   于是源文档那条被清掉，对应 Word 里"全部清除"之后的段落。
 */
function tabsElementXml(stops: readonly unknown[]): string {
  const inner = stops.map(tabElementXml).join("");
  return inner ? `<w:tabs>${inner}</w:tabs>` : "";
}

/**
 * 一条 `<w:tab/>`，属性按 Word 自己的书写顺序（`w:val` → `w:leader` → `w:pos`）。
 *
 * @param stop attrs 里的一个制表位；不是对象或位置不是数字时产不出元素。
 * @returns 元素字符串；无法建模时返回空串（跳过这一条，不写一个假位置）。
 */
function tabElementXml(stop: unknown): string {
  const { positionPt, alignment, leader } = (stop ?? {}) as Partial<TabStop>;
  const position = Number(positionPt);
  if (!Number.isFinite(position)) return "";
  const val = (TAB_ALIGNMENTS as readonly string[]).includes(alignment ?? "")
    ? alignment
    : "left";
  const leaderAttr = (TAB_LEADERS as readonly string[]).includes(leader ?? "")
    ? ` w:leader="${leader}"`
    : "";
  return `<w:tab w:val="${val}"${leaderAttr} w:pos="${Math.round(position * 20)}"/>`;
}

/**
 * 写 OOXML 的开关型段落属性，是 `onOffValue`（`xml.ts`）的逆运算。
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

/**
 * `w:jc`/@w:val → 编辑器的 `textAlign`。
 *
 * @param jc OOXML 的对齐值。
 * @returns left 与读不懂的写法返回 undefined：那是 Word 的默认值，写成直接格式
 *   会在保存时多出用户没做过的改动。样式链解析与正文读侧共用这一份判断。
 */
export function jcToTextAlign(jc: string | undefined): string | undefined {
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
  // 复合域的 begin/separate/end 跨多个 run，而域的结果里还能再套域（目录的每条条目
  // 里就有一个 PAGEREF 域）。所以按深度计数：只有最外层那一对 begin/end 产出 field 节点，
  // 内层的指令与文字一律当作"缓存结果的文本"收走——否则内层那次 begin 会把外层的指令清零。
  const fieldTypes = fldCharTypes(p);
  // Word 生成的目录把 begin+instr+separate 写在首段、end 写在末段：这种域在本段配不上对，
  // 只能落 `fieldBegin`/`fieldEnd` 边界标记，让边界之后的内容回到正常结构（条目里的
  // PAGEREF 仍是 field 节点、制表符仍是 run），而不是被吞成一段缓存文本。
  const isBoundary = fieldTypes.includes("begin") && !closesInParagraph(fieldTypes);
  let fieldDepth = 0;
  let fieldInstr: string | null = null;
  let fieldDisplay = "";
  let boundaryDone = false;

  const flushField = () => {
    if (fieldInstr === null) return;
    out.push({
      type: "field",
      attrs: {
        kind: fieldKind(fieldInstr),
        instr: fieldInstr.trim(),
        // 缓存结果按字面留着，不 trim：目录条目的"标题⇥页码"就靠那个制表符对齐。
        display: fieldDisplay,
      },
    });
    fieldInstr = null;
    fieldDisplay = "";
  };

  /** 起始边界落成节点：指令归它自己，域内容从这里起按结构读，深度清零。 */
  const flushBoundary = (isSeparated: boolean) => {
    if (fieldInstr === null) return;
    out.push({
      type: "fieldBegin",
      attrs: {
        kind: fieldKind(fieldInstr),
        instr: fieldInstr.trim(),
        isSeparated,
      },
    });
    fieldInstr = null;
    fieldDepth = 0;
    boundaryDone = true;
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
            if (fieldDepth === 0) {
              fieldInstr = "";
              fieldDisplay = "";
            }
            fieldDepth++;
          } else if (t === "separate") {
            if (isBoundary && !boundaryDone && fieldDepth === 1) flushBoundary(true);
          } else if (t === "end") {
            if (fieldDepth === 0) {
              // 上一段的 begin 在这里收尾（目录末段）：没有配对的指令，只落一个结束边界。
              out.push({ type: "fieldEnd" });
            } else {
              fieldDepth = Math.max(0, fieldDepth - 1);
              if (fieldDepth === 0) flushField();
            }
          }
          consumed = true;
        } else if (sub.name === "w:instrText") {
          // 只有最外层那一段指令属于这个域；内层域的指令是它自己结果的一部分。
          if (fieldDepth === 1 && fieldInstr !== null) fieldInstr += textOf(sub);
          consumed = true;
        } else if (fieldDepth > 0) {
          fieldDisplay += fieldResultText(sub);
          consumed = true;
        }
      }
      if (!consumed) out.push(...runToInlines(child, ctx));
    } else if (fieldDepth > 0 && (child.name === "w:fldSimple" || child.name === "w:hyperlink")) {
      // 外层域的结果里嵌了简单域或超链接（`\h` 生成的目录条目就是整条被 w:hyperlink 包着的）：
      // 另立节点会让外层那条域的 begin/end 配不上对，所以整段折成缓存文本。
      fieldDisplay += fieldResultText(child);
    } else if (child.name === "w:fldSimple") {
      // Simple field: <w:fldSimple w:instr="PAGE"><w:r>...</w:r></w:fldSimple>
      const instr = (child.attributes?.["w:instr"] as string | undefined) ?? "";
      const texts = collectDescendants(child, "w:t").map(textOf).join("");
      out.push({
        type: "field",
        attrs: {
          kind: fieldKind(instr),
          instr: instr.trim(),
          display: texts,
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
    } else if (child.name === "w:bookmarkStart") {
      // 书签是 <w:p> 的直接子元素、与 <w:r> 平级，成对靠 w:id 匹配（ECMA-376 §17.5.2.1）。
      // `_GoBack`（Word 自己的"回到上次编辑处"）也照原样建模：终点只带 id，
      // 单独丢掉起点会留下规范明确判为 non-conformant 的孤儿终点。
      const id = child.attributes?.["w:id"] as string | undefined;
      const name = child.attributes?.["w:name"] as string | undefined;
      if (id && BOOKMARK_ID.test(id)) {
        out.push({ type: "bookmarkStart", attrs: name ? { id, name } : { id } });
      }
    } else if (child.name === "w:bookmarkEnd") {
      const id = child.attributes?.["w:id"] as string | undefined;
      if (id && BOOKMARK_ID.test(id)) out.push({ type: "bookmarkEnd", attrs: { id } });
    } else if (isElement(child)) {
      for (const sub of children(child)) {
        if (sub.name === "w:r") out.push(...runToInlines(sub, ctx));
      }
    }
  }
  // If a field was unterminated, flush what we have so it isn't lost.
  if (isBoundary && !boundaryDone) {
    // 有 begin 有指令但没写到 separate（Word 标记为 dirty、等着自己重算的域）：
    // 仍按起始边界建模，指令之后的内容折成文字，免得整段丢掉。
    flushBoundary(false);
    if (fieldDisplay) out.push({ type: "text", text: fieldDisplay });
    fieldDisplay = "";
  } else {
    flushField();
  }
  const inlines = mergeAdjacentText(out);
  if (!ctx.floats) return inlines;
  // The marks only say "a box sat here"; the boxes are block-level and were
  // collected into `ctx.floats` in source order.
  return inlines.filter((n) => n !== FLOAT_SLOT_MARK);
}

/**
 * 段落里所有 `w:fldChar` 的类型，按出现顺序。
 *
 * @param p 段落元素。
 * @returns 形如 `["begin", "separate", "end"]` 的序列；没有域时为空数组。
 */
function fldCharTypes(p: Element): string[] {
  const types: string[] = [];
  for (const child of children(p)) {
    if (child.name !== "w:r") continue;
    for (const sub of children(child)) {
      if (sub.name === "w:fldChar") types.push(String(sub.attributes?.["w:fldCharType"] ?? ""));
    }
  }
  return types;
}

/**
 * 最外层那对 begin/end 是否在同一段落里闭合。
 *
 * @param types `fldCharTypes` 的结果。
 * @returns true 表示段内域（用 field 节点），false 表示跨段域（用边界标记）。
 */
function closesInParagraph(types: string[]): boolean {
  let depth = 0;
  for (const t of types) {
    if (t === "begin") depth++;
    else if (t === "end") {
      depth = Math.max(0, depth - 1);
      if (depth === 0) return true;
    }
  }
  return false;
}

/**
 * 域结果里的任意元素 → 缓存文本（递归到 run 的子元素）。
 *
 * OOXML 的制表符与换行是**元素**（`<w:tab/>`、`<w:br/>`）而不是字符，只取 `w:t`
 * 会让目录条目的"标题⇥页码"塌成"标题页码"，所以这里换成字面字符；
 * `fieldResultRuns` 是它的反面。
 *
 * @param el 域结果里遇到的一个元素（run、run 的子元素、超链接或简单域）。
 * @returns 这个元素贡献的文字；不贡献文字的（rPr 等）返回空串。
 */
function fieldResultText(el: Element): string {
  if (el.name === "w:t") return textOf(el);
  if (el.name === "w:tab") return "\t";
  if (el.name === "w:br") return "\n";
  if (el.name === "w:noBreakHyphen") return "\u2011";
  if (el.name === "w:rPr" || el.name === "w:pPr") return "";
  return children(el).map(fieldResultText).join("");
}

function fieldKind(instr: string): string {
  const m = instr.trim().match(/^([A-Za-z]+)/);
  return m && m[1] ? m[1].toUpperCase() : "FIELD";
}

/**
 * `<w:rPr>` 里的直接字符格式，读成编辑器 attrs 的词汇：字号/字距是 CSS px，颜色 #rrggbb。
 *
 * run 与样式定义读的是同一张属性表，因此只留这一份实现：样式链解析（`style-format.ts`）
 * 复制一遍就会出现"正文里读得到、样式里读不到"的分歧（AGENTS §3.2）。
 *
 * @param rPr 已解析的 `w:rPr` 元素；undefined 表示这份 run/样式没写字符属性。
 * @returns 只带这份 rPr 真写过的项。四个开关是三态：true/false 是写过——
 *   `<w:b w:val="0"/>` 是 Word 用来**显式取消**从样式继承来的加粗——undefined 才是"没提过"。
 *   颜色读不懂时整个不出现，宁可丢色也不把 Word 认不了的值往下带。
 */
export interface RunStyleProps {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  /** 中文那一支字体（`w:eastAsia`），没有它才回落到西文那支。 */
  fontFamily?: string;
  /** 中西文各一支时（模板标配：宋体 + Times New Roman）的西文那支。 */
  fontFamilyWest?: string;
  /** CSS px，不是 pt。 */
  fontSize?: string;
  /** CSS px，允许负值（OOXML 的 `w:spacing` 收紧字距）。 */
  letterSpacing?: string;
  color?: string;
  /** `w:highlight` 的荧光笔色，其次才是 `w:shd` 的底纹填充。 */
  backgroundColor?: string;
}

/**
 * 读一个 `<w:rPr>`（或 `<w:rPrDefault>` 里的 `w:rPr`）。
 *
 * @param rPr 已解析的元素，undefined 表示没有字符属性。
 * @returns 见 {@link RunStyleProps}；元素为 undefined 时是空对象。
 */
export function readRunProperties(rPr: Element | undefined): RunStyleProps {
  const out: RunStyleProps = {};
  if (!rPr) return out;
  // 元素在就是"写过"，值决定开还是关；元素不在才是 undefined。
  for (const [key, tag] of [
    ["bold", "w:b"],
    ["italic", "w:i"],
    ["underline", "w:u"],
    ["strike", "w:strike"],
  ] as const) {
    const el = firstChild(rPr, tag);
    if (el) out[key] = flagOn(el);
  }
  const rFonts = firstChild(rPr, "w:rFonts");
  if (rFonts) {
    const attrs = rFonts.attributes ?? {};
    const eastAsia = attrs["w:eastAsia"] as string | undefined;
    const west =
      (attrs["w:ascii"] as string | undefined) || (attrs["w:hAnsi"] as string | undefined);
    if (eastAsia || west) out.fontFamily = eastAsia || west;
    if (eastAsia && west && eastAsia !== west) out.fontFamilyWest = west;
  }
  const sz = firstChild(rPr, "w:sz");
  if (sz) {
    const half = Number(val(sz));
    if (half && !Number.isNaN(half)) out.fontSize = ptToCssFontSize(half / 2);
  }
  const spacing = firstChild(rPr, "w:spacing");
  if (spacing) {
    const twips = Number(val(spacing));
    if (Number.isFinite(twips) && twips !== 0) {
      out.letterSpacing = ptToCssLetterSpacing(twips / 20);
    }
  }
  const color = firstChild(rPr, "w:color");
  if (color) {
    const c = canonColor(colorVal(val(color)));
    if (c) out.color = c;
  }
  const highlight = firstChild(rPr, "w:highlight");
  if (highlight) {
    const v = val(highlight);
    const c = v ? canonColor(WORD_HIGHLIGHT[v] ?? colorVal(v)) : null;
    if (c) out.backgroundColor = c;
  }
  const shd = firstChild(rPr, "w:shd");
  if (shd && out.backgroundColor === undefined) {
    const c = canonColor(colorVal(shd.attributes?.["w:fill"] as string | undefined));
    if (c) out.backgroundColor = c;
  }
  return out;
}

function runToInlines(r: Element, ctx: InlineCtx = {}): PMNode[] {
  const props = readRunProperties(firstChild(r, "w:rPr"));
  const marks: PMMark[] = [];
  if (props.bold === true) marks.push({ type: "bold" });
  if (props.italic === true) marks.push({ type: "italic" });
  if (!ctx.inLink && props.underline === true) marks.push({ type: "underline" });
  if (props.strike === true) marks.push({ type: "strike" });
  const textStyle: NonNullable<Extract<PMMark, { type: "textStyle" }>["attrs"]> = {};
  if (props.fontFamily !== undefined) textStyle.fontFamily = props.fontFamily;
  if (props.fontFamilyWest !== undefined) textStyle.fontFamilyWest = props.fontFamilyWest;
  if (props.fontSize !== undefined) textStyle.fontSize = props.fontSize;
  if (props.letterSpacing !== undefined) textStyle.letterSpacing = props.letterSpacing;
  // 链接文字的颜色由 link 决定，rPr 里那条 w:color 不能再盖上去。
  if (props.color !== undefined && !ctx.inLink) textStyle.color = props.color;
  if (props.backgroundColor !== undefined) textStyle.backgroundColor = props.backgroundColor;
  if (Object.keys(textStyle).length) marks.push({ type: "textStyle", attrs: textStyle });

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
      return paragraphXml(node, { link: ctx.link, styles: ctx.styles });
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
      return horizontalRuleXml();
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
    /** 标题节点现在的级别；非标题段落不要传。 */
    headingLevel?: number;
    styles?: StyleMap;
  } = {},
): string {
  const a = node.attrs ?? {};
  const pPr = paragraphPprXml(a, {
    pStyle: opts.pStyle,
    numPr: opts.numPr,
    rawPpr: opts.rawPpr ?? a.rawPpr,
    headingLevel: opts.headingLevel,
    styles: opts.styles,
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
 * 按给定顺序表稳定排序子元素，拼成内部 XML。
 *
 * @param parts 子元素列表，允许乱序；同名元素的相对次序保持传入顺序。
 * @param order 顺序表（都是带 `w:` 前缀的元素名）；表里没有的名字排到最后。
 * @returns 排好序的子元素串；空列表返回空串。
 */
function sortParts(parts: PprPart[], order: readonly string[]): string {
  const rankOf = (name: string): number => {
    const i = order.indexOf(name);
    return i < 0 ? order.length : i;
  };
  return parts
    .map((part, index) => ({ index, rank: rankOf(part.name), xml: part.xml }))
    .sort((a, b) => a.rank - b.rank || a.index - b.index)
    .map((part) => part.xml)
    .join("");
}

/**
 * 按 CT_PPr 子元素顺序拼出 `<w:pPr>`。
 *
 * @param parts 待写出的子元素。
 * @returns 完整的 pPr 字符串；一个子元素都没有时返回空串（段落直接不带 pPr）。
 */
function pprFromParts(parts: PprPart[]): string {
  const inner = sortParts(parts, PPR_CHILD_ORDER);
  return inner ? `<w:pPr>${inner}</w:pPr>` : "";
}

/**
 * 按 CT_PPr 子元素顺序拼出 pPr 的**内部** XML（不含外层标签）。
 *
 * @param parts 待写出的子元素。
 * @returns 内部串；空列表返回空串。样式定义里的 `w:pPr` 存的是内部 XML，
 *   所以读侧与写侧共用同一个形状。
 */
function pprInnerFromParts(parts: PprPart[]): string {
  return sortParts(parts, PPR_CHILD_ORDER);
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
  // 制表位是唯一"空清单也要记名"的一项：用户点了全部清除时 attrs 是 []，
  // 写出的是没有 w:tabs 的 pPr，但 covered 必须有它，否则源文档那条又被透传回来。
  if (Array.isArray(a.tabStops)) {
    parts.push({ name: "w:tabs", xml: tabsElementXml(a.tabStops) });
  }
  push("w:jc", jcXml(a.textAlign));
  const lvl = a.outlineLevel;
  if (typeof lvl === "number" && Number.isInteger(lvl) && lvl >= 0 && lvl <= 9) {
    push("w:outlineLvl", `<w:outlineLvl w:val="${lvl}"/>`);
  }
  return parts;
}

/**
 * 源文档 pPr 里我们**没建模过**的子元素，原样带走。
 *
 * @param rawEls 已解析的 rawPpr 子元素。
 * @param covered attrs 这一轮已经写出的元素名，同名元素由 attrs 说了算。
 * @returns 待排序写出的子元素。`w:pStyle` 永远不在这里生成：它由 `pStylePart` 单独
 *   决定，同一份 pPr 里出现两个时 Word 取第一个，套哪个样式变得不可预测；样式定义里
 *   写它更是自己引用自己。`w:numPr` 与其余没建模的写法（`w:pBdr`、 attrs 里没有
 *   `tabStops` 时的 `w:tabs`）一样带走——我们从选区抓不到列表编号，更没有权力删掉
 *   用户原本挂在段落或样式上的编号。
 */
function unmodeledRawParts(rawEls: Element[], covered: Set<string>): PprPart[] {
  return rawEls
    .map((el) => ({ name: el.name, xml: serialize(el) }))
    .filter((part) => !covered.has(part.name) && part.name !== "w:pStyle");
}

/**
 * 生成段落的 `<w:pPr>`，两条路径（有 rawPpr 的导入段落 / 编辑器新建段落）共用。
 *
 * 合并规则：attrs 建模过的元素顶替 rawPpr 里的同名元素——spacing/ind 内部还要
 * 逐个属性合并（见 `mergedElementXml`），因为源文档里可能有我们没建模的写法
 * （`w:*Chars`、`w:*Lines`、`w:beforeLines`）。其余元素原样保留并按 CT_PPr
 * 顺序排回该在的位置，源文档带 `<w:rPr>`（段落标记的字符属性，Word 写得非常多）
 * 时不会再被挤到 spacing 前面。
 *
 * @param a 段落 attrs。
 * @param opts.pStyle 从节点类型推断的样式 id（Heading1 / Quote / ListNumber）；
 *   attrs 建模了具名样式、或源文档自己带 pStyle 时都不用它。
 * @param opts.numPr 列表段落重新分组后由我们给出的 numPr。
 * @param opts.rawPpr 源文档的 pPr 内部 XML。
 * @param opts.headingLevel 标题节点现在的级别（1–6），非标题段落为 undefined。
 * @param opts.styles 这份文档的样式表，用来复核 styleId 还配不配这段落。
 * @returns 完整 pPr 字符串；段落没有任何格式属性时返回空串。
 */
function paragraphPprXml(
  a: PPrAttrs,
  opts: {
    pStyle?: string;
    numPr?: string;
    rawPpr?: string;
    headingLevel?: number;
    styles?: StyleMap;
  } = {},
): string {
  const modeled = modeledPprParts(a, opts.rawPpr);
  const rawEls = opts.rawPpr ? parseFragment(opts.rawPpr) : [];
  const rawStyle = rawEls.find((el) => el.name === "w:pStyle");
  // pStyle 只由 pStylePart 决定，别让它从两个来源各进来一次——
  // 同一份 pPr 里出现两个 w:pStyle 时 Word 取第一个，套哪个样式变得不可预测。
  const rawParts = unmodeledRawParts(rawEls, new Set(modeled.map((part) => part.name)));
  const parts: PprPart[] = [];
  const applied = fitStyleId(a.styleId, opts.headingLevel, opts.styles);
  const style = pStylePart(applied, rawStyle, opts.pStyle);
  if (style) parts.push(style);
  parts.push(...rawParts);
  if (opts.numPr) parts.push({ name: "w:numPr", xml: opts.numPr });
  parts.push(...modeled);
  return pprFromParts(parts);
}

/**
 * 由"从选区抓到的格式"生成一条**样式定义**的 pPr / rPr 内部 XML。
 *
 * 与段落写侧共用同一批生成器（`modeledPprParts` / `runPropertiesParts`），所以字号、
 * 缩进、行距的换算在整个仓库里只有一份（AGENTS §2.1）。样式定义不写 `w:pStyle`：
 * 那是自己引用自己。我们从选区抓不到列表编号，所以也从不生成 `w:numPr`，
 * 但被改样式原本挂着的编号与其他没建模的写法一样原样带走。
 *
 * @param format.paragraph 段落 attrs（段前/缩进是 pt，lineHeight 是倍数或 pt:N / atLeast:N）。
 * @param format.marks 选区那段文字的行内标记。
 * @param existing 被改样式原有的 pPrInner / rPrInner。同名元素被新值顶替，我们没
 *   建模过的写法（`w:tabs`、`w:pBdr`、`w:lang`）原样留着——Word 的"更新样式以匹配
 *   所选内容"是在原定义上合并，不是整条重建。
 * @returns 内部 XML（不含外层标签）；两个空串表示这条样式不带任何属性。
 */
export function stylePropertiesXml(
  format: { paragraph: PPrAttrs; marks: PMMark[] },
  existing: { pPrInner?: string; rPrInner?: string } = {},
): { pPrInner: string; rPrInner: string } {
  const modeled = modeledPprParts(format.paragraph, existing.pPrInner);
  const rawParts = unmodeledRawParts(
    existing.pPrInner ? parseFragment(existing.pPrInner) : [],
    new Set(modeled.map((part) => part.name)),
  );
  const runParts = runPropertiesParts(format.marks, null);
  const covered = new Set(
    runParts.flatMap((part) => [part.name, ...(part.alsoCovers ?? [])]),
  );
  // 源定义里没被这次抓取覆盖的元素留着；w:szCs 跟着 w:sz 走（两者永远成对写出），
  // 否则旧字号留在后面，Word 按先出现的那条排，字号只改了一半。
  const keptRunParts = (existing.rPrInner ? parseFragment(existing.rPrInner) : [])
    .map((el) => ({ name: el.name, xml: serialize(el) }))
    .filter((part) => !covered.has(part.name));
  return {
    pPrInner: pprInnerFromParts([...rawParts, ...modeled]),
    rPrInner: rprInnerFromParts([...keptRunParts, ...runParts]),
  };
}

/**
 * 决定这段落写出的 `<w:pStyle>`。
 *
 * @param styleId attrs 里的具名样式：字符串=套用它（用户在样式窗格里选的）；
 *   null=用户"清除样式"，源文档那条必须删；undefined=没建模过，保留源文档写法。
 * @param rawEl 源文档 pPr 里原来的 `w:pStyle` 元素，可能没有。
 * @param inferred 从节点类型推断的 id，只在两侧都没有时用。
 * @returns 要写出的 pPr 子元素；不该出现时返回 null。id 与源文档一致时按原元素
 *   写回，避免一次无关的保存改动样式引用的写法。
 */
function pStylePart(
  styleId: string | null | undefined,
  rawEl: Element | undefined,
  inferred?: string,
): PprPart | null {
  if (styleId === null) return null;
  if (typeof styleId === "string") {
    if (rawEl && String(rawEl.attributes?.["w:val"] ?? "") === styleId) {
      return { name: "w:pStyle", xml: serialize(rawEl) };
    }
    return { name: "w:pStyle", xml: `<w:pStyle w:val="${escapeAttr(styleId)}"/>` };
  }
  if (rawEl) return { name: "w:pStyle", xml: serialize(rawEl) };
  return inferred ? { name: "w:pStyle", xml: `<w:pStyle w:val="${inferred}"/>` } : null;
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
    headingLevel: Number(node.attrs.level),
    link: ctx.link,
    styles: ctx.styles,
  });
}

function blockquoteXml(
  node: Extract<PMNode, { type: "blockquote" }>,
  ctx: WriteCtx = {},
): string {
  return (node.content ?? [])
    .map((child) => {
      if (child.type === "paragraph") {
        return paragraphXml(child, { pStyle: "Quote", link: ctx.link, styles: ctx.styles });
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

/**
 * 分隔线导出成"段落底部单线边框"，也就是 Word 里连打三个 `-` 自动更正后的形态。
 *
 * 之前写的是 `<w:pStyle w:val="HorizontalRule"/>`，而本仓不产出 styles.xml，
 * 那个样式定义根本不存在，Word 只会排出一个空段落 —— 分隔线静默消失。
 * `w:sz` 的单位是 1/8 磅（ECMA-376 §17.3.4：sz=18 即 2.25 磅），6 = 0.75 磅。
 *
 * @returns 一个只带 <w:pBdr> 的完整段落 XML，无正文。
 */
function horizontalRuleXml(): string {
  return `<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr></w:pPr></w:p>`;
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
          return paragraphXml(withBold, { link: ctx.link, styles: ctx.styles });
        }
        return paragraphXml(c, { link: ctx.link, styles: ctx.styles });
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
    // 这几类各自成组，不跟文本合用一个 <w:r>：书签压根不是 run 的内容，另三类各带自己的 run 结构。
    if (
      n.type === "image" ||
      n.type === "field" ||
      n.type === "pageBreak" ||
      n.type === "bookmarkStart" ||
      n.type === "bookmarkEnd" ||
      n.type === "fieldBegin" ||
      n.type === "fieldEnd"
    ) {
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
  // 书签单独成组时直接吐元素：它属于 <w:p> 的内容模型，包进 <w:r> Word 就不认了。
  if (nodes.length === 1 && (nodes[0]!.type === "bookmarkStart" || nodes[0]!.type === "bookmarkEnd")) {
    return bookmarkXml(nodes[0]!);
  }
  // 跨段域的边界标记同理单独成组：它们各自的 run 里只有 fldChar / instrText，
  // 合进文本 run 会把域边界和正文混在一个 run 里（Word 会把指令当成正文显示）。
  if (nodes.length === 1 && (nodes[0]!.type === "fieldBegin" || nodes[0]!.type === "fieldEnd")) {
    return fieldBoundaryXml(nodes[0]!);
  }
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
    if (!buffer.length) return;
    // 制表符在 OOXML 里是 run 的元素 `<w:tab/>`，不是字符：留在 `<w:t>` 里只是一个空白字符，
    // 目录条目"标题⇥页码"的对齐就没了（换行同一个道理，见下面那个 replace）。
    const segments = buffer.split("\t");
    segments.forEach((segment, index) => {
      if (index > 0) parts.push("<w:tab/>");
      if (!segment.length) return;
      // Escape newlines as a <w:br/> inside the run（先转义再插记号，反了会把记号也转义掉）。
      const escaped = escapeXml(segment).replace(/\n/g, `</w:t><w:br/><w:t xml:space="preserve">`);
      parts.push(`<w:t xml:space="preserve">${escaped}</w:t>`);
    });
    buffer = "";
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

/** RPR_ORDER 的带前缀写法，供 `sortParts` 与段落那边用同一套排序。 */
const RPR_ELEMENT_ORDER = RPR_ORDER.map((key) => `w:${key}`);

/**
 * 按 CT_RPr 顺序拼出 `<w:rPr>` 的内部 XML。
 *
 * @param parts 子元素列表，允许乱序（同名元素的相对次序保持传入顺序）。
 * @returns 内部串；空列表返回空串。
 */
function rprInnerFromParts(parts: PprPart[]): string {
  return sortParts(parts, RPR_ELEMENT_ORDER);
}

/** A font name as an XML attribute value. */
function fontAttr(name: string): string {
  return escapeXml(name).replace(/"/g, "");
}

/** `<w:rPr>` 的一个子元素 + 它顶替掉的原定义元素名（`sz` 要连 `szCs` 一起换）。 */
interface RunPropertyPart extends PprPart {
  /** 除自身之外还要被这条顶替掉的原元素名。 */
  alsoCovers?: string[];
}

function runPropertiesXml(marks: PMMark[], link?: LinkStyle | null): string {
  const inner = rprInnerFromParts(runPropertiesParts(marks, link));
  return inner ? `<w:rPr>${inner}</w:rPr>` : "";
}

/**
 * 由行内标记生成 `<w:rPr>` 的子元素序列。
 *
 * @param marks 这段文字的标记（编辑器里的字号是 CSS px、颜色是 #rrggbb）。
 * @param link 文档级超链接样式；null/undefined 表示这段不是链接。
 * @returns 每个元素名一项（`w:sz` 那一项同时带 `w:szCs`，两者永远成对写出）。
 *   没有任何属性时返回空数组。
 */
function runPropertiesParts(marks: PMMark[], link?: LinkStyle | null): RunPropertyPart[] {
  // One writer per element name. A duplicated `<w:color>` is not merely
  // untidy: Word honours the first and ignores the rest, which is how a link
  // that also carried the user's own colour used to lose it.
  const el = new Map<string, RunPropertyPart>();
  const styleMark = marks.find(
    (m): m is Extract<PMMark, { type: "textStyle" }> => m.type === "textStyle",
  );
  const ownColor = styleMark?.attrs.color;
  const ownUnderline = marks.some((m) => m.type === "underline");
  // A link's colour and rule are the document-level setting, and only when the
  // run has none of its own.
  const linkColor = cssColorToOoxmlHex(link?.color ?? "") ?? "0563C1";
  const linkUnderline = link?.underline !== false;

  const put = (part: RunPropertyPart): void => {
    el.set(part.name, part);
  };

  if (marks.some((m) => m.type === "bold")) put({ name: "w:b", xml: "<w:b/>" });
  if (marks.some((m) => m.type === "italic")) put({ name: "w:i", xml: "<w:i/>" });
  if (ownUnderline) put({ name: "w:u", xml: `<w:u w:val="single"/>` });
  if (marks.some((m) => m.type === "strike")) put({ name: "w:strike", xml: "<w:strike/>" });
  if (marks.some((m) => m.type === "code")) {
    put({
      name: "w:rFonts",
      xml: `<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/>`,
    });
  }
  if (marks.some((m) => m.type === "link")) {
    if (!ownColor) put({ name: "w:color", xml: `<w:color w:val="${linkColor}"/>` });
    if (!ownUnderline && linkUnderline) put({ name: "w:u", xml: `<w:u w:val="single"/>` });
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
      put({ name: "w:rFonts", xml: `<w:rFonts ${attrs}/>` });
    }
    if (fontSize) {
      const pt = cssFontSizeToPt(String(fontSize));
      const n = pt === null ? 0 : Math.round(pt * 2);
      if (n > 0)
        put({
          name: "w:sz",
          xml: `<w:sz w:val="${n}"/><w:szCs w:val="${n}"/>`,
          alsoCovers: ["w:szCs"],
        });
    }
    if (letterSpacing) {
      // rPr `w:spacing` is character spacing in twentieths of a point, and it
      // may be negative (tightening) — only zero means "inherit".
      const pt = cssLetterSpacingToPt(String(letterSpacing));
      const twips = pt === null ? 0 : Math.round(pt * 20);
      if (twips !== 0) put({ name: "w:spacing", xml: `<w:spacing w:val="${twips}"/>` });
    }
    if (color) {
      const hex = cssColorToOoxmlHex(color);
      if (hex) put({ name: "w:color", xml: `<w:color w:val="${hex}"/>` });
    }
    if (backgroundColor) {
      // Word's 16 highlight names win when the colour is one of them; any
      // other colour goes through shading. Both accept only hex, so an
      // unreadable colour is dropped here rather than written broken.
      const raw = String(backgroundColor).trim().toLowerCase();
      const hex = cssColorToOoxmlHex(raw);
      const named = HIGHLIGHT_FROM_CSS[hex ? `#${hex.toLowerCase()}` : raw];
      if (named) {
        put({ name: "w:highlight", xml: `<w:highlight w:val="${named}"/>` });
      } else if (hex) {
        put({
          name: "w:shd",
          xml: `<w:shd w:val="clear" w:color="auto" w:fill="${hex}"/>`,
        });
      }
    }
  }

  return RPR_ORDER.map((k) => el.get(`w:${k}`)).filter((p): p is RunPropertyPart => !!p);
}

function fieldXml(node: Extract<PMNode, { type: "field" }>): string {
  const instr = node.attrs.instr ?? node.attrs.kind ?? "PAGE";
  const display = node.attrs.display ?? "";
  return (
    `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
    `<w:r><w:instrText xml:space="preserve"> ${escapeXml(instr)} </w:instrText></w:r>` +
    `<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
    fieldResultRuns(display) +
    `<w:r><w:fldChar w:fldCharType="end"/></w:r>`
  );
}

/**
 * 跨段落域的边界节点 → OOXML 的 run 片段。
 *
 * `fieldBegin` 吐 begin 记号 + 指令 run（指令必须待在 `w:instrText` 里，写成 `w:t`
 * 会被 Word 当成正文显示出来）+ 可选的 separate 记号；`fieldEnd` 只吐 end 记号。
 *
 * @param node `fieldBegin` 或 `fieldEnd` 节点。
 * @returns 直接插进 `<w:p>` 的 `<w:r>` 序列。
 */
function fieldBoundaryXml(
  node: Extract<PMNode, { type: "fieldBegin" }> | Extract<PMNode, { type: "fieldEnd" }>,
): string {
  if (node.type === "fieldEnd") return `<w:r><w:fldChar w:fldCharType="end"/></w:r>`;
  const instr = node.attrs.instr ?? node.attrs.kind ?? "TOC";
  const separate = node.attrs.isSeparated
    ? `<w:r><w:fldChar w:fldCharType="separate"/></w:r>`
    : "";
  return (
    `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
    `<w:r><w:instrText xml:space="preserve"> ${escapeXml(instr)} </w:instrText></w:r>` +
    separate
  );
}

/**
 * 缓存结果文本 → 结果 run 序列，是 `fieldResultText` 的反向换算。
 *
 * display 里的 `\t`、`\n` 必须还原成 `<w:tab/>`、`<w:br/>` 元素：OOXML 的 run 里
 * 没有"制表符字符"，把它写进 `<w:t>` 只会变成一个不可见字符，目录条目的页码对齐就没了。
 *
 * @param display 编辑器里的缓存结果，可含 `\t` / `\n` / `\u2011`。
 * @returns 零个或多个 `<w:r>`；display 为空时返回空串——Word 允许域还没有算过结果。
 */
function fieldResultRuns(display: string): string {
  if (!display) return "";
  let xml = "";
  let pending = "";
  const flush = () => {
    if (pending) xml += `<w:r><w:t xml:space="preserve">${escapeXml(pending)}</w:t></w:r>`;
    pending = "";
  };
  for (const char of display) {
    if (char === "\t") {
      flush();
      xml += `<w:r><w:tab/></w:r>`;
    } else if (char === "\n") {
      flush();
      xml += `<w:r><w:br/></w:r>`;
    } else {
      pending += char;
    }
  }
  flush();
  return xml;
}

/**
 * 书签节点 → OOXML。`w:bookmarkStart`/`w:bookmarkEnd` 是 `<w:p>` 的直接子元素，
 * 与 `<w:r>` 平级（§17.5.2.1 的例子就夹在两个 run 之间），所以绝不能包进 run 里——
 * 包了就成了 run 的内容，Word 读不认，书签等于没写。
 *
 * @param node 已通过 `dropUnpairedBookmarks` 配对的书签节点。
 * @returns 插进 `<w:p>` 的片段；id 不是非负整数时返回空串（宁可不写也不产出违规元素）。
 */
function bookmarkXml(node: BookmarkNode): string {
  const id = String(node.attrs.id ?? "");
  if (!BOOKMARK_ID.test(id)) return "";
  if (node.type === "bookmarkEnd") return `<w:bookmarkEnd w:id="${id}"/>`;
  const name = node.attrs.name;
  return `<w:bookmarkStart w:id="${id}"${name ? ` w:name="${escapeXml(name)}"` : ""}/>`;
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
    .replace(/>/g, "&gt;")
    // 属性值一律带引号（w:name / w:anchor / w:r：…），不转义 `"` 的话一个引号
    // 就能把元素截断，Word 直接判文档打不开。
    .replace(/"/g, "&quot;");
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
