/**
 * Geometry and XML helpers for Word floating shapes that carry text
 * (`wps:txbx` inside a `wp:anchor`, and the legacy `v:textbox` inside a
 * `w:pict`).
 *
 * Word lays out whole documents (resume templates especially) out of
 * absolutely-positioned text boxes: ~100% of the visible text can live inside
 * them. The editor models those as `floatGroup`/`floatBox` blocks, and this
 * module answers the two questions the converter needs:
 *
 *   import  — where does each shape sit, and how does it look? Shapes inside a
 *             group (`wpg:grpSp`) are expressed in the group's *child*
 *             coordinate space, so they have to be scaled into the anchor's
 *             space before they mean anything.
 *   export  — how do we put edited text back into the untouched shape XML?
 *
 * The round-trip strategy is deliberately conservative: we keep the source
 * element's XML verbatim as a template and splice `<w:txbxContent>` bodies into
 * it, so fills, gradients, group transforms and the `mc:Fallback` VML branch
 * survive byte-for-byte unless the user actually retypes inside a box.
 */
import {
  children,
  collectDescendants,
  firstChild,
  text,
  type Element,
} from "./xml.js";

/** EMU per CSS px (914400 per inch at 96 dpi). */
const EMU_PER_PX = 9525;

export interface AnchorGeom {
  /** Horizontal offset measured from `relH`, in CSS px. */
  x: number;
  /** Vertical offset from `relV`, in CSS px. */
  y: number;
  /** `wp:positionH/@relativeFrom` (page | column | margin | paragraph | character). */
  relH: string;
  /** `wp:positionV/@relativeFrom`. */
  relV: string;
  /** `wp:extent`, in CSS px. */
  w: number;
  h: number;
  /** `wp:anchor/@relativeHeight`, used only as a stacking hint. */
  zOrder: number;
}

export interface ShapeInfo {
  /** Position inside the anchor's box, in CSS px. */
  x: number;
  y: number;
  w: number;
  h: number;
  /** CSS background for the shape fill, or null when the shape has no fill. */
  fill: string | null;
  /** CSS border color, or null when the shape is not outlined. */
  stroke: string | null;
  strokeWidth: number | null;
  /** `wps:bodyPr/@anchor` mapped to a CSS vertical alignment. */
  vAlign: string | null;
  /** The shape's `<w:txbxContent>` element; null for a decorative shape. */
  txbxContentEl: Element | null;
  name: string | null;
}

const int = (v: string | undefined | null): number | undefined => {
  if (v === undefined || v === null || v === "") return undefined;
  const n = Number(v);
  return Number.isFinite(n) ? n : undefined;
};

const attr = (el: Element | undefined, name: string): string | undefined =>
  el?.attributes?.[name] as string | undefined;

function numAttr(el: Element | undefined, name: string): number | undefined {
  return int(attr(el, name));
}

/** `<a:off x y/>` + `<a:ext cx cy/>` under an `<a:xfrm>`, in EMU. */
function readXfrm(xfrm: Element | undefined) {
  const off = xfrm ? firstChild(xfrm, "a:off") : undefined;
  const ext = xfrm ? firstChild(xfrm, "a:ext") : undefined;
  const chOff = xfrm ? firstChild(xfrm, "a:chOff") : undefined;
  const chExt = xfrm ? firstChild(xfrm, "a:chExt") : undefined;
  return {
    x: numAttr(off, "x") ?? 0,
    y: numAttr(off, "y") ?? 0,
    cx: numAttr(ext, "cx") ?? 0,
    cy: numAttr(ext, "cy") ?? 0,
    // A connector line really is 0 tall, which is not the same as carrying no
    // `<a:ext>` at all and meaning "as big as the anchor".
    hasExt: !!ext,
    chX: numAttr(chOff, "x") ?? 0,
    chY: numAttr(chOff, "y") ?? 0,
    chCx: numAttr(chExt, "cx") ?? 0,
    chCy: numAttr(chExt, "cy") ?? 0,
  };
}

const toPx = (emu: number): number => Math.round(emu / EMU_PER_PX);

export function anchorGeom(anchor: Element): AnchorGeom {
  const posH = firstChild(anchor, "wp:positionH");
  const posV = firstChild(anchor, "wp:positionV");
  const extent = firstChild(anchor, "wp:extent");
  const geom: AnchorGeom = {
    x: 0,
    y: 0,
    relH: attr(posH, "relativeFrom") ?? "column",
    relV: attr(posV, "relativeFrom") ?? "paragraph",
    w: toPx(numAttr(extent, "cx") ?? 0),
    h: toPx(numAttr(extent, "cy") ?? 0),
    zOrder: numAttr(anchor, "relativeHeight") ?? 0,
  };
  // `<wp:posOffset>` carries its value as element text, not an attribute.
  const xo = posH ? firstChild(posH, "wp:posOffset") : undefined;
  const yo = posV ? firstChild(posV, "wp:posOffset") : undefined;
  if (xo) geom.x = toPx(Number(text(xo)) || 0);
  if (yo) geom.y = toPx(Number(text(yo)) || 0);
  return geom;
}

const SCHEME_COLORS: Record<string, string> = {
  bg1: "#FFFFFF",
  lt1: "#FFFFFF",
  tx1: "#000000",
  dk1: "#000000",
  bg2: "#E7E6E6",
  lt2: "#E7E6E6",
  dk2: "#444444",
  tx2: "#444444",
  hlink: "#0563C1",
  folHlink: "#954F72",
};

function colorFrom(el: Element | undefined): string | null {
  if (!el) return null;
  const srgb = firstChild(el, "a:srgbClr");
  if (srgb) {
    const v = attr(srgb, "val");
    return v ? `#${v}` : null;
  }
  const sys = firstChild(el, "a:sysClr");
  if (sys) {
    const last = attr(sys, "lastClr");
    return last ? `#${last}` : null;
  }
  const scheme = firstChild(el, "a:schemeClr");
  if (scheme) {
    const v = attr(scheme, "val");
    return (v && SCHEME_COLORS[v]) || null;
  }
  return null;
}

/** Shape fill → a CSS background. `a:noFill` and unmapped colors give null. */
function cssFill(spPr: Element | undefined): string | null {
  if (!spPr) return null;
  if (firstChild(spPr, "a:noFill")) return null;
  const solid = firstChild(spPr, "a:solidFill");
  const solidColor = colorFrom(solid);
  if (solidColor) return solidColor;
  const grad = firstChild(spPr, "a:gradFill");
  if (grad) return cssGradient(grad);
  const blipFill = firstChild(spPr, "a:blipFill");
  if (blipFill) return null; // picture fills are handled by the image path
  return null;
}

/**
 * `a:gradFill` → CSS gradient. Word measures the linear angle clockwise from
 * east in 1/60000 deg; CSS measures clockwise from north in deg.
 */
function cssGradient(grad: Element): string | null {
  const gsLst = firstChild(grad, "a:gsLst");
  if (!gsLst) return null;
  const stops = children(gsLst)
    .filter((c) => c.name === "a:gs")
    .map((gs) => ({ pos: numAttr(gs, "pos") ?? 0, color: colorFrom(gs) }))
    .filter((s): s is { pos: number; color: string } => s.color != null)
    .sort((a, b) => a.pos - b.pos);
  if (!stops.length) return null;
  const lin = firstChild(grad, "a:lin");
  const ang = lin ? (numAttr(lin, "ang") ?? 0) / 60000 - 90 : 90;
  const parts = stops.map(
    (s) => `${s.color} ${Math.round((s.pos / 1000) * 10) / 10}%`,
  );
  return `linear-gradient(${Math.round(ang)}deg, ${parts.join(", ")})`;
}

/** `a:ln` inside `wps:spPr` → a CSS border color and width in px. */
function cssStroke(spPr: Element | undefined): { color: string | null; width: number | null } {
  const ln = spPr ? firstChild(spPr, "a:ln") : undefined;
  if (!ln || firstChild(ln, "a:noFill")) return { color: null, width: null };
  const widthPx = numAttr(ln, "w");
  return {
    color: colorFrom(firstChild(ln, "a:solidFill")),
    width: widthPx ? Math.max(1, Math.round(widthPx / EMU_PER_PX)) : 1,
  };
}

const VALIGN: Record<string, string> = {
  t: "flex-start",
  ctr: "center",
  b: "flex-end",
};

function txbxContentOf(wsp: Element): Element | null {
  const txbx = firstChild(wsp, "wps:txbx");
  return txbx ? (firstChild(txbx, "w:txbxContent") ?? null) : null;
}

/**
 * Every shape drawn by one `<a:graphicData>`, in document order, mapped into
 * the anchor's own coordinate space (px, origin = top-left of `wp:extent`).
 */
export function shapesIn(
  container: Element,
  fallback?: { w: number; h: number },
): ShapeInfo[] {
  const shapes = collectDescendants(container, "wps:wsp");
  if (!shapes.length) return [];

  // A group maps its children through a child-coordinate space: subtract
  // chOff, then scale by ext/chExt. Documents produced by template tools
  // sometimes mis-name the group element, so we key off grpSpPr instead.
  const groupPr = collectDescendants(container, "wpg:grpSpPr")[0];
  const group = groupPr ? readXfrm(firstChild(groupPr, "a:xfrm")) : null;
  const useGroup = !!group && group.chCx > 0 && group.chCy > 0;
  const sx = useGroup ? group.cx / group.chCx : 1;
  const sy = useGroup ? group.cy / group.chCy : 1;

  return shapes.map((wsp) => {
    const spPr = firstChild(wsp, "wps:spPr");
    const x = readXfrm(spPr ? firstChild(spPr, "a:xfrm") : undefined);
    const inGroup = useGroup && (x.cx > 0 || x.cy > 0);
    const px = inGroup ? (x.x - group.chX) * sx + group.x : x.x;
    const py = inGroup ? (x.y - group.chY) * sy + group.y : x.y;
    // Fall back to the anchor's extent only when a shape carries no xfrm of
    // its own; a zero from a real `<a:ext>` (a connector line) stays zero.
    const rawW = inGroup ? x.cx * sx : x.cx;
    const rawH = inGroup ? x.cy * sy : x.cy;
    const pw = x.hasExt ? rawW : rawW || (fallback?.w ?? 0) * EMU_PER_PX;
    const ph = x.hasExt ? rawH : rawH || (fallback?.h ?? 0) * EMU_PER_PX;
    const stroke = cssStroke(spPr);
    const bodyPr = firstChild(wsp, "wps:bodyPr");
    const anchorAttr = attr(bodyPr, "anchor");
    return {
      x: toPx(px),
      y: toPx(py),
      w: toPx(pw),
      h: toPx(ph),
      fill: cssFill(spPr),
      stroke: stroke.color,
      strokeWidth: stroke.width,
      vAlign: (anchorAttr && VALIGN[anchorAttr]) || null,
      txbxContentEl: txbxContentOf(wsp),
      name: attr(firstChild(wsp, "wps:cNvPr"), "name") ?? null,
    };
  });
}

/**
 * Legacy `w:pict` shapes (pre-2010 docs, and the `mc:Fallback` branch Word
 * writes for every group). We read the VML `style` string for position and
 * size. Inside a `v:group` VML expresses offsets as `margin-left`/`margin-top`
 * relative to the parent, which we accumulate; a group's `coordsize` scaling
 * is ignored, so a legacy group can sit a little off — such documents no
 * longer carry the `wps:` form we could measure exactly.
 */
export function shapesFromPict(pict: Element): ShapeInfo[] {
  const out: ShapeInfo[] = [];
  const walk = (el: Element, parentX: number, parentY: number, isRoot: boolean) => {
    const isShape = el.name === "v:shape" || el.name === "v:rect";
    const style = parseVmlStyle(attr(el, "style"));
    const x = isShape ? (isRoot ? style.x : parentX + style.x) : parentX;
    const y = isShape ? (isRoot ? style.y : parentY + style.y) : parentY;
    const textbox = isShape ? firstChild(el, "v:textbox") : undefined;
    if (textbox) {
      out.push({
        x: Math.round(x),
        y: Math.round(y),
        w: Math.round(style.w),
        h: Math.round(style.h),
        fill: vmlColor(attr(el, "fillcolor")),
        stroke: attr(el, "stroked") === "f" ? null : vmlColor(attr(el, "strokecolor")),
        strokeWidth: attr(el, "strokeweight") ? 1 : null,
        vAlign: null,
        txbxContentEl: firstChild(textbox, "w:txbxContent") ?? null,
        name: attr(el, "id") ?? attr(el, "alt") ?? null,
      });
    }
    children(el).forEach((c) => walk(c, x, y, false));
  };
  children(pict).forEach((c) => walk(c, 0, 0, true));
  return out;
}

function parseVmlStyle(style: string | undefined): {
  x: number;
  y: number;
  w: number;
  h: number;
} {
  const read = (prop: string): number => {
    const m = style?.match(new RegExp(`(?:^|[;\\s])${prop}:\\s*(-?[\\d.]+)pt`));
    return m ? ptToPx(Number(m[1])) : 0;
  };
  return {
    x: read("left") || read("margin-left"),
    y: read("top") || read("margin-top"),
    w: read("width"),
    h: read("height"),
  };
}

/** VML colors are either `#RRGGBB` or a name (`white`, `lightgray`). */
const vmlColor = (v: string | undefined): string | null => {
  if (!v || v === "auto" || v === "none") return null;
  if (v.startsWith("#")) return v;
  return /^[0-9a-fA-F]{6}$/.test(v) ? `#${v}` : v;
};

/** CSS px from PostScript points (72 pt/in at 96 dpi). */
export function ptToPx(pt: number): number {
  return Math.round((pt * 96) / 72);
}

export function hasPicture(el: Element): boolean {
  return collectDescendants(el, "a:blip").length > 0;
}

const TXBX_CONTENT = /<w:txbxContent\b[^>]*?(?:\/>|>[\s\S]*?<\/w:txbxContent>)/g;

/**
 * Replace every `<w:txbxContent>` body in a preserved shape XML template with
 * the corresponding entry of `innerXml`, positionally. A `mc:AlternateContent`
 * carries the same shape twice (the `wps:` branch and the VML `mc:Fallback`),
 * both in the same order, so the list cycles when it runs out — that keeps the
 * two branches in step without having to understand which branch we are in.
 */
export function spliceTxbxContents(containerXml: string, innerXml: string[]): string {
  if (!innerXml.length) return containerXml;
  let i = 0;
  return containerXml.replace(TXBX_CONTENT, () => {
    const body = innerXml[i % innerXml.length];
    i++;
    return `<w:txbxContent>${body}</w:txbxContent>`;
  });
}
