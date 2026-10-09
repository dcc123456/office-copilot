import { Extension, Node, mergeAttributes } from "@tiptap/core";
import { Plugin, PluginKey, type Transaction } from "@tiptap/pm/state";
import type { Node as PMNode } from "@tiptap/pm/model";
import type { EditorView, ViewMutationRecord } from "@tiptap/pm/view";
import { FLOAT_PAGES, floatPagesOf } from "./PagedLayout.js";

/**
 * Zero height: the shapes a group holds float over the page and must not push
 * the following lines down.
 */
const GROUP_STYLE = "position:relative;height:0;overflow:visible;";

/**
 * Floating text boxes — the `<wp:anchor>` shapes Word template documents (all
 * resumes, most flyers) use to place text at absolute coordinates.
 *
 * The DOM mirrors how Word positions them, which keeps the text editable as
 * ordinary ProseMirror content:
 *
 *   <div data-float-group>   an in-flow block of zero height, sitting where the
 *                            anchor's host paragraph sat
 *     <div data-float-box>   `position:absolute` at the shape's offset inside
 *                            the anchor, painted with the shape's fill/border
 *       <p>…</p>             the box's text
 *
 * The group carries the anchor's own offset as `--oc-float-dx/-dy`, and each box
 * adds its offset within the anchor, because "the anchor's origin" depends on
 * what `relativeFrom` says and on which page the group currently falls — which
 * is FloatingTextLayout's job to work out.
 *
 * On save, the source shape XML is written back; a box whose text was actually
 * retyped carries a `dirty` flag so only its `<w:txbxContent>` is rebuilt (see
 * packages/docx-io/src/ooxml-pm.ts).
 */

const px = (v: number | null | undefined): string => `${v ?? 0}px`;

/**
 * A connector's outline is the line itself: `<a:ext cx="0">` means a vertical
 * rule, `cy="0"` a horizontal one. Painting a full CSS border on such a box
 * adds caps at both ends and doubles the stroke, so Word's divider turns into a
 * narrow outlined rectangle.
 */
function strokeCss(a: {
  w: number;
  h: number;
  stroke: string | null;
  strokeWidth: number | null;
}): string {
  if (!a.stroke) return "";
  const line = `${px(a.strokeWidth ?? 1)} solid ${a.stroke}`;
  if (a.w <= 0 && a.h > 0) return `border-left:${line}`;
  if (a.h <= 0 && a.w > 0) return `border-top:${line}`;
  return `border:${line}`;
}

/** A box's own attributes, as `renderHTML` and the node view both need them. */
interface BoxAttrs {
  x: number;
  y: number;
  w: number;
  h: number;
  fill: string | null;
  stroke: string | null;
  strokeWidth: number | null;
  vAlign: string | null;
  name: string | null;
}

/**
 * The DOM for `relayout` to clip: a box whose text runs past its page's body
 * band is cut at a line boundary and the rest is continued on the next page.
 * Only that pass writes a clip, and the clip is layout, not document content.
 */
const boxClips = new WeakMap<HTMLElement, (clip: number | null) => void>();

function boxData(a: BoxAttrs): Record<string, string | undefined> {
  return {
    "data-x": String(a.x),
    "data-y": String(a.y),
    "data-w": String(a.w),
    "data-h": String(a.h),
    "data-fill": a.fill ?? undefined,
    "data-stroke": a.stroke ?? undefined,
    "data-stroke-width":
      a.strokeWidth === null ? undefined : String(a.strokeWidth),
    "data-valign": a.vAlign ?? undefined,
    "data-name": a.name ?? undefined,
  };
}

function boxStyle(a: BoxAttrs, clip: number | null = null): string {
  const style = [
    `left:calc(var(--oc-float-dx, 0px) + ${px(a.x)})`,
    `top:calc(var(--oc-float-dy, 0px) + ${px(a.y)})`,
    `width:${px(a.w)}`,
    // Word clips a text box to its height; growing it keeps an edit from
    // hiding text behind the shape below. A shape with no text of its own —
    // a connector line — has nothing to grow with, so it stays exact, else
    // the empty paragraph that keeps it selectable would push it open.
    a.h > 0
      ? `min-height:${px(a.h)};overflow:visible`
      : "height:0;overflow:hidden",
    a.fill ? `background:${a.fill}` : "",
    strokeCss(a),
    // Vertical centering only makes sense against the shape's own height; a
    // clipped box would push its remaining lines down.
    a.vAlign && clip === null
      ? `display:flex;flex-direction:column;justify-content:${a.vAlign}`
      : "",
    "position:absolute",
    "box-sizing:border-box",
    // Word insets a text box by 0.1cm on the sides, 0 on top and bottom.
    "padding:0 5px",
  ]
    .filter(Boolean)
    .join(";");

  // `min-height` has to go before `height` can bound anything, and the cut edge
  // must hide rather than show the line it lands inside.
  return clip === null
    ? style
    : `${style};min-height:0;height:${px(clip)};overflow:hidden`;
}

function numData(name: string, fallback: number) {
  return (element: HTMLElement): number => {
    const raw = element.getAttribute(name);
    if (raw === null) return fallback;
    const n = Number(raw);
    return Number.isFinite(n) ? n : fallback;
  };
}

function dataString(name: string, fallback: string) {
  return (element: HTMLElement): string =>
    element.getAttribute(name) ?? fallback;
}

function stringData(name: string) {
  return (element: HTMLElement): string | null => element.getAttribute(name);
}

function nullableNumData(name: string) {
  return (element: HTMLElement): number | null => {
    const raw = element.getAttribute(name);
    const n = raw === null ? NaN : Number(raw);
    return Number.isFinite(n) ? n : null;
  };
}

export const FloatBox = Node.create({
  name: "floatBox",
  group: "block",
  content: "block+",

  addAttributes() {
    // `rendered: false` everywhere: the DOM carries `data-*` names, which
    // `renderHTML` writes itself, so the raw attribute names must not leak into
    // the DOM as invalid bare attributes like `w="205"`.
    return {
      x: { default: 0, rendered: false, parseHTML: numData("data-x", 0) },
      y: { default: 0, rendered: false, parseHTML: numData("data-y", 0) },
      w: { default: 0, rendered: false, parseHTML: numData("data-w", 0) },
      h: { default: 0, rendered: false, parseHTML: numData("data-h", 0) },
      fill: {
        default: null,
        rendered: false,
        parseHTML: stringData("data-fill"),
      },
      stroke: {
        default: null,
        rendered: false,
        parseHTML: stringData("data-stroke"),
      },
      strokeWidth: {
        default: null,
        rendered: false,
        parseHTML: nullableNumData("data-stroke-width"),
      },
      vAlign: {
        default: null,
        rendered: false,
        parseHTML: stringData("data-valign"),
      },
      name: {
        default: null,
        rendered: false,
        parseHTML: stringData("data-name"),
      },
      // The box's original `<w:txbxContent>` body travels only through the JSON
      // document, never the DOM.
      txbx: { default: null, rendered: false },
      dirty: { default: false, rendered: false },
    };
  },

  parseHTML() {
    return [{ tag: "div[data-float-box]" }];
  },

  renderHTML({ node, HTMLAttributes }) {
    const a = node.attrs as BoxAttrs;
    return [
      "div",
      mergeAttributes(HTMLAttributes, {
        "data-float-box": "",
        ...boxData(a),
        title: a.name ?? undefined,
        style: boxStyle(a),
      }),
      0,
    ];
  },

  /**
   * A node view, because a box whose text outgrows its page has to be clipped
   * at a line boundary, and that height is layout the plugin works out — not
   * something ProseMirror should read back as an edit. `boxClips` hands the
   * element to `relayout`, which is the only writer of the clip.
   */
  addNodeView() {
    return ({ node, HTMLAttributes }) => {
      const dom = document.createElement("div");
      let current = node;
      let clip: number | null = null;
      const write = () => {
        const a = current.attrs as BoxAttrs;
        for (const [k, v] of Object.entries(boxData(a))) {
          if (v === undefined) dom.removeAttribute(k);
          else dom.setAttribute(k, v);
        }
        dom.setAttribute("style", boxStyle(a, clip));
      };
      dom.setAttribute("data-float-box", "");
      for (const [k, v] of Object.entries(HTMLAttributes)) {
        if (typeof v === "string") dom.setAttribute(k, v);
      }
      write();
      boxClips.set(dom, (next) => {
        if (next === clip) return;
        clip = next;
        write();
      });
      return {
        dom,
        contentDOM: dom,
        ignoreMutation: (mutation: ViewMutationRecord) =>
          mutation.type === "attributes" && mutation.target === dom,
        update: (next: PMNode) => {
          if (next.type !== current.type) return false;
          current = next;
          write();
          return true;
        },
        destroy: () => {
          boxClips.delete(dom);
        },
      };
    };
  },
});

export const FloatGroup = Node.create({
  name: "floatGroup",
  group: "block",
  content: "floatBox*",

  addAttributes() {
    return {
      ooxml: { default: "", rendered: false },
      x: { default: 0, rendered: false, parseHTML: numData("data-anchor-x", 0) },
      y: { default: 0, rendered: false, parseHTML: numData("data-anchor-y", 0) },
      relH: {
        default: "column",
        rendered: false,
        parseHTML: dataString("data-rel-h", "column"),
      },
      relV: {
        default: "paragraph",
        rendered: false,
        parseHTML: dataString("data-rel-v", "paragraph"),
      },
      z: { default: 0, rendered: false, parseHTML: numData("data-z", 0) },
    };
  },

  parseHTML() {
    return [{ tag: "div[data-float-group]" }];
  },

  /**
   * A node view, because FloatingTextLayout writes the anchor offsets into this
   * element's inline style every frame. ProseMirror reads foreign attribute
   * writes as an edit and re-parses the range, which rebuilds the element from
   * `renderHTML` — wiping the offsets and, worse, dropping `ooxml` and `txbx`,
   * which never appear in the DOM. Owning the element stops both.
   */
  addNodeView() {
    return ({ node }) => {
      const dom = document.createElement("div");
      dom.setAttribute("data-float-group", "");
      dom.style.cssText = GROUP_STYLE;

      const write = (target: PMNode) => {
        const a = target.attrs as {
          x: number;
          y: number;
          relH: string;
          relV: string;
          z: number;
        };
        dom.dataset.anchorX = String(a.x);
        dom.dataset.anchorY = String(a.y);
        dom.dataset.relH = a.relH;
        dom.dataset.relV = a.relV;
        dom.dataset.z = String(a.z);
        dom.style.zIndex = String(a.z);
      };
      write(node);

      return {
        dom,
        contentDOM: dom,
        ignoreMutation: (mutation: ViewMutationRecord) =>
          mutation.type === "attributes" && mutation.target === dom,
        update: (next: PMNode) => {
          if (next.type !== node.type) return false;
          write(next);
          return true;
        },
      };
    };
  },

  renderHTML({ node, HTMLAttributes }) {
    const a = node.attrs as {
      x: number;
      y: number;
      relH: string;
      relV: string;
      z: number;
    };
    return [
      "div",
      mergeAttributes(HTMLAttributes, {
        "data-float-group": "",
        "data-anchor-x": String(a.x),
        "data-anchor-y": String(a.y),
        "data-rel-h": a.relH,
        "data-rel-v": a.relV,
        "data-z": String(a.z),
        style: `${GROUP_STYLE}z-index:${a.z};`,
      }),
      0,
    ];
  },
});

interface PageGeom {
  /** Distance between two page tops, in layout px. */
  pitch: number;
  /** Usable body band per page, in layout px. */
  band: number;
  marginTop: number;
  marginLeft: number;
  /** Stage transform scale, so client rects can be turned into layout px. */
  scale: number;
}

/** Geometry comes from the same CSS variables PagedLayout reads. */
function readPageGeom(dom: HTMLElement): PageGeom {
  const cs = getComputedStyle(dom);
  const num = (value: string, fallback: number): number => {
    const n = parseFloat(value);
    return Number.isFinite(n) ? n : fallback;
  };
  const pageHeight = num(cs.getPropertyValue("--oc-page-height"), 1123);
  const contentHeight = num(cs.getPropertyValue("--oc-content-height"), 931);
  const marginTop = num(cs.getPropertyValue("--oc-margin-top"), 96);
  const marginBottom = num(cs.getPropertyValue("--oc-margin-bottom"), 96);
  const pageGap = num(cs.getPropertyValue("--oc-page-gap"), 24);
  const usedWidth = parseFloat(cs.width);
  return {
    // Fall back to the content band plus both margins when the host has not
    // published a page height yet.
    pitch:
      pageHeight > 0
        ? pageHeight + pageGap
        : contentHeight + marginTop + marginBottom + pageGap,
    band: contentHeight,
    marginTop,
    marginLeft: num(cs.getPropertyValue("--oc-margin-left"), 96),
    // The stage is zoomed with a CSS transform, so rects are screen px while
    // every measurement here is layout px.
    scale: usedWidth > 0 ? dom.getBoundingClientRect().width / usedWidth : 1,
  };
}

/** Kept clear of a page's bottom edge, the same margin PagedLayout gives body lines. */
const SLACK = 2;
const EPS = 0.5;
/** The layer a box's continuations are painted in — outside the editable doc. */
const SPLIT_CLASS = "oc-float-split";

/**
 * Content offsets, in layout px below the top of a box, at which its text is
 * cut so that every part fits one page's body band. `lines` are the bottom
 * edges of the box's painted lines, ascending, also from the box's own top.
 *
 * Returns [] when not even the next line fits the room left on the box's page:
 * there is no line boundary to cut at, so the box stays as Word draws it.
 */
export function splitCuts(
  lines: number[],
  firstAvail: number,
  band: number,
  maxParts = 24,
): number[] {
  const cuts: number[] = [];
  const last = lines.length ? lines[lines.length - 1]! : 0;
  let from = 0;
  let room = firstAvail;
  for (let i = 0; i < maxParts && last > from + room + EPS; i++) {
    let cut = from;
    for (const bottom of lines) {
      if (bottom > from + room) break;
      cut = bottom;
    }
    if (cut <= from) return [];
    cuts.push(cut);
    from = cut;
    room = band;
  }
  return cuts;
}

/**
 * Bottom edges of every painted line inside `el`, in layout px below its top.
 * A range over the whole content reports one client rect per line fragment,
 * and a line is the smallest thing that can move to the next page whole.
 */
function lineBottoms(
  el: HTMLElement,
  origin: number,
  top: number,
  scale: number,
): number[] {
  const range = document.createRange();
  range.selectNodeContents(el);
  const rects = Array.from(range.getClientRects()).sort(
    (a, b) => a.bottom - b.bottom,
  );
  const out: number[] = [];
  let prev = Number.NaN;
  for (const rect of rects) {
    if (rect.width <= 0 || rect.height <= 0) continue;
    // Half-pixel buckets: a line broken into several fragments reports the same
    // bottom a fraction apart, and each duplicate would become a cut.
    const bottom = Math.round(((rect.bottom - origin) / scale - top) * 2) / 2;
    if (bottom !== prev) out.push(bottom);
    prev = bottom;
  }
  return out;
}

interface SplitPart {
  left: number;
  /** Layout px from the top of page 1's band. */
  top: number;
  width: number;
  height: number;
  /** How far the box's own coordinates the part begins at. */
  offset: number;
  box: HTMLElement;
}

/**
 * Paint the continuations. ProseMirror gives a node exactly one DOM element, so
 * the parts of a split box cannot be document content — they are a copy in a
 * layer the editor never reads.
 */
function renderParts(host: HTMLElement, parts: SplitPart[]): void {
  let layer = host.querySelector<HTMLElement>(`:scope > .${SPLIT_CLASS}`);
  if (!parts.length) {
    if (layer) layer.replaceChildren();
    return;
  }
  if (!layer) {
    layer = document.createElement("div");
    layer.className = SPLIT_CLASS;
    // Zero size, at the host's own origin: the parts place themselves from the
    // top of page 1's band, in the same coordinates `relayout` uses.
    layer.style.cssText =
      "position:absolute;left:0;top:0;width:0;height:0;pointer-events:none;z-index:3";
    host.appendChild(layer);
  }

  const fragment = document.createDocumentFragment();
  for (const p of parts) {
    const wrap = document.createElement("div");
    wrap.className = "oc-float-part";
    wrap.style.cssText = `position:absolute;left:${px(p.left)};top:${px(
      p.top,
    )};width:${px(p.width)};height:${px(
      p.height,
    )};overflow:hidden;pointer-events:none`;

    const clone = p.box.cloneNode(true) as HTMLElement;
    // A page-gap widget would belong to the original block's pagination; a
    // continuation has none of its own.
    clone
      .querySelectorAll<HTMLElement>(".oc-page-gap")
      .forEach((el) => el.remove());
    const cs = getComputedStyle(p.box);
    // The shape continues with its text: a box Word fills dark holds white
    // text, so a bare continuation would paint it onto the page and lose it.
    // `min-height` is what the shape's own height was, and the cut edge is the
    // page's, so neither carries over.
    clone.style.cssText =
      `position:absolute;left:0;top:${px(-p.offset)};width:${px(p.width)};` +
      "box-sizing:border-box;min-height:0;padding:0 5px;overflow:visible;" +
      `font-family:${cs.fontFamily};font-size:${cs.fontSize};` +
      `line-height:${cs.lineHeight};color:${cs.color};` +
      `background:${cs.background};border-top:${cs.borderTop};` +
      `border-right:${cs.borderRight};border-bottom:${cs.borderBottom};` +
      `border-left:${cs.borderLeft};border-radius:${cs.borderRadius}`;

    wrap.appendChild(clone);
    fragment.appendChild(wrap);
  }
  layer.replaceChildren(fragment);
}

/**
 * Keep every floating box's text inside a page's body band: cut the boxes that
 * outgrow it at a line boundary, continue what falls below on the next pages,
 * and ask for as many pages as the continuations need.
 */
function applySplits(
  view: EditorView,
  geom: PageGeom,
  host: HTMLElement,
  origin: number,
  originLeft: number,
): void {
  const boxes = Array.from(
    (view.dom as HTMLElement).querySelectorAll<HTMLElement>("[data-float-box]"),
  );
  // Read the lines with every clip gone. A clipped-out line is a line the range
  // may stop reporting, and cutting on that reading would shorten the box a
  // little more every frame until its text disappeared.
  for (const el of boxes) boxClips.get(el)?.(null);

  const parts: SplitPart[] = [];
  const clips: { el: HTMLElement; cut: number }[] = [];
  let needed = 1;

  for (const el of boxes) {
    const rect = el.getBoundingClientRect();
    const width = rect.width / geom.scale;
    const top = (rect.top - origin) / geom.scale;
    let page = Math.max(0, Math.floor(top / geom.pitch));
    // A box that begins below a page's band — in its bottom margin — is
    // measured against the band it actually runs into.
    if (top - page * geom.pitch > geom.band) page += 1;
    const avail = page * geom.pitch + geom.band - SLACK - top;
    if (avail <= SLACK) continue;

    const lines = lineBottoms(el, origin, top, geom.scale);
    if (!lines.length) continue;
    const cuts = splitCuts(lines, avail, geom.band - SLACK);
    if (!cuts.length) continue;

    const lastLine = lines[lines.length - 1]!;
    for (let k = 1; k <= cuts.length; k++) {
      const offset = cuts[k - 1]!;
      const end = k < cuts.length ? cuts[k]! : lastLine;
      parts.push({
        left: (rect.left - originLeft) / geom.scale,
        top: (page + k) * geom.pitch + SLACK,
        width,
        height: Math.min(end - offset, geom.band - SLACK),
        offset,
        box: el,
      });
    }
    clips.push({ el, cut: cuts[0]! });
    needed = Math.max(needed, page + cuts.length + 1);
  }

  renderParts(host, parts);
  for (const c of clips) boxClips.get(c.el)?.(c.cut);

  if (needed !== floatPagesOf(view.state)) {
    view.dispatch(
      view.state.tr
        .setMeta(FLOAT_PAGES, needed)
        .setMeta("addToHistory", false),
    );
  }
}

/**
 * Publish, per group, the shift from where the group sits in the text flow to
 * where Word put its anchor. Boxes pick it up through `calc()` in their own
 * `left`/`top`, so nothing has to be written into the boxes themselves.
 */
function relayout(view: EditorView): void {
  const dom = view.dom as HTMLElement;
  const groups = Array.from(dom.querySelectorAll<HTMLElement>("[data-float-group]"));

  // Page 1's content band starts at the host's top edge; the editor is a single
  // long column, so page p's band starts p * pitch below it.
  const host = dom.closest<HTMLElement>(".prosemirror-host") ?? dom;
  const geom = readPageGeom(dom);
  const hostRect = host.getBoundingClientRect();
  const origin = hostRect.top;

  for (const el of groups) {
    const relV = el.dataset.relV ?? "paragraph";
    const relH = el.dataset.relH ?? "column";
    const x = Number(el.dataset.anchorX ?? 0);
    const y = Number(el.dataset.anchorY ?? 0);
    const flowTop = (el.getBoundingClientRect().top - origin) / geom.scale;
    const bandTop = Math.max(0, Math.floor(flowTop / geom.pitch)) * geom.pitch;

    // `relativeFrom="page"` measures from the paper, one top margin above the
    // band; `column` and `margin` both measure from the band's own edges, which
    // is where the host already begins.
    let anchorTop: number;
    if (relV === "paragraph") anchorTop = flowTop + y;
    else if (relV === "page") anchorTop = bandTop - geom.marginTop + y;
    else anchorTop = bandTop + y;

    const dx = relH === "page" ? x - geom.marginLeft : x;

    el.style.setProperty("--oc-float-dx", px(dx));
    el.style.setProperty("--oc-float-dy", px(anchorTop - flowTop));
  }

  // After the anchors, so the boxes are measured where they now sit.
  applySplits(view, geom, host, origin, hostRect.left);
}

export const FloatingTextLayout = Extension.create({
  name: "floatingTextLayout",

  addProseMirrorPlugins() {
    const collectBoxes = (doc: PMNode): { node: PMNode; pos: number }[] => {
      const boxes: { node: PMNode; pos: number }[] = [];
      doc.descendants((node, pos) => {
        if (node.type.name === "floatBox") boxes.push({ node, pos });
        return true;
      });
      return boxes;
    };

    return [
      new Plugin({
        key: new PluginKey("floatingTextLayout"),
        view: (view) => {
          let queued = false;
          const schedule = () => {
            if (queued) return;
            queued = true;
            requestAnimationFrame(() => {
              queued = false;
              relayout(view);
            });
          };
          const ro = new ResizeObserver(schedule);
          ro.observe(view.dom);
          window.addEventListener("resize", schedule);
          view.dom.addEventListener("oc:pagecount", schedule);
          document.fonts?.ready.then(schedule).catch(() => undefined);
          schedule();
          return {
            // Any transaction can move a group: an edit above it, or a page gap
            // PagedLayout inserted, changes which page it falls on.
            update: schedule,
            destroy: () => {
              ro.disconnect();
              window.removeEventListener("resize", schedule);
              view.dom.removeEventListener("oc:pagecount", schedule);
            },
          };
        },
      }),
      new Plugin({
        key: new PluginKey("floatBoxDirty"),
        appendTransaction(trs, oldState, newState) {
          if (!trs.some((tr) => tr.docChanged)) return null;
          const before = collectBoxes(oldState.doc);
          const after = collectBoxes(newState.doc);
          // A change in the number of boxes means boxes were added or removed,
          // so the positional mapping onto the saved template is gone.
          if (before.length !== after.length) return null;
          let tr: Transaction | null = null;
          for (let i = 0; i < before.length; i++) {
            const old = before[i]!;
            const next = after[i]!;
            if (next.node.attrs.dirty || next.node.content.eq(old.node.content)) {
              continue;
            }
            tr = tr ?? newState.tr.setMeta("addToHistory", false);
            tr.setNodeMarkup(next.pos, undefined, {
              ...next.node.attrs,
              dirty: true,
            });
          }
          return tr && tr.docChanged ? tr : null;
        },
      }),
    ];
  },
});
