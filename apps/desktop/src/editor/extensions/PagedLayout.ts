import { Extension } from "@tiptap/core";
import { Plugin, PluginKey, TextSelection, type EditorState } from "@tiptap/pm/state";
import { Decoration, DecorationSet } from "@tiptap/pm/view";
import type { Node as PMNode } from "@tiptap/pm/model";

/**
 * PagedLayout paginates the editing view like Word: every top-level block is
 * kept inside its page's content band, a block taller than that band is
 * continued on the next page at a line boundary, and an explicit page break
 * starts the rest of its block on the next page even when it would have fit.
 *
 * Geometry comes from CSS variables the host writes on the editor DOM:
 *   --oc-page-height     full page height (px, unscaled layout px)
 *   --oc-content-height  usable body band per page (px)
 *   --oc-page-gap        gray gap between two pages (px)
 *   --oc-margin-top/-bottom
 *
 * Page p's band is [p*pitch, p*pitch + contentHeight] measured from the top of
 * the editor's content box, where pitch = pageHeight + pageGap. Because a
 * spacer only ever pushes content down, a spacer must absorb the space the
 * previous page left unused: height = targetTop - currentTop. That is the
 * difference from a fixed spacer height, which assumes every page fills
 * exactly and drifts one page over whenever a block ends early.
 *
 * Planning happens in *natural* coordinates — the layout with no spacers at
 * all — and is read with the existing spacers hidden. That hiding cannot be
 * replaced by subtracting the rendered gaps: a spacer inside a paragraph forces
 * a line break there, so the rest of the paragraph wraps differently than it
 * would without it, and a subtractive reading settles on a layout that the plan
 * it produces does not reproduce.
 *
 * Because that read mutates the DOM, it never runs while the mouse button is
 * down, and the transaction that follows puts back any selection it collapsed.
 */

const KEY = new PluginKey<PagedState>("pagedLayout");
const EPS = 0.5;
const MAX_BREAKS_PER_BLOCK = 256;
/** Breathing room kept clear of the header and footer bands, in layout px. */
const SLACK = 2;

/**
 * Meta by which another plugin asks for more pages than the text flow needs —
 * a floating box whose continuation reaches past the last page.
 */
export const FLOAT_PAGES = new PluginKey("floatPages");

/** How many pages the floating shapes currently require. */
export function floatPagesOf(state: EditorState): number {
  return KEY.getState(state)?.floatPages ?? 1;
}

const DEFAULTS = {
  pageHeight: 1123,
  contentHeight: 931,
  pageGap: 24,
  marginTop: 96,
  marginBottom: 96,
};

type EditorViewLike = import("@tiptap/pm/view").EditorView;

interface PlannedBreak {
  /** Document position the spacer is inserted at. */
  pos: number;
  /** Spacer height in layout px. */
  height: number;
  /** True for a break inside a block's inline content. */
  inline: boolean;
}

interface PagedState {
  deco: DecorationSet;
  /** Pages the text flow needs, on its own, without the floating shapes. */
  pages: number;
  /** Pages the floating shapes need. The host is shown the larger of the two. */
  floatPages: number;
  /** Page count already reported to the host, so a re-plan never re-announces. */
  announced: number;
  breaks: PlannedBreak[];
}

interface Geom {
  /** Usable body band per page. */
  band: number;
  /** Distance between two page tops. */
  pitch: number;
  /** Stage transform scale, so client rects can be turned into layout px. */
  scale: number;
  /** Client-space y that layout px 0 refers to (top of page 1's band). */
  origin: number;
}

interface ForcedBreak {
  /** Doc position of the page-break node. */
  pos: number;
  /** Natural top of the break line, layout px from content top. */
  top: number;
  /**
   * Natural top of whatever the break moves to the next page. The break line is
   * an editing affordance with margins of its own, so this — not `top` — is the
   * edge that has to land on the next page, or every page after the break starts
   * one line late.
   */
  resume: number;
}

interface Unit {
  node: PMNode;
  /** Doc position of the node; also where a block-level spacer goes. */
  pos: number;
  el: HTMLElement;
  /** Border-box top/bottom in natural layout px, relative to content top. */
  top: number;
  bottom: number;
  marginTop: number;
  marginBottom: number;
  splittable: boolean;
  /** Explicit page breaks inside this block, in natural top order. */
  forced: ForcedBreak[];
}

/**
 * The explicit page breaks inside one block — a `w:br w:type="page"` from the
 * source, or one the toolbar inserted. Their tops come from the same
 * spacer-free pass as the block's own, so the plan can drop the break line at
 * the bottom of its page and let the rest of the block continue below it.
 */
function forcedBreaksIn(
  el: HTMLElement,
  view: EditorViewLike,
  geom: Geom,
): { pos: number; top: number }[] {
  const spans = el.querySelectorAll<HTMLElement>("span[data-pagebreak]");
  if (!spans.length) return [];
  const out: { pos: number; top: number }[] = [];
  for (const span of Array.from(spans)) {
    let pos: number;
    try {
      pos = view.posAtDOM(span, 0);
    } catch {
      continue;
    }
    if (pos <= 0) continue;
    out.push({
      pos,
      top: (span.getBoundingClientRect().top - geom.origin) / geom.scale,
    });
  }
  return out.sort((a, b) => a.top - b.top);
}

export const PagedLayout = Extension.create({
  name: "pagedLayout",

  addProseMirrorPlugins() {
    return [
      new Plugin<PagedState>({
        key: KEY,
        state: {
          init: (): PagedState => ({
            deco: DecorationSet.empty,
            pages: 1,
            floatPages: 1,
            announced: 0,
            breaks: [],
          }),
          apply(tr, prev): PagedState {
            const meta = tr.getMeta(KEY) as PagedState | undefined;
            if (meta) return meta;
            const floatPages = tr.getMeta(FLOAT_PAGES) as number | undefined;
            if (typeof floatPages === "number" && floatPages !== prev.floatPages) {
              return { ...prev, floatPages };
            }
            if (tr.docChanged) {
              return { ...prev, deco: prev.deco.map(tr.mapping, tr.doc) };
            }
            return prev;
          },
        },
        props: {
          decorations: (s) => KEY.getState(s)?.deco ?? DecorationSet.empty,
        },
        view: (view) => {
          let pending = false;
          let lastWidth = -1;
          const measure = () => {
            if (pending) return;
            pending = true;
            requestAnimationFrame(() => {
              pending = false;
              runMeasure(view);
            });
          };

          // The host draws one card per page, so it only learns the count from
          // here — whichever of the two page needs last changed.
          const announce = () => {
            const state = KEY.getState(view.state);
            if (!state) return;
            const pages = Math.max(1, state.pages, state.floatPages);
            if (pages === state.announced) return;
            view.dispatch(
              view.state.tr
                .setMeta(KEY, { ...state, announced: pages })
                .setMeta("addToHistory", false),
            );
            view.dom.dispatchEvent(
              new CustomEvent("oc:pagecount", {
                bubbles: true,
                detail: { pages, breaks: state.breaks.map((b) => b.pos) },
              }),
            );
          };

          measure();

          // A drag must not be interrupted by the measuring pass, so re-plan
          // only once the button is back up.
          let buttonDown = false;
          let deferred = false;
          const down = () => {
            buttonDown = true;
          };
          const up = () => {
            buttonDown = false;
            if (deferred) {
              deferred = false;
              measure();
            }
          };
          const guarded = () => {
            if (buttonDown) {
              deferred = true;
              return;
            }
            measure();
          };

          // Only a width change is worth reacting to here: height changes are
          // either caused by our own spacers or follow a document change, which
          // the plugin view already reports.
          const ro = new ResizeObserver((entries) => {
            const rect = entries[entries.length - 1]?.contentRect;
            if (!rect) return;
            if (Math.abs(rect.width - lastWidth) < 0.5) return;
            lastWidth = rect.width;
            guarded();
          });
          ro.observe(view.dom);
          view.dom.addEventListener("load", guarded, true);
          view.dom.addEventListener("mousedown", down);
          window.addEventListener("mouseup", up);
          window.addEventListener("blur", up);
          if (document.fonts) {
            document.fonts.ready.then(guarded).catch(() => undefined);
          }

          return {
            update: (v, prev) => {
              announce();
              if (!prev || prev.doc !== v.state.doc) guarded();
            },
            destroy: () => {
              ro.disconnect();
              view.dom.removeEventListener("load", guarded, true);
              view.dom.removeEventListener("mousedown", down);
              window.removeEventListener("mouseup", up);
              window.removeEventListener("blur", up);
            },
          };
        },
      }),
    ];
  },

  onCreate() {
    this.editor.view.dom.classList.add("paged-doc");
  },
});

/**
 * Client-space top of page 1's content band.
 *
 * This is the host wrapper, not the editor's own border box: the first block's
 * margin-top collapses out of both the editor and the wrapper, so the editor
 * sits a margin below the band it is meant to fill. Measuring from the editor
 * would slide every page band down by that margin and let the body run into the
 * footer.
 */
function bandTopClient(dom: HTMLElement): number {
  const host = dom.closest<HTMLElement>(".prosemirror-host");
  return (host ?? dom).getBoundingClientRect().top;
}

function readGeom(dom: HTMLElement): Geom {
  const cs = getComputedStyle(dom);
  const num = (v: string, fallback: number): number => {
    const n = parseFloat(v);
    return Number.isFinite(n) ? n : fallback;
  };
  const contentHeight = num(
    cs.getPropertyValue("--oc-content-height"),
    DEFAULTS.contentHeight,
  );
  const pageHeight = num(
    cs.getPropertyValue("--oc-page-height"),
    DEFAULTS.pageHeight,
  );
  const pageGap = num(cs.getPropertyValue("--oc-page-gap"), DEFAULTS.pageGap);
  const marginTop = num(
    cs.getPropertyValue("--oc-margin-top"),
    DEFAULTS.marginTop,
  );
  const marginBottom = num(
    cs.getPropertyValue("--oc-margin-bottom"),
    DEFAULTS.marginBottom,
  );
  // The stage is scaled with a CSS transform, so client rects are screen px
  // while every band here is layout px. getComputedStyle().width reports the
  // used (untransformed) width, which gives an exact scale factor.
  const usedWidth = parseFloat(cs.width);
  const rectWidth = dom.getBoundingClientRect().width;
  const scale = usedWidth > 0 ? rectWidth / usedWidth : 1;
  return {
    band: contentHeight,
    // pageHeight already equals band + margins; fall back to the margin sum if
    // the host has not published --oc-page-height yet.
    pitch:
      pageHeight > 0
        ? pageHeight + pageGap
        : contentHeight + marginTop + marginBottom + pageGap,
    scale,
    origin: bandTopClient(dom),
  };
}

/**
 * Measure top-level blocks with every spacer hidden, so tops and heights are
 * the natural layout. Lists contribute one unit per item, which lets a list
 * break between items instead of only before the whole list.
 *
 * Returns null when a block has no DOM to read — a view that is still mounting,
 * or one already torn down. A partial read would price that block at nothing and
 * the plan built from it would drop the spacers holding the following pages
 * down, so the whole pass is worthless.
 */
function collectUnits(view: EditorViewLike, geom: Geom): Unit[] | null {
  const toPx = (v: number) => v / geom.scale;
  const toY = (clientTop: number) => toPx(clientTop - geom.origin);
  const units: Unit[] = [];
  let unread = 0;

  const push = (node: PMNode, pos: number) => {
    let dom: HTMLElement | null;
    try {
      dom = view.nodeDOM(pos) as HTMLElement | null;
    } catch {
      dom = null;
    }
    if (!dom || dom.nodeType !== Node.ELEMENT_NODE) {
      unread++;
      return;
    }
    const el = dom;
    const cs = getComputedStyle(el);
    const rect = el.getBoundingClientRect();
    // The stage is scaled, so rects are screen px; convert to layout px.
    const top = toY(rect.top);
    const bottom = top + rect.height / geom.scale;
    const lineHeight =
      parseFloat(cs.lineHeight) || (parseFloat(cs.fontSize) || 14) * 1.2;
    units.push({
      node,
      pos,
      el,
      top,
      bottom,
      marginTop: parseFloat(cs.marginTop) || 0,
      marginBottom: parseFloat(cs.marginBottom) || 0,
      splittable: node.type.name === "paragraph" || node.type.name === "heading",
      forced: forcedBreaksIn(el, view, geom).map((f) => ({
        pos: f.pos,
        top: f.top,
        // Content resumes one line below the break, or at the block's own edge
        // when the break ends the block.
        resume:
          f.top + lineHeight >= bottom - lineHeight
            ? bottom
            : f.top + lineHeight,
      })),
    });
  };

  view.state.doc.forEach((node, pos) => {
    if (node.type.name === "bulletList" || node.type.name === "orderedList") {
      node.forEach((item, inner) => push(item, pos + 1 + inner));
    } else {
      push(node, pos);
    }
  });
  return unread ? null : units;
}

interface SplitPoint {
  /** Inline doc position, just before the first overflowing character. */
  pos: number;
  /** Natural top of the overflowing line, layout px from content top. */
  top: number;
}

/**
 * Find the first line inside `unit` whose bottom crosses `limit`, and return
 * the document position right before it. Returns null when even the first line
 * does not fit, which means the block has to move whole.
 */
function findLineSplit(
  view: EditorViewLike,
  unit: Unit,
  fromTop: number,
  limit: number,
  geom: Geom,
): SplitPoint | null {
  const nodes: Text[] = [];
  const starts: number[] = [];
  let total = 0;
  const walker = document.createTreeWalker(unit.el, NodeFilter.SHOW_TEXT);
  while (walker.nextNode()) {
    const text = walker.currentNode as Text;
    if (!text.data.length) continue;
    starts.push(total);
    nodes.push(text);
    total += text.data.length;
  }
  if (!nodes.length) return null;

  const domTop = geom.origin;
  const toY = (clientTop: number) => (clientTop - domTop) / geom.scale;

  const lineAt = (
    index: number,
  ): { top: number; bottom: number; node: Text; offset: number } | null => {
    // Walk forward a few characters when the offset sits on whitespace that
    // produces no client rect of its own.
    for (let probe = index; probe < Math.min(index + 4, total); probe++) {
      let lo = 0;
      while (lo + 1 < starts.length && starts[lo + 1]! <= probe) lo += 1;
      const node = nodes[lo]!;
      const offset = probe - starts[lo]!;
      if (offset >= node.data.length) continue;
      const range = document.createRange();
      range.setStart(node, offset);
      range.setEnd(node, offset + 1);
      const rects = range.getClientRects();
      if (!rects.length) continue;
      const rect = rects[0]!;
      return {
        top: toY(rect.top),
        bottom: toY(rect.bottom),
        node,
        offset,
      };
    }
    return null;
  };

  // The first character whose line bottom passes the limit. Line bottoms are
  // non-decreasing along the text, so this is a monotone predicate.
  let low = 0;
  let high = total - 1;
  let found = -1;
  while (low <= high) {
    const mid = (low + high) >> 1;
    const line = lineAt(mid);
    if (line && line.bottom > limit) {
      found = mid;
      high = mid - 1;
    } else {
      low = mid + 1;
    }
  }
  if (found < 0) return null;
  const line = lineAt(found);
  if (!line || line.top <= fromTop + EPS) return null;

  const pos = view.posAtDOM(line.node, line.offset);
  if (pos <= unit.pos || pos >= unit.pos + unit.node.nodeSize) return null;
  return { pos, top: line.top };
}

interface Plan {
  breaks: PlannedBreak[];
  pages: number;
}

function planLayout(view: EditorViewLike, geom: Geom): Plan | null {
  const units = collectUnits(view, geom);
  if (!units) return null;
  const { band, pitch } = geom;
  const breaks: PlannedBreak[] = [];
  if (!units.length) return { breaks, pages: 1 };

  // Total spacer height inserted above the cursor, converting natural y into
  // live y.
  let acc = 0;
  let page = 0;
  let pages = 1;

  for (let i = 0; i < units.length; i++) {
    const unit = units[i]!;
    const previous = units[i - 1];
    // A spacer sibling stops the two blocks' margins from collapsing, so the
    // block lands that much lower than the natural layout suggested.
    const collapse = previous
      ? Math.min(previous.marginBottom, unit.marginTop)
      : unit.marginTop;

    let segTop = unit.top;
    let anchor = unit.pos;
    let isFirstSegment = true;
    let guard = 0;
    // Index into `unit.forced`: each break is consumed once the layout reaches
    // the line it sits on.
    let fi = 0;

    while (guard++ < MAX_BREAKS_PER_BLOCK) {
      // A text rect sticks a little outside its line box (CJK metrics), so keep
      // a small band inside each page edge to stay clear of the header/footer.
      const pageTop = page * pitch + SLACK;
      const pageBottom = page * pitch + band - SLACK;
      const liveTop = segTop + acc;
      const liveBottom = unit.bottom + acc;
      const nextForced = unit.forced[fi];
      const forcedPending = nextForced !== undefined && nextForced.top < unit.bottom - EPS;

      if (liveBottom <= pageBottom + EPS && !forcedPending) break;

      // The segment sits above this page's band (a previous block spilled over
      // it): push it down to the band top.
      if (liveTop < pageTop - EPS) {
        const height = pageTop - liveTop - (isFirstSegment ? collapse : 0);
        if (height > EPS) {
          breaks.push({ pos: anchor, height, inline: !isFirstSegment });
          acc += height;
          continue;
        }
        page = Math.max(0, Math.floor(liveTop / pitch));
        continue;
      }

      // An explicit break: put the break line at the bottom of its page and
      // resume the rest of the block at the top of the next one, then carry on
      // from there.
      if (forcedPending && nextForced) {
        fi += 1;
        const liveForced = nextForced.top + acc;
        const target = Math.max(page, Math.floor(liveForced / pitch)) + 1;
        const height = target * pitch + SLACK - (nextForced.resume + acc);
        if (height > EPS) {
          breaks.push({ pos: nextForced.pos, height, inline: true });
          acc += height;
        }
        page = target;
        segTop = nextForced.resume;
        anchor = nextForced.pos;
        isFirstSegment = false;
        continue;
      }

      // Fill the rest of this page with as many lines as fit, then continue the
      // block on the next page.
      if (unit.splittable) {
        const split = findLineSplit(view, unit, segTop, pageBottom - acc, geom);
        if (split) {
          const height = (page + 1) * pitch + SLACK - (split.top + acc);
          if (height > EPS) {
            breaks.push({ pos: split.pos, height, inline: true });
            acc += height;
          }
          page += 1;
          segTop = split.top;
          anchor = split.pos;
          isFirstSegment = false;
          continue;
        }
      }

      // Not splittable, or nothing fits on this page: move the whole block to
      // the next page top, unless this page is still empty.
      if (liveTop > pageTop + EPS) {
        const height =
          (page + 1) * pitch + SLACK - liveTop - (isFirstSegment ? collapse : 0);
        if (height > EPS) {
          breaks.push({ pos: anchor, height, inline: !isFirstSegment });
          acc += height;
          page += 1;
          isFirstSegment = false;
          continue;
        }
      }

      // A block taller than a full page on an empty page has to spill. Resume
      // the following content at the first page top below the spill.
      page = Math.max(page + 1, Math.ceil(liveBottom / pitch));
      break;
    }

    pages = Math.max(pages, Math.max(1, Math.ceil((unit.bottom + acc) / pitch)));
  }

  return { breaks: dedupe(breaks), pages };
}

function dedupe(breaks: PlannedBreak[]): PlannedBreak[] {
  const out: PlannedBreak[] = [];
  for (const b of breaks) {
    const last = out[out.length - 1];
    if (last && last.pos === b.pos) {
      last.height += b.height;
      continue;
    }
    out.push({ ...b });
  }
  return out;
}

function samePlan(prev: PlannedBreak[], next: PlannedBreak[]): boolean {
  if (prev.length !== next.length) return false;
  return prev.every(
    (b, i) => b.pos === next[i]!.pos && Math.abs(b.height - next[i]!.height) < 1,
  );
}

function buildDecorations(
  view: EditorViewLike,
  breaks: PlannedBreak[],
): DecorationSet {
  const decorations = breaks.map((b, i) => {
    const widget = document.createElement("div");
    widget.className = "oc-page-gap";
    widget.setAttribute("contenteditable", "false");
    widget.style.height = `${b.height}px`;
    widget.dataset.page = String(i + 1);
    if (b.inline) widget.dataset.split = "1";
    return Decoration.widget(b.pos, widget, {
      side: -1,
      key: `pagegap-${i}-${b.pos}`,
      ignoreSelection: true,
    });
  });
  return DecorationSet.create(view.state.doc, decorations);
}

/**
 * The live selection as document positions, or null when there is nothing of
 * the user's to protect. Widget decorations never shift positions, so a range
 * captured before a re-plan is still the right range after it.
 */
function captureSelection(
  view: EditorViewLike,
): { from: number; to: number } | null {
  const sel = window.getSelection();
  if (!sel || sel.isCollapsed || sel.rangeCount === 0) return null;
  const range = sel.getRangeAt(0);
  if (
    !view.dom.contains(range.startContainer) ||
    !view.dom.contains(range.endContainer)
  ) {
    return null;
  }
  try {
    const from = view.posAtDOM(range.startContainer, range.startOffset);
    const to = view.posAtDOM(range.endContainer, range.endOffset);
    return from === to ? null : { from, to };
  } catch {
    return null;
  }
}

/**
 * Swap the decoration set without losing the range the user just dragged.
 *
 * Re-planning rebuilds the blocks around every in-paragraph spacer, which can
 * detach the ends of a DOM selection that ProseMirror has not read yet — a mouse
 * drag lands its range in the document a beat after mouseup, while the re-plan
 * runs on the next frame. Restoring through ProseMirror rather than re-adding a
 * saved Range is what makes this survive: the old nodes may no longer be in the
 * document by then, but their positions still are.
 */
function dispatchPlan(view: EditorViewLike, state: PagedState): void {
  const keep = captureSelection(view);

  const tr = view.state.tr;
  tr.setMeta(KEY, state);
  tr.setMeta("addToHistory", false);
  view.dispatch(tr);

  if (!keep) return;
  const sel = window.getSelection();
  if (sel && !sel.isCollapsed) return; // the range came through on its own

  const size = view.state.doc.content.size;
  const from = Math.min(keep.from, size);
  const to = Math.min(keep.to, size);
  if (from === to) return;
  const restore = view.state.tr;
  try {
    restore.setSelection(TextSelection.create(view.state.doc, from, to));
  } catch {
    return;
  }
  restore.setMeta("addToHistory", false);
  view.dispatch(restore);
}

function runMeasure(view: EditorViewLike, attempt = 0): void {
  const dom = view.dom;
  const geom = readGeom(dom);
  const spacers = Array.from(dom.querySelectorAll<HTMLElement>(".oc-page-gap"));
  for (const el of spacers) el.style.display = "none";

  let plan: Plan | null;
  try {
    plan = planLayout(view, geom);
  } finally {
    for (const el of spacers) el.style.removeProperty("display");
  }

  // Nothing can be read yet. Keep the current plan rather than replacing it
  // with one measured against half a document, and try again shortly.
  if (!plan) {
    if (attempt < 20) setTimeout(() => runMeasure(view, attempt + 1), 250);
    return;
  }

  const prev = KEY.getState(view.state);
  const floatPages = prev?.floatPages ?? 1;
  if (!prev || !samePlan(prev.breaks, plan.breaks)) {
    dispatchPlan(view, {
      deco: buildDecorations(view, plan.breaks),
      pages: plan.pages,
      floatPages,
      announced: prev?.announced ?? 0,
      breaks: plan.breaks,
    });
  }
}
