import { Extension } from "@tiptap/core";
import type { Node as PMNode } from "@tiptap/pm/model";
import { Plugin, PluginKey } from "@tiptap/pm/state";
import type { EditorView, NodeView } from "@tiptap/pm/view";

/**
 * Allow dragging floating (layout="anchor") images to reposition them.
 * The Drawing node renders a zero-width <span data-drawing="anchor">
 * containing an absolutely-positioned <img>; we attach native mouse
 * listeners to those <img> elements and update the node's posX/posY
 * attrs via a ProseMirror transaction. The in-flow text anchor stays
 * where it was (so the paragraph the picture is logically anchored to
 * doesn't change), but the picture itself moves.
 *
 * On save, when ooxml is absent, ooxml-pm synthesizes a fresh
 * <wp:anchor> with the new coordinates.
 */

const key = new PluginKey("floatingImageDrag");

interface DragState {
  pos: number;
  startX: number;
  startY: number;
  origX: number;
  origY: number;
  img: HTMLImageElement;
}

function eventPosInDoc(view: EditorView, event: MouseEvent): number {
  const coords = { left: event.clientX, top: event.clientY };
  return view.posAtCoords(coords)?.pos ?? -1;
}

function findImageAtPos(doc: PMNode, pos: number): PMNode | null {
  if (pos < 0 || pos > doc.content.size) return null;
  const $pos = doc.resolve(pos);
  // The drawn image lives at pos-1 (it's an atom inline node).
  for (let depth = $pos.depth; depth >= 0; depth--) {
    const node = depth === $pos.depth ? doc.nodeAt(pos) : $pos.node(depth);
    if (node?.type.name === "image") return node;
    const before = doc.nodeAt(pos - 1);
    if (before?.type.name === "image") return before;
  }
  return doc.nodeAt(pos - 1);
}

export const FloatingImageDrag = Extension.create({
  name: "floatingImageDrag",

  addProseMirrorPlugins() {
    let drag: DragState | null = null;

    const onMouseDown = (view: EditorView, e: MouseEvent) => {
      const target = e.target as HTMLElement;
      const img = target.closest(
        'span[data-drawing="anchor"] img[data-image="anchor"]',
      ) as HTMLImageElement | null;
      if (!img) return;
      // Don't hijack simple clicks (selection) — only start dragging on
      // an explicit drag threshold.
      const posInfo = view.posAtDOM(img.parentElement!, 0);
      const node = view.state.doc.nodeAt(posInfo);
      if (!node || node.type.name !== "image") return;
      if (node.attrs.layout !== "anchor") return;

      e.preventDefault();
      view.focus();
      drag = {
        pos: posInfo,
        startX: e.clientX,
        startY: e.clientY,
        origX: node.attrs.posX ?? 0,
        origY: node.attrs.posY ?? 0,
        img,
      };
      img.style.cursor = "grabbing";
      img.style.opacity = "0.8";

      const onMove = (ev: MouseEvent) => {
        if (!drag) return;
        const dx = ev.clientX - drag.startX;
        const dy = ev.clientY - drag.startY;
        const nextX = Math.round(drag.origX + dx);
        const nextY = Math.round(drag.origY + dy);
        // Visual live update without a transaction for 60fps feel.
        drag.img.style.left = `${nextX}px`;
        drag.img.style.top = `${nextY}px`;
      };
      const onUp = (ev: MouseEvent) => {
        if (!drag) return;
        const dx = ev.clientX - drag.startX;
        const dy = ev.clientY - drag.startY;
        const moved = Math.abs(dx) + Math.abs(dy) > 2;
        if (moved) {
          const nextX = Math.round(drag.origX + dx);
          const nextY = Math.round(drag.origY + dy);
          const tr = view.state.tr.setNodeMarkup(drag.pos, undefined, {
            ...view.state.doc.nodeAt(drag.pos)!.attrs,
            posX: nextX,
            posY: nextY,
            // Clear source OOXML so the converter regenerates anchor XML
            // with the new position on save.
            ooxml: null,
          });
          view.dispatch(tr);
        }
        drag.img.style.cursor = "";
        drag.img.style.opacity = "";
        drag = null;
        window.removeEventListener("mousemove", onMove);
        window.removeEventListener("mouseup", onUp);
      };
      window.addEventListener("mousemove", onMove);
      window.addEventListener("mouseup", onUp);
    };

    return [
      new Plugin({
        key,
        props: {
          handleDOMEvents: {
            mousedown: onMouseDown,
          },
          decorations(state) {
            // Mark floating images with a grab cursor via a DOM
            // attribute handled in CSS.
            void state;
            return null;
          },
        },
      }),
    ];
  },
});
