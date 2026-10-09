import { useCallback } from "react";
import { startPointerDrag } from "../hooks/usePointerDrag.js";

type Edge = "n" | "s" | "e" | "w" | "ne" | "nw" | "se" | "sw";

const EDGES: { edge: Edge; cls: string; title: string }[] = [
  { edge: "e", cls: "edge-e", title: "拖动调整窗口宽度" },
  { edge: "w", cls: "edge-w", title: "拖动调整窗口宽度" },
  { edge: "s", cls: "edge-s", title: "拖动调整窗口高度" },
  { edge: "n", cls: "edge-n", title: "拖动调整窗口高度" },
  { edge: "se", cls: "edge-se", title: "拖动调整窗口大小" },
  { edge: "sw", cls: "edge-sw", title: "拖动调整窗口大小" },
  { edge: "ne", cls: "edge-ne", title: "拖动调整窗口大小" },
  { edge: "nw", cls: "edge-nw", title: "拖动调整窗口大小" },
];

/**
 * DOM hot-zones on the window edges that resize the Electron window:
 * press to start, move to resize, release to stop. Edge semantics + min-size
 * clamping live in the main process ("window:drag-resize").
 *
 * The pointer is captured because grabbing an edge moves it straight out of the
 * web contents: with plain mouse events `mouseup` never arrived, the drag stayed
 * bound, and the window kept resizing on the next hover.
 */
export function EdgeResizer() {
  const startResize = useCallback(
    (edge: Edge) => (e: React.PointerEvent<HTMLDivElement>) => {
      if (e.button !== 0 || !window.officeCopilot?.dragResizeWindow) return;
      e.preventDefault();
      // Screen coordinates, not clientX/clientY: dragging the left/top edge moves
      // the window under the pointer, which would rewrite clientX and feed the
      // displacement back into itself.
      const startX = e.screenX;
      const startY = e.screenY;
      document.body.classList.add("is-window-resizing");
      document.body.dataset.windowResize = edge;

      let base: { x: number; y: number; width: number; height: number } | null =
        null;
      void window.officeCopilot.getWindowBounds().then((b) => (base = b));

      startPointerDrag(e.currentTarget, e.pointerId, {
        move: (ev) => {
          if (!base) return;
          window.officeCopilot.dragResizeWindow({
            base,
            dx: ev.screenX - startX,
            dy: ev.screenY - startY,
            edge,
          });
        },
        end: () => {
          document.body.classList.remove("is-window-resizing");
          delete document.body.dataset.windowResize;
        },
      });
    },
    [],
  );

  if (!window.officeCopilot?.dragResizeWindow) return null;

  return (
    <>
      {EDGES.map((d) => (
        <div
          key={d.edge}
          className={`edge-resizer ${d.cls}`}
          title={d.title}
          onPointerDown={startResize(d.edge)}
        />
      ))}
    </>
  );
}
