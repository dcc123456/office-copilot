import { useCallback } from "react";
import { startPointerDrag } from "./usePointerDrag.js";

const clamp = (v: number, min: number, max: number) =>
  Math.min(max, Math.max(min, v));

/**
 * Width after dragging a panel's resizer by `dx` CSS px. A "left" panel (the
 * outline) widens as the pointer moves right; a "right" panel (the AI assistant)
 * widens as it moves left.
 */
export function nextWidth(
  side: "left" | "right",
  startWidth: number,
  dx: number,
  min: number,
  max: number,
): number {
  return clamp(side === "left" ? startWidth + dx : startWidth - dx, min, max);
}

/**
 * Pointer-driven width drag for a side panel. `side` says which edge the
 * panel sits on: a "left" panel widens as the pointer moves right, a "right"
 * panel widens as it moves left.
 */
export function usePanelResize({
  side,
  width,
  setWidth,
  min,
  max,
  resetTo,
}: {
  side: "left" | "right";
  width: number;
  setWidth: (w: number) => void;
  min: number;
  max: number;
  resetTo: number;
}) {
  const onPointerDown = useCallback(
    (e: React.PointerEvent<HTMLDivElement>) => {
      if (e.button !== 0) return;
      e.preventDefault();
      const startX = e.clientX;
      const startWidth = width;
      document.body.classList.add("is-resizing");

      startPointerDrag(e.currentTarget, e.pointerId, {
        move: (ev) =>
          setWidth(nextWidth(side, startWidth, ev.clientX - startX, min, max)),
        end: () => document.body.classList.remove("is-resizing"),
      });
    },
    [width, setWidth, side, min, max],
  );

  const onDoubleClick = useCallback(() => setWidth(resetTo), [setWidth, resetTo]);

  return { onPointerDown, onDoubleClick };
}
