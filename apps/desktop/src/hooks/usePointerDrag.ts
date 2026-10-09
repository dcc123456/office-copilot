/**
 * A drag that is guaranteed to end.
 *
 * `mousemove` + `mouseup` on `window` leaks: grab a window edge and the pointer
 * leaves the web contents immediately, so the browser never delivers `mouseup`,
 * the listeners stay bound, and the *next* hover resizes the panel or window with
 * no button held. Pointer capture keeps routing events to the element until the
 * real release, wherever it happens.
 */
export function startPointerDrag(
  el: HTMLElement,
  pointerId: number,
  handlers: { move: (ev: PointerEvent) => void; end?: () => void },
): void {
  let active = true;

  const move = (ev: PointerEvent) => {
    if (active && ev.pointerId === pointerId) handlers.move(ev);
  };
  const stop = (ev: PointerEvent) => {
    if (!active || ev.pointerId !== pointerId) return;
    active = false;
    if (el.hasPointerCapture?.(pointerId)) el.releasePointerCapture(pointerId);
    el.removeEventListener("pointermove", move);
    el.removeEventListener("pointerup", stop);
    el.removeEventListener("pointercancel", stop);
    handlers.end?.();
  };

  try {
    el.setPointerCapture(pointerId);
  } catch {
    // Pointer already released (synthetic events in tests); the listeners below
    // still end the drag.
  }
  el.addEventListener("pointermove", move);
  el.addEventListener("pointerup", stop);
  el.addEventListener("pointercancel", stop);
}
