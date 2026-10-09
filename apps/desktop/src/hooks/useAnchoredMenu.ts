import { useCallback, useEffect, useRef, useState } from "react";

export interface MenuPosition {
  left: number;
  top: number;
}

/**
 * Where a menu of `width` belongs under a trigger, in viewport coordinates.
 * Clamped so a menu near the right edge stays entirely on screen.
 */
export function anchoredRect(
  trigger: { left: number; bottom: number },
  width: number,
  viewportWidth: number,
  gap = 6,
): MenuPosition {
  return {
    left: Math.max(8, Math.min(trigger.left, viewportWidth - width - 8)),
    top: trigger.bottom + gap,
  };
}

/**
 * Anchor a dropdown to its trigger with `position: fixed`.
 *
 * The toolbar bands scroll horizontally, which turns them into clipping boxes:
 * a menu rendered `absolute` inside a band opened invisibly, hanging below the
 * band's own box. Viewport coordinates escape the clip. `width` is the menu's
 * own width, used to keep it inside the window.
 */
export function useAnchoredMenu(width: number) {
  const anchorRef = useRef<HTMLDivElement>(null);
  const [rect, setRect] = useState<MenuPosition | null>(null);

  const show = useCallback(() => {
    const el = anchorRef.current;
    if (!el) return;
    const r = el.getBoundingClientRect();
    setRect(anchoredRect({ left: r.left, bottom: r.bottom }, width, window.innerWidth));
  }, [width]);

  const hide = useCallback(() => setRect(null), []);
  const toggle = useCallback(() => (rect ? hide() : show()), [rect, hide, show]);

  useEffect(() => {
    if (!rect) return;
    // The anchor moves when the band scrolls or the window resizes, so the
    // cached rect goes stale. Editor scrolls must not close the menu — the
    // popover's own commands focus the editor.
    const band = anchorRef.current?.closest(".tb-band");
    band?.addEventListener("scroll", hide);
    window.addEventListener("resize", hide);
    return () => {
      band?.removeEventListener("scroll", hide);
      window.removeEventListener("resize", hide);
    };
  }, [rect, hide]);

  return { anchorRef, rect, open: rect !== null, show, hide, toggle };
}
