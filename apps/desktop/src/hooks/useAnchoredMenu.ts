import { useCallback, useEffect, useRef, useState } from "react";

export interface MenuPosition {
  left: number;
  top: number;
}

/**
 * Where a menu of `width` belongs under a trigger, in viewport coordinates.
 * Clamped so a menu near the right edge stays entirely on screen.
 *
 * @param trigger 触发元素的视口矩形（只用到 left/right/bottom，单位 px）。
 * @param width 浮层自己的宽度，px。
 * @param viewportWidth 视口宽度，px。
 * @param gap 浮层与触发元素之间的间距，px。
 * @param anchor 水平对齐的那条边："left" 贴触发元素左边（默认），"right" 贴右边——
 *   触发元素在窄容器（如右侧面板）里时，贴左会把浮层甩出容器外。
 */
export function anchoredRect(
  trigger: { left: number; right: number; bottom: number },
  width: number,
  viewportWidth: number,
  gap = 6,
  anchor: "left" | "right" = "left",
): MenuPosition {
  const wanted = anchor === "right" ? trigger.right - width : trigger.left;
  return {
    left: Math.max(8, Math.min(wanted, viewportWidth - width - 8)),
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
export function useAnchoredMenu(width: number, anchor: "left" | "right" = "left") {
  const anchorRef = useRef<HTMLDivElement>(null);
  const [rect, setRect] = useState<MenuPosition | null>(null);

  const show = useCallback(() => {
    const el = anchorRef.current;
    if (!el) return;
    const r = el.getBoundingClientRect();
    setRect(
      anchoredRect({ left: r.left, right: r.right, bottom: r.bottom }, width, window.innerWidth, 6, anchor),
    );
  }, [width, anchor]);

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
