import { describe, expect, it } from "vitest";
import { anchoredRect } from "./useAnchoredMenu.js";

describe("anchoredRect", () => {
  it("hangs the menu just below the trigger", () => {
    expect(anchoredRect({ left: 120, right: 144, bottom: 64 }, 150, 1200)).toEqual({
      left: 120,
      top: 70,
    });
  });

  it("keeps a menu near the right edge fully on screen", () => {
    expect(anchoredRect({ left: 1160, right: 1184, bottom: 64 }, 150, 1200).left).toBe(1042);
  });

  it("never pushes the menu off the left edge", () => {
    expect(anchoredRect({ left: -20, right: 260, bottom: 30 }, 280, 1200).left).toBe(8);
  });

  it("anchors to the trigger's right edge for a narrow right-hand panel", () => {
    // 触发元素贴在 340px 面板里、面板右边在 1200：菜单右边缘应与触发右边缘对齐。
    expect(anchoredRect({ left: 900, right: 1188, bottom: 40 }, 250, 1200, 6, "right")).toEqual({
      left: 938,
      top: 46,
    });
  });

  it("slides a right-anchored menu back on screen when the trigger is at the left", () => {
    expect(anchoredRect({ left: 20, right: 60, bottom: 40 }, 250, 1200, 6, "right").left).toBe(8);
  });
});
