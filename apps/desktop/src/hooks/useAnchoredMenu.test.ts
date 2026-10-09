import { describe, expect, it } from "vitest";
import { anchoredRect } from "./useAnchoredMenu.js";

describe("anchoredRect", () => {
  it("hangs the menu just below the trigger", () => {
    expect(anchoredRect({ left: 120, bottom: 64 }, 150, 1200)).toEqual({ left: 120, top: 70 });
  });

  it("keeps a menu near the right edge fully on screen", () => {
    expect(anchoredRect({ left: 1160, bottom: 64 }, 150, 1200).left).toBe(1042);
  });

  it("never pushes the menu off the left edge", () => {
    expect(anchoredRect({ left: -20, bottom: 30 }, 280, 1200).left).toBe(8);
  });
});
