import { describe, expect, it } from "vitest";
import { splitCuts } from "./FloatingTextBox.js";

describe("splitCuts", () => {
  it("leaves a box that fits its page alone", () => {
    expect(splitCuts([18, 36, 54, 72], 100, 900)).toEqual([]);
    expect(splitCuts([], 100, 900)).toEqual([]);
  });

  it("cuts at the last line that still fits, and continues a full band after that", () => {
    // 55px left on the box's own page, then a whole band on each next page.
    expect(splitCuts([20, 40, 60, 80, 100], 55, 50)).toEqual([40, 80]);
  });

  it("gives up when the next line does not fit the room left", () => {
    // Not even one line fits on this page, so there is no boundary to cut at.
    expect(splitCuts([20, 40, 60], 10, 500)).toEqual([]);
  });

  it("spends as many pages as the text needs", () => {
    const lines = Array.from({ length: 50 }, (_, i) => (i + 1) * 20);
    const cuts = splitCuts(lines, 100, 100);
    expect(cuts).toEqual([100, 200, 300, 400, 500, 600, 700, 800, 900]);
  });

  it("stops at maxParts and leaves the tail on the last page it reached", () => {
    const lines = Array.from({ length: 50 }, (_, i) => (i + 1) * 20);
    expect(splitCuts(lines, 100, 100, 3)).toHaveLength(3);
  });
});
