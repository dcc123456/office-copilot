import { describe, expect, it } from "vitest";
import { argbToSheetColor, sheetColorToArgb } from "./color.js";

describe("ARGB ↔ #rrggbb（实测形态：<color rgb=\"FFFF0000\"/> 与 theme 引用并存）", () => {
  it("不透明色丢掉 alpha，读写互逆", () => {
    expect(argbToSheetColor("FFFF0000")).toEqual({ hex: "#ff0000" });
    expect(sheetColorToArgb({ hex: "#ff0000" })).toBe("FFFF0000");
    expect(sheetColorToArgb({ hex: "#0000ff" })).toBe("FF0000FF");
  });

  it("半透明把 alpha 带出来", () => {
    const color = argbToSheetColor("80FF0000");
    expect(color.hex).toBe("#ff0000");
    expect(color.alpha).toBeCloseTo(0.502, 3);
    expect(sheetColorToArgb(color)).toBe("80FF0000");
  });

  it("读不懂的颜色保留原串而不是编造一个", () => {
    expect(argbToSheetColor("ZZZ").raw).toBe("ZZZ");
    expect(sheetColorToArgb({ hex: "" })).toBeUndefined();
  });
});
