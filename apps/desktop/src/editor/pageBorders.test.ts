import { describe, expect, it } from "vitest";
import type { SectionPageBorders } from "@office-copilot/docx-io";
import {
  borderLinesFor,
  borderSideToCss,
  emptyPageBorders,
  hasAnyBorderSide,
} from "./pageBorders.js";

const side = (style: string, sizePt = 1, spacePt = 1, color: string | null = null) => ({
  style,
  sizePt,
  spacePt,
  color,
});

/** 四条边都相同的边框，方便只改一个字段。 */
function allSides(patch: Partial<SectionPageBorders> = {}): SectionPageBorders {
  return {
    offsetFrom: "page",
    top: side("single"),
    left: side("single"),
    bottom: side("single"),
    right: side("single"),
    ...patch,
  };
}

const PAGE = {
  pageWidth: 794,
  pageHeight: 1123,
  marginTop: 96,
  marginRight: 96,
  marginBottom: 96,
  marginLeft: 96,
};

describe("页面边框线型的 CSS 换算", () => {
  it("八分之一 pt 的线宽换成 px，自动色就是黑", () => {
    // sz=24 是 3pt，@96DPI = 4px；这是渲染层的换算，模型里始终保持 pt。
    expect(borderSideToCss(side("single", 3, 1, null))).toEqual({
      widthPx: 4,
      style: "solid",
      color: "#000000",
    });
    expect(borderSideToCss(side("single", 0.5, 1, "#FF0000")).color).toBe("#FF0000");
  });

  it("线型按最接近的观感降级，双线至少 3px 才画得出两根", () => {
    expect(borderSideToCss(side("dotted", 1)).style).toBe("dotted");
    expect(borderSideToCss(side("dashed", 1)).style).toBe("dashed");
    expect(borderSideToCss(side("dotDotDash", 1)).style).toBe("dotted");
    expect(borderSideToCss(side("thick", 2.25)).style).toBe("solid");
    expect(borderSideToCss(side("triple", 1)).style).toBe("double");
    expect(borderSideToCss(side("wave", 1)).style).toBe("solid");
    expect(borderSideToCss(side("double", 0.5)).widthPx).toBe(3);
  });
});

describe("页面边框的位置", () => {
  it("距页面时从页边往里量，四条线在四个角相接", () => {
    const lines = borderLinesFor(allSides({ top: side("single", 1, 2) }), PAGE);
    const byKey = Object.fromEntries(lines.map((l) => [l.key, l]));
    // 1pt ≈ 1.333px、2pt ≈ 2.667px：上边按自己的间距量，其余三条用缺省的 1pt。
    expect(byKey.top!.horizontal).toBe(true);
    expect(byKey.top!.y).toBeCloseTo(2.667, 2);
    expect(byKey.top!.x).toBeCloseTo(1.333, 2);
    expect(byKey.top!.length).toBeCloseTo(794 - 1.333 * 2, 2);
    expect(byKey.left!.horizontal).toBe(false);
    expect(byKey.left!.x).toBeCloseTo(1.333, 2);
    expect(byKey.bottom!.y).toBeCloseTo(1123 - 1.333, 2);
    expect(byKey.right!.x).toBeCloseTo(794 - 1.333, 2);
    // 每条线把自己那条边的设置带出去，调用方不必再判断"这条边存在吗"。
    expect(byKey.top!.side.spacePt).toBe(2);
  });

  it("距正文时边框落在正文区与页边之间", () => {
    const lines = borderLinesFor(
      allSides({ offsetFrom: "text", top: side("single", 1, 4), bottom: side("single", 1, 4) }),
      PAGE,
    );
    const byKey = Object.fromEntries(lines.map((l) => [l.key, l]));
    // 4pt ≈ 5.333px：上线在正文上缘之上，下线在正文下缘之下。
    expect(byKey.top!.y).toBeCloseTo(96 - 5.333, 2);
    expect(byKey.bottom!.y).toBeCloseTo(1123 - 96 + 5.333, 2);
  });

  it("关掉的边不产生线，全关时数组为空", () => {
    expect(borderLinesFor(emptyPageBorders(), PAGE)).toEqual([]);
    const onlyTop = borderLinesFor(
      { ...emptyPageBorders(), top: side("single") },
      PAGE,
    );
    expect(onlyTop.map((l) => l.key)).toEqual(["top"]);
  });

  it("hasAnyBorderSide 只在至少一条边开着时为真", () => {
    expect(hasAnyBorderSide(null)).toBe(false);
    expect(hasAnyBorderSide(emptyPageBorders())).toBe(false);
    expect(hasAnyBorderSide({ ...emptyPageBorders(), left: side("dashed") })).toBe(true);
  });
});
