import { describe, expect, it } from "vitest";
import { textEmCount, watermarkLayerFor } from "./pageWatermark.js";
import type { SectionWatermark } from "@office-copilot/docx-io";

/** A4 竖排的整数版几何，CSS px。 */
const PAGE = {
  pageWidth: 800,
  pageHeight: 1000,
  marginTop: 100,
  marginRight: 120,
  marginBottom: 100,
  marginLeft: 120,
};

/** 一条水印设置，只覆盖测试关心的字段。 */
function wm(patch: Partial<SectionWatermark> = {}): SectionWatermark {
  return {
    text: "草稿",
    font: "宋体",
    sizePt: null,
    color: "#c0c0c0",
    semiTransparent: true,
    layout: "diagonal",
    ...patch,
  };
}

describe("水印形状换算成画布与纸上的呈现", () => {
  it("盒子取 Word 的 468pt × 117pt，并居中在页边距之内", () => {
    const layer = watermarkLayerFor(wm(), PAGE);
    // pt → px 只在这一处发生（AGENTS §2.1）：468pt = 624px，117pt = 156px。
    expect(layer?.widthPx).toBe(624);
    expect(layer?.heightPx).toBe(156);
    expect(layer?.leftPx).toBe(88); // 120 + (560 - 624) / 2
    expect(layer?.topPx).toBe(422); // 100 + (800 - 156) / 2
  });

  it("字号留空（Word 的自动）时按框宽缩到放得下，且不超过框高", () => {
    expect(watermarkLayerFor(wm({ text: "草稿" }), PAGE)?.fontSizePx).toBe(156);
    // 8 个全角字要 8 em，624px 宽只放得下 78px 字号。
    expect(
      watermarkLayerFor(wm({ text: "内部资料严禁外传" }), PAGE)?.fontSizePx,
    ).toBe(78);
  });

  it("写死字号时按 pt 换算，不再自动缩放", () => {
    expect(watermarkLayerFor(wm({ sizePt: 48 }), PAGE)?.fontSizePx).toBe(64);
  });

  it("对角版式是 -45 度，横排是 0 度", () => {
    expect(watermarkLayerFor(wm(), PAGE)?.rotateDeg).toBe(-45);
    expect(
      watermarkLayerFor(wm({ layout: "horizontal" }), PAGE)?.rotateDeg,
    ).toBe(0);
  });

  it("半透明是 0.5，不勾就是实心", () => {
    expect(watermarkLayerFor(wm(), PAGE)?.opacity).toBe(0.5);
    expect(
      watermarkLayerFor(wm({ semiTransparent: false }), PAGE)?.opacity,
    ).toBe(1);
  });

  it("文字为空或全是空格时不画", () => {
    expect(watermarkLayerFor(wm({ text: "" }), PAGE)).toBeNull();
    expect(watermarkLayerFor(wm({ text: "   " }), PAGE)).toBeNull();
  });

  it("字体名里的注入字符整串退回 serif，正常中文名照用", () => {
    expect(watermarkLayerFor(wm({ font: '宋体" onclick="x' }), PAGE)?.fontFamily).toBe(
      '"serif"',
    );
    expect(watermarkLayerFor(wm({ font: "  " }), PAGE)?.fontFamily).toBe('"serif"');
    expect(watermarkLayerFor(wm({ font: "Microsoft YaHei" }), PAGE)?.fontFamily).toBe(
      '"Microsoft YaHei"',
    );
  });

  it("读不懂的颜色按黑色画，不把非法值写进 CSS（AGENTS §2.2）", () => {
    expect(watermarkLayerFor(wm({ color: "red" }), PAGE)?.color).toBe("#000000");
    expect(watermarkLayerFor(wm({ color: "#FF0000" }), PAGE)?.color).toBe("#FF0000");
  });

  it("全角算一个 em，拉丁与数字算 0.55 个", () => {
    expect(textEmCount("草稿")).toBe(2);
    expect(textEmCount("DRAFT")).toBeCloseTo(2.75);
  });
});
