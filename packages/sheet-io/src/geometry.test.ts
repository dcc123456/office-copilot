import { describe, expect, it } from "vitest";
import {
  DEFAULT_MAX_DIGIT_WIDTH_PX,
  columnWidthCharsToPx,
  columnWidthPxToChars,
  rowHeightPtToPx,
} from "./geometry.js";

describe("列宽与行高的换算（spec 决策 D1：换算单点）", () => {
  it("Excel 默认列宽 8.43 字符在 96DPI 下就是 64 px", () => {
    // 8.43 是 Excel 空白工作簿写在 @width 上的值，64 px 是它的默认列宽，两个数必须对上
    expect(columnWidthCharsToPx(2158 / 256)).toBe(64);
  });

  it("整数字符宽度按 px = chars × MDW + 5 展开", () => {
    expect(columnWidthCharsToPx(12)).toBe(89);
    expect(columnWidthCharsToPx(14)).toBe(103);
    expect(columnWidthCharsToPx(0)).toBe(5);
  });

  it("拖动出来的像素换成字符再换回像素，不许逐次漂移", () => {
    for (let px = 20; px <= 400; px += 7) {
      expect(columnWidthCharsToPx(columnWidthPxToChars(px))).toBe(px);
    }
  });

  it("默认字体宽度是 Calibri 11 的 7 px，行高 15 磅是 20 px", () => {
    expect(DEFAULT_MAX_DIGIT_WIDTH_PX).toBe(7);
    expect(rowHeightPtToPx(15)).toBe(20);
    expect(rowHeightPtToPx(22.5)).toBe(30);
  });
});
