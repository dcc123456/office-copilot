import { describe, expect, it } from "vitest";
import { buildPrintHtml } from "./printHtml.js";

const base = {
  bodyHtml: "<p>正文</p>",
  pageWidthPx: 794,
  pageHeightPx: 1123,
  marginTopPx: 96,
  marginRightPx: 96,
  marginBottomPx: 96,
  marginLeftPx: 96,
  headerPx: 48,
  footerPx: 48,
};

describe("打印 / PDF 导出的分栏", () => {
  it("两栏时正文区用 CSS columns 按 Word 的顺序填满第一栏", () => {
    const html = buildPrintHtml({ ...base, columns: { count: 2, gapPx: 30 } });
    expect(html).toContain("column-count: 2");
    expect(html).toContain("column-gap: 30px");
    expect(html).toContain("column-fill: auto");
    // 正文区高度 = 页高 - 上下页边距，`column-fill: auto` 要有实高才肯填满一栏再换栏。
    expect(html).toContain("height: 931px");
  });

  it("单栏（或没设过分栏）时不写任何列样式", () => {
    expect(buildPrintHtml({ ...base, columns: { count: 1, gapPx: 30 } })).not.toContain(
      "column-count",
    );
    expect(buildPrintHtml(base)).not.toContain("column-count");
  });
});
