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

describe("打印 / PDF 导出的页面边框", () => {
  const borders = {
    offsetFrom: "page",
    top: { style: "single", sizePt: 1, spacePt: 2, color: "#FF0000" },
    left: null,
    bottom: null,
    right: { style: "dashed", sizePt: 3, spacePt: 1, color: null },
  };

  it("每条开着的边画一根绝对定位的线，位置与颜色和画布预览一致", () => {
    const html = buildPrintHtml({ ...base, pageBorders: borders });
    // 上边：间距 2pt ≈ 2.67px，1pt 单线 ≈ 1.3px；横线从左边（此处没开，x=0）量到右边。
    expect(html).toContain(
      '<div class="pgb" style="left: 0px; top: 2.67px; width: 792.67px; border-top: 1.3px solid #FF0000;"></div>',
    );
    // 右边：3pt = 4px 虚线，颜色为 Word 的"自动"时按黑色画。
    expect(html).toContain(
      '<div class="pgb" style="left: 792.67px; top: 2.67px; height: 1120.33px; border-left: 4px dashed #000000;"></div>',
    );
    // 关掉的左/下边不产生任何盒子。
    expect(html.match(/class="pgb"/g)).toHaveLength(2);
  });

  it("没设边框或四条边都关掉时纸上什么都不画", () => {
    expect(buildPrintHtml(base)).not.toContain('class="pgb"');
    expect(
      buildPrintHtml({
        ...base,
        pageBorders: { offsetFrom: "page", top: null, left: null, bottom: null, right: null },
      }),
    ).not.toContain('class="pgb"');
  });
});

describe("打印 / PDF 导出的首字下沉", () => {
  it("纸上是画布那一条 ::first-letter 规则，不是第二份实现", () => {
    const html = buildPrintHtml(base);
    expect(html).toContain(".oc-drop-cap::first-letter");
    expect(html).toContain("initial-letter: var(--oc-drop-cap");
  });
});
