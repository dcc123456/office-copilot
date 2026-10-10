import { describe, expect, it } from "vitest";
import { docStyleSheet } from "./docStyleSheet.js";
import type { ResolvedStyle } from "@office-copilot/docx-io";

/**
 * 样式表注入的是用户 .docx 里的内容，逐条盯住"选择器怎么写、读不懂的值怎么办"。
 * 数值本身由 `docx-io` 的 style-format 测试覆盖，这里只验 pt/倍数 → CSS 的落地。
 */
describe("docStyleSheet", () => {
  it("没有解析结果时什么都不注入", () => {
    expect(docStyleSheet([])).toBe("");
  });

  it("文档默认段落样式落在没写 data-style-id 与已清除样式的段落上", () => {
    const css = docStyleSheet([
      {
        styleId: null,
        format: {
          textAlign: "justify",
          fontSize: "14px",
          fontFamily: "宋体",
          fontFamilyWest: "Times New Roman",
        },
      },
    ]);
    expect(css).toContain(
      '.ProseMirror p:not([data-style-id]), .ProseMirror [data-style-id=""]',
    );
    // 西文在前：拉丁字母取 Times New Roman，汉字回退到宋体，正是 ascii/eastAsia 的分工。
    expect(css).toContain("font-family: Times New Roman, 宋体");
    expect(css).toContain("text-align: justify");
  });

  it("具名样式按 pt 与倍数落成 CSS，标题的粗体与固定行距一起带上", () => {
    const css = docStyleSheet([
      {
        styleId: "2",
        format: {
          spacingBefore: 17,
          spacingAfter: 16.5,
          lineHeight: "atLeast:9",
          fontSize: "29.3px",
          bold: true,
        },
      },
    ]);
    expect(css).toContain('.ProseMirror [data-style-id="2"]');
    expect(css).toContain("margin-top: 22.7px");
    expect(css).toContain("margin-bottom: 22px");
    expect(css).toContain("line-height: max(12px, 1.2em)");
    expect(css).toContain("font-weight: bold");
  });

  it("样式明确取消加粗时写 normal，基线粗体的 h1 才压得回去", () => {
    const css = docStyleSheet([{ styleId: "5", format: { bold: false } }]);
    expect(css).toContain("font-weight: normal");
    expect(css).not.toContain("font-style");
  });

  it("下划线与删除线合成一条 text-decoration", () => {
    const css = docStyleSheet([
      { styleId: "7", format: { underline: true, strike: true } },
    ]);
    expect(css).toContain("text-decoration: underline line-through");
  });

  it("引用了本文档没声明的 id 时按默认样式兜底，具名样式不受影响", () => {
    const css = docStyleSheet([
      { styleId: null, format: { fontSize: "14px", textAlign: "justify" } },
      { styleId: "1", format: { fontSize: "14px" } },
    ]);
    // Word 对未定义引用是按默认样式排的；:is() 让这条兜底规则的特异度固定在 0,2,0，
    // 排在具名规则之前，所以声明过的 id 一律由自己的那条接管。
    expect(css).toContain(
      '.ProseMirror [data-style-id]:not(:is([data-style-id="1"]))',
    );
    expect(css.indexOf('[data-style-id="1"] {')).toBeGreaterThan(
      css.indexOf(":not(:is("),
    );
  });

  it("一条具名样式都没有时兜底规则不用写 :not，直接覆盖所有带引用的段落", () => {
    const css = docStyleSheet([
      { styleId: null, format: { fontSize: "14px" } },
      { styleId: `x"][onfocus=alert(1)`, format: { bold: true } },
    ]);
    expect(css).toContain(".ProseMirror [data-style-id] {");
    expect(css).not.toContain(":not(:is(");
    expect(css).not.toContain("onfocus");
  });

  it("styleId 不合字符集时整条跳过，不让它进选择器", () => {
    const hostile = { styleId: `x"][onfocus=alert(1)`, format: { bold: true } };
    const css = docStyleSheet([
      hostile,
      { styleId: "3", format: { fontSize: "21.3px" } },
    ]);
    expect(css).not.toContain("onfocus");
    expect(css).toContain('.ProseMirror [data-style-id="3"]');
  });

  it("读不懂的字体名与颜色按没写过处理，其余字段照常生效", () => {
    const css = docStyleSheet([
      {
        styleId: "4",
        format: {
          fontFamily: '宋体"},*{color:red',
          fontSize: "18pt",
          color: "rgb(192, 0, 0)",
          letterSpacing: "-0.35px",
        },
      },
    ]);
    expect(css).not.toContain("color:red");
    expect(css).not.toContain("18pt");
    expect(css).not.toContain("rgb(");
    expect(css).toContain("letter-spacing: -0.35px");
  });
});
