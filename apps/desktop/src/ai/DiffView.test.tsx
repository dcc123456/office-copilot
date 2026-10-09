import { describe, expect, it } from "vitest";
import { renderToStaticMarkup } from "react-dom/server";
import { DiffView, InlineDiff } from "./DiffView.js";
import type { Block } from "./docEdit.js";

const para = (i: number, text: string, style: Block["style"] = {}): Block => ({
  index: i,
  type: "paragraph",
  text,
  style,
});

describe("DiffView", () => {
  const before = [para(0, "一"), para(1, "二"), para(2, "三"), para(3, "四"), para(4, "五"), para(5, "六"), para(6, "第七段 原文"), para(7, "八")];
  const after = [para(0, "一"), para(1, "二"), para(2, "三"), para(3, "四", { bold: true }), para(4, "五"), para(5, "六"), para(6, "第七段 改文"), para(7, "八")];
  const html = renderToStaticMarkup(<DiffView before={before} after={after} />);

  it("shows text changes as numbered -/+ lines", () => {
    expect(html).toContain('<div class="diff-line del"><span class="diff-ln">7</span><span class="diff-ln"></span><span class="diff-sign">-');
    expect(html).toContain('<div class="diff-line add"><span class="diff-ln"></span><span class="diff-ln">7</span><span class="diff-sign">+');
    expect(html).toContain("<span>第七段 </span><del>原</del><span>文</span>");
    expect(html).toContain("<span>第七段 </span><ins>改</ins><span>文</span>");
  });

  it("collapses the unchanged stretch around them", () => {
    expect(html).toContain("未改动 4 行");
    expect(html.match(/diff-line ctx/g)).toHaveLength(3);
  });

  it("lists a style-only change on its own lines", () => {
    expect(html).toContain("段落格式变化");
    expect(html).toContain("样式：默认样式");
    expect(html).toContain("样式：加粗");
  });

  it("says so when nothing is visible", () => {
    expect(renderToStaticMarkup(<DiffView before={before} after={before} />)).toContain("没有可见变化");
  });

  it("renders an inline word diff for rewrites", () => {
    const h = renderToStaticMarkup(<InlineDiff before="我们讨论一下" after="我们讨论下" />);
    expect(h).toContain("<del>一</del>");
    expect(h).not.toContain("diff-ln");
  });

  it("names tables and images instead of printing nothing", () => {
    const b = [para(0, "前言")];
    const a = [
      para(0, "前言"),
      { index: 1, type: "table" as const, text: "条款 | 责任方", rows: [["条款", "责任方"]], style: {} },
    ];
    expect(renderToStaticMarkup(<DiffView before={b} after={a} />)).toContain("[表格] 条款 | 责任方");
  });
});
