import { describe, expect, it } from "vitest";
import { indentFirstFor, specialFormatOf, tabStopPositionPt } from "./ParagraphPopover.js";

describe("段落特殊格式与 indentFirst 的换算", () => {
  it("把属性值读回 Word 的三个取值", () => {
    expect(specialFormatOf(undefined)).toBe("none");
    expect(specialFormatOf(0)).toBe("none");
    expect(specialFormatOf(21)).toBe("firstLine");
    expect(specialFormatOf(-21)).toBe("hanging");
  });

  it("导入的悬挂缩进不会被面板改成正数", () => {
    // 旧实现只有一个"首行缩进"数字框且 min=0，读进来 -21 一改就变 21。
    expect(indentFirstFor("hanging", 21)).toBe(-21);
    expect(specialFormatOf(indentFirstFor("hanging", 21))).toBe("hanging");
  });

  it("切换特殊格式时给一个空磅值补 2 字符的初值", () => {
    expect(indentFirstFor("firstLine", undefined)).toBe(21);
    expect(indentFirstFor("hanging", 0)).toBe(-21);
    expect(indentFirstFor("none", 21)).toBeUndefined();
  });
});

describe("制表位位置的厘米录入", () => {
  it("按 Word 的口径显示厘米，存进段落属性的仍是磅", () => {
    expect(tabStopPositionPt("2")).toBe(56.7);
    expect(tabStopPositionPt(" 14.65 ")).toBe(415.3);
    expect(tabStopPositionPt("0")).toBe(0);
  });

  it("空、负数、非数字都不算一个位置——界面据此禁用「设置」", () => {
    for (const raw of ["", "   ", "-1", "abc", "1.2.3", "Infinity"]) {
      expect(tabStopPositionPt(raw)).toBeNull();
    }
  });
});
