import { describe, expect, it } from "vitest";
import { indentFirstFor, specialFormatOf } from "./ParagraphPopover.js";

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
