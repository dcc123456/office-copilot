import { describe, expect, it } from "vitest";
import { cellDisplayText, formatNumberForDisplay, numberFormatFromCode, parseUserCellText } from "./cell-text.js";
import type { SheetCell } from "./model.js";

/** 造一个测试用的格；只填被测函数会读的字段。 */
function cell(parts: Partial<SheetCell>): SheetCell {
  return { address: "A1", colIndex: 0, rowIndex: 0, kind: "number", value: null, ...parts };
}

describe("格式码准入（spec 决策 D10 第 3 条：认不出的一律拒）", () => {
  it("内建码按 ECMA-376 的 id 落，日期那条带 isDateLike，大小写与首尾空格不参与判定", () => {
    expect(numberFormatFromCode("General")).toEqual({ numFmtId: 0, isDateLike: false });
    expect(numberFormatFromCode("0")).toEqual({ numFmtId: 1, isDateLike: false });
    expect(numberFormatFromCode("0.00")).toEqual({ numFmtId: 2, isDateLike: false });
    expect(numberFormatFromCode("#,##0")).toEqual({ numFmtId: 3, isDateLike: false });
    expect(numberFormatFromCode(" 0% ")).toEqual({ numFmtId: 9, isDateLike: false });
    expect(numberFormatFromCode("0.00%")).toEqual({ numFmtId: 10, isDateLike: false });
    expect(numberFormatFromCode(" M/D/YYYY ")).toEqual({ numFmtId: 14, isDateLike: true });
    expect(numberFormatFromCode("@")).toEqual({ numFmtId: 49, isDateLike: false });
  });

  it("内建表盖不住的小数位走自定义码，与键入时产的是同一套形状", () => {
    expect(numberFormatFromCode("0.0")).toEqual({ numFmtId: 0, formatCode: "0.0", isDateLike: false });
    expect(numberFormatFromCode("0.000")).toEqual({ numFmtId: 0, formatCode: "0.000", isDateLike: false });
    expect(numberFormatFromCode("0.0%")).toEqual({ numFmtId: 0, formatCode: "0.0%", isDateLike: false });
    expect(numberFormatFromCode("0.000%")).toEqual({ numFmtId: 0, formatCode: "0.000%", isDateLike: false });
  });

  it("日期与时间的自定义写法收下，带引号字面量的那种不收", () => {
    expect(numberFormatFromCode("yyyy-mm-dd")).toEqual({ numFmtId: 0, formatCode: "yyyy-mm-dd", isDateLike: true });
    expect(numberFormatFromCode('yyyy"年"m"月"')).toBeNull();
    expect(numberFormatFromCode("yyyy年mm月")).toBeNull();
  });

  it("认不出的码回 null，调用方必须把这句原样报给用户", () => {
    for (const code of ["", "$#,##0.00", "0.0000%", "0.0000", "??"]) expect(numberFormatFromCode(code)).toBeNull();
  });
});

describe("数字格式投影（spec 验收 A3/A4 的显示侧，换算只发生在 sheet-io）", () => {
  it("内建日期 id 与自定义格式码各按自己的写法显示同一个序列号", () => {
    const serial = 46053; // 实测样本 D2，2026-01-31
    expect(formatNumberForDisplay(serial, { numFmtId: 14, isDateLike: true })).toBe("1/31/2026");
    expect(formatNumberForDisplay(serial, { numFmtId: 164, formatCode: "yyyy-mm-dd", isDateLike: true })).toBe("2026-01-31");
    expect(formatNumberForDisplay(serial, { numFmtId: 15, isDateLike: true })).toBe("31-Jan-26");
    expect(formatNumberForDisplay(serial, { numFmtId: 17, isDateLike: true })).toBe("Jan-26");
  });

  it("带引号字面量的中文格式保留字面量，m 在 h 之后当分钟", () => {
    const format = { numFmtId: 165, formatCode: 'yyyy"年"m"月"', isDateLike: true };
    expect(formatNumberForDisplay(46053, format)).toBe("2026年1月");
    expect(formatNumberForDisplay(0.5, { numFmtId: 21, isDateLike: false })).toBe("12:00:00");
    expect(formatNumberForDisplay(0.5, { numFmtId: 18, isDateLike: false })).toBe("12:00 PM");
  });

  it("百分比与千分位按格式码缩放，普通数字保持原值", () => {
    expect(formatNumberForDisplay(0.12, { numFmtId: 10, isDateLike: false })).toBe("12.00%");
    expect(formatNumberForDisplay(0.1234, { numFmtId: 0, formatCode: "0.0%", isDateLike: false })).toBe("12.3%");
    expect(formatNumberForDisplay(1234567, { numFmtId: 3, isDateLike: false })).toBe("1,234,567");
    expect(formatNumberForDisplay(1234.5678, { numFmtId: 0, isDateLike: false })).toBe("1234.5678");
  });

  it("认不出的格式不猜：显示数值本身；非有限数原样露出来", () => {
    expect(formatNumberForDisplay(0.12, { numFmtId: 11, isDateLike: false })).toBe("0.12");
    expect(formatNumberForDisplay(Number.NaN, { numFmtId: 0, isDateLike: false })).toBe("NaN");
  });
});

describe("格子显示文本", () => {
  it("公式格显示缓存值，没有缓存值就空着（UI 另标待重算，不许显示过期数字）", () => {
    expect(cellDisplayText(cell({ formula: "SUM(C2:C4)", cachedValue: 360, numberFormat: { numFmtId: 3, isDateLike: false } }))).toBe("360");
    expect(cellDisplayText(cell({ formula: "SUM(C2:C4)", cachedValue: null }))).toBe("");
  });

  it("字符串、布尔、只有格式的格各按其语义", () => {
    expect(cellDisplayText(cell({ kind: "string", value: "华东" }))).toBe("华东");
    expect(cellDisplayText(cell({ kind: "boolean", value: true }))).toBe("TRUE");
    expect(cellDisplayText(cell({ kind: "blank", value: null, styleIndex: 3 }))).toBe("");
    expect(cellDisplayText(cell({ kind: "error", value: "#DIV/0!" }))).toBe("#DIV/0!");
  });
});

describe("用户输入解析（§2.1：日期 ↔ 序列号只在这里发生）", () => {
  it("纯数字吃成 number，且不指定格式（不改用户已有的显示）", () => {
    expect(parseUserCellText(" 120 ")).toEqual({ kind: "number", value: 120 });
    expect(parseUserCellText("1.5e3")).toEqual({ kind: "number", value: 1500 });
  });

  it("完整日期吃成序列号并带内建日期格式；两位月日这类歧义写法当文本", () => {
    expect(parseUserCellText("2026-01-31")).toEqual({
      kind: "number",
      value: 46053,
      numberFormat: { numFmtId: 14, isDateLike: true },
    });
    expect(parseUserCellText("1-31")).toEqual({ kind: "string", value: "1-31" });
  });

  it("百分号吃成小数并带百分比格式；TRUE/FALSE 吃成布尔；空串是空格", () => {
    expect(parseUserCellText("12%")).toEqual({ kind: "number", value: 0.12, numberFormat: { numFmtId: 9, isDateLike: false } });
    expect(parseUserCellText("true")).toEqual({ kind: "boolean", value: true });
    expect(parseUserCellText("   ")).toEqual({ kind: "blank", value: null });
    expect(parseUserCellText("合计 3 项")).toEqual({ kind: "string", value: "合计 3 项" });
  });

  it("百分比显示回敲进去的原话：几位小数就显示几位（内建 9/10 盖不住的走自定义码）", () => {
    for (const typed of ["12%", "-12%", "0%", "12.5%", "12.50%", "12.500%"]) {
      const parsed = parseUserCellText(typed);
      expect(cellDisplayText(cell({ kind: parsed.kind, value: parsed.value, numberFormat: parsed.numberFormat }))).toBe(typed);
    }
    // 两位以内不凭空造自定义格式，Excel 的样式表才不会为每个百分号长一条 numFmt
    expect(parseUserCellText("12%")?.numberFormat).toEqual({ numFmtId: 9, isDateLike: false });
    expect(parseUserCellText("12.50%")?.numberFormat).toEqual({ numFmtId: 10, isDateLike: false });
    expect(parseUserCellText("12.5%")?.numberFormat).toEqual({ numFmtId: 0, formatCode: "0.0%", isDateLike: false });
  });
});
