import { describe, expect, it } from "vitest";
import { colNameToIndex, formatA1, formatRange, indexToColName, parseA1, parseRange, rangeContains } from "./address.js";

describe("A1 记法换算（全仓唯一一处）", () => {
  it("列名与下标双向一致", () => {
    expect(colNameToIndex("A")).toBe(0);
    expect(colNameToIndex("Z")).toBe(25);
    expect(colNameToIndex("AA")).toBe(26);
    expect(colNameToIndex("AZ")).toBe(51);
    expect(indexToColName(0)).toBe("A");
    expect(indexToColName(26)).toBe("AA");
    for (const name of ["A", "M", "Z", "AA", "BQ", "ZZ"]) {
      expect(indexToColName(colNameToIndex(name) ?? -1)).toBe(name);
    }
  });

  it("读不懂的列名返回 undefined，不猜一个列", () => {
    expect(colNameToIndex("A1")).toBeUndefined();
    expect(colNameToIndex("")).toBeUndefined();
    expect(colNameToIndex("ABCD")).toBeUndefined();
    expect(indexToColName(-1)).toBe("");
  });

  it("绝对引用 $B$7 与 b7 都落到同一个格", () => {
    expect(parseA1("$B$7")).toEqual({ colIndex: 1, rowIndex: 6 });
    expect(parseA1("b7")).toEqual({ colIndex: 1, rowIndex: 6 });
    expect(parseA1("7B")).toBeUndefined();
    expect(formatA1({ colIndex: 1, rowIndex: 6 })).toBe("B7");
  });

  it("区域归一化成左上/右下，整列按传入的行数收口", () => {
    expect(parseRange("C2:A1")).toEqual({ start: { colIndex: 0, rowIndex: 0 }, end: { colIndex: 2, rowIndex: 1 } });
    expect(parseRange("A1")).toEqual({ start: { colIndex: 0, rowIndex: 0 }, end: { colIndex: 0, rowIndex: 0 } });
    expect(parseRange("A:B", 3)).toEqual({ start: { colIndex: 0, rowIndex: 0 }, end: { colIndex: 1, rowIndex: 3 } });
    expect(parseRange("2:4")).toEqual({ start: { colIndex: 0, rowIndex: 1 }, end: { colIndex: 16383, rowIndex: 3 } });
    expect(parseRange("A1-C3")).toBeUndefined();
    expect(formatRange(parseRange("A1:C3")!)).toBe("A1:C3");
    expect(formatRange(parseRange("B2")!)).toBe("B2");
  });

  it("边界含在区域里", () => {
    const range = parseRange("B2:C3")!;
    expect(rangeContains(range, parseA1("B2")!)).toBe(true);
    expect(rangeContains(range, parseA1("C3")!)).toBe(true);
    expect(rangeContains(range, parseA1("D3")!)).toBe(false);
  });
});
