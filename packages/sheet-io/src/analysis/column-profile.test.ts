import { describe, expect, it } from "vitest";
import { cellNumberValue, numericSummary, profileColumn, quantileInc } from "./column-profile.js";
import type { SheetCell } from "../model.js";

/**
 * 造一格。
 *
 * @param letter 列字母（A、B…）。
 * @param rowNumber 1 基行号，与网格里看到的行号一致。
 * @param value 格子里的值；null 表示只有格式的空格。
 * @param extra 覆盖字段（`kind`/`numberFormat`/`formula`/`cachedValue`）。
 * @returns 一个够画像层用的格子。
 */
function cellOf(letter: string, rowNumber: number, value: string | number | boolean | null, extra: Partial<SheetCell> = {}): SheetCell {
  return {
    address: `${letter}${rowNumber}`,
    colIndex: letter.charCodeAt(0) - 65,
    rowIndex: rowNumber - 1,
    kind: value === null ? "blank" : typeof value === "number" ? "number" : typeof value === "boolean" ? "boolean" : "string",
    value,
    ...extra,
  };
}

/** 带公式的格子：`kind` 必须是 number，缓存值缺省时代表"还没重算"。 */
function formulaCell(letter: string, rowNumber: number, formula: string, cachedValue?: number): SheetCell {
  return cellOf(letter, rowNumber, null, { kind: "number", formula, ...(cachedValue === undefined ? {} : { cachedValue }) });
}

describe("quantileInc：与 Excel PERCENTILE.INC 同口径", () => {
  it("四个数取四分位是线性插值，不是取整", () => {
    expect(quantileInc([1, 2, 3, 4], 0.25)).toBe(1.75);
    expect(quantileInc([1, 2, 3, 4], 0.5)).toBe(2.5);
    expect(quantileInc([1, 2, 3, 4], 0.75)).toBe(3.25);
  });

  it("一个数时任何分位都等于它本身，空数组返回 null 而不是 0", () => {
    expect(quantileInc([7], 0.25)).toBe(7);
    expect(quantileInc([], 0.5)).toBeNull();
  });
});

describe("numericSummary：Excel 统计口径的描述统计", () => {
  it("乱序输入自己排好再取中位数", () => {
    expect(numericSummary([4, 1, 3, 2])!.median).toBe(2.5);
  });

  it("给出 sum/average/min/max/median 与 Tukey 栅栏", () => {
    const summary = numericSummary([10, 12, 11, 13, 100]);
    expect(summary).toMatchObject({ count: 5, sum: 146, min: 10, max: 100 });
    expect(summary!.average).toBeCloseTo(29.2, 9);
    expect(summary!.median).toBe(12);
    expect(summary!.q1).toBe(11);
    expect(summary!.q3).toBe(13);
    expect(summary!.iqr).toBe(2);
    expect(summary!.upperFence).toBe(16);
    expect(summary!.lowerFence).toBe(8);
  });

  it("空集合返回 null：不许用 0 冒充「没有可统计的数字」", () => {
    expect(numericSummary([])).toBeNull();
  });
});

describe("cellNumberValue：只吃数字", () => {
  it("文本与布尔返回 null，公式格取缓存值", () => {
    expect(cellNumberValue(cellOf("A", 1, "华东"))).toBeNull();
    expect(cellNumberValue(cellOf("A", 1, true))).toBeNull();
    expect(cellNumberValue(null)).toBeNull();
    expect(cellNumberValue(formulaCell("A", 1, "SUM(B1:B3)", 360))).toBe(360);
  });

  it("带公式但没有缓存值时返回 null（还没重算，不许拿过期数字统计）", () => {
    expect(cellNumberValue(formulaCell("A", 1, "TODAY()"))).toBeNull();
  });
});

describe("profileColumn：类型推断、缺失与异常值", () => {
  it("数字列标出离群值并给出栅栏外的计数", () => {
    const profile = profileColumn({
      index: 2,
      letter: "C",
      header: "销售额",
      cells: [cellOf("C", 2, 100), cellOf("C", 3, 102), cellOf("C", 4, 98), cellOf("C", 5, 5000), null],
    });
    expect(profile.type).toBe("number");
    expect(profile.missing).toBe(1);
    expect(profile.numeric!.count).toBe(4);
    expect(profile.outlierCount).toBe(1);
    expect(profile.outliers).toEqual([{ address: "C5", display: "5000", value: 5000, side: "high" }]);
  });

  it("日期列不进数字统计，只给区间", () => {
    const dateFormat = { numFmtId: 14, isDateLike: true };
    const profile = profileColumn({
      index: 3,
      letter: "D",
      header: "日期",
      cells: [cellOf("D", 2, 46053, { numberFormat: dateFormat }), cellOf("D", 3, 46054, { numberFormat: dateFormat })],
    });
    expect(profile.type).toBe("date");
    expect(profile.numeric).toBeUndefined();
    expect(profile.dateRange).toEqual({ from: "1/31/2026", to: "2/1/2026" });
  });

  it("同一列里数字与文本并存判成 mixed，不给其中任何一种口径的数字", () => {
    const profile = profileColumn({ index: 0, letter: "A", header: "混合", cells: [cellOf("A", 2, 1), cellOf("A", 3, "待补"), cellOf("A", 4, 2)] });
    expect(profile.type).toBe("mixed");
    expect(profile.numeric).toBeUndefined();
    expect(profile.distinct).toBe(3);
    expect(profile.topValues.map((item) => item.text)).toEqual(["1", "2", "待补"]);
  });

  it("全空列 type=empty，统计量都是 0，不会给出假的平均值", () => {
    const profile = profileColumn({ index: 7, letter: "H", header: "备注", cells: [null, null, cellOf("H", 2, null)] });
    expect(profile.type).toBe("empty");
    expect(profile.numeric).toBeUndefined();
    expect(profile.outliers).toEqual([]);
    expect(profile.missing).toBe(3);
  });

  it("带公式但没有缓存值的格数记进 pendingRecalc", () => {
    const profile = profileColumn({
      index: 6,
      letter: "G",
      header: "合计",
      cells: [formulaCell("G", 2, "SUM(C2:C4)", 360), formulaCell("G", 3, "SUM(C5:C7)")],
    });
    expect(profile.pendingRecalc).toBe(1);
    expect(profile.numeric?.count).toBe(1);
  });

  it("百分比列的显示串按该列格式渲染，上层不再换算第二次", () => {
    const percent = { numFmtId: 10, isDateLike: false };
    const profile = profileColumn({
      index: 1,
      letter: "B",
      header: "占比",
      cells: [cellOf("B", 2, 0.125, { numberFormat: percent }), cellOf("B", 3, 0.25, { numberFormat: percent })],
    });
    expect(profile.numericText?.sum).toBe("37.50%");
    expect(profile.numericText?.average).toBe("18.75%");
  });
});
