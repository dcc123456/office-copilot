import { describe, expect, it } from "vitest";
import { aggregateSheet, describeSheet, duplicateGroups, resolveColumnRef } from "./sheet-table.js";
import { indexToColName } from "../address.js";
import type { SheetCell } from "../model.js";
import type { SheetPayload } from "../sheet-dto.js";

/**
 * 造一格（数字 `kind:number`，其余 `kind:string`）。
 *
 * @param colIndex 0 基列号。
 * @param rowIndex 0 基行号。
 * @param value 值。
 * @param extra 覆盖字段。
 * @returns 格子。
 */
function cellOf(colIndex: number, rowIndex: number, value: string | number, extra: Partial<SheetCell> = {}): SheetCell {
  return {
    address: `${indexToColName(colIndex)}${rowIndex + 1}`,
    colIndex,
    rowIndex,
    kind: typeof value === "number" ? "number" : "string",
    value,
    ...extra,
  };
}

/**
 * 用二维数组造一张表（第 0 行是表头）。
 *
 * @param rows 逐行的值，null 代表这格空着。
 * @returns `SheetPayload`；已用区域按实际给出的行列算，空数组得到空表。
 */
function sheetOf(rows: Array<Array<string | number | null>>): SheetPayload {
  const cells: Record<string, SheetCell> = {};
  rows.forEach((row, rowIndex) => {
    row.forEach((value, colIndex) => {
      if (value === null) return;
      const cell = cellOf(colIndex, rowIndex, value);
      cells[cell.address] = cell;
    });
  });
  return {
    name: "销售",
    sheetId: 1,
    partPath: "xl/worksheets/sheet1.xml",
    lastColIndex: rows.reduce((acc, row) => Math.max(acc, row.length - 1), -1),
    lastRowIndex: rows.length - 1,
    cells,
    merges: [],
    columns: [],
    rows: [],
    warnings: [],
  };
}

/** 样本表：4 个数据行、两个区域、第 2 行与第 5 行整行重复。 */
function sampleSheet(): SheetPayload {
  return sheetOf([
    ["月份", "区域", "销售额", "占比"],
    ["1月", "华东", 100, 0.1],
    ["2月", "华北", 200, 0.2],
    ["3月", "华东", 300, 0.3],
    ["1月", "华东", 100, 0.1],
  ]);
}

describe("describeSheet：结构摘要 + 逐列统计", () => {
  it("表头行不进统计，行列数与已用区域对得上", () => {
    const profile = describeSheet(sampleSheet());
    expect(profile).toMatchObject({ sheetName: "销售", usedRange: "A1:D5", rowCount: 4, columnCount: 4, hasHeader: true });
    expect(profile.headerWarnings).toEqual([]);
    const sales = profile.columns[2]!;
    expect(sales).toMatchObject({ header: "销售额", type: "number", populated: 4, missing: 0, distinct: 3 });
    expect(sales.numeric).toMatchObject({ count: 4, sum: 700, min: 100, max: 300 });
  });

  it("hasHeader=false 时首行当数据，列名退回列字母", () => {
    const profile = describeSheet(sampleSheet(), { hasHeader: false });
    expect(profile.rowCount).toBe(5);
    expect(profile.columns[0]!.header).toBe("A");
    expect(profile.columns[0]!.type).toBe("string");
  });

  it("空表头与重名表头都给提示，不许静默按列号猜", () => {
    const warnings = describeSheet(sheetOf([["区域", "", "区域"], ["华东", 1, 2]])).headerWarnings.join("|");
    expect(warnings).toContain("B 列表头是空的");
    expect(warnings).toContain("C 列与 A 列表头同名");
  });

  it("空表返回 rowCount 0 / columns 空，不抛异常", () => {
    const profile = describeSheet(sheetOf([]));
    expect(profile.rowCount).toBe(0);
    expect(profile.columnCount).toBe(0);
    expect(profile.usedRange).toBeUndefined();
    expect(profile.columns).toEqual([]);
  });
});

describe("resolveColumnRef：列字母与表头都认", () => {
  it("字母、表头原文、重名的序号三种写法", () => {
    const sheet = sampleSheet();
    expect(resolveColumnRef(sheet, "c", true)).toBe(2);
    expect(resolveColumnRef(sheet, "销售额", true)).toBe(2);
    const doubled = sheetOf([["区域", "合计", "区域"], ["华东", 1, 2]]);
    expect(resolveColumnRef(doubled, "区域", true)).toBe(0);
    expect(resolveColumnRef(doubled, "区域#2", true)).toBe(2);
  });

  it("认不出来返回 null（工具层据此回错误串，不是崩溃）", () => {
    expect(resolveColumnRef(sampleSheet(), "库存", true)).toBeNull();
    expect(resolveColumnRef(sampleSheet(), "", true)).toBeNull();
    expect(resolveColumnRef(sampleSheet(), "Z", true)).toBeNull();
  });
});

describe("aggregateSheet：数字由引擎算（验收 B10）", () => {
  it("按区域分组求和/计数，组内行号给到 1 基", () => {
    const result = aggregateSheet(sampleSheet(), {
      groupBy: ["区域"],
      aggregates: [
        { fn: "sum", column: "销售额" },
        { fn: "count", column: "销售额" },
      ],
    });
    expect(result.ok).toBe(true);
    expect(result.groupCount).toBe(2);
    const east = result.groups.find((group) => group.key === "华东")!;
    expect(east.rowNumbers).toEqual([2, 4, 5]);
    expect(east.values).toEqual({ "sum(销售额)": 500, "count(销售额)": 3 });
    expect(east.valueTexts["sum(销售额)"]).toBe("500");
  });

  it("average 跳过文本与空格，只按数字格算；rowCount 仍是组内行数", () => {
    const result = aggregateSheet(sheetOf([["组", "值"], ["A", 10], ["A", "待补"], ["A", 20], ["A", null]]), {
      groupBy: ["组"],
      aggregates: [
        { fn: "average", column: "值" },
        { fn: "count", column: "值" },
      ],
    });
    expect(result.groups[0]!.values).toEqual({ "average(值)": 15, "count(值)": 2 });
    expect(result.groups[0]!.rowCount).toBe(4);
  });

  it("一组里没有任何数字时 average 给 null，显示串说明原因而不是 0", () => {
    const result = aggregateSheet(sheetOf([["组", "值"], ["A", "文字"], ["A", "更多"]]), {
      groupBy: ["组"],
      aggregates: [{ fn: "average", column: "值" }],
    });
    expect(result.groups[0]!.values["average(值)"]).toBeNull();
    expect(result.groups[0]!.valueTexts["average(值)"]).toBe("（无可用数字）");
  });

  it("列名认不出来给可读理由，groups 空", () => {
    const result = aggregateSheet(sampleSheet(), { groupBy: ["部门"], aggregates: [{ fn: "sum", column: "库存" }] });
    expect(result.ok).toBe(false);
    expect(result.reason).toContain("部门");
    expect(result.reason).toContain("库存");
    expect(result.groups).toEqual([]);
  });

  it("limit 截断组数但如实报 groupCount 与 truncated", () => {
    const result = aggregateSheet(sampleSheet(), { groupBy: ["月份"], aggregates: [{ fn: "sum", column: "销售额" }], limit: 2 });
    expect(result.groups.length).toBe(2);
    expect(result.groupCount).toBe(3);
    expect(result.truncated).toBe(true);
  });

  it("浮点尾差收进 15 位有效数字，0.1+0.2 不算出 0.30000000000000004", () => {
    const result = aggregateSheet(sheetOf([["组", "值"], ["A", 0.1], ["A", 0.2]]), {
      groupBy: ["组"],
      aggregates: [
        { fn: "sum", column: "值" },
        { fn: "average", column: "值" },
      ],
    });
    expect(result.groups[0]!.values["sum(值)"]).toBe(0.3);
    expect(result.groups[0]!.values["average(值)"]).toBe(0.15);
  });

  it("百分比列的聚合值按该列格式显示", () => {
    const percent = { numFmtId: 10, isDateLike: false };
    const sheet = sheetOf([
      ["组", "占比"],
      ["A", 0.125],
      ["A", 0.25],
    ]);
    sheet.cells.B2!.numberFormat = percent;
    sheet.cells.B3!.numberFormat = percent;
    const result = aggregateSheet(sheet, { groupBy: ["组"], aggregates: [{ fn: "sum", column: "占比" }] });
    expect(result.groups[0]!.valueTexts["sum(占比)"]).toBe("37.50%");
  });
});

describe("duplicateGroups：去重键的行数算得出来", () => {
  it("整行重复只算一次多余，redundantRowCount 就是能删掉的行数", () => {
    const result = duplicateGroups(sampleSheet(), { keys: ["月份", "区域", "销售额"] });
    expect(result.ok).toBe(true);
    expect(result.groups).toEqual([{ key: "1月 / 华东 / 100", rowNumbers: [2, 5], count: 2 }]);
    expect(result.redundantRowCount).toBe(1);
  });

  it("单键分组时同月的不同区域也算同组", () => {
    const result = duplicateGroups(sampleSheet(), { keys: ["月份"] });
    expect(result.groups.map((group) => `${group.key}×${group.count}`)).toEqual(["1月×2"]);
    expect(result.redundantRowCount).toBe(1);
  });

  it("没有任何重复时给空数组而不是 null", () => {
    expect(duplicateGroups(sheetOf([["月份"], ["1月"], ["2月"]]), { keys: ["月份"] })).toEqual({ ok: true, groups: [], redundantRowCount: 0 });
  });

  it("键列认不出来时同样给理由", () => {
    const result = duplicateGroups(sampleSheet(), { keys: ["库存"] });
    expect(result.ok).toBe(false);
    expect(result.reason).toContain("库存");
  });
});
