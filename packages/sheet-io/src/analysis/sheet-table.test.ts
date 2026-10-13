import { describe, expect, it } from "vitest";
import { aggregateSheet, describeSheet, duplicateGroups, findRows, readRange, resolveColumnRef } from "./sheet-table.js";
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

  it("宽表只报前 maxColumns 列，columnCount 仍是真实列数（验收 B11 的结构摘要）", () => {
    const columnCount = 60;
    const wide = sheetOf([
      Array.from({ length: columnCount }, (_, position) => `列${position + 1}`),
      Array.from({ length: columnCount }, (_, position) => position),
    ]);
    const capped = describeSheet(wide, { maxColumns: 8 });
    expect(capped.columns.map((column) => column.header)).toEqual(["列1", "列2", "列3", "列4", "列5", "列6", "列7", "列8"]);
    expect(capped).toMatchObject({ columnCount, columnsTruncated: true });
    expect(describeSheet(wide).columns.length).toBe(50);
    expect(describeSheet(sampleSheet()).columnsTruncated).toBe(false);
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

describe("readRange：上送给模型看的行样本（截断在这里收口，验收 B11）", () => {
  it("默认读满整块区域，行列数与实际给出的范围一致", () => {
    const read = readRange(sampleSheet(), { range: "A1:C3" });
    expect(read).toMatchObject({ ok: true, ref: "A1:C3", rowCount: 3, colCount: 3, truncated: false });
    expect(read.rows.map((row) => row.rowNumber)).toEqual([1, 2, 3]);
    expect(read.rows[1]!.cells).toEqual(["1月", "华东", "100"]);
  });

  it("单格写法按 1×1 区域读，空格拉成空串而不是 null", () => {
    const read = readRange(sheetOf([["月份", null], ["1月", 100]]), { range: "B1" });
    expect(read).toMatchObject({ ok: true, ref: "B1", rowCount: 1, colCount: 1 });
    expect(read.rows[0]!.cells).toEqual([""]);
  });

  it("maxRows 只截行，ref 报到实际读出的那一行、rowCount 仍是区域真实行数", () => {
    const rows: Array<[string, number]> = [["月份", "销售额"]];
    for (let month = 1; month <= 60; month += 1) rows.push([`${month}月`, month]);
    const read = readRange(sheetOf(rows), { range: "A2:B60", maxRows: 5 });
    expect(read).toMatchObject({ ok: true, ref: "A2:B6", rowCount: 59, colCount: 2, truncated: true });
    expect(read.rows.map((row) => row.rowNumber)).toEqual([2, 3, 4, 5, 6]);
    expect(read.rows[4]!.cells).toEqual(["5月", "5"]);
  });

  it("一格长文本按 120 字收过，免得一格撑爆上下文", () => {
    const read = readRange(sheetOf([["备注"], ["长".repeat(200)]]), { range: "A1:A2" });
    expect(read.rows[1]!.cells[0]!.length).toBe(121);
    expect(read.rows[1]!.cells[0]!.endsWith("…")).toBe(true);
  });

  it("带公式但没有缓存值的格显示成空串，不把公式体当数字上送", () => {
    const sheet = sheetOf([
      ["月份", "销售额"],
      ["1月", 100],
      ["2月", null],
    ]);
    sheet.cells.B3 = { address: "B3", colIndex: 1, rowIndex: 2, kind: "number", value: null, formula: "SUM(B2:B2)" };
    expect(readRange(sheet, { range: "A1:B3" }).rows[2]!.cells).toEqual(["2月", ""]);
  });

  it("请求范围超出数据区时按数据区收口，行列数报的是求交后的真实值", () => {
    const read = readRange(sampleSheet(), { range: "A1:F10" });
    expect(read).toMatchObject({ ok: true, ref: "A1:D5", rowCount: 5, colCount: 4, truncated: false });
    expect(read.rows.length).toBe(5);
  });

  it("整块都在数据区之外时给一句理由，不回几百列空格（CLI 探针实测的缺陷）", () => {
    const read = readRange(sampleSheet(), { range: "F10:H20" });
    expect(read.ok).toBe(false);
    expect(read.reason).toContain("数据区");
    expect(read.reason).toContain("A1:D5");
    expect(read.rows).toEqual([]);
    const empty = readRange(sheetOf([]), { range: "A1:B2" });
    expect(empty.ok).toBe(false);
    expect(empty.reason).toContain("空表");
  });

  it("列数同样有上限：超宽区域只读前 50 列，colCount 仍是真实列数", () => {
    const columnCount = 60;
    const wide = sheetOf([
      Array.from({ length: columnCount }, (_, position) => `列${position + 1}`),
      Array.from({ length: columnCount }, (_, position) => position),
    ]);
    const read = readRange(wide, { range: `A1:${indexToColName(columnCount - 1)}2` });
    expect(read).toMatchObject({ ok: true, ref: `A1:${indexToColName(49)}2`, colCount: columnCount, truncated: true });
    expect(read.rows[0]!.cells.length).toBe(50);
  });

  it("区域写法认不出来时给理由，不抛异常（验收 B9）", () => {
    const read = readRange(sampleSheet(), { range: "这块区域写错了" });
    expect(read.ok).toBe(false);
    expect(read.reason).toContain("这块区域写错了");
    expect(read.rows).toEqual([]);
  });
});

describe("findRows：条件判定与 COUNTIF 共用一套口径（验收 B10）", () => {
  it("数值条件命中若干行，地址与行号都是网格上看到的 1 基", () => {
    const result = findRows(sampleSheet(), { column: "销售额", criteria: ">150" });
    expect(result).toMatchObject({ ok: true, matchedCount: 2, truncated: false });
    expect(result.hits.map((hit) => `${hit.address}:${hit.display}`)).toEqual(["C3:200", "C4:300"]);
  });

  it("文本条件按相等比较，<> 取反；表头行不参与命中", () => {
    expect(findRows(sampleSheet(), { column: "区域", criteria: "华北" }).hits.map((hit) => hit.rowNumber)).toEqual([3]);
    expect(findRows(sampleSheet(), { column: "区域", criteria: "<>华东" }).matchedCount).toBe(1);
    expect(findRows(sampleSheet(), { column: "销售额", criteria: ">=100" }).hits.map((hit) => hit.rowNumber)).toEqual([2, 3, 4, 5]);
  });

  it("带公式的格用缓存值参与比较，没有缓存值就当空", () => {
    const sheet = sheetOf([
      ["月份", "销售额"],
      ["1月", 100],
      ["2月", null],
    ]);
    sheet.cells.B3 = { address: "B3", colIndex: 1, rowIndex: 2, kind: "number", value: null, formula: "SUM(B2:B2)", cachedValue: 500 };
    expect(findRows(sheet, { column: "销售额", criteria: ">150" }).hits.map((hit) => hit.rowNumber)).toEqual([3]);
    sheet.cells.B3!.cachedValue = null;
    expect(findRows(sheet, { column: "销售额", criteria: ">150" }).matchedCount).toBe(0);
  });

  it("limit 只截 hits，matchedCount 不受影响", () => {
    const result = findRows(sampleSheet(), { column: "销售额", criteria: ">=100", limit: 1 });
    expect(result).toMatchObject({ ok: true, matchedCount: 4, truncated: true });
    expect(result.hits.length).toBe(1);
  });

  it("列认不出来时给可读理由而不是崩溃（验收 B9）", () => {
    const result = findRows(sampleSheet(), { column: "库存", criteria: ">1" });
    expect(result.ok).toBe(false);
    expect(result.reason).toContain("库存");
    expect(result.hits).toEqual([]);
  });
});
