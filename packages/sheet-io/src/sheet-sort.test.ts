import { describe, expect, it } from "vitest";
import { isIdentityPermutation, sortKeyOf, sortPermutation } from "./sheet-sort.js";
import type { SheetCell } from "./model.js";

/**
 * 造一个排序键用的格子。
 *
 * @param value 格里的值；`null` 表示空位（连同 `kind:"blank"` 一起当作空格）。
 * @param kind `@t` 解出来的类型，默认按 `value` 的 JS 类型推。
 * @returns 一个只填了排序用得上的字段的格子。
 */
function cellOf(value: string | number | boolean | null, kind?: SheetCell["kind"]): SheetCell {
  const resolved = kind ?? (typeof value === "number" ? "number" : typeof value === "boolean" ? "boolean" : "string");
  return {
    address: "A1",
    colIndex: 0,
    rowIndex: 0,
    kind: value === null ? "blank" : resolved,
    value: value === null ? null : value,
  };
}

/**
 * 把一组值按同一口径排出来，读回值的顺序。
 *
 * @param values 每格一个值（`null` 是空格）。
 * @param direction 升序还是降序。
 * @returns 排好后的值队列。
 */
function ordered(values: Array<string | number | boolean | null>, direction: "asc" | "desc" = "asc"): Array<string | number | boolean | null> {
  const permutation = sortPermutation(values.map((value) => sortKeyOf(value === null ? undefined : cellOf(value))), direction);
  return permutation.map((source) => values[source] ?? null);
}

describe("排序键与置换（spec 决策 D7 的口径）", () => {
  it("数字按数值比，日期序列号也按数值（不是按显示串）", () => {
    expect(ordered([10, 9, 100])).toEqual([9, 10, 100]);
    expect(ordered([10, 9, 100], "desc")).toEqual([100, 10, 9]);
    // 45000 与 45100 是日期序列号：字典序会把 45100 排在 45000 前面，数值序不会
    expect(ordered([45100, 45000])).toEqual([45000, 45100]);
  });

  it("中文文本走 `zh-Hans-CN` 且带 `numeric`：`第2章` 排在 `第10章` 之前", () => {
    expect(ordered(["第10章", "第2章", "第1章"])).toEqual(["第1章", "第2章", "第10章"]);
    // sensitivity:"base"：大小写不参与定序
    expect(ordered(["banana", "Apple"])).toEqual(["Apple", "banana"]);
  });

  it("空格无论升降都落在带尾", () => {
    expect(ordered([3, null, 1, null, 2])).toEqual([1, 2, 3, null, null]);
    expect(ordered([3, null, 1, null, 2], "desc")).toEqual([3, 2, 1, null, null]);
    expect(ordered(["北京", null])).toEqual(["北京", null]);
  });

  it("跨类型分秩：数字 < 文本 < 布尔 < 错误值；降序把秩一起反转，只有空格恒在带尾", () => {
    // 秩来自 `kind`，所以这一组直接给格子而不是原始值
    const entries: Array<{ label: string; cell: SheetCell | undefined }> = [
      { label: "真", cell: cellOf(true) },
      { label: "空", cell: undefined },
      { label: "文本", cell: cellOf("北京") },
      { label: "二", cell: cellOf(2) },
      { label: "错误", cell: cellOf("#DIV/0!", "error") },
      { label: "一", cell: cellOf(1) },
    ];
    const read = (direction: "asc" | "desc") =>
      sortPermutation(entries.map((entry) => sortKeyOf(entry.cell)), direction).map((source) => entries[source]!.label);
    expect(read("asc")).toEqual(["一", "二", "文本", "真", "错误", "空"]);
    expect(read("desc")).toEqual(["错误", "真", "文本", "二", "一", "空"]);
  });

  it("同键的行保持原相对顺序（稳定，不依赖运行时的 sort）", () => {
    const rows = [
      { key: 10, name: "甲" },
      { key: 5, name: "乙" },
      { key: 10, name: "丙" },
      { key: 10, name: "丁" },
    ];
    const permutation = sortPermutation(rows.map((row) => sortKeyOf(cellOf(row.key))));
    expect(permutation.map((source) => rows[source]!.name)).toEqual(["乙", "甲", "丙", "丁"]);
    const descending = sortPermutation(rows.map((row) => sortKeyOf(cellOf(row.key))), "desc");
    expect(descending.map((source) => rows[source]!.name)).toEqual(["甲", "丙", "丁", "乙"]);
  });

  it("恒等置换能被认出来（调用方据此拒绝产出一条什么都不改的 op）", () => {
    expect(isIdentityPermutation(sortPermutation([1, 2, 3].map((value) => sortKeyOf(cellOf(value)))))).toBe(true);
    expect(isIdentityPermutation(sortPermutation([3, 2, 1].map((value) => sortKeyOf(cellOf(value)))))).toBe(false);
  });

  it("只有 `kind` 是 error 时按错误值秩，普通字符串按文本秩", () => {
    expect(sortKeyOf(cellOf("#DIV/0!", "error")).rank).toBeGreaterThan(sortKeyOf(cellOf("Zebra")).rank);
  });
});
