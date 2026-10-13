/**
 * 表级分析：结构画像、分组聚合与重复行定位（spec 阶段 4「分析数据 / 数据整理」的引擎侧）。
 *
 * 三条约束决定了本文件的形状：
 * - **数字由这里算**（`AGENTS.md` §2.8、spec §零.4，验收 B10）：模型只允许指定"按哪几列分组、对哪列做什么聚合"。
 * - **认不出来给理由而不是崩溃**（验收 B9）：列引用解析失败返回 `ok:false` + `reason`，工具层把这条错误串回给模型。
 * - **上送内容可控**（spec §零.5）：分组报告默认只回前 `GROUP_REPORT_LIMIT` 组，其余只给计数。
 * 显示串一律走 `cellDisplayText`/`formatNumberForDisplay`，本文件不自己换算单位（§2.1）。
 */
import { colNameToIndex, formatRange, indexToColName } from "../address.js";
import { cellDisplayText, formatNumberForDisplay } from "../cell-text.js";
import type { SheetCell } from "../model.js";
import type { SheetPayload } from "../sheet-dto.js";
import { cellNumberValue, profileColumn, type ColumnProfile } from "./column-profile.js";

/** 允许的聚合函数（刻意只留这五个，其余交给 Excel 公式，见 spec 决策 D4）。 */
export type AggregateFunction = "count" | "sum" | "average" | "min" | "max";

/** 一个聚合：作用在哪一列（列字母或表头文字）+ 用什么函数。 */
export interface AggregateSpec {
  fn: AggregateFunction;
  column: string;
}

/** `describeSheet` 的输入选项。 */
export interface DescribeOptions {
  /** 首行是不是表头（默认 true）。为 true 时首行不进统计，列名取首行文本。 */
  hasHeader?: boolean;
}

/** `aggregateSheet` 的输入。 */
export interface AggregateInput {
  /** 分组列，1~3 个（再多就不是"分析"而是"把整表搬过来"了）。 */
  groupBy: string[];
  aggregates: AggregateSpec[];
  hasHeader?: boolean;
  /** 返回的分组数上限，默认 `GROUP_REPORT_LIMIT`；超出的组只报计数。 */
  limit?: number;
}

/** 一张表的画像。 */
export interface SheetProfile {
  sheetName: string;
  partPath: string;
  /** 已用区域的 A1 写法；空表为 undefined。 */
  usedRange?: string;
  /** 数据行数（不含表头）。 */
  rowCount: number;
  columnCount: number;
  hasHeader: boolean;
  /** 表头缺失或重复时的提示，上层必须带给人看，不许静默按列号猜。 */
  headerWarnings: string[];
  columns: ColumnProfile[];
}

/** 一个分组：键 + 组内行号 + 引擎算出的聚合值。 */
export interface AggregateGroup {
  /** 分组键的连接文本，给人看。 */
  key: string;
  groupValues: string[];
  /** 组内数据行数。 */
  rowCount: number;
  /** 组内行号（1 基，与网格里看到的行号一致）。 */
  rowNumbers: number[];
  /** 每个聚合的标签（`sum(C)` 这类）→ 数值；该组没有可用数字时为 null（不许用 0 冒充结果）。 */
  values: Record<string, number | null>;
  /** 同一批聚合值的显示串（按该列代表性格式渲染，上层不再换算）。 */
  valueTexts: Record<string, string>;
}

/** `aggregateSheet` 的结果。 */
export interface AggregateResult {
  ok: boolean;
  /** 失败原因（列引用认不出来）；成功时为 undefined。 */
  reason?: string;
  groups: AggregateGroup[];
  /** 分组总数（`groups` 可能被 `limit` 截断）。 */
  groupCount: number;
  truncated: boolean;
}

/** 一组重复行。 */
export interface DuplicateGroup {
  key: string;
  rowNumbers: number[];
  /** 重复次数（= `rowNumbers.length`）。 */
  count: number;
}

/** `duplicateGroups` 的结果。 */
export interface DuplicateResult {
  ok: boolean;
  reason?: string;
  groups: DuplicateGroup[];
  /** 每组多出来的行数之和，即"按这些键去重能安全删掉几行"。 */
  redundantRowCount: number;
}

/** 分组报告的默认条数上限（与 `column-profile` 同一口径：只报前若干条，其余给计数）。 */
const GROUP_REPORT_LIMIT = 50;

/**
 * 把浮点尾差收进 15 位有效数字（Excel 的显示口径，`formatNumberForDisplay` 也这么收）。
 *
 * @param value 引擎算出的原始数值。
 * @returns 收过的数值；非有限数原样返回。
 */
function tidy(value: number): number {
  return Number.isFinite(value) ? Number(value.toPrecision(15)) : value;
}

/**
 * 数据行的行号区间（0 基）。
 *
 * @param sheet 表载荷。
 * @param hasHeader 首行是否算表头。
 * @returns `{ from, to }`；空表返回 `{ from: 0, to: -1 }`，即一条数据行都没有。
 */
export function dataRowRange(sheet: SheetPayload, hasHeader: boolean): { from: number; to: number } {
  if (sheet.lastColIndex < 0 || sheet.lastRowIndex < 0) return { from: 0, to: -1 };
  return { from: hasHeader ? 1 : 0, to: sheet.lastRowIndex };
}

/**
 * 取一列在数据行上的格子，按行序排好（缺的格用 null 占位，缺失统计要数得出空行）。
 *
 * @param sheet 表载荷。
 * @param colIndex 列号（0 基）。
 * @param hasHeader 首行是否算表头。
 * @returns 与数据行等长的格子数组。
 */
export function columnCellsOf(sheet: SheetPayload, colIndex: number, hasHeader: boolean): Array<SheetCell | null> {
  const { from, to } = dataRowRange(sheet, hasHeader);
  const letter = indexToColName(colIndex);
  const rows: Array<SheetCell | null> = [];
  for (let rowIndex = from; rowIndex <= to; rowIndex += 1) rows.push(sheet.cells[`${letter}${rowIndex + 1}`] ?? null);
  return rows;
}

/**
 * 表头里等于某个名字的列号（按出现顺序）。
 *
 * @param sheet 表载荷。
 * @param name 表头文字（比较时去首尾空白、不区分大小写）。
 * @param hasHeader 首行是否算表头；false 时没有任何"表头"，返回空数组。
 * @returns 命中的 0 基列号数组。
 */
function headerIndexesOf(sheet: SheetPayload, name: string, hasHeader: boolean): number[] {
  if (!hasHeader || sheet.lastRowIndex < 0) return [];
  const target = name.trim().toLowerCase();
  if (target === "") return [];
  const indexes: number[] = [];
  for (let colIndex = 0; colIndex <= sheet.lastColIndex; colIndex += 1) {
    const cell = sheet.cells[`${indexToColName(colIndex)}1`];
    if (cell && cellDisplayText(cell).trim().toLowerCase() === target) indexes.push(colIndex);
  }
  return indexes;
}

/**
 * 把模型给的列引用解析成列号。
 *
 * 支持三种写法：列字母（`C`/`c`）、表头文字（`销售额`，去首尾空白后不区分大小写）、
 * 表头重名时的第 N 次出现（`销售额#2`，1 基）。
 *
 * @param sheet 表载荷。
 * @param ref 列引用。
 * @param hasHeader 首行是否算表头（决定表头文字从哪一行找）。
 * @returns 0 基列号；解析不出来返回 null（调用方必须把 null 变成给模型看的错误串）。
 */
export function resolveColumnRef(sheet: SheetPayload, ref: string, hasHeader: boolean): number | null {
  const trimmed = ref.trim();
  if (trimmed === "") return null;
  const letterMatch = /^[A-Za-z]+$/.exec(trimmed);
  // 纯字母先看是不是表头文字（表头本身可能全是字母），不是才当列号。
  if (letterMatch && headerIndexesOf(sheet, trimmed, hasHeader).length === 0) {
    const index = colNameToIndex(trimmed.toUpperCase());
    if (index !== undefined && index <= sheet.lastColIndex) return index;
  }
  const ordinal = /#(\d+)$/.exec(trimmed);
  const indexes = headerIndexesOf(sheet, trimmed.replace(/#\d+$/, "").trim(), hasHeader);
  if (indexes.length === 0) return null;
  return indexes[ordinal ? Number(ordinal[1]) - 1 : 0] ?? null;
}

/**
 * 生成表画像：逐列的类型推断与统计（数字一律由 `column-profile` 算，模型不参与心算）。
 *
 * @param sheet 表载荷。
 * @param options 见 `DescribeOptions`；默认把首行当表头。
 * @returns `SheetProfile`；空表返回 `rowCount: 0`、`columns: []`，不抛异常。
 */
export function describeSheet(sheet: SheetPayload, options: DescribeOptions = {}): SheetProfile {
  const hasHeader = options.hasHeader ?? true;
  const empty = sheet.lastColIndex < 0 || sheet.lastRowIndex < 0;
  const columns: ColumnProfile[] = [];
  const headerWarnings: string[] = [];
  const seen = new Map<string, number>();

  if (!empty) {
    for (let colIndex = 0; colIndex <= sheet.lastColIndex; colIndex += 1) {
      const letter = indexToColName(colIndex);
      const headerCell = sheet.cells[`${letter}1`];
      const header = hasHeader && headerCell ? cellDisplayText(headerCell).trim() : "";
      if (hasHeader && header === "") headerWarnings.push(`${letter} 列表头是空的，只能按列号引用`);
      if (header !== "") {
        const previous = seen.get(header.toLowerCase());
        if (previous !== undefined)
          headerWarnings.push(`${letter} 列与 ${indexToColName(previous)} 列表头同名「${header}」，引用要写成「${header}#2」这类带序号的形式`);
        seen.set(header.toLowerCase(), colIndex);
      }
      columns.push(profileColumn({ index: colIndex, letter, header: header || letter, cells: columnCellsOf(sheet, colIndex, hasHeader) }));
    }
  }

  const { from, to } = dataRowRange(sheet, hasHeader);
  const range = empty ? undefined : formatRange({ start: { colIndex: 0, rowIndex: 0 }, end: { colIndex: sheet.lastColIndex, rowIndex: sheet.lastRowIndex } });
  return {
    sheetName: sheet.name,
    partPath: sheet.partPath,
    ...(range ? { usedRange: range } : {}),
    rowCount: empty ? 0 : to - from + 1,
    columnCount: empty ? 0 : sheet.lastColIndex + 1,
    hasHeader,
    headerWarnings,
    columns,
  };
}

/**
 * 按若干列把数据行分桶。
 *
 * @param sheet 表载荷。
 * @param columns 已解析的分组列号（0 基）。
 * @param hasHeader 首行是否算表头。
 * @returns 键 → `{ groupValues, rowNumbers }`（保持首次出现顺序）。
 */
function bucketRows(sheet: SheetPayload, columns: number[], hasHeader: boolean): Map<string, { groupValues: string[]; rowNumbers: number[] }> {
  const { from, to } = dataRowRange(sheet, hasHeader);
  const letters = columns.map((index) => indexToColName(index));
  const buckets = new Map<string, { groupValues: string[]; rowNumbers: number[] }>();
  for (let rowIndex = from; rowIndex <= to; rowIndex += 1) {
    const values = columns.map((index, position) => {
      const cell = sheet.cells[`${letters[position]}${rowIndex + 1}`] ?? null;
      return cell ? cellDisplayText(cell).trim() : "";
    });
    const key = values.join(" / ");
    const bucket = buckets.get(key);
    if (bucket) bucket.rowNumbers.push(rowIndex + 1);
    else buckets.set(key, { groupValues: values, rowNumbers: [rowIndex + 1] });
  }
  return buckets;
}

/**
 * 算一组行在某个聚合下的值。
 *
 * @param sheet 表载荷。
 * @param colIndex 数值列（0 基）。
 * @param rowNumbers 该组内数据行号（1 基）。
 * @param fn 聚合函数。
 * @returns 数值；该组一个可用数字都没有且函数不是 `count` 时返回 null（不许返回 0 冒充结果）。
 */
function aggregateValue(sheet: SheetPayload, colIndex: number, rowNumbers: number[], fn: AggregateFunction): number | null {
  const letter = indexToColName(colIndex);
  const values = rowNumbers
    .map((rowNumber) => cellNumberValue(sheet.cells[`${letter}${rowNumber}`] ?? null))
    .filter((value): value is number => value !== null);
  if (fn === "count") return values.length;
  if (values.length === 0) return null;
  if (fn === "sum") return tidy(values.reduce((acc, value) => acc + value, 0));
  if (fn === "average") return tidy(values.reduce((acc, value) => acc + value, 0) / values.length);
  return tidy(fn === "min" ? Math.min(...values) : Math.max(...values));
}

/**
 * 把一个聚合值显示出来：沿用该列第一个数值格的数字格式（换算仍在 `sheet-io` 内，§2.1）。
 *
 * @param value 引擎算出的数值。
 * @param sheet 表载荷。
 * @param colIndex 该聚合作用的列（0 基）。
 * @param hasHeader 首行是否算表头。
 * @returns 给人看的显示串；该列没有可读格式时按 General 显示。
 */
function displayAggregate(value: number, sheet: SheetPayload, colIndex: number, hasHeader: boolean): string {
  const format = columnCellsOf(sheet, colIndex, hasHeader).find((cell) => cellNumberValue(cell) !== null)?.numberFormat;
  return formatNumberForDisplay(tidy(value), format);
}

/**
 * 分组聚合：把 `groupBy` 指定的列当键，对 `aggregates` 指定的列算数字。
 *
 * @param sheet 表载荷。
 * @param input 见 `AggregateInput`。
 * @returns `AggregateResult`；列认不出来时 `ok: false` + `reason`，`groups` 为空。
 */
export function aggregateSheet(sheet: SheetPayload, input: AggregateInput): AggregateResult {
  const hasHeader = input.hasHeader ?? true;
  const limit = input.limit ?? GROUP_REPORT_LIMIT;
  const unrecognized = [...input.groupBy, ...input.aggregates.map((spec) => spec.column)].filter((ref) => resolveColumnRef(sheet, ref, hasHeader) === null);
  if (unrecognized.length > 0) {
    return { ok: false, reason: `认不出这些列：${unrecognized.join("、")}。请改用列字母（如 C）或表头原文。`, groups: [], groupCount: 0, truncated: false };
  }

  const buckets = [...bucketRows(sheet, input.groupBy.map((ref) => resolveColumnRef(sheet, ref, hasHeader)!), hasHeader).values()];
  const sorted = buckets.sort((left, right) => right.rowNumbers.length - left.rowNumbers.length);
  const groups = sorted.slice(0, limit).map((bucket) => {
    const values: Record<string, number | null> = {};
    const valueTexts: Record<string, string> = {};
    input.aggregates.forEach((spec, position) => {
      const colIndex = resolveColumnRef(sheet, spec.column, hasHeader)!;
      const label = `${spec.fn}(${spec.column || indexToColName(colIndex)})`;
      const value = aggregateValue(sheet, colIndex, bucket.rowNumbers, spec.fn);
      values[label] = value;
      valueTexts[label] = value === null ? "（无可用数字）" : displayAggregate(value, sheet, colIndex, hasHeader);
    });
    return {
      key: bucket.groupValues.join(" / "),
      groupValues: bucket.groupValues,
      rowCount: bucket.rowNumbers.length,
      rowNumbers: bucket.rowNumbers,
      values,
      valueTexts,
    };
  });

  return { ok: true, groups, groupCount: sorted.length, truncated: sorted.length > groups.length };
}

/**
 * 找出键列完全相同的重复行组（数据整理里的"去重"判据）。
 *
 * @param sheet 表载荷。
 * @param input `keys` 要求全行一致的列（1~5 个）；`hasHeader` 同上；`limit` 只影响返回条数，计数不受影响。
 * @returns `DuplicateResult`；`redundantRowCount` = 每组多出来的行数之和（即"能安全删掉几行"）。
 */
export function duplicateGroups(sheet: SheetPayload, input: { keys: string[]; hasHeader?: boolean; limit?: number }): DuplicateResult {
  const hasHeader = input.hasHeader ?? true;
  const limit = input.limit ?? GROUP_REPORT_LIMIT;
  const unrecognized = input.keys.filter((ref) => resolveColumnRef(sheet, ref, hasHeader) === null);
  if (unrecognized.length > 0) return { ok: false, reason: `认不出这些列：${unrecognized.join("、")}。请改用列字母或表头原文。`, groups: [], redundantRowCount: 0 };

  const columns = input.keys.map((ref) => resolveColumnRef(sheet, ref, hasHeader)!);
  const duplicated = [...bucketRows(sheet, columns, hasHeader).values()]
    .filter((bucket) => bucket.rowNumbers.length > 1)
    .sort((left, right) => right.rowNumbers.length - left.rowNumbers.length);
  return {
    ok: true,
    groups: duplicated.slice(0, limit).map((bucket) => ({ key: bucket.groupValues.join(" / "), rowNumbers: bucket.rowNumbers, count: bucket.rowNumbers.length })),
    redundantRowCount: duplicated.reduce((sum, bucket) => sum + bucket.rowNumbers.length - 1, 0),
  };
}
