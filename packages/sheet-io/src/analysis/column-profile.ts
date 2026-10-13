/**
 * 一列的画像：类型推断、缺失、众数、分位数与异常值（spec 阶段 4「分析数据」的引擎侧）。
 *
 * 为什么数字在这里算：`AGENTS.md` §2.8/spec §零.4 规定模型只允许指定"对哪一列做什么聚合"，
 * 实际数值由 `sheet-io` 算出——模型心算的数字与引擎不符时以引擎为准。
 * 统计口径照 Excel 的统计函数：**只吃数字**，跳过空、文本、布尔与逻辑值；
 * 分位数用 `PERCENTILE.INC` 的线性插值（rank = k·(n−1)），异常值用 Tukey 栅栏（Q1−1.5·IQR、Q3+1.5·IQR）。
 * 显示串一律走 `formatNumberForDisplay`/`cellDisplayText`，本文件不自己换算单位（§2.1）。
 */
import { cellDisplayText, formatNumberForDisplay } from "../cell-text.js";
import type { ResolvedNumberFormat, SheetCell } from "../model.js";

/** 一列的主导值类型。`mixed` = 同一列里出现两种以上非空类型，整理阶段要按它决定怎么清洗。 */
export type ColumnValueType = "empty" | "number" | "date" | "boolean" | "string" | "mixed";

/** 数字列的描述统计。全部按 Excel 口径：`count` 只数参与统计的数字个数。 */
export interface NumericSummary {
  count: number;
  sum: number;
  average: number;
  min: number;
  max: number;
  median: number;
  q1: number;
  q3: number;
  iqr: number;
  lowerFence: number;
  upperFence: number;
}

/** 一个离群值：地址 + 给人看的文本 + 判定它离群的那一侧。 */
export interface ColumnOutlier {
  address: string;
  display: string;
  value: number;
  side: "low" | "high";
}

/** profileColumn 的输入：某一列在数据行上的格子（空格用 null，位置必须按行序排好）。 */
export interface ColumnCells {
  index: number;
  letter: string;
  header: string;
  cells: Array<SheetCell | null>;
}

/** 一列的完整画像。`numeric`/`numericText`/`dateRange` 只在对应类型上出现。 */
export interface ColumnProfile {
  index: number;
  letter: string;
  header: string;
  type: ColumnValueType;
  /** 非空格数。 */
  populated: number;
  /** 数据行里空着的格数（含只有格式的格）。 */
  missing: number;
  /** 非空格的互异显示文本个数。 */
  distinct: number;
  /** 出现最多的前几个值及其次数；`mixed`/`string` 列靠它看数据长什么样。 */
  topValues: Array<{ text: string; count: number }>;
  /** 带公式但没有缓存值的格数：这些格参与不了统计，Excel 重算后数字可能变化，必须让上层看见。 */
  pendingRecalc: number;
  numeric?: NumericSummary;
  /** 同一批统计量的显示串（按该列代表性数字格式渲染），上层拿到就不再换算。 */
  numericText?: Record<keyof NumericSummary, string>;
  dateRange?: { from: string; to: string };
  outliers: ColumnOutlier[];
  /** 栅栏外的格总数（`outliers` 只截前 `OUTLIER_REPORT_LIMIT` 条）。 */
  outlierCount: number;
}

/** 报告里最多列几条异常值；再多只给计数，免得画像本身把上下文撑爆（spec §零.5 的同一口径）。 */
const OUTLIER_REPORT_LIMIT = 10;
/** 众数列表的条数上限。 */
const TOP_VALUE_LIMIT = 5;

/**
 * 取一格用于统计的数值。
 *
 * @param cell 格子（null 表示空）。
 * @returns 有限数字；空格、文本、布尔、以及"带公式但没有缓存值"的格一律返回 null（Excel 也这么跳）。
 */
export function cellNumberValue(cell: SheetCell | null): number | null {
  if (!cell || cell.kind === "blank" || cell.kind === "error") return null;
  const raw = cell.formula ? cell.cachedValue : cell.value;
  if (typeof raw !== "number" || !Number.isFinite(raw)) return null;
  return raw;
}

/**
 * 这格是不是按日期显示的数字（序列号 ↔ 日期的判定只读 reader 已经解好的 `numberFormat`，本文件不换算）。
 *
 * @param cell 格子（null 表示空）。
 * @returns 数字型且 `isDateLike` 时为 true。
 */
function cellIsDate(cell: SheetCell | null): boolean {
  if (!cell || cell.kind !== "number") return false;
  return Boolean(cell.numberFormat?.isDateLike) && cellNumberValue(cell) !== null;
}

/**
 * 只吃数字的线性插值分位数（Excel `PERCENTILE.INC` / `QUARTILE.INC` 的算法：rank = k·(n−1)，两邻值按小数权重插）。
 *
 * @param sorted 升序排列的数字数组，长度 ≥ 1。
 * @param fraction 0~1 的分位点。
 * @returns 分位值；空数组返回 null（调用方据此跳过统计，不许返回 0 冒充结果）。
 */
export function quantileInc(sorted: number[], fraction: number): number | null {
  const n = sorted.length;
  if (n === 0) return null;
  if (n === 1) return sorted[0]!;
  const rank = fraction * (n - 1);
  const low = Math.floor(rank);
  const high = Math.min(low + 1, n - 1);
  return sorted[low]! + (sorted[high]! - sorted[low]!) * (rank - low);
}

/**
 * 一组数字的描述统计（含四分位与 Tukey 栅栏）。
 *
 * @param values 数字集合，顺序不限（内部排序）。
 * @returns `NumericSummary`；空集合返回 null，含义是"这列没有可统计的数字"。
 */
export function numericSummary(values: number[]): NumericSummary | null {
  if (values.length === 0) return null;
  const sorted = [...values].sort((left, right) => left - right);
  const q1 = quantileInc(sorted, 0.25)!;
  const q3 = quantileInc(sorted, 0.75)!;
  const iqr = q3 - q1;
  const sum = sorted.reduce((acc, value) => acc + value, 0);
  return {
    count: sorted.length,
    sum,
    average: sum / sorted.length,
    min: sorted[0]!,
    max: sorted[sorted.length - 1]!,
    median: quantileInc(sorted, 0.5)!,
    q1,
    q3,
    iqr,
    lowerFence: q1 - 1.5 * iqr,
    upperFence: q3 + 1.5 * iqr,
  };
}

/**
 * 该列显示数字用的代表性格式：取第一个带数字格式的数值格。
 *
 * @param cells 列内的格子。
 * @returns `ResolvedNumberFormat` 或 undefined（undefined 时按 General 显示）。
 */
function representativeFormat(cells: Array<SheetCell | null>): ResolvedNumberFormat | undefined {
  for (const cell of cells) {
    if (cell && cellNumberValue(cell) !== null && cell.numberFormat) return cell.numberFormat;
  }
  return undefined;
}

/**
 * 数出非空格里的显示文本频次，按次数从多到少排好。
 *
 * @param cells 列内的格子。
 * @returns 计数表（全量，未截断）；空文本（只有格式的格）不计入。
 */
function valueCounts(cells: Array<SheetCell | null>): Array<{ text: string; count: number }> {
  const counts = new Map<string, number>();
  for (const cell of cells) {
    if (!cell || cell.kind === "blank" || cell.kind === "error") continue;
    const text = cellDisplayText(cell);
    if (text === "") continue;
    counts.set(text, (counts.get(text) ?? 0) + 1);
  }
  return [...counts].map(([text, count]) => ({ text, count })).sort((left, right) => right.count - left.count || left.text.localeCompare(right.text));
}

/**
 * 推断一列的主导类型并算出全部统计量。
 *
 * @param column 该列在数据行上的格子（`ColumnCells`）。
 * @returns `ColumnProfile`；`type` 为 `empty` 时其余统计量都是 0/空数组，不会给出假的数字。
 */
export function profileColumn(column: ColumnCells): ColumnProfile {
  const { cells } = column;
  const populated = cells.filter(
    (cell): cell is SheetCell => cell !== null && cell.kind !== "blank" && cell.kind !== "error" && cellDisplayText(cell) !== ""
  );
  const hasNumber = cells.some((cell) => cellNumberValue(cell) !== null && !cellIsDate(cell));
  const hasDate = cells.some(cellIsDate);
  const hasBoolean = cells.some((cell) => cell?.kind === "boolean");
  const hasString = cells.some((cell) => cell?.kind === "string" && cellDisplayText(cell) !== "");
  const kinds = [hasNumber, hasDate, hasBoolean, hasString].filter(Boolean).length;

  const counts = valueCounts(cells);
  // 日期格不进数字统计：序列号混进平均值会算出几万这种荒谬数字。
  const numericCells = cells.filter((cell): cell is SheetCell => cellNumberValue(cell) !== null && !cellIsDate(cell));
  const numbers = numericCells.map((cell) => cellNumberValue(cell)!);
  const base: ColumnProfile = {
    index: column.index,
    letter: column.letter,
    header: column.header,
    type: populated.length === 0 ? "empty" : kinds > 1 ? "mixed" : hasDate ? "date" : hasNumber ? "number" : hasBoolean ? "boolean" : "string",
    populated: populated.length,
    missing: cells.length - populated.length,
    distinct: counts.length,
    topValues: counts.slice(0, TOP_VALUE_LIMIT),
    pendingRecalc: cells.filter((cell) => cell?.formula && cellNumberValue(cell) === null).length,
    outliers: [],
    outlierCount: 0,
  };

  const numeric = numericSummary(numbers);
  if (base.type === "number" && numeric) {
    const format = representativeFormat(cells);
    base.numeric = numeric;
    base.numericText = Object.fromEntries(
      Object.entries(numeric).map(([key, value]) => [key, formatNumberForDisplay(value as number, format)])
    ) as Record<keyof NumericSummary, string>;
    const flagged = numericCells
      .filter((cell) => cellNumberValue(cell)! < numeric.lowerFence || cellNumberValue(cell)! > numeric.upperFence)
      .map((cell) => ({
        address: cell.address,
        display: cellDisplayText(cell),
        value: cellNumberValue(cell)!,
        side: cellNumberValue(cell)! < numeric.lowerFence ? ("low" as const) : ("high" as const),
      }));
    base.outlierCount = flagged.length;
    base.outliers = flagged.slice(0, OUTLIER_REPORT_LIMIT);
  }
  if (base.type === "date") {
    const serials = cells.filter(cellIsDate).map((cell) => cellNumberValue(cell)!);
    if (serials.length > 0) {
      const min = Math.min(...serials);
      const max = Math.max(...serials);
      const format = representativeFormat(cells);
      base.dateRange = { from: formatNumberForDisplay(min, format), to: formatNumberForDisplay(max, format) };
    }
  }
  return base;
}
