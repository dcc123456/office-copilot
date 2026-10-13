/**
 * 图表规格（spec 决策 D5 · 阶段 5）：把"选哪块数据画什么图"变成结构化数据，并从**现场表**里把数取出来。
 *
 * 为什么落在 `sheet-io`：数字必须由引擎算/读，模型只给区域引用（`AGENTS.md` §2.8、验收 B10），
 * 而引用解析、A1 格式化、显示文本这三件事这里都已经有实现（`address.ts`、`cell-text.ts`），不再抄第二套。
 * 取数结果里带的是缓存值与显示串，`xlsx-chart.ts` 只负责把它拼成 OOXML，渲染层只负责画。
 */
import { formatA1, indexToColName, parseA1, parseRange, type A1Address, type A1Range } from "./address.js";
import { cellDisplayText } from "./cell-text.js";
import type { SheetCell, SheetLike as SheetPayload } from "./model.js";

/** 一期只支持三种图型：足够覆盖"看趋势 / 比大小 / 看占比"，更多图型等真实需求再开。 */
export type ChartKind = "bar" | "line" | "pie";

/** 支持的图型清单，工具面与校验都从这一份里取。 */
export const CHART_KINDS: readonly ChartKind[] = ["bar", "line", "pie"];

/** 一个系列：`values` 是数据区域，`name` 是给系列名用的单格 A1 引用（可省，省了就用列字母当名）。 */
export interface ChartSeriesInput {
  values: string;
  name?: string;
}

/** 模型（或 UI）交出来的图表请求，全部是 A1 引用，不含任何数值。 */
export interface ChartSpecInput {
  kind: string;
  title?: string;
  /** 类别轴区域，如 `A2:A4`。 */
  categories: string;
  series: ChartSeriesInput[];
  /** 图表左上角落在哪个格（A1，如 `E2`）。 */
  anchor: string;
  /** 尺寸，单位 CSS px（96 DPI）；`sheet-io` 进出口之外不许再换算（§2.1）。 */
  widthPx?: number;
  heightPx?: number;
}

/** 解析后的一个系列：名字 + 给 OOXML 的引用串 + 引擎读到的数值。 */
export interface ChartSeries {
  name: string;
  /** `'销售'!$C$1` 形态；没有名字格时是 null，写入时省略 `<c:tx>`。 */
  nameRef: string | null;
  /** `'销售'!$C$2:$C$4` 形态。 */
  valueRef: string;
  /** 与类别轴同长的数值；`NaN` 不会出现——取不到数就整条请求被拒。 */
  values: number[];
}

/** 一张可以画、也可以写进 `.xlsx` 的图表的全部现场数据。 */
export interface ResolvedChart {
  kind: ChartKind;
  title: string;
  anchor: A1Address;
  /** 类别轴的显示串（文本类别时用它填 `strCache`）。 */
  categories: string[];
  categoryRef: string;
  /** 类别轴是否真是数字：决定写 `numRef` 还是 `strRef`（Excel 两种都收，但文本走 numRef 会画成 0）。 */
  categoriesAreNumeric: boolean;
  categoryNumbers: number[];
  series: ChartSeries[];
  widthPx: number;
  heightPx: number;
}

/** 校验结果。失败时 `reason` 是给模型/用户看的一句话，不抛异常（验收 B9 的口径）。 */
export type ChartResolveResult = { ok: true; chart: ResolvedChart } | { ok: false; reason: string };

/** 默认图表尺寸，与 Excel 插入图表的默认框一致量级（px，@96DPI）。 */
const DEFAULT_WIDTH_PX = 480;
const DEFAULT_HEIGHT_PX = 288;

/** 一次请求里最多几个系列：再多就不是"选区出图"，是让模型画矩阵了。 */
const MAX_SERIES = 8;

/** px → EMU：96 DPI 下 1 px = 9525 EMU，这是本包里唯一一处尺寸换算（§2.1）。 */
const EMU_PER_PX = 9525;

/**
 * 把表名拼成 OOXML 引用里需要的带引号前缀。
 *
 * @param name 工作表名，可能含空格或单引号。
 * @returns `'销售'!` 这样的前缀；名字里的单引号按规范 doubling 成两个。
 */
export function quotedSheetPrefix(name: string): string {
  return `'${name.replace(/'/g, "''")}'!`;
}

/**
 * 把一段 A1 区域拼成引用串（绝对地址形态，Excel 自己插图表就是这么写的）。
 *
 * @param sheetName 工作表名。
 * @param range 已解析好的区域。
 * @returns 单格时 `'销售'!$C$1`，多格时 `'销售'!$C$1:$C$4`。
 */
function rangeRef(sheetName: string, range: A1Range): string {
  const start = `$${indexToColName(range.start.colIndex)}$${range.start.rowIndex + 1}`;
  const end = `$${indexToColName(range.end.colIndex)}$${range.end.rowIndex + 1}`;
  const sameCell = range.start.colIndex === range.end.colIndex && range.start.rowIndex === range.end.rowIndex;
  return quotedSheetPrefix(sheetName) + (sameCell ? start : `${start}:${end}`);
}

/**
 * 取一格的数值口径：公式格取缓存值，常量数字格取 `value`。
 *
 * @param cell 目标格，可能是 undefined（这格根本没内容）。
 * @returns 能当数字用的值取不到时返回 null，调用方据此给出理由。
 */
function numericOf(cell: SheetCell | undefined): number | null {
  if (!cell || cell.kind !== "number") return null;
  const raw = cell.formula ? cell.cachedValue : cell.value;
  return typeof raw === "number" && Number.isFinite(raw) ? raw : null;
}

/**
 * 按区域展开成一串格子（只允许单列或单行——图表的一条系列必须是一维的）。
 *
 * @param sheet 现场表载荷。
 * @param ref 用户/模型写的 A1 区域。
 * @returns 顺序固定的格子数组；区域写法不是一维时返回理由。
 */
function cellsOfLine(sheet: SheetPayload, ref: string): { cells: SheetCell[] } | { reason: string } {
  const range = parseRange(ref.trim());
  if (!range) return { reason: `认不出区域「${ref}」，请写成 A2:A4 这样的形式。` };
  const isColumn = range.start.colIndex === range.end.colIndex;
  const isRow = range.start.rowIndex === range.end.rowIndex;
  if (!isColumn && !isRow) return { reason: `区域「${ref}」是二维的，一条系列只能是一列或一行。` };
  const cells: SheetCell[] = [];
  if (isColumn) {
    const letter = indexToColName(range.start.colIndex);
    for (let row = range.start.rowIndex; row <= range.end.rowIndex; row += 1) {
      const cell = sheet.cells[`${letter}${row + 1}`];
      if (cell) cells.push(cell);
    }
  } else {
    for (let col = range.start.colIndex; col <= range.end.colIndex; col += 1) {
      const letter = indexToColName(col);
      const cell = sheet.cells[`${letter}${range.start.rowIndex + 1}`];
      if (cell) cells.push(cell);
    }
  }
  if (!cells.length) return { reason: `区域「${ref}」里没有任何内容。` };
  return { cells };
}

/**
 * 校验一份图表请求，并把引用现场解析成可画/可写的数据。
 *
 * 失败一律给一句话理由，不抛异常：模型看得懂才改得动（`AGENTS.md` §2.8、spec 验收 B9/B10）。
 *
 * @param sheet 当前工作表载荷。
 * @param input 模型或 UI 交出来的请求。
 * @returns `ChartResolveResult`；`ok:false` 时调用方不许落地。
 */
export function resolveChartSpec(sheet: SheetPayload, input: ChartSpecInput): ChartResolveResult {
  if (!CHART_KINDS.includes(input.kind as ChartKind)) {
    return { ok: false, reason: `图型「${input.kind}」不支持，能用的是 ${CHART_KINDS.join(" / ")}。` };
  }
  const anchor = parseA1(input.anchor.trim().toUpperCase());
  if (!anchor) return { ok: false, reason: `落点「${input.anchor}」不是一个格子地址，请写成 E2 这样。` };
  if (!input.series?.length) return { ok: false, reason: "至少要一条系列（要画哪一列数据）。" };
  if (input.series.length > MAX_SERIES) return { ok: false, reason: `一条图表最多 ${MAX_SERIES} 条系列，现在 ${input.series.length} 条。` };

  const cats = cellsOfLine(sheet, input.categories);
  if ("reason" in cats) return { ok: false, reason: cats.reason };
  const categories = cats.cells.map(cellDisplayText);
  const categoryNumbers = cats.cells.map((cell) => numericOf(cell));
  const categoriesAreNumeric = categoryNumbers.every((value) => value !== null);
  if (categories.every((text) => !text)) return { ok: false, reason: `类别轴「${input.categories}」全是空的。` };

  const prefix = quotedSheetPrefix(sheet.name);
  const series: ChartSeries[] = [];
  for (const item of input.series) {
    const taken = cellsOfLine(sheet, item.values);
    if ("reason" in taken) return { ok: false, reason: taken.reason };
    if (taken.cells.length !== categories.length) {
      return {
        ok: false,
        reason: `区域「${item.values}」有 ${taken.cells.length} 个值，类别轴「${input.categories}」有 ${categories.length} 个，对不上。`,
      };
    }
    const values: number[] = [];
    for (const cell of taken.cells) {
      const numeric = numericOf(cell);
      if (numeric === null) {
        return { ok: false, reason: `「${cell.address}」不是数值，图表只画数字；空行请先去掉。` };
      }
      values.push(numeric);
    }
    let nameRef: string | null = null;
    let name = indexToColName(parseRange(item.values)!.start.colIndex);
    if (item.name?.trim()) {
      const nameCell = parseA1(item.name.trim().toUpperCase());
      if (!nameCell) return { ok: false, reason: `系列名「${item.name}」不是一个格子地址。` };
      const cell = sheet.cells[formatA1(nameCell)];
      if (!cell || !cellDisplayText(cell)) return { ok: false, reason: `系列名「${item.name}」那格是空的。` };
      name = cellDisplayText(cell);
      nameRef = rangeRef(sheet.name, { start: nameCell, end: nameCell });
    }
    series.push({ name, nameRef, valueRef: rangeRef(sheet.name, parseRange(item.values)!), values });
  }

  const categoriesRef = parseRange(input.categories)!;
  return {
    ok: true,
    chart: {
      kind: input.kind as ChartKind,
      title: input.title?.trim() || series[0]!.name,
      anchor,
      categories,
      categoryRef: rangeRef(sheet.name, categoriesRef),
      categoriesAreNumeric,
      categoryNumbers: categoryNumbers.map((value) => value ?? 0),
      series,
      widthPx: input.widthPx ?? DEFAULT_WIDTH_PX,
      heightPx: input.heightPx ?? DEFAULT_HEIGHT_PX,
    },
  };
}

/**
 * 图表框的 EMU 尺寸（`<xdr:ext>` 用的单位）。
 *
 * @param chart 已解析的图表。
 * @returns `{ cx, cy }`，整数 EMU。
 */
export function chartExtentEmu(chart: ResolvedChart): { cx: number; cy: number } {
  return { cx: Math.round(chart.widthPx * EMU_PER_PX), cy: Math.round(chart.heightPx * EMU_PER_PX) };
}
