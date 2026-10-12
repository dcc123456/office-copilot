/**
 * A1 记法与 0 基下标的互转。
 *
 * 全仓只有这里允许把 `AB12` 这种 Excel 记法换算成内部下标（spec §零.2）：
 * 渲染层、AI 工具层拿到的一律是 `{ colIndex, rowIndex }` 或原样字符串，不再各算一遍。
 */

/** A1 单元格地址：列与行都是 0 基。 */
export interface A1Address {
  colIndex: number;
  rowIndex: number;
}

/** A1 矩形区域：`start` 恒在 `end` 的左上（读入时已归一化）。 */
export interface A1Range {
  start: A1Address;
  end: A1Address;
}

const COL_LETTERS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

/**
 * 列字母 → 0 基列号。
 *
 * @param name 纯字母列名，如 `A`、`AB`；大小写均可，不允许带数字。
 * @returns 0 基列号；含非字母或超出 `ZZZ`（18278 列）时返回 undefined，
 *   调用方按"读不懂"处理，不许猜一个列。
 */
export function colNameToIndex(name: string): number | undefined {
  const upper = name.toUpperCase();
  if (!/^[A-Z]{1,3}$/.test(upper)) return undefined;
  let acc = 0;
  for (const char of upper) {
    const digit = COL_LETTERS.indexOf(char);
    if (digit < 0) return undefined;
    acc = acc * 26 + digit + 1;
  }
  return acc - 1;
}

/**
 * 0 基列号 → 列字母。
 *
 * @param index 0 基列号。
 * @returns `A`/`Z`/`AA` 形式的列名；越界（<0 或 >18277）返回空串。
 */
export function indexToColName(index: number): string {
  if (!Number.isInteger(index) || index < 0 || index > 18277) return "";
  let rest = index + 1;
  let name = "";
  while (rest > 0) {
    const remainder = (rest - 1) % 26;
    name = COL_LETTERS[remainder] + name;
    rest = Math.floor((rest - 1) / 26);
  }
  return name;
}

/**
 * 解析单个 A1 地址。
 *
 * @param ref 形如 `B7`，允许 `$B$7` 的绝对引用写法（`$` 被忽略，我们不做引用平移）。
 * @returns 0 基地址；写法不对返回 undefined。
 */
export function parseA1(ref: string): A1Address | undefined {
  const match = /^\$?([A-Za-z]{1,3})\$?([1-9][0-9]{0,6})$/.exec(ref.trim());
  if (!match) return undefined;
  const colIndex = colNameToIndex(match[1] ?? "");
  if (colIndex === undefined) return undefined;
  const rowIndex = Number(match[2]) - 1;
  if (!Number.isInteger(rowIndex) || rowIndex < 0) return undefined;
  return { colIndex, rowIndex };
}

/**
 * 0 基地址 → A1 字符串（用于展示与 op 描述，不用于撤销定位，见 spec D3）。
 *
 * @param address 0 基地址。
 * @returns 大写的 A1 记法；列号越界返回空串。
 */
export function formatA1(address: A1Address): string {
  const col = indexToColName(address.colIndex);
  if (!col) return "";
  return `${col}${address.rowIndex + 1}`;
}

/**
 * 解析区域写法。
 *
 * @param ref `A1:C3`、`A1`（单格）、整列 `A:C`、整行 `2:4` 四种之一，允许带 `$`。
 * @returns 归一化后的矩形；整列/整行按 `maxRowIndex` 参数收口（调用方传实际已用行数，
 *   避免"整列"变成 1048576 行的假区域）。写法不可解析返回 undefined。
 */
export function parseRange(ref: string, maxRowIndex = 1048575): A1Range | undefined {
  const trimmed = ref.trim().replace(/\$/g, "");
  const wholeRow = /^([1-9][0-9]{0,6}):([1-9][0-9]{0,6})$/.exec(trimmed);
  if (wholeRow) {
    const from = Number(wholeRow[1]) - 1;
    const to = Number(wholeRow[2]) - 1;
    return { start: { colIndex: 0, rowIndex: Math.min(from, to) }, end: { colIndex: 16383, rowIndex: Math.max(from, to) } };
  }
  const wholeCol = /^([A-Za-z]{1,3}):([A-Za-z]{1,3})$/.exec(trimmed);
  if (wholeCol) {
    const from = colNameToIndex(wholeCol[1] ?? "");
    const to = colNameToIndex(wholeCol[2] ?? "");
    if (from === undefined || to === undefined) return undefined;
    return { start: { colIndex: Math.min(from, to), rowIndex: 0 }, end: { colIndex: Math.max(from, to), rowIndex: maxRowIndex } };
  }
  const pair = /^([A-Za-z]{1,3}[1-9][0-9]{0,6})(?::([A-Za-z]{1,3}[1-9][0-9]{0,6}))?$/.exec(trimmed);
  if (!pair) return undefined;
  const anchor = parseA1(pair[1] ?? "");
  if (!anchor) return undefined;
  if (!pair[2]) return { start: anchor, end: anchor };
  const far = parseA1(pair[2]);
  if (!far) return undefined;
  return {
    start: { colIndex: Math.min(anchor.colIndex, far.colIndex), rowIndex: Math.min(anchor.rowIndex, far.rowIndex) },
    end: { colIndex: Math.max(anchor.colIndex, far.colIndex), rowIndex: Math.max(anchor.rowIndex, far.rowIndex) },
  };
}

/**
 * 矩形 → A1 区域字符串。
 *
 * @param range 归一化矩形。
 * @returns 单格时是 `B7`，否则 `A1:C3`；不可格式化返回空串。
 */
export function formatRange(range: A1Range): string {
  const from = formatA1(range.start);
  const to = formatA1(range.end);
  if (!from || !to) return "";
  return from === to ? from : `${from}:${to}`;
}

/**
 * 矩形是否包含某个地址。
 *
 * @param range 判定区域。
 * @param address 待判地址。
 * @returns 落在区域内为 true（边界含在内）。
 */
export function rangeContains(range: A1Range, address: A1Address): boolean {
  return (
    address.colIndex >= range.start.colIndex &&
    address.colIndex <= range.end.colIndex &&
    address.rowIndex >= range.start.rowIndex &&
    address.rowIndex <= range.end.rowIndex
  );
}
