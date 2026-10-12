/**
 * 网格里"位置 ↔ 数据"的纯函数层（spec 阶段 2 · 提交 7）。
 *
 * 之所以独立一个文件：桌面端的 vitest 是 `environment: "node"`，没有 jsdom，
 * 组件本身测不了；把几何、换算、编辑落地这些真正会出错的部分收成纯函数，
 * 就能在 node 里钉死，组件只剩装配（`AGENTS.md` §3.3）。
 */
import {
  cellDisplayText,
  columnWidthCharsToPx,
  columnWidthPxToChars,
  formatA1,
  parseA1,
  parseUserCellText,
  rowHeightPtToPx,
  type A1Address,
  type A1Range,
  type SheetCell,
  type SheetPayload,
} from "@office-copilot/sheet-io";

/** Excel 默认列宽是 8.43 字符（2158/256），换算只在 `sheet-io` 里发生，这里拿到的已是 px。 */
const DEFAULT_COLUMN_WIDTH_PX = columnWidthCharsToPx(2158 / 256);

/** 默认行高（磅），与 Excel 空白工作簿一致；px 由 IO 包换算。 */
const DEFAULT_ROW_HEIGHT_PT = 15;

/**
 * 已用区域之外再多给的空列/空行：用户要能在表尾接着打字，Excel 就是这个行为。
 *
 * 空列按默认宽度 64px 排，8 列只有 512px——实测在 1200px 宽的窗口里右侧会露出一条没有格线的空白，
 * 所以要给到 24 列（≈1536px）才铺得满常见窗口。
 */
const TRAILING_COLUMNS = 24;
const TRAILING_ROWS = 40;

/** 虚拟滚动上下各多渲染这么多行，避免快速滚动露白。 */
export const OVERSCAN_ROWS = 8;

/** 行号表头列的宽度（px），属于窗口装饰而不是文档格式，因此不进 IO 包。 */
export const ROW_HEADER_WIDTH_PX = 46;

/** 列字母表头的高度（px）。 */
export const COLUMN_HEADER_HEIGHT_PX = 24;

/** 一屏要画多少格：只算可见行，列全画（列数受 `colCount` 约束，通常几十到几百）。 */
export interface SheetLayout {
  /** 可见列的真实列号，隐藏列不在里面。 */
  colIndexes: number[];
  /** 可见行的真实行号。 */
  rowIndexes: number[];
  /** 与 `colIndexes` 同序的列宽（px）。 */
  colWidthPx: number[];
  /** 与 `rowIndexes` 同序的行高（px）。 */
  rowHeightPx: number[];
  /** 前缀和，长度比 `colIndexes` 多 1，用来二分可见窗口。 */
  colOffsetPx: number[];
  rowOffsetPx: number[];
  totalWidthPx: number;
  totalHeightPx: number;
}

/**
 * 把一张表排成网格几何。
 *
 * @param sheet 跨界载荷（`cells` 以 A1 地址为键）。
 * @returns 可见行列的宽度、高度与前缀和；隐藏的行列（`<col/@hidden>`、`<row/@hidden>`）被排除，
 *   但数据仍在 `cells` 里，保存时原样带回，不许因为不显示就丢掉。
 */
export function buildSheetLayout(sheet: SheetPayload): SheetLayout {
  const colLimit = Math.max(sheet.lastColIndex + 1, 1) + TRAILING_COLUMNS;
  const rowLimit = Math.max(sheet.lastRowIndex + 1, 1) + TRAILING_ROWS;
  const hiddenCols = hiddenIndexSet(sheet.columns.map((column) => [column.minColIndex, column.maxColIndex, column.hidden] as const), colLimit);
  const hiddenRows = hiddenIndexSet(sheet.rows.map((row) => [row.rowIndex, row.rowIndex, row.hidden] as const), rowLimit);
  const widthByCol = widthByColumnIndex(sheet.columns, colLimit);
  const heightByRow = heightByRowIndex(sheet.rows);

  const colIndexes: number[] = [];
  const colWidthPx: number[] = [];
  for (let colIndex = 0; colIndex < colLimit; colIndex += 1) {
    if (hiddenCols.has(colIndex)) continue;
    colIndexes.push(colIndex);
    colWidthPx.push(widthByCol.get(colIndex) ?? DEFAULT_COLUMN_WIDTH_PX);
  }
  const rowIndexes: number[] = [];
  const rowHeightPx: number[] = [];
  const defaultRowPx = rowHeightPtToPx(DEFAULT_ROW_HEIGHT_PT);
  for (let rowIndex = 0; rowIndex < rowLimit; rowIndex += 1) {
    if (hiddenRows.has(rowIndex)) continue;
    rowIndexes.push(rowIndex);
    rowHeightPx.push(heightByRow.get(rowIndex) ?? defaultRowPx);
  }

  const colOffsetPx = prefixOffsets(colWidthPx);
  const rowOffsetPx = prefixOffsets(rowHeightPx);
  return {
    colIndexes,
    rowIndexes,
    colWidthPx,
    rowHeightPx,
    colOffsetPx,
    rowOffsetPx,
    totalWidthPx: colOffsetPx[colOffsetPx.length - 1] ?? 0,
    totalHeightPx: rowOffsetPx[rowOffsetPx.length - 1] ?? 0,
  };
}

/**
 * 求滚动窗口内要渲染的槽位区间（行、列两个方向共用这一条）。
 *
 * @param offsets 该方向的前缀和（`layout.rowOffsetPx` 或 `colOffsetPx`，长度比槽位数多 1）。
 * @param scrollPx 该方向的滚动距离（px）。
 * @param viewportPx 视口在该方向上的尺寸（px）。
 * @param overscan 两端各多算几个槽位。
 * @returns `[start, end)` 的槽位下标区间；已夹到有效范围。
 */
export function visibleSlotRange(offsets: number[], scrollPx: number, viewportPx: number, overscan = OVERSCAN_ROWS): [number, number] {
  const count = offsets.length - 1;
  const start = Math.max(0, lowerBound(offsets, scrollPx) - 1 - overscan);
  const end = Math.min(count, lowerBound(offsets, scrollPx + viewportPx) + overscan);
  return [start, Math.max(start, end)];
}

/**
 * 一格显示出来的文字。
 *
 * @param cell 格子；undefined 表示空格。
 * @returns 显示文本；空内容是空串。日期与百分比的读法全在 `cellDisplayText` 里，这里不重复。
 */
export function displayTextOf(cell: SheetCell | undefined): string {
  return cell ? cellDisplayText(cell) : "";
}

/**
 * 用户在网格里输入的一行文字如何落到表上。
 *
 * 保留原格的样式（`styleIndex`/`attrs`）是硬约束：`AGENTS.md` §2.3 不许把局部格式拍平，
 * 只清掉与旧值绑定的东西（公式、缓存值、原文、共享字符串索引）。
 * 清空内容时，原本带格式的格退化成"只有格式的格"，与 Excel 行为一致。
 *
 * @param sheet 当前表的载荷。
 * @param address A1 地址（与 `colIndex/rowIndex` 同源，由本函数算出，调用方不必自己数）。
 * @param text 输入框里的原始文字。
 * @returns 新载荷（`cells` 复制一份），并把已用区域放大到这一格。
 */
export function withCellText(sheet: SheetPayload, address: string, text: string): SheetPayload {
  const parsedAddress = parseA1(address);
  if (!parsedAddress) return sheet;
  const parsed = parseUserCellText(text);
  const previous = sheet.cells[address];
  const cells = { ...sheet.cells };

  if (parsed.kind === "blank" && (previous?.styleIndex === undefined || previous?.rawXml)) {
    delete cells[address];
  } else {
    const next: SheetCell = {
      address,
      colIndex: parsedAddress.colIndex,
      rowIndex: parsedAddress.rowIndex,
      kind: parsed.kind,
      value: parsed.value,
      ...(previous ? { styleIndex: previous.styleIndex, attrs: previous.attrs } : {}),
      ...(parsed.kind === "blank" ? { styleOnly: true } : {}),
      ...(parsed.numberFormat ? { numberFormat: parsed.numberFormat } : {}),
    };
    cells[address] = next;
  }

  return {
    ...sheet,
    cells,
    lastColIndex: Math.max(sheet.lastColIndex, parsedAddress.colIndex),
    lastRowIndex: Math.max(sheet.lastRowIndex, parsedAddress.rowIndex),
  };
}

/**
 * 由按下点与拖拽点组成归一化选区。
 *
 * @param anchor 起点（鼠标按下或键盘锚点）。
 * @param head 终点（当前格）。
 * @returns 归一化后的矩形；两点相同时就是单格选区。
 */
export function selectionRectOf(anchor: A1Address, head: A1Address): A1Range {
  return {
    start: { colIndex: Math.min(anchor.colIndex, head.colIndex), rowIndex: Math.min(anchor.rowIndex, head.rowIndex) },
    end: { colIndex: Math.max(anchor.colIndex, head.colIndex), rowIndex: Math.max(anchor.rowIndex, head.rowIndex) },
  };
}

/**
 * 选区右下角的格子地址，状态栏显示 `A1:C3` 用。
 *
 * @param range 选区。
 * @returns 单格时是 `B2`，多格时是 `B2:D5`。
 */
export function selectionLabel(range: A1Range): string {
  const start = formatA1(range.start);
  const end = formatA1(range.end);
  return start === end ? start : `${start}:${end}`;
}

/**
 * 把键盘位移落在一个格子上。
 *
 * @param address 当前格。
 * @param dCol 列位移（-1/0/1）。
 * @param dRow 行位移（-1/0/1）。
 * @param maxColIndex 网格最大列号（含尾部空列）。
 * @param maxRowIndex 网格最大行号。
 * @returns 新的格子地址；越界方向夹到边上，不循环。
 */
export function moveAddress(address: A1Address, dCol: number, dRow: number, maxColIndex: number, maxRowIndex: number): A1Address {
  return {
    colIndex: Math.min(Math.max(address.colIndex + dCol, 0), maxColIndex),
    rowIndex: Math.min(Math.max(address.rowIndex + dRow, 0), maxRowIndex),
  };
}

/**
 * 把点在网格里的坐标换成格子地址。
 *
 * @param layout 网格几何。
 * @param x 相对画布左上角的横坐标（px，不含行号表头）。
 * @param y 相对画布左上角的纵坐标（px，不含列字母表头）。
 * @returns 真实行列号；点在表头留白之外时夹到最后一格，不返回 undefined。
 */
export function addressAtPoint(layout: SheetLayout, x: number, y: number): A1Address {
  const colSlot = slotForOffset(layout.colOffsetPx, x, layout.colIndexes.length);
  const rowSlot = slotForOffset(layout.rowOffsetPx, y, layout.rowIndexes.length);
  return {
    colIndex: layout.colIndexes[colSlot] ?? 0,
    rowIndex: layout.rowIndexes[rowSlot] ?? 0,
  };
}

/**
 * 拖动某一列的边界后改写这一列的宽度。
 *
 * `<col min max>` 是一段范围，只改其中一列必须把范围劈成三段（左段原样、中间这列新宽度、右段原样），
 * 否则会把相邻列的宽度一起牵连掉。
 *
 * @param sheet 当前表的载荷。
 * @param colIndex 被拖的真实列号。
 * @param widthPx 拖出来的像素宽度。
 * @returns 新载荷；`columns` 换了一份数组，其余字段原样。
 */
export function withColumnWidth(sheet: SheetPayload, colIndex: number, widthPx: number): SheetPayload {
  const widthChars = columnWidthPxToChars(widthPx);
  const covering = sheet.columns.find((column) => colIndex >= column.minColIndex && colIndex <= column.maxColIndex);
  const others = sheet.columns.filter((column) => column !== covering);
  const rebuilt: SheetPayload["columns"] = [];
  if (covering) {
    if (covering.minColIndex < colIndex) {
      rebuilt.push({ ...covering, minColIndex: covering.minColIndex, maxColIndex: colIndex - 1 });
    }
    rebuilt.push({ ...covering, minColIndex: colIndex, maxColIndex: colIndex, widthChars, customWidth: true });
    if (covering.maxColIndex > colIndex) {
      rebuilt.push({ ...covering, minColIndex: colIndex + 1, maxColIndex: covering.maxColIndex });
    }
  } else {
    rebuilt.push({ minColIndex: colIndex, maxColIndex: colIndex, widthChars, customWidth: true });
  }
  return { ...sheet, columns: [...others, ...rebuilt] };
}

/** 前缀和里找 `x` 落在哪一段；段号即槽位下标，越界夹到最后一段。 */
function slotForOffset(offsets: number[], value: number, count: number): number {
  let low = 0;
  let high = count - 1;
  while (low < high) {
    const middle = (low + high + 1) >> 1;
    if ((offsets[middle] ?? 0) <= value) low = middle;
    else high = middle - 1;
  }
  return Math.max(0, Math.min(low, count - 1));
}

/** 列范围里可见列的宽度；`hidden` 的列由调用方排除，这里只算 px。 */
function widthByColumnIndex(columns: SheetPayload["columns"], colLimit: number): Map<number, number> {
  const acc = new Map<number, number>();
  for (const column of columns) {
    if (column.widthChars === undefined) continue;
    const max = Math.min(column.maxColIndex, colLimit - 1);
    for (let colIndex = Math.max(column.minColIndex, 0); colIndex <= max; colIndex += 1) {
      acc.set(colIndex, columnWidthCharsToPx(column.widthChars));
    }
  }
  return acc;
}

/** 行元信息里的自定义行高（磅）换成 px。 */
function heightByRowIndex(rows: SheetPayload["rows"]): Map<number, number> {
  const acc = new Map<number, number>();
  for (const row of rows) {
    if (row.heightPt !== undefined) acc.set(row.rowIndex, rowHeightPtToPx(row.heightPt));
  }
  return acc;
}

/** 把 `[min, max, hidden]` 三段范围摊平成隐藏下标集合，只摊到 `limit` 为止。 */
function hiddenIndexSet(ranges: ReadonlyArray<readonly [number, number, boolean | undefined]>, limit: number): Set<number> {
  const acc = new Set<number>();
  for (const [min, max, hidden] of ranges) {
    if (!hidden) continue;
    for (let index = Math.max(min, 0); index <= Math.min(max, limit - 1); index += 1) acc.add(index);
  }
  return acc;
}

function prefixOffsets(sizes: number[]): number[] {
  const offsets = [0];
  let total = 0;
  for (const size of sizes) {
    total += size;
    offsets.push(total);
  }
  return offsets;
}

/** 第一个 `offsets[i] > value` 的下标；用于按滚动距离定位起始行。 */
function lowerBound(offsets: number[], value: number): number {
  let low = 0;
  let high = offsets.length - 1;
  while (low < high) {
    const middle = (low + high) >> 1;
    if ((offsets[middle] ?? 0) > value) high = middle;
    else low = middle + 1;
  }
  return low;
}
