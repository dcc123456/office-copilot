/**
 * 表格的结构化改动（spec 阶段 3）：每个 op 自带逆操作，撤销按内容指纹重定位。
 *
 * 为什么落在 `sheet-io` 而不是桌面的 `src/ai`：阶段 4 的 `packages/sheet-tools` 要产出同一批 op，
 * 而依赖方向只允许 `apps → packages`（`AGENTS.md` §5.1）；op 长在渲染层就会被人复制第二套。
 * 形状与 Word 的 `CoreOp` 同构：`before` 与 `after` 都存在 op 里，逆操作就是两者交换，
 * 落地前先核对现场等于 `before`，对不上就整条跳过并报"部分撤销"（§2.4/§2.5）。
 */
import { formatA1, parseA1 } from "./address.js";
import { parseUserCellText } from "./cell-text.js";
import { columnWidthPxToChars } from "./geometry.js";
import type { SheetCell, SheetColumn, SheetLike as SheetPayload } from "./model.js";
import { colFingerprints, recalculateSheetExtent, rowFingerprints } from "./xlsx-sheet.js";

/** 行/列两个轴：插删与尺寸调整的语义完全对称，所以一套类型管两轴。 */
export type SheetAxis = "row" | "col";

/**
 * 一条行或列的内容寻址。
 *
 * 位置会随插删漂移，所以只有内容指纹能认出"还是那一条"；`index` 是空白条的退路。
 */
export interface SheetTrackAddress {
  axis: SheetAxis;
  /** 建 op 时该条的内容指纹（`rowFingerprints` / `colFingerprints` 的产物）；空白条是伪指纹 `EMPTY_TRACK`。 */
  fingerprint: string;
  /** 同指纹出现的第几个（0 基）；两条内容完全相同的记录靠它区分。 */
  nth: number;
  /** 建 op 时的 0 基下标，只在空白条上用作"必须还在这里且仍然为空"的校验。 */
  index: number;
}

/** 一格的目标：行用内容寻址，列号按位置——列要是被插删挪过，现场核对会拦下来而不是写错格。 */
export interface SheetCellTarget {
  row: SheetTrackAddress;
  colIndex: number;
}

/** 改一格的值（清空也算，就是这格从有到无）。`null` 表示这格不存在。 */
export interface SetCellValueOp {
  kind: "setCellValue";
  /** 目标表的包内部件路径；对不上就整条跳过，不许静默写进别的表。 */
  partPath: string;
  target: SheetCellTarget;
  before: SheetCell | null;
  after: SheetCell | null;
}

/** 改一列的宽度。存的是"覆盖这一列的那几段 `<col>`"的整段前后像，逆操作把三段换回去即可。 */
export interface SetColumnWidthOp {
  kind: "setColumnWidth";
  partPath: string;
  /** `axis` 恒为 `col` 的地址。 */
  target: SheetTrackAddress;
  beforeColumns: SheetColumn[];
  afterColumns: SheetColumn[];
}

/** 表格的 op 全集（阶段 3 后续会加插删行列与合并）。 */
export type SheetCoreOp = SetCellValueOp | SetColumnWidthOp;

/** 一条 op 没落地时的记录，UI 据此报"部分撤销（x/y）"。 */
export interface SheetSkippedOp {
  op: SheetCoreOp;
  reason: string;
}

/** 一批 op 的执行结果。`ok` 为 false 不代表白做：`after` 是已落地那些 op 的效果。 */
export interface SheetOpsResult {
  ok: boolean;
  before: SheetPayload;
  after: SheetPayload;
  /** 真正落地的 op（跳过的不在里面）。 */
  ops: SheetCoreOp[];
  /** 与 `ops` 同序的逆操作，交给 `revertSheetOps`。 */
  inverses: SheetCoreOp[];
  skipped: SheetSkippedOp[];
}

/** 单条 op 的落地结果：成功给新载荷与逆操作，失败只给理由。 */
type ApplyOutcome = { sheet: SheetPayload; inverse: SheetCoreOp } | { reason: string };

/**
 * 一批 op 的唯一落地入口：在载荷的副本上顺序应用，失败的挑出去，成功的合成新载荷。
 *
 * @param sheet 当前表的载荷（入参不被改动）。
 * @param ops 待应用的一批 op，必须都指向同一张表。
 * @returns `SheetOpsResult`；`after` 永远是新对象。
 */
export function runSheetOps(sheet: SheetPayload, ops: SheetCoreOp[]): SheetOpsResult {
  return applyOps(sheet, ops);
}

/**
 * 撤销：把逆操作队列倒序应用回去。
 *
 * @param sheet 现场载荷（可能已被用户后来的手工编辑改过）。
 * @param inverses `runSheetOps` 记下的逆操作队列。
 * @returns `SheetOpsResult`；认不出目标（内容变了、行列被挪过）的条目进 `skipped`，
 *   调用方必须据此报"部分撤销"，不许猜一个位置继续写（§2.5）。
 */
export function revertSheetOps(sheet: SheetPayload, inverses: SheetCoreOp[]): SheetOpsResult {
  return applyOps(sheet, inverses.slice().reverse());
}

/**
 * 由"用户在某格输入了这段文字"造一条改值 op。
 *
 * 保留原格的样式（`styleIndex`/`attrs`）是硬约束：`AGENTS.md` §2.3 不许把局部格式拍平，
 * 只清掉与旧值绑定的东西（公式、缓存值、原文、共享字符串索引）。
 * 清空内容时，原本带格式的格退化成"只有格式的格"，与 Excel 行为一致。
 *
 * @param sheet 当前表的载荷。
 * @param address 大写 A1 地址。
 * @param text 输入框里的原始文字。
 * @returns 可直接交给 `runSheetOps` 的 op；地址解析不出来时返回 null。
 */
export function setCellTextOp(sheet: SheetPayload, address: string, text: string): SetCellValueOp | null {
  const parsedAddress = parseA1(address);
  if (!parsedAddress) return null;
  const parsed = parseUserCellText(text);
  const previous = sheet.cells[address] ?? null;
  const cleared = parsed.kind === "blank" && (previous?.styleIndex === undefined || previous?.rawXml !== undefined);
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
  return {
    kind: "setCellValue",
    partPath: sheet.partPath,
    target: { row: trackAddress(sheet, "row", parsedAddress.rowIndex), colIndex: parsedAddress.colIndex },
    before: previous,
    after: cleared ? null : next,
  };
}

/**
 * 由"用户把某列拖到这么宽"造一条改列宽 op。
 *
 * `<col min max>` 是一段范围，只改其中一列必须把范围劈成三段（左段原样、中间这列新宽度、右段原样），
 * 否则会把相邻列的宽度一起牵连掉。像素换字符只在 `geometry.ts` 发生（§2.1）。
 *
 * @param sheet 当前表的载荷。
 * @param colIndex 被拖的列号（0 基）。
 * @param widthPx 拖出来的宽度（CSS px，非 pt）。
 * @returns 改列宽 op；这一列本来就已经是单独一段且宽度相同（没有可改的）时返回 null。
 */
export function setColumnWidthOp(sheet: SheetPayload, colIndex: number, widthPx: number): SetColumnWidthOp | null {
  const widthChars = columnWidthPxToChars(widthPx);
  const covering = sheet.columns.filter((column) => colIndex >= column.minColIndex && colIndex <= column.maxColIndex);
  const single = covering.length === 1 ? covering[0] : undefined;
  if (single && single.minColIndex === colIndex && single.maxColIndex === colIndex && single.widthChars === widthChars) {
    return null;
  }
  const rebuilt: SheetColumn[] = [];
  for (const column of covering) {
    if (column.minColIndex < colIndex) rebuilt.push({ ...column, maxColIndex: colIndex - 1 });
    if (column.maxColIndex > colIndex) rebuilt.push({ ...column, minColIndex: colIndex + 1 });
  }
  const inherited = covering[0];
  rebuilt.push({ ...(inherited ?? {}), minColIndex: colIndex, maxColIndex: colIndex, widthChars, customWidth: true });
  rebuilt.sort((left, right) => left.minColIndex - right.minColIndex);
  return {
    kind: "setColumnWidth",
    partPath: sheet.partPath,
    target: trackAddress(sheet, "col", colIndex),
    beforeColumns: covering.map((column) => ({ ...column })),
    afterColumns: rebuilt,
  };
}

/**
 * 把内容寻址解成现场下标。
 *
 * @param sheet 目标表载荷。
 * @param address 建 op 时记下的地址。
 * @returns 0 基下标；内容对不上（这一条被用户改过）返回 null，调用方整条跳过。
 *   空白条没有可认的内容，只能按位置回，真正拦住写错地方的是现场核对（格子或宽度段的 `before`）。
 */
function resolveTrack(sheet: SheetPayload, address: SheetTrackAddress): number | null {
  const fingerprints = fingerprintsFor(sheet, address.axis);
  if (address.fingerprint === EMPTY_TRACK) return address.index;
  let seen = 0;
  for (const index of Object.keys(fingerprints).map(Number).sort((left, right) => left - right)) {
    if (fingerprints[index] !== address.fingerprint) continue;
    if (seen === address.nth) return index;
    seen += 1;
  }
  return null;
}

/** 某一轴的全部指纹：行用 `rowFingerprints`，列用 `colFingerprints`。 */
function fingerprintsFor(sheet: SheetPayload, axis: SheetAxis): Record<number, string> {
  return axis === "row" ? rowFingerprints(sheet) : colFingerprints(sheet);
}

/** 空白条的伪指纹：让它可寻址，但仍只在下标处仍然为空时才认。 */
const EMPTY_TRACK = "∅";

/**
 * 记下"这一条现在装的是什么"，之后能凭内容找回来。
 *
 * @param sheet 目标表载荷。
 * @param axis 行轴还是列轴。
 * @param index 0 基下标。
 * @returns 内容寻址；空白条的指纹是伪指纹 `EMPTY_TRACK`，只有位置校验（那条还空着）兜得住。
 */
function trackAddress(sheet: SheetPayload, axis: SheetAxis, index: number): SheetTrackAddress {
  const fingerprints = fingerprintsFor(sheet, axis);
  const fingerprint = fingerprints[index] ?? EMPTY_TRACK;
  const nth = Object.keys(fingerprints)
    .map(Number)
    .filter((other) => other < index && (fingerprints[other] ?? EMPTY_TRACK) === fingerprint).length;
  return { axis, fingerprint, nth, index };
}

/**
 * 两格内容是否等价——只比"这次改动会不会牵连的东西"。
 *
 * 地址不参与比较：一次改动只认"值、类型、公式、样式索引"这一组，地址前缀写进去会因为插删行
 * 而必然失配，等于承认"位移之后的格不是原来那格"，撤销就再也落不回去了。
 * 富文本 run 与原文（`rawXml`）只比"在不在"，因为把红字拍平成纯文本本身就是 §2.3 禁止的事，
 * 现场多出或少掉这类结构就不许覆盖。
 */
function sameCellContent(left: SheetCell | null, right: SheetCell | null): boolean {
  if (left === null || right === null) return left === right;
  return (
    left.kind === right.kind &&
    left.value === right.value &&
    (left.formula ?? "") === (right.formula ?? "") &&
    left.styleIndex === right.styleIndex &&
    left.styleOnly === right.styleOnly &&
    Boolean(left.runs) === Boolean(right.runs) &&
    Boolean(left.rawXml) === Boolean(right.rawXml)
  );
}

/** 两段 `<col>` 是否等价（含我们没有解释、但必须原样搬的其余属性）。 */
function sameColumn(left: SheetColumn, right: SheetColumn): boolean {
  return (
    left.minColIndex === right.minColIndex &&
    left.maxColIndex === right.maxColIndex &&
    left.widthChars === right.widthChars &&
    left.hidden === right.hidden &&
    left.customWidth === right.customWidth &&
    sameAttributes(left.attrs, right.attrs)
  );
}

/** 其余属性的逐条比较；键序不稳定，所以按键名排序后再比。 */
function sameAttributes(left: Record<string, string> | undefined, right: Record<string, string> | undefined): boolean {
  const leftKeys = Object.keys(left ?? {}).sort();
  const rightKeys = Object.keys(right ?? {}).sort();
  if (leftKeys.length !== rightKeys.length) return false;
  return leftKeys.every((key) => left?.[key] === right?.[key]);
}

/**
 * 落一条改值 op。
 *
 * @param sheet 现场载荷。
 * @param op 改值 op（逆操作形状相同，`before`/`after` 已交换）。
 * @returns 新载荷与逆操作，或跳过理由。
 */
function applyCellOp(sheet: SheetPayload, op: SetCellValueOp): ApplyOutcome {
  const rowIndex = resolveTrack(sheet, op.target.row);
  if (rowIndex === null) return { reason: "认不出这一行：内容已被改动，或它被插入/删除的行挪走了" };
  const address = formatA1({ colIndex: op.target.colIndex, rowIndex });
  if (!address) return { reason: "目标地址算不出来" };
  const current = sheet.cells[address] ?? null;
  if (!sameCellContent(current, op.before)) {
    return { reason: `${address} 现在装的不是这条改动记录时的内容，不覆盖它` };
  }
  const cells = { ...sheet.cells };
  if (op.after === null) delete cells[address];
  else cells[address] = { ...op.after, address, colIndex: op.target.colIndex, rowIndex };
  const next = { ...sheet, cells };
  recalculateSheetExtent(next);
  // 逆操作按"落地之后的样子"重新寻址：这次改动本身就换了行内容指纹，沿用旧指纹会认不出自己刚写的那行。
  const inverseTarget: SheetCellTarget = { row: trackAddress(next, "row", rowIndex), colIndex: op.target.colIndex };
  return {
    sheet: next,
    inverse: { ...op, target: inverseTarget, before: op.after, after: op.before },
  };
}

/**
 * 落一条改列宽 op：把覆盖这一列的那几段 `<col>` 换成记录里的新段。
 *
 * @param sheet 现场载荷。
 * @param op 改列宽 op。
 * @returns 新载荷与逆操作，或跳过理由。
 */
function applyColumnWidthOp(sheet: SheetPayload, op: SetColumnWidthOp): ApplyOutcome {
  const colIndex = resolveTrack(sheet, op.target);
  if (colIndex === null) return { reason: "认不出这一列：它的宽度段已被改动" };
  const shift = colIndex - op.target.index;
  const expected = op.beforeColumns.map((column) => shiftColumn(column, shift));
  const band = columnsInBand(sheet.columns, expected, colIndex);
  const sameShape =
    band.length === expected.length &&
    band.every((column, slot) => {
      const recorded = expected[slot];
      return recorded !== undefined && sameColumn(column, recorded);
    });
  if (!sameShape) return { reason: "这一列的宽度段与记录时的分段不一致，不重排用户的列宽" };
  const others = sheet.columns.filter((column) => !band.includes(column));
  const rebuilt = op.afterColumns.map((column) => shiftColumn(column, shift));
  return {
    // 按列号排好再放回：Excel 自己写出的 `<cols>` 是有序的，我们打散顺序会让文件与它产的对不上。
    sheet: { ...sheet, columns: [...others, ...rebuilt].sort((left, right) => left.minColIndex - right.minColIndex) },
    inverse: { ...op, beforeColumns: op.afterColumns, afterColumns: op.beforeColumns },
  };
}

/** 整段列宽定义按插删的位移平移；位移为 0 时原样返回，避免造出多余副本。 */
function shiftColumn(column: SheetColumn, shift: number): SheetColumn {
  if (shift === 0) return column;
  return { ...column, minColIndex: column.minColIndex + shift, maxColIndex: column.maxColIndex + shift };
}

/**
 * 取出现场 `<col>` 里落在"这条 op 要动的那几段"范围内的条目。
 *
 * @param columns 现场的列定义。
 * @param expected 记录里的原分段（已按位移平移过）。
 * @param colIndex 解析出的列号；分段为空（这列原本没有自定义宽度）时用它当唯一范围。
 * @returns 落在范围内的现场分段，数组顺序保持现场顺序。
 */
function columnsInBand(columns: SheetColumn[], expected: SheetColumn[], colIndex: number): SheetColumn[] {
  const bandStart = expected.length ? Math.min(...expected.map((column) => column.minColIndex)) : colIndex;
  const bandEnd = expected.length ? Math.max(...expected.map((column) => column.maxColIndex)) : colIndex;
  return columns.filter((column) => column.maxColIndex >= bandStart && column.minColIndex <= bandEnd);
}

/** 单条 op 的分派；指向别的表就整条跳过。 */
function applyOne(sheet: SheetPayload, op: SheetCoreOp): ApplyOutcome {
  if (op.partPath !== sheet.partPath) return { reason: "这条 op 指向的是另一张表" };
  return op.kind === "setCellValue" ? applyCellOp(sheet, op) : applyColumnWidthOp(sheet, op);
}

/** 顺序应用一批 op，失败的挑出去；`inverses` 与已落地的 `ops` 同序。 */
function applyOps(sheet: SheetPayload, ops: SheetCoreOp[]): SheetOpsResult {
  let current = sheet;
  const opsApplied: SheetCoreOp[] = [];
  const inverses: SheetCoreOp[] = [];
  const skipped: SheetSkippedOp[] = [];
  for (const op of ops) {
    const outcome = applyOne(current, op);
    if ("reason" in outcome) {
      skipped.push({ op, reason: outcome.reason });
      continue;
    }
    current = outcome.sheet;
    opsApplied.push(op);
    inverses.push(outcome.inverse);
  }
  return { ok: skipped.length === 0, before: sheet, after: current, ops: opsApplied, inverses, skipped };
}
