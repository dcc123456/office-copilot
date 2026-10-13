/**
 * 表格的结构化改动（spec 阶段 3）：每个 op 自带逆操作，撤销按内容指纹重定位。
 *
 * 为什么落在 `sheet-io` 而不是桌面的 `src/ai`：阶段 4 的 `packages/sheet-tools` 要产出同一批 op，
 * 而依赖方向只允许 `apps → packages`（`AGENTS.md` §5.1）；op 长在渲染层就会被人复制第二套。
 * 形状与 Word 的 `CoreOp` 同构：`before` 与 `after` 都存在 op 里，逆操作就是两者交换，
 * 落地前先核对现场等于 `before`，对不上就整条跳过并报"部分撤销"（§2.4/§2.5）。
 */
import { formatA1, formatRange, indexToColName, parseA1, parseRange, rangeContains, type A1Address, type A1Range } from "./address.js";
import { resolveColumnRef } from "./analysis/sheet-table.js";
import { parseUserCellText } from "./cell-text.js";
import { evaluateFormula, formulaReferencedRanges } from "./formula-eval.js";
import { shiftFormula } from "./formula-refs.js";
import { columnWidthPxToChars } from "./geometry.js";
import type { CellScalarValue, SheetCell, SheetColumn, SheetLike as SheetPayload, SheetMerge, SheetRowMeta } from "./model.js";
import { isIdentityPermutation, sortKeyOf, sortPermutation, type SortKey } from "./sheet-sort.js";
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

/**
 * 改一列的宽度。
 *
 * 存的是"覆盖这一列的那几段 `<col>`"的整段前后像，逆操作把三段换回去即可；
 * 分段与记录不一致时（同一批里前一条 op 劈过段）只要目标列自己的宽度没变，
 * 就按现场重算劈法落地，不许因为表示法变了就把整条跳过。
 */
export interface SetColumnWidthOp {
  kind: "setColumnWidth";
  partPath: string;
  /** `axis` 恒为 `col` 的地址。 */
  target: SheetTrackAddress;
  beforeColumns: SheetColumn[];
  afterColumns: SheetColumn[];
}

/** 一条带（连续若干行或若干列）在改动前的全部内容，删除与撤销都凭它核对现场。 */
export interface SheetTrackBand {
  cells: SheetCell[];
  rows: SheetRowMeta[];
  columns: SheetColumn[];
  merges: SheetMerge[];
}

/**
 * 插入或删除连续的整行/整列。
 *
 * 与其它 op 同构：`removed` 是这一带改动前的内容前像，逆操作就是"kind 取反 + 同一个前像"。
 * 插入时前像是空集，撤销插入 = 删除这一带，只有它**还空着**才允许动手（否则会把用户后来写的数据删掉）。
 */
export interface ShiftTracksOp {
  kind: "insertTracks" | "deleteTracks";
  partPath: string;
  axis: SheetAxis;
  /** 这一带起始位置的内容寻址：插入时锚在"被挤下去的那一条"上，删除时锚在"被删的第一条"上。 */
  target: SheetTrackAddress;
  count: number;
  /** 被删带（或待插带）在构造时的内容。 */
  removed: SheetTrackBand;
}

/**
 * 按一列重排一块矩形里的整行（spec 决策 D7）。
 *
 * `before` 与 `after` 是**同一块矩形的两份镜像**（排之前、排之后），逆操作就是两者交换——
 * 所以 `revertSheetOps` 不需要认识"排序"这件事，撤销也不靠整表前像。
 * 带外一格都不碰；带内的公式、探出区域的记录尾巴、劈不开的合并区在造 op 时就被拒了（D7 三条拒绝）。
 */
export interface SortRowsOp {
  kind: "sortRows";
  partPath: string;
  /** 带首行的内容寻址；`before`/`after` 里的位置与 `target.index` 同框（与 `ShiftTracksOp` 同一约定）。 */
  target: SheetTrackAddress;
  /** 参与重排的矩形（0 基，含边界），建 op 时的坐标系；落地时按现场位移平移。 */
  range: A1Range;
  before: SheetTrackBand;
  after: SheetTrackBand;
}

/** 表格的 op 全集（合并区的编辑还没落地，届时再加一种 op）。 */
export type SheetCoreOp = SetCellValueOp | SetColumnWidthOp | ShiftTracksOp | SortRowsOp;

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
 * @returns `SheetOpsResult`；`after` 永远是新对象。批里改过格子内容时，被牵连的公式缓存会作为
 *   **额外的 `setCellValue` op** 追加进 `ops`/`inverses`（spec D8），所以调用方看到的条数可能比提交的多。
 */
export function runSheetOps(sheet: SheetPayload, ops: SheetCoreOp[]): SheetOpsResult {
  return applyOps(sheet, ops, false);
}

/**
 * 撤销：把逆操作队列倒序应用回去。
 *
 * @param sheet 现场载荷（可能已被用户后来的手工编辑改过）。
 * @param inverses `runSheetOps` 记下的逆操作队列（含 D8 那条缓存刷新自己的逆操作）。
 * @returns `SheetOpsResult`；认不出目标（内容变了、行列被挪过）的条目进 `skipped`，
 *   调用方必须据此报"部分撤销"，不许猜一个位置继续写（§2.5）。撤销批**不再刷新一轮公式缓存**：
 *   它的契约是回到批前那个样子，而不是顺手把别的格子"算对"。
 */
export function revertSheetOps(sheet: SheetPayload, inverses: SheetCoreOp[]): SheetOpsResult {
  return applyOps(sheet, inverses.slice().reverse(), true);
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
 * 公式算出来的值对应到格子的 `kind`（写公式与刷缓存两处要用同一套判定，§3.2）。
 *
 * @param value 求值结果；`null` 是空格。
 * @returns `kind` 字段；数值与空都记 `number`（与读入端 `t` 缺失时的判定一致）。
 */
function formulaCellParts(value: CellScalarValue): { kind: SheetCell["kind"] } {
  return { kind: typeof value === "string" ? "string" : typeof value === "boolean" ? "boolean" : "number" };
}

/**
 * 写一条公式，并按 spec 决策 D4 同时算出缓存值。
 *
 * `<f>` 是让 Excel/WPS 打开后能重算的活公式，`<v>` 是让 app 当场能显示数字。本地算不出来时
 * （白名单外的函数、volatile、循环引用）**照常写 `<f>` 但不写 `<v>`**，`recalcNeeded` 为 true，
 * 调用方必须据此提示"该格需 Excel 重算"，不许拿过期数字顶上（验收 A4）。
 * 与 `setCellTextOp` 同一条硬约束：保留原格的 `styleIndex`/`attrs`，只清掉与旧值绑定的东西（§2.3）。
 *
 * @param sheet 当前表的载荷。
 * @param address 大写 A1 地址。
 * @param formula 公式体；带前导 `=` 会被去掉，空串视为无效。
 * @returns `{ op, recalcNeeded }`；地址解析不出来或公式为空时返回 null。
 */
export function setCellFormulaOp(sheet: SheetPayload, address: string, formula: string): { op: SetCellValueOp; recalcNeeded: boolean } | null {
  const parsedAddress = parseA1(address);
  const body = formula.trim().replace(/^=/, "");
  if (!parsedAddress || body === "") return null;
  const evaluated = evaluateFormula(body, sheet, address.toUpperCase());
  const previous = sheet.cells[address] ?? null;
  const value: CellScalarValue = evaluated.ok ? evaluated.value : null;
  const next: SheetCell = {
    address,
    colIndex: parsedAddress.colIndex,
    rowIndex: parsedAddress.rowIndex,
    ...formulaCellParts(value),
    value,
    formula: body,
    ...(evaluated.ok ? { cachedValue: value } : {}),
    ...(previous ? { styleIndex: previous.styleIndex, attrs: previous.attrs } : {}),
  };
  return {
    op: {
      kind: "setCellValue",
      partPath: sheet.partPath,
      target: { row: trackAddress(sheet, "row", parsedAddress.rowIndex), colIndex: parsedAddress.colIndex },
      before: previous,
      after: next,
    },
    recalcNeeded: !evaluated.ok,
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
  const { covering, rebuilt } = splitColumnBand(sheet.columns, colIndex, widthChars);
  const single = covering.length === 1 ? covering[0] : undefined;
  if (single && single.minColIndex === colIndex && single.maxColIndex === colIndex && single.widthChars === widthChars) {
    return null;
  }
  return {
    kind: "setColumnWidth",
    partPath: sheet.partPath,
    target: trackAddress(sheet, "col", colIndex),
    beforeColumns: covering.map((column) => ({ ...column })),
    afterColumns: rebuilt,
  };
}

/**
 * 把"这一列改成这么宽"落到列定义上：覆盖它的分段按左右两截保留，中间单独一段用新宽度。
 *
 * @param columns 列定义（现场的或建 op 时的，取决于哪一批要劈）。
 * @param colIndex 目标列号（0 基）。
 * @param widthChars 目标宽度（字符宽，像素换算只在 `geometry.ts` 发生，§2.1）。
 * @returns `covering` 是被盖住的旧分段，`rebuilt` 是劈完之后替换它们的新分段（按列号排好）。
 */
function splitColumnBand(columns: SheetColumn[], colIndex: number, widthChars: number): { covering: SheetColumn[]; rebuilt: SheetColumn[] } {
  const covering = columns.filter((column) => colIndex >= column.minColIndex && colIndex <= column.maxColIndex);
  const rebuilt: SheetColumn[] = [];
  for (const column of covering) {
    if (column.minColIndex < colIndex) rebuilt.push({ ...column, maxColIndex: colIndex - 1 });
    if (column.maxColIndex > colIndex) rebuilt.push({ ...column, minColIndex: colIndex + 1 });
  }
  // 相邻列的其余属性（样式、隐藏等）原样继承，只有宽度和 customWidth 是我们改的。
  const inherited = covering[0];
  rebuilt.push({ ...(inherited ?? {}), minColIndex: colIndex, maxColIndex: colIndex, widthChars, customWidth: true });
  rebuilt.sort((left, right) => left.minColIndex - right.minColIndex);
  return { covering, rebuilt };
}

/** 空带：插入时的前像，撤销插入凭它确认"这一带还空着"。 */
const EMPTY_BAND: SheetTrackBand = { cells: [], rows: [], columns: [], merges: [] };

/**
 * 造一条"在第 `index` 条之前插入 `count` 条空行/空列"的 op。
 *
 * @param sheet 当前表的载荷。
 * @param axis 插行还是插列。
 * @param index 0 基插入点（这条 op 锚在它当前所在的那一条上）。
 * @param count 插入条数，默认 1。
 * @returns 交给 `runSheetOps` 的 op；`count < 1` 或下标非法时返回 null（没什么可插的）。
 */
export function insertTracksOp(sheet: SheetPayload, axis: SheetAxis, index: number, count = 1): ShiftTracksOp | null {
  if (count < 1 || !Number.isInteger(index) || index < 0) return null;
  return { kind: "insertTracks", partPath: sheet.partPath, axis, target: trackAddress(sheet, axis, index), count, removed: EMPTY_BAND };
}

/**
 * 造一条"从第 `index` 条起删除 `count` 条整行/整列"的 op。
 *
 * 构造时就把这一带的内容拍成前像：落地前要凭它核对现场，撤销时要凭它把内容插回去。
 *
 * @param sheet 当前表的载荷。
 * @param axis 删行还是删列。
 * @param index 0 基起点。
 * @param count 删除条数，默认 1。
 * @returns op；`count < 1` 或下标非法时返回 null。
 */
export function deleteTracksOp(sheet: SheetPayload, axis: SheetAxis, index: number, count = 1): ShiftTracksOp | null {
  if (count < 1 || !Number.isInteger(index) || index < 0) return null;
  return { kind: "deleteTracks", partPath: sheet.partPath, axis, target: trackAddress(sheet, axis, index), count, removed: bandSnapshot(sheet, axis, index, count) };
}

/** 一次能安全记账的带内格子数上限（D7：op 里同时装前后两份镜像，还要整个进 `modlog` 的 JSON）。 */
const SORT_BAND_CELL_LIMIT = 50_000;

/** `sortRowsOp` 的入参。 */
export interface SortRowsInput {
  /** A1 区域写法，如 `A2:D10`。行范围就是参与重排的数据行，列范围必须盖住整条记录；整列/整行写法（`A:C`、`2:4`）一律拒。 */
  range: string;
  /** 键列引用：列字母（`C`）或表头原文（`销售额`），与读工具共用同一个 `resolveColumnRef` 口径。 */
  column: string;
  /** 升序还是降序，默认升序；只反转非空格之间的次序，空格恒在带尾。 */
  order?: "asc" | "desc";
  /** 首行是否当表头，只影响表头文字形式的键列引用；区域本身按字面落地，不做"自动跳过表头"。 */
  hasHeader?: boolean;
}

/** `sortRowsOp` 的结果：要么给一条能落地的 op，要么给一句给人和模型看的理由（不抛异常）。 */
export type SortRowsResult = { ok: true; op: SortRowsOp } | { ok: false; reason: string };

/**
 * 造一条"按某一列重排这一块矩形里的整行"的 op（spec 决策 D7）。
 *
 * 行序由引擎从现场表算（`AGENTS.md` §2.11）：调用方只交区域、键列与升降，不许自己报一个新行序。
 * 造出来的 `before`/`after` 是同一矩形的两份镜像，撤销就是把两者交换；三条拒绝（带内公式、
 * 这些行在区域外还有格子、劈不开的合并区）在这里就拦掉，不给"落地一半"的机会。
 *
 * @param sheet 当前表的载荷。
 * @param input 见 `SortRowsInput`。
 * @returns `{ok:true, op}`；区域写错、键列认不出、只有 0/1 行、触发三条拒绝、或已经是要的次序时 `{ok:false, reason}`。
 */
export function sortRowsOp(sheet: SheetPayload, input: SortRowsInput): SortRowsResult {
  const parsed = parseRange(input.range.trim().toUpperCase());
  if (!parsed) return { ok: false, reason: `区域「${input.range}」没法解析，请写成 A2:D10 这种起止都明确的形式。` };
  if (parsed.end.rowIndex > sheet.lastRowIndex || parsed.end.colIndex > sheet.lastColIndex) {
    return { ok: false, reason: `区域「${input.range}」超出了这张表的数据区（到第 ${sheet.lastRowIndex + 1} 行、${indexToColName(sheet.lastColIndex)} 列），先缩小区域。` };
  }
  if (parsed.end.rowIndex <= parsed.start.rowIndex) return { ok: false, reason: `区域「${input.range}」不足两行，没什么可排的。` };
  const direction = input.order === "desc" ? "desc" : "asc";
  const keyColIndex = resolveColumnRef(sheet, input.column, input.hasHeader ?? true);
  if (keyColIndex === null) return { ok: false, reason: `认不出键列「${input.column}」，请改用列字母（如 C）或表头原文。` };
  if (keyColIndex < parsed.start.colIndex || keyColIndex > parsed.end.colIndex) {
    return { ok: false, reason: `键列 ${indexToColName(keyColIndex)} 不在排序区域 ${formatRange(parsed)} 里。` };
  }
  const guard = sortBandGuardReason(sheet, parsed);
  if (guard) return { ok: false, reason: guard };

  const keys = rectKeys(sheet, parsed, keyColIndex);
  const permutation = sortPermutation(keys, direction);
  if (isIdentityPermutation(permutation)) {
    return { ok: false, reason: `${formatRange(parsed)} 按 ${indexToColName(keyColIndex)} 列已经是${direction === "desc" ? "降序" : "升序"}了，不需要改动。` };
  }
  return {
    ok: true,
    op: {
      kind: "sortRows",
      partPath: sheet.partPath,
      target: trackAddress(sheet, "row", parsed.start.rowIndex),
      range: parsed,
      before: rectImage(sheet, parsed),
      after: rectImage(sheet, parsed, permutation),
    },
  };
}

/**
 * 排序带的硬约束（D7）：这些行在区域外不许还有格子、区域内不许有公式、合并区不许劈开、格子总数不许超上限。
 *
 * 造 op 与落地各查一次：用户在提案排队期间往带里补了公式或写了一格区域外的数据，同样不许动手。
 *
 * @param sheet 现场载荷。
 * @param range 排序矩形（0 基，含边界）。
 * @returns 需要拒绝时给一句中文理由；可以排返回 null。
 */
function sortBandGuardReason(sheet: SheetPayload, range: A1Range): string | null {
  const inBandRows = (rowIndex: number) => rowIndex >= range.start.rowIndex && rowIndex <= range.end.rowIndex;
  const inRectCols = (colIndex: number) => colIndex >= range.start.colIndex && colIndex <= range.end.colIndex;
  let cellCount = 0;
  for (const cell of Object.values(sheet.cells)) {
    if (!inBandRows(cell.rowIndex)) continue;
    if (!inRectCols(cell.colIndex)) {
      return `${cell.address} 在排序区域之外还有内容，只搬区域内的列会把这一条记录劈成两半，不动它`;
    }
    if (cell.formula || cell.rawXml?.includes("<f")) {
      return `${cell.address} 带公式：Excel 搬动公式时会改写相对引用，本机没有 Excel 能实测它改成什么样（spec 决策 D7），不猜`;
    }
    cellCount += 1;
  }
  if (cellCount > SORT_BAND_CELL_LIMIT) {
    return `要排的这一带有 ${cellCount} 格，超过一次能安全记账的 ${SORT_BAND_CELL_LIMIT} 格，先把区域缩小`;
  }
  const straddle = sheet.merges.find(
    (merge) =>
      merge.range.end.rowIndex >= range.start.rowIndex &&
      merge.range.start.rowIndex <= range.end.rowIndex &&
      !(merge.range.start.rowIndex === merge.range.end.rowIndex && inRectCols(merge.range.start.colIndex) && inRectCols(merge.range.end.colIndex)),
  );
  if (straddle) return `${straddle.ref} 这个合并区跨行或探出排序区域，重排后合不回原样，不动它`;
  return null;
}

/**
 * 取排序键列在这一带上逐行的键值（行序即现场行序）。
 *
 * @param sheet 载荷。
 * @param range 排序矩形。
 * @param keyColIndex 键列号（0 基，调用方已确认落在矩形列范围内）。
 * @returns 与带高等长的键数组；空位在 `sortKeyOf` 里退化成空格秩。
 */
function rectKeys(sheet: SheetPayload, range: A1Range, keyColIndex: number): SortKey[] {
  const letter = indexToColName(keyColIndex);
  const acc: SortKey[] = [];
  for (let rowIndex = range.start.rowIndex; rowIndex <= range.end.rowIndex; rowIndex += 1) {
    acc.push(sortKeyOf(sheet.cells[`${letter}${rowIndex + 1}`]));
  }
  return acc;
}

/**
 * 拍下这一矩形的镜像。
 *
 * @param sheet 载荷。
 * @param range 排序矩形（0 基，含边界）。
 * @param permutation 给出时按"新位置 ← 带内原行"重排（`sortPermutation` 的产物）；省略即现场原样。
 *   行元信息与单行合并区随记录走——用户藏起来的那一行本该跟着它的数据一起挪。
 * @returns `SheetTrackBand` 形状的前像/后像（`columns` 恒空：行轴不动列宽）。
 */
function rectImage(sheet: SheetPayload, range: A1Range, permutation?: number[]): SheetTrackBand {
  const rowCount = range.end.rowIndex - range.start.rowIndex + 1;
  const cellsByRow = new Map<number, SheetCell[]>();
  for (const cell of Object.values(sheet.cells)) {
    if (cell.rowIndex < range.start.rowIndex || cell.rowIndex > range.end.rowIndex) continue;
    if (cell.colIndex < range.start.colIndex || cell.colIndex > range.end.colIndex) continue;
    const bucket = cellsByRow.get(cell.rowIndex);
    if (bucket) bucket.push(cell);
    else cellsByRow.set(cell.rowIndex, [cell]);
  }
  const metaByRow = new Map(sheet.rows.filter((row) => row.rowIndex >= range.start.rowIndex && row.rowIndex <= range.end.rowIndex).map((row) => [row.rowIndex, row]));
  const mergesByRow = new Map<number, SheetMerge[]>();
  for (const merge of sheet.merges) {
    if (merge.range.start.rowIndex < range.start.rowIndex || merge.range.start.rowIndex > range.end.rowIndex) continue;
    const bucket = mergesByRow.get(merge.range.start.rowIndex);
    if (bucket) bucket.push(merge);
    else mergesByRow.set(merge.range.start.rowIndex, [merge]);
  }
  const cells: SheetCell[] = [];
  const rows: SheetRowMeta[] = [];
  const merges: SheetMerge[] = [];
  for (let position = 0; position < rowCount; position += 1) {
    const source = range.start.rowIndex + (permutation?.[position] ?? position);
    const target = range.start.rowIndex + position;
    for (const cell of cellsByRow.get(source) ?? []) {
      cells.push({ ...cell, rowIndex: target, address: formatA1({ colIndex: cell.colIndex, rowIndex: target }) });
    }
    const meta = metaByRow.get(source);
    if (meta) rows.push({ ...meta, rowIndex: target });
    for (const merge of mergesByRow.get(source) ?? []) merges.push(moveMerge(merge, true, target - source, 0));
  }
  return { cells, rows, columns: [], merges };
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
 * 落一条改列宽 op。
 *
 * 先走**精确路径**：现场的分段与记录里逐段一致时整段换回去，撤销能把 `<cols>` 还原成原来那一段。
 * 不一致时走**按现场重算**：只核对目标列自己的宽度没被改过，就地劈段（相邻列的宽度按现场保留）。
 * 有了后者，同一批里前一条列宽 op 劈开的段才不会把后面几条全拦下来（阶段 4 的批内死角）。
 *
 * @param sheet 现场载荷。
 * @param op 改列宽 op。
 * @returns 新载荷与逆操作（两者都在落地时的坐标系里），或跳过理由。
 */
function applyColumnWidthOp(sheet: SheetPayload, op: SetColumnWidthOp): ApplyOutcome {
  const colIndex = resolveTrack(sheet, op.target);
  if (colIndex === null) return { reason: "认不出这一列：它已被插入/删除的列挪走，或整列内容被改过" };
  const widthChars = pieceCovering(op.afterColumns, op.target.index)?.widthChars;
  if (widthChars === undefined) return { reason: "这条列宽改动没记下目标宽度" };
  const shift = colIndex - op.target.index;
  const expected = op.beforeColumns.map((column) => shiftColumn(column, shift));
  const band = columnsInBand(sheet.columns, expected, colIndex);
  const exact = sameBandShape(band, expected);
  if (!exact && !sameColumnWidth(pieceCovering(band, colIndex), pieceCovering(expected, colIndex))) {
    return { reason: `${columnLabel(colIndex)} 列现在的宽度与这条改动记录时的不一致，不覆盖用户的列宽` };
  }
  const split = exact
    ? { covering: band, rebuilt: op.afterColumns.map((column) => shiftColumn(column, shift)) }
    : splitColumnBand(sheet.columns, colIndex, widthChars);
  const others = sheet.columns.filter((column) => !split.covering.includes(column));
  const applied: SetColumnWidthOp = {
    ...op,
    target: { ...op.target, index: colIndex },
    beforeColumns: split.covering.map((column) => ({ ...column })),
    afterColumns: split.rebuilt,
  };
  return {
    // 按列号排好再放回：Excel 自己写出的 `<cols>` 是有序的，我们打散顺序会让文件与它产的对不上。
    sheet: { ...sheet, columns: [...others, ...split.rebuilt].sort((left, right) => left.minColIndex - right.minColIndex) },
    inverse: { ...applied, beforeColumns: applied.afterColumns, afterColumns: applied.beforeColumns },
  };
}

/** 现场的分段是否与记录里逐段一致（一致才有把握整段换回，把 `<cols>` 还原成原来的写法）。 */
function sameBandShape(band: SheetColumn[], expected: SheetColumn[]): boolean {
  return (
    band.length === expected.length &&
    band.every((column, slot) => {
      const recorded = expected[slot];
      return recorded !== undefined && sameColumn(column, recorded);
    })
  );
}

/** 取分段里盖住这一列的那一段（`<col>` 段本不该重叠，真重叠时取第一段）。 */
function pieceCovering(columns: SheetColumn[], colIndex: number): SheetColumn | undefined {
  return columns.find((column) => colIndex >= column.minColIndex && colIndex <= column.maxColIndex);
}

/**
 * 这一列自己的宽度有没有被改过。
 *
 * 不比覆盖范围：把 `{0–3}` 劈成三段、我们那列的宽度没动，属于表示法变化而不是用户改动，
 * 拦下来会让批内的后续 op 全落不了地（`AGENTS.md` §2.5 要的是不牵连用户的数据）。
 *
 * @param left 现场那一段（这列没有自定义宽度时是 undefined）。
 * @param right 建 op 时那一段。
 * @returns 宽度与其余属性都一致时为 true；一段有一无时为 false。
 */
function sameColumnWidth(left: SheetColumn | undefined, right: SheetColumn | undefined): boolean {
  if (!left || !right) return left === right;
  return (
    left.widthChars === right.widthChars &&
    left.hidden === right.hidden &&
    left.customWidth === right.customWidth &&
    sameAttributes(left.attrs, right.attrs)
  );
}

/** 列号 → 给人看的列名（A、B、…）；越界时退回下标，不让错误消息本身出错。 */
function columnLabel(colIndex: number): string {
  return indexToColName(colIndex) || `${colIndex + 1} 列`;
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

/** 整段列宽定义按插删的位移平移；位移为 0 时原样返回，避免造出多余副本。 */
function shiftColumn(column: SheetColumn, shift: number): SheetColumn {
  if (shift === 0) return column;
  return { ...column, minColIndex: column.minColIndex + shift, maxColIndex: column.maxColIndex + shift };
}

/**
 * 落一条插删行列 op：核对现场 → 位移格子/行元信息/列宽段/合并区 → 平移公式引用。
 *
 * 位移规则逐条对着 Excel 的行为实现（出处见 `formula-refs.ts` 与 `docs/excel-support-spec.md` 阶段 3）：
 * 跨边的 `<col>` 段与合并区在插入时**变宽**、在删除时**收缩**，而不是劈开或整块消失。
 *
 * @param sheet 现场载荷。
 * @param op 插删 op（逆操作形状相同，只有 `kind` 取反）。
 * @returns 新载荷与逆操作，或跳过理由。
 */
function applyShiftOp(sheet: SheetPayload, op: ShiftTracksOp): ApplyOutcome {
  const isRow = op.axis === "row";
  const trackLabel = isRow ? "行" : "列";
  const from = resolveTrack(sheet, op.target);
  if (from === null) return { reason: `认不出第 ${op.target.index + 1} ${trackLabel}：内容已被改动，或它被插入/删除的${trackLabel}挪走了` };
  const shift = from - op.target.index;
  if (op.kind === "deleteTracks") {
    const mismatch = bandMismatchReason(sheet, op, from, shift);
    if (mismatch) return { reason: mismatch };
  }
  // 只有带内容的插删（删除、以及撤销删除）需要这道检查；纯插入的几何来回是精确的。
  if (!isEmptyBand(op.removed)) {
    const split = mergeSplitReason(sheet, op.axis, from, from + op.count - 1);
    if (split) return { reason: split };
  }
  const moved = moveTracks(sheet, op, from, shift);
  if ("reason" in moved) return moved;
  const next: SheetPayload = { ...sheet, ...moved.tracks };
  // 列号变了以后旧 `spans` 就是假提示；行轴插删不改一行之内的列分布，不用重算。
  if (!isRow) recomputeRowSpans(next);
  recalculateSheetExtent(next);
  return {
    // 逆操作的前像要搬到落地之后的坐标系：`removed` 里的位置与 `target.index` 始终是同一套坐标。
    sheet: next,
    inverse: {
      ...op,
      kind: op.kind === "insertTracks" ? "deleteTracks" : "insertTracks",
      target: trackAddress(next, op.axis, from),
      removed: shiftBand(op.removed, op.axis, shift),
    },
  };
}

/** 插删之后的四样东西（格子、行元信息、列宽段、合并区）；认不出安全落点时只给理由。 */
type MoveOutcome = { tracks: Pick<SheetPayload, "cells" | "rows" | "columns" | "merges"> } | { reason: string };

/**
 * 算出位移之后的表内容。
 *
 * @param sheet 现场载荷。
 * @param op 插删 op。
 * @param from 解析出的 0 基带起点。
 * @param shift `from` 与建 op 时记下的下标之差（行列被用户挪过位置时用它把前像搬到位）。
 * @returns 新的四样东西，或拒绝理由。
 */
function moveTracks(sheet: SheetPayload, op: ShiftTracksOp, from: number, shift: number): MoveOutcome {
  const isRow = op.axis === "row";
  const inserting = op.kind === "insertTracks";
  const to = from + op.count - 1;
  const delta = inserting ? op.count : -op.count;
  // 插入时带是新生的，现场没有旧内容要跳过；删除时"起点之后"从带的下一条算起。
  const moveFrom = inserting ? from : to + 1;
  const positionOf = (cell: SheetCell) => (isRow ? cell.rowIndex : cell.colIndex);

  const cells: Record<string, SheetCell> = {};
  for (const cell of Object.values(sheet.cells)) {
    if (!inserting && positionOf(cell) >= from && positionOf(cell) <= to) continue;
    const movedBy = positionOf(cell) >= moveFrom ? delta : 0;
    const colIndex = cell.colIndex + (isRow ? 0 : movedBy);
    const rowIndex = cell.rowIndex + (isRow ? movedBy : 0);
    if (cell.formula) {
      const shiftedFormula = shiftFormula(cell.formula, op.axis, from, delta, inserting ? undefined : { startIndex: from, endIndex: to });
      if (!shiftedFormula.ok) return { reason: shiftedFormula.reason };
      placeCell(cells, { ...cell, formula: shiftedFormula.formula }, colIndex, rowIndex);
      continue;
    }
    if (movedBy !== 0 && cell.rawXml?.includes("<f")) return { reason: "有原样保留的格子带着公式，位移会把公式和它引用的格子拆散" };
    placeCell(cells, cell, colIndex, rowIndex);
  }
  let rows = isRow ? shiftRowMetas(sheet.rows, moveFrom, delta, inserting ? undefined : { start: from, end: to }) : sheet.rows;
  let columns = isRow ? sheet.columns : shiftColumnBands(sheet.columns, from, op.count, inserting, to);
  let merges = shiftMerges(sheet.merges, isRow, from, op.count, inserting, to);
  if (inserting && !isEmptyBand(op.removed)) {
    for (const cell of op.removed.cells) {
      placeCell(cells, cell, cell.colIndex + (isRow ? 0 : shift), cell.rowIndex + (isRow ? shift : 0));
    }
    if (isRow) rows = [...rows, ...op.removed.rows.map((row) => ({ ...row, rowIndex: row.rowIndex + shift }))];
    else columns = [...columns, ...op.removed.columns.map((column) => shiftColumn(column, shift))];
    merges = [...merges, ...op.removed.merges.map((merge) => moveMerge(merge, isRow, shift, shift))];
  }
  return { tracks: { cells, rows, columns: sortColumns(columns), merges: sortMerges(merges) } };
}

/** 把一个格子按新位置放回结果表：`address` 与 `colIndex/rowIndex` 必须同步（spec D1）。 */
function placeCell(cells: Record<string, SheetCell>, cell: SheetCell, colIndex: number, rowIndex: number): void {
  const address = formatA1({ colIndex, rowIndex });
  const moved: SheetCell = { ...cell, address, colIndex, rowIndex };
  const rawXml = cell.rawXml ? retargetRawCellXml(cell.rawXml, address) : undefined;
  if (rawXml) moved.rawXml = rawXml;
  else delete moved.rawXml;
  cells[address] = moved;
}

/**
 * 改写 `<c>` 原文里的 `r` 属性，属性顺序与其余属性不动。
 *
 * @param rawXml 读入时留下的整格原文。
 * @param address 位移后的新 A1 地址。
 * @returns 换好地址的原文；匹配不上 `r` 属性时原样返回（这类格读入时就没有地址，已被记进 warnings）。
 */
function retargetRawCellXml(rawXml: string, address: string): string {
  return rawXml.replace(/(<c\b[^>]*?\br=")[^"]*(")/, `$1${address}$2`);
}

/**
 * 位移行元信息。
 *
 * @param rows 现场行元信息。
 * @param moveFrom 从哪一行起跟着挪。
 * @param delta 位移量。
 * @param cut 删除时的带区间；整行落在里面的行元信息随带消失。
 * @returns 新的行元信息队列。
 */
function shiftRowMetas(rows: SheetRowMeta[], moveFrom: number, delta: number, cut?: { start: number; end: number }): SheetRowMeta[] {
  const acc: SheetRowMeta[] = [];
  for (const row of rows) {
    if (cut && row.rowIndex >= cut.start && row.rowIndex <= cut.end) continue;
    acc.push(row.rowIndex >= moveFrom ? { ...row, rowIndex: row.rowIndex + delta } : row);
  }
  return acc;
}

/**
 * 位移 `<col>` 分段。
 *
 * Excel 的行为：插入落在段里时这段**变宽**，删除吃掉一段时这段**收缩**，整段被删则消失。
 *
 * @param columns 现场分段。
 * @param from 0 基插入点/带起点。
 * @param count 插入或删除的列数。
 * @param inserting 是插入还是删除。
 * @param to 删除时的带终点。
 * @returns 新的分段队列（按列号排好）。
 */
function shiftColumnBands(columns: SheetColumn[], from: number, count: number, inserting: boolean, to: number): SheetColumn[] {
  const acc: SheetColumn[] = [];
  for (const column of columns) {
    if (inserting) {
      const [minColIndex, maxColIndex] = growRange(column.minColIndex, column.maxColIndex, from, count);
      acc.push({ ...column, minColIndex, maxColIndex });
      continue;
    }
    if (column.minColIndex >= from && column.maxColIndex <= to) continue;
    const [minColIndex, maxColIndex] = shrinkRange(column.minColIndex, column.maxColIndex, from, to, count);
    acc.push({ ...column, minColIndex, maxColIndex });
  }
  return acc;
}

/** 插入时段的两个端点各自"到了插入点之后就跟着一块右移"，于是跨边的段自动变宽。 */
function growRange(start: number, end: number, from: number, count: number): [number, number] {
  return [start >= from ? start + count : start, end >= from ? end + count : end];
}

/** 删除时段里剩下的那截仍然连续：吃掉重叠量，跨左边时收在带的上一条，跨右边时起点落到带上。 */
function shrinkRange(start: number, end: number, from: number, to: number, count: number): [number, number] {
  return [start < from ? start : start > to ? start - count : from, end < from ? end : end > to ? end - count : from - 1];
}

/**
 * 位移合并区。
 *
 * @param merges 现场合并区。
 * @param isRow 动的是行还是列。
 * @param from 0 基插入点/带起点。
 * @param count 条数。
 * @param inserting 插入还是删除。
 * @param to 删除时的带终点。
 * @returns 新的合并区队列。
 */
function shiftMerges(merges: SheetMerge[], isRow: boolean, from: number, count: number, inserting: boolean, to: number): SheetMerge[] {
  const acc: SheetMerge[] = [];
  for (const merge of merges) {
    const start = isRow ? merge.range.start.rowIndex : merge.range.start.colIndex;
    const end = isRow ? merge.range.end.rowIndex : merge.range.end.colIndex;
    if (!inserting && start >= from && end <= to) continue;
    const [movedStart, movedEnd] = inserting ? growRange(start, end, from, count) : shrinkRange(start, end, from, to, count);
    acc.push(moveMerge(merge, isRow, movedStart - start, movedEnd - end));
  }
  return acc;
}

/** 按"行方向的位移量 + 列方向的位移量"重建一个合并区（`ref` 与 `range` 必须同步）。 */
function moveMerge(merge: SheetMerge, isRow: boolean, rowShift: number, colShift: number): SheetMerge {
  const range = {
    start: { colIndex: merge.range.start.colIndex + (isRow ? 0 : colShift), rowIndex: merge.range.start.rowIndex + (isRow ? rowShift : 0) },
    end: { colIndex: merge.range.end.colIndex + (isRow ? 0 : colShift), rowIndex: merge.range.end.rowIndex + (isRow ? rowShift : 0) },
  };
  return { ref: formatRange(range), range };
}

/** 列宽分段按列号排好；合并区按左上角排好——Excel 自己写出的顺序就是这样。 */
function sortColumns(columns: SheetColumn[]): SheetColumn[] {
  return columns.slice().sort((left, right) => left.minColIndex - right.minColIndex);
}

/** 合并区的稳定顺序：先按起始列再按起始行。 */
function sortMerges(merges: SheetMerge[]): SheetMerge[] {
  return merges.slice().sort((left, right) => left.range.start.rowIndex - right.range.start.rowIndex || left.range.start.colIndex - right.range.start.colIndex);
}

/**
 * 重算每行的 `spans`（"这行里有格子的列范围"，1 基）。
 *
 * @param sheet 就地更新的载荷。列轴插删之后列分布变了，旧 `spans` 就成了假提示；
 *   一行里没有格子时干脆不写（Excel 也是这么写的）。
 */
function recomputeRowSpans(sheet: SheetPayload): void {
  const colsByRow = new Map<number, number[]>();
  for (const cell of Object.values(sheet.cells)) {
    const bucket = colsByRow.get(cell.rowIndex);
    if (bucket) bucket.push(cell.colIndex);
    else colsByRow.set(cell.rowIndex, [cell.colIndex]);
  }
  sheet.rows = sheet.rows.map((row) => {
    const rest: Record<string, string> = {};
    for (const [key, value] of Object.entries(row.attrs ?? {})) if (key !== "spans") rest[key] = value;
    const cols = colsByRow.get(row.rowIndex);
    if (cols?.length) rest.spans = `${Math.min(...cols) + 1}:${Math.max(...cols) + 1}`;
    return { ...row, attrs: rest };
  });
}

/** 前像是不是空集（纯插入时为真，撤销插入才不需要"插回内容"）。 */
function isEmptyBand(band: SheetTrackBand): boolean {
  return band.cells.length === 0 && band.rows.length === 0 && band.columns.length === 0 && band.merges.length === 0;
}

/**
 * 记下"这一带现在装的是什么"，之后凭它核对现场与插回原内容。
 *
 * @param sheet 建 op 时的载荷。
 * @param axis 行还是列。
 * @param start 0 基带起点。
 * @param count 条数。
 * @returns 带内的格子、行元信息、整段落在带里的 `<col>`、整块落在带里的合并区（都是副本）。
 */
function bandSnapshot(sheet: SheetPayload, axis: SheetAxis, start: number, count: number): SheetTrackBand {
  const isRow = axis === "row";
  const end = start + count - 1;
  const inside = (position: number) => position >= start && position <= end;
  return {
    cells: Object.values(sheet.cells)
      .filter((cell) => inside(isRow ? cell.rowIndex : cell.colIndex))
      .map((cell) => ({ ...cell })),
    rows: isRow ? sheet.rows.filter((row) => inside(row.rowIndex)).map((row) => ({ ...row })) : [],
    ...bandShape(sheet, axis, start, end),
  };
}

/**
 * 把前像整体搬到另一个坐标系。
 *
 * op 的约定：`removed` 里的位置永远与 `target.index` 同框。落地时把逆操作的前像搬到位，
 * 撤销时才算得出正确的插回点（行列被用户挪过位置时 `shift` 不为 0）。
 *
 * @param band 前像。
 * @param axis 行还是列。
 * @param shift 位移量。
 * @returns 搬动后的前像（`shift` 为 0 时原样返回，避免多余副本）。
 */
function shiftBand(band: SheetTrackBand, axis: SheetAxis, shift: number): SheetTrackBand {
  if (shift === 0) return band;
  const isRow = axis === "row";
  return {
    cells: band.cells.map((cell) => ({ ...cell, colIndex: cell.colIndex + (isRow ? 0 : shift), rowIndex: cell.rowIndex + (isRow ? shift : 0) })),
    rows: band.rows.map((row) => ({ ...row, rowIndex: row.rowIndex + shift })),
    columns: band.columns.map((column) => shiftColumn(column, shift)),
    merges: band.merges.map((merge) => moveMerge(merge, isRow, shift, shift)),
  };
}

/**
 * 取出"落在这一带里的那部分"前像形状：列宽段按带裁剪，合并区只收整块落在带里的。
 *
 * 列宽段裁剪而不是整段收：整段盖住带的（两边都探出带外）只按插删增减端点，来回本来就精确，
 * 收进前像反而会在撤销时重出一份；只有被带劈开的那种才需要把带内那截单独存下来，
 * 撤销后带内一截与带外一截相邻同宽，Excel 读出与原来一段等价。合并区不能这么合，见 `mergeSplitReason`。
 *
 * @param sheet 载荷。
 * @param axis 行还是列。
 * @param start 0 基带起点。
 * @param end 0 基带终点。
 * @returns 带内的列宽段与合并区（都是副本；行轴不动列宽，`columns` 为空）。
 */
function bandShape(sheet: SheetPayload, axis: SheetAxis, start: number, end: number): Pick<SheetTrackBand, "columns" | "merges"> {
  const isRow = axis === "row";
  const clip = (itemStart: number, itemEnd: number): [number, number] | null => {
    if (itemStart < start && itemEnd > end) return null;
    const clippedStart = Math.max(itemStart, start);
    const clippedEnd = Math.min(itemEnd, end);
    return clippedStart <= clippedEnd ? [clippedStart, clippedEnd] : null;
  };
  const columns = isRow ? [] : sheet.columns.flatMap((column) => {
    const span = clip(column.minColIndex, column.maxColIndex);
    return span ? [{ ...column, minColIndex: span[0], maxColIndex: span[1] }] : [];
  });
  const merges = sheet.merges
    .filter((merge) => {
      const startOf = isRow ? merge.range.start.rowIndex : merge.range.start.colIndex;
      const endOf = isRow ? merge.range.end.rowIndex : merge.range.end.colIndex;
      return startOf >= start && endOf <= end;
    })
    .map((merge) => ({ ref: merge.ref, range: { start: { ...merge.range.start }, end: { ...merge.range.end } } }));
  return { columns, merges };
}

/**
 * 这一带是否把某个合并区劈成两半。
 *
 * 整块落在带里的合并区由前像整块存下、撤销时原样插回；整块盖住带的只增减端点，来回也精确。
 * 唯独"一半在带里一半在外面"没有可逆的表示法：收缩掉的那半没处存，撤销时又只能补出一截假合并区
 * （真合并区被劈成两块就再也合不回去），所以整条拒绝（`AGENTS.md` §2.5：不能安全回滚就明说，不许猜）。
 * 列宽段没有这个问题（相邻同宽的两段与一段等价），所以列宽段走裁剪、合并区走拒绝。
 *
 * @param sheet 现场载荷。
 * @param axis 行还是列。
 * @param from 0 基带起点。
 * @param to 0 基带终点。
 * @returns 有劈不开的合并区时给理由；没有返回 null。
 */
function mergeSplitReason(sheet: SheetPayload, axis: SheetAxis, from: number, to: number): string | null {
  const isRow = axis === "row";
  const split = sheet.merges.find((merge) => {
    const startOf = isRow ? merge.range.start.rowIndex : merge.range.start.colIndex;
    const endOf = isRow ? merge.range.end.rowIndex : merge.range.end.colIndex;
    return endOf >= from && startOf <= to && !(startOf >= from && endOf <= to) && !(startOf < from && endOf > to);
  });
  if (!split) return null;
  return `${split.ref} 这个合并区正好被这一带劈开，撤销合不回原样，不动它`;
}

/**
 * 删除前核对现场：这一带必须还是建 op 时那份内容。
 *
 * @param sheet 现场载荷。
 * @param op 删除 op。
 * @param from 解析出的带起点。
 * @param shift 现场与建 op 时的下标差（行列被用户挪过位置时用它搬动记录）。
 * @returns 不一致的理由；一致返回 null。
 */
function bandMismatchReason(sheet: SheetPayload, op: ShiftTracksOp, from: number, shift: number): string | null {
  const isRow = op.axis === "row";
  const to = from + op.count - 1;
  const inside = (position: number) => position >= from && position <= to;
  const liveCells = Object.values(sheet.cells).filter((cell) => inside(isRow ? cell.rowIndex : cell.colIndex));
  if (liveCells.length !== op.removed.cells.length) return `第 ${from + 1} ${isRow ? "行" : "列"}起这一带的格子数与记录时不同，不删`;
  for (const recorded of op.removed.cells) {
    const colIndex = recorded.colIndex + (isRow ? 0 : shift);
    const rowIndex = recorded.rowIndex + (isRow ? shift : 0);
    const live = liveCells.find((cell) => cell.colIndex === colIndex && cell.rowIndex === rowIndex);
    const address = formatA1({ colIndex, rowIndex });
    if (!live || !sameCellContent(live, recorded)) return `${address} 现在装的不是这条记录时的内容，不删`;
  }
  const liveRows = isRow ? sheet.rows.filter((row) => inside(row.rowIndex)) : [];
  if (liveRows.length !== op.removed.rows.length) return `第 ${from + 1} 行的行元信息与记录时不同，不删`;
  for (const row of op.removed.rows) {
    const live = liveRows.find((item) => item.rowIndex === row.rowIndex + shift);
    if (!live || !sameRowMeta(live, row)) return `第 ${row.rowIndex + 1 + shift} 行的行高/格式与记录时不同，不删`;
  }
  const live = bandShape(sheet, op.axis, from, to);
  const liveColumns = live.columns;
  const recordedColumns = op.removed.columns.map((column) => shiftColumn(column, shift));
  if (liveColumns.length !== recordedColumns.length || !recordedColumns.every((column) => liveColumns.some((liveColumn) => sameColumn(liveColumn, column)))) {
    return "这一带的列宽分段与记录时不同，不删";
  }
  const liveMergeRefs = live.merges.map((merge) => merge.ref);
  const recordedMergeRefs = op.removed.merges.map((merge) => moveMerge(merge, isRow, shift, shift).ref);
  if (liveMergeRefs.length !== recordedMergeRefs.length || !recordedMergeRefs.every((ref) => liveMergeRefs.includes(ref))) {
    return "这一带的合并区与记录时不同，不删";
  }
  return null;
}

/** 两条行元信息是否等价（含没建模但必须原样搬的其余属性）。 */
function sameRowMeta(left: SheetRowMeta, right: SheetRowMeta): boolean {
  return (
    left.heightPt === right.heightPt &&
    left.customHeight === right.customHeight &&
    left.hidden === right.hidden &&
    left.styleIndex === right.styleIndex &&
    sameAttributes(left.attrs, right.attrs)
  );
}

/**
 * 落一条排序 op：核对现场 → 整块换成另一份镜像。
 *
 * 前进与撤销走同一个函数（逆操作就是 `before`/`after` 交换），所以这里没有"排序专属"的撤销分支：
 * 现场必须等于 `before` 那份镜像，对不上就整条跳过并报理由（`AGENTS.md` §2.5）。
 * 行上的 `spans` 描述的是"这一行自己的列分布"，整行搬走时它跟着记录一起走，不需要重算；
 * 已用区域也不变——排的是同一批行，非空行的下标集合前后一致。
 *
 * @param sheet 现场载荷。
 * @param op 排序 op。
 * @returns 新载荷与逆操作，或跳过理由。
 */
function applySortRowsOp(sheet: SheetPayload, op: SortRowsOp): ApplyOutcome {
  const from = resolveTrack(sheet, op.target);
  if (from === null) return { reason: "认不出要重排的第一行：它的内容已被改动，或它被插入/删除的行挪走了" };
  const shift = from - op.target.index;
  const range: A1Range = {
    start: { colIndex: op.range.start.colIndex, rowIndex: from },
    end: { colIndex: op.range.end.colIndex, rowIndex: from + (op.range.end.rowIndex - op.range.start.rowIndex) },
  };
  const guard = sortBandGuardReason(sheet, range);
  if (guard) return { reason: guard };
  const before = shiftBand(op.before, "row", shift);
  const after = shiftBand(op.after, "row", shift);
  const mismatch = rectMismatchReason(sheet, before, range);
  if (mismatch) return { reason: mismatch };

  const cells: Record<string, SheetCell> = {};
  for (const [key, cell] of Object.entries(sheet.cells)) {
    if (cell.rowIndex >= range.start.rowIndex && cell.rowIndex <= range.end.rowIndex && cell.colIndex >= range.start.colIndex && cell.colIndex <= range.end.colIndex) continue;
    cells[key] = cell;
  }
  for (const cell of after.cells) placeCell(cells, cell, cell.colIndex, cell.rowIndex);

  const inBand = (rowIndex: number) => rowIndex >= range.start.rowIndex && rowIndex <= range.end.rowIndex;
  const rows = [...sheet.rows.filter((row) => !inBand(row.rowIndex)), ...after.rows.map((row) => ({ ...row }))].sort((left, right) => left.rowIndex - right.rowIndex);
  const merges = sortMerges([
    ...sheet.merges.filter((merge) => !inBand(merge.range.start.rowIndex)),
    ...after.merges.map((merge) => ({ ref: merge.ref, range: { start: { ...merge.range.start }, end: { ...merge.range.end } } })),
  ]);
  const next: SheetPayload = { ...sheet, cells, rows, merges };
  return {
    // 逆操作的镜像已经在现场坐标系里（`shift` 搬过），锚点也换成落地后那一行的指纹
    sheet: next,
    inverse: { ...op, target: trackAddress(next, "row", from), before: after, after: before },
  };
}

/**
 * 现场这一矩形是否与镜像逐格一致（排序落地与撤销共用）。
 *
 * 与 `bandMismatchReason` 同一套判据，但范围收在矩形内：带外那一列不属于这次改动，
 * 用户在带外写东西不该把撤销整体拦死（区域外的格子由 `sortBandGuardReason` 单独拒绝）。
 *
 * @param sheet 现场载荷。
 * @param image 要核对的那份镜像（前进时是 `before`，撤销时是排好之后的那份）。
 * @param range 现场坐标系里的矩形。
 * @returns 不一致的理由；一致返回 null。
 */
function rectMismatchReason(sheet: SheetPayload, image: SheetTrackBand, range: A1Range): string | null {
  const inBand = (rowIndex: number) => rowIndex >= range.start.rowIndex && rowIndex <= range.end.rowIndex;
  const inRect = (colIndex: number) => colIndex >= range.start.colIndex && colIndex <= range.end.colIndex;
  const liveCells = Object.values(sheet.cells).filter((cell) => inBand(cell.rowIndex) && inRect(cell.colIndex));
  if (liveCells.length !== image.cells.length) return `第 ${range.start.rowIndex + 1}~${range.end.rowIndex + 1} 行的格子数与这条记录时不同，不覆盖它`;
  for (const recorded of image.cells) {
    const live = liveCells.find((cell) => cell.colIndex === recorded.colIndex && cell.rowIndex === recorded.rowIndex);
    if (!live || !sameCellContent(live, recorded)) return `${recorded.address} 现在装的不是这条记录时的内容，不覆盖它`;
  }
  const liveRows = sheet.rows.filter((row) => inBand(row.rowIndex));
  if (liveRows.length !== image.rows.length) return "这一带的行元信息与记录时不同，不动它";
  for (const recorded of image.rows) {
    const live = liveRows.find((row) => row.rowIndex === recorded.rowIndex);
    if (!live || !sameRowMeta(live, recorded)) return `第 ${recorded.rowIndex + 1} 行的行高/格式与记录时不同，不动它`;
  }
  const liveRefs = sheet.merges.filter((merge) => inBand(merge.range.start.rowIndex)).map((merge) => merge.ref).sort();
  const recordedRefs = image.merges.map((merge) => merge.ref).sort();
  if (liveRefs.length !== recordedRefs.length || !recordedRefs.every((ref, slot) => ref === liveRefs[slot])) return "这一带的合并区与记录时不同，不动它";
  return null;
}

/** 单条 op 的分派；指向别的表就整条跳过。 */
function applyOne(sheet: SheetPayload, op: SheetCoreOp): ApplyOutcome {
  if (op.partPath !== sheet.partPath) return { reason: "这条 op 指向的是另一张表" };
  if (op.kind === "setCellValue") return applyCellOp(sheet, op);
  if (op.kind === "setColumnWidth") return applyColumnWidthOp(sheet, op);
  if (op.kind === "sortRows") return applySortRowsOp(sheet, op);
  return applyShiftOp(sheet, op);
}

/**
 * 规划"受影响公式格的缓存刷新"（spec 决策 D8），产出的是**普通 `setCellValue` op**，
 * 所以它天然带着自己的逆操作：撤销时先回缓存再回数据，逐格还原到批前那份（C15/C17 的契约不破）。
 *
 * 只动"值确实变了"的公式格；算不出来的（白名单外 / volatile / 跨表 / 循环）**原样保留**——
 * 我们既证明不了它过期、也给不出更好的数字，抹掉会让用户在自己没碰过的格子里看到空白（与 §2.3 同一条口径）。
 *
 * @param sheet 本批用户 op 全部落地之后的载荷。
 * @param changedAddresses 这次改动碰到的格子（0 基）。几何变动（插删、排序）不走这里，见 D8 的边界说明。
 * @returns 待应用的 op 列表；没有公式格需要刷时为 `[]`。
 */
function planFormulaRefresh(sheet: SheetPayload, changedAddresses: A1Address[]): SheetCoreOp[] {
  const planned: SheetCoreOp[] = [];
  for (const [address, cell] of Object.entries(sheet.cells)) {
    if (!cell.formula || cell.rawXml) continue;
    const ranges = formulaReferencedRanges(cell.formula);
    // null = 引用里有我们不解释的结构（跨表、数组常量等），证明不了有没有被牵连，按 D8 不碰。
    if (!ranges || !ranges.some((range) => changedAddresses.some((changed) => rangeContains(range, changed)))) continue;
    const parsedAddress = parseA1(address);
    if (!parsedAddress) continue;
    const evaluated = evaluateFormula(cell.formula, sheet, address);
    if (!evaluated.ok) continue;
    const parts = formulaCellParts(evaluated.value);
    if (cell.value === evaluated.value && cell.kind === parts.kind) continue;
    planned.push({
      kind: "setCellValue",
      partPath: sheet.partPath,
      target: { row: trackAddress(sheet, "row", parsedAddress.rowIndex), colIndex: parsedAddress.colIndex },
      before: cell,
      after: { ...cell, ...parts, value: evaluated.value, cachedValue: evaluated.value },
    });
  }
  return planned;
}

/**
 * 顺序应用一批 op，失败的挑出去；`inverses` 与已落地的 `ops` 同序。
 *
 * 一批 op 都是照**同一份快照**锚定的，而第一条 op 本身就会改掉它那一行的内容指纹——
 * 不重新锚定的话，同一行的第二条 op 会以"认不出这一行"被跳过（阶段 4 的交叉读探针实测到的）。
 * 所以这里记一张"批内改过的条：旧指纹 → 新指纹"的表，只在这批 op 内部生效；
 * 批外的改动（用户后来的手工编辑）仍然按原样拦下来，`AGENTS.md` §2.5 的语义不受影响。
 *
 * @param sheet 现场载荷（入参不被改动）。
 * @param ops 待应用的一批 op。
 * @param isUndo 是否撤销批：撤销批落地后不再刷新一轮公式缓存（spec D8），正向批会在 op 全部落地后追加刷新。
 * @returns `SheetOpsResult`；正向批的 `ops`/`inverses` 可能比提交的条数多（多出的是缓存刷新 op）。
 */
function applyOps(sheet: SheetPayload, ops: SheetCoreOp[], isUndo: boolean): SheetOpsResult {
  let current = sheet;
  const opsApplied: SheetCoreOp[] = [];
  const inverses: SheetCoreOp[] = [];
  const skipped: SheetSkippedOp[] = [];
  const reanchors = new Map<string, SheetTrackAddress>();
  const changedCells: A1Address[] = [];
  for (const planned of ops) {
    const op = reanchorOp(planned, reanchors);
    const landedOn = trackLandedOn(current, op);
    const outcome = applyOne(current, op);
    if ("reason" in outcome) {
      skipped.push({ op: planned, reason: outcome.reason });
      continue;
    }
    if (landedOn) rememberReanchor(reanchors, outcome.sheet, op, landedOn);
    current = outcome.sheet;
    opsApplied.push(op);
    inverses.push(outcome.inverse);
    if (op.kind === "setCellValue") {
      const changed = parseA1(op.after?.address ?? op.before?.address ?? "");
      if (changed) changedCells.push(changed);
    }
  }
  // 正向批：把被牵连的公式缓存刷成真实值（D8）。撤销批不刷——撤销的契约是回到批前那个样子。
  if (!isUndo) {
    for (const refreshOp of planFormulaRefresh(current, changedCells)) {
      const outcome = applyOne(current, refreshOp);
      if ("reason" in outcome) continue;
      current = outcome.sheet;
      opsApplied.push(refreshOp);
      inverses.push(outcome.inverse);
    }
  }
  return { ok: skipped.length === 0, before: sheet, after: current, ops: opsApplied, inverses, skipped };
}

/** 一条"条"的身份证：轴 + 内容指纹 + 同指纹里的第几个。 */
function trackKey(track: SheetTrackAddress): string {
  return `${track.axis}|${track.fingerprint}|${track.nth}`;
}

/**
 * 把 op 的锚点换成"本批已经改过之后的那把指纹"。
 *
 * @param op 建 op 时锚定的那条（可能与现场对不上，因为同一行已被这批里的前一条改过）。
 * @param reanchors 批内锚点表。
 * @returns 能认出就换成新锚点的 op，否则原样返回（同一条对象，不复制）。
 */
function reanchorOp(op: SheetCoreOp, reanchors: Map<string, SheetTrackAddress>): SheetCoreOp {
  if (reanchors.size === 0) return op;
  if (op.kind === "insertTracks" || op.kind === "deleteTracks") return op;
  const track = op.kind === "setCellValue" ? op.target.row : op.target;
  const reanchored = reanchorTrack(track, reanchors);
  if (reanchored === track) return op;
  return op.kind === "setCellValue"
    ? { ...op, target: { ...op.target, row: reanchored } }
    : { ...op, target: reanchored };
}

/**
 * 沿锚点表往下走，走到现场认得出的那把指纹。
 *
 * @param track 起点锚点。
 * @param reanchors 批内锚点表。
 * @returns 最后一把锚点；没有记录时原样返回。
 */
function reanchorTrack(track: SheetTrackAddress, reanchors: Map<string, SheetTrackAddress>): SheetTrackAddress {
  let resolved = track;
  const seen = new Set<string>([trackKey(track)]);
  for (;;) {
    const next = reanchors.get(trackKey(resolved));
    if (!next || seen.has(trackKey(next))) return resolved;
    resolved = next;
    seen.add(trackKey(next));
  }
}

/**
 * 这条 op 现在落在第几条（按内容寻址在**改动前**的载荷上解）。
 *
 * @param sheet 改动前的载荷。
 * @param op 即将应用的 op。
 * @returns `{axis,index}`；插删与排序（改的是整段的位置，不是某一条的内容）与空白条（伪指纹，只按位置认）
 *   返回 null，表示这次改动不需要记批内锚点。
 */
function trackLandedOn(sheet: SheetPayload, op: SheetCoreOp): { axis: SheetAxis; index: number } | null {
  // 插删与排序都是"把记录整体挪走"，不是"把某一条改成新内容"：给它们记锚点会把后面的 op 领到错的行上，
  // 而按内容寻址本来就能在新区位认出那条记录（§2.5 要的就是这个）。
  if (op.kind === "insertTracks" || op.kind === "deleteTracks" || op.kind === "sortRows") return null;
  const track = op.kind === "setCellValue" ? op.target.row : op.target;
  if (track.fingerprint === EMPTY_TRACK) return null;
  const index = resolveTrack(sheet, track);
  return index === null ? null : { axis: track.axis, index };
}

/**
 * 记下"这一条被这次改动改成了什么指纹"，让同一批里后面的 op 还认得它。
 *
 * @param reanchors 批内锚点表（就地写入）。
 * @param after 改动后的载荷。
 * @param op 已落地的 op。
 * @param landedOn 这条 op 落在的第几条。
 */
function rememberReanchor(
  reanchors: Map<string, SheetTrackAddress>,
  after: SheetPayload,
  op: SheetCoreOp,
  landedOn: { axis: SheetAxis; index: number },
): void {
  const track = op.kind === "setCellValue" ? op.target.row : op.target;
  const fresh = trackAddress(after, landedOn.axis, landedOn.index);
  if (fresh.fingerprint !== track.fingerprint || fresh.nth !== track.nth) reanchors.set(trackKey(track), fresh);
}
