/**
 * 单张工作表（`xl/worksheets/sheetN.xml`）的读取与**定点改写**。
 *
 * 为什么不是"解析成对象再序列化整篇"：实测（spec §二 P5）`sheet1.xml` 里就带着
 * `<conditionalFormatting>`、`<dataValidations>`、`<pageMargins>`、`x14ac:dyDescent` 这类
 * 我们没有建模的结构，整篇重序列化会把它们抹掉。所以这里只替换 `<sheetData>` 区间与
 * `<dimension>`，其余字节原样保留（`AGENTS.md` §2.4 / spec §零.6）。
 */
import { formatA1, parseA1, parseRange } from "./address.js";
import { isSheetEmpty } from "./model.js";
import type {
  CellScalarValue,
  SheetCell,
  SheetColumn,
  SheetLike,
  SheetMerge,
  SheetModel,
  SheetRowMeta,
} from "./model.js";
import type { StyleTable } from "./xlsx-styles.js";
import type { SharedStringTable } from "./xlsx-strings.js";
import { type Element, children, escapeAttr, escapeText, firstChild, isElement, parseDocument, serialize, text } from "./xml.js";

/** 解析一张表的输入；字符串表与样式表由工作簿共享，所以传进来而不是重新解析。 */
export interface ParseSheetInput {
  xmlText: string;
  partPath: string;
  name: string;
  sheetId: number;
  sharedStrings: SharedStringTable;
  styles: StyleTable;
}

/**
 * 解析一张工作表。
 *
 * @param input 部件文本 + 共享表 + 表元信息。
 * @returns `SheetModel`，含 XML 原文（写回时用它做定点替换）。读不懂的构造记进 `warnings`，
 *   并把原片段挂在格上（`rawXml`），保证写回时不丢。
 */
export function parseSheet(input: ParseSheetInput): SheetModel {
  const root = parseDocument(input.xmlText);
  const warnings: string[] = [];
  const cells: Record<string, SheetCell> = {};
  const rows: SheetRowMeta[] = [];
  const columns: SheetColumn[] = [];
  const merges: SheetMerge[] = [];
  let lastColIndex = -1;
  let lastRowIndex = -1;

  const sheetData = firstChild(root, "sheetData");
  if (!sheetData) warnings.push("missing-sheetData");
  for (const rowEl of sheetData ? findRows(sheetData) : []) {
    const rowIndex = Number(rowEl.attributes["r"]) - 1;
    if (!Number.isInteger(rowIndex) || rowIndex < 0) {
      warnings.push("row-without-index");
      continue;
    }
    const meta: SheetRowMeta = { rowIndex, attrs: restAttributes(rowEl, ["r"]) };
    const height = rowEl.attributes["ht"];
    if (height !== undefined) meta.heightPt = Number(height);
    if (rowEl.attributes["customHeight"] === "1") meta.customHeight = true;
    if (rowEl.attributes["hidden"] === "1") meta.hidden = true;
    const rowStyle = rowEl.attributes["s"];
    if (rowStyle !== undefined) meta.styleIndex = Number(rowStyle);
    rows.push(meta);
    lastRowIndex = Math.max(lastRowIndex, rowIndex);

    for (const cellEl of children(rowEl).filter((n) => isElement(n) && n.name === "c")) {
      const cell = parseCell(cellEl, input, warnings);
      if (!cell) continue;
      cells[cell.address] = cell;
      lastColIndex = Math.max(lastColIndex, cell.colIndex);
    }
  }

  for (const colEl of findChildrenByName(firstChild(root, "cols"), "col")) {
    const min = Number(colEl.attributes["min"]) - 1;
    const max = Number(colEl.attributes["max"]) - 1;
    if (!Number.isInteger(min) || !Number.isInteger(max) || min < 0) continue;
    const column: SheetColumn = { minColIndex: min, maxColIndex: max, attrs: restAttributes(colEl, []) };
    const width = colEl.attributes["width"];
    if (width !== undefined) column.widthChars = Number(width);
    if (colEl.attributes["hidden"] === "1") column.hidden = true;
    if (colEl.attributes["customWidth"] === "1") column.customWidth = true;
    columns.push(column);
  }

  for (const mergeEl of findChildrenByName(firstChild(root, "mergeCells"), "mergeCell")) {
    const ref = String(mergeEl.attributes["ref"] ?? "");
    const range = parseRange(ref);
    if (!range) continue;
    merges.push({ ref: ref.toUpperCase(), range });
    lastColIndex = Math.max(lastColIndex, range.end.colIndex);
    lastRowIndex = Math.max(lastRowIndex, range.end.rowIndex);
  }

  const sheetViews = firstChild(root, "sheetViews");
  const sheetView = sheetViews ? firstChild(sheetViews, "sheetView") : undefined;
  const pane = sheetView ? firstChild(sheetView, "pane") : undefined;
  let freeze: SheetModel["freeze"];
  if (pane && pane.attributes["state"] === "frozen") {
    freeze = { colCount: Number(pane.attributes["xSplit"] ?? 0), rowCount: Number(pane.attributes["ySplit"] ?? 0) };
  }

  for (const el of children(root)) {
    if (!MODELLED_ELEMENTS.has(el.name)) warnings.push(`unmodelled:${el.name}`);
  }

  return {
    name: input.name,
    sheetId: input.sheetId,
    partPath: input.partPath,
    lastColIndex,
    lastRowIndex,
    cells,
    merges,
    columns,
    rows,
    freeze,
    sourceXml: input.xmlText,
    warnings,
  };
}

/**
 * 生成一张表的新 XML。
 *
 * @param sheet 可能已被改过格子、列宽或合并区的表。
 * @param sharedStrings 工作簿共享的字符串表（新字符串会往里追加）。
 * @returns 完整的 `sheetN.xml` 部件文本：只有 `<sheetData>`、`<dimension>`、`<cols>` 与 `<mergeCells>` 变了。
 */
export function serializeSheetXml(sheet: SheetModel, sharedStrings: SharedStringTable): string {
  const inner = buildSheetDataInner(sheet, sharedStrings);
  const replacement = inner ? `<sheetData>${inner}</sheetData>` : "<sheetData/>";
  let out = sheet.sourceXml.replace(SHEET_DATA_PATTERN, () => replacement);
  out = out.replace(DIMENSION_PATTERN, () => `<dimension ${dimensionAttributes(sheet)}/>`);
  out = replaceOrInsertCols(out, sheet.columns);
  return replaceOrInsertMergeCells(out, sheet.merges);
}

/** `<mergeCells>` 区间；与 `<sheetData>` 一样按原文定点替换，不许整篇重序列化。 */
const MERGE_CELLS_PATTERN = /<mergeCells(?:\s[^>]*)?\s*\/>|<mergeCells(?:\s[^>]*)?>[\s\S]*?<\/mergeCells>/;

/**
 * CT_Worksheet 里排在 `<mergeCells>` **之后**的元素（按规范顺序，缺哪个跳哪个）。
 *
 * 表原来一个合并区都没有时包里根本没有这一段，得凭这些邻居把它插在正确位置；顺序错了 Excel 会报文件损坏。
 */
const AFTER_MERGE_CELLS_PATTERN = /<(?:conditionalFormatting|dataValidations|hyperlinks|printOptions|pageMargins|pageSetup|headerFooter|rowBreaks|colBreaks|drawing|legacyDrawing|picture|oleObjects|controls|tableParts|smartTags|cellWatches|ignoredErrors|customProperties|phoneticPr)\b/;

/**
 * 把合并区写回 `<mergeCells>` 区间。
 *
 * 为什么必须动它：插删行列会把合并区整体挪位（实测 `H1:I2` 插一列后应是 `I1:J2`），
 * 只写 `<sheetData>` 的话文件里的合并区还停在旧位置，读回来与内存模型不一致。
 *
 * @param xml 已替换过 `<sheetData>`、`<dimension>` 与 `<cols>` 的部件文本。
 * @param merges 表当前的合并区（`sheet-edit.ts` 已按左上角排好）。
 * @returns 带上新 `<mergeCells>` 的文本；一个合并区都没有时删掉原区间（空区间 Excel 也不写）。
 */
function replaceOrInsertMergeCells(xml: string, merges: SheetMerge[]): string {
  const region = merges.length
    ? `<mergeCells count="${merges.length}">${merges.map((merge) => `<mergeCell ref="${escapeAttr(merge.ref)}"/>`).join("")}</mergeCells>`
    : "";
  if (MERGE_CELLS_PATTERN.test(xml)) return xml.replace(MERGE_CELLS_PATTERN, () => region);
  if (!region) return xml;
  const anchor = AFTER_MERGE_CELLS_PATTERN.exec(xml);
  if (anchor?.index !== undefined) return xml.slice(0, anchor.index) + region + xml.slice(anchor.index);
  return xml.replace(/<\/worksheet>/, () => `${region}</worksheet>`);
}

/** `<cols>` 区间；与 `<sheetData>` 一样按原文定点替换，不许整篇重序列化。 */
const COLS_PATTERN = /<cols(?:\s[^>]*)?\s*\/>|<cols(?:\s[^>]*)?>[\s\S]*?<\/cols>/;

/**
 * 把列宽定义写回 `<cols>` 区间，其余字节仍然原样带着。
 *
 * 为什么必须动它：`<col min max>` 是一段范围，拖一列的宽度会把它劈成三段，
 * 只写 `<sheetData>` 的话用户的列宽根本存不进文件（实测：写出去读回来还是原来那两段）。
 *
 * @param xml 已替换过 `<sheetData>` 与 `<dimension>` 的部件文本。
 * @param columns 表当前的列宽分段（`sheet-edit.ts` 已按列号排好）。
 * @returns 带上新 `<cols>` 的文本；一段自定义宽度都没有时删掉原区间（Excel 也是这么写的）。
 */
function replaceOrInsertCols(xml: string, columns: SheetColumn[]): string {
  const region = columns.length ? `<cols>${columns.map(colXml).join("")}</cols>` : "";
  if (COLS_PATTERN.test(xml)) return xml.replace(COLS_PATTERN, () => region);
  // 原来没有 `<cols>`（这张表从没设过自定义宽度）：按 CT_Worksheet 的顺序插在 `<sheetData>` 之前
  return xml.replace(SHEET_DATA_PATTERN, (sheetData) => region + sheetData);
}

/**
 * 序列化一段 `<col>`。
 *
 * `attrs` 存的是读入时的全部属性（含 `min`/`max`/`width`/`customWidth`/`hidden`），
 * 这几个一律以模型字段为准重写——否则劈段后会写出与 `minColIndex` 矛盾的旧值；
 * 其余属性（`style`、`bestFit`、`x14ac:width`…）原样搬。
 *
 * @param column 一个列宽分段。
 * @returns `<col .../>` 自闭合标签（`min`/`max` 是 1 基，这里 +1 换算）。
 */
function colXml(column: SheetColumn): string {
  const attrs: string[] = [`min="${column.minColIndex + 1}"`, `max="${column.maxColIndex + 1}"`];
  if (column.widthChars !== undefined) attrs.push(`width="${column.widthChars}"`);
  for (const [key, value] of Object.entries(column.attrs ?? {})) {
    if (["min", "max", "width", "hidden", "customWidth"].includes(key)) continue;
    attrs.push(`${key}="${escapeAttr(value)}"`);
  }
  if (column.customWidth) attrs.push('customWidth="1"');
  if (column.hidden) attrs.push('hidden="1"');
  return `<col ${attrs.join(" ")}/>`;
}

/**
 * 表内所有格子的值指纹（阶段 3 的内容寻址撤销要用，spec D3）。
 *
 * @param sheet 目标表（载荷也行，只读 `cells`）。
 * @returns 行号 → 该行内容指纹。串里带的是**列号**而不是 A1 地址：插入一行会把整表下移，
 *   地址跟着变，用地址当身份就等于承认"位移之后的行不是原来那行"，撤销会认错。
 */
export function rowFingerprints(sheet: SheetLike): Record<number, string> {
  const byRow = groupCellsByRow(sheet);
  const acc: Record<number, string> = {};
  for (const [rowIndex, rowCells] of byRow) {
    acc[rowIndex] = rowCells.map((cell) => `${cell.colIndex}=${cell.value ?? ""}`).join("|");
  }
  return acc;
}

/**
 * 列方向的内容指纹（与 `rowFingerprints` 对称，供列宽与列插删的撤销定位）。
 *
 * @param sheet 目标表（载荷也行，只读 `cells`）。
 * @returns 列号 → 该列内容指纹，串里带行号，所以插入列不会让指纹认错。
 */
export function colFingerprints(sheet: SheetLike): Record<number, string> {
  const byCol = new Map<number, SheetCell[]>();
  for (const cell of Object.values(sheet.cells)) {
    const bucket = byCol.get(cell.colIndex);
    if (bucket) bucket.push(cell);
    else byCol.set(cell.colIndex, [cell]);
  }
  const acc: Record<number, string> = {};
  for (const [colIndex, colCells] of byCol) {
    acc[colIndex] = colCells
      .slice()
      .sort((a, b) => a.rowIndex - b.rowIndex)
      .map((cell) => `${cell.rowIndex}=${cell.value ?? ""}`)
      .join("|");
  }
  return acc;
}

/** `<dimension/@ref>` 需要的写法：由实际存在的格子推出，不猜 Excel 没写的区域。 */
function dimensionAttributes(sheet: SheetModel): string {
  if (isSheetEmpty(sheet)) return 'ref="A1"';
  const far = formatA1({ colIndex: sheet.lastColIndex, rowIndex: sheet.lastRowIndex });
  return `ref="A1${far ? `:${far}` : ""}"`;
}

const MODELLED_ELEMENTS = new Set([
  "sheetPr",
  "dimension",
  "sheetViews",
  "sheetFormatPr",
  "cols",
  "sheetData",
  "mergeCells",
]);

const SHEET_DATA_PATTERN = /<sheetData(?:\s[^>]*)?\s*\/>|<sheetData(?:\s[^>]*)?>[\s\S]*?<\/sheetData>/;
const DIMENSION_PATTERN = /<dimension(?:\s[^>]*)?\/>/;

/**
 * 重算已用区域右下角。
 *
 * @param sheet 原地更新的表（调用方改完 `cells` 后必须调它，否则 `<dimension>` 与实际格子不一致）。
 */
export function recalculateSheetExtent(sheet: SheetLike): void {
  let lastColIndex = -1;
  let lastRowIndex = -1;
  for (const cell of Object.values(sheet.cells)) {
    lastColIndex = Math.max(lastColIndex, cell.colIndex);
    lastRowIndex = Math.max(lastRowIndex, cell.rowIndex);
  }
  for (const merge of sheet.merges) {
    lastColIndex = Math.max(lastColIndex, merge.range.end.colIndex);
    lastRowIndex = Math.max(lastRowIndex, merge.range.end.rowIndex);
  }
  sheet.lastColIndex = lastColIndex;
  sheet.lastRowIndex = lastRowIndex;
}

function findRows(sheetData: Element): Element[] {
  return children(sheetData).filter((n) => isElement(n) && n.name === "row");
}

function findChildrenByName(el: Element | undefined, name: string): Element[] {
  return el ? children(el).filter((n) => isElement(n) && n.name === name) : [];
}

function restAttributes(el: Element, exclude: string[]): Record<string, string> {
  const acc: Record<string, string> = {};
  for (const [key, value] of Object.entries(el.attributes ?? {})) {
    if (exclude.includes(key)) continue;
    acc[key] = String(value);
  }
  return acc;
}

/**
 * 解析一个 `<c>`。
 *
 * @param cellEl 单元格元素。
 * @param input 解析上下文（共享表、样式表、warning 收集器）。
 * @returns 单元格；完全无法定位（缺 `@r`）时返回 undefined 并记 warning。
 */
function parseCell(
  cellEl: Element,
  input: ParseSheetInput,
  warnings: string[],
): SheetCell | undefined {
  const rawRef = String(cellEl.attributes["r"] ?? "");
  const address = rawRef.toUpperCase();
  const position = parseA1(address);
  if (!position) {
    warnings.push(`cell-without-address:${rawRef || "(none)"}`);
    return undefined;
  }
  const cell: SheetCell = {
    address,
    colIndex: position.colIndex,
    rowIndex: position.rowIndex,
    kind: "blank",
    value: null,
  };
  const extraAttrs = restAttributes(cellEl, ["r", "s", "t"]);
  if (Object.keys(extraAttrs).length > 0) cell.attrs = extraAttrs;
  const styleIndexRaw = cellEl.attributes["s"];
  if (styleIndexRaw !== undefined) cell.styleIndex = Number(styleIndexRaw);

  const valueEl = firstChild(cellEl, "v");
  const formulaEl = firstChild(cellEl, "f");
  const rawValue = valueEl ? text(valueEl) : undefined;
  const typeAttr = String(cellEl.attributes["t"] ?? "");

  switch (typeAttr) {
    case "":
    case "n": {
      if (rawValue === undefined) break;
      const numeric = Number(rawValue);
      if (Number.isFinite(numeric)) {
        cell.kind = "number";
        cell.value = numeric;
      } else {
        warnings.push(`unreadable-number:${address}=${rawValue}`);
        cell.rawXml = serialize(cellEl);
      }
      break;
    }
    case "s": {
      const index = Number(rawValue);
      const entry = Number.isInteger(index) ? input.sharedStrings.at(index) : undefined;
      if (!entry) {
        warnings.push(`bad-string-reference:${address}→${rawValue}`);
        cell.rawXml = serialize(cellEl);
      } else {
        cell.kind = "string";
        cell.value = entry.text;
        cell.sharedStringIndex = index;
        if (entry.runs) cell.runs = entry.runs;
      }
      break;
    }
    case "inlineStr": {
      const isEl = firstChild(cellEl, "is");
      cell.kind = "string";
      cell.value = isEl ? text(isEl) : "";
      break;
    }
    case "b":
      cell.kind = "boolean";
      cell.value = rawValue === "1";
      break;
    case "e":
      cell.kind = "error";
      cell.value = rawValue ?? "";
      break;
    case "str":
      cell.kind = "string";
      cell.value = rawValue ?? "";
      break;
    default:
      warnings.push(`unknown-cell-type:${typeAttr}`);
      cell.rawXml = serialize(cellEl);
  }

  if (formulaEl && !cell.rawXml) {
    cell.formula = text(formulaEl);
    cell.cachedValue = cell.value;
  }
  if (cell.kind === "blank" && cell.styleIndex !== undefined) cell.styleOnly = true;
  // 日期只在样式里（实测 §二 P3）：kind 保持 number，上层靠 numberFormat.isDateLike 判定显示。
  if (cell.styleIndex !== undefined) cell.numberFormat = input.styles.numberFormatFor(cell.styleIndex);
  return cell;
}

function groupCellsByRow(sheet: SheetLike): Map<number, SheetCell[]> {
  const byRow = new Map<number, SheetCell[]>();
  for (const cell of Object.values(sheet.cells)) {
    const bucket = byRow.get(cell.rowIndex);
    if (bucket) bucket.push(cell);
    else byRow.set(cell.rowIndex, [cell]);
  }
  for (const bucket of byRow.values()) bucket.sort((a, b) => a.colIndex - b.colIndex);
  return byRow;
}

/**
 * 生成 `<sheetData>` 的内部 XML。
 *
 * @param sheet 表模型。
 * @param sharedStrings 共享字符串表（新字符串在这里追加）。
 * @returns 行按 `r` 升序、格按列升序的串；空表返回空串。
 */
function buildSheetDataInner(sheet: SheetModel, sharedStrings: SharedStringTable): string {
  const byRow = groupCellsByRow(sheet);
  const rowIndexes = new Set<number>(byRow.keys());
  for (const meta of sheet.rows) rowIndexes.add(meta.rowIndex);
  const metas = new Map(sheet.rows.map((meta) => [meta.rowIndex, meta]));

  const parts: string[] = [];
  for (const rowIndex of [...rowIndexes].sort((a, b) => a - b)) {
    const meta = metas.get(rowIndex);
    const attrs: string[] = [`r="${rowIndex + 1}"`];
    if (meta?.heightPt !== undefined) attrs.push(`ht="${meta.heightPt}"`);
    if (meta?.customHeight) attrs.push('customHeight="1"');
    if (meta?.hidden) attrs.push('hidden="1"');
    if (meta?.styleIndex !== undefined) attrs.push(`s="${meta.styleIndex}"`);
    for (const [key, value] of Object.entries(meta?.attrs ?? {})) {
      if (["ht", "customHeight", "hidden", "s"].includes(key)) continue;
      attrs.push(`${key}="${escapeAttr(value)}"`);
    }
    const cellsXml = (byRow.get(rowIndex) ?? []).map((cell) => cellXml(cell, sharedStrings)).join("");
    parts.push(cellsXml ? `<row ${attrs.join(" ")}>${cellsXml}</row>` : `<row ${attrs.join(" ")}/>`);
  }
  return parts.join("");
}

function cellXml(cell: SheetCell, sharedStrings: SharedStringTable): string {
  if (cell.rawXml) return cell.rawXml;
  const attrs: string[] = [`r="${cell.address}"`];
  if (cell.styleIndex !== undefined) attrs.push(`s="${cell.styleIndex}"`);
  for (const [key, value] of Object.entries(cell.attrs ?? {})) {
    if (key === "t") continue;
    attrs.push(`${key}="${escapeAttr(value)}"`);
  }

  const inner: string[] = [];
  let typeAttr = "";
  let valueText: string | undefined;

  if (cell.kind === "string") {
    if (cell.formula) {
      typeAttr = "str";
      valueText = String(cell.value ?? "");
    } else {
      typeAttr = "s";
      // 内容没变就沿用读入时那条 `<si>`：换索引等于换条目，格式只剩我们能重新序列化出来的那部分。
      valueText = String(sharedStrings.reuseOrAppend(String(cell.value ?? ""), cell.runs, cell.sharedStringIndex));
    }
  } else if (cell.kind === "boolean") {
    typeAttr = "b";
    valueText = cell.value ? "1" : "0";
  } else if (cell.kind === "error") {
    typeAttr = "e";
    valueText = String(cell.value ?? "");
  } else if (cell.kind === "number") {
    valueText = formatNumeric(cell.value);
  }

  if (cell.formula) inner.push(`<f>${escapeText(cell.formula)}</f>`);
  if (valueText !== undefined && !cell.styleOnly) inner.push(`<v>${escapeText(valueText)}</v>`);
  if (typeAttr) attrs.push(`t="${typeAttr}"`);

  const head = attrs.join(" ");
  return inner.length ? `<c ${head}>${inner.join("")}</c>` : `<c ${head}/>`;
}

/**
 * 数字写回文本。
 *
 * @param value 单元格数值。
 * @returns 定点写法；不是有限数时返回 undefined，调用方据此不写 `<v>`（不许把 NaN 塞进 OOXML）。
 */
function formatNumeric(value: CellScalarValue): string | undefined {
  if (typeof value !== "number" || !Number.isFinite(value)) return undefined;
  return Number.isInteger(value) ? String(value) : String(Number(value.toPrecision(15)));
}
