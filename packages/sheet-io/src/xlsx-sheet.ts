/**
 * 单张工作表（`xl/worksheets/sheetN.xml`）的读取与**定点改写**。
 *
 * 为什么不是"解析成对象再序列化整篇"：实测（spec §二 P5）`sheet1.xml` 里就带着
 * `<conditionalFormatting>`、`<dataValidations>`、`<pageMargins>`、`x14ac:dyDescent` 这类
 * 我们没有建模的结构，整篇重序列化会把它们抹掉。所以这里只替换 `<sheetData>` 区间与
 * `<dimension>`，其余字节原样保留（`AGENTS.md` §2.4 / spec §零.6）。
 */
import { formatA1, parseA1, parseRange } from "./address.js";
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
 * @param sheet 可能已被改过格子的表。
 * @param sharedStrings 工作簿共享的字符串表（新字符串会往里追加）。
 * @returns 完整的 `sheetN.xml` 部件文本：只有 `<sheetData>` 与 `<dimension>` 变了。
 */
export function serializeSheetXml(sheet: SheetModel, sharedStrings: SharedStringTable): string {
  const inner = buildSheetDataInner(sheet, sharedStrings);
  const replacement = inner ? `<sheetData>${inner}</sheetData>` : "<sheetData/>";
  let out = sheet.sourceXml.replace(SHEET_DATA_PATTERN, () => replacement);
  out = out.replace(DIMENSION_PATTERN, () => `<dimension ${dimensionAttributes(sheet)}/>`)
  return out;
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
  if (sheet.lastColIndex < 0 || sheet.lastRowIndex < 0) return 'ref="A1"';
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
      const reusable =
        cell.sharedStringIndex !== undefined &&
        !cell.runs?.length &&
        sharedStrings.at(cell.sharedStringIndex)?.text === String(cell.value ?? "")
          ? cell.sharedStringIndex
          : sharedStrings.intern(String(cell.value ?? ""), cell.runs);
      valueText = String(reusable);
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
