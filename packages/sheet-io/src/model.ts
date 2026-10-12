/**
 * 工作簿的数据模型（spec D1）。
 *
 * 刻意**不套用 ProseMirror 树**：工作簿是稀疏网格 + 字符串表 + 样式表，
 * 与 docx 的节点树没有结构共享点。单位语义写在字段名上（`…Pt`/`…Px`/`…Chars`/`serial`）。
 */
import type { A1Range } from "./address.js";
import type { FontSpec } from "./color.js";

/** 单元格值的原生形态；日期是 `number`（序列号），由 `numberFormat` 决定显示。 */
export type CellScalarValue = string | number | boolean | null;

/** 对应 `<c/@t>`：`s`=共享字符串、`inlineStr`、`b`、`e`、缺省=数字。`blank` 是只有格式的格。 */
export type CellKind = "string" | "number" | "boolean" | "error" | "blank";

/** 富文本 run（实测来自 `<si><r><rPr>…`，spec §二 P2）。只装我们读得懂的标记，其余保留原串。 */
export interface CellRun extends FontSpec {
  text: string;
}

/** 由 `cellXfs` + `numFmts` 解出来的数字格式。 */
export interface ResolvedNumberFormat {
  numFmtId: number;
  /** 自定义格式码；内建 id 时为 undefined（Excel 自己知道怎么显示）。 */
  formatCode?: string;
  isDateLike: boolean;
}

/** 一个非空或有格式的格。`address` 与 `colIndex/rowIndex` 必须同步，写回时以两者一致为前提。 */
export interface SheetCell {
  address: string;
  colIndex: number;
  rowIndex: number;
  kind: CellKind;
  value: CellScalarValue;
  /** 公式体，不含前导 `=`；实测形态 `<c r="G2"><f>SUM(C2:C4)</f><v>360</v></c>`。 */
  formula?: string;
  /** `<v>` 里的缓存值（读入时与 `value` 同值；写入公式时必须给，见 spec D4）。 */
  cachedValue?: CellScalarValue;
  /** 原 `@s`（cellXfs 位置索引）。写回时优先带着它，避免无谓地新建样式。 */
  styleIndex?: number;
  numberFormat?: ResolvedNumberFormat;
  runs?: CellRun[];
  /** `t="s"` 时读到的 string 表索引；值未变则写回复用它，不追加重复条目。 */
  sharedStringIndex?: number;
  /** 只有格式没有值时的原样 XML（`<c r="H2" s="3"/>`），保证往返不丢格。 */
  styleOnly?: boolean;
  /** `<c>` 上除 `r`/`s`/`t` 之外我们不解释的属性（`vm`、`cm`…），写回时逐条原样搬。 */
  attrs?: Record<string, string>;
  /** 整格原文；遇到无法解释的结构时用它原样回写，不许丢。 */
  rawXml?: string;
}

/** 列宽：OOXML `@width` 的单位是"字符数"，不是 px 也不是 pt，所以原样存不做换算。 */
export interface SheetColumn {
  minColIndex: number;
  maxColIndex: number;
  widthChars?: number;
  hidden?: boolean;
  customWidth?: boolean;
  /** `<col>` 上其余属性（`style`、`bestFit`…），写回时原样搬。 */
  attrs?: Record<string, string>;
}

/** 行元信息（`<row>` 上除单元格之外的属性）。 */
export interface SheetRowMeta {
  rowIndex: number;
  heightPt?: number;
  customHeight?: boolean;
  hidden?: boolean;
  styleIndex?: number;
  /** `<row>` 上其余属性（`spans`、`x14ac:dyDescent`、`thickBot`…）。实测每张表都带 `x14ac:dyDescent`，丢了就是丢用户格式。 */
  attrs?: Record<string, string>;
}

/** 合并区域（`<mergeCell/@ref>`）。 */
export interface SheetMerge {
  ref: string;
  range: A1Range;
}

/** 冻结窗格：`<pane xSplit ySplit>` 换算成"冻结了几列几行"。 */
export interface SheetFreeze {
  colCount: number;
  rowCount: number;
}

/** 一张工作表的完整建模结果，外加它的原始 XML（定点改写用）。 */
export interface SheetModel {
  name: string;
  sheetId: number;
  /** 包内部件路径，如 `xl/worksheets/sheet1.xml`。 */
  partPath: string;
  /** 已用区域右下角（0 基）。空表为 `-1`。 */
  lastColIndex: number;
  lastRowIndex: number;
  cells: Record<string, SheetCell>;
  merges: SheetMerge[];
  columns: SheetColumn[];
  rows: SheetRowMeta[];
  freeze?: SheetFreeze;
  /** 该表的 XML 原文；`writeSheet` 只替换 `<sheetData>` 区间（spec §零.6）。 */
  sourceXml: string;
  warnings: string[];
}

/** `xl/sharedStrings.xml` 的一个 `<si>` 条目。 */
export interface SharedStringEntry {
  text: string;
  runs?: CellRun[];
  /** `<si>` 的序列化原文（含 `<rPh>` 注音等我们不解释的子元素），未改动时按它原样回写。 */
  rawSi?: string;
}

/** 打开时发现的"我们没有建模但它还在文件里"的东西，只用于提示，不参与改写。 */
export interface UnsupportedFeature {
  partPath: string;
  reason: string;
}

/** 整个工作簿的建模结果。 */
export interface WorkbookModel {
  fileName: string;
  sheets: SheetModel[];
  sharedStrings: SharedStringEntry[];
  /** 表级警告（读不懂的属性、跳过的部件等），必须一路带到 UI，不许静默。 */
  warnings: string[];
  unsupported: UnsupportedFeature[];
}

/** 空表判定：`lastColIndex < 0` 即没有任何已用格。 */
export function isSheetEmpty(sheet: SheetModel): boolean {
  return sheet.lastColIndex < 0 || sheet.lastRowIndex < 0;
}

/**
 * 按 A1 地址取格。
 *
 * @param sheet 目标表。
 * @param address 大写 A1 串（调用方负责已归一化）。
 * @returns 单元格；不存在返回 undefined（含义是"这格没有内容也没有格式"）。
 */
export function cellAt(sheet: SheetModel, address: string): SheetCell | undefined {
  return sheet.cells[address];
}
