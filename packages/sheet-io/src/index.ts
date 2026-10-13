/**
 * `@office-copilot/sheet-io` 的公开入口（spec D1）。
 *
 * 这里是渲染层 / CLI / AI 工具层能碰到的**唯一**表格 IO 面：
 * 单位（日期序列号、pt、字符宽度、ARGB）的解释只在本包内部发生，外面只拿到结构化模型。
 */
export { SheetFile, SheetOpenError, type SheetOpenOptions, type SheetRef, type SheetSnapshot } from "./sheet-file.js";
export {
  cellAt,
  isSheetEmpty,
  type CellRun,
  type CellScalarValue,
  type CellKind,
  type ResolvedNumberFormat,
  type SheetCell,
  type SheetColumn,
  type SheetFreeze,
  type SheetMerge,
  type SheetModel,
  type SheetRowMeta,
  type SharedStringEntry,
  type UnsupportedFeature,
  type WorkbookModel,
} from "./model.js";
export {
  colNameToIndex,
  formatA1,
  formatRange,
  indexToColName,
  parseA1,
  parseRange,
  rangeContains,
  type A1Address,
  type A1Range,
} from "./address.js";
export {
  BUILTIN_DATE_TIME_NUMFMT_IDS,
  isDateFormatCode,
  serialToUtcDate,
  utcDateToSerial,
} from "./datetime.js";
export { argbToSheetColor, sheetColorToArgb, type FontSpec, type SheetColor } from "./color.js";
export {
  COLUMN_PADDING_PX,
  DEFAULT_MAX_DIGIT_WIDTH_PX,
  DEFAULT_SHEET_ROW_HEIGHT_PT,
  columnWidthCharsToPx,
  columnWidthPxToChars,
  rowHeightPtToPx,
} from "./geometry.js";
export { colFingerprints, parseSheet, recalculateSheetExtent, rowFingerprints, serializeSheetXml, type ParseSheetInput } from "./xlsx-sheet.js";
export { SharedStringTable } from "./xlsx-strings.js";
export { StyleTable, type CellStyleParts } from "./xlsx-styles.js";
export { cellDisplayText, formatNumberForDisplay, parseUserCellText, type ParsedCellInput } from "./cell-text.js";
export {
  applySheetPayload,
  EDITABLE_CELL_LIMIT,
  readOnlyReason,
  sheetPayloadOf,
  usedCellCount,
  workbookPayload,
  type SheetPayload,
  type WorkbookPayload,
} from "./sheet-dto.js";
export {
  deleteTracksOp,
  insertTracksOp,
  revertSheetOps,
  runSheetOps,
  setCellFormulaOp,
  setCellTextOp,
  setColumnWidthOp,
  type SheetCoreOp,
  type SheetOpsResult,
} from "./sheet-edit.js";
export {
  aggregateSheet,
  describeSheet,
  duplicateGroups,
  resolveColumnRef,
  type AggregateFunction,
  type AggregateGroup,
  type AggregateInput,
  type AggregateResult,
  type AggregateSpec,
  type DescribeOptions,
  type DuplicateGroup,
  type DuplicateResult,
  type SheetProfile,
} from "./analysis/sheet-table.js";
export {
  type ColumnOutlier,
  type ColumnProfile,
  type ColumnValueType,
  type NumericSummary,
} from "./analysis/column-profile.js";
