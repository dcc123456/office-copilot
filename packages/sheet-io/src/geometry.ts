/**
 * 网格几何的单位换算（spec 决策 D1：换算只在本包发生，渲染层只拿 px）。
 *
 * 一手依据：OOXML `@width` 的单位是"字符数"，而 SheetJS 的列属性文档把口径写死了
 * （`docs.sheetjs.com/docs/csf/features/colprops/`）——"总宽度里含 5 像素内边距"、
 * "字符宽度以 MDW 为单位、四舍五入到 1/256"。于是：
 * `px = chars × MDW + 5`，反向 `chars = round(((px − 5) / MDW) × 256) / 256`。
 * 用 Excel 默认值自检：`8.43 字符 → 8.4297 × 7 + 5 = 64 px`，与 Excel 默认列宽 64 px 一致。
 *
 * 行高走另一条：OOXML `@ht` 的单位是**磅**，磅→px 是 96/72（`15pt → 20px`，同样是 Excel 默认行高）。
 */

/** Calibri 11 @96DPI 的最大数字宽度（MDW），单位 px；换默认字体时由调用方传别的值。 */
export const DEFAULT_MAX_DIGIT_WIDTH_PX = 7;

/** Excel 给每格左右内边距与网格线留的固定像素，字符宽度里**不含**这一段。 */
export const COLUMN_PADDING_PX = 5;

/** 未设 `@ht` 的行用这个默认行高（磅），与 Excel 空白工作簿一致。 */
export const DEFAULT_SHEET_ROW_HEIGHT_PT = 15;

/** 磅到 CSS px 的系数（96DPI）。 */
const PX_PER_PT = 96 / 72;

/**
 * 列宽（字符）→ CSS 像素。
 *
 * @param widthChars OOXML `@width`，单位是"最大数字字符个数"，不是 px 也不是 pt。
 * @param mdwPx 当前默认字体的最大数字宽度，默认 7px（Calibri 11 @96DPI）。
 * @returns 整数像素；传 NaN 时返回 NaN，调用方按"这列没设宽度"处理。
 */
export function columnWidthCharsToPx(widthChars: number, mdwPx = DEFAULT_MAX_DIGIT_WIDTH_PX): number {
  return Math.round(widthChars * mdwPx + COLUMN_PADDING_PX);
}

/**
 * CSS 像素 → 列宽（字符），拖列宽时写回 `@width` 用。
 *
 * @param widthPx 用户拖出来的像素宽度。
 * @param mdwPx 同 `columnWidthCharsToPx`。
 * @returns 按 1/256 字符对齐后的宽度；`pxToCharsToPx` 与传入值一致，拖动不会逐次漂移。
 */
export function columnWidthPxToChars(widthPx: number, mdwPx = DEFAULT_MAX_DIGIT_WIDTH_PX): number {
  return Math.round(((widthPx - COLUMN_PADDING_PX) / mdwPx) * 256) / 256;
}

/**
 * 行高（磅）→ CSS 像素。
 *
 * @param heightPt OOXML `@ht`，单位是磅（与 docx 那边同源，但换算只在此处发生）。
 * @returns 整数像素；`15 → 20`。
 */
export function rowHeightPtToPx(heightPt: number): number {
  return Math.round(heightPt * PX_PER_PT);
}
