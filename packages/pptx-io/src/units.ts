/**
 * DrawingML 的单位换算：EMU ↔ CSS px、半磅 ↔ pt。全仓唯一的一处（`AGENTS.md` §2.1 / spec D4）。
 *
 * 读模型交给 UI 的数字一律是 px 与 pt，写侧才换回 EMU 与百分磅——渲染层、CLI、AI 工具层再出现一次
 * `9525` 或 `×100` 就是缺陷。
 */

/** 形状的位置/尺寸（CSS px @96DPI），四个值都给齐——读模型交给视图的就是这一份。 */
export interface ShapeBoxPx {
  /** 左偏移（px，可为负表示挪出幻灯片）。 */
  leftPx: number;
  /** 上偏移（px，可为负）。 */
  topPx: number;
  /** 宽（px，须为正）。 */
  widthPx: number;
  /** 高（px，须为正）。 */
  heightPx: number;
}

/** 要落地的几何改动：只带用户真改过的那几项（`setShapeGeometry` 的入参）。 */
export type ShapeGeometryPatch = Partial<ShapeBoxPx>;

/** CSS px @96DPI → EMU 的系数（914400 EMU/inch ÷ 96 px/inch）。 */
const EMU_PER_PX = 9525;

/**
 * EMU → CSS px。
 *
 * @param emu OOXML 里的长度（`a:off @x`、`a:ext @cx` 等）。
 * @returns 四舍五入到整 px 的显示值。
 */
export function emuToPx(emu: number): number {
  return Math.round(emu / EMU_PER_PX);
}

/**
 * CSS px → EMU。
 *
 * @param px 视图里的位置或尺寸（96DPI 像素，可为负表示挪出幻灯片）。
 * @returns 写进 OOXML 的整数 EMU。
 */
export function pxToEmu(px: number): number {
  return Math.round(px * EMU_PER_PX);
}

/**
 * OOXML `sz`（半磅，1pt = 100）→ pt。
 *
 * @param sz `<a:rPr @sz>` 的原值。
 * @returns CSS pt（字号的显示单位）。
 */
export function halfPointToPt(sz: number): number {
  return sz / 100;
}

/**
 * pt → OOXML `sz`（半磅）。
 *
 * @param pt 用户给的字号 pt，须为正数（校验在调用方，这里只换算）。
 * @returns 写进 `sz` 的整数值。
 */
export function ptToHalfPoint(pt: number): number {
  return Math.round(pt * 100);
}
