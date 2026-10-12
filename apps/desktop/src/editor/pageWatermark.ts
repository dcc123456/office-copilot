import type { SectionWatermark } from "@office-copilot/docx-io";
import { PT_TO_PX } from "./pageBorders.js";

/**
 * 水印在屏幕与纸面上的呈现值。
 *
 * Word 把水印画成页眉里的一个 VML 文字艺术形状（`v:shape` + `v:textpath`），
 * 形状框固定 468pt × 117pt、居中于页边距之内。这里把那一份形状换算成一份 CSS，
 * 编辑画布与打印/PDF 共用，避免出现两套"看起来像水印"的东西（AGENTS §3.2）。
 * 单位约定同 `pageBorders.ts`：模型里是 pt，只有落到屏幕才乘 `PT_TO_PX`（§2.1）。
 */

/** Word 给水印形状的那个框，pt——`v:shape/@style` 里的 width/height。 */
const SHAPE_BOX_PT = { width: 468, height: 117 };

/** 半透明对应 `v:fill/@opacity="50%"`。 */
const SEMI_TRANSPARENT_ALPHA = 0.5;

/** 对角版式对应 VML `rotation="315"`，也就是 CSS 的 -45 度（左下→右上）。 */
const DIAGONAL_ROTATE_DEG = -45;

/** 读不懂的颜色按黑色画，不把非法值写进 CSS（AGENTS §2.2）。 */
const FALLBACK_COLOR = "#000000";

/** 一条水印落到页面上要用的盒子与文字呈现，全部 CSS px。 */
export interface WatermarkLayer {
  text: string;
  /** 相对页面左上角的盒子位置。 */
  leftPx: number;
  topPx: number;
  widthPx: number;
  heightPx: number;
  fontSizePx: number;
  rotateDeg: number;
  /** 已经带引号、可直接拼进 CSS 的 `font-family` 值。 */
  fontFamily: string;
  color: string;
  opacity: number;
}

/** 页面几何里水印只用到这六项，CSS px。 */
export interface WatermarkPage {
  pageWidth: number;
  pageHeight: number;
  marginTop: number;
  marginRight: number;
  marginBottom: number;
  marginLeft: number;
}

/**
 * 估算一行文字在给定字号下占多宽（CSS px）。
 *
 * 打印 HTML 是一段字符串，没有 DOM 可量，所以按经验值估：CJK 与全角标点算 1 em，
 * 其余（拉丁字母、数字、空格）算 0.55 em。只用于"自动"字号的缩放，不参与落盘。
 *
 * @param text 水印文字。
 * @returns 文字总宽度相当于多少个 em。
 */
export function textEmCount(text: string): number {
  let em = 0;
  for (const ch of text) em += ch.codePointAt(0)! > 0x2e80 ? 1 : 0.55;
  return em;
}

/**
 * 把字体名变成可以安全拼进 `style="..."` 的 CSS 值。
 *
 * 字体名是用户在对话框里填的，引号、尖括号、分号都可能进来；打印 HTML 是字符串拼接，
 * 一个漏网的 `"` 就能把水印那一行变成注入。所以这里按黑名单整串拒绝——白名单会误伤
 * `宋体`、`MS Mincho`、`Segoe UI` 这类正常名字。
 *
 * @param font 字体族名。
 * @returns 带双引号的 CSS `font-family` 值；含注入字符或为空时给 `"serif"`。
 */
function safeFontFamily(font: string): string {
  const trimmed = font.trim();
  if (!trimmed || /["'<>();\\&{}]/.test(trimmed)) return '"serif"';
  return `"${trimmed}"`;
}

/**
 * 把水印模型换算成画布与纸面共用的那一份呈现值。
 *
 * @param wm   水印设置。`sizePt` 为 null 是 Word 的"自动"（`fitshape`），
 *   Word 会把文字非等比拉满形状框，这里按等比取"框高"和"框宽/字数"里的较小值，
 *   所以预览里的自动水印比 Word 略小，但字号与位置关系一致。
 * @param page 页面几何（CSS px）。
 * @returns 盒子与文字呈现；文字为空时返回 null，调用方什么都不画。
 */
export function watermarkLayerFor(
  wm: SectionWatermark,
  page: WatermarkPage,
): WatermarkLayer | null {
  const text = wm.text.trim();
  if (!text) return null;
  const widthPx = SHAPE_BOX_PT.width * PT_TO_PX;
  const heightPx = SHAPE_BOX_PT.height * PT_TO_PX;
  const contentWidth = page.pageWidth - page.marginLeft - page.marginRight;
  const contentHeight = page.pageHeight - page.marginTop - page.marginBottom;
  const fontSizePx =
    wm.sizePt && wm.sizePt > 0
      ? wm.sizePt * PT_TO_PX
      : Math.min(heightPx, widthPx / Math.max(textEmCount(text), 1));
  return {
    text,
    leftPx: page.marginLeft + (contentWidth - widthPx) / 2,
    topPx: page.marginTop + (contentHeight - heightPx) / 2,
    widthPx,
    heightPx,
    fontSizePx,
    rotateDeg: wm.layout === "diagonal" ? DIAGONAL_ROTATE_DEG : 0,
    fontFamily: safeFontFamily(wm.font),
    color: /^#[0-9a-f]{6}$/i.test(wm.color) ? wm.color : FALLBACK_COLOR,
    opacity: wm.semiTransparent ? SEMI_TRANSPARENT_ALPHA : 1,
  };
}
