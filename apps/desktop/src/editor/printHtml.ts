/**
 * Build a self-contained HTML document for printing / PDF export.
 *
 * We render the editor's current HTML (from `editor.getHTML()`) inside
 * A4-sized `<section class="page">` blocks. Print media CSS forces
 * page breaks after each section. Because the editor itself does
 * visual pagination via Decoration widgets, we strip those before
 * printing and let Chromium re-paginate natively — that produces
 * crisp PDFs without relying on our measurement.
 */

import type {
  LinkStyle,
  SectionPageBorders,
  SectionWatermark,
} from "@office-copilot/docx-io";
import { borderLinesFor, borderSideToCss } from "./pageBorders.js";
import { watermarkLayerFor } from "./pageWatermark.js";
import { escapeHtml } from "./renderHeaderFooter.js";
// 首字下沉那条 ::first-letter 规则只有全局一份才不会出现两份实现（§3.2）：纸上直接引
// 画布用的同一份样式文本，而不是再抄一遍 initial-letter。
import dropCapCss from "../styles/dropCap.css?raw";
import pageWatermarkCss from "../styles/pageWatermark.css?raw";

export function buildPrintHtml(opts: {
  bodyHtml: string;
  headerHtml?: string;
  footerHtml?: string;
  pageWidthPx: number;
  pageHeightPx: number;
  marginTopPx: number;
  marginRightPx: number;
  marginBottomPx: number;
  marginLeftPx: number;
  headerPx: number;
  footerPx: number;
  /**
   * 分栏（`w:cols`）：栏数与栏间距，CSS px。栏数小于 2 时不写任何列样式，
   * 和 Word 的单栏文档一致。
   */
  columns?: { count: number; gapPx: number } | null;
  /**
   * 页面边框（`w:pgBorders`）：四条边的线型/线宽/间距/颜色。null 或四条边都关时
   * 纸上什么都不画，和 Word 一致。
   */
  pageBorders?: SectionPageBorders | null;
  /**
   * 水印（存在页眉里的那个文字艺术形状）。null 或没设过时纸上什么都不画。
   */
  watermark?: SectionWatermark | null;
  link?: LinkStyle | null;
}): string {
  const {
    bodyHtml,
    headerHtml,
    footerHtml,
    pageWidthPx,
    pageHeightPx,
    marginTopPx,
    marginRightPx,
    marginBottomPx,
    marginLeftPx,
    headerPx,
    footerPx,
    columns,
    pageBorders,
    watermark,
    link,
  } = opts;

  const linkColor =
    link?.color && /^#[0-9a-f]{6}$/i.test(link.color) ? link.color : "#0563c1";
  const linkRule = link?.underline === false ? "none" : "underline";

  // 纸上的分栏用 CSS multi-column：`column-fill: auto` 才和 Word 一样"第一栏排满再排
  // 第二栏"，而不是把内容平衡成两截短栏。这里**不能**给容器定高——定高之后 Blink 会把
  // 装不下的内容横向溢出成第三、第四栏再被页面裁掉，整份文档只剩第一页（实测见
  // docs/word-parity-spec.md §8.13）。跨页断栏交给打印分页自己发生。
  const columnRule =
    columns && columns.count > 1
      ? `column-count: ${columns.count}; column-gap: ${columns.gapPx}px; column-fill: auto;`
      : "";

  // 页面边框：每条边一个绝对定位的空盒子，靠 border-top / border-left 画出线来。
  // 几何与线型换算都取自 pageBorders 那份共享实现，画布预览与纸上才会一模一样。
  const px = (v: number) => Math.round(v * 100) / 100;
  const borderHtml = pageBorders
    ? borderLinesFor(pageBorders, {
        pageWidth: pageWidthPx,
        pageHeight: pageHeightPx,
        marginTop: marginTopPx,
        marginRight: marginRightPx,
        marginBottom: marginBottomPx,
        marginLeft: marginLeftPx,
      })
        .map((line) => {
          const css = borderSideToCss(line.side);
          const border = `${css.widthPx}px ${css.style} ${css.color}`;
          const box = line.horizontal
            ? `width: ${px(line.length)}px; border-top: ${border};`
            : `height: ${px(line.length)}px; border-left: ${border};`;
          return `<div class="pgb" style="left: ${px(line.x)}px; top: ${px(line.y)}px; ${box}"></div>`;
        })
        .join("")
    : "";

  // 水印：Word 把它放在页眉里，所以每一页都有一份。纸上靠 `position: fixed` 达到同样的
  // 效果——Chromium 打印时固定元素每页重画一次。几何与画布共用 watermarkLayerFor。
  const layer = watermark
    ? watermarkLayerFor(watermark, {
        pageWidth: pageWidthPx,
        pageHeight: pageHeightPx,
        marginTop: marginTopPx,
        marginRight: marginRightPx,
        marginBottom: marginBottomPx,
        marginLeft: marginLeftPx,
      })
    : null;
  const watermarkHtml = layer
    ? `<div class="oc-page-watermark" style="left: ${px(layer.leftPx)}px; ` +
      `top: ${px(layer.topPx)}px; width: ${px(layer.widthPx)}px; ` +
      `height: ${px(layer.heightPx)}px; font-size: ${px(layer.fontSizePx)}px; ` +
      `font-family: ${layer.fontFamily}; color: ${layer.color}; ` +
      `opacity: ${layer.opacity}; transform: rotate(${layer.rotateDeg}deg);">` +
      `${escapeHtml(layer.text)}</div>`
    : "";

  // Reuse the editor's stylesheet? We inline a minimal stylesheet that
  // makes the output look close to what the user sees.
  const css = `
    @page { size: ${pageWidthPx}px ${pageHeightPx}px; margin: 0; }
    * { box-sizing: border-box; }
    html, body { margin: 0; padding: 0; background: #fff; color: #000;
      font-family: "Microsoft YaHei", "PingFang SC", "Hiragino Sans GB",
        "Source Han Sans SC", "Noto Sans CJK SC", "Segoe UI", sans-serif;
      font-size: 14px; line-height: 1.5; }
    .page {
      position: relative;
      width: ${pageWidthPx}px;
      /* 只能是 min-height：定高 + overflow:hidden 会让 Blink 把这整块当成不可断的
         裁剪容器，第二页以后的正文直接消失（一份 60 段的文档只出 1 页）。 */
      min-height: ${pageHeightPx}px;
      padding: ${marginTopPx}px ${marginRightPx}px ${marginBottomPx}px ${marginLeftPx}px;
      page-break-after: always;
      break-after: page;
    }
    .page:last-child { page-break-after: auto; }
    /* 页面边框固定定位：Word 的边框是每页一条，Chromium 打印时固定元素逐页重画。 */
    .pgb { position: fixed; margin: 0; padding: 0; background: none; }
    .ph { position: absolute; left: ${marginLeftPx}px; right: ${marginRightPx}px;
           top: ${headerPx}px; font-size: 10px; color: #666; }
    .pf { position: absolute; left: ${marginLeftPx}px; right: ${marginRightPx}px;
           bottom: ${footerPx}px; font-size: 10px; color: #666; text-align: center; }
    /* Editor block styles */
    .ProseMirror { outline: none; ${columnRule} }
    p { margin: 0 0 6pt; }
    h1 { font-size: 22pt; margin: 12pt 0 6pt; }
    h2 { font-size: 16pt; margin: 10pt 0 4pt; }
    h3 { font-size: 13pt; margin: 8pt 0 4pt; }
    h4 { font-size: 12pt; margin: 7pt 0 3pt; }
    h5 { font-size: 11pt; margin: 6pt 0 2pt; }
    h6 { font-size: 10.5pt; margin: 5pt 0 2pt; }
    ul, ol { margin: 4pt 0 4pt 24pt; padding: 0; }
    blockquote { border-left: 3px solid #ccc; padding-left: 10px; color: #555;
      margin: 6pt 0; }
    table { border-collapse: collapse; }
    table td, table th { border: 1px solid #888; padding: 4px 8px; }
    img { max-width: 100%; }
    a { color: ${linkColor}; text-decoration: ${linkRule}; }
    .pm-field {
      display: inline-block; padding: 0 4px; border-radius: 3px;
      background: #eaf3fb; color: #0078d4;
    }
    /* The page-gap widgets from our editor shouldn't appear on paper. */
    .oc-page-gap { display: none !important; }
    ${pageWatermarkCss}
    /* 纸上的水印每页都要有一份（Word 是放在页眉里的），固定元素在 Chromium 打印时逐页重画。 */
    .oc-page-watermark { position: fixed; }
    ${dropCapCss}
  `;

  return `<!doctype html>
<html><head><meta charset="utf-8"/><title>Print</title>
<style>${css}</style></head>
<body>
  <section class="page">
    ${headerHtml ? `<div class="ph">${headerHtml}</div>` : ""}
    <div class="ProseMirror">${bodyHtml}</div>
    ${footerHtml ? `<div class="pf">${footerHtml}</div>` : ""}
    ${borderHtml}
    ${watermarkHtml}
  </section>
</body></html>`;
}
