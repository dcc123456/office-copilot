import type {
  PageBorderSide,
  SectionPageBorders,
} from "@office-copilot/docx-io";

/**
 * 页面边框（`w:pgBorders`）在桌面上共享的一份翻译层：对话框、编辑画布、打印/PDF
 * 三处都要把「OOXML 的线型 + 八分之一 pt 线宽 + pt 间距」变成能画的东西，
 * 所以只在这里写一次（AGENTS §3.2/§3.3）。
 *
 * 单位约定：模型里的 `sizePt`/`spacePt` 就是 OOXML 的原生单位（pt），
 * 只有落到屏幕/纸面时才乘 `PT_TO_PX` 换算成 CSS px（AGENTS §2.1 的换算发生在
 * docx-io 的进出口，这里属于渲染层，不做第二次单位制转换）。
 */

/** CSS px / pt：96 DPI 下一 pt 占多少 px。 */
export const PT_TO_PX = 96 / 72;

/** `w:pgBorders` 的四条边；顺序就是 ECMA-376 示例里子元素的书写顺序。 */
export type BorderSideKey = "top" | "left" | "bottom" | "right";

/** 页面边框对话框里可选的线型；`val` 是 ECMA-376 §17.3.4.14 `ST_Border` 的字面量。 */
export const BORDER_STYLE_OPTIONS: readonly { val: string; label: string }[] = [
  { val: "none", label: "无" },
  { val: "single", label: "单线" },
  { val: "thick", label: "粗线" },
  { val: "double", label: "双线" },
  { val: "dotted", label: "点线" },
  { val: "dashed", label: "虚线" },
  { val: "dotDash", label: "点划线" },
  { val: "dotDotDash", label: "双点划线" },
  { val: "triple", label: "三线" },
  { val: "thinThickSmallGap", label: "细粗线" },
  { val: "thickThinSmallGap", label: "粗细线" },
  { val: "wave", label: "波浪线" },
];

/**
 * OOXML 线型 → CSS `border-style`。
 *
 * CSS 只有 4 种实/虚线型，所以 Word 的复合线型按"最接近的观感"降级：
 * 三线与一粗一细归到 `double`，波浪线归到 `solid`（浏览器画不出波浪）。
 * 这只影响预览与 PDF；写进 `.docx` 的永远是 `BORDER_STYLE_OPTIONS` 里的原始字面量。
 */
export function borderStyleToCss(val: string): string {
  switch (val) {
    case "dotted":
    case "dotDotDash":
      return "dotted";
    case "dashed":
    case "dotDash":
    case "dashSmallGap":
    case "dashDotStroked":
      return "dashed";
    case "double":
    case "triple":
    case "thinThickSmallGap":
    case "thinThickLargeGap":
    case "thinThickMediumGap":
    case "thickThinSmallGap":
    case "thickThinLargeGap":
    case "thickThinMediumGap":
    case "doubleWave":
      return "double";
    default:
      return "solid";
  }
}

/** 一条边框画成什么样：CSS 的 border-width / style / color 三段。 */
export interface BorderCss {
  widthPx: number;
  style: string;
  color: string;
}

/**
 * 把模型里的一条边转成 CSS 尺寸。
 *
 * @param side 一条边的设置（线宽与间距是 pt，`color` 为 null 表示 Word 的"自动"）。
 * @returns 画这条线用的 CSS 值；`double` 至少 3px 才画得出两根线，
 *   其余线型至少 1px。
 */
export function borderSideToCss(side: PageBorderSide): BorderCss {
  const style = borderStyleToCss(side.style);
  const raw = side.sizePt * PT_TO_PX;
  const min = style === "double" ? 3 : 1;
  return {
    widthPx: Math.max(min, Math.round(raw * 10) / 10),
    style,
    color: side.color ?? "#000000",
  };
}

/** 用户新加一条边框时的初值：½pt 单线、距边 1pt、自动色（Word 对话框的缺省）。 */
export function defaultBorderSide(): PageBorderSide {
  return { style: "single", sizePt: 0.5, spacePt: 1, color: null };
}

/** 四条边都关掉的边框，作为"文档原本没有边框"时对话框的起点。 */
export function emptyPageBorders(): SectionPageBorders {
  return { offsetFrom: "page", top: null, left: null, bottom: null, right: null };
}

/** 是否至少有一条边开着；全关时保存要把 `w:pgBorders` 整个删掉。 */
export function hasAnyBorderSide(b: SectionPageBorders | null): boolean {
  if (!b) return false;
  return !!(b.top || b.left || b.bottom || b.right);
}

/** 一条边框线在页面坐标里的位置（CSS px，相对页面左上角）。 */
export interface BorderLine {
  /** 页面方向：true 表示横线（上下边），false 表示竖线（左右边）。 */
  horizontal: boolean;
  /** 线起点 x。 */
  x: number;
  /** 线起点 y。 */
  y: number;
  /** 线的长度。 */
  length: number;
  /** 对应 `SectionPageBorders` 上的字段名。 */
  key: BorderSideKey;
  /** 这条边的设置，省掉调用方再一次"它存不存在"的判断。 */
  side: PageBorderSide;
}

/**
 * 画边框只要页面外框与四个页边距，所以用这个窄接口：编辑视图传整份
 * `SectionGeometry`，打印层用手头的几个 px 数值就能拼出来。
 */
export interface PageBoxForBorders {
  pageWidth: number;
  pageHeight: number;
  marginTop: number;
  marginRight: number;
  marginBottom: number;
  marginLeft: number;
}

/**
 * 算出四条边框线的几何。
 *
 * @param borders 边框设置（`spacePt` 是 pt，这里才换算成 px）。
 * @param g 页面几何，全部 CSS px。
 * @returns 只包含开着的那些边；线画在"距页面/距正文"的位置上，
 *   与 `w:offsetFrom` 一致（`page` 从页边往里量，`text` 从正文边界往外量）。
 */
export function borderLinesFor(
  borders: SectionPageBorders,
  g: PageBoxForBorders,
): BorderLine[] {
  const fromPage = borders.offsetFrom !== "text";
  const spaceOf = (side: PageBorderSide | null) => (side?.spacePt ?? 0) * PT_TO_PX;
  // `w:space` 是"边框离参照物多远"，参照物由 `w:offsetFrom` 决定：以页面为基准时从
  // 页边往里量，以正文为基准时边框落在正文区与页边之间，所以上/左要往回退、下/右要往外扩。
  const topY = fromPage ? spaceOf(borders.top) : g.marginTop - spaceOf(borders.top);
  const leftX = fromPage ? spaceOf(borders.left) : g.marginLeft - spaceOf(borders.left);
  const bottomY = fromPage
    ? g.pageHeight - spaceOf(borders.bottom)
    : g.pageHeight - g.marginBottom + spaceOf(borders.bottom);
  const rightX = fromPage
    ? g.pageWidth - spaceOf(borders.right)
    : g.pageWidth - g.marginRight + spaceOf(borders.right);

  const lines: BorderLine[] = [];
  if (borders.top)
    lines.push({ horizontal: true, x: leftX, y: topY, length: rightX - leftX, key: "top", side: borders.top });
  if (borders.bottom)
    lines.push({
      horizontal: true,
      x: leftX,
      y: bottomY,
      length: rightX - leftX,
      key: "bottom",
      side: borders.bottom,
    });
  if (borders.left)
    lines.push({ horizontal: false, x: leftX, y: topY, length: bottomY - topY, key: "left", side: borders.left });
  if (borders.right)
    lines.push({
      horizontal: false,
      x: rightX,
      y: topY,
      length: bottomY - topY,
      key: "right",
      side: borders.right,
    });
  return lines;
}
