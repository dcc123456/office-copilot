/**
 * OOXML 的 8 位 ARGB 与编辑器/CSS 的 `#rrggbb` 互转（spec §零.2 的单点换算）。
 *
 * 实测（spec §二 P3）：`.xlsx` 的颜色有两种并存形态 `<color rgb="FFFF0000"/>`（ARGB，
 * 透明度在前）与 `<color theme="1"/>`（主题色引用）。主题色解析依赖 `theme1.xml`，
 * 阶段 1 只做**原样保留**（`themeRef`），不替 Excel 猜它渲染成什么色。
 */

/** 颜色：`hex` 是 `#rrggbb`，`alpha` 只在非不透明时存在（0–1）。 */
export interface SheetColor {
  hex: string;
  alpha?: number;
  /** 主题色引用（`<color theme="1" tint="0.4"/>`）原样带着，不建模成具体颜色。 */
  themeRef?: { theme: number; tint?: number };
  /** 读不懂时的原始串，写回时逐字节搬运，禁止塞进不认识的值。 */
  raw?: string;
}

/** 字体属性（`sizePt` 保持 pt，换算成 px 是出口/渲染层唯一一处的事）。 */
export interface FontSpec {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  sizePt?: number;
  name?: string;
  color?: SheetColor;
}

/**
 * ARGB（8 位十六进制）→ SheetColor。
 *
 * @param argb 形如 `FFFF0000`；不含 `#`。
 * @returns `hex` 为 `#rrggbb`，全不透明时省略 `alpha`；长度或字符合法性不对时返回
 *   `{ raw }`，含义是"这里没读懂，别写它"。
 */
export function argbToSheetColor(argb: string): SheetColor {
  if (!/^[0-9a-fA-F]{8}$/.test(argb)) return { hex: "", raw: argb };
  const alpha = parseInt(argb.slice(0, 2), 16) / 255;
  const hex = `#${argb.slice(2).toLowerCase()}`;
  return alpha >= 0.999 ? { hex } : { hex, alpha: Math.round(alpha * 1000) / 1000 };
}

/**
 * SheetColor → ARGB，用于写回 `@rgb`。
 *
 * @param color 颜色对象。
 * @returns 8 位大写 ARGB；没有可用十六进制时返回 undefined（调用方不写这个属性）。
 */
export function sheetColorToArgb(color: SheetColor): string | undefined {
  const match = /^#?([0-9a-fA-F]{6})$/.exec(color.hex);
  if (!match) return undefined;
  const alpha = color.alpha === undefined ? 1 : Math.min(Math.max(color.alpha, 0), 1);
  return `${Math.round(alpha * 255).toString(16).padStart(2, "0")}${match[1]}`.toUpperCase();
}
