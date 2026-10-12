/**
 * Excel 日期序列号 ↔ JS Date。
 *
 * 实测（spec §二 P3）：日期格在 `sheet1.xml` 里**没有类型标记**，长这样
 * `<c r="D2" s="2"><v>46053</v></c>` —— 值就是序列号，"它是日期"这件事只由样式
 * `s` 指向的 `cellXfs/@numFmtId` 决定。所以序列号换算在全仓只允许出现在这里。
 */

const MS_PER_DAY = 86_400_000;

/**
 * 1900 起算的"第 0 天"：Excel 的 1 是 1900-01-01，且它把不存在的 1900-02-29 当作第 60 天。
 * 于是序列号 ≥61 用 1899-12-30 起算，≤59 用 1899-12-31 起算，60 是无法表示的假日期的兼容值。
 */
const EPOCH_1899_12_30_UTC = Date.UTC(1899, 11, 30);
const EPOCH_1899_12_31_UTC = Date.UTC(1899, 11, 31);

/**
 * 序列号 → UTC 零点的 Date。
 *
 * @param serial Excel 序列号（可含小数表示一天内的时间）。
 * @returns 对应的时间；`serial === 60`（Excel 的假 1900-02-29）与非法输入返回 undefined，
 *   调用方按"不是日期"处理显示，不要凭空造一天。
 */
export function serialToUtcDate(serial: number): Date | undefined {
  if (!Number.isFinite(serial)) return undefined;
  if (serial === 60) return undefined;
  const dayParts = Math.floor(serial);
  const fractionMs = Math.round((serial - dayParts) * MS_PER_DAY);
  if (dayParts < 0) return undefined;
  const base = dayParts < 60 ? EPOCH_1899_12_31_UTC : EPOCH_1899_12_30_UTC;
  return new Date(base + dayParts * MS_PER_DAY + fractionMs);
}

/**
 * Date → 序列号（按 UTC 日历日，与 `serialToUtcDate` 严格互逆）。
 *
 * @param date 任意 Date；只取它的 UTC 年月日与时分秒，不做本地时区解释。
 * @returns Excel 序列号；早于 1900-01-01 或算出第 60 天时返回 undefined（Excel 无处安放）。
 */
export function utcDateToSerial(date: Date): number | undefined {
  const msOfDay = date.getUTCHours() * 3_600_000 + date.getUTCMinutes() * 60_000 + date.getUTCSeconds() * 1000 + date.getUTCMilliseconds();
  const dayParts = Math.round((Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate()) - EPOCH_1899_12_30_UTC) / MS_PER_DAY);
  if (dayParts <= 0) return undefined;
  const serial = dayParts + msOfDay / MS_PER_DAY;
  // 从 1900-01-01 到 1900-02-28 这段要减掉那个不存在的第 60 天，否则比 Excel 多一天。
  return dayParts <= 60 ? serial - 1 : serial;
}

/** 内建数字格式里属于日期/时间的那些 `@numFmtId`（ECMA-376 18.8.30 的表）。 */
export const BUILTIN_DATE_TIME_NUMFMT_IDS: ReadonlySet<number> = new Set([
  14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47,
]);

/**
 * 判断一个格式码是不是日期/时间格式。
 *
 * @param formatCode 自定义格式码，如 `yyyy-mm-dd`、`#,##0`、`0.00%`。
 * @returns 含日期 token 且不在引号/转义里时为 true；空串为 false。
 */
export function isDateFormatCode(formatCode: string | undefined): boolean {
  if (!formatCode) return false;
  // 引号内与 `\` 后的字符是字面量，不能当成日期 token（否则 `"yyyy"` 会被误判成日期）。
  const stripped = formatCode.replace(/"[^"]*"/g, "").replace(/\\./g, "");
  return /[ymdhs]|ggsla/i.test(stripped);
}
