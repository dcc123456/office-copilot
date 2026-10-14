/**
 * 把 `numFmtId`/`formatCode` 投影成给人看的文本（spec 验收 A3/A4 的显示侧）。
 *
 * 为什么放在 `sheet-io` 而不是网格里：`AGENTS.md` §2.1 要求序列号 ↔ 日期、数字 ↔ 显示
 * 只允许在 IO 层换算一次；渲染层再写一份 `new Date(serial)` 就是第二处换算，两边迟早不一致。
 * 因此这里只吃 reader 已经解好的 `numberFormat`，输出字符串，不回写任何模型字段。
 *
 * 覆盖面如实限定（够用即可，不是完整 Excel 格式引擎）：内建 id 的常用集，自定义码里的
 * `yyyy/yy/mmmm/mmm/mm/m/dd/dd/d/hh/h/mm(分)/ss/s/AM\/PM`、千分位 `,`、小数位 `0`/`#`、
 * 百分号 `%`、引号字面量与 `\` 转义。读不懂的格式**原样显示数值**，不猜格式。
 */
import type { CellKind, CellScalarValue, ResolvedNumberFormat, SheetCell } from "./model.js";
import { serialToUtcDate, utcDateToSerial } from "./datetime.js";

/**
 * 内建 `numFmtId` 的规范格式码（ECMA-376 18.8.30 附表的常用行）。
 * 表外 id 视为"我们不解释"，显示原始数值。
 */
const BUILTIN_FORMAT_CODES: Readonly<Record<number, string>> = {
  0: "General",
  1: "0",
  2: "0.00",
  3: "#,##0",
  4: "#,##0.00",
  9: "0%",
  10: "0.00%",
  14: "m/d/yyyy",
  15: "d-mmm-yy",
  16: "d-mmm",
  17: "mmm-yy",
  18: "h:mm AM/PM",
  19: "h:mm:ss AM/PM",
  20: "h:mm",
  21: "h:mm:ss",
  22: "m/d/yyyy h:mm",
  37: "#,##0;(#,##0)",
  39: "#,##0.00;(#,##0.00)",
  45: "mm:ss",
  46: "[h]:mm:ss",
  49: "@",
};

const MONTH_SHORT = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
const MONTH_FULL = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
const WEEKDAY_SHORT = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];

/** 格式码里一个日期/时间占位符（`len` 是原码里重复字母的个数；`meridiem` 是 AM/PM 指示符）。 */
interface DatePlaceholder {
  kind: "year" | "month" | "day" | "hour" | "minute" | "second" | "meridiem";
  len: number;
}

/** 一次格式码扫描的结果。 */
interface FormatPlan {
  /** 输出序列：字面量与日期占位符按格式码里的先后排好。 */
  parts: Array<string | DatePlaceholder>;
  isDateTime: boolean;
  isPercent: boolean;
  hasThousands: boolean;
  /** 小数位数（小数点后 `0`/`#` 的个数）；undefined = 格式码没规定，按原值显示。 */
  fractionDigits?: number;
  /** 出现 `[h]`：小时按总小时数显示（`[h]:mm:ss` 这类经过时间的写法）。 */
  isElapsedHours: boolean;
}

/** 扫描格式码第一段，得到显示计划。 */
function planFormat(formatCode: string | undefined): FormatPlan {
  const code = (formatCode ?? "General").split(";")[0]!;
  const parts: Array<string | DatePlaceholder> = [];
  const plan: FormatPlan = { parts, isDateTime: false, isPercent: false, hasThousands: false, isElapsedHours: false };
  let afterHours = false;
  let afterDecimalPoint = false;

  let index = 0;
  while (index < code.length) {
    const char = code[index]!;
    if (char === '"') {
      const end = code.indexOf('"', index + 1);
      parts.push(code.slice(index + 1, end < 0 ? code.length : end));
      index = end < 0 ? code.length : end + 1;
      continue;
    }
    if (char === "\\") {
      parts.push(code[index + 1] ?? "");
      index += 2;
      continue;
    }
    if (char === "[") {
      const end = code.indexOf("]", index);
      const inner = code.slice(index + 1, end < 0 ? code.length : end).toLowerCase();
      if (inner === "h") {
        plan.isElapsedHours = true;
        plan.isDateTime = true;
        parts.push({ kind: "hour", len: 1 });
      }
      index = end < 0 ? code.length : end + 1;
      continue;
    }
    const runLen = letterRunLength(code, index);
    const lower = char.toLowerCase();
    if (runLen > 0 && "ymdhs".includes(lower)) {
      // `m` 在 `h` 之后是分钟，否则是月份 —— Excel 自己也按位置判。
      const kind: DatePlaceholder["kind"] =
        lower === "y" ? "year" : lower === "d" ? "day" : lower === "h" ? "hour" : lower === "s" ? "second" : afterHours ? "minute" : "month";
      if (kind === "hour") afterHours = true;
      plan.isDateTime = true;
      parts.push({ kind, len: runLen });
      index += runLen;
      continue;
    }
    if (lower === "a" && code.slice(index, index + 5).toUpperCase() === "AM/PM") {
      // AM/PM 出现即改用 12 小时制，占位符本身在渲染时换成 "AM"/"PM"（空格留给格式码里的字面量）。
      parts.push({ kind: "meridiem", len: 1 });
      index += 5;
      continue;
    }
    if (char === "%") plan.isPercent = true;
    if (char === "," && /[0#]/.test(code.slice(index + 1))) plan.hasThousands = true;
    if (char === ".") afterDecimalPoint = true;
    if (afterDecimalPoint && (char === "0" || char === "#")) plan.fractionDigits = (plan.fractionDigits ?? 0) + 1;
    // 数字占位符与格式控制符不当地面文字；`-` `()` `:` `/` 空格在日期码里都是要照抄的分隔符。
    if (!"0#?,.%eE*@_".includes(char)) parts.push(char);
    index += 1;
  }
  return plan;
}

/** 从 `start` 起同类字母（忽略大小写）的连续个数；非字母返回 0。 */
function letterRunLength(code: string, start: number): number {
  const base = code[start];
  if (!base || !/[a-z]/i.test(base)) return 0;
  let count = 0;
  while (start + count < code.length && code[start + count]!.toLowerCase() === base.toLowerCase()) count += 1;
  return count;
}

/**
 * 按格式码显示一个数值。
 *
 * @param value 单元格数值（日期是序列号，不是 `Date`）。
 * @param format reader 解出的数字格式；undefined 按 General 显示。
 * @returns 给人看的字符串；非有限数原样返回字面量（`NaN` 之类要能被用户看见，不许伪装成 0）。
 */
export function formatNumberForDisplay(value: number, format?: ResolvedNumberFormat): string {
  if (!Number.isFinite(value)) return String(value);
  const code = format?.formatCode ?? (format ? BUILTIN_FORMAT_CODES[format.numFmtId] : undefined);
  const plan = planFormat(code);
  if (plan.isDateTime) return renderDateTime(value, plan);

  let shown = plan.isPercent ? value * 100 : value;
  shown = Number(shown.toPrecision(15));
  let text = plan.fractionDigits === undefined ? String(shown) : shown.toFixed(plan.fractionDigits);
  if (plan.hasThousands) {
    const [intPart, decPart] = text.split(".");
    text = intPart!.replace(/\B(?=(\d{3})+(?!\d))/g, ",") + (decPart === undefined ? "" : `.${decPart}`);
  }
  return plan.isPercent ? `${text}%` : text;
}

/**
 * 把序列号按格式计划渲染成日期/时间文本。
 *
 * @param serial OOXML 日期序列号（1900 历元，换算细节见 `datetime.ts`）。
 * @param plan `planFormat` 的结果。
 * @returns 渲染串；序列号超出可表示范围时退回数字原文，不许渲染成 `Invalid Date`。
 */
function renderDateTime(serial: number, plan: FormatPlan): string {
  const date = serialToUtcDate(Math.floor(serial));
  if (!date) return String(serial);
  const seconds = Math.round((serial - Math.floor(serial)) * 86400);
  const hours = Math.floor(seconds / 3600);
  const minutes = Math.floor((seconds % 3600) / 60);
  const secondsRest = seconds % 60;
  const month = date.getUTCMonth() + 1;
  const day = date.getUTCDate();
  const year = date.getUTCFullYear();
  const uses12HourClock = plan.parts.some((part) => typeof part === "object" && part.kind === "meridiem");

  let text = "";
  for (const part of plan.parts) {
    if (typeof part === "string") {
      text += part;
      continue;
    }
    switch (part.kind) {
      case "year":
        text += part.len >= 4 ? String(year) : String(year % 100).padStart(2, "0");
        break;
      case "month":
        text += part.len >= 4 ? MONTH_FULL[month - 1] : part.len === 3 ? MONTH_SHORT[month - 1] : part.len === 2 ? String(month).padStart(2, "0") : String(month);
        break;
      case "day":
        text += part.len >= 3 ? WEEKDAY_SHORT[date.getUTCDay()] : part.len === 2 ? String(day).padStart(2, "0") : String(day);
        break;
      case "hour": {
        // `[h]` 是"经过的小时数"，不做 12 小时制换算。
        const clock = !plan.isElapsedHours && uses12HourClock ? (hours % 12 === 0 ? 12 : hours % 12) : hours;
        text += part.len === 2 && !plan.isElapsedHours ? String(clock).padStart(2, "0") : String(clock);
        break;
      }
      case "meridiem":
        // 前导空格是格式码自带的字面量（内建 18 是 `h:mm AM/PM`），这里不能再补一个。
        text += hours < 12 ? "AM" : "PM";
        break;
      case "minute":
        text += part.len === 2 ? String(minutes).padStart(2, "0") : String(minutes);
        break;
      case "second":
        text += part.len === 2 ? String(secondsRest).padStart(2, "0") : String(secondsRest);
        break;
    }
  }
  return text;
}

/**
 * 单元格在网格里应当显示的文本。
 *
 * @param cell 单元格（`kind`/`value`/`numberFormat`/`cachedValue` 已由 reader 解好）。
 * @returns 显示串。日期与百分比按格式显示；公式格显示**缓存值**（无缓存值时显示空串，
 *   由 UI 另标"待 Excel 重算"）；只有格式没有值的格返回空串。
 */
export function cellDisplayText(cell: SheetCell): string {
  switch (cell.kind) {
    case "blank":
      return "";
    case "string":
    case "error":
      return String(cell.value ?? "");
    case "boolean":
      return cell.value ? "TRUE" : "FALSE";
    case "number": {
      const numeric = cell.formula ? cell.cachedValue : cell.value;
      if (numeric === null || numeric === undefined) return "";
      if (typeof numeric !== "number") return String(numeric);
      return formatNumberForDisplay(numeric, cell.numberFormat);
    }
  }
}

/** 用户输入解析的结果：正好是写进 `SheetCell` 所需的那几个字段。 */
export interface ParsedCellInput {
  kind: CellKind;
  value: CellScalarValue;
  /** 输入本身暗示的数字格式（日期、百分比）；纯数字与文本为 undefined = 不改格式。 */
  numberFormat?: ResolvedNumberFormat;
}

/** 内建"短日期"格式 id（ECMA-376 18.8.30），也是新输入日期采用的显示格式。 */
/**
 * 把给人看的格式码解析成写进格子的显示意图（spec 决策 D10 第 3 条）。
 *
 * 只认两张表：`BUILTIN_FORMAT_CODES` 里的内建码，以及投影层盖得住的"固定小数位"形状
 * （`0.0`、`0.000`、`0.0%`、`0.000%` 这类，键入时 `percentFormatWithDecimals` 产的就是同一套）。
 * 表外一律回 null——调用方必须把这句话报给人，不许把认不出的串原样塞进 `<numFmt>`。
 *
 * @param formatCode 格式码原文（大小写不敏感，去首尾空格）。
 * @returns 写进 `SheetCell.numberFormat` 的显示意图；认不出时 `null`。
 */
export function numberFormatFromCode(formatCode: string): ResolvedNumberFormat | null {
  const code = formatCode.trim();
  for (const [rawId, builtin] of Object.entries(BUILTIN_FORMAT_CODES)) {
    if (builtin.toLowerCase() !== code.toLowerCase()) continue;
    const numFmtId = Number(rawId);
    return { numFmtId, isDateLike: BUILTIN_DATE_NUMFMT_IDS.has(numFmtId) };
  }
  const percent = /^0(?:\.(0+))?%$/.exec(code);
  if (percent) {
    const decimals = percent[1]?.length ?? 0;
    // 四位以上小数位不是我们要给的习惯写法，且 Excel 里也不常见，按认不出处理。
    return decimals > 3 ? null : percentFormatWithDecimals(decimals);
  }
  const fixed = /^0\.(0+)$/.exec(code);
  if (fixed) {
    const decimals = fixed[1]!.length;
    if (decimals > 3) return null;
    if (decimals === 2) return { numFmtId: 2, isDateLike: false };
    return { numFmtId: 0, formatCode: code, isDateLike: false };
  }
  // 日期/时间的自定义写法（`yyyy-mm-dd` 这类）：显示由上面的投影层按 formatCode 算，所以认得。
  if (/^[ymdhs][ymdhs/\-.:, ]*$/i.test(code) && /[ymd]/i.test(code)) return { numFmtId: 0, formatCode: code, isDateLike: true };
  return null;
}

/** ECMA-376 18.8.30 里属日期/时间的那几个内建 id（`BUILTIN_FORMAT_CODES` 表中的 14–22、45、46）。 */
const BUILTIN_DATE_NUMFMT_IDS = new Set([14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46]);

const BUILTIN_DATE_FORMAT_NUMFMT_ID = 14;

/** 只有 ISO 风格的完整日期才当日期吃：`1-2` 在 Excel 里是 jan-02，在本工具里宁可当文本。 */
const ISO_DATE_PATTERN = /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/;
const PLAIN_NUMBER_PATTERN = /^-?\d+(?:\.\d+)?(?:[eE][+-]?\d+)?$/;
const PERCENT_PATTERN = /^-?\d+(?:\.\d+)?%$/;

/**
 * 选一个"显示出来跟用户敲进去的一模一样"的百分比格式。
 *
 * 为什么不是所有百分比都套内建 10 号：ECMA-376 里 10 是 `0.00%`、9 才是 `0%`，
 * 一律给 10 会让键入 `12%` 显示成 `12.00%`（实测记在 spec §七）。内建表盖不住的位数
 * 用自定义格式码表达，保存时由 `StyleTable` 追加进 `<numFmts>`。
 *
 * @param decimals 用户输入里小数点后的位数（`12.50%` → 2）。
 * @returns 写进 `SheetCell.numberFormat` 的显示意图。
 */
function percentFormatWithDecimals(decimals: number): ResolvedNumberFormat {
  if (decimals === 0) return { numFmtId: 9, isDateLike: false };
  if (decimals === 2) return { numFmtId: 10, isDateLike: false };
  return { numFmtId: 0, formatCode: `0.${"0".repeat(decimals)}%`, isDateLike: false };
}

/**
 * 把用户在手改单元格里输入的文本解析成格子内容（`AGENTS.md` §2.1：日期 ↔ 序列号只能在这里换算）。
 *
 * @param text 编辑器里的原文；前后空白会被忽略。
 * @returns `kind`/`value`，以及输入自己暗示的数字格式（`2026-01-31` → 内建日期格式、
 *   `12%` → 与输入小数位数一致的百分号格式）；空文本返回 `kind:"blank"` 且 `value:null`，调用方据此删格。
 */
export function parseUserCellText(text: string): ParsedCellInput {
  const trimmed = text.trim();
  if (trimmed === "") return { kind: "blank", value: null };
  if (trimmed.toUpperCase() === "TRUE") return { kind: "boolean", value: true };
  if (trimmed.toUpperCase() === "FALSE") return { kind: "boolean", value: false };
  if (PLAIN_NUMBER_PATTERN.test(trimmed)) return { kind: "number", value: Number(trimmed) };
  if (PERCENT_PATTERN.test(trimmed)) {
    const body = trimmed.slice(0, -1);
    const decimals = body.includes(".") ? body.split(".")[1]!.length : 0;
    return { kind: "number", value: Number(body) / 100, numberFormat: percentFormatWithDecimals(decimals) };
  }
  const dateMatch = ISO_DATE_PATTERN.exec(trimmed);
  if (dateMatch) {
    const [, year, month, day] = dateMatch;
    const serial = utcDateToSerial(new Date(Date.UTC(Number(year), Number(month) - 1, Number(day))));
    if (serial !== undefined) {
      return { kind: "number", value: serial, numberFormat: { numFmtId: BUILTIN_DATE_FORMAT_NUMFMT_ID, isDateLike: true } };
    }
  }
  return { kind: "string", value: trimmed };
}
