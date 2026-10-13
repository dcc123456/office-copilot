/**
 * 公式引用的位移（阶段 3 插删行列的配套，spec D4 的前置一步）。
 *
 * 规则有出处，不是猜的：
 * - `$` 是**复制/填充时**的钉子，不是插删时的钉子——
 *   [在相对引用、绝对引用和混合引用间切换](https://support.microsoft.com/zh-cn/Excel/switch-between-relative-absolute-and-mixed-references)
 *   原文只说"如果想要在**复制**时保留…原始单元格引用，则可以在…之前加上美元符号（$），使单元格引用绝对"，
 *   整篇没有把 `$` 与插入/删除挂钩。插删改的是格子本身，引用认的是格子身份，所以 `SUM($A$1:$A$10)`
 *   在第 1 行上面插一行照样变成 `SUM($A$2:$A$11)`——**两处的位移我们都照搬，只保留 `$` 的写法**。
 * - 引用了的格被删掉，Excel 把公式变成 `#REF!`——
 *   [How to correct a #REF! error](https://support.microsoft.com/en-us/office/how-to-correct-a-ref-error-822c8e46-e610-4d02-bf29-ec4b8c5ff4be)：
 *   "when cells that were referenced by formulas get deleted"。我们**不产** `#REF!`：
 *   打开就是坏公式的文件不如不写，整条 op 拦下来并把理由报给用户。
 * - 删掉区间**中间**的行时区间收缩（`SUM(C2:C4)` 删第 3 行 → `SUM(C2:C3)`），这与 Excel 一致：
 *   区间两个端点都不在被删带里，就只是各自上移。
 *
 * 认不出的写法（跨表 `!`、外部工作簿 `[1]`、结构化表 `[...]`、数组常量 `{}`、已有的 `#`）
 * 一律拒绝：它们的位移要连表名/表结构一起解析，那属于阶段 4 的公式求值器。
 */
import { colNameToIndex, indexToColName } from "./address.js";
import type { SheetAxis } from "./sheet-edit.js";

/** 位移结果：能安全平移就给出新公式，否则给拒绝理由（调用方整条 op 跳过，不许猜位置）。 */
export type FormulaShift = { ok: true; formula: string } | { ok: false; reason: string };

/** 被删掉的连续区间（0 基，含两端）；插入时没有这个。 */
export interface FormulaCut {
  startIndex: number;
  endIndex: number;
}

/**
 * 一次扫描的三种命中：字符串字面量 / 我们不解析的构造 / A1 风格引用。
 *
 * 引用必须"列字母 + 行数字"都齐（`C2`、`$C$2`）；整列引用 `C:C` 匹配不上，正符合它不动的语义。
 */
const TOKEN_PATTERN = /("(?:[^"]|"")*")|([!\[\]{}#])|((\$?)([A-Za-z]{1,3})(\$?)([1-9][0-9]{0,6}))/g;

/** 引用前后出现字母/数字/下划线/点/`$`，说明它不是独立引用而是某个名字的一部分（如 `Table1`、`ATAN2`）。 */
const NAME_CHARS = /[A-Za-z0-9_.$]/;

/** 一条 A1 引用的拆解；两处的 `$` 标记必须逐部分保留，否则平移会把绝对引用变成相对引用。 */
interface RefParts {
  isColAbsolute: boolean;
  colIndex: number;
  isRowAbsolute: boolean;
  rowIndex: number;
}

/**
 * 按轴平移一条公式里的引用。
 *
 * @param formula 公式体，不含前导 `=`。
 * @param axis 这次动的是行还是列。
 * @param from 0 基插入点/被删带起点；该位置及其之后的引用都要挪（`$` 也一起挪，见文件头）。
 * @param delta 位移量：插入为 `+count`，删除为 `-count`。
 * @param cut 删除时的被删区间；任何引用（含绝对）落在里面就整条拒绝——Excel 这时给的是 `#REF!`。
 * @returns `{ok:true, formula}`：没有引用受影响时返回原串；`{ok:false, reason}`：调用方必须整条跳过。
 */
export function shiftFormula(formula: string, axis: SheetAxis, from: number, delta: number, cut?: FormulaCut): FormulaShift {
  const pattern = new RegExp(TOKEN_PATTERN.source, "g");
  let out = "";
  let cursor = 0;
  for (const match of formula.matchAll(pattern)) {
    const index = match.index ?? 0;
    out += formula.slice(cursor, index);
    cursor = index + match[0].length;
    if (match[1] !== undefined) {
      // 字符串字面量里的 "A1" 是文字不是引用，原样带过
      out += match[0];
      continue;
    }
    if (match[2] !== undefined) {
      return { ok: false, reason: `公式里有我们不解析的写法（出现「${match[2]}」），位移会把引用弄错：${formula}` };
    }
    const nameProblem = checkNotAName(match[0], formula, cursor);
    if (nameProblem === SKIP_TOKEN) {
      // `LOG10(`、`ATAN2(` 这类函数名恰好长得像引用：紧跟 `(` 就是调用，原样带过
      out += match[0];
      continue;
    }
    if (nameProblem) return { ok: false, reason: `公式里有我们不解析的名字（${nameProblem}）：${formula}` };
    const parts = refParts(match);
    if (!parts) return { ok: false, reason: `公式里的引用认不出列号（「${match[0]}」）：${formula}` };
    if (cut && hitsCut(parts, axis, cut)) {
      return { ok: false, reason: `公式引用的格子正好在要删掉的区间里，Excel 会把它变成 #REF!：${formula}` };
    }
    const moved = moveOnAxis(parts, axis, from, delta);
    out += moved ? `${moved.isColAbsolute ? "$" : ""}${indexToColName(moved.colIndex)}${moved.isRowAbsolute ? "$" : ""}${moved.rowIndex + 1}` : match[0];
  }
  return { ok: true, formula: out + formula.slice(cursor) };
}

/** `checkNotAName` 的哨兵值：这段匹配是函数名的一部分，当普通文字带过。 */
const SKIP_TOKEN = Symbol("skip");

/**
 * 判断这次匹配是不是一段更长名字的一部分。
 *
 * @param token 匹配到的串。
 * @param formula 整条公式。
 * @param after 匹配结束后的下标。
 * @returns `SKIP_TOKEN`（函数名）、违规描述、或 null（确实是独立引用）。
 */
function checkNotAName(token: string, formula: string, after: number): typeof SKIP_TOKEN | string | null {
  const followed = formula[after];
  if (followed === "(") return SKIP_TOKEN;
  if (followed !== undefined && NAME_CHARS.test(followed)) return `「${token}」后面还跟着「${followed}」`;
  const previous = formula[after - token.length - 1];
  if (previous !== undefined && NAME_CHARS.test(previous)) return `「${token}」不是独立引用`;
  return null;
}

/** 从正则分组拆出引用的列号、行号与两处的 `$` 标记；列字母解析不出来时返回 null。 */
function refParts(match: RegExpMatchArray): RefParts | null {
  const colIndex = colNameToIndex(match[5] ?? "");
  if (colIndex === undefined) return null;
  return {
    isColAbsolute: match[4] === "$",
    colIndex,
    isRowAbsolute: match[6] === "$",
    rowIndex: Number(match[7]) - 1,
  };
}

/** 引用落在被删区间里吗（绝对引用也算——钉子钉在原位，原位没了就是没了）。 */
function hitsCut(parts: RefParts, axis: SheetAxis, cut: FormulaCut): boolean {
  const position = axis === "row" ? parts.rowIndex : parts.colIndex;
  return position >= cut.startIndex && position <= cut.endIndex;
}

/**
 * 算出引用该落到哪个新位置。
 *
 * @param parts 拆解后的引用。
 * @param axis 位移发生的轴。
 * @param from 插入点/被删带起点（0 基）。
 * @param delta 位移量。
 * @returns 新位置（带着原来的两个 `$` 标记，供逐部分重写）；在这次改动之前的引用不受影响，返回 null 原样带过。
 */
function moveOnAxis(parts: RefParts, axis: SheetAxis, from: number, delta: number): RefParts | null {
  const isRow = axis === "row";
  const position = isRow ? parts.rowIndex : parts.colIndex;
  // 插删挪的是格子本身，`$` 不参与判断（它只钉复制/填充），所以钉住的引用照样要跟着走。
  if (position < from) return null;
  const moved = position + delta;
  return isRow ? { ...parts, rowIndex: moved } : { ...parts, colIndex: moved };
}
