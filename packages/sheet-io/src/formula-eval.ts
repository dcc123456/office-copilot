/**
 * 公式子集求值：只为"写公式时同时算出缓存值"服务（spec 决策 D4）。
 *
 * 分工要先说清：`.xlsx` 里的 `<f>` 才是事实，Excel/WPS 打开后自己重算；本文件算出来的东西只进 `<v>`，
 * 用来让 app 里当场看得见数字。因此**宁可留空也不许猜**——认不出的函数、volatile 函数（`TODAY()`）、
 * 循环引用、跨表/外部引用一律返回 `{ok:false, reason}`，调用方照常写 `<f>` 但不写 `<v>`，
 * UI 走既有的"需 Excel 重算"提示（阶段 2 的 `.sheet-recalc-hint`）。
 * 白名单来自 spec D4：`SUM/AVERAGE/MIN/MAX/COUNT/COUNTIF/SUMIF/ROUND/ABS/IF/CONCAT` + 算术表达式；
 * 求值器**永不抛异常**（验收 B9：给模型的得是错误串，不是崩溃）。
 * 取数口径与 Excel 一致：区间里的空格、文本、布尔一律跳过；比较不区分大小写；`<>` 是不等于。
 */
import { formatA1, indexToColName, parseA1, parseRange, type A1Range } from "./address.js";
import type { CellScalarValue } from "./model.js";
import type { SheetPayload } from "./sheet-dto.js";

/** 一条公式算出来的值；`null` 是空格。 */
export type FormulaValue = CellScalarValue;

/** 求值结果：`ok:false` 时缓存值必须留空，`reason` 给人看也给日志。 */
export type FormulaEvalResult = { ok: true; value: FormulaValue } | { ok: false; reason: string };

/** 白名单外又"每次都变"的函数：写了缓存值就等于写一个过期数字，一律拒绝。 */
const VOLATILE_FUNCTIONS = new Set(["TODAY", "NOW", "RAND", "RANDBETWEEN", "OFFSET", "INDIRECT", "CELL", "INFO"]);

/** 认不出的构造：跨表 `!`、外部工作簿 `[1]`、数组常量 `{}`、已有的错误值 `#`。 */
const UNSUPPORTED_MARKER = /[![\]{}#]/;

/** 递归下降的游标与求值环境。 */
interface ParseState {
  tokens: FormulaToken[];
  position: number;
  sheet: SheetPayload;
  /** 正在求值的格（含起点自己）：命中即循环引用。 */
  evaluating: Set<string>;
}

type TokenKind = "number" | "text" | "name" | "ref" | "range" | "op" | "lparen" | "rparen" | "comma" | "percent";

/** 一个词法单元；`ref`/`range` 在扫描阶段就换算成 0 基坐标。 */
interface FormulaToken {
  kind: TokenKind;
  text: string;
  address?: { colIndex: number; rowIndex: number };
  range?: { start: { colIndex: number; rowIndex: number }; end: { colIndex: number; rowIndex: number } };
}

/** 表达式解析统一返回形状。 */
type Parsed = { ok: true; value: FormulaValue } | { ok: false; reason: string };

/** 函数实参：区间在这里就是一个值数组。 */
type ArgumentValue = FormulaValue | FormulaValue[];

/** 循环引用与算不出的依赖共用一条出口：凭它跨层冒到最外层，转成 reason 而不是被咽成 0。 */
class EvaluationBlocked extends Error {}

/**
 * 求值一条公式。**永不抛异常**：一切认不出的情况都返回 `{ok:false, reason}`。
 *
 * @param formula 公式体，不含前导 `=`（`<f>` 里的原样）。
 * @param sheet 求值所在的表载荷（引用按本表解释，跨表引用被拒绝）。
 * @param origin 公式所在格的大写 A1 地址；给出它才能认出 `A1 = SUM(A1)` 这类自我引用。
 * @returns `FormulaEvalResult`；`ok:false` 时调用方留空 `<v>`，不许拿近似值顶上。
 */
export function evaluateFormula(formula: string, sheet: SheetPayload, origin?: string): FormulaEvalResult {
  const evaluating = new Set(origin ? [origin] : []);
  try {
    return runFormula(formula, sheet, evaluating);
  } catch (error) {
    if (error instanceof EvaluationBlocked) return { ok: false, reason: error.message };
    return { ok: false, reason: "公式求值没走通" };
  }
}

/**
 * 列出一条公式在本表内引用到的地址区间（复用求值器自己的词法扫描，不再写第二份引用解析，§3.1）。
 *
 * @param formula 公式体，不含前导 `=`。
 * @returns 0 基的引用区间数组，可为空（纯常量公式）；含跨表 `!`、外部 `[1]`、数组常量 `{}`、已有错误值 `#`
 *   这些我们不解释的结构时返回 **null**——调用方据此判"依赖证明不了"，不许把它当成"没有依赖"。
 */
export function formulaReferencedRanges(formula: string): A1Range[] | null {
  const tokens = tokenize(formula);
  if (!tokens) return null;
  const ranges: A1Range[] = [];
  for (const token of tokens) {
    if (token.kind === "range" && token.range) ranges.push(token.range);
    else if (token.kind === "ref" && token.address) ranges.push({ start: token.address, end: token.address });
  }
  return ranges;
}

/**
 * 走一遍表达式层，不做异常兜底。
 *
 * 递归求值依赖格时也走这里：循环引用与算不出的依赖必须一路冒到 `evaluateFormula` 那唯一一处
 * `catch`，中间层自己吞掉的话就会把"算不出"变成"算出个 0"（spec D4 的红线）。
 *
 * @param formula 公式体。
 * @param sheet 表载荷。
 * @param evaluating 已在求值的格子集合（引用链上的每一格）。
 * @returns 求值结果；`evaluating` 会被就地增删。
 */
function runFormula(formula: string, sheet: SheetPayload, evaluating: Set<string>): FormulaEvalResult {
  const tokens = tokenize(formula);
  if (!tokens) return { ok: false, reason: "公式里有我们不解释的结构（跨表引用、数组常量等）" };
  const state: ParseState = { tokens, position: 0, sheet, evaluating };
  const value = parseExpression(state);
  if (!value.ok) return value;
  if (state.position !== tokens.length) return { ok: false, reason: "公式末尾有多余内容，没吃干净" };
  return { ok: true, value: value.value };
}

/**
 * 把公式切成 token。
 *
 * @param formula 公式体。
 * @returns token 序列；遇到上面列出的不支持结构返回 null（调用方留空缓存值）。
 */
function tokenize(formula: string): FormulaToken[] | null {
  // 只在没有引号包裹的部分找不支持的结构：`CONCAT(A1,"!")` 里的 `!` 是文本，不是跨表引用。
  if (UNSUPPORTED_MARKER.test(formula.replace(/"[^"]*"/g, '""'))) {
    return null;
  }
  const tokens: FormulaToken[] = [];
  let index = 0;
  while (index < formula.length) {
    const char = formula[index]!;
    if (/\s/.test(char)) {
      index += 1;
      continue;
    }
    if (char === '"') {
      const end = formula.indexOf('"', index + 1);
      if (end < 0) return null;
      tokens.push({ kind: "text", text: formula.slice(index + 1, end) });
      index = end + 1;
      continue;
    }
    const single = { "(": "lparen", ")": "rparen", ",": "comma", "%": "percent" } as const;
    if (char in single) {
      tokens.push({ kind: single[char as keyof typeof single], text: char });
      index += 1;
      continue;
    }
    const twoChar = formula.slice(index, index + 2);
    if (twoChar === "<>" || twoChar === "<=" || twoChar === ">=") {
      tokens.push({ kind: "op", text: twoChar });
      index += 2;
      continue;
    }
    if ("+-*/^&=<>".includes(char)) {
      tokens.push({ kind: "op", text: char });
      index += 1;
      continue;
    }
    const scanned = scanReferenceOrName(formula, index);
    if (!scanned.ok) return null;
    tokens.push(scanned.token);
    index += scanned.length;
  }
  return tokens;
}

/** 扫描的结果：一个 token 与吃掉的长度。 */
type ScanResult = { ok: true; token: FormulaToken; length: number } | { ok: false };

/**
 * 扫一段引用或名字：`A1`/`$A$1`/`A1:B3`/函数名/TRUE。
 *
 * 数字字面量给 `number` token；地址里的 `$` 是复制时的钉子，取数前在这一步就去掉了。
 *
 * @param formula 公式体。
 * @param start 起点下标。
 * @returns token 与长度；认不出返回 `ok:false`。
 */
function scanReferenceOrName(formula: string, start: number): ScanResult {
  const rest = formula.slice(start);
  const number = /^[0-9]*\.?[0-9]+(?:[eE][+-]?[0-9]+)?/.exec(rest);
  if (number) return { ok: true, length: number[0].length, token: { kind: "number", text: number[0] } };

  const rangeRef = /^(\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6}):(\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6})/.exec(rest);
  if (rangeRef) {
    const range = parseRange(`${stripDollar(rangeRef[1]!)}:${stripDollar(rangeRef[2]!)}`);
    if (!range) return { ok: false };
    return { ok: true, length: rangeRef[0].length, token: { kind: "range", text: rangeRef[0], range } };
  }
  const cellRef = /^\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6}/.exec(rest);
  if (cellRef) {
    const parsed = parseA1(stripDollar(cellRef[0]));
    if (!parsed) return { ok: false };
    return { ok: true, length: cellRef[0].length, token: { kind: "ref", text: cellRef[0], address: parsed } };
  }
  const name = /^[A-Za-z_][A-Za-z0-9_.]*/.exec(rest);
  if (name) return { ok: true, length: name[0].length, token: { kind: "name", text: name[0].toUpperCase() } };
  return { ok: false };
}

/** 去掉引用里的 `$`（`parseA1`/`parseRange` 只认干净地址）。 */
function stripDollar(ref: string): string {
  return ref.replace(/\$/g, "");
}

/** 看一眼当前 token 但不前进。 */
function peek(state: ParseState): FormulaToken | null {
  return state.tokens[state.position] ?? null;
}

/** 吃掉当前 token 并前进；越界返回 null。 */
function take(state: ParseState): FormulaToken | null {
  const token = state.tokens[state.position] ?? null;
  if (token) state.position += 1;
  return token;
}

/** 当前 token 是不是这个运算符（不前进，由调用方决定吃不吃）。 */
function peekIsOp(state: ParseState, text: string): boolean {
  const token = peek(state);
  return token?.kind === "op" && token.text === text;
}

/**
 * 比较层（最低优先级）：`=` 相等、`<>` 不等、`<` `<=` `>` `>=`。
 *
 * @param state 解析游标。
 * @returns 布尔值；没有比较运算符时原样返回下层结果。
 */
function parseExpression(state: ParseState): Parsed {
  const left = parseConcat(state);
  if (!left.ok) return left;
  const token = peek(state);
  if (!token || token.kind !== "op" || !["=", "<>", "<", "<=", ">", ">="].includes(token.text)) return left;
  state.position += 1;
  const right = parseConcat(state);
  if (!right.ok) return right;
  return { ok: true, value: compare(left.value, right.value, token.text) };
}

/** 连接层：`&` 把两边按文本拼起来（布尔按 Excel 显示成 TRUE/FALSE）。 */
function parseConcat(state: ParseState): Parsed {
  let left = parseAdditive(state);
  while (left.ok && peekIsOp(state, "&")) {
    state.position += 1;
    const right = parseAdditive(state);
    if (!right.ok) return right;
    left = { ok: true, value: asText(left.value) + asText(right.value) };
  }
  return left;
}

/** 加减层。 */
function parseAdditive(state: ParseState): Parsed {
  let left = parseMultiplicative(state);
  while (left.ok && (peekIsOp(state, "+") || peekIsOp(state, "-"))) {
    const operator = state.tokens[state.position]!.text;
    state.position += 1;
    const right = parseMultiplicative(state);
    if (!right.ok) return right;
    const numbers = numberPair(left.value, right.value);
    if (!numbers) return { ok: false, reason: "#VALUE!：把文本加进算式了" };
    left = { ok: true, value: operator === "+" ? numbers[0] + numbers[1] : numbers[0] - numbers[1] };
  }
  return left;
}

/** 乘除层（除零给 `#DIV/0!` 的理由，与 Excel 的显示一致）。 */
function parseMultiplicative(state: ParseState): Parsed {
  let left = parsePower(state);
  while (left.ok && (peekIsOp(state, "*") || peekIsOp(state, "/"))) {
    const operator = state.tokens[state.position]!.text;
    state.position += 1;
    const right = parsePower(state);
    if (!right.ok) return right;
    const numbers = numberPair(left.value, right.value);
    if (!numbers) return { ok: false, reason: "#VALUE!：把文本放进乘法了" };
    if (operator === "/") {
      if (numbers[1] === 0) return { ok: false, reason: "#DIV/0!：除数是 0" };
      left = { ok: true, value: numbers[0] / numbers[1] };
      continue;
    }
    left = { ok: true, value: numbers[0] * numbers[1] };
  }
  return left;
}

/** 幂层：Excel 里 `^` 右结合且优先级高于乘法。 */
function parsePower(state: ParseState): Parsed {
  const base = parseUnary(state);
  if (!base.ok) return base;
  if (!peekIsOp(state, "^")) return base;
  state.position += 1;
  const exponent = parsePower(state);
  if (!exponent.ok) return exponent;
  const numbers = numberPair(base.value, exponent.value);
  if (!numbers) return { ok: false, reason: "#VALUE!：幂里有文本" };
  return { ok: true, value: numbers[0] ** numbers[1] };
}

/** 一元正负号。 */
function parseUnary(state: ParseState): Parsed {
  if (peekIsOp(state, "-")) {
    state.position += 1;
    const inner = parseUnary(state);
    if (!inner.ok) return inner;
    const numeric = asNumber(inner.value);
    return numeric === null ? { ok: false, reason: "#VALUE!：对文本取负" } : { ok: true, value: -numeric };
  }
  if (peekIsOp(state, "+")) state.position += 1;
  return parsePostfix(state);
}

/** 百分号后缀：`50%` 就是 `0.5`（Excel 的写法）。 */
function parsePostfix(state: ParseState): Parsed {
  let value = parsePrimary(state);
  while (value.ok && peek(state)?.kind === "percent") {
    state.position += 1;
    const numeric = asNumber(value.value);
    if (numeric === null) return { ok: false, reason: "#VALUE!：对文本取百分比" };
    value = { ok: true, value: numeric / 100 };
  }
  return value;
}

/**
 * 最小单元：数字 / 文本 / 布尔 / 引用 / 区间 / 函数 / 括号。
 *
 * @param state 解析游标。
 * @returns 值；区间出现在标量位置时给 `#VALUE!`。
 */
function parsePrimary(state: ParseState): Parsed {
  const token = take(state);
  if (!token) return { ok: false, reason: "公式不完整，后面没内容了" };
  switch (token.kind) {
    case "number":
      return { ok: true, value: Number(token.text) };
    case "text":
      return { ok: true, value: token.text };
    case "name":
      if (token.text === "TRUE") return { ok: true, value: true };
      if (token.text === "FALSE") return { ok: true, value: false };
      return parseFunction(state, token.text);
    case "ref":
      return { ok: true, value: readCell(state, token.address!) };
    case "range":
      return { ok: false, reason: "#VALUE!：整个区间放到了要单个数字的位置" };
    case "lparen": {
      const inner = parseExpression(state);
      if (!inner.ok) return inner;
      if (take(state)?.kind !== "rparen") return { ok: false, reason: "括号没闭合" };
      return inner;
    }
    default:
      return { ok: false, reason: `认不出的写法：${token.text}` };
  }
}

/**
 * 函数调用：只认白名单，其余一律 `{ok:false}`，让调用方留空缓存值。
 *
 * @param state 解析游标。
 * @param name 已大写的函数名。
 * @returns 求值结果。
 */
function parseFunction(state: ParseState, name: string): Parsed {
  if (peek(state)?.kind !== "lparen") return { ok: false, reason: `认不出的名字：${name}` };
  state.position += 1;
  const args = parseArguments(state);
  if (!args.ok) return args;
  if (take(state)?.kind !== "rparen") return { ok: false, reason: "函数括号没闭合" };
  if (VOLATILE_FUNCTIONS.has(name)) return { ok: false, reason: `${name}() 每次都变，不写缓存值` };
  const called = callFunction(name, args.value);
  return called.ok ? { ok: true, value: called.value } : { ok: false, reason: called.reason };
}

/** 实参列表；空括号得到空数组。 */
function parseArguments(state: ParseState): { ok: true; value: ArgumentValue[] } | { ok: false; reason: string } {
  if (peek(state)?.kind === "rparen") return { ok: true, value: [] };
  const args: ArgumentValue[] = [];
  for (;;) {
    const arg = parseArgument(state);
    if (!arg.ok) return arg;
    args.push(arg.value);
    if (peek(state)?.kind === "comma") {
      state.position += 1;
      continue;
    }
    return { ok: true, value: args };
  }
}

/** 一个实参：区间整体是一个值数组，其它走完整表达式层。 */
function parseArgument(state: ParseState): { ok: true; value: ArgumentValue } | { ok: false; reason: string } {
  const token = peek(state);
  if (token?.kind === "range") {
    state.position += 1;
    return { ok: true, value: readRange(state, token.range!) };
  }
  return parseExpression(state);
}

/** 函数返回形状。 */
type Called = { ok: true; value: FormulaValue } | { ok: false; reason: string };

/**
 * 调一个白名单函数。
 *
 * @param name 函数名（大写）。
 * @param args 实参（区间以数组形式进来）。
 * @returns 值或拒绝理由。
 */
function callFunction(name: string, args: ArgumentValue[]): Called {
  const flat = args.flatMap((arg) => (Array.isArray(arg) ? arg : [arg]));
  switch (name) {
    case "SUM":
      return { ok: true, value: numbersOf(flat).reduce((acc, value) => acc + value, 0) };
    case "AVERAGE": {
      const numbers = numbersOf(flat);
      return numbers.length === 0
        ? { ok: false, reason: "#DIV/0!：AVERAGE 没有一个可用数字" }
        : { ok: true, value: numbers.reduce((acc, value) => acc + value, 0) / numbers.length };
    }
    case "MIN":
      return { ok: true, value: numbersOf(flat).length === 0 ? 0 : Math.min(...numbersOf(flat)) };
    case "MAX":
      return { ok: true, value: numbersOf(flat).length === 0 ? 0 : Math.max(...numbersOf(flat)) };
    case "COUNT":
      return { ok: true, value: numbersOf(flat).length };
    case "ABS": {
      const numeric = asNumber(scalar(args[0] ?? null));
      return numeric === null ? { ok: false, reason: "ABS 只吃数字" } : { ok: true, value: Math.abs(numeric) };
    }
    case "ROUND": {
      const numeric = asNumber(scalar(args[0] ?? null));
      if (numeric === null) return { ok: false, reason: "ROUND 只吃数字" };
      return { ok: true, value: roundAwayFromZero(numeric, asNumber(scalar(args[1] ?? 0)) ?? 0) };
    }
    case "IF": {
      if (args.length < 2) return { ok: false, reason: "IF 至少要有条件和结果" };
      if (truthy(args[0]!)) return { ok: true, value: scalar(args[1]!) };
      return { ok: true, value: args.length >= 3 ? scalar(args[2]!) : false };
    }
    case "CONCAT":
    case "CONCATENATE":
      return { ok: true, value: flat.map(asText).join("") };
    case "COUNTIF":
      return countIf(args);
    case "SUMIF":
      return sumIf(args);
    default:
      return { ok: false, reason: `还没实现 ${name}()，交给 Excel 重算` };
  }
}

/** `COUNTIF(区间, 条件)`。 */
function countIf(args: ArgumentValue[]): Called {
  if (args.length < 2) return { ok: false, reason: "COUNTIF 要有区间和条件" };
  const criteria = toCriteria(args[1]!);
  if (!criteria) return { ok: false, reason: "COUNTIF 的条件认不出" };
  return { ok: true, value: valuesOf(args[0]!).filter(criteria).length };
}

/** `SUMIF(条件区间, 条件, [求和区间])`：第三个参数缺省时对条件区间自身求和。 */
function sumIf(args: ArgumentValue[]): Called {
  if (args.length < 2) return { ok: false, reason: "SUMIF 要有条件区间和条件" };
  const criteria = toCriteria(args[1]!);
  if (!criteria) return { ok: false, reason: "SUMIF 的条件认不出" };
  const conditionValues = valuesOf(args[0]!);
  const sumValues = args.length >= 3 ? valuesOf(args[2]!) : conditionValues;
  let sum = 0;
  conditionValues.forEach((value, position) => {
    if (!criteria(value)) return;
    const addend = sumValues[position];
    if (typeof addend === "number") sum += addend;
  });
  return { ok: true, value: sum };
}

/** 取值数组（区间原样、标量套一层）。 */
function valuesOf(arg: ArgumentValue): FormulaValue[] {
  return Array.isArray(arg) ? arg : [arg];
}

/** 取标量位置的值：万一实参是区间，只取第一个（Excel 这时用左上角）。 */
function scalar(arg: ArgumentValue): FormulaValue {
  return Array.isArray(arg) ? arg[0] ?? null : arg;
}

/**
 * 统计函数用的一组数字：区间里的空格、文本、布尔一律跳过（Excel 的 `SUM`/`COUNT` 就这么算）。
 * "看着像数字的文本"也算文本——它来自格子，Excel 凭格式判断类型，不因为我们能解析出数字就把它计进来。
 * 需要强转的是算式里直接写出来的字面量（`A2+"15"`），那条路走 `asNumber`，不走这里。
 *
 * @param values 展平后的实参。
 * @returns 数字数组。
 */
function numbersOf(values: FormulaValue[]): number[] {
  return values.filter((value): value is number => typeof value === "number");
}

/** 纯数字文本判定（不含千分位——那属于格式，不是值）。 */
function isNumericText(text: string): boolean {
  return /^-?\d+(?:\.\d+)?(?:[eE][+-]?\d+)?$/.test(text.trim());
}

/**
 * 生成条件判定：支持 `>100`、`<=2`、`<>x`、`=文本`、纯文本（等于）与纯数字（相等）。
 *
 * 除了 `COUNTIF`/`SUMIF` 求值，`analysis/sheet-table` 的按条件定位也用它，
 * 好让"工具层筛出来的行"与"公式算出来的行"是同一套口径（验收 B10）。
 *
 * @param raw 条件实参。
 * @returns 判定函数；写不出可比较的式子返回 null。
 */
export function toCriteria(raw: ArgumentValue): ((value: FormulaValue) => boolean) | null {
  const value = scalar(raw);
  if (typeof value === "number") return (candidate) => compare(candidate, value, "=");
  if (typeof value !== "string") return null;
  const match = /^(<=|>=|<>|=|<|>)(.*)$/.exec(value.trim());
  const operator = match ? match[1]! : "=";
  const text = (match ? match[2]! : value).trim();
  const expected: number | string = isNumericText(text) ? Number(text) : text;
  if (operator === "<>") return (candidate) => !compare(candidate, expected, "=");
  return (candidate) => compare(candidate, expected, operator);
}

/**
 * Excel 风格的比较：数字按数值、文本忽略大小写、空格与 0 相等。
 *
 * @param left 左侧值。
 * @param right 右侧值。
 * @param operator 比较符原文。
 * @returns 布尔结果。
 */
function compare(left: FormulaValue, right: FormulaValue, operator: string): boolean {
  const a = asNumber(left);
  const b = asNumber(right);
  const order = a !== null && b !== null ? a - b : asText(left).localeCompare(asText(right), undefined, { sensitivity: "base" });
  switch (operator) {
    case "=":
      return order === 0;
    case "<>":
      return order !== 0;
    case "<":
      return order < 0;
    case "<=":
      return order <= 0;
    case ">":
      return order > 0;
    default:
      return order >= 0;
  }
}

/** IF 的条件位：空格与 0 为假，文本只认 "TRUE" 为真（Excel 的隐式转换就这么窄）。 */
function truthy(arg: ArgumentValue): boolean {
  const value = scalar(arg);
  if (value === null || value === "" || value === false) return false;
  if (value === true) return true;
  if (typeof value === "number") return value !== 0;
  return value.toUpperCase() === "TRUE";
}

/**
 * 任意值 → 数字。
 *
 * @param value 任意值。
 * @returns 数字；空格按 Excel 当 0，文本只有纯数字才算，布尔返回 null（不参与算术）。
 */
function asNumber(value: FormulaValue): number | null {
  if (typeof value === "number") return value;
  if (value === null || value === "") return 0;
  if (typeof value === "boolean") return null;
  return isNumericText(value) ? Number(value.trim()) : null;
}

/** 任意值 → 文本（`&` 与 `CONCAT` 用；布尔显示成 TRUE/FALSE，空格是空串）。 */
function asText(value: FormulaValue): string {
  if (value === null || value === undefined) return "";
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  return String(value);
}

/** 两个操作数都能化成数字时给出来，否则 null（调用方据此报 `#VALUE!`）。 */
function numberPair(left: FormulaValue, right: FormulaValue): [number, number] | null {
  const a = asNumber(left);
  const b = asNumber(right);
  return a !== null && b !== null ? [a, b] : null;
}

/**
 * Excel 的 `ROUND`：离零方向四舍五入（不是银行家舍入）。
 *
 * @param value 待舍数字。
 * @param digits 小数位数，可为负（`ROUND(1234,-2)` → 1200）。
 * @returns 舍入结果，尾差收进 15 位有效数字。
 */
function roundAwayFromZero(value: number, digits: number): number {
  const factor = 10 ** digits;
  const rounded = Math.sign(value * factor) * Math.round(Math.abs(value * factor));
  return Number((rounded / factor).toPrecision(15));
}

/**
 * 读一格的值。
 *
 * @param state 解析游标（带表与已求值集合）。
 * @param address 0 基坐标。
 * @returns 值；空格返回 null。带公式的格递归求值，命中正在求值的格即循环引用。
 */
function readCell(state: ParseState, address: { colIndex: number; rowIndex: number }): FormulaValue {
  const key = formatA1(address) || `${indexToColName(address.colIndex)}${address.rowIndex + 1}`;
  // 正在求值的那一格被自己（或引用链上的下一格）指到，就是循环引用，Excel 也这么判。
  if (state.evaluating.has(key)) throw new EvaluationBlocked("循环引用，交给 Excel 重算");
  const cell = state.sheet.cells[key] ?? null;
  if (!cell || cell.kind === "blank") return null;
  if (cell.formula) {
    state.evaluating.add(key);
    try {
      const result = runFormula(cell.formula, state.sheet, state.evaluating);
      if (result.ok) return result.value;
      // 依赖格算不出等于这条也算不出：拿过期缓存值顶上就是写一个假数字（spec D4）。
      throw new EvaluationBlocked(`${key} 算不出来：${result.reason}`);
    } finally {
      state.evaluating.delete(key);
    }
  }
  return cell.value;
}

/**
 * 读一个区间的值（按行优先展开，与 `SUM(A1:B2)` 的取数顺序一致）。
 *
 * @param state 解析游标。
 * @param range 已归一化的矩形（`start` 在左上）。
 * @returns 值数组；跨表引用在切词阶段就被拒绝，到不了这里。
 */
function readRange(state: ParseState, range: { start: { colIndex: number; rowIndex: number }; end: { colIndex: number; rowIndex: number } }): FormulaValue[] {
  const values: FormulaValue[] = [];
  for (let rowIndex = range.start.rowIndex; rowIndex <= range.end.rowIndex; rowIndex += 1) {
    for (let colIndex = range.start.colIndex; colIndex <= range.end.colIndex; colIndex += 1) values.push(readCell(state, { colIndex, rowIndex }));
  }
  return values;
}
