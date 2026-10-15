/**
 * 内容流的最小解析：只认"显示文字"与"进 Form XObject"这两类算子，其余一概跳过。
 *
 * 为什么不用一个完整的 PDF 解释器：W1 这一档要做的是"同框同字体同长度替换"（spec D8），
 * 需要的是**文字操作数的字节位置**与**当时生效的字体资源名**，不需要排版状态。算子语法取自
 * ISO 32000-1 §7.5.3（词法）与 §9.4.3（文字算子 `Tj TJ ' "` 与 `Tf`、图形 XObject 的 `Do`）。
 */
import { readHexOperand, readLiteralOperand, type StreamOperand } from "./strings.js";

/** 一条显示文字的算子与它的操作数。 */
export interface TextRun {
  /** 算子名。`'` 与 `"` 等价于先换行再显示，`TJ` 的操作数来自数组的某一个元素。 */
  operator: "Tj" | "TJ" | "'" | '"';
  /** 操作数在解码后内容流里的位置与字节。 */
  operand: StreamOperand;
  /** 当时生效的字体资源名（最近一次 `Tf` 给的那个 `/Name`）；没出现过 `Tf` 时为 null。 */
  fontName: string | null;
}

/** 一次 `Do` 调用：调用方据此递归进 Form XObject。 */
export interface XObjectCall {
  /** 资源名（`/XObject` 字典里的键）。 */
  name: string;
  /** `Do` 算子在内容流里的下标，仅用于定位。 */
  at: number;
  /** 当时生效的字体资源名；Form 内部会另起一份图形状态，所以递归时不带走。 */
  fontName: string | null;
}

/** 一个内容流扫出来的全部结果。 */
export interface ContentScan {
  /** 显示文字的算子，按内容流顺序。 */
  runs: TextRun[];
  /** 调用图形 XObject 的位置。 */
  xobjects: XObjectCall[];
}

const WHITESPACE = new Set(["\u0000", "\t", "\n", "\f", "\r", " "]);
const DELIMITERS = new Set(["(", ")", "<", ">", "[", "]", "{", "}", "/", "%"]);

/** 按下标取一个字节字符；越界给空串（内容流按 latin1 处理，一个字符就是一个字节）。 */
const at = (text: string, index: number): string => text[index] ?? '';

/**
 * 扫一个解码后的内容流。
 *
 * @param src 内容流文本（latin1 字符串：一个字符就是一个字节，下标即字节位置）。
 * @returns 文字操作数与 `Do` 调用；语法残破时按"跳过一位"继续，不抛异常（读侧宽容，写侧才严格）。
 */
export function scanContentStream(src: string): ContentScan {
  const runs: TextRun[] = [];
  const xobjects: XObjectCall[] = [];
  let index = 0;
  let fontName: string | null = null;
  let pendingName: string | null = null;
  let pendingString: StreamOperand | null = null;
  let pendingArray: StreamOperand[] | null = null;

  while (index < src.length) {
    const char = at(src, index);
    if (WHITESPACE.has(char)) {
      index += 1;
      continue;
    }
    if (char === "%") {
      while (index < src.length && at(src, index) !== "\n" && at(src, index) !== "\r") index += 1;
      continue;
    }
    if (char === "(") {
      pendingString = readLiteralOperand(src, index);
      pendingArray = null;
      index = pendingString.end;
      continue;
    }
    if (char === "<" && at(src, index + 1) === "<") {
      index = skipDictionary(src, index);
      pendingString = null;
      pendingArray = null;
      continue;
    }
    if (char === "<") {
      pendingString = readHexOperand(src, index);
      pendingArray = null;
      index = pendingString.end;
      continue;
    }
    if (char === "[") {
      pendingArray = readArrayOperands(src, index);
      pendingString = null;
      while (index < src.length && at(src, index) !== "]") index += 1;
      index += 1;
      continue;
    }
    if (char === "]" || char === "{".charAt(0) || char === "}" || char === ">") {
      index += 1;
      continue;
    }
    if (char === "/") {
      pendingName = readName(src, index).name;
      index = readName(src, index).end;
      continue;
    }
    if (isNumberStart(char)) {
      index = skipNumber(src, index);
      continue;
    }
    const word = readKeyword(src, index);
    index = word.end;
    if (word.value === "Tf" && pendingName) fontName = pendingName;
    else if (word.value === "Do" && pendingName) xobjects.push({ name: pendingName, at: word.end, fontName });
    else if ((word.value === "Tj" || word.value === "'" || word.value === '"') && pendingString) {
      runs.push({ operator: word.value, operand: pendingString, fontName });
    } else if (word.value === "TJ" && pendingArray) {
      for (const operand of pendingArray) runs.push({ operator: "TJ", operand, fontName });
    }
    pendingName = null;
    pendingString = null;
    pendingArray = null;
  }
  return { runs, xobjects };
}

/**
 * 读一个 PDF 名字（`/Name`），并解开 `#XX` 转义。
 *
 * @param src 内容流文本。
 * @param start `/` 所在下标。
 * @returns 名字（已解开转义）与结束下标。
 */
function readName(src: string, start: number): { name: string; end: number } {
  let index = start + 1;
  let raw = "";
  while (index < src.length && !WHITESPACE.has(at(src, index)) && !DELIMITERS.has(at(src, index))) {
    raw += at(src, index);
    index += 1;
  }
  return { name: raw.replace(/#([0-9a-fA-F]{2})/g, (_, hex) => String.fromCharCode(parseInt(hex, 16))), end: index };
}

/** 跳过 `<< ... >>` 字典，返回 `>>` 之后一位的下标。 */
function skipDictionary(src: string, start: number): number {
  let index = start + 2;
  let depth = 1;
  while (index < src.length && depth > 0) {
    if (at(src, index) === "<" && at(src, index + 1) === "<") {
      depth += 1;
      index += 2;
      continue;
    }
    if (at(src, index) === ">" && at(src, index + 1) === ">") {
      depth -= 1;
      index += 2;
      continue;
    }
    if (at(src, index) === "(") {
      index = readLiteralOperand(src, index).end;
      continue;
    }
    index += 1;
  }
  return index;
}

/**
 * 收集 `[ ... ]` 数组里的字符串操作数（`TJ` 用得上，其余元素是字距数字，忽略）。
 *
 * @param src 内容流文本。
 * @param start `[` 所在下标。
 * @returns 数组内的字符串操作数；嵌套数组里的操作数不单独成 run（真实产物里极罕见，且改它会破坏字距结构）。
 */
function readArrayOperands(src: string, start: number): StreamOperand[] {
  const operands: StreamOperand[] = [];
  let index = start + 1;
  let depth = 1;
  while (index < src.length && depth > 0) {
    const char = at(src, index);
    if (char === "[") {
      depth += 1;
      index += 1;
      continue;
    }
    if (char === "]") {
      depth -= 1;
      index += 1;
      continue;
    }
    if (char === "(") {
      const operand = readLiteralOperand(src, index);
      if (depth === 1) operands.push(operand);
      index = operand.end;
      continue;
    }
    if (char === "<" && at(src, index + 1) !== "<") {
      const operand = readHexOperand(src, index);
      if (depth === 1) operands.push(operand);
      index = operand.end;
      continue;
    }
    index += 1;
  }
  return operands;
}

/** 数字 token 是否以此字符开头（§7.4.6 文法：可带符号与小数点）。 */
function isNumberStart(char: string): boolean {
  return (char >= "0" && char <= "9") || char === "." || char === "+" || char === "-";
}

/** 跳过一个数字 token，返回结束下标。 */
function skipNumber(src: string, start: number): number {
  let index = start;
  if (at(src, index) === "+" || at(src, index) === "-") index += 1;
  while (index < src.length && /[0-9.]/.test(at(src, index))) index += 1;
  return index === start ? start + 1 : index;
}

/** 读一个算子/关键字 token（连续的非常规字符）。 */
function readKeyword(src: string, start: number): { value: string; end: number } {
  let index = start;
  while (index < src.length && !WHITESPACE.has(at(src, index)) && !DELIMITERS.has(at(src, index))) index += 1;
  return { value: src.slice(start, index), end: index };
}
