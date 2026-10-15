/**
 * PDF 字符串字面量的两种书写形态（ISO 32000-1 §7.4.3）与它们的编解码。
 *
 * 内容流里的文字一律以这两种形态之一出现：literal-string `( ... )` 与 hex-string `< ... >`。
 * 本机实测：我们自产的 PDF（pdf-lib 那条通路）用的是 **hex 形态**，而 Word/Chrome 导出的常见是 literal 形态，
 * 所以定位与替换必须两种都认（`docs/office-parity-spec.md` D8 的 W1 那条"按 hex/字面量两种串形态定位"）。
 */

/** 一条字符串操作数在解码后的内容流里的位置与信息。 */
export interface StreamOperand {
  /** 书写形态。 */
  form: "literal" | "hex";
  /** 起始下标（含定界符），单位是字节（内容流按 latin1 处理，一个字符就是一个字节）。 */
  start: number;
  /** 结束下标（不含），即定界符 `)` 或 `>` 的下一位。 */
  end: number;
  /** 解出来的原始字节（0-255），编码前的真值。 */
  bytes: number[];
}

/** literal 形态里的转义（§7.4.3 表 4）。 */
/** 按下标取一个字节字符；越界给空串（内容流与串字面量都按 latin1 处理，一个字符就是一个字节）。 */
const at = (text: string, index: number): string => text[index] ?? '';

const LITERAL_ESCAPES: Record<string, number> = {
  n: 10,
  r: 13,
  t: 9,
  b: 8,
  f: 12,
  "(": 40,
  ")": 41,
  "\\": 92,
};

/**
 * 解一个 literal-string 操作数 `( ... )`。
 *
 * @param src 解码后的内容流（latin1 字符串）。
 * @param start `(` 所在的下标。
 * @returns 操作数信息；`end` 指向 `)` 之后一位。定界不配对时按"读到流尾"处理，不抛异常（读侧宽容，写侧才严格）。
 */
export function readLiteralOperand(src: string, start: number): StreamOperand {
  const bytes: number[] = [];
  let index = start + 1;
  let depth = 1;
  while (index < src.length) {
    const char = at(src, index);
    if (char === "\\") {
      const next = at(src, index + 1) ?? "";
      if (next >= "0" && next <= "7") {
        let octal = "";
        index += 1;
        while (octal.length < 3 && at(src, index) >= "0" && at(src, index) <= "7") {
          octal += at(src, index);
          index += 1;
        }
        bytes.push(parseInt(octal, 8) & 255);
        continue;
      }
      if (next === "\n") {
        index += 2;
        continue;
      }
      if (next === "\r") {
        index += at(src, index + 2) === "\n" ? 3 : 2;
        continue;
      }
      bytes.push(LITERAL_ESCAPES[next] ?? next.charCodeAt(0));
      index += 2;
      continue;
    }
    if (char === "(") depth += 1;
    if (char === ")") {
      depth -= 1;
      if (depth === 0) return { form: "literal", start, end: index + 1, bytes };
    }
    bytes.push(char.charCodeAt(0) & 255);
    index += 1;
  }
  return { form: "literal", start, end: index, bytes };
}

/**
 * 解一个 hex-string 操作数 `< ... >`。
 *
 * @param src 解码后的内容流。
 * @param start `<` 所在的下标（调用方须已排除 `<<` 字典开头）。
 * @returns 操作数信息；字节数为奇数个十六进制位时按规范补一个 0（§7.4.3）。
 */
export function readHexOperand(src: string, start: number): StreamOperand {
  const bytes: number[] = [];
  let nibble = "";
  let index = start + 1;
  for (; index < src.length; index += 1) {
    const char = at(src, index);
    if (char === ">") break;
    if (/\s/.test(char)) continue;
    nibble += char;
    if (nibble.length === 2) {
      bytes.push(parseInt(nibble, 16) & 255);
      nibble = "";
    }
  }
  if (nibble.length === 1) bytes.push(parseInt(nibble + "0", 16) & 255);
  return { form: "hex", start, end: index + 1 <= src.length ? index + 1 : src.length, bytes };
}

/**
 * 把字节写成 literal 形态的 `( ... )`。
 *
 * @param bytes 要写出的字节（0-255）。
 * @returns 完整的操作数字符串；括号、反斜杠与非可打印字节走八进制转义。
 */
export function encodeLiteralOperand(bytes: number[]): string {
  let out = "(";
  for (const byte of bytes) {
    if (byte === 40 || byte === 41 || byte === 92) out += "\\" + String.fromCharCode(byte);
    else if (byte >= 32 && byte <= 126) out += String.fromCharCode(byte);
    else out += "\\" + byte.toString(8).padStart(3, "0");
  }
  return out + ")";
}

/**
 * 把字节写成 hex 形态的 `< ... >`（大写十六进制，无空格）。
 *
 * @param bytes 要写出的字节（0-255）。
 * @returns 完整的操作数字符串。
 */
export function encodeHexOperand(bytes: number[]): string {
  return "<" + bytes.map((byte) => byte.toString(16).toUpperCase().padStart(2, "0")).join("") + ">";
}
