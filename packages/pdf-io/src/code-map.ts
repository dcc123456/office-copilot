/**
 * 字符码 ↔ 文字：标准单字节编码表、`/ToUnicode` CMap 解析、按码长切字节。
 *
 * 为什么要 `/ToUnicode`：ISO 32000-1 §9.10.2 明说"没有办法确定字符码代表什么字符"，§9.10.3 的
 * `/ToUnicode` 是**可选**的且经常缺或错（spec D8 的 P11 实测：2 个嵌入字体只有 1 个有）。
 * 所以本模块给不出映射时一律返回 null，调用方据此拒绝改动并给出结构化理由——**禁止猜**。
 */
import { encodeHexOperand, encodeLiteralOperand, type StreamOperand } from "./strings.js";

/** 一个字体可用的码↔文字映射。 */
export interface CodeMap {
  /** 一个字符码占几个字节：简单字体 1，`Identity-H` 复合字体 2；0 表示本引擎不认这种编码。 */
  codeBytes: number;
  /** 码 → 文字；表里没有这个码时 null。 */
  decode(code: number): string | null;
  /** 文字（一个码点，或 CMap 里给出的连字串）→ 码；子集里没有这个字时 null。 */
  encode(char: string): number | null;
}

/**
 * `WinAnsiEncoding` 高位区（0x80-0xFF），按下标即码位减 128，`\u0000` 是表里的未定义位。
 *
 * 取的是 Windows-1252 码表——Adobe 的 WinAnsi 与它的差别只在 0x80 以下的控制区，高位区逐位相同；
 * 本机用 python 的 cp1252 解码逐位生成，未定义位是 129/141/143/144/157（手抄易错，故生成后粘进来）。
 */
const WIN_ANSI_HIGH = "€\u0000‚ƒ„…†‡ˆ‰Š‹Œ\u0000Ž\u0000\u0000\u0000‘’“”•–—˜™š›œ\u0000žŸ\u00a0¡¢£¤¥¦§¨©ª«¬\u00ad®¯°±²³´µ¶·¸¹º»¼½¾¿ÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖ×ØÙÚÛÜÝÞßàáâãäåæçèéêëìíîïðñòóôõö÷øùúûüýþÿ";

/** `T*`、`Tab` 这些控制码在两表里的固定含义（§9.6.4.2 表 24）。 */
const WIN_ANSI_CONTROL: Record<number, string> = { 8: "\b", 9: "\t", 10: "\n", 12: "\f", 13: "\r" };

/**
 * 单字节标准编码表（`codeBytes` 恒为 1）。
 *
 * @param name 编码名；本引擎只认 `WinAnsiEncoding`（Chrome / Word / pdf-lib 三条产物实测都走这条），
 *   其余标准编码（PDFDoc、MacRoman、带 `/Differences` 的自定义表）不给 Unicode 就不猜，返回 null。
 * @returns 码表，或不认识时的 null。
 */
export function standardEncodingMap(name: string): CodeMap | null {
  if (name !== "WinAnsiEncoding") return null;
  const toCode = new Map<string, number>();
  for (let code = 32; code < 128; code += 1) toCode.set(String.fromCharCode(code), code);
  for (let index = 0; index < WIN_ANSI_HIGH.length; index += 1) {
    const char = (WIN_ANSI_HIGH[index] ?? '');
    if (char !== "\u0000") toCode.set(char, 128 + index);
  }
  return {
    codeBytes: 1,
    decode: (code) => {
      if (code < 32) return WIN_ANSI_CONTROL[code] ?? null;
      if (code < 128) return String.fromCharCode(code);
      const char = (WIN_ANSI_HIGH[code - 128] ?? '');
      return char && char !== "\u0000" ? char : null;
    },
    encode: (char) => toCode.get(char) ?? null,
  };
}

/**
 * 解一份 `/ToUnicode` CMap（§9.10.3 的 `bfchar` 与 `bfrange` 两种，含数组形式的目的码）。
 *
 * @param cmap CMap 流解码后的文本。
 * @returns 码表；一条都解不出来时 null（调用方据此拒绝编辑）。
 */
export function parseToUnicodeCMap(cmap: string): CodeMap | null {
  const map = new Map<number, string>();
  const toCode = new Map<string, number>();
  let codeBytes = 0;
  for (const [, rawFirst] of cmap.matchAll(/begincodespacerange\s*<([0-9a-fA-F]+)>/g)) {
    codeBytes = Math.max(codeBytes, Math.ceil((rawFirst ?? '').length / 2));
  }
  for (const block of matchBlocks(cmap, "beginbfchar", "endbfchar")) {
    for (const [, codeHex, dstHex] of block.matchAll(/<([0-9a-fA-F]+)>\s*<([0-9a-fA-F]+)>/g)) {
      record(map, toCode, parseInt(codeHex ?? '', 16), hexToUtf16(dstHex ?? ''));
    }
  }
  for (const block of matchBlocks(cmap, "beginbfrange", "endbfrange")) {
    for (const [, fromHex, toHex, dst] of block.matchAll(/<([0-9a-fA-F]+)>\s*<([0-9a-fA-F]+)>\s*(<[0-9a-fA-F]+>|\[[^\]]*\])/g)) {
      const from = parseInt(fromHex ?? '', 16);
      const count = parseInt(toHex ?? '', 16) - from;
      if (count < 0 || count > 65535) continue;
      const destinations = (dst ?? '').startsWith("[")
        ? [...(dst ?? '').slice(1, -1).matchAll(/<([0-9a-fA-F]+)>/g)].map((entry) => hexToUtf16(entry[1] ?? ''))
        : Array.from({ length: count + 1 }, (_, offset) => hexToUtf16((parseInt((dst ?? '').slice(1, -1), 16) + offset).toString(16)));
      destinations.forEach((text, offset) => record(map, toCode, from + offset, text));
    }
  }
  if (map.size === 0) return null;
  if (!codeBytes) codeBytes = Math.max(...[...map.keys()].map((code) => (code > 255 ? 2 : 1)));
  return {
    codeBytes,
    decode: (code) => map.get(code) ?? null,
    encode: (char) => toCode.get(char) ?? null,
  };
}

/**
 * 把操作数字节切成字符码，再解成文字。
 *
 * @param operand 内容流里的一条字符串操作数。
 * @param codes 码表。
 * @returns 解出的文字；任一码查不到时 null（不返回半截文字，避免把残缺内容当成全文交给模型或用户）。
 */
export function decodeOperandText(operand: StreamOperand, codes: CodeMap | null): string | null {
  if (!codes || !codes.codeBytes) return null;
  if (operand.bytes.length % codes.codeBytes !== 0) return null;
  let text = "";
  for (let index = 0; index < operand.bytes.length; index += codes.codeBytes) {
    let code = 0;
    for (let offset = 0; offset < codes.codeBytes; offset += 1) code = code * 256 + (operand.bytes[index + offset] ?? 0);
    const char = codes.decode(code);
    if (char === null) return null;
    text += char;
  }
  return text;
}

/**
 * 把新文字编回字节，并写出与原操作数**同长**的操作数 token。
 *
 * PDF 没有回流这回事（spec D8 W1 的硬判据），字节数一变就装不回原来的位置：
 * 这里先按原字节数编码，再在 literal / hex 两种形态里挑一个 token 长度与原来完全相同的。
 *
 * @param text 要写出的文字。
 * @param operand 原操作数的两个尺寸：`encodedBytes` 是原串字节数，`tokenLength` 是原 token 的字符数（含定界符）。
 * @param codes 码表。
 * @returns `{ token }`，或结构化的失败码（`no-encoding` 字体认不出 / `char-outside-subset` 子集缺字 /
 *   `length-mismatch` 装不回去）。文案由调用方组装（`AGENTS.md` §5.7）。
 */
export function encodeOperandToken(
  text: string,
  operand: { encodedBytes: number; tokenLength: number },
  codes: CodeMap | null,
): { token: string } | { failure: "no-encoding" | "char-outside-subset" | "length-mismatch" } {
  if (!codes || !codes.codeBytes) return { failure: "no-encoding" };
  const bytes: number[] = [];
  for (const char of text) {
    const code = codes.encode(char);
    if (code === null) return { failure: "char-outside-subset" };
    for (let shift = (codes.codeBytes - 1) * 8; shift >= 0; shift -= 8) bytes.push((code >>> shift) & 255);
  }
  if (bytes.length !== operand.encodedBytes) return { failure: "length-mismatch" };
  const token = [encodeLiteralOperand(bytes), encodeHexOperand(bytes)].find((candidate) => candidate.length === operand.tokenLength);
  return token ? { token } : { failure: "length-mismatch" };
}

/** 往两张表里写一条码↔文字（目的码解不出来就整条跳过，不写半截）。 */
function record(map: Map<number, string>, toCode: Map<string, number>, code: number, text: string | null): void {
  if (!text || Number.isNaN(code)) return;
  map.set(code, text);
  if (!toCode.has(text)) toCode.set(text, code);
}

/** 十六进制串 → UTF-16 文字（ToUnicode 的目的码可以是一串，例如连字 `fi`）。 */
function hexToUtf16(hex: string): string | null {
  if (hex.length % 4 !== 0 || !hex) return null;
  let text = "";
  for (let index = 0; index < hex.length; index += 4) {
    const unit = parseInt(hex.slice(index, index + 4), 16);
    if (Number.isNaN(unit)) return null;
    text += String.fromCharCode(unit);
  }
  return text;
}

/** 取出 CMap 里成对关键字包住的文本块。 */
function matchBlocks(text: string, begin: string, end: string): string[] {
  return [...text.matchAll(new RegExp(`${begin}([\\s\\S]*?)${end}`, "g"))].map((match) => match[1] ?? '');
}
