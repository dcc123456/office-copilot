/**
 * 内容流的最小解析：认"显示文字"、"进 Form XObject"与"文字定位"这三类算子，其余一概跳过。
 *
 * 为什么不用一个完整的 PDF 解释器：W1 这一档要做的是"同框同字体同长度替换"（spec D8），
 * 需要的是**文字操作数的字节位置**与**当时生效的字体资源名**。而 W2 要把逐字形的一条条 run
 * 拼成逻辑文本行（真实中文产物一个字一条 `Tj`），靠的是**最近一次显式定位算子给的原点**。
 * 两种生产者形态都实测过（2026-10-15，两份真文件）：
 * ① Word / Google Docs 一行一条 `Tj`，换行发 `Td`/`T*`/`Tm`——同一行内的 run 共享同一个原点；
 * ② WPS 一个字发一条 `… … TD`（`<1D2C>Tj 441.594 -0 TD<1DBC>Tj …`），所以每个字形原点都不同，
 *   但**同一行的 y 相同**、换栏时 y 相同而 x 跳一大截（左栏 x=90…142、右栏 x=289…341 在同一行上）。
 * 于是分行的判据是"y 在容差内相同"，分栏还得再看 x 是否断开——只按 y 分组会把两栏串成一行。
 * **这里算的是原点，不是每条 run 的左右边界**：字形推进量要字体 Widths，本包扫描器不碰对象树。
 * 算子语义取自 ISO 32000-1 §7.5.3（词法）、§9.4.2（图形状态 `q Q cm`）与 §9.4.3-9.4.4（`Tj TJ ' "`、`Tf`、`Tm Td TD T* TL`、`BT ET`），
 * 并与 pdf.js 的 `moveText`/`setMatrix` 实现交叉核对过（`Td` 走的是**行矩阵**累加，不是当前笔位；
 * `T*` 是 `0 -Tl Td`，往下走要减 leading；`"` 前两个数是词/字间距而不是 leading）。
 */
import { readHexOperand, readLiteralOperand, type StreamOperand } from "./strings.js";

/** PDF 的矩阵 `[a b c d e f]`（行向量约定，ISO 32000-1 表 80）。只在文件内部用：外部要的是算完的原点，不是矩阵。 */
type PdfMatrix = readonly [number, number, number, number, number, number];

/** 单位矩阵。 */
const IDENTITY: PdfMatrix = [1, 0, 0, 1, 0, 0];

/**
 * 两个矩阵相乘（`m` 后接 `n`，即点先过 `m` 再过 `n`）。
 *
 * @param m 左边的矩阵。
 * @param n 右边的矩阵。
 * @returns 按表 80 合成后的新矩阵。
 */
function multiplyMatrix(m: PdfMatrix, n: PdfMatrix): PdfMatrix {
  return [
    m[0] * n[0] + m[1] * n[2],
    m[0] * n[1] + m[1] * n[3],
    m[2] * n[0] + m[3] * n[2],
    m[2] * n[1] + m[3] * n[3],
    m[4] * n[0] + m[5] * n[2] + n[4],
    m[4] * n[1] + m[5] * n[3] + n[5],
  ];
}

/** 平移矩阵 `[1 0 0 1 tx ty]`。 */
const translate = (tx: number, ty: number): PdfMatrix => [1, 0, 0, 1, tx, ty];

/**
 * 从算子的数值操作数前六位取一个矩阵。
 *
 * @param numbers 这条算子攒到的数值（顺序即操作数顺序）。
 * @returns 够六个就成矩阵；不够（内容流被截断或写歪了）回 null，调用方保持原矩阵而不是拿 0 顶。
 */
function matrixOf(numbers: number[]): PdfMatrix | null {
  if (numbers.length < 6) return null;
  const [a = 0, b = 0, c = 0, d = 0, e = 0, f = 0] = numbers;
  return [a, b, c, d, e, f];
}

/** `q` 压进去的一份状态：图形状态与文字状态一起（ISO 32000-1 §8.4.1 / §9.4 说 `q` 两者都存）。 */
interface SavedState {
  ctm: PdfMatrix;
  lineMatrix: PdfMatrix;
  textMatrix: PdfMatrix;
  leading: number;
  /** `Tf` 字面写的字号（pt），还没乘文字矩阵与 CTM 的缩放。 */
  declaredSizePt: number | null;
  fontName: string | null;
}

/** 一条显示文字的算子与它的操作数。 */
export interface TextRun {
  /** 算子名。`'` 与 `"` 等价于先换行再显示，`TJ` 的操作数来自数组的某一个元素。 */
  operator: "Tj" | "TJ" | "'" | '"';
  /** 操作数在解码后内容流里的位置与字节。 */
  operand: StreamOperand;
  /** 当时生效的字体资源名（最近一次 `Tf` 给的那个 `/Name`）；没出现过 `Tf` 时为 null。 */
  fontName: string | null;
  /**
   * 生效字号（**pt，PDF 用户空间**）：`Tf` 那个数**乘过**文字矩阵与 CTM 的缩放。
   *
   * 为什么要乘：WPS 一类生产者在内容流里写 `0.05 0 0 -0.05 x y Tm` + `/FT8 439 Tf`，字面的 439 乘上
   * 0.05 才是看到的 22pt（实测逐字形推进量 22.08pt 对得上）。直接交 439 会让下一个读这个字段的人排版排错。
   * 旋转/剪切时按面积取（`sqrt(|行列式|)`）；没出现过 `Tf` 时为 null。
   */
  fontSizePt: number | null;
  /**
   * 这条 run 的字形原点（**PDF 用户空间单位，pt**，y 轴朝上、页面左下角是 (0,0)，与页面高度那种 y 朝下相反）。
   *
   * 它是"最近一次显式定位算子"（`Tm`/`Td`/`TD`/`T*`/`'`/`"`，再叠上 `cm` 的图形变换）算出来的，
   * **不是**这条 run 的左边界：Word 一行一条 `Tj` 时同一行共享同一个原点，WPS 一字一条 `TD` 时每个原点都不同，
   * 但两种形态下**同一行的 y 都一样**——分行看 y、分栏看 x 是否断开。
   * 内容流里从没出现过定位算子时按规范就是身份矩阵，于是原点是 `(0, 0)`（这是**真值**，不是没测到）。
   */
  origin: { x: number; y: number };
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

/**
 * 一个内容流扫出来的全部结果。
 *
 * 一条 run 的 `origin` 属于**它自己那一层坐标空间**（页面主内容流，或某个 Form XObject 内部）：
 * Form 的 `/BBox`+`/Matrix` 到页面坐标的映射这里不做，所以**跨层比 x/y 没有意义**，
 * 拼行只能拿同一条内容流、同一层 Form 链的 run 相比。
 */
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
 * @returns 文字操作数（含字形原点与字号）与 `Do` 调用；语法残破时按"跳过一位"继续，不抛异常（读侧宽容，写侧才严格）。
 */
export function scanContentStream(src: string): ContentScan {
  const runs: TextRun[] = [];
  const xobjects: XObjectCall[] = [];
  let index = 0;
  let fontName: string | null = null;
  /** `Tf` 字面写的字号（pt）：交出去前还要乘文字矩阵与 CTM 的缩放（见 {@link fontSizeInUse}）。 */
  let declaredSizePt: number | null = null;
  let pendingName: string | null = null;
  let pendingString: StreamOperand | null = null;
  let pendingArray: StreamOperand[] | null = null;
  /** 本条算子累积到的数值操作数（`Tm` 要六个、`Td` 要两个、`Tf` 要字号）。 */
  let pendingNumbers: number[] = [];
  /** 图形状态：CTM（`cm` 复合）、`q`/`Q` 的状态栈，以及文字状态三项（行矩阵、文字矩阵、leading）。 */
  let ctm: PdfMatrix = IDENTITY;
  const stateStack: SavedState[] = [];
  let lineMatrix: PdfMatrix = IDENTITY;
  let textMatrix: PdfMatrix = IDENTITY;
  let leading = 0;

  /** 沿行矩阵往下走一个 leading（`T*`／`'`／`"` 共用那一步）。 */
  const nextLine = (): void => {
    // 规范把 `T*` 定义为 `0 -Tl Td`：leading 存的是正数，往下走就是加一个负的 y。
    lineMatrix = multiplyMatrix(translate(0, -leading), lineMatrix);
    textMatrix = lineMatrix;
  };

  /** 沿行矩阵平移（`Td`／`TD` 那一步：规范是在**行矩阵**上累加，不是当前笔位）。 */
  const moveLine = (tx: number, ty: number): void => {
    lineMatrix = multiplyMatrix(translate(tx, ty), lineMatrix);
    textMatrix = lineMatrix;
  };

  /** 当前"文字矩阵接 CTM"的合成矩阵：原点与缩放都从它取，避免两处各算一遍（§3.2）。 */
  const deviceMatrix = (): PdfMatrix => multiplyMatrix(textMatrix, ctm);

  /** 把当前文字矩阵叠上 CTM，取字形原点（矩阵作用在 (0,0) 上就是那两个平移量）。 */
  const originOf = (): { x: number; y: number } => {
    const device = deviceMatrix();
    return { x: device[4], y: device[5] };
  };

  /**
   * 生效字号：`Tf` 的字面数乘上合成矩阵的缩放。
   *
   * @returns 没出现过 `Tf` 时 null；否则 `字号 × sqrt(|行列式|)`（等比缩放就是那个缩放系数，旋转/剪切按面积）。
   */
  const fontSizeInUse = (): number | null => {
    if (declaredSizePt === null) return null;
    const device = deviceMatrix();
    return declaredSizePt * Math.sqrt(Math.abs(device[0] * device[3] - device[1] * device[2]));
  };

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
      const name = readName(src, index);
      pendingName = name.name;
      index = name.end;
      continue;
    }
    if (isNumberStart(char)) {
      const number = readNumber(src, index);
      if (Number.isFinite(number.value)) pendingNumbers.push(number.value);
      index = number.end;
      continue;
    }
    const word = readKeyword(src, index);
    index = word.end;
    switch (word.value) {
      case "Tf": {
        if (pendingName) fontName = pendingName;
        // 字号是 `/F1 12 Tf` 里那个数；没带数（写歪的 `Tf /F1 Tf`）就保持上一个，而不是清空已知的值。
        const size = pendingNumbers[pendingNumbers.length - 1];
        if (size !== undefined) declaredSizePt = size;
        break;
      }
      case "Do":
        if (pendingName) xobjects.push({ name: pendingName, at: word.end, fontName });
        break;
      case "q":
        stateStack.push({ ctm, lineMatrix, textMatrix, leading, declaredSizePt, fontName });
        break;
      case "Q": {
        const saved = stateStack.pop();
        if (saved) {
          ctm = saved.ctm;
          lineMatrix = saved.lineMatrix;
          textMatrix = saved.textMatrix;
          leading = saved.leading;
          declaredSizePt = saved.declaredSizePt;
          fontName = saved.fontName;
        }
        break;
      }
      case "cm": {
        const pushed = matrixOf(pendingNumbers);
        if (pushed) ctm = multiplyMatrix(pushed, ctm);
        break;
      }
      case "BT":
        // 规范说 BT 把 Tm 与 Tlm 一起置为身份矩阵（leading 与字体状态留着不动）。
        lineMatrix = IDENTITY;
        textMatrix = IDENTITY;
        break;
      case "Tm": {
        // `Tm` 是**直接指定**（不是平移），所以行矩阵与文字矩阵一起换成它。
        const given = matrixOf(pendingNumbers);
        if (given) {
          lineMatrix = given;
          textMatrix = given;
        }
        break;
      }
      case "Td": {
        const [tx, ty] = pendingNumbers;
        if (tx !== undefined && ty !== undefined) moveLine(tx, ty);
        break;
      }
      case "TD": {
        // `TD tx ty`：规范说先把 leading 设成 `-ty`，再执行 `tx ty Td`。
        const [tx, ty] = pendingNumbers;
        if (tx !== undefined && ty !== undefined) {
          leading = -ty;
          moveLine(tx, ty);
        }
        break;
      }
      case "TL": {
        const [lineLeading] = pendingNumbers;
        if (lineLeading !== undefined) leading = lineLeading;
        break;
      }
      case "T*":
        nextLine();
        break;
      case '"':
        // `aw ac (string) "`：规范说那两个数是**词间距与字间距**（不是 leading），等价于先设这两项再执行 `'`。
        // 本包不算字形推进量，所以这两个数没有可作用的地方，剩下的一步就是换行加显示。
        nextLine();
        if (pendingString) runs.push({ operator: '"', operand: pendingString, fontName, fontSizePt: fontSizeInUse(), origin: originOf() });
        break;
      case "'":
        nextLine();
        if (pendingString) runs.push({ operator: "'", operand: pendingString, fontName, fontSizePt: fontSizeInUse(), origin: originOf() });
        break;
      case "Tj":
        if (pendingString) runs.push({ operator: "Tj", operand: pendingString, fontName, fontSizePt: fontSizeInUse(), origin: originOf() });
        break;
      case "TJ":
        // 数组里那些数值是 kerning（负值右移、正值左移），本包不算宽度，所以整条数组共享同一个原点。
        for (const operand of pendingArray ?? []) {
          runs.push({ operator: "TJ", operand, fontName, fontSizePt: fontSizeInUse(), origin: originOf() });
        }
        break;
      default:
        break;
    }
    pendingName = null;
    pendingString = null;
    pendingArray = null;
    pendingNumbers = [];
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

/**
 * 读一个数字 token。
 *
 * @param src 内容流文本。
 * @param start 该数字起始下标。
 * @returns 解析出的数值（读不成数字回 `NaN`，调用方据此放弃这次定位而不是拿 0 顶）与结束下标。
 */
function readNumber(src: string, start: number): { value: number; end: number } {
  let index = start;
  if (at(src, index) === "+" || at(src, index) === "-") index += 1;
  while (index < src.length && /[0-9.]/.test(at(src, index))) index += 1;
  const end = index === start ? start + 1 : index;
  return { value: Number.parseFloat(src.slice(start, end)), end };
}

/** 读一个算子/关键字 token（连续的非常规字符）。 */
function readKeyword(src: string, start: number): { value: string; end: number } {
  let index = start;
  while (index < src.length && !WHITESPACE.has(at(src, index)) && !DELIMITERS.has(at(src, index))) index += 1;
  return { value: src.slice(start, index), end: index };
}
