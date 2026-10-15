/**
 * `.pdf` 读模型里**不含 Node 能力**的那一半：数据类型与"这条 run 能不能改"的判定。
 *
 * 拆出这个文件是因为 `AGENTS.md` §5.2：渲染层不许碰 Node。`pdf-file.ts` 要 `node:zlib` 解内容流、
 * 要 `@cantoo/pdf-lib` 做增量更新，那些能力只属于主进程；而判定与定位键是渲染层（PDF 视图、AI 面板）
 * 也要读的同一份实现，所以放在这里，由 `pdf-file.ts` 反向引用——判定仍然只有一份（§3.2）。
 *
 * 渲染层引用这一层的唯一入口是子路径 `@office-copilot/pdf-io/pure`（见包内 `exports`）；
 * 主进程继续用 `@office-copilot/pdf-io` 那个总入口。
 */
import type { TextRun } from "./content-stream.js";

/** 一条显示文字的算子在模型里的样子。 */
export interface PdfTextRun {
  /** 页号（0 基）。 */
  pageIndex: number;
  /** 所在内容流的对象引用（如 `6 0 R`）；内联（非间接对象）内容流为 null，那种改不了。 */
  streamRef: string | null;
  /** 从页面内容流进第几层 Form XObject 的资源名链（如 `Fm0/Fm1`）；主内容流为 null。 */
  withinForm: string | null;
  /** 算子名。 */
  operator: TextRun["operator"];
  /** 串的书写形态。 */
  form: "literal" | "hex";
  /** 操作数在**解压后**内容流里的起始下标（字节）。与 `streamRef`、`withinForm` 合起来才是一条 run 的身份。 */
  start: number;
  /** 原串的字节数。 */
  encodedBytes: number;
  /** 原 token 的字符数（含定界符），改回去必须等长。 */
  tokenLength: number;
  /** 生效字体的 `/BaseFont`；资源链里查不到这个字体时 null。 */
  baseFont: string | null;
  /** 是不是 Type0/CID 复合字体（中文产物基本都是）。 */
  composite: boolean;
  /**
   * 生效字号（**pt，PDF 用户空间**，不是 CSS px）：`Tf` 字面那个数**已经乘过**文字矩阵与 CTM 的缩放。
   *
   * 乘缩放是因为 WPS 一类产物写 `0.05 0 0 -0.05 x y Tm` + `/FT8 439 Tf`，看到的其实是 22pt；
   * 直接把 439 交出去会让下一个读这个字段的人（视图、AI 概要）按 439pt 排版。没出现过 `Tf` 时为 null。
   */
  fontSizePt: number | null;
  /**
   * 字形原点（**PDF 用户空间单位，pt**，y 轴朝上、页面左下角是 (0,0)）。
   *
   * 取自最近一次显式定位算子（`Tm`/`Td`/`TD`/`T*`/`'`/`"`）再叠上 `cm`，**不是这条 run 的左边界**：
   * Word 一行一条 `Tj` 时同一行共享原点，WPS 一字一条 `TD` 时每个字形原点都不同，但两种形态下**同一行的 y 相同**，
   * 所以拼行按 y（带容差）、分栏再看 x 是否断开。跨内容流/跨 Form 层相比没有意义（Form 的 `/Matrix` 映射这里不做）。
   * 从没出现过定位算子时按规范就是身份矩阵，于是 `(0, 0)` 是**真值**。
   */
  origin: { x: number; y: number };
  /** 解出来的文字；字体认不出或子集缺码时为 null。 */
  text: string | null;
}

/** `.pdf` 的读模型。 */
export interface PdfReadModel {
  /** 页数。 */
  pageCount: number;
  /** 全部文字 run，按页、按内容流、按下标排序。 */
  runs: PdfTextRun[];
}

/** 一条改写的定位键（取自 {@link PdfTextRun} 的同名字段）。 */
export interface PdfTextTarget {
  /** 页号（0 基）。 */
  pageIndex: number;
  /** 内容流对象引用。 */
  streamRef: string;
  /** Form XObject 资源名链；主内容流给 null。 */
  withinForm: string | null;
  /** 操作数起始下标。 */
  start: number;
}

/** 一处待落地的文字改写。 */
export interface PdfTextEdit {
  /** 改哪一条 run。 */
  target: PdfTextTarget;
  /** 新文字（Unicode）。 */
  text: string;
}

/** 一处已落地的改动：撤销要靠它把旧文字换回去，所以旧文字必须随结果一起交出去（§2.5）。 */
export interface PdfAppliedEdit {
  /** 改的那条 run 的定位键。 */
  target: PdfTextTarget;
  /** 改之前的文字。 */
  fromText: string;
  /** 改之后的文字。 */
  toText: string;
}

/** 一处改不了的原因（结构化，文案由调用方组装，`AGENTS.md` §5.7）。 */
export interface PdfEditFailure {
  /** 失败的那一条定位键。 */
  target: PdfTextTarget;
  /**
   * 原因码：找不到这条 run / 看不出这段文字念作什么（字体认不出、缺 `/ToUnicode` 或码表覆盖不到，
   * 不知道原文就不许改） / 子集里没有这个字 / 新串装不回原字节数 /
   * 撤销时那条已经不是我写出去的样子（内容寻址，不牵连后来的编辑，§2.5）。
   */
  reason: "run-not-found" | "no-encoding" | "char-outside-subset" | "length-mismatch" | "text-drifted";
}

/** 一条 run 改不了的类别（结构化码；回给人听的中文由调用方组装，`AGENTS.md` §5.7）。 */
export type PdfRunBlock = "inline-stream" | "undecoded";

/**
 * 这条 run 改不了的话，卡在哪一类（判定只在这一处：手工视图与 AI 代理都读它）。
 *
 * @param run 读模型里的一条显示文字。
 * @returns `inline-stream`＝文字内联在页面里、没有可标记的间接对象；`undecoded`＝认不出这段念作什么
 *   （连旧文字都不知道，§2.5 的撤销无从比对）；null＝可以改。
 */
export function runEditBlock(run: PdfTextRun): PdfRunBlock | null {
  if (run.streamRef === null) return "inline-stream";
  if (run.text === null) return "undecoded";
  return null;
}

/**
 * 这条 run 能不能改。
 *
 * @param run 读模型里的一条显示文字。
 * @returns {@link runEditBlock} 为 null 时 true。
 */
export function isRunEditable(run: PdfTextRun): boolean {
  return runEditBlock(run) === null;
}

/**
 * 从一条 run 抽出改写定位键。
 *
 * @param run 读模型里的一条显示文字。
 * @returns 四个字段组成的定位键；内联流没有可标记的对象，回 null（这条也就改不了）。
 */
export function targetOfRun(run: PdfTextRun): PdfTextTarget | null {
  if (run.streamRef === null) return null;
  return { pageIndex: run.pageIndex, streamRef: run.streamRef, withinForm: run.withinForm, start: run.start };
}

/**
 * 定位键的稳定写法：四个字段拼成一句，供"同一条 run"的比对与草稿索引。
 *
 * @param target 定位键。
 * @returns `页号|内容流引用|Form 链|起始下标`；写入侧的撤销核对与渲染层的草稿表共用这一份实现。
 */
export function pdfTargetKey(target: PdfTextTarget): string {
  return `${target.pageIndex}|${target.streamRef}|${target.withinForm ?? ""}|${target.start}`;
}

/**
 * 一条逻辑文本行：把逐字形的 run 拼回"人能读的一句"。
 *
 * 存在的原因（2026-10-15 两份真文件实测）：真实中文产物**一个字发一条显示算子**
 * （WPS 简历 6 页 5196 条 run，每字形一次 `Tj` + 一次 `TD`），所以 run 不是行、也不是词。
 * 拼行只靠几何：同一层里 **y 在容差内相同**就是同一行；**x 间隔超过字号 4 倍**判为新行（那是换栏或标签位）。
 * 行是**连续下标段**（按传入顺序，即内容流顺序），所以 `fromRun`/`toRun` 就足够换回定位键。
 */
export interface PdfTextLine {
  /** 页号（0 基）。 */
  pageIndex: number;
  /** 所在内容流的对象引用；与 run 一样，内联流为 null。 */
  streamRef: string | null;
  /** Form XObject 资源名链；主内容流为 null。 */
  withinForm: string | null;
  /** 这条行覆盖的 run 下标区间：含 `fromRun`、不含 `toRun`（指向 {@link pdfTextLinesOf} 的入参数组）。 */
  fromRun: number;
  /** 区间的右端（不含）：`toRun - fromRun` 就是这条行由几条 run 拼成。 */
  toRun: number;
  /**
   * 拼出来的文字（Unicode）。
   *
   * 段与段之间**不补空格**：中文逐字形的相邻 run 补空格会凭空多出一个字符。
   * 含解不出原文的 run 时那些位置直接缺字，靠 {@link hasUndecoded} 如实标出来。
   */
  text: string;
  /** 行首字形原点（PDF 用户空间 pt，y 朝上）。 */
  x: number;
  /** 这条行的 y（取行首那条 run 的原点）。 */
  y: number;
  /** 行内最后一条 run 的原点 x（**不是右边界**：本包不读字形宽度）。 */
  lastOriginX: number;
  /** 生效字号（pt）：取行内第一条有字号的 run；整行都没有时 null。 */
  fontSizePt: number | null;
  /** 这条行里有没有解不出原文的 run（true 时 `text` 是缺字的，不许当成完整句子送给模型或用户）。 */
  hasUndecoded: boolean;
}

/** y 容差（pt）：WPS 那份简历里同一条视觉行内相邻字形的 y 实测最多差 0.119pt（`TD` 的累加抖动），而真是两条行时最近的两行差 15.12pt，0.6pt 落在两者之间。 */
const LINE_Y_TOLERANCE_PT = 0.6;
/**
 * 断行的 x 间隔上限，按**字号归一**（em）。
 *
 * 为什么不写死 pt：实测同一行内相邻字形的推进量 p50=1.00em、p90=1.01em（中日韩字形都是一个 em 宽），
 * 而换栏/标签位的间隔是 9em 起（94pt/157pt 那些），所以 4em 落在两类之间最空的带子里。
 * 直接写 40pt 会在大字号标题上把一行切开、在小字号正文里把跨栏的并成一行。
 */
const LINE_BREAK_GAP_EM = 4;

/**
 * 把一条页内、同一坐标层的 run 拼成逻辑文本行（spec 阶段 7 的 W2 前置）。
 *
 * @param runs 读模型里的全部 run，**必须按内容流顺序**（{@link PdfReadModel} 的 `runs` 就是：页 → 流 → 起始下标）。
 * @returns 逻辑行，按传入顺序；跨页、跨内容流、跨 Form 链的一律不合并（三者的原点不在同一个坐标空间里）。
 *
 * 已知不做：① 旋转排版的文字（y 会一路漂，会被切成一条一个字一行）；② 字形宽度与真正的右边界——
 * 只比原点，所以 `lastOriginX` 是"最后一个字形的起点"。
 */
export function pdfTextLinesOf(runs: readonly PdfTextRun[]): PdfTextLine[] {
  const lines: PdfTextLine[] = [];

  /** 从一条 run 起一条新行（下标 `index` 是它在入参数组里的位置）。 */
  const startLine = (index: number, run: PdfTextRun): PdfTextLine => ({
    pageIndex: run.pageIndex,
    streamRef: run.streamRef,
    withinForm: run.withinForm,
    fromRun: index,
    toRun: index + 1,
    text: run.text ?? "",
    x: run.origin.x,
    y: run.origin.y,
    lastOriginX: run.origin.x,
    fontSizePt: run.fontSizePt,
    hasUndecoded: run.text === null,
  });

  /**
   * 这条 run 是否该另起一行（只在同一坐标层内判）。
   *
   * @param line 当前攒着的行。
   * @param run 下一条 run。
   * @returns y 出了容差、x 倒退、或空出超过 4em 时 true。
   */
  const isBreak = (line: PdfTextLine, run: PdfTextRun): boolean => {
    const em = line.fontSizePt ?? run.fontSizePt ?? 0;
    const gap = run.origin.x - line.lastOriginX;
    return Math.abs(run.origin.y - line.y) > LINE_Y_TOLERANCE_PT || gap < -LINE_Y_TOLERANCE_PT || (em > 0 && gap > LINE_BREAK_GAP_EM * em);
  };

  /** 当前正在攒的那条行。 */
  let open: PdfTextLine | null = null;
  for (let index = 0; index < runs.length; index += 1) {
    const run = runs[index];
    if (!run) continue;
    const current = open;
    const sameLayer = current !== null && current.pageIndex === run.pageIndex && current.streamRef === run.streamRef && current.withinForm === run.withinForm;
    if (current === null || !sameLayer || isBreak(current, run)) {
      open = startLine(index, run);
      lines.push(open);
      continue;
    }
    current.toRun = index + 1;
    current.text += run.text ?? "";
    current.lastOriginX = run.origin.x;
    if (current.fontSizePt === null) current.fontSizePt = run.fontSizePt;
    if (run.text === null) current.hasUndecoded = true;
  }
  return lines;
}
