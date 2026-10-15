/**
 * `.pdf` 的分析读面：把读模型摊成**给 AI 看的结构化画像**（`docs/office-parity-spec.md` 阶段 7、D6/D8）。
 *
 * 与表格侧 `sheet-io` 的 `describeSheet`/`readRange`、幻灯侧 `pptx-io` 的 `describeDeck`/`readSlideText`
 * 同一个分工（`AGENTS.md` §3.3、§2.11）：页数、每页 run 数、逐条 run 的原文与字节预算这些**事实**由引擎从现场
 * 文件算，工具包（`pdf-tools`）只管入参合不合法，代理层（`apps/desktop/src/ai/pdfAgent.ts`）只搬运——
 * 模型不许自己数页、数 run，更不许自己编 `streamRef` 与起始下标（那是引擎内部坐标，跨不出这道边）。
 *
 * 模型看到的定位坐标只有三个：**页号 1 基**、**该页内的 run 序号 0 基**、**该页内的行号 0 基**。这些数由这里的排序唯一决定
 * （读模型已按页、内容流、起始下标排好，行由 `pdfTextLinesOf` 按同一顺序拼），所以 `describePdf`、`readPdfPageText`
 * 与 `pdfLineEditsOf` 报的序号必定一致，代理再按同一份模型把它们翻回 `PdfTextTarget`。
 *
 * 上送内容默认截断（§2.9）：页数与每页 run 数各有上限，单条 run 文字 ≤120 字，被截掉的部分如实置
 * `pagesTruncated` / `runsTruncated` / `textTruncated`，让模型接着加大上限往下读，而不是误以为那就是全部。
 */
import { clipText } from "@office-copilot/shared";
import { isRunEditable, pdfTextLinesOf, runEditBlock, targetOfRun, type PdfReadModel, type PdfRunBlock, type PdfTextLine, type PdfTextEdit, type PdfTextRun } from "./pdf-model.js";

/** 单条 run 文字上送的上限（字数；与幻灯侧 `RUN_TEXT_LIMIT`、表格侧单格 120 字同一口径，§2.9）。 */
const RUN_TEXT_LIMIT = 120;

/** 一页的文字概要上送的上限（概要只用来让模型认出是哪一页，全文要靠 `readPdfPageText` 逐条读）。 */
const PAGE_TEXT_LIMIT = 400;

/** `describePdf` 的入参上限。 */
export interface PdfDescribeOptions {
  /** 最多描述多少页（1~200）。 */
  maxPages: number;
}

/** 一页给模型看的概要。 */
export interface PdfPageOutline {
  /** 页号（1 基，与 `pdf_read_page_text` 的入参同一个口径）。 */
  pageNumber: number;
  /** 该页显示文字条数（`maxRunsPerPage` 截断时用它说明只列了前面几条）。 */
  runCount: number;
  /** 其中能改的条数（内联流与认不出原文的那两种在这里已经排除）。 */
  editableCount: number;
  /** 该页的逻辑行数（逐字形的中文产物里这个数远小于 run 数；`set_line_text` 的行号就是按它编号的）。 */
  lineCount: number;
  /** 其中能改的行数（行内任何一条 run 改不了，整行就算改不了）。 */
  editableLineCount: number;
  /** 这一页的文字概要（按**逻辑文本行**拼，逐字形的产物已拼回整行；超长已截，行之间用 `⏎`）。 */
  text: string;
  /** 概要是否被 {@link PAGE_TEXT_LIMIT} 截断。 */
  textTruncated: boolean;
}

/** `describePdf` 的结果。 */
export interface PdfOutline {
  /** 总页数。 */
  pageCount: number;
  /** 列出来的页（`maxPages` 截断时后面的页不在这里，但 `pageCount` 仍是全量）。 */
  pages: PdfPageOutline[];
  /** 页清单是否被截断。 */
  pagesTruncated: boolean;
}

/** 一条 run 摊开给模型看的样子：定位序号 + 原文 + 尺寸 + 改不了的话卡在哪一类。 */
export interface PdfRunSnapshot {
  /** run 序号（0 基，该页内连续编号；模型提交改动时用它与页号一起定位）。 */
  runIndex: number;
  /** 这条 run 的现文字；认不出原文时是空串。 */
  text: string;
  /** 文字是否被截断（§2.9）。 */
  textTruncated: boolean;
  /** 原文的字数（按码点计）；改完的文字必须与它**完全相同**，否则装不回原字节数（spec D8 的 W1 档）。 */
  charCount: number;
  /** 原串在内容流里的字节数（只作说明用；等长判定由引擎编码时做，这里不参与决策，§2.1）。 */
  encodedBytes: number;
  /** 生效字体的 `/BaseFont`；资源链里查不到这个字体时 null。 */
  baseFont: string | null;
  /** 是不是 Type0/CID 复合字体（中文产物基本都是）。 */
  composite: boolean;
  /** Form XObject 资源名链（如 `Fm0`）；主内容流为 null。 */
  withinForm: string | null;
  /** 这条能不能改（true 等价于 `block === null`）。 */
  editable: boolean;
  /** 改不了的那一类码；能改时为 null。中文理由由调用方组装（§5.7）。 */
  block: PdfRunBlock | null;
}

/** 一条逻辑行摊给模型看的样子：行号（提交改动用的把手）+ 整行文字 + 它由哪几条 run 拼成。 */
export interface PdfLineSnapshot {
  /** 行号（0 基，该页内连续编号；与 {@link pdfLineEditsOf} 的 `lineIndex` 同一个口径）。 */
  lineIndex: number;
  /** 这一行的现文字（认不出原文的那些位置是缺字的，见 {@link hasUndecoded}）。 */
  text: string;
  /** 文字是否被截断（§2.9）。 */
  textTruncated: boolean;
  /** 整行的字数（按码点计）；改成新文字时**总字数必须与它相同**（spec D8 的 W1 档：等长是逐 run 判的）。 */
  charCount: number;
  /** 这一行覆盖的 run 序号区间：含 `runFrom`、不含 `runTo`（与 {@link PdfRunSnapshot.runIndex} 同一个排序）。 */
  runFrom: number;
  /** 区间的右端（不含）；`runTo - runFrom` 就是这一行由几条 run 拼成（逐字形的中文产物通常是几十个）。 */
  runTo: number;
  /** 这一行能不能改（行内任何一条 run 改不了，整行就改不了——等长替换要逐 run 落，缺一条就拼不回这句话）。 */
  editable: boolean;
  /** 改不了时**第一条**卡住的 run 属于哪一类；能改时为 null。中文理由由调用方组装（§5.7）。 */
  block: PdfRunBlock | null;
  /** 行里有没有认不出原文的 run（true 时 `text` 是缺字的，那句不能当完整话读，更不许改）。 */
  hasUndecoded: boolean;
}

/** `readPdfPageText` 的结果。 */
export interface PdfPageTextReport {
  /** 页号（1 基）。 */
  pageNumber: number;
  /** 该页全部 run 的快照，按 `runIndex` 升序。 */
  runs: PdfRunSnapshot[];
  /** run 清单是否被 `maxRunsPerPage` 截断（截断时后面的 run 没上送，但 `runs.length` 不是全量）。 */
  runsTruncated: boolean;
  /** 该页 run 总数（全量，与 `runs.length` 对比就知道截没截）。 */
  runCount: number;
  /** 该页逻辑行的快照，按 `lineIndex` 升序（同一个 `maxRunsPerPage` 上限）。 */
  lines: PdfLineSnapshot[];
  /** 行清单是否被截断。 */
  linesTruncated: boolean;
  /** 该页逻辑行总数（全量）。 */
  lineCount: number;
}

/**
 * 描述整份 PDF：页数与每页的文字概要。
 *
 * @param model 主进程读出来的模型。
 * @param options 页数上限（工具入参已校验过范围，这里只照做）。
 * @returns 结构化画像；被截断的页置了 `pagesTruncated`，不静默丢。
 */
export function describePdf(model: PdfReadModel, options: PdfDescribeOptions): PdfOutline {
  const byPage = groupRunsByPage(model.runs, model.pageCount);
  const pages: PdfPageOutline[] = [];
  for (let index = 0; index < Math.min(byPage.length, options.maxPages); index += 1) {
    const runs = byPage[index] ?? [];
    // 按**逻辑行**拼，不是按 run 拼：真实中文产物一个字一条 run，按 run 拼出来的概要是"个⏎人⏎简⏎历"，
    // 模型读不出这是一句话，也就提不出正确的改动。
    const lines = pdfTextLinesOf(runs);
    const joined = lines
      .map((line) => line.text)
      .filter((text) => text !== "")
      .join("⏎");
    const clipped = clipText(joined, PAGE_TEXT_LIMIT);
    pages.push({
      pageNumber: index + 1,
      runCount: runs.length,
      editableCount: runs.filter(isRunEditable).length,
      lineCount: lines.length,
      editableLineCount: lines.filter((line) => runs.slice(line.fromRun, line.toRun).every(isRunEditable)).length,
      text: clipped.text,
      textTruncated: clipped.truncated,
    });
  }
  return { pageCount: model.pageCount, pages, pagesTruncated: byPage.length > pages.length };
}

/**
 * 把某一页摊开：逐条 run 与逐条逻辑行（同一份排序的两个粒度），供模型对准 `runIndex`/`lineIndex` 并写出等长的新文字。
 *
 * @param model 主进程读出来的模型。
 * @param pageNumber 页号（1 基；越界时返回 `ok:false` 与给模型看的中文理由，不抛异常）。
 * @param maxRunsPerPage 每页最多列多少条 run（超出置 `runsTruncated`）。
 * @returns 该页的文字报告，或一句中文失败理由（页号越界）。
 */
export function readPdfPageText(model: PdfReadModel, pageNumber: number, maxRunsPerPage: number): { ok: true; report: PdfPageTextReport } | { ok: false; reason: string } {
  const groups = groupRunsByPage(model.runs, model.pageCount);
  const runs = groups[pageNumber - 1];
  if (!runs) return { ok: false, reason: `没有第 ${pageNumber} 页：这份 PDF 一共 ${model.pageCount} 页，先用 pdf_describe 看页号。` };
  const listed = runs.slice(0, maxRunsPerPage).map((run, runIndex) => snapshotOfRun(run, runIndex));
  const lines = pdfTextLinesOf(runs);
  const listedLines = lines.slice(0, maxRunsPerPage).map((line, lineIndex) => snapshotOfLine(line, runs, lineIndex));
  return {
    ok: true,
    report: {
      pageNumber,
      runs: listed,
      runsTruncated: runs.length > listed.length,
      runCount: runs.length,
      lines: listedLines,
      linesTruncated: lines.length > listedLines.length,
      lineCount: lines.length,
    },
  };
}

/**
 * 把"改这一行"展开成逐 run 的等长替换（spec D8 的 W1.5：整行那一档，等长语义与 W1 完全相同）。
 *
 * 展开规则只有一条：**每条 run 分到它原来那几个字**。新文字按码点排好后依次切给行内各条 run，
 * 于是逐字形的中文产物里"姓名：丁长长"这一行会展开成 6 条各一个字的替换，而 Word 那种一行一条 run
 * 的产物展开成一条整行替换——两种产物都不许改字数，等长这条由 {@link PdfRunSnapshot.charCount} 事先判掉，
 * 最终字节数仍由引擎编码时判（§2.1）。
 *
 * @param model 主进程读出来的模型。
 * @param pageNumber 页号（1 基）。
 * @param lineIndex 该页内的行号（0 基，与 `pdf_read_page_text` 报的 `lineIndex` 同一个排序）。
 * @param newText 这一行要换成的新文字（总字数必须与该行相同）。
 * @returns `edits` 是按行内顺序排好的逐条替换（定位键由引擎填，模型看不到也不许编）；翻不动时给一句中文理由。
 */
export function pdfLineEditsOf(model: PdfReadModel, pageNumber: number, lineIndex: number, newText: string): { ok: true; edits: PdfTextEdit[] } | { ok: false; reason: string } {
  const groups = groupRunsByPage(model.runs, model.pageCount);
  const runs = groups[pageNumber - 1];
  if (!runs) return { ok: false, reason: `没有第 ${pageNumber} 页：这份 PDF 一共 ${model.pageCount} 页，先用 pdf_describe 看页号。` };
  const lines = pdfTextLinesOf(runs);
  const line = lines[lineIndex];
  if (!line) return { ok: false, reason: `第 ${pageNumber} 页只有 ${lines.length} 行，没有第 ${lineIndex} 行：用 pdf_read_page_text 核对行号，别自己数。` };
  const members = runs.slice(line.fromRun, line.toRun);
  const blocked = members.find((run) => !isRunEditable(run));
  if (blocked) return { ok: false, reason: `第 ${pageNumber} 页第 ${lineIndex} 行改不了：${blockReason(blocked)}（这一行由 ${members.length} 条 run 拼成，缺一条就拼不回这句话，整行不许动）。` };
  const characters = [...newText];
  const total = members.reduce((sum, run) => sum + [...(run.text ?? "")].length, 0);
  if (characters.length !== total) {
    return { ok: false, reason: `第 ${pageNumber} 页第 ${lineIndex} 行原文是 ${total} 个字，新文字是 ${characters.length} 个字：PDF 不回流，字数不等装不回原位置。` };
  }
  if (newText === line.text) return { ok: false, reason: `第 ${pageNumber} 页第 ${lineIndex} 行的新文字与原文一样，不用改。` };
  const edits: PdfTextEdit[] = [];
  let offset = 0;
  for (const run of members) {
    const count = [...(run.text ?? "")].length;
    const target = targetOfRun(run);
    if (!target) return { ok: false, reason: `第 ${pageNumber} 页第 ${lineIndex} 行里有一条文字在内联内容流里，没有可标记的对象，整行不动。` };
    edits.push({ target, text: characters.slice(offset, offset + count).join("") });
    offset += count;
  }
  return { ok: true, edits };
}

/**
 * 一条 run 改不了的中文理由（这一层的理由直接回给模型看，与 `readPdfPageText` 的越界理由同一处组装）。
 *
 * @param run 改不了的那条显示文字。
 * @returns 一句中文：内联内容流 / 认不出原文。
 */
function blockReason(run: PdfTextRun): string {
  switch (runEditBlock(run)) {
    case "inline-stream":
      return "那条文字内联在页面内容流里，没有可单独标记的对象";
    case "undecoded":
      return "认不出那条文字原文（字体缺 `/ToUnicode` 或子集码表覆盖不到），旧文字都不知道就没法安全撤销";
    default:
      return "那条文字改不了";
  }
}

/**
 * 一条逻辑行 → 给模型看的快照。
 *
 * @param line 引擎拼出来的一条逻辑行。
 * @param runs 该页的全部 run（`line.fromRun`/`line.toRun` 是相对它的下标，用来查行内第一条改不了的 run）。
 * @param lineIndex 该页内的行号（0 基）。
 * @returns 快照；行内任何一条 run 改不了时整行标成不可改，并带上第一条卡住的那一类的码。
 */
function snapshotOfLine(line: PdfTextLine, runs: readonly PdfTextRun[], lineIndex: number): PdfLineSnapshot {
  const members = runs.slice(line.fromRun, line.toRun);
  const blocked = members.find((run) => !isRunEditable(run));
  const clipped = clipText(line.text, RUN_TEXT_LIMIT);
  return {
    lineIndex,
    text: clipped.text,
    textTruncated: clipped.truncated,
    charCount: [...line.text].length,
    runFrom: line.fromRun,
    runTo: line.toRun,
    editable: blocked === undefined,
    block: blocked ? runEditBlock(blocked) : null,
    hasUndecoded: line.hasUndecoded,
  };
}

/**
 * 一条 run → 给模型看的快照。
 *
 * @param run 读模型里的一条显示文字。
 * @param runIndex 该页内的序号（0 基）。
 * @returns 快照；认不出原文时 `text` 是空串、`charCount` 是 0，并带上 `block` 码说明为什么改不了。
 */
function snapshotOfRun(run: PdfTextRun, runIndex: number): PdfRunSnapshot {
  const text = run.text ?? "";
  const clipped = clipText(text, RUN_TEXT_LIMIT);
  return {
    runIndex,
    text: clipped.text,
    textTruncated: clipped.truncated,
    charCount: [...text].length,
    encodedBytes: run.encodedBytes,
    baseFont: run.baseFont,
    composite: run.composite,
    withinForm: run.withinForm,
    editable: isRunEditable(run),
    block: runEditBlock(run),
  };
}

/**
 * 按页分组。
 *
 * @param runs 全部 run（读模型已按页、内容流、起始下标排好序）。
 * @param pageCount 页数（来自读模型，不是 run 里出现过的最大页号）。
 * @returns 长度等于页数的数组；没有文字的页是空数组，这样页号（1 基 = 下标 + 1）与页始终对得上。
 */
function groupRunsByPage(runs: PdfTextRun[], pageCount: number): PdfTextRun[][] {
  const groups: PdfTextRun[][] = Array.from({ length: pageCount }, () => []);
  for (const run of runs) if (run.pageIndex >= 0 && run.pageIndex < pageCount) groups[run.pageIndex]?.push(run);
  return groups;
}
