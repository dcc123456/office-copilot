/**
 * `.pdf` 的分析读面：把读模型摊成**给 AI 看的结构化画像**（`docs/office-parity-spec.md` 阶段 7、D6/D8）。
 *
 * 与表格侧 `sheet-io` 的 `describeSheet`/`readRange`、幻灯侧 `pptx-io` 的 `describeDeck`/`readSlideText`
 * 同一个分工（`AGENTS.md` §3.3、§2.11）：页数、每页 run 数、逐条 run 的原文与字节预算这些**事实**由引擎从现场
 * 文件算，工具包（`pdf-tools`）只管入参合不合法，代理层（`apps/desktop/src/ai/pdfAgent.ts`）只搬运——
 * 模型不许自己数页、数 run，更不许自己编 `streamRef` 与起始下标（那是引擎内部坐标，跨不出这道边）。
 *
 * 模型看到的定位坐标只有两个：**页号 1 基**与**该页内的 run 序号 0 基**。这两个数由这里的排序唯一决定
 * （读模型已按页、内容流、起始下标排好），所以 `describePdf` 与 `readPdfPageText` 报的序号必定一致，
 * 代理再按同一份模型把它们翻回 `PdfTextTarget`。
 *
 * 上送内容默认截断（§2.9）：页数与每页 run 数各有上限，单条 run 文字 ≤120 字，被截掉的部分如实置
 * `pagesTruncated` / `runsTruncated` / `textTruncated`，让模型接着加大上限往下读，而不是误以为那就是全部。
 */
import { clipText } from "@office-copilot/shared";
import { isRunEditable, runEditBlock, type PdfReadModel, type PdfRunBlock, type PdfTextRun } from "./pdf-file.js";

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
  /** 这一页的文字概要（按内容流顺序拼接，超长已截，换行用 `⏎`）。 */
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
    const joined = runs.map((run) => run.text ?? "").filter((text) => text !== "").join("⏎");
    const clipped = clipText(joined, PAGE_TEXT_LIMIT);
    pages.push({
      pageNumber: index + 1,
      runCount: runs.length,
      editableCount: runs.filter(isRunEditable).length,
      text: clipped.text,
      textTruncated: clipped.truncated,
    });
  }
  return { pageCount: model.pageCount, pages, pagesTruncated: byPage.length > pages.length };
}

/**
 * 把某一页的 run 逐条摊开，供模型对准 `runIndex` 并写出等长的新文字。
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
  return { ok: true, report: { pageNumber, runs: listed, runsTruncated: runs.length > listed.length, runCount: runs.length } };
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
