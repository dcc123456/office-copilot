/**
 * PDF 视图要用的那几个纯函数（组件只做装配，判定放这里，`AGENTS.md` §3.3）。
 *
 * 两条来自引擎事实的守卫都在这里收成一句中文：① 内联内容流（`streamRef` 为 null）改不了；
 * ② 解不出原文的 run（`text` 为 null：字体认不出或缺 `/ToUnicode`）连"从什么改成什么"都不知道，
 * 所以只读（spec D8 的 W1 档）。**字节预算不在这里判**：编码规则（WinAnsi / CID 每字 2 字节）属于
 * `pdf-io`，视图再算一遍就是第二套换算（§2.1），所以草稿交给引擎、由它回 `length-mismatch`。
 */
import type { PdfEditFailure, PdfReadModel, PdfTextEdit, PdfTextRun, PdfTextTarget } from "@office-copilot/pdf-io";

/** 一条草稿在渲染层里的样子：`key` 用 {@link pdfTargetKey}，`text` 是用户敲的新文字。 */
export type PdfDrafts = Record<string, string>;

/**
 * 一条 run 的身份键（草稿表用它索引）。
 *
 * @param target 定位键。
 * @returns `页号|内容流引用|Form 链|起始下标` 拼出的串，与引擎的现场核对同一条口径。
 */
export function pdfTargetKey(target: PdfTextTarget): string {
  return `${target.pageIndex}|${target.streamRef}|${target.withinForm ?? ""}|${target.start}`;
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
 * 这条 run 能不能改。
 *
 * @param run 读模型里的一条显示文字。
 * @returns 内联流与解不出原文的 run 都回 false——前者没有可标记的对象，后者连旧文字都不知道（§2.5 的撤销要靠它）。
 */
export function isRunEditable(run: PdfTextRun): boolean {
  return run.streamRef !== null && run.text !== null;
}

/**
 * 改不了的 run 给一句中文理由。
 *
 * @param run 一条不可编辑的显示文字。
 * @returns 说明卡在哪里的短句；可编辑时回空串（调用方据此不显示）。
 */
export function readOnlyReason(run: PdfTextRun): string {
  if (run.streamRef === null) return "文字内联在页面里，没有独立对象可标记，改不了";
  if (run.text === null) return `认不出这段文字念作什么（字体 ${run.baseFont ?? "?"}${run.composite ? "，CID" : ""}），原文都不知道就不改`;
  return "";
}

/**
 * 取某一页的全部 run。
 *
 * @param model 读模型。
 * @param pageIndex 页号（0 基）。
 * @returns 该页的 run，按内容流与起始下标的原序（读模型已是这个序）。
 */
export function runsOfPage(model: PdfReadModel, pageIndex: number): PdfTextRun[] {
  return model.runs.filter((run) => run.pageIndex === pageIndex);
}

/**
 * 把草稿表里真正改过、且引擎收得到定位键的那几条收成一次请求。
 *
 * @param model 当前读模型（草稿只对这份模型里的 run 有意义）。
 * @param drafts 草稿表。
 * @returns 交给 `pdf:edit` 的改动清单，顺序按页与下标；空数组表示没有待保存的改动。
 */
export function pendingPdfEdits(model: PdfReadModel, drafts: PdfDrafts): PdfTextEdit[] {
  const edits: PdfTextEdit[] = [];
  for (const run of model.runs) {
    const target = targetOfRun(run);
    if (!target || !isRunEditable(run)) continue;
    const draft = drafts[pdfTargetKey(target)];
    if (draft === undefined || draft === run.text) continue;
    edits.push({ target, text: draft });
  }
  return edits;
}

/**
 * 一条草稿的字节预算提示（只作参考，真正的判定在引擎）。
 *
 * @param run 一条可编辑的显示文字。
 * @returns 一句中文，告诉用户"改完得装得回这么大的位置"，PDF 没有回流（spec D8）。
 */
export function byteBudgetNote(run: PdfTextRun): string {
  return `原位只有 ${run.encodedBytes} 字节，改长了装不回（PDF 不回流）`;
}

/** 引擎的错误码 → 回给人听的那句中文（文案留在渲染层，§5.7）。 */
export const PDF_FAILURE_LABEL: Record<PdfEditFailure["reason"], string> = {
  "run-not-found": "那段文字已经不在原来的位置了",
  "no-encoding": "认不出那段文字念作什么",
  "char-outside-subset": "文件里的字体没带这个字，补不出来",
  "length-mismatch": "改完的文字装不回原来的字节数",
  "text-drifted": "那段文字后来又被改过，不能拿旧内容覆盖",
};

/**
 * 把一批失败原因汇成一句 toast 文案。
 *
 * @param failures 引擎回包里的失败清单。
 * @returns 去重后的中文理由，用「；」连接。
 */
export function pdfFailureText(failures: PdfEditFailure[]): string {
  return [...new Set(failures.map((failure) => PDF_FAILURE_LABEL[failure.reason]))].join("；");
}
