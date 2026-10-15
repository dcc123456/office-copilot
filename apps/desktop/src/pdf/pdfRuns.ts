/**
 * PDF 视图与 AI 代理共用的那几个纯函数（组件只做装配，文案与攒草稿放这里，`AGENTS.md` §3.3）。
 *
 * "这条 run 能不能改、卡在哪一类"的判定只有一份，在 `@office-copilot/pdf-io` 的 `runEditBlock`（内联内容流与
 * 解不出原文那两种）；这里只把它的码翻成中文。**字节预算不在这里判**：编码规则（WinAnsi / CID 每字 2 字节）属于
 * `pdf-io`，视图再算一遍就是第二套换算（§2.1），所以草稿交给引擎、由它回 `length-mismatch`。
 */
import {
  isRunEditable,
  pdfTargetKey,
  runEditBlock,
  targetOfRun,
  type PdfEditFailure,
  type PdfReadModel,
  type PdfTextEdit,
  type PdfTextRun,
} from "@office-copilot/pdf-io/pure";

/** 一条草稿在渲染层里的样子：`key` 用引擎的 `pdfTargetKey`，`text` 是用户敲的新文字。 */
export type PdfDrafts = Record<string, string>;

/** 引擎的"改不了"类别 → 回给人听的那句中文开头（文案留在渲染层，§5.7）。 */
export const PDF_BLOCK_LABEL: Record<NonNullable<ReturnType<typeof runEditBlock>>, string> = {
  "inline-stream": "文字内联在页面里，没有独立对象可标记，改不了",
  undecoded: "认不出这段文字念作什么",
};

/** 引擎的错误码 → 回给人听的那句中文（文案留在渲染层，§5.7）。 */
export const PDF_FAILURE_LABEL: Record<PdfEditFailure["reason"], string> = {
  "run-not-found": "那段文字已经不在原来的位置了",
  "no-encoding": "认不出那段文字念作什么",
  "char-outside-subset": "文件里的字体没带这个字，补不出来",
  "length-mismatch": "改完的文字装不回原来的字节数",
  "text-drifted": "那段文字后来又被改过，不能拿旧内容覆盖",
};

/**
 * 改不了的 run 给一句中文理由。
 *
 * @param run 一条显示文字。
 * @returns 说明卡在哪里的短句；可编辑时回空串（调用方据此不显示）。CID 字体在理由里点名，因为它连"缺 `/ToUnicode`"
 *   还是"子集码表没这条"都表现成同一个码，得让用户自己判断（spec D8）。
 */
export function readOnlyReason(run: PdfTextRun): string {
  const block = runEditBlock(run);
  if (!block) return "";
  if (block === "inline-stream") return PDF_BLOCK_LABEL[block];
  return `${PDF_BLOCK_LABEL[block]}（字体 ${run.baseFont ?? "?"}${run.composite ? "，CID" : ""}），原文都不知道就不改`;
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

/**
 * 把一批失败原因汇成一句 toast 文案。
 *
 * @param failures 引擎回包里的失败清单。
 * @returns 去重后的中文理由，用「；」连接。
 */
export function pdfFailureText(failures: PdfEditFailure[]): string {
  return [...new Set(failures.map((failure) => PDF_FAILURE_LABEL[failure.reason]))].join("；");
}
