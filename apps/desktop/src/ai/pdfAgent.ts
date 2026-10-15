/**
 * PDF 的 AI 代理（spec 阶段 5 同一条形状，pdf 一线 / 决策 D6）：把"把这页标题里的 2026 换成 2027"这类诉求
 * 跑成一串工具回合，落地只认 `pdf_submit_proposal` 交出的那批改动，并且**只交提案，不写文件**（`AGENTS.md` §2.8）。
 *
 * 与 `slideAgent.ts`/`sheetAgent.ts` 同一条形状，三条约束照抄：
 * ① 页数、每页 run 数、逐条原文与字数这些事实由 `@office-copilot/pdf-io` 从现场模型读（`describePdf`/`readPdfPageText`），这里只搬运（§2.11）；
 * ② 上送内容默认截断（§2.9）：页数与每页 run 数都有上限，超出部分引擎会标出来；
 * ③ 模型给的非法入参与越界页号都换成回给它的一句话，永不抛异常（§7.1 的 mock 用例逐条钉住）。
 *
 * 交出去的 `edits` 就是视图"保存"用的那批 `PdfTextEdit`（`pdf-io/pdf-file.ts`）：确认后原样送 `pdf:edit`，
 * 主进程做增量写盘并回 `applied`（撤销凭据）——所以这里不需要、也不许再长第二套落地或撤销通道（§3.2）。
 * 模型视角的坐标只有**页号（1 基）与页内 run 序号（0 基）**；内容流引用、Form 链、起始下标由本文件的
 * `pdfRequestsOfProposal` 按现场模型翻出来，这是唯一一处做这个映射（§3.3）。
 */
import {
  describePdf,
  isRunEditable,
  readPdfPageText,
  targetOfRun,
  type PdfReadModel,
  type PdfTextEdit,
} from "@office-copilot/pdf-io";
import { parseToolArguments, PDF_TOOLS, type PdfEditT, type PdfToolName, type SubmitProposalArgsT } from "@office-copilot/pdf-tools";
import type { ToolCall, ToolTurnMessage, ToolTurnRequest, ToolTurnResult } from "../global.js";
import { readOnlyReason, runsOfPage } from "../pdf/pdfRuns.js";

/** 一次任务里最多转多少轮：模型无限要工具就该停下来问人，而不是无限烧 token。 */
const MAX_TOOL_TURNS = 8;

/** 首条用户消息里那份画像的上限：给模型看得见全貌，但不许把整份 PDF 摊上去（§2.9）。 */
const CONTEXT_MAX_PAGES = 20;

const PDF_SYSTEM = `你是 PDF（.pdf）代理，看得见的是这份文档的结构画像与逐条显示文字，不是内容流原文。
- 动手前先调 pdf_describe 拿页号；要改某条文字再调 pdf_read_page_text 确认 runIndex 与那条的 charCount（原文字数），不许凭印象数，也不许自己编坐标。
- 要改只能调 pdf_submit_proposal，条目只有一种 set_run_text，指到"第几页第几条显示文字"，把那条整条换成新文字。
- 硬约束是**新文字的字数必须与 charCount 完全相同**：PDF 没有回流，长度变了装不回原来的位置，整批一个字节都不会写。
  需要短一点就自己删字凑等长、需要留白就用空格补位，凑不出等长的说法就别提这条，直接回用户做不了。
- 新文字里每个字都得是这一页已经出现过的字——文件里的字体是子集，没带的字补不出来（这一条由引擎判定，但你能先避开）。
- editable=false 的那条改不了（内联内容流 / 认不出原文），别对着它提改动。
- 改字体、改颜色、改几何、删页、把跨条 run 拼成一句完整的话再改，这些现在都做不了，遇到直接说做不了并解释原因。
- 换算字节、编内容流引用这类事不归你管，只交给页号与 run 序号。
- 纯分析的问题直接回答，用中文，带上你实际读到的页号与文字，不要为了回答去提提案。`;

/** `runPdfAgent` 的输入。 */
export interface PdfAgentInput {
  /** 主进程读出来的当前模型（只读：代理不改它，落地由上层确认后走 `pdf:edit`）。 */
  doc: PdfReadModel;
  /** 用户那句诉求原文。 */
  instruction: string;
  /** 已经跑过的 transcript：暂停后继续时原样带回。 */
  messages?: ToolTurnMessage[];
  /** 指定服务商；省略即用设置里激活的那个。 */
  providerId?: string;
  /** 发一轮请求的通道，默认走 preload 的 `ai:tool-turn`，单测里换成桩。 */
  ask?: (payload: ToolTurnRequest) => Promise<ToolTurnResult>;
  /** 用户按了暂停：下一轮请求发出前停下，transcript 原样带回。 */
  isPaused?: () => boolean;
}

/** 一次 PDF 任务的结果。`messages` 每支都带，UI 靠它接着往下跑。 */
export type PdfAgentOutcome =
  | {
      kind: "proposal";
      summary: string;
      /** 翻成引擎坐标的那批改动（定位键 + 新文字），确认后原样送 `pdf:edit` 即可落地并拿到撤销凭据（§2.5）。 */
      edits: PdfTextEdit[];
      /** 提案里没能变成改动的条目，逐条中文理由，不许静默丢掉（§2.5）。 */
      notes: string[];
      /** 模型原样的提案条目（给用户看"它到底提了什么"，坐标是模型视角）。 */
      proposed: PdfEditT[];
      messages: ToolTurnMessage[];
    }
  | { kind: "answer"; text: string; messages: ToolTurnMessage[] }
  | { kind: "paused"; messages: ToolTurnMessage[] }
  | { kind: "failed"; message: string; needsSetup?: boolean; messages: ToolTurnMessage[] };

/**
 * 组装上送的文档画像：页数 + 每页 run 数与文字概要（截断由引擎做，这里不再二次裁剪）。
 *
 * @param doc 当前模型。
 * @returns 多行中文文本，作为本轮用户消息的前缀。
 */
export function pdfContextOf(doc: PdfReadModel): string {
  const outline = describePdf(doc, { maxPages: CONTEXT_MAX_PAGES });
  const lines = [`PDF：${outline.pageCount} 页，其中 ${outline.pages.reduce((sum, page) => sum + page.editableCount, 0)} 条显示文字可改。`, `画像：${JSON.stringify(outline.pages)}`];
  if (outline.pagesTruncated) lines.push(`页太多，只描述了前 ${outline.pages.length} 页，剩下的页要用 pdf_describe 加大 maxPages 才看得见。`);
  return lines.join("\n");
}

/**
 * 把提案里的条目翻成引擎认的跨界改动（内容流引用 + Form 链 + 起始下标）。
 *
 * 页号是模型视角的 1 基，run 序号是该页显示文字的 0 基序号——两者都按现场模型算，这是唯一一处做这个映射（§3.3）。
 * 翻不动的（页号越界、该页没有这么条 run、那条本来改不了、字数不等装不回去、与原文一模一样）逐条记进 `notes`，
 * 不许交一批主进程必然拒绝、整批不落地的改动（§2.4 的事务语义）。
 *
 * @param doc 提案要落在这份文档上（页号 → 内容流坐标按它算）。
 * @param edits `pdf_submit_proposal` 校验过的条目。
 * @returns `edits` 按提案顺序排好；`notes` 是没变成改动的条目的中文说明。
 */
export function pdfRequestsOfProposal(doc: PdfReadModel, edits: PdfEditT[]): { requests: PdfTextEdit[]; notes: string[] } {
  const requests: PdfTextEdit[] = [];
  const notes: string[] = [];
  for (const edit of edits) {
    if (edit.pageNumber > doc.pageCount) {
      notes.push(`没有第 ${edit.pageNumber} 页（这份 PDF 共 ${doc.pageCount} 页），这条跳过。`);
      continue;
    }
    const runs = runsOfPage(doc, edit.pageNumber - 1);
    const run = runs[edit.runIndex];
    if (!run) {
      notes.push(`第 ${edit.pageNumber} 页只有 ${runs.length} 条显示文字，没有第 ${edit.runIndex} 条，这条跳过（用 pdf_read_page_text 核对 runIndex）。`);
      continue;
    }
    if (!isRunEditable(run)) {
      notes.push(`第 ${edit.pageNumber} 页第 ${edit.runIndex} 条改不了：${readOnlyReason(run)}`);
      continue;
    }
    const target = targetOfRun(run);
    if (!target) {
      notes.push(`第 ${edit.pageNumber} 页第 ${edit.runIndex} 条没有可标记的内容流对象，这条跳过。`);
      continue;
    }
    if (edit.text === run.text) {
      notes.push(`第 ${edit.pageNumber} 页第 ${edit.runIndex} 条的新文字与原文一样，不用改，这条跳过。`);
      continue;
    }
    // 等长这条由引擎最终判定（字节才算得准）；这里按模型视角的字数先筛掉必然整批失败的条目（§2.4）。
    if ([...edit.text].length !== [...(run.text ?? "")].length) {
      notes.push(`第 ${edit.pageNumber} 页第 ${edit.runIndex} 条原文是 ${[...(run.text ?? "")].length} 个字，新文字是 ${[...edit.text].length} 个字：PDF 不回流，字数不等装不回原位置，这条跳过。`);
      continue;
    }
    requests.push({ target, text: edit.text });
  }
  return { requests, notes };
}

/**
 * 跑一次工具调用：入参先过 `parseToolArguments`，再交给 `pdf-io` 算。
 *
 * @param doc 当前模型（只读）。
 * @param call 模型点名的工具与入参原文。
 * @returns 回给模型的文本；`proposal` 非空表示模型交了提案，回路该收尾。
 */
function executeToolCall(doc: PdfReadModel, call: ToolCall): { content: string; proposal?: SubmitProposalArgsT } {
  switch (call.name as PdfToolName) {
    case "pdf_describe": {
      const parsed = parseToolArguments("pdf_describe", call.arguments);
      return { content: parsed.ok ? JSON.stringify(describePdf(doc, parsed.args)) : parsed.reason };
    }
    case "pdf_read_page_text": {
      const parsed = parseToolArguments("pdf_read_page_text", call.arguments);
      if (!parsed.ok) return { content: parsed.reason };
      const read = readPdfPageText(doc, parsed.args.pageNumber, parsed.args.maxRunsPerPage);
      return { content: read.ok ? JSON.stringify(read.report) : read.reason };
    }
    case "pdf_submit_proposal": {
      const parsed = parseToolArguments("pdf_submit_proposal", call.arguments);
      if (!parsed.ok) return { content: parsed.reason };
      return {
        content: `提案已收下（${parsed.args.edits.length} 条改动），等用户确认后才会落地。要继续核对就接着调读工具，别重复提交。`,
        proposal: parsed.args,
      };
    }
    default:
      return { content: `不认识的工具：${call.name}。可用面只有 pdf_describe / pdf_read_page_text / pdf_submit_proposal。` };
  }
}

/**
 * 跑一次 PDF 任务：转工具回合，直到模型交提案、直接作答、被暂停或出错。
 *
 * @param input 见 `PdfAgentInput`；`doc` 全程只读。
 * @returns `PdfAgentOutcome`。没配服务商、HTTP 失败、超过 8 轮都走 `failed`，不抛异常。
 */
export async function runPdfAgent(input: PdfAgentInput): Promise<PdfAgentOutcome> {
  const ask = input.ask ?? ((payload: ToolTurnRequest) => window.officeCopilot.aiToolTurn(payload));
  const messages: ToolTurnMessage[] = input.messages?.length
    ? input.messages.map((message) => ({ ...message }))
    : [
        { role: "system", content: PDF_SYSTEM },
        { role: "user", content: `${pdfContextOf(input.doc)}\n\n诉求：${input.instruction}` },
      ];

  for (let turn = 0; turn < MAX_TOOL_TURNS; turn += 1) {
    if (input.isPaused?.()) return { kind: "paused", messages };
    const result = await ask({ messages, tools: PDF_TOOLS, providerId: input.providerId });
    if (!result.ok) return { kind: "failed", message: result.message, needsSetup: result.needsSetup, messages };
    // 模型没点工具就是说完话了：这一轮的正文即为答复（纯分析类任务走这里，不出提案）。
    if (!result.toolCalls.length) return { kind: "answer", text: result.content, messages };

    messages.push({
      role: "assistant",
      ...(result.content ? { content: result.content } : {}),
      tool_calls: result.toolCalls.map((call) => ({
        id: call.id,
        type: "function" as const,
        function: { name: call.name, arguments: call.arguments },
      })),
    });
    for (const call of result.toolCalls) {
      const executed = executeToolCall(input.doc, call);
      messages.push({ role: "tool", tool_call_id: call.id, content: executed.content });
      if (!executed.proposal) continue;
      const { requests, notes } = pdfRequestsOfProposal(input.doc, executed.proposal.edits);
      return { kind: "proposal", summary: executed.proposal.summary, edits: requests, notes, proposed: executed.proposal.edits, messages };
    }
  }
  return { kind: "failed", message: `工具回合到了 ${MAX_TOOL_TURNS} 轮上限，模型还没交提案或给出结论，先停下。`, messages };
}
