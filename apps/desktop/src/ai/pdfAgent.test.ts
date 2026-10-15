/**
 * PDF 代理的用例（`pdfAgent.ts`，spec 阶段 7 的 AI 那一半）：工具回合怎么转、提案翻成引擎坐标、非法入参与越界怎么变成回给模型的话。
 *
 * `ask` 一律打桩（`AGENTS.md` §7.1 禁止测试碰真实端点）；真实端点行为由一次性探针负责（§6.5）。
 * 这里钉的是两条红线：**模型给的页号/run 序号对不上现场模型时，那一条不许交出去**——主进程的事务语义是
 * "任一条失败就整批不写"（§2.4），代理交一批注定失败的改动等于让用户白点一次确认；以及
 * **内容流引用与起始下标只能由代理按现场模型填**，模型自己报的坐标一律不进请求。
 */
import { describe, expect, it } from "vitest";
import type { PdfReadModel, PdfTextRun } from "@office-copilot/pdf-io";
import type { ToolTurnRequest, ToolTurnResult } from "../global.js";
import { pdfContextOf, pdfRequestsOfProposal, runPdfAgent } from "./pdfAgent.js";

/** 造一条 run。 */
function run(overrides: Partial<PdfTextRun> = {}): PdfTextRun {
  return {
    pageIndex: 0,
    streamRef: "6 0 R",
    withinForm: null,
    operator: "Tj",
    form: "literal",
    start: 32,
    encodedBytes: 21,
    tokenLength: 23,
    baseFont: "Helvetica",
    composite: false,
    text: "Quarterly Report 2026",
    ...overrides,
  };
}

/** 两页夹具：第 1 页两条可改（第二条在 Form 里），第 2 页一条内联（改不了）+ 一条认不出原文。 */
function sampleDoc(): PdfReadModel {
  return {
    pageCount: 2,
    runs: [
      run(),
      run({ start: 60, withinForm: "Fm0", text: "Appendix A", encodedBytes: 10, tokenLength: 12 }),
      run({ pageIndex: 1, streamRef: null, text: "内联的文字" }),
      run({ pageIndex: 1, start: 12, text: null, baseFont: null }),
    ],
  };
}

/** 依次吐出这些回包；用完就一直要工具，用来钉轮数上限。 */
function stubAsk(replies: ToolTurnResult[]): { ask: (payload: ToolTurnRequest) => Promise<ToolTurnResult>; requests: ToolTurnRequest[] } {
  const requests: ToolTurnRequest[] = [];
  return {
    requests,
    ask: async (payload) => {
      requests.push(payload);
      return replies.shift() ?? { ok: true, content: "", toolCalls: [{ id: `call_loop_${requests.length}`, name: "pdf_no_such_tool", arguments: "{}" }] };
    },
  };
}

/** 模型点一个工具的一轮回包。 */
function turnWithTool(name: string, args: string, content = ""): ToolTurnResult {
  return { ok: true, content, toolCalls: [{ id: `call_${name}`, name, arguments: args }] };
}

/** 模型直接答一句话、不点工具的一轮回包。 */
function turnWithAnswer(content: string): ToolTurnResult {
  return { ok: true, content, toolCalls: [] };
}

describe("pdfContextOf：上送的是画像不是原文", () => {
  it("报页数与可改条数，画像里只有页号与概要，没有内容流坐标", () => {
    const context = pdfContextOf(sampleDoc());
    expect(context).toContain("PDF：2 页");
    expect(context).toContain("2 条显示文字可改");
    expect(context).toContain('"pageNumber":1');
    expect(context).not.toContain("6 0 R");
    expect(context).not.toContain("streamRef");
  });
});

describe("pdfRequestsOfProposal：模型视角→引擎坐标", () => {
  it("合法那条翻出内容流引用与起始下标，Form 链原样带上", () => {
    const { requests, notes } = pdfRequestsOfProposal(sampleDoc(), [
      { type: "set_run_text", pageNumber: 1, runIndex: 1, text: "Appendix B" },
    ]);
    expect(notes).toEqual([]);
    expect(requests).toEqual([{ target: { pageIndex: 0, streamRef: "6 0 R", withinForm: "Fm0", start: 60 }, text: "Appendix B" }]);
  });

  it("越界页号、越界序号、改不了的那两种、字数不等、与原文一样：逐条给理由，一条都不交", () => {
    const { requests, notes } = pdfRequestsOfProposal(sampleDoc(), [
      { type: "set_run_text", pageNumber: 9, runIndex: 0, text: "一样长啊哈" },
      { type: "set_run_text", pageNumber: 1, runIndex: 7, text: "一样长啊哈" },
      { type: "set_run_text", pageNumber: 2, runIndex: 0, text: "一样长啊哈" },
      { type: "set_run_text", pageNumber: 2, runIndex: 1, text: "一样长啊哈" },
      { type: "set_run_text", pageNumber: 1, runIndex: 0, text: "装不下下下下" },
      { type: "set_run_text", pageNumber: 1, runIndex: 0, text: "Quarterly Report 2026" },
    ]);
    expect(requests).toEqual([]);
    expect(notes).toHaveLength(6);
    expect(notes[0]).toContain("没有第 9 页");
    expect(notes[1]).toContain("只有 2 条显示文字");
    expect(notes[2]).toContain("内联");
    expect(notes[3]).toContain("认不出这段文字念作什么");
    expect(notes[4]).toContain("原文是 21 个字，新文字是 6 个字");
    expect(notes[5]).toContain("与原文一样");
  });
});

describe("runPdfAgent：工具回合与收口", () => {
  it("交提案那一轮就收尾，edits 是引擎坐标、proposed 原样留着给界面显示", async () => {
    const ask = stubAsk([turnWithTool("pdf_submit_proposal", '{"summary":"把年份换掉","edits":[{"type":"set_run_text","pageNumber":1,"runIndex":0,"text":"Quarterly Report 2027"}]}')]);
    const outcome = await runPdfAgent({ doc: sampleDoc(), instruction: "把标题的年份换成 2027", ask: ask.ask });
    expect(outcome.kind).toBe("proposal");
    if (outcome.kind !== "proposal") return;
    expect(outcome.summary).toBe("把年份换掉");
    expect(outcome.edits).toEqual([{ target: { pageIndex: 0, streamRef: "6 0 R", withinForm: null, start: 32 }, text: "Quarterly Report 2027" }]);
    expect(outcome.proposed[0]).toMatchObject({ pageNumber: 1, runIndex: 0 });
    expect(outcome.notes).toEqual([]);
    // transcript 末尾是那条工具回包，说明代理确实执行了工具而不是直接抄模型的话。
    expect(outcome.messages.at(-1)?.role).toBe("tool");
    expect(ask.requests).toHaveLength(1);
    expect(JSON.stringify(ask.requests[0]!.tools)).toContain("pdf_read_page_text");
  });

  it("先读工具再交提案：读工具的回包来自引擎（带 charCount），模型照它写等长文字", async () => {
    const ask = stubAsk([
      turnWithTool("pdf_read_page_text", '{"pageNumber":1}'),
      turnWithTool("pdf_submit_proposal", '{"summary":"换个说法","edits":[{"type":"set_run_text","pageNumber":1,"runIndex":1,"text":"Appendix C"}]}'),
    ]);
    const outcome = await runPdfAgent({ doc: sampleDoc(), instruction: "把附录那条改一下", ask: ask.ask });
    expect(outcome.kind).toBe("proposal");
    if (outcome.kind !== "proposal") return;
    const toolMessage = outcome.messages.find((message) => message.role === "tool");
    expect(toolMessage?.content).toContain('"charCount":10');
    expect(toolMessage?.content).toContain('"runIndex":1');
    expect(outcome.edits).toHaveLength(1);
  });

  it("非法入参回给模型一句理由，不抛异常也不结束回合", async () => {
    const ask = stubAsk([
      turnWithTool("pdf_read_page_text", "{}"),
      turnWithAnswer("第 2 页那两条都改不了：一条内联在页面里，一条认不出原文。"),
    ]);
    const outcome = await runPdfAgent({ doc: sampleDoc(), instruction: "改第 2 页", ask: ask.ask });
    expect(outcome.kind).toBe("answer");
    if (outcome.kind !== "answer") return;
    expect(outcome.text).toContain("改不了");
    expect(outcome.messages.find((message) => message.role === "tool")?.content).toContain("pageNumber");
  });

  it("模型点了不存在的方法名时回一句可用面，继续往下跑", async () => {
    const ask = stubAsk([turnWithTool("pdf_edit_page", "{}"), turnWithAnswer("这个做不了。")]);
    const outcome = await runPdfAgent({ doc: sampleDoc(), instruction: "删掉第 2 页", ask: ask.ask });
    expect(outcome.kind).toBe("answer");
    if (outcome.kind !== "answer") return;
    expect(outcome.messages.find((message) => message.role === "tool")?.content).toContain("不认识的工具");
  });

  it("没配服务商时把 needsSetup 原样带回去，不抛异常", async () => {
    const ask = stubAsk([{ ok: false, message: "还没配置服务商", needsSetup: true }]);
    const outcome = await runPdfAgent({ doc: sampleDoc(), instruction: "改标题", ask: ask.ask });
    expect(outcome.kind).toBe("failed");
    if (outcome.kind !== "failed") return;
    expect(outcome.needsSetup).toBe(true);
    expect(outcome.message).toContain("还没配置服务商");
  });

  it("暂停在发请求之前：transcript 原样带回，接着跑就用它", async () => {
    const ask = stubAsk([turnWithAnswer("不该跑到这里")]);
    const outcome = await runPdfAgent({ doc: sampleDoc(), instruction: "改标题", ask: ask.ask, isPaused: () => true });
    expect(outcome.kind).toBe("paused");
    expect(ask.requests).toHaveLength(0);
    expect(outcome.messages).toHaveLength(2);
  });

  it("模型无限要工具时到 8 轮上限停下并给一句理由", async () => {
    const ask = stubAsk([]);
    const outcome = await runPdfAgent({ doc: sampleDoc(), instruction: "把整份改一遍", ask: ask.ask });
    expect(outcome.kind).toBe("failed");
    if (outcome.kind !== "failed") return;
    expect(outcome.message).toContain("8 轮上限");
    expect(ask.requests).toHaveLength(8);
  });
});
