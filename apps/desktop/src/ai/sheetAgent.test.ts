import { afterEach, describe, expect, it, vi } from "vitest";
import { columnWidthCharsToPx, runSheetOps, type SheetPayload } from "@office-copilot/sheet-io";
import { SHEET_TOOLS, type SheetEditT } from "@office-copilot/sheet-tools";
import type { ToolTurnRequest, ToolTurnResult } from "../global.js";
import { runSheetAgent, sheetContextOf, sheetOpsOfProposal } from "./sheetAgent.js";

/**
 * 表格代理的用例：上下文怎么截断、工具回合怎么转、提案翻成 op、非法入参怎么变成回给模型的话。
 *
 * `ask` 一律打桩（`AGENTS.md` §7.1 禁止测试碰真实端点）；真实端点由一次性探针负责（§6.5）。
 */

afterEach(() => {
  vi.unstubAllGlobals();
});

/** 一张四行数据的小表：表头 + 三个区域，最后一行故意与上一行重复，好测去重。 */
function sampleSheet(): SheetPayload {
  const cell = (address: string, colIndex: number, rowIndex: number, value: string | number): SheetPayload["cells"][string] => ({
    address,
    colIndex,
    rowIndex,
    kind: typeof value === "number" ? "number" : "string",
    value,
  });
  return {
    name: "销售",
    sheetId: 1,
    partPath: "xl/worksheets/sheet1.xml",
    lastColIndex: 2,
    lastRowIndex: 4,
    cells: {
      A1: cell("A1", 0, 0, "月份"),
      B1: cell("B1", 1, 0, "区域"),
      C1: cell("C1", 2, 0, "销售额"),
      A2: cell("A2", 0, 1, "1月"),
      B2: cell("B2", 1, 1, "华东"),
      C2: cell("C2", 2, 1, 100),
      A3: cell("A3", 0, 2, "2月"),
      B3: cell("B3", 1, 2, "华北"),
      C3: cell("C3", 2, 2, 120),
      A4: cell("A4", 0, 3, "3月"),
      B4: cell("B4", 1, 3, "华东"),
      C4: cell("C4", 2, 3, 150),
      A5: cell("A5", 0, 4, "3月"),
      B5: cell("B5", 1, 4, "华东"),
      C5: cell("C5", 2, 4, 150),
    },
    merges: [],
    columns: [{ minColIndex: 0, maxColIndex: 0, widthChars: 12, customWidth: true }],
    rows: [],
    warnings: [],
  };
}

/** 模型点一个工具的一轮回包。 */
function turnWithTool(name: string, args: string, content = ""): ToolTurnResult {
  return { ok: true, content, toolCalls: [{ id: `call_${name}`, name, arguments: args }] };
}

describe("sheetContextOf：上送的只有摘要与截断过的样本", () => {
  it("首行当表头：行数不含表头，画像与样本都带进来", () => {
    const context = sheetContextOf(sampleSheet(), true);
    expect(context).toContain("工作表「销售」：4 行 × 3 列，已用区域 A1:C5。");
    expect(context).toContain('"header":"销售额"');
    expect(context).toContain("3月");
    // 五行的表在 50 行上限内，不该出现截断说明。
    expect(context).not.toContain("样本只到第");
  });

  it("超过 50 行时样本收住，并告诉模型后面要接着读", () => {
    const wide = sampleSheet();
    for (let rowIndex = 5; rowIndex < 70; rowIndex += 1) {
      wide.cells[`A${rowIndex + 1}`] = { address: `A${rowIndex + 1}`, colIndex: 0, rowIndex, kind: "string", value: `行${rowIndex}` };
      wide.lastRowIndex = rowIndex;
    }
    const context = sheetContextOf(wide, true);
    expect(context).toContain("样本只到第 50 行");
    expect(context).toContain("共 70 行");
  });
});

describe("sheetOpsOfProposal：只有能安全撤销的四类才变成 op", () => {
  it("改值 / 列宽 / 删行各成一条 op，按现场核对全部落得下", () => {
    const sheet = sampleSheet();
    const edits: SheetEditT[] = [
      { type: "set_cell", address: "d2", text: "=C2*2" },
      { type: "set_column_width", column: "区域", widthPx: 130 },
      { type: "delete_rows", rows: [5, 4] },
    ];
    const { ops, notes } = sheetOpsOfProposal(sheet, edits);
    expect(ops.map((op) => op.kind)).toEqual(["setCellValue", "setColumnWidth", "deleteTracks", "deleteTracks"]);
    expect(notes).toEqual([]);

    const result = runSheetOps(sheet, ops);
    expect(result.ok).toBe(true);
    expect(result.skipped).toEqual([]);
    expect(result.ops).toHaveLength(4);
    // 从大到小删：两条 deleteTracks 都落在自己那一行上，第二行的公式没被牵连。
    expect(result.after.cells.D2?.formula).toBe("C2*2");
  });

  it("提案里删掉公式正引用的行时，引擎的现场核对会挡住，代理不自己猜", () => {
    const sheet = sampleSheet();
    const { ops } = sheetOpsOfProposal(sheet, [
      { type: "set_cell", address: "C6", text: "=SUM(C2:C5)" },
      { type: "delete_rows", rows: [5] },
    ]);
    const result = runSheetOps(sheet, ops);
    expect(result.ops).toHaveLength(1);
    expect(result.skipped[0]?.reason).toContain("#REF!");
  });

  it("认不出的列与写不了的地址逐条说明，不静默丢掉半份提案", () => {
    const sheet = sampleSheet();
    const { ops, notes } = sheetOpsOfProposal(sheet, [
      { type: "set_column_width", column: "不存在列", widthPx: 120 },
      { type: "set_cell", address: "不是地址", text: "1" },
    ]);
    expect(ops).toEqual([]);
    expect(notes).toEqual(["认不出列「不存在列」，那条列宽跳过。", "不是地址 那一格写不了，已跳过。"]);
  });

  it("宽度本来就一样时不算改动：既不出 op 也不报错", () => {
    const sheet = sampleSheet();
    // A 列存的是 12 个字符宽，用它自己换算出来的 px 提回去，等于没改。
    expect(sheetOpsOfProposal(sheet, [{ type: "set_column_width", column: "A", widthPx: columnWidthCharsToPx(12) }])).toEqual({
      ops: [],
      notes: [],
    });
  });
});

describe("runSheetAgent：工具回合", () => {
  it("没点工具就是直接作答，上下文里带着摘要与那句诉求", async () => {
    const ask = vi.fn(async (_payload: ToolTurnRequest): Promise<ToolTurnResult> => ({
      ok: true,
      content: "销售额最高的是 3 月，150。",
      toolCalls: [],
    }));
    const outcome = await runSheetAgent({ sheet: sampleSheet(), instruction: "哪个月卖得最好？", ask });

    expect(outcome.kind).toBe("answer");
    expect(outcome.kind === "answer" && outcome.text).toBe("销售额最高的是 3 月，150。");
    const request = ask.mock.calls[0]?.[0];
    expect(request?.messages.map((message) => message.role)).toEqual(["system", "user"]);
    expect(request?.messages[1]?.content).toContain("诉求：哪个月卖得最好？");
    expect(request?.tools).toBe(SHEET_TOOLS);
  });

  it("读工具的回包原样进 transcript，交提案时停下并给出 op", async () => {
    const sheet = sampleSheet();
    const script: ToolTurnResult[] = [
      turnWithTool("sheet_find", '{"column":"销售额","criteria":">130"}'),
      turnWithTool(
        "sheet_submit_proposal",
        '{"summary":"删掉重复的 3 月记录","edits":[{"type":"delete_rows","rows":[5]}]}',
      ),
    ];
    const ask = vi.fn(async () => script.shift()!);
    const outcome = await runSheetAgent({ sheet, instruction: "把重复的 3 月那行删掉", ask });

    expect(outcome.kind).toBe("proposal");
    if (outcome.kind !== "proposal") return;
    expect(outcome.summary).toBe("删掉重复的 3 月记录");
    expect(outcome.ops.map((op) => op.kind)).toEqual(["deleteTracks"]);
    expect(outcome.notes).toEqual([]);
    expect(outcome.messages.map((message) => message.role)).toEqual([
      "system",
      "user",
      "assistant",
      "tool",
      "assistant",
      "tool",
    ]);
    // 回给模型的命中行带着真实行号，模型不必自己数。
    expect(outcome.messages[3]?.content).toContain('"rowNumber":4');
    // 续跑时要把 assistant 点名的工具原样带回，否则网关认不出这条 tool 回复属于谁。
    expect(outcome.messages[2]?.tool_calls?.[0]?.function.name).toBe("sheet_find");
  });

  it("出图那条：模型只给区域，数值由引擎从现场表读，回合停在图表提案上", async () => {
    const script: ToolTurnResult[] = [
      turnWithTool("sheet_create_chart", '{"kind":"bar","categories":"A2:A4","series":[{"values":"C2:C4","name":"C1"}],"anchor":"E2"}'),
    ];
    const outcome = await runSheetAgent({ sheet: sampleSheet(), instruction: "把每月销售额画成柱状图", ask: async () => script.shift()! });
    if (outcome.kind !== "chart") throw new Error(`期望 chart，实际 ${outcome.kind}`);
    expect(outcome.chart.series[0]!.values).toEqual([100, 120, 150]);
    expect(outcome.chart.series[0]!.name).toBe("销售额");
    expect(outcome.chart.categoryRef).toBe("'销售'!$A$2:$A$4");
    expect(outcome.chart.anchor).toEqual({ colIndex: 4, rowIndex: 1 });
    // 回给模型的那句话里带着引擎读到的引用，模型不必也不许自己报数
    const toolMessage = outcome.messages.at(-1)!;
    expect(JSON.stringify(toolMessage)).toContain("'销售'!$C$2:$C$4");
  });

  it("区域取不到数时把理由回给模型、回路继续，不许当成图表交出去", async () => {
    const script: ToolTurnResult[] = [
      turnWithTool("sheet_create_chart", '{"kind":"bar","categories":"A2:A4","series":[{"values":"B2:B4"}],"anchor":"E2"}'),
      { ok: true, content: "区域里 B 列是文本，画不出数值，改成 C 列就可以了。", toolCalls: [] },
    ];
    const outcome = await runSheetAgent({ sheet: sampleSheet(), instruction: "按区域画个图", ask: async () => script.shift()! });
    if (outcome.kind !== "answer") throw new Error(`期望 answer，实际 ${outcome.kind}`);
    const toolMessage = outcome.messages[3];
    expect(JSON.stringify(toolMessage)).toContain("不是数值");
  });

  it("非法入参与不存在的方法名都变成回给模型的一句话，回路继续走", async () => {
    const script: ToolTurnResult[] = [
      {
        ok: true,
        content: "",
        toolCalls: [
          { id: "a", name: "sheet_find", arguments: "{这不是 JSON" },
          { id: "b", name: "sheet_sort_rows", arguments: "{}" },
        ],
      },
      { ok: true, content: "没法排序，我只能改格子、列宽和整行增删。", toolCalls: [] },
    ];
    const outcome = await runSheetAgent({ sheet: sampleSheet(), instruction: "按销售额排序", ask: async () => script.shift()! });

    expect(outcome.kind).toBe("answer");
    const toolContents = outcome.messages.filter((message) => message.role === "tool").map((message) => message.content);
    expect(toolContents[0]).toContain("不是合法 JSON");
    expect(toolContents[1]).toContain("不认识的工具：sheet_sort_rows");
  });

  it("只调工具不给结论时转到上限就停，不许无限烧 token", async () => {
    const ask = vi.fn(async (_payload: ToolTurnRequest) => turnWithTool("sheet_describe", "{}"));
    const outcome = await runSheetAgent({ sheet: sampleSheet(), instruction: "看看这张表", ask });

    expect(outcome.kind).toBe("failed");
    expect(outcome.kind === "failed" && outcome.message).toContain("8 轮上限");
    expect(ask).toHaveBeenCalledTimes(8);
  });

  it("按暂停就在下一轮请求发出前停下，transcript 原样带回好接着跑", async () => {
    const ask = vi.fn(async (_payload: ToolTurnRequest) => turnWithTool("sheet_describe", "{}"));
    const outcome = await runSheetAgent({ sheet: sampleSheet(), instruction: "看看这张表", ask, isPaused: () => true });

    expect(ask).not.toHaveBeenCalled();
    expect(outcome.kind).toBe("paused");
    if (outcome.kind !== "paused") return;
    expect(outcome.messages.map((message) => message.role)).toEqual(["system", "user"]);
  });

  it("暂停后继续：接着用上文，不重新组装一次摘要", async () => {
    const previous = [
      { role: "system" as const, content: "你是 Excel 表格代理" },
      { role: "user" as const, content: "上一轮已经上送过的摘要" },
    ];
    const ask = vi.fn(async (_payload: ToolTurnRequest): Promise<ToolTurnResult> => ({
      ok: true,
      content: "接着上文答完了。",
      toolCalls: [],
    }));
    const outcome = await runSheetAgent({ sheet: sampleSheet(), instruction: "继续", ask, messages: previous });

    expect(outcome.kind).toBe("answer");
    expect(ask.mock.calls[0]?.[0]?.messages).toEqual(previous);
  });

  it("没配服务商 / HTTP 失败都走 failed，并把 needsSetup 原样带上去", async () => {
    const outcome = await runSheetAgent({
      sheet: sampleSheet(),
      instruction: "算一下合计",
      ask: async () => ({ ok: false, message: "还没有配置模型。", needsSetup: true }),
    });
    expect(outcome.kind).toBe("failed");
    if (outcome.kind !== "failed") return;
    expect(outcome.needsSetup).toBe(true);
    expect(outcome.messages).toHaveLength(2);
  });

  it("不注入 ask 时走 preload 白名单上的 ai:tool-turn，渲染层不直接碰 Node 能力", async () => {
    const aiToolTurn = vi.fn(async (_payload: ToolTurnRequest): Promise<ToolTurnResult> => ({
      ok: true,
      content: "看完了。",
      toolCalls: [],
    }));
    vi.stubGlobal("window", { officeCopilot: { aiToolTurn } });

    const outcome = await runSheetAgent({ sheet: sampleSheet(), instruction: "看看这张表" });

    expect(outcome.kind).toBe("answer");
    expect(aiToolTurn).toHaveBeenCalledTimes(1);
  });
});
