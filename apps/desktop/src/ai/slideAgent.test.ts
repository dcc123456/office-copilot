/**
 * 幻灯片代理的用例（`slideAgent.ts`，spec 阶段 5）：工具回合怎么转、提案翻成引擎坐标、非法入参与越界怎么变成回给模型的话。
 *
 * `ask` 一律打桩（`AGENTS.md` §7.1 禁止测试碰真实端点）；真实端点行为由一次性探针负责（§6.5、验收 B5）。
 * 这里钉的是红线：**模型给的页号/形状号/段/run 号对不上现场模型时，那一条不许交出去**——
 * 主进程的事务语义是"任一条失败就整批不写"（§2.4），代理交一批注定失败的改动等于让用户白点一次确认。
 */
import { describe, expect, it } from "vitest";
import type { PresentationModel } from "@office-copilot/pptx-io";
import type { ToolTurnRequest, ToolTurnResult } from "../global.js";
import { runSlideAgent, slideContextOf, slideRequestsOfProposal } from "./slideAgent.js";

/** 两页夹具：第 1 页有标题占位符（id=2，单 run）与一个自带几何的文本框（id=4，一段两条 run）与一张图片（id=7）。 */
function sampleDeck(): PresentationModel {
  return {
    slideSizeEmu: { cx: 9144000, cy: 6858000 },
    slides: [
      {
        partPath: "ppt/slides/slide1.xml",
        order: 0,
        slideId: 256,
        shapes: [
          {
            shapeId: 2,
            name: "Title 1",
            kind: "shape",
            isPlaceholder: true,
            placeholderRef: "title",
            hasOwnGeometry: false,
            paragraphTexts: ["标题甲"],
            paragraphRunCounts: [1],
            paragraphRunTexts: [["标题甲"]],
            paragraphRunFormats: [[{}]],
          },
          {
            shapeId: 4,
            name: "TextBox 3",
            kind: "shape",
            isPlaceholder: false,
            hasOwnGeometry: true,
            paragraphTexts: ["粗细两截"],
            paragraphRunCounts: [2],
            paragraphRunTexts: [["粗", "细"]],
            paragraphRunFormats: [[{ bold: true }, { sizePt: 18 }]],
            boxPx: { leftPx: 96, topPx: 384, widthPx: 288, heightPx: 96 },
          },
          {
            shapeId: 7,
            name: "Picture 6",
            kind: "picture",
            isPlaceholder: false,
            hasOwnGeometry: true,
            paragraphTexts: [],
            paragraphRunCounts: [],
            paragraphRunTexts: [],
            paragraphRunFormats: [],
            boxPx: { leftPx: 0, topPx: 0, widthPx: 100, heightPx: 100 },
          },
        ],
      },
      { partPath: "ppt/slides/slide2.xml", order: 1, slideId: 257, shapes: [] },
    ],
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

/**
 * 按脚本依次回包的 `ask` 桩：模型每调一次工具就前进一步，脚本用完即报错（用例不该依赖无限轮）。
 *
 * @param turns 逐轮回包。
 * @returns 可直接塞进 `runSlideAgent({ ask })` 的函数，带一个 `requests` 记录供断言。
 */
function stubAsk(turns: ToolTurnResult[]) {
  const requests: ToolTurnRequest[] = [];
  const ask = async (payload: ToolTurnRequest) => {
    requests.push(payload);
    const next = turns.shift();
    if (!next) throw new Error("脚本用完了：用例该给够回合数");
    return next;
  };
  return { ask, requests };
}

describe("slideRequestsOfProposal：模型视角的坐标翻成引擎坐标", () => {
  it("四类改动都落到正确的部件路径与序号上", () => {
    const { requests, notes } = slideRequestsOfProposal(sampleDeck(), [
      { type: "set_run_text", slideNumber: 1, shapeId: 2, paraIndex: 0, runIndex: 0, text: "新标题" },
      { type: "set_run_font", slideNumber: 1, shapeId: 2, paraIndex: 0, runIndex: 0, sizePt: 28 },
      { type: "set_run_color", slideNumber: 1, shapeId: 4, paraIndex: 0, runIndex: 1, color: "#C00000" },
      { type: "set_shape_geometry", slideNumber: 1, shapeId: 4, leftPx: 120, heightPx: 200 },
    ]);
    expect(notes).toEqual([]);
    expect(requests).toEqual([
      { kind: "text", target: { partPath: "ppt/slides/slide1.xml", shapeId: 2, paraIndex: 0, runIndex: 0 }, text: "新标题" },
      { kind: "font", target: { partPath: "ppt/slides/slide1.xml", shapeId: 2, paraIndex: 0, runIndex: 0 }, format: { sizePt: 28 } },
      { kind: "color", target: { partPath: "ppt/slides/slide1.xml", shapeId: 4, paraIndex: 0, runIndex: 1 }, hex: "#C00000" },
      { kind: "geometry", target: { partPath: "ppt/slides/slide1.xml", shapeId: 4 }, box: { leftPx: 120, heightPx: 200 } },
    ]);
  });

  it("多 run 的段落按模型指的 run 改，不整段覆盖（§2.3）", () => {
    const { requests } = slideRequestsOfProposal(sampleDeck(), [
      { type: "set_run_text", slideNumber: 1, shapeId: 4, paraIndex: 0, runIndex: 1, text: "只改第二截" },
    ]);
    expect(requests).toHaveLength(1);
    if (requests[0]!.kind !== "text") throw new Error("应该是文字改动");
    expect(requests[0]!.target).toMatchObject({ shapeId: 4, paraIndex: 0, runIndex: 1 });
  });

  it("页号、形状号、段号、run 号对不上现场模型的一律不交出，逐条给中文理由（§2.4 整批不落的代价）", () => {
    const { requests, notes } = slideRequestsOfProposal(sampleDeck(), [
      { type: "set_run_text", slideNumber: 9, shapeId: 2, paraIndex: 0, runIndex: 0, text: "越界的页" },
      { type: "set_run_text", slideNumber: 1, shapeId: 99, paraIndex: 0, runIndex: 0, text: "没有的形状" },
      { type: "set_run_text", slideNumber: 1, shapeId: 2, paraIndex: 5, runIndex: 0, text: "越界的段" },
      { type: "set_run_text", slideNumber: 1, shapeId: 2, paraIndex: 0, runIndex: 3, text: "越界的 run" },
      { type: "set_run_text", slideNumber: 1, shapeId: 7, paraIndex: 0, runIndex: 0, text: "图片没有文字" },
      { type: "set_shape_geometry", slideNumber: 2, shapeId: 4, leftPx: 10 },
    ]);
    expect(requests).toEqual([]);
    expect(notes).toHaveLength(6);
    expect(notes.join("\n")).toContain("没有第 9 页");
    expect(notes.join("\n")).toContain("没有 shapeId=99");
    expect(notes.join("\n")).toContain("越界（只有 1 段）");
    expect(notes.join("\n")).toContain("没有第 3 个");
    expect(notes.join("\n")).toContain("没有文字体");
    expect(notes.join("\n")).toContain("第 2 页没有 shapeId=4");
  });

  it("文字来自字段（页码那种 0 条 run 的段落）也拒，别让引擎去回『run 序号越界』", () => {
    const deck = sampleDeck();
    deck.slides[0]!.shapes[0] = { ...deck.slides[0]!.shapes[0]!, paragraphRunCounts: [0], paragraphRunTexts: [[]], paragraphRunFormats: [[]] };
    const { requests, notes } = slideRequestsOfProposal(deck, [
      { type: "set_run_font", slideNumber: 1, shapeId: 2, paraIndex: 0, runIndex: 0, bold: true },
    ]);
    expect(requests).toEqual([]);
    expect(notes[0]).toContain("文字来自字段");
  });
});

describe("runSlideAgent：工具回合", () => {
  it("上送的首条消息带结构画像，并说清单位与截断规则", async () => {
    const { ask, requests } = stubAsk([turnWithAnswer("这份文稿有 2 页。")]);
    const outcome = await runSlideAgent({ deck: sampleDeck(), instruction: "有几页？", ask });
    expect(outcome.kind).toBe("answer");
    const firstUser = requests[0]!.messages[1]!.content ?? "";
    expect(firstUser).toContain("2 页，页面 960×720 px");
    expect(firstUser).toContain("诉求：有几页？");
    expect((outcome as { text: string }).text).toBe("这份文稿有 2 页。");
  });

  it("模型调 ppt_describe 与 ppt_read_slide_text 后交提案：回合收尾，交出的是引擎坐标", async () => {
    const proposal = { summary: "把标题改成新标题", edits: [{ type: "set_run_text", slideNumber: 1, shapeId: 2, paraIndex: 0, runIndex: 0, text: "新标题" }] };
    const { ask } = stubAsk([
      turnWithTool("ppt_describe", "{}"),
      turnWithTool("ppt_read_slide_text", JSON.stringify({ slideNumber: 1 })),
      turnWithTool("ppt_submit_proposal", JSON.stringify(proposal), "我改标题"),
    ]);
    const outcome = await runSlideAgent({ deck: sampleDeck(), instruction: "标题改成新标题", ask });
    if (outcome.kind !== "proposal") throw new Error(`应该出提案，实际是 ${outcome.kind}`);
    expect(outcome.summary).toBe("把标题改成新标题");
    expect(outcome.edits).toEqual([
      { kind: "text", target: { partPath: "ppt/slides/slide1.xml", shapeId: 2, paraIndex: 0, runIndex: 0 }, text: "新标题" },
    ]);
    expect(outcome.notes).toEqual([]);
    // 回给模型的工具消息：describe 带页号，read_slide_text 带逐 run 的文字与格式。
    const toolMessages = outcome.messages.filter((message) => message.role === "tool");
    expect(toolMessages).toHaveLength(3);
    expect(toolMessages[0]!.content).toContain('"slideNumber":1');
    expect(toolMessages[1]!.content).toContain("标题甲");
    expect(toolMessages[1]!.content).toContain('"runIndex":1'); // 文本框那段两条 run 都列出来
    expect(toolMessages[2]!.content).toContain("等用户确认后才会落地");
  });

  it("坏 JSON 与非法入参回中文理由，回路继续转而不是抛异常（§7.1、验收 B9）", async () => {
    const { ask } = stubAsk([
      turnWithTool("ppt_submit_proposal", "{这不是 JSON"),
      turnWithTool("ppt_submit_proposal", JSON.stringify({ summary: "越界", edits: [{ type: "set_run_font", slideNumber: 1, shapeId: 2, paraIndex: 0, runIndex: 0 }] })),
      turnWithTool("ppt_submit_proposal", JSON.stringify({ summary: "没有这种改动", edits: [{ type: "add_slide", slideNumber: 1 }] })),
      turnWithTool("ppt_no_such_tool", "{}"),
      turnWithAnswer("这些我都做不了。"),
    ]);
    const outcome = await runSlideAgent({ deck: sampleDeck(), instruction: "随便改改", ask });
    expect(outcome.kind).toBe("answer");
    const toolContents = outcome.messages.filter((message) => message.role === "tool").map((message) => message.content);
    expect(toolContents[0]).toContain("不是合法 JSON");
    expect(toolContents[1]).toContain("bold 与 sizePt 至少给一个");
    expect(toolContents[2]).toContain("edits.0");
    expect(toolContents[3]).toContain("不认识的工具：ppt_no_such_tool");
  });

  it("提案里越界的那几条只进 notes，绝不混进要落地的批次", async () => {
    const { ask } = stubAsk([
      turnWithTool("ppt_submit_proposal", JSON.stringify({ summary: "两页一起改", edits: [
        { type: "set_run_text", slideNumber: 1, shapeId: 2, paraIndex: 0, runIndex: 0, text: "新标题" },
        { type: "set_run_text", slideNumber: 3, shapeId: 2, paraIndex: 0, runIndex: 0, text: "不存在的页" },
      ] })),
    ]);
    const outcome = await runSlideAgent({ deck: sampleDeck(), instruction: "改标题", ask });
    if (outcome.kind !== "proposal") throw new Error(`应该出提案，实际是 ${outcome.kind}`);
    expect(outcome.edits).toHaveLength(1);
    expect(outcome.notes).toEqual(["没有第 3 页（这份文稿共 2 页），这条跳过。"]);
    expect(outcome.proposed).toHaveLength(2); // 模型原样的提案留着给用户看
  });

  it("暂停在下一轮请求前停下，transcript 原样带回", async () => {
    const { ask } = stubAsk([turnWithAnswer("不该走到这里")]);
    const outcome = await runSlideAgent({ deck: sampleDeck(), instruction: "改标题", ask, isPaused: () => true });
    expect(outcome.kind).toBe("paused");
    expect(outcome.messages).toHaveLength(2);
  });

  it("模型一直要点工具不出结论：到 8 轮上限停下并报原因", async () => {
    const ask = async () => turnWithTool("ppt_describe", "{}");
    const outcome = await runSlideAgent({ deck: sampleDeck(), instruction: "改标题", ask });
    expect(outcome.kind).toBe("failed");
    expect((outcome as { message: string }).message).toContain("8 轮上限");
  });

  it("HTTP 层失败（没配服务商）走 failed，并把 needsSetup 透传给 UI", async () => {
    const { ask } = stubAsk([{ ok: false, message: "没配置服务商", needsSetup: true, toolCalls: [], content: "" } as unknown as ToolTurnResult]);
    const outcome = await runSlideAgent({ deck: sampleDeck(), instruction: "改标题", ask });
    expect(outcome.kind).toBe("failed");
    expect(outcome).toMatchObject({ needsSetup: true });
  });

  it("画像超上限时上送文本里带一句截断说明，而不是把整份文稿摊上去（§2.9）", () => {
    const deck = sampleDeck();
    deck.slides = Array.from({ length: 25 }, (_, index) => ({ ...deck.slides[1]!, order: index, partPath: `ppt/slides/slide${index + 1}.xml` }));
    const context = slideContextOf(deck);
    expect(context).toContain("25 页");
    expect(context).toContain("只描述了前 20 页");
  });
});
