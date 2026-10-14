/**
 * 视图交出去的那批改动到底包含什么（`slideDeckModel.ts`）。
 *
 * 这几个断言钉的是**落地范围**：保存会把哪些段落写进文件由它决定，
 * 所以一条都不能靠"看起来对"——尤其那条"多条 run 的段落不许整段覆盖"（`AGENTS.md` §2.3）。
 */
import { describe, expect, it } from "vitest";
import type { PptxEditRequest, PresentationModel, SlideShape } from "@office-copilot/pptx-io";
import { draftKeyOf, pendingTextEdits } from "./slideDeckModel.js";

/**
 * 取一条文字改动的定位键，顺便把"视图只该交出文字改动"这条约束钉住。
 *
 * @param edit 模型层交出的一条改动。
 * @returns `partPath#shapeId#paraIndex`；不是文字改动时直接失败，不让几何改动混进这批。
 */
function textEditKey(edit: PptxEditRequest): string {
  if (edit.kind !== "text") throw new Error(`视图只该交出文字改动，却收到 ${edit.kind}`);
  return draftKeyOf(edit.target.partPath, edit.target.shapeId, edit.target.paraIndex);
}

/**
 * 造一个只填了文字与 run 条数的形状快照，其余字段与本逻辑无关。
 *
 * @param shapeId 形状号。
 * @param paragraphs 每段的 `[文字, 该段 <a:r> 条数]`。
 * @returns 一个形状快照。
 */
function shapeOf(shapeId: number, paragraphs: [string, number][]): SlideShape {
  return {
    shapeId,
    name: `形状 ${shapeId}`,
    kind: "shape",
    isPlaceholder: false,
    hasOwnGeometry: true,
    paragraphTexts: paragraphs.map(([runText]) => runText),
    paragraphRunCounts: paragraphs.map(([, runCount]) => runCount),
  };
}

/**
 * 造一份只有一张幻灯片的模型。
 *
 * @param shapes 这张幻灯片上的形状。
 * @returns 演示文稿模型（页面尺寸与几何无关，给个真实值即可）。
 */
function deckOf(shapes: SlideShape[]): PresentationModel {
  return {
    slideSizeEmu: { cx: 9144000, cy: 6858000 },
    slides: [{ partPath: "ppt/slides/slide1.xml", order: 0, slideId: 256, shapes }],
  };
}

describe("pendingTextEdits：视图只把能安全落地的段落交出去", () => {
  it("只有真的改了文字的单 run 段落才成一条", () => {
    const deck = deckOf([shapeOf(2, [["标题甲", 1]])]);
    expect(pendingTextEdits(deck, {})).toEqual([]);
    expect(pendingTextEdits(deck, { [draftKeyOf("ppt/slides/slide1.xml", 2, 0)]: "标题甲" })).toEqual([]);
    expect(pendingTextEdits(deck, { [draftKeyOf("ppt/slides/slide1.xml", 2, 0)]: "改过的标题" })).toEqual([
      {
        kind: "text",
        target: { partPath: "ppt/slides/slide1.xml", shapeId: 2, paraIndex: 0, runIndex: 0 },
        text: "改过的标题",
      },
    ]);
  });

  it("多条 run 的段落一次都不交出去（整段覆盖会拍平局部格式，§2.3）", () => {
    const deck = deckOf([shapeOf(2, [["粗体那一截和 24pt 那一截", 2]])]);
    const drafts = { [draftKeyOf("ppt/slides/slide1.xml", 2, 0)]: "我想整段替换" };
    expect(pendingTextEdits(deck, drafts)).toEqual([]);
  });

  it("段落的文字来自字段（0 条 run）时也不交出去", () => {
    const deck = deckOf([shapeOf(5, [["2", 0]])]);
    expect(pendingTextEdits(deck, { [draftKeyOf("ppt/slides/slide1.xml", 5, 0)]: "99" })).toEqual([]);
  });

  it("顺序跟着「页—形状—段」的显示顺序，同一批里后面的改动在主进程看得到前面已落地的结果", () => {
    const deck = {
      slideSizeEmu: { cx: 9144000, cy: 6858000 },
      slides: [
        {
          partPath: "ppt/slides/slide1.xml",
          order: 0,
          slideId: 256,
          shapes: [shapeOf(2, [["第一段", 1], ["第二段", 1]]), shapeOf(3, [["另一个形状", 1]])],
        },
        {
          partPath: "ppt/slides/slide2.xml",
          order: 1,
          slideId: 257,
          shapes: [shapeOf(2, [["第二页标题", 1]])],
        },
      ],
    } satisfies PresentationModel;
    const drafts = {
      [draftKeyOf("ppt/slides/slide2.xml", 2, 0)]: "第二页改过",
      [draftKeyOf("ppt/slides/slide1.xml", 3, 0)]: "另一个形状改过",
      [draftKeyOf("ppt/slides/slide1.xml", 2, 1)]: "第二段改过",
    };
    expect(pendingTextEdits(deck, drafts).map(textEditKey)).toEqual([
      "ppt/slides/slide1.xml#2#1",
      "ppt/slides/slide1.xml#3#0",
      "ppt/slides/slide2.xml#2#0",
    ]);
  });
});
