/**
 * 视图交出去的那批改动到底包含什么（`slideDeckModel.ts`）。
 *
 * 这几个断言钉的是**落地范围**：保存会把哪些段落、哪些字段写进文件由它决定，
 * 所以一条都不能靠"看起来对"——尤其那条"多条 run 的段落不许整段覆盖"（`AGENTS.md` §2.3）。
 */
import { describe, expect, it } from "vitest";
import type { PptxEditRequest, PresentationModel, RunFormat, SlideShape } from "@office-copilot/pptx-io";
import { boxDraftKeyOf, draftKeyOf, pendingEdits } from "./slideDeckModel.js";

/** 一张幻灯片上的部件路径，全部用例共用。 */
const PART = "ppt/slides/slide1.xml";

/**
 * 取一条文字改动的定位键，顺便把"这批里的文字改动都指到 run 0"钉住。
 *
 * @param edit 模型层交出的一条改动。
 * @returns `partPath#shapeId#paraIndex`；不是文字改动时直接失败。
 */
function textEditKey(edit: PptxEditRequest): string {
  if (edit.kind !== "text") throw new Error(`这条不是文字改动：${edit.kind}`);
  if (edit.target.runIndex !== 0) throw new Error(`本视图只改单 run 段落，runIndex 该是 0`);
  return draftKeyOf(edit.target.partPath, edit.target.shapeId, edit.target.paraIndex);
}

/**
 * 造一个形状快照：文字、每段 run 条数与第一段那截 run 的现场格式。
 *
 * @param shapeId 形状号。
 * @param paragraphs 每段的 `[文字, 该段 <a:r> 条数, run 0 的现场格式]`；第三项只给段 0 用。
 * @returns 一个形状快照；没有自己的 `<a:xfrm>`（`boxPx` 留空），几何用例各自再补。
 */
function shapeOf(shapeId: number, paragraphs: [string, number, RunFormat?][]): SlideShape {
  return {
    shapeId,
    name: `形状 ${shapeId}`,
    kind: "shape",
    isPlaceholder: false,
    hasOwnGeometry: false,
    paragraphTexts: paragraphs.map(([paragraphText]) => paragraphText),
    paragraphRunCounts: paragraphs.map(([, runCount]) => runCount),
    paragraphRunFormats: paragraphs.map(([, , runFormat]) => [runFormat ?? {}]),
  };
}

/**
 * 造一份只有一张幻灯片的模型。
 *
 * @param shapes 这张幻灯片上的形状。
 * @returns 演示文稿模型（页面尺寸与几何无关，给个真实值即可）。
 */
function deckOf(shapes: SlideShape[]): PresentationModel {
  return { slideSizeEmu: { cx: 9144000, cy: 6858000 }, slides: [{ partPath: PART, order: 0, slideId: 256, shapes }] };
}

describe("pendingEdits：视图只把能安全落地的字段交出去", () => {
  it("只有真的改了文字的单 run 段落才成一条", () => {
    const deck = deckOf([shapeOf(2, [["标题甲", 1]])]);
    expect(pendingEdits(deck, {})).toEqual([]);
    expect(pendingEdits(deck, { [draftKeyOf(PART, 2, 0)]: { text: "标题甲" } })).toEqual([]);
    expect(pendingEdits(deck, { [draftKeyOf(PART, 2, 0)]: { text: "改过的标题" } })).toEqual([
      { kind: "text", target: { partPath: PART, shapeId: 2, paraIndex: 0, runIndex: 0 }, text: "改过的标题" },
    ]);
  });

  it("多条 run 的段落一次都不交出去（整段覆盖会拍平局部格式，§2.3）", () => {
    const deck = deckOf([shapeOf(2, [["粗体那一截和 24pt 那一截", 2]])]);
    expect(pendingEdits(deck, { [draftKeyOf(PART, 2, 0)]: { text: "我想整段替换", sizePt: 32 } })).toEqual([]);
  });

  it("段落的文字来自字段（0 条 run）时也不交出去", () => {
    const deck = deckOf([shapeOf(5, [["2", 0]])]);
    expect(pendingEdits(deck, { [draftKeyOf(PART, 5, 0)]: { text: "99" } })).toEqual([]);
  });

  it("字号与颜色各成一条，且只带真改过的那一项", () => {
    const deck = deckOf([shapeOf(2, [["正文", 1, { sizePt: 18, bold: true, hex: "#ff0000" }]])]);
    // 只改字号：不动 bold，主进程那条 op 也只会写 sz 一个属性（§2.3 不拍平）。
    expect(pendingEdits(deck, { [draftKeyOf(PART, 2, 0)]: { sizePt: 24 } })).toEqual([
      { kind: "font", target: { partPath: PART, shapeId: 2, paraIndex: 0, runIndex: 0 }, format: { sizePt: 24 } },
    ]);
    // 现场值原样敲回去：什么都不交出，空点保存不写盘。
    expect(pendingEdits(deck, { [draftKeyOf(PART, 2, 0)]: { sizePt: 18, hex: "#ff0000" } })).toEqual([]);
    expect(pendingEdits(deck, { [draftKeyOf(PART, 2, 0)]: { hex: "#00ff00", text: "正文" } })).toEqual([
      { kind: "color", target: { partPath: PART, shapeId: 2, paraIndex: 0, runIndex: 0 }, hex: "#00ff00" },
    ]);
  });

  it("继承来的字号（现场没有 sz）也一样：给了新值才成一条", () => {
    const deck = deckOf([shapeOf(2, [["标题", 1, {}]])]);
    expect(pendingEdits(deck, { [draftKeyOf(PART, 2, 0)]: { sizePt: undefined } })).toEqual([]);
    expect(pendingEdits(deck, { [draftKeyOf(PART, 2, 0)]: { sizePt: 28 } })).toEqual([
      { kind: "font", target: { partPath: PART, shapeId: 2, paraIndex: 0, runIndex: 0 }, format: { sizePt: 28 } },
    ]);
  });

  it("形状自己写了 <a:xfrm> 才交几何改动，只带改过的那几个字段", () => {
    const withBox = { ...shapeOf(4, [["文本", 1]]), boxPx: { leftPx: 96, topPx: 384, widthPx: 288, heightPx: 96 } };
    const deck = deckOf([withBox]);
    expect(pendingEdits(deck, { [boxDraftKeyOf(PART, 4)]: { box: { leftPx: 96, widthPx: 288 } } })).toEqual([]);
    expect(pendingEdits(deck, { [boxDraftKeyOf(PART, 4)]: { box: { topPx: 400, widthPx: 300 } } })).toEqual([
      { kind: "geometry", target: { partPath: PART, shapeId: 4 }, box: { topPx: 400, widthPx: 300 } },
    ]);
  });

  it("几何继承自版式（没有自己的 <a:xfrm>）时一个字段都不写（spec D3 不许烘焙）", () => {
    const deck = deckOf([shapeOf(2, [["标题", 1]])]);
    expect(pendingEdits(deck, { [boxDraftKeyOf(PART, 2)]: { box: { leftPx: 10, widthPx: 20 } } })).toEqual([]);
  });

  it("顺序跟着「页—形状（几何先，再各段）」的显示顺序，同一批里后面的改动在主进程看得到前面已落地的结果", () => {
    const secondPageShape = shapeOf(2, [["第二页标题", 1]]);
    const deck = {
      slideSizeEmu: { cx: 9144000, cy: 6858000 },
      slides: [
        {
          partPath: PART,
          order: 0,
          slideId: 256,
          shapes: [
            shapeOf(2, [["第一段", 1], ["第二段", 1]]),
            { ...shapeOf(3, [["另一个形状", 1]]), boxPx: { leftPx: 0, topPx: 0, widthPx: 100, heightPx: 50 } },
          ],
        },
        { partPath: "ppt/slides/slide2.xml", order: 1, slideId: 257, shapes: [secondPageShape] },
      ],
    } satisfies PresentationModel;
    const drafts = {
      [draftKeyOf("ppt/slides/slide2.xml", 2, 0)]: { text: "第二页改过" },
      [boxDraftKeyOf(PART, 3)]: { box: { widthPx: 120 } },
      [draftKeyOf(PART, 2, 1)]: { text: "第二段改过" },
      [draftKeyOf(PART, 3, 0)]: { hex: "#123456" },
    };
    expect(pendingEdits(deck, drafts).map((edit) => (edit.kind === "text" ? textEditKey(edit) : `${edit.kind}:${edit.target.shapeId}`))).toEqual([
      `${PART}#2#1`,
      "geometry:3",
      "color:3",
      "ppt/slides/slide2.xml#2#0",
    ]);
  });
});
