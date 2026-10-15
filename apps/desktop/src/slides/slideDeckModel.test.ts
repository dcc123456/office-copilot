/**
 * 视图交出去的那批改动到底包含什么（`slideDeckModel.ts`）。
 *
 * 这几个断言钉的是**落地范围**：保存会把哪些段落、哪些字段写进文件由它决定，
 * 所以一条都不能靠"看起来对"——尤其那条"多条 run 的段落不许整段覆盖"（`AGENTS.md` §2.3）。
 */
import { describe, expect, it } from "vitest";
import type { PptxEditRequest, PresentationModel, RunFormat, Slide, SlideNotes, SlideShape } from "@office-copilot/pptx-io";
import { boxDraftKeyOf, draftKeyOf, dropShapeDrafts, inheritedBoxNote, layerEdgeNote, notesRequestOf, notesSaveBlock, pendingEdits, sizeFieldHint } from "./slideDeckModel.js";

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
 * 把交出的一条改动折成断言里好读的短串。
 *
 * @param edit 模型层交出的一条改动。
 * @returns 文字改动是 `partPath#shapeId#paraIndex`，其余是 `kind:形状号`；
 *   整页粒度的那两条（加文本框）在这里出现说明用例写错了——草稿里没有这种键。
 */
function describeEdit(edit: PptxEditRequest): string {
  if (edit.kind === "text") return textEditKey(edit);
  return "shapeId" in edit.target ? `${edit.kind}:${edit.target.shapeId}` : `${edit.kind}:整页`;
}

/**
 * 造一个形状快照：文字、每段 run 条数与第一段那截 run 的现场格式。
 *
 * 每段的 run 级现场（文字与格式）这里只给一条，与 `paragraphRunCounts` 不一致也无妨：
 * 落地范围由条数说了算（多条 run 就不许整段覆盖），本文件的断言不读 run 级文字。
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
    paragraphRunTexts: paragraphs.map(([paragraphText]) => [paragraphText]),
    paragraphRunFormats: paragraphs.map(([, , runFormat]) => [runFormat ?? {}]),
    effectiveBoxPx: {},
    boxPxSources: {},
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
    expect(pendingEdits(deck, drafts).map(describeEdit)).toEqual([
      `${PART}#2#1`,
      "geometry:3",
      "color:3",
      "ppt/slides/slide2.xml#2#0",
    ]);
  });

  it("草稿里凭空写一个整页粒度的键也不会交出那种改动（加/删只由按钮直接送主进程）", () => {
    const deck = deckOf([shapeOf(2, [["第一段", 1]])]);
    const drafts = { "ppt/slides/slide1.xml#2": { text: "不像 run 键" }, [`${PART}@`]: { box: { widthPx: 10 } } };
    expect(pendingEdits(deck, drafts).map(describeEdit)).toEqual([]);
  });
});

describe("layerEdgeNote：到头的那个方向压根不许交出去", () => {
  /** 三个形状按文档顺序（= 叠放顺序：第一个最底、最后一个最顶）排在同一页上。 */
  const bottom = shapeOf(2, [["最底", 1]]);
  const middle = shapeOf(4, [["中间", 1]]);
  const top = shapeOf(7, [["最顶", 1]]);
  const shapes = [bottom, middle, top];

  it("中间的形状四个方向都能挪", () => {
    for (const direction of ["up", "down", "top", "bottom"] as const) {
      expect(layerEdgeNote(shapes, middle, direction), direction).toBeNull();
    }
  });

  it("最顶的那个只给了向下那两头（交出去会被引擎整批拒，§2.4）", () => {
    expect(layerEdgeNote(shapes, top, "up"), "up").toBe("它已经在最顶层");
    expect(layerEdgeNote(shapes, top, "top"), "top").toBe("它已经在最顶层");
    expect(layerEdgeNote(shapes, top, "down"), "down").toBeNull();
  });

  it("最底的那个反过来：向上的两头到头", () => {
    expect(layerEdgeNote(shapes, bottom, "up"), "up").toBeNull();
    expect(layerEdgeNote(shapes, bottom, "bottom"), "bottom").toBe("它已经在最底层");
  });

  it("这一页只有一个形状时四个按钮都不给点（挪了也是白挪）", () => {
    const lonely = shapeOf(2, [["只有一个", 1]]);
    for (const direction of ["up", "down", "top", "bottom"] as const) {
      expect(layerEdgeNote([lonely], lonely, direction), direction).toBe("这一页只有这一个形状，没有可叠放的邻居");
    }
  });
});

describe("组合那一层：叠放只比组内的邻居，几何那一格只给一句话", () => {
  /** 一页五格：标题、组、组里两格、一条连接线（读模型给的平铺顺序就是文档顺序）。 */
  const title = shapeOf(2, [["标题", 1]]);
  const group = { ...shapeOf(9, [["", 0]]), kind: "group" as const };
  const first = { ...shapeOf(3, [["组内甲", 1]]), groupShapeId: 9 };
  const second = { ...shapeOf(4, [["组内乙", 1]]), groupShapeId: 9 };
  const connector = { ...shapeOf(10, [["", 0]]), kind: "connector" as const };
  const shapes = [title, group, first, second, connector];

  it("组里最后一格向上到头：拿整页平铺清单比会误判成还能挪，交出去主进程整批不落地（§2.4）", () => {
    expect(layerEdgeNote(shapes, second, "up"), "up").toBe("它已经在组内最顶层");
    expect(layerEdgeNote(shapes, second, "top"), "top").toBe("它已经在组内最顶层");
    expect(layerEdgeNote(shapes, second, "down"), "down").toBeNull();
    expect(layerEdgeNote(shapes, first, "down"), "down").toBe("它已经在组内最底层");
    expect(layerEdgeNote(shapes, first, "up"), "up").toBeNull();
  });

  it("页面那一层的三格照旧按整页比：连接线在最顶、标题在最底", () => {
    expect(layerEdgeNote(shapes, connector, "up"), "up").toBe("它已经在最顶层");
    expect(layerEdgeNote(shapes, title, "bottom"), "bottom").toBe("它已经在最底层");
    expect(layerEdgeNote(shapes, group, "up"), "up").toBeNull();
  });

  it("组里只剩一格时四个方向都不给点，理由说的是组内", () => {
    const onlyOne = { ...shapeOf(3, [["组内只有一个", 1]]), groupShapeId: 9 };
    for (const direction of ["up", "down", "top", "bottom"] as const) {
      expect(layerEdgeNote([group, onlyOne], onlyOne, direction), direction).toBe("这个组合里只有这一个形状，组内没有可叠放的邻居");
    }
  });

  it("组内那格的几何说明不许念成「自己没写几何」：它写了，只是写在组内子坐标系里", () => {
    expect(inheritedBoxNote({ ...first, hasOwnGeometry: true })).toContain("组内子坐标系");
    // 组自己那一格同理：它把几何写在 `<p:grpSpPr>` 的组框里，读侧没接，不许念成"没写"。
    expect(inheritedBoxNote(group)).toContain("组框");
    // 连接线在自己身上写了几何时不受影响（它不在组里），走原来那条继承说明。
    expect(inheritedBoxNote(connector)).toContain("这个形状自己没写几何");
  });
});

describe("dropShapeDrafts：删掉的形状不留未保存的草稿", () => {
  it("只去掉那一个形状的各段与它的几何，别的形状一条不少", () => {
    const drafts = {
      [draftKeyOf(PART, 2, 0)]: { text: "留着" },
      [draftKeyOf(PART, 4, 0)]: { text: "要删的" },
      [draftKeyOf(PART, 4, 1)]: { hex: "#ff0000" },
      [boxDraftKeyOf(PART, 4)]: { box: { widthPx: 200 } },
    };
    expect(Object.keys(dropShapeDrafts(drafts, PART, 4))).toEqual([`${PART}#2#0`]);
  });

  it("形状号 4 的清理不会顺手带走形状号 41 的草稿（键里那个分隔符是算上的）", () => {
    const drafts = { [draftKeyOf(PART, 4, 0)]: { text: "四号" }, [draftKeyOf(PART, 41, 0)]: { text: "四一号" } };
    expect(Object.keys(dropShapeDrafts(drafts, PART, 4))).toEqual([`${PART}#41#0`]);
  });

  it("这一页这个形状压根没草稿时原对象返回，不白造一份新状态", () => {
    const drafts = { [draftKeyOf(PART, 2, 0)]: { text: "留着" } };
    expect(dropShapeDrafts(drafts, "ppt/slides/slide2.xml", 9)).toBe(drafts);
  });
});

/**
 * 没有自己几何的那一格显示什么（`inheritedBoxNote`）。
 *
 * 数字来自读模型算完继承链的 `effectiveBoxPx`（spec D3 要求"有效值 + 来源层"两个都给），
 * 但这一句必须**只读**：草稿照样不交出几何字段，否则就是把继承值烤进 slide。
 */
describe("inheritedBoxNote：继承来的几何如实报数，但一个字都不写", () => {
  /**
   * 造一个"自己没写几何"的占位符快照。
   *
   * @param sources 每一轴的来源层（读模型给的那份）。
   * @returns 一个形状快照；`boxPx` 留空，与真夹具里那个空 `<p:spPr/>` 同形。
   */
  function inheritedShape(sources: SlideShape["boxPxSources"], box: SlideShape["effectiveBoxPx"] = { leftPx: 96, topPx: 48, widthPx: 768, heightPx: 120 }): SlideShape {
    return { ...shapeOf(2, [["标题", 1]]), isPlaceholder: true, placeholderRef: "ctrTitle", effectiveBoxPx: box, boxPxSources: sources };
  }

  it("四轴都来自母版时句子念出四个数与「母版」", () => {
    const note = inheritedBoxNote(inheritedShape({ leftPx: "master", topPx: "master", widthPx: "master", heightPx: "master" }));
    expect(note).toContain("继承自母版");
    expect(note).toContain("左 96 · 上 48 · 宽 768 · 高 120 px");
  });

  it("左上来自版式、宽高来自母版时两个层都点名（逐轴回退是 python-pptx 那套，不许只说一个层）", () => {
    const note = inheritedBoxNote(inheritedShape({ leftPx: "layout", topPx: "layout", widthPx: "master", heightPx: "master" }));
    expect(note).toContain("继承自版式、母版");
  });

  it("三层都没写几何时不给数字，只说没写（视图不许造值）", () => {
    const note = inheritedBoxNote(inheritedShape({}, {}));
    expect(note).toContain("版式与母版都没给这一格写几何");
    expect(note).not.toContain("px"); // 一个数都不许出现：带数就得带单位
  });

  it("非占位符又没写自己的几何时说的是「自己没写」，不拿版式那格的数来凑", () => {
    const note = inheritedBoxNote({ ...shapeOf(4, [["正文", 1]]), effectiveBoxPx: {}, boxPxSources: {} });
    expect(note).toContain("这个形状自己没写几何");
  });

  it("报告了继承数之后，草稿里填几轴仍然一个字段都不交出（只读这条是红线）", () => {
    const deck = deckOf([inheritedShape({ leftPx: "master" }, { leftPx: 96 })]);
    expect(pendingEdits(deck, { [boxDraftKeyOf(PART, 2)]: { box: { leftPx: 500 } } })).toEqual([]);
  });
});

describe("sizeFieldHint：字号框说的是「实际画多大」，写侧仍只看显式 sz（spec D3）", () => {
  it("自己写了 sz 时只说自己写的，一句不提继承", () => {
    expect(sizeFieldHint({ sizePt: 18, effectiveSizePt: 18, effectiveSizeSource: "run" })).toBe("当前 18pt（这条 run 自己写的）");
  });

  it("没写 sz 时给出链上算出来的数与来源层，单位写着 pt", () => {
    expect(sizeFieldHint({ effectiveSizePt: 40, effectiveSizeSource: "layout" })).toBe("当前 40pt，继承自版式那一格（本形状没写字号，写死就切断母版联动）");
  });

  it("整条链都没写时一个数都不给，不许拿 18pt 这种常见默认凑", () => {
    expect(sizeFieldHint({})).toBe("这一格没写字号，继承链上也没找到（视图不写死，spec D3）");
  });

  it("继承值只进说明、不进输入框：草稿没敲字号时一条 op 都不交出", () => {
    const deck = deckOf([shapeOf(2, [["标题", 1, { effectiveSizePt: 40, effectiveSizeSource: "layout" }]])]);
    expect(pendingEdits(deck, { [draftKeyOf(PART, 2, 0)]: { sizePt: undefined } })).toEqual([]);
  });
});

/** 一页只有备注这一件事的夹具：`notes` 传 null 表示这一页没有备注部件。 */
function slideWithNotes(notes: SlideNotes | null): Slide {
  return { partPath: PART, order: 0, slideId: 256, shapes: [], ...(notes ? { notes } : {}) };
}

describe("notesSaveBlock / notesRequestOf：备注那一格的落地范围", () => {
  const plain: SlideNotes = {
    partPath: "ppt/notesSlides/notesSlide1.xml",
    paragraphTexts: ["旧备注一句", "旧备注二句"],
    paragraphRunCounts: [1, 1],
    styledRunCount: 0,
  };

  it("这一页没有备注部件时不给存，并说清新建那一档还没做", () => {
    expect(notesSaveBlock(slideWithNotes(null), "随便写点")).toContain("没有备注页");
  });

  it("输入框与现场逐字一样时不给存（磁盘上就是一行一段，join 与 split 是同一份对应关系）", () => {
    const slide = slideWithNotes(plain);
    expect(notesSaveBlock(slide, plain.paragraphTexts.join("\n"))).toBe("备注没有改动。");
  });

  it("现有备注带多截格式时不给存，理由是会被拍平（§2.3，判定与引擎共用那一份）", () => {
    const styled = slideWithNotes({ ...plain, paragraphRunCounts: [1, 2], styledRunCount: 1 });
    expect(notesSaveBlock(styled, "换成一句")).toContain("压平");
  });

  it("同一种格式的备注改过就给存：坐标是现场那份备注部件路径，一行一段", () => {
    const slide = slideWithNotes(plain);
    expect(notesSaveBlock(slide, "新备注一句\n\n新备注三句")).toBeNull();
    expect(notesRequestOf(plain, "新备注一句\n\n新备注三句")).toEqual({
      kind: "set_notes",
      target: { notesPartPath: "ppt/notesSlides/notesSlide1.xml" },
      paragraphTexts: ["新备注一句", "", "新备注三句"],
    });
  });
});
