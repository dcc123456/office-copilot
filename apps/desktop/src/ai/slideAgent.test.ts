/**
 * 幻灯片代理的用例（`slideAgent.ts`，spec 阶段 5）：工具回合怎么转、提案翻成引擎坐标、非法入参与越界怎么变成回给模型的话。
 *
 * `ask` 一律打桩（`AGENTS.md` §7.1 禁止测试碰真实端点）；真实端点行为由一次性探针负责（§6.5、验收 B5）。
 * 这里钉的是红线：**模型给的页号/形状号/段/run 号对不上现场模型时，那一条不许交出去**——
 * 主进程的事务语义是"任一条失败就整批不写"（§2.4），代理交一批注定失败的改动等于让用户白点一次确认。
 */
import { describe, expect, it } from "vitest";
import type { PresentationModel, SlideNotes } from "@office-copilot/pptx-io";
import type { ToolTurnRequest, ToolTurnResult } from "../global.js";
import { runSlideAgent, slideContextOf, slideRequestsOfProposal } from "./slideAgent.js";

/** 第 1 页那份备注：一段一截、没字符格式，正是"整篇换"允许的那一种。 */
const SAMPLE_NOTES: SlideNotes = {
  partPath: "ppt/notesSlides/notesSlide1.xml",
  paragraphTexts: ["旧备注一句"],
  paragraphRunCounts: [1],
  styledRunCount: 0,
};

/**
 * 两页夹具：第 1 页有标题占位符（id=2，单 run）与一个自带几何的文本框（id=4，一段两条 run）与一张图片（id=7），
 * 第 2 页是空页且**没有备注部件**（两档备注各管一半：`set_notes` 只改已有的，`create_notes` 只给没有的那一页新建）。
 *
 * @param notesOfFirstSlide 第 1 页的备注快照；默认给一份一段一截的，传 null 表示这页没有备注页。
 * @param hasNotesMaster 这份文件里有没有备注母版；默认 true，传 false 用来测"新建这一档整机提不了"那条预拒。
 */
function sampleDeck(notesOfFirstSlide: SlideNotes | null = SAMPLE_NOTES, hasNotesMaster = true): PresentationModel {
  return {
    slideSizeEmu: { cx: 9144000, cy: 6858000 },
    hasNotesMaster,
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
            // 自己没写几何（`boxPx` 留空），数从版式来：代理那条"要给就给全四轴"的守卫看的是 `boxPx`，不是这份只读的有效值。
            effectiveBoxPx: { leftPx: 88, topPx: 60, widthPx: 576, heightPx: 72 },
            boxPxSources: { leftPx: "layout", topPx: "layout", widthPx: "layout", heightPx: "layout" },
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
            effectiveBoxPx: { leftPx: 96, topPx: 384, widthPx: 288, heightPx: 96 },
            boxPxSources: { leftPx: "shape", topPx: "shape", widthPx: "shape", heightPx: "shape" },
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
            effectiveBoxPx: { leftPx: 0, topPx: 0, widthPx: 100, heightPx: 100 },
            boxPxSources: { leftPx: "shape", topPx: "shape", widthPx: "shape", heightPx: "shape" },
          },
        ],
        ...(notesOfFirstSlide ? { notes: notesOfFirstSlide } : {}),
      },
      { partPath: "ppt/slides/slide2.xml", order: 1, slideId: 257, shapes: [] },
    ],
  };
}

/**
 * 在第 1 页末尾再加一个组（id=9）与组里两格（id=5、id=6，各一段单 run 文字）。
 *
 * `sampleDeck()` 每次现造一个新对象，所以这里直接往它返回的数组上追加，不影响别的用例。
 * 页面平铺顺序成了 2、4、7、9、5、6：组自己那一层在最顶，组里 6 在最顶——正是"必须按层比"的那两种情形。
 */
function deckWithGroup(): PresentationModel {
  const deck = sampleDeck();
  const shapes = deck.slides[0]!.shapes;
  const picture = shapes[2]!;
  /** 组里的一格：一段一截文字，自己写了几何但那是组内子坐标（读模型这时候不给 `boxPx`）。 */
  const groupedChild = (shapeId: number, paragraphText: string) => ({
    ...picture,
    shapeId,
    name: `TextBox ${shapeId}`,
    paragraphTexts: [paragraphText],
    paragraphRunCounts: [1],
    paragraphRunTexts: [[paragraphText]],
    paragraphRunFormats: [[{}]],
    hasOwnGeometry: true,
    boxPx: undefined,
    effectiveBoxPx: {},
    boxPxSources: {},
    groupShapeId: 9,
  });
  shapes.push({ ...picture, shapeId: 9, name: "Group 8", kind: "group" });
  shapes.push(groupedChild(5, "组内甲"));
  shapes.push(groupedChild(6, "组内乙"));
  return deck;
}

/**
 * 在第 1 页末尾再加一个图形框（id=10）装一张表：第 0 行两格各一段一截，
 * 第 1 行第 0 列是合并起点（跨两列、两段——第二段两条 run），第 1 行第 1 列是它并进来的续格。
 *
 * 表格的形状自己**没有** `paragraphTexts`（文字体在 `<a:tbl>` 的每格里），所以这一格专测装配层那条"逐格"的分岔。
 */
function deckWithTable(): PresentationModel {
  const deck = sampleDeck();
  const picture = deck.slides[0]!.shapes[2]!;
  deck.slides[0]!.shapes.push({
    ...picture,
    shapeId: 10,
    name: "Table 9",
    kind: "graphicFrame",
    paragraphTexts: [],
    paragraphRunCounts: [],
    paragraphRunTexts: [],
    paragraphRunFormats: [],
    table: {
      rows: [
        [
          { paragraphTexts: ["姓名"], paragraphRunCounts: [1], paragraphRunFormats: [[{}]], gridSpan: 1, isMergeContinuation: false },
          { paragraphTexts: ["城市"], paragraphRunCounts: [1], paragraphRunFormats: [[{}]], gridSpan: 1, isMergeContinuation: false },
        ],
        [
          { paragraphTexts: ["张三", "上海两截"], paragraphRunCounts: [1, 2], paragraphRunFormats: [[{}], [{}, { sizePt: 20 }]], gridSpan: 2, isMergeContinuation: false },
          { paragraphTexts: [""], paragraphRunCounts: [0], paragraphRunFormats: [[]], gridSpan: 1, isMergeContinuation: true },
        ],
      ],
    },
  });
  return deck;
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
  it("文字/字体/颜色/几何四条都落到正确的部件路径与序号上", () => {
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

  it("叠放次序翻成 layer 请求；已在要去那一端的别交出（引擎会整批拒，§2.4）", () => {
    const { requests, notes } = slideRequestsOfProposal(sampleDeck(), [
      { type: "move_shape_layer", slideNumber: 1, shapeId: 4, direction: "up" },
      { type: "move_shape_layer", slideNumber: 1, shapeId: 2, direction: "bottom" },
      { type: "move_shape_layer", slideNumber: 1, shapeId: 7, direction: "top" },
    ]);
    // 第 1 页的形状顺序（= 叠放顺序）是 2 在最底、7 在最顶，所以只有中间那条能挪
    expect(requests).toEqual([{ kind: "layer", target: { partPath: "ppt/slides/slide1.xml", shapeId: 4 }, direction: "up" }]);
    expect(notes.join("\n")).toContain("已经在最底层");
    expect(notes.join("\n")).toContain("已经在最顶层");
  });

  it("组合里的那一格：改文字照旧交出去，改几何与越界的调层级换成一句理由（引擎那两条会整批拒）", () => {
    const { requests, notes } = slideRequestsOfProposal(deckWithGroup(), [
      { type: "set_run_text", slideNumber: 1, shapeId: 5, paraIndex: 0, runIndex: 0, text: "改后的组内甲" },
      { type: "set_shape_geometry", slideNumber: 1, shapeId: 5, leftPx: 10, topPx: 10, widthPx: 100, heightPx: 100 },
      { type: "move_shape_layer", slideNumber: 1, shapeId: 6, direction: "up" },
      { type: "move_shape_layer", slideNumber: 1, shapeId: 9, direction: "up" },
    ]);
    expect(requests).toEqual([
      { kind: "text", target: { partPath: "ppt/slides/slide1.xml", shapeId: 5, paraIndex: 0, runIndex: 0 }, text: "改后的组内甲" },
    ]);
    const joined = notes.join("\n");
    expect(joined).toContain("组内子坐标系"); // 几何那条：它的坐标写在组的子坐标系里
    expect(joined).toContain("组内最顶层"); // 组里最后一格向上：只比组内的邻居
    expect(joined).toContain("已经在最顶层"); // 组自己向上：页面那一层它确实在最顶（后面那两格是它的内容，不是邻居）
  });

  it("加文本框只定位到整页（新形状号要引擎落地时才分配），删形状定位到形状", () => {
    const { requests, notes } = slideRequestsOfProposal(sampleDeck(), [
      { type: "add_text_box", slideNumber: 2, leftPx: 96, topPx: 96, widthPx: 480, heightPx: 144, text: "谢谢观看" },
      { type: "remove_shape", slideNumber: 1, shapeId: 7 },
    ]);
    expect(notes).toEqual([]);
    expect(requests).toEqual([
      {
        kind: "add_text_box",
        target: { partPath: "ppt/slides/slide2.xml" },
        box: { leftPx: 96, topPx: 96, widthPx: 480, heightPx: 144 },
        text: "谢谢观看",
      },
      { kind: "remove_shape", target: { partPath: "ppt/slides/slide1.xml", shapeId: 7 } },
    ]);

    const pageMissing = slideRequestsOfProposal(sampleDeck(), [
      { type: "add_text_box", slideNumber: 9, leftPx: 0, topPx: 0, widthPx: 100, heightPx: 100, text: "甲" },
    ]);
    expect(pageMissing.requests).toEqual([]);
    expect(pageMissing.notes[0]).toContain("没有第 9 页");
  });

  it("同一批里先删掉的形状，后面那条改它的不交出（按顺序落地时那条必然失败，整批就不写了，§2.4）", () => {
    const { requests, notes } = slideRequestsOfProposal(sampleDeck(), [
      { type: "remove_shape", slideNumber: 1, shapeId: 4 },
      { type: "set_run_text", slideNumber: 1, shapeId: 4, paraIndex: 0, runIndex: 0, text: "改不到了" },
      { type: "remove_shape", slideNumber: 1, shapeId: 4 },
      { type: "set_run_text", slideNumber: 1, shapeId: 2, paraIndex: 0, runIndex: 0, text: "这条还能落" },
    ]);
    expect(requests).toEqual([
      { kind: "remove_shape", target: { partPath: "ppt/slides/slide1.xml", shapeId: 4 } },
      { kind: "text", target: { partPath: "ppt/slides/slide1.xml", shapeId: 2, paraIndex: 0, runIndex: 0 }, text: "这条还能落" },
    ]);
    expect(notes).toHaveLength(2);
    expect(notes.join("\n")).toContain("已经被删掉");
    // 重复删同一个形状同理拦下：现场模型是提案前那一份，光查"这一页有没有它"拦不住第二次删
    expect(notes[1]).toContain("「TextBox 3」在同一批里已经被删掉");
  });

  it("几何继承自版式的形状（标题占位符那种）只给一轴时不交出：引擎写出来会缺必填属性，整批不落地（§2.4）", () => {
    const { requests, notes } = slideRequestsOfProposal(sampleDeck(), [
      { type: "set_shape_geometry", slideNumber: 1, shapeId: 2, topPx: 60 },
      { type: "set_shape_geometry", slideNumber: 1, shapeId: 4, heightPx: 200 },
      { type: "set_shape_geometry", slideNumber: 1, shapeId: 2, leftPx: 96, topPx: 60, widthPx: 480, heightPx: 120 },
    ]);
    // 形状 4 自己写了几何，单改高度合法；形状 2 没写，只有四个一起给的那条能落
    expect(requests).toEqual([
      { kind: "geometry", target: { partPath: "ppt/slides/slide1.xml", shapeId: 4 }, box: { heightPx: 200 } },
      { kind: "geometry", target: { partPath: "ppt/slides/slide1.xml", shapeId: 2 }, box: { leftPx: 96, topPx: 60, widthPx: 480, heightPx: 120 } },
    ]);
    expect(notes).toHaveLength(1);
    expect(notes[0]).toContain("从版式继承");
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

  it("表格那一格：带行与列翻成 table_text，shapeId 用图形框那个", () => {
    const { requests, notes } = slideRequestsOfProposal(deckWithTable(), [
      { type: "set_table_cell_text", slideNumber: 1, shapeId: 10, row: 0, col: 1, paraIndex: 0, runIndex: 0, text: "省份" },
      { type: "set_table_cell_text", slideNumber: 1, shapeId: 10, row: 1, col: 0, paraIndex: 1, runIndex: 1, text: "改为两截的第二截" },
    ]);
    expect(notes).toEqual([]);
    expect(requests).toEqual([
      { kind: "table_text", target: { partPath: "ppt/slides/slide1.xml", shapeId: 10, row: 0, col: 1, paraIndex: 0, runIndex: 0 }, text: "省份" },
      {
        kind: "table_text",
        target: { partPath: "ppt/slides/slide1.xml", shapeId: 10, row: 1, col: 0, paraIndex: 1, runIndex: 1 },
        text: "改为两截的第二截",
      },
    ]);
  });

  it("表格那一格的四种改不了：图形框里不是表、行列越界、合并续格、段/run 越界——一条都不许交出（§2.4）", () => {
    const { requests, notes } = slideRequestsOfProposal(deckWithTable(), [
      { type: "set_table_cell_text", slideNumber: 1, shapeId: 7, row: 0, col: 0, paraIndex: 0, runIndex: 0, text: "图片里没有表" },
      { type: "set_table_cell_text", slideNumber: 1, shapeId: 10, row: 5, col: 0, paraIndex: 0, runIndex: 0, text: "越界的行" },
      { type: "set_table_cell_text", slideNumber: 1, shapeId: 10, row: 0, col: 3, paraIndex: 0, runIndex: 0, text: "越界的列" },
      { type: "set_table_cell_text", slideNumber: 1, shapeId: 10, row: 1, col: 1, paraIndex: 0, runIndex: 0, text: "并进来的那一格" },
      { type: "set_table_cell_text", slideNumber: 1, shapeId: 10, row: 0, col: 0, paraIndex: 2, runIndex: 0, text: "越界的段" },
      { type: "set_table_cell_text", slideNumber: 1, shapeId: 10, row: 1, col: 0, paraIndex: 0, runIndex: 1, text: "越界的 run" },
    ]);
    expect(requests).toEqual([]);
    expect(notes).toHaveLength(6);
    const all = notes.join("\n");
    expect(all).toContain("没有表格");
    expect(all).toContain("行号 5 越界（这张表有 2 行）");
    expect(all).toContain("列号 3 越界（第 0 行有 2 格）");
    expect(all).toContain("它不画自己的字");
    expect(all).toContain("只有 1 段，没有第 2 段");
    expect(all).toContain("只有 1 个 run，没有第 1 个");
  });

  it("表格那一格与同批别的改动混着提：拒掉的那条只牺牲自己，别条照旧交出", () => {
    const { requests, notes } = slideRequestsOfProposal(deckWithTable(), [
      { type: "set_table_cell_text", slideNumber: 1, shapeId: 10, row: 1, col: 1, paraIndex: 0, runIndex: 0, text: "并进来的那一格" },
      { type: "set_run_text", slideNumber: 1, shapeId: 2, paraIndex: 0, runIndex: 0, text: "新标题" },
    ]);
    expect(notes).toHaveLength(1);
    expect(requests).toHaveLength(1);
    expect(requests[0]!.kind).toBe("text");
  });

  it("表格那一格的字号/颜色：带行与列就翻成带行列的 font / color 坐标（引擎那两条 op 认得，不新开 kinds）", () => {
    const { requests, notes } = slideRequestsOfProposal(deckWithTable(), [
      { type: "set_run_font", slideNumber: 1, shapeId: 10, row: 0, col: 1, paraIndex: 0, runIndex: 0, sizePt: 20 },
      { type: "set_run_color", slideNumber: 1, shapeId: 10, row: 1, col: 0, paraIndex: 1, runIndex: 1, color: "#ff0000" },
      // 不带行列的旧条目照旧走普通形状那条链：一个字段都不多给
      { type: "set_run_font", slideNumber: 1, shapeId: 2, paraIndex: 0, runIndex: 0, bold: false },
    ]);
    expect(notes).toEqual([]);
    expect(requests).toEqual([
      { kind: "font", target: { partPath: "ppt/slides/slide1.xml", shapeId: 10, row: 0, col: 1, paraIndex: 0, runIndex: 0 }, format: { bold: undefined, sizePt: 20 } },
      { kind: "color", target: { partPath: "ppt/slides/slide1.xml", shapeId: 10, row: 1, col: 0, paraIndex: 1, runIndex: 1 }, hex: "#ff0000" },
      { kind: "font", target: { partPath: "ppt/slides/slide1.xml", shapeId: 2, paraIndex: 0, runIndex: 0 }, format: { bold: false, sizePt: undefined } },
    ]);
  });

  it("表格那一格的字号/颜色四种改不了：不是表、行列越界、合并续格、run 越界——一条都不许交出，别条照旧落地（§2.4）", () => {
    const { requests, notes } = slideRequestsOfProposal(deckWithTable(), [
      { type: "set_run_font", slideNumber: 1, shapeId: 7, row: 0, col: 0, paraIndex: 0, runIndex: 0, bold: true },
      { type: "set_run_color", slideNumber: 1, shapeId: 10, row: 5, col: 0, paraIndex: 0, runIndex: 0, color: "#00ff00" },
      { type: "set_run_font", slideNumber: 1, shapeId: 10, row: 1, col: 1, paraIndex: 0, runIndex: 0, sizePt: 18 },
      { type: "set_run_color", slideNumber: 1, shapeId: 10, row: 0, col: 0, paraIndex: 0, runIndex: 3, color: "#0000ff" },
      // 同批里合法的那一条不能因为上面四条被牺牲
      { type: "set_run_font", slideNumber: 1, shapeId: 10, row: 0, col: 0, paraIndex: 0, runIndex: 0, bold: true },
    ]);
    expect(requests).toHaveLength(1);
    expect(requests[0]!.kind).toBe("font");
    expect(notes).toHaveLength(4);
    const all = notes.join("\n");
    expect(all).toContain("没有表格");
    expect(all).toContain("行号 5 越界（这张表有 2 行）");
    expect(all).toContain("它不画自己的字");
    expect(all).toContain("只有 1 个 run，没有第 3 个");
  });

  it("set_notes 把页号翻成备注部件路径：这一档只改整篇，带的是那一页自己的 partPath", () => {
    const { requests, notes } = slideRequestsOfProposal(sampleDeck(), [
      { type: "set_notes", slideNumber: 1, paragraphTexts: ["新备注一句", "第二句"] },
    ]);
    expect(notes).toEqual([]);
    expect(requests).toEqual([
      {
        kind: "set_notes",
        target: { notesPartPath: "ppt/notesSlides/notesSlide1.xml" },
        paragraphTexts: ["新备注一句", "第二句"],
      },
    ]);
  });

  it("那一页根本没有备注部件时不交出：那是 create_notes 的地盘，set 这一档交出去主进程必然拒，整批就不写了（§2.4）", () => {
    const { requests, notes } = slideRequestsOfProposal(sampleDeck(), [
      { type: "set_notes", slideNumber: 2, paragraphTexts: ["第 2 页没有备注页"] },
    ]);
    expect(requests).toEqual([]);
    expect(notes[0]).toContain("没有备注页");
  });

  it("备注里有多截不同格式的文字时不交出：整篇覆盖会压平那一截（§2.3），理由用引擎那份共用判定", () => {
    const styled = { ...SAMPLE_NOTES, paragraphRunCounts: [1, 2], styledRunCount: 1 };
    const { requests, notes } = slideRequestsOfProposal(sampleDeck(styled), [
      { type: "set_notes", slideNumber: 1, paragraphTexts: ["压平了不行"] },
    ]);
    expect(requests).toEqual([]);
    expect(notes[0]).toContain("压平");
  });

  it("create_notes 把页号翻成**幻灯片**部件路径：引擎要在那一页上挂新备注页，不是拿备注路径当坐标", () => {
    const { requests, notes } = slideRequestsOfProposal(sampleDeck(), [
      { type: "create_notes", slideNumber: 2, paragraphTexts: ["第 2 页的第一句备注"] },
    ]);
    expect(notes).toEqual([]);
    expect(requests).toEqual([
      { kind: "create_notes", target: { partPath: "ppt/slides/slide2.xml" }, paragraphTexts: ["第 2 页的第一句备注"] },
    ]);
  });

  it("create_notes 的三种提不了：已经有备注页、这份文件没有备注母版、同一批里已经建过——一条都不许交出（§2.4 整批不落地）", () => {
    const { requests, notes } = slideRequestsOfProposal(sampleDeck(), [
      // ① 那一页已经有备注部件，那是 set_notes 的地盘；提 create 会在同一页上造出第二份备注页
      { type: "create_notes", slideNumber: 1, paragraphTexts: ["第 1 页已经有备注页了"] },
      // ③ 同一页在同一批里排了两条：落地按顺序执行，第二条在引擎那儿必然看到"已经有备注页了"
      { type: "create_notes", slideNumber: 2, paragraphTexts: ["第一份"] },
      { type: "create_notes", slideNumber: 2, paragraphTexts: ["多余的第二份"] },
    ]);
    expect(requests).toHaveLength(1);
    expect(requests[0]!.kind).toBe("create_notes");
    const all = notes.join("\n");
    expect(all).toContain("已经有备注页了");
    expect(all).toContain("已经建过一次");

    // ② 这份文件没有备注母版：新建的备注页没有可继承的版式，绝不自己造一份（§2.7）
    const noMaster = slideRequestsOfProposal(sampleDeck(null, false), [
      { type: "create_notes", slideNumber: 2, paragraphTexts: ["没有备注母版就建不了"] },
    ]);
    expect(noMaster.requests).toEqual([]);
    expect(noMaster.notes[0]).toContain("没有备注母版");
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
