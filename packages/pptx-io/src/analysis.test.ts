/**
 * 分析读面的用例（`analysis.ts`，spec 阶段 5 / D6）：画像与逐 run 清单到底给模型看什么。
 *
 * 钉的是两件红线事：**上送内容必须截断并说明截断了**（`AGENTS.md` §2.9）、
 * **页号/形状号/段号/run 号这些事实由引擎算**（§2.11，模型不许自己数）。
 * 夹具走 `buildDeckBytes`：让断言落在真解析器读出来的模型上，而不是我手搓的一份对象。
 */
import { describe, expect, it } from "vitest";
import { readPresentation } from "./read.js";
import { buildDeckBytes, minimalDeckParts } from "./test-deck.js";
import { describeDeck, readSlideText } from "./analysis.js";

/** 一条 130 字的 run 文字（超过上送上限 120 字，用来验截断真的发生）。 */
const LONG_TEXT = "长".repeat(130);

/**
 * 造一份"给分析面用"的 deck：一页三个形状（占位符 / 自带几何的文本框 / 图片），文本框里一段两条 run、一段一条 run。
 *
 * @returns `.pptx` 的字节。
 */
async function analysisDeck(): Promise<Uint8Array> {
  const parts = minimalDeckParts();
  parts["ppt/slides/slide1.xml"] =
    '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    '<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>' +
    '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/>' +
    "<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>标题甲</a:t></a:r></a:p></p:txBody></p:sp>" +
    '<p:sp><p:nvSpPr><p:cNvPr id="4" name="TextBox 3"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>' +
    '<p:spPr><a:xfrm><a:off x="914400" y="3657600"/><a:ext cx="2743200" cy="914400"/></a:xfrm></p:spPr>' +
    "<p:txBody><a:bodyPr/><a:lstStyle/>" +
    `<a:p><a:r><a:rPr b="1" sz="1800"/><a:t>${LONG_TEXT}</a:t></a:r><a:r><a:rPr><a:solidFill><a:srgbClr val="C00000"/></a:solidFill></a:rPr><a:t>尾</a:t></a:r></a:p>` +
    "<a:p><a:r><a:t>最后一行</a:t></a:r></a:p></p:txBody></p:sp>" +
    '<p:pic><p:nvPicPr><p:cNvPr id="7" name="Picture 6"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill/><p:spPr/></p:pic>' +
    '<p:sp><p:nvSpPr><p:cNvPr id="8" name="Slide Number 7"/><p:cNvSpPr/><p:nvPr><p:ph type="sldNum" idx="4"/></p:nvPr></p:nvSpPr><p:spPr/>' +
    "<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:fld id=\"{6E7F9B92-8F4B-4D8E-9E2A-1C0F0A7B3E11}\" type=\"slidenum\"><a:t>2</a:t></a:fld></a:p></p:txBody></p:sp>" +
    "</p:spTree></p:cSld></p:sld>";
  return buildDeckBytes(parts);
}

describe("describeDeck：给模型的画像", () => {
  it("页数与页面尺寸换算成 px，形状带着 shapeId / 占位符角色 / 是否自带几何", async () => {
    const { model } = await readPresentation(await analysisDeck());
    const outline = describeDeck(model, { maxSlides: 50, maxShapesPerSlide: 40 });
    expect(outline.slideCount).toBe(1);
    expect(outline.slidesTruncated).toBe(false);
    // 9144000 EMU = 960 px @96DPI：模型视角一律 px（§2.1 换算留在引擎边）。
    expect(outline.slideSizePx).toEqual({ widthPx: 960, heightPx: 720 });
    const slide = outline.slides[0]!;
    expect(slide.slideNumber).toBe(1);
    expect(slide.shapeCount).toBe(4);
    expect(slide.shapes.map((shape) => shape.shapeId)).toEqual([2, 4, 7, 8]);
    expect(slide.shapes[0]).toMatchObject({ placeholderRef: "title", hasOwnGeometry: false, text: "标题甲" });
    expect(slide.shapes[1]).toMatchObject({ placeholderRef: null, hasOwnGeometry: true });
    expect(slide.shapes[2]).toMatchObject({ kind: "picture", text: "" });
  });

  it("上限一到就只列前面那些，并如实标 truncated——页数与形状数仍是全量（§2.9）", async () => {
    const { model } = await readPresentation(await analysisDeck());
    const outline = describeDeck(model, { maxSlides: 50, maxShapesPerSlide: 2 });
    expect(outline.slides[0]!.shapes.map((shape) => shape.shapeId)).toEqual([2, 4]);
    expect(outline.slides[0]!.shapeCount).toBe(4);
    expect(outline.slides[0]!.shapesTruncated).toBe(true);
  });

  it("单条概要超过 120 字就切断并加省略号", async () => {
    const { model } = await readPresentation(await analysisDeck());
    const textbox = describeDeck(model, { maxSlides: 50, maxShapesPerSlide: 40 }).slides[0]!.shapes[1]!;
    expect(textbox.text.length).toBe(121); // 120 字 + 一个省略号
    expect(textbox.text.endsWith("…")).toBe(true);
  });
});

describe("readSlideText：逐 run 的文字与格式", () => {
  it("按「段—run」列出定位坐标与现格式，长文字截断但坐标不截", async () => {
    const { model } = await readPresentation(await analysisDeck());
    const read = readSlideText(model, 1, 50);
    if (!read.ok) throw new Error(read.reason);
    // 图片没有文本体：列出来只会诱导模型去改它，所以整个跳过。
    expect(read.report.shapes.map((shape) => shape.shapeId)).toEqual([2, 4, 8]);
    const textbox = read.report.shapes[1]!;
    expect(textbox.singleRunParagraphs).toEqual([false, true]);
    expect(textbox.runs).toHaveLength(3);
    expect(textbox.runs[0]).toMatchObject({ paraIndex: 0, runIndex: 0, textTruncated: true, format: { bold: true, sizePt: 18 } });
    expect(textbox.runs[0]!.text.length).toBe(121);
    expect(textbox.runs[1]).toMatchObject({ paraIndex: 0, runIndex: 1, text: "尾", format: { hex: "#c00000" } });
    expect(textbox.runs[2]).toMatchObject({ paraIndex: 1, runIndex: 0, text: "最后一行", format: {} });
  });

  it("文字来自字段（`a:fld`：页码 / 日期）的段落列出 0 条 run，模型就没法指到一个 run", async () => {
    const { model } = await readPresentation(await analysisDeck());
    const read = readSlideText(model, 1, 50);
    if (!read.ok) throw new Error(read.reason);
    const slideNumberShape = read.report.shapes[2]!;
    expect(slideNumberShape.runs).toEqual([]);
    expect(slideNumberShape.singleRunParagraphs).toEqual([false]); // 视图与引擎都据此拒绝整段覆盖
  });

  it("每段 run 数有上限，超出只列前面那些并标 runsTruncated", async () => {
    const { model } = await readPresentation(await analysisDeck());
    const read = readSlideText(model, 1, 1);
    if (!read.ok) throw new Error(read.reason);
    expect(read.report.runsTruncated).toBe(true);
    expect(read.report.shapes[1]!.runs.map((run) => [run.paraIndex, run.runIndex])).toEqual([
      [0, 0],
      [1, 0],
    ]);
  });

  it("页号越界回一句中文理由，不抛异常也不返回半份数据", async () => {
    const { model } = await readPresentation(await analysisDeck());
    const read = readSlideText(model, 9, 50);
    expect(read.ok).toBe(false);
    if (!read.ok) expect(read.reason).toContain("一共 1 页");
  });

  it("**这一页读不动**时回理由而不是回一份空形状清单（spec line 439）", async () => {
    const parts = minimalDeckParts();
    delete parts["ppt/slides/slide1.xml"]; // 关系表还指向它，部件本身不在包里
    const { model } = await readPresentation(await buildDeckBytes(parts));
    const read = readSlideText(model, 1, 50);
    expect(read.ok).toBe(false);
    if (!read.ok) expect(read.reason).toContain("读不动");
    if (!read.ok) expect(read.reason).toContain("部件");
    const outline = describeDeck(model, { maxSlides: 20, maxShapesPerSlide: 20 });
    expect(outline.slides[0]!.readFailure).toBe("missing-part");
    expect(outline.slides[0]!.shapeCount).toBe(0);
  });

  it("读得动的那一页画像里失败码是 null（空形状才真的是没形状）", async () => {
    const { model } = await readPresentation(await analysisDeck());
    expect(describeDeck(model, { maxSlides: 20, maxShapesPerSlide: 20 }).slides[0]!.readFailure).toBeNull();
  });
});
