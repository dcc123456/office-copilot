/**
 * `.pdf` 的 AI 读面用例（spec 阶段 7 的 D6 那一半）。
 *
 * 钉五条：① 模型看到的只有页号（1 基）与该页内 run 序号（0 基），`streamRef` 与起始下标跨不出引擎；
 * ② `describePdf` 与 `readPdfPageText` 的序号必须由同一个排序算出来，否则模型照概要提的号会指错那条；
 * ③ 上送内容截断要如实标记（§2.9），越界页号回一句中文理由而不是抛异常（B9）；
 * ④ 改不了的 run 照样列出来并带码说明，别让模型对着一页看不见的文字瞎猜；
 * ⑤ 概要那一栏按**逻辑文本行**拼（一个行内多条 run 之间不再有 `⏎`），但逐条那一面仍是 run 粒度——
 *   改了行就不许顺手改定位坐标，那是模型提交改动时唯一的把手（spec D8 的 W1 档）。
 */
import { describe, expect, it } from "vitest";
import { describePdf, pdfTextLinesOf, readPdfModel, readPdfPageText, type PdfReadModel, type PdfTextRun } from "./index.js";
import { simpleWinAnsiFixturePdf } from "./test-pdf.js";

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
    fontSizePt: 11,
    origin: { x: 72, y: 720 },
    text: "Quarterly Report 2026",
    ...overrides,
  };
}

/** 造一份三页的模型：第 1 页两条（其中一条在 Form 里）、第 2 页没有文字、第 3 页一条内联 + 一条 CID。 */
function model(): PdfReadModel {
  return {
    pageCount: 3,
    runs: [
      run(),
      run({ start: 60, withinForm: "Fm0", text: "中文标题", composite: true, baseFont: "SourceHanSansCN-Regular", encodedBytes: 8, origin: { x: 72, y: 700 } }),
      run({ pageIndex: 2, streamRef: null, text: "内联的文字" }),
      run({ pageIndex: 2, start: 10, text: null, baseFont: null }),
    ],
  };
}

describe("describePdf：页号 1 基，没文字的页也在场", () => {
  it("每页给 run 数、可改数与概要，页号从 1 开始连续", () => {
    const outline = describePdf(model(), { maxPages: 50 });
    expect(outline.pageCount).toBe(3);
    expect(outline.pages.map((page) => page.pageNumber)).toEqual([1, 2, 3]);
    expect(outline.pages[1]).toMatchObject({ runCount: 0, editableCount: 0, text: "" });
    expect(outline.pages[0]).toMatchObject({ runCount: 2, editableCount: 2, text: "Quarterly Report 2026⏎中文标题" });
    expect(outline.pages[2]).toMatchObject({ runCount: 2, editableCount: 0, text: "内联的文字" });
    expect(outline.pagesTruncated).toBe(false);
  });

  it("maxPages 截断时只列前面那些，页数仍是全量并置标记", () => {
    const outline = describePdf(model(), { maxPages: 2 });
    expect(outline.pages.map((page) => page.pageNumber)).toEqual([1, 2]);
    expect(outline.pageCount).toBe(3);
    expect(outline.pagesTruncated).toBe(true);
  });

  it("同一逻辑行的多条 run（WPS 那种逐字形产物）在概要里拼成一行，逐条面仍是 run 粒度", () => {
    const glyphs: PdfReadModel = {
      pageCount: 1,
      runs: [
        run({ text: "姓", origin: { x: 90, y: 615 }, fontSizePt: 10.45 }),
        run({ start: 24, text: "名", origin: { x: 100.45, y: 615 }, fontSizePt: 10.45 }),
      ],
    };
    const page = describePdf(glyphs, { maxPages: 5 }).pages[0];
    expect(page?.text).toBe("姓名");
    expect(page?.runCount).toBe(2);
    const read = readPdfPageText(glyphs, 1, 50);
    if (!read.ok) throw new Error("该读出来的：" + read.reason);
    expect(read.report.runs.map((item) => item.text)).toEqual(["姓", "名"]);
  });

  it("概要里的 `⏎` 只出现在换行处：y 差超过 0.6pt 才断行", () => {
    const two: PdfReadModel = {
      pageCount: 1,
      runs: [run({ text: "第一行", origin: { x: 90, y: 615 } }), run({ start: 60, text: "第二行", origin: { x: 90, y: 597 } })],
    };
    expect(describePdf(two, { maxPages: 5 }).pages[0]?.text).toBe("第一行⏎第二行");
  });

  it("概要超 400 字时截断并标记（§2.9：截了要说出来）", () => {
    const long: PdfReadModel = { pageCount: 1, runs: [run({ text: "长".repeat(401) })] };
    const page = describePdf(long, { maxPages: 5 }).pages[0];
    expect(page?.textTruncated).toBe(true);
    expect(page?.text.startsWith("长".repeat(400))).toBe(true);
    expect(page?.text.endsWith("…")).toBe(true);
  });
});

describe("readPdfPageText：逐条 run 的序号、字数与改不了的码", () => {
  it("页号越界回一句中文理由，不抛异常", () => {
    expect(readPdfPageText(model(), 4, 50)).toEqual({ ok: false, reason: "没有第 4 页：这份 PDF 一共 3 页，先用 pdf_describe 看页号。" });
    expect(readPdfPageText(model(), 0, 50).ok).toBe(false);
  });

  it("runIndex 从 0 起连续；字数按码点算，等长规则由它决定（spec D8 的 W1 档）", () => {
    const read = readPdfPageText(model(), 1, 50);
    if (!read.ok) throw new Error("该读出来的：" + read.reason);
    expect(read.report.runs.map((item) => item.runIndex)).toEqual([0, 1]);
    expect(read.report.runs[0]).toMatchObject({ text: "Quarterly Report 2026", charCount: 21, encodedBytes: 21, editable: true, block: null, withinForm: null });
    expect(read.report.runs[1]).toMatchObject({ text: "中文标题", charCount: 4, encodedBytes: 8, composite: true, withinForm: "Fm0" });
  });

  it("改不了的那两条照样列出来，带上各自的码；空文字与 0 字数", () => {
    const read = readPdfPageText(model(), 3, 50);
    if (!read.ok) throw new Error("该读出来的：" + read.reason);
    expect(read.report.runs).toHaveLength(2);
    expect(read.report.runs[0]).toMatchObject({ editable: false, block: "inline-stream", text: "内联的文字", charCount: 5 });
    expect(read.report.runs[1]).toMatchObject({ editable: false, block: "undecoded", text: "", charCount: 0, baseFont: null });
  });

  it("单条超 120 字只上送前 120 字并标记；maxRunsPerPage 截断时给出全量条数", () => {
    const many: PdfReadModel = {
      pageCount: 1,
      runs: [run({ text: "字".repeat(121) }), run({ start: 99, text: "第二条" })],
    };
    const read = readPdfPageText(many, 1, 1);
    if (!read.ok) throw new Error("该读出来的：" + read.reason);
    expect(read.report.runs).toHaveLength(1);
    expect(read.report.runs[0]).toMatchObject({ text: "字".repeat(120) + "…", textTruncated: true, charCount: 121 });
    expect(read.report.runsTruncated).toBe(true);
    expect(read.report.runCount).toBe(2);
  });

  it("真文件上概要是逻辑行拼的、逐条是 run 列的，两边同源：模型照任一边的号提改动都指同一条", async () => {
    const bytes = await simpleWinAnsiFixturePdf();
    const real = await readPdfModel(bytes);
    const outline = describePdf(real, { maxPages: 50 });
    for (const page of outline.pages) {
      const read = readPdfPageText(real, page.pageNumber, 50);
      if (!read.ok) throw new Error("该读出来的：" + read.reason);
      expect(read.report.runCount).toBe(page.runCount);
      const runsOfPage = real.runs.filter((item) => item.pageIndex === page.pageNumber - 1);
      const lines = pdfTextLinesOf(runsOfPage);
      if (!page.textTruncated) expect(lines.map((line) => line.text).filter((text) => text !== "").join("⏎")).toBe(page.text);
      // 分组只改"怎么断行"，不许改文字本身：把行首尾相接必须等于把 run 首尾相接。
      expect(lines.map((line) => line.text).join("")).toBe(runsOfPage.map((item) => item.text ?? "").join(""));
    }
  });
});
