/**
 * `.pdf` 的 AI 读面用例（spec 阶段 7 的 D6 那一半）。
 *
 * 钉四条：① 模型看到的只有页号（1 基）与该页内 run 序号（0 基），`streamRef` 与起始下标跨不出引擎；
 * ② `describePdf` 与 `readPdfPageText` 的序号必须由同一个排序算出来，否则模型照概要提的号会指错那条；
 * ③ 上送内容截断要如实标记（§2.9），越界页号回一句中文理由而不是抛异常（B9）；
 * ④ 改不了的 run 照样列出来并带码说明，别让模型对着一页看不见的文字瞎猜。
 */
import { describe, expect, it } from "vitest";
import { describePdf, readPdfModel, readPdfPageText, type PdfReadModel, type PdfTextRun } from "./index.js";
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
      run({ start: 60, withinForm: "Fm0", text: "中文标题", composite: true, baseFont: "SourceHanSansCN-Regular", encodedBytes: 8 }),
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

  it("真文件上 describe 的概要就是逐条文字拼出来的：模型照任一边的号提改动都指同一条", async () => {
    const bytes = await simpleWinAnsiFixturePdf();
    const real = await readPdfModel(bytes);
    const outline = describePdf(real, { maxPages: 50 });
    for (const page of outline.pages) {
      const read = readPdfPageText(real, page.pageNumber, 50);
      if (!read.ok) throw new Error("该读出来的：" + read.reason);
      expect(read.report.runCount).toBe(page.runCount);
      expect(read.report.runs.map((item) => item.text).filter((text) => text !== "").join("⏎")).toBe(page.text);
    }
  });
});
