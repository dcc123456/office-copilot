/**
 * PDF 视图的判定层（阶段 7 的装配）。
 *
 * 钉三条事实：① 能不能改的判定只有一份（`pdf-io` 的 `runEditBlock`），这里只验它翻出来的中文理由；② 草稿只对
 * **当前读模型**里的可编辑 run 生效，改回原文的那条不进批次；③ 字节预算不在这里判——渲染层再算一遍编码就是第二套换算
 * （§2.1），所以只把引擎的错误码翻成中文。
 */
import { describe, expect, it } from "vitest";
import { isRunEditable, pdfTargetKey, type PdfReadModel, type PdfTextRun, type PdfTextTarget } from "@office-copilot/pdf-io/pure";
import { byteBudgetNote, pdfFailureText, pendingPdfEdits, readOnlyReason, runsOfPage } from "./pdfRuns.js";

/** 造一条可读的 run。 */
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

/** 造一条定位键（草稿表与失败项都按它索引）。 */
function tgt(overrides: Partial<PdfTextTarget> = {}): PdfTextTarget {
  return { pageIndex: 0, streamRef: "6 0 R", withinForm: null, start: 32, ...overrides };
}

/** 造一份两页的读模型：第一页一条可改 + 一条内联，第二页一条 CID 与一条解不出原文的。 */
function model(): PdfReadModel {
  return {
    pageCount: 2,
    runs: [
      run(),
      run({ streamRef: null, start: 60, text: "内联的文字" }),
      run({ pageIndex: 1, start: 10, text: "中文标题", composite: true, baseFont: "SourceHanSansCN-Regular", encodedBytes: 8 }),
      run({ pageIndex: 1, start: 40, text: null }),
    ],
  };
}

describe("pdf 草稿的守卫", () => {
  it("内联流与解不出原文的 run 不给编辑框，理由各说各的", () => {
    expect(isRunEditable(run())).toBe(true);
    expect(isRunEditable(run({ streamRef: null }))).toBe(false);
    expect(readOnlyReason(run({ streamRef: null }))).toContain("内联");
    expect(isRunEditable(run({ text: null }))).toBe(false);
    expect(readOnlyReason(run({ text: null, composite: true }))).toContain("CID");
    expect(readOnlyReason(run())).toBe("");
  });

  it("草稿只对可编辑的 run 生效：改回原文那条不进批次，内联那条压根不收", () => {
    const drafts = {
      [pdfTargetKey(tgt())]: "Quarterly Report 2027",
      [pdfTargetKey({ ...tgt(), start: 60 })]: "不该出现",
      [pdfTargetKey({ ...tgt(), pageIndex: 1, start: 10 })]: "中文标题",
      [pdfTargetKey({ ...tgt(), pageIndex: 1, start: 40 })]: "也不该出现",
    };
    const edits = pendingPdfEdits(model(), drafts);
    expect(edits).toHaveLength(1);
    expect(edits[0]).toEqual({ target: { pageIndex: 0, streamRef: "6 0 R", withinForm: null, start: 32 }, text: "Quarterly Report 2027" });
  });

  it("错误码只翻成中文，视图不自己判字节数：等长那条由引擎回 length-mismatch", () => {
    expect(byteBudgetNote(run())).toContain("21 字节");
    expect(pdfFailureText([{ target: tgt(), reason: "length-mismatch" }, { target: tgt(), reason: "char-outside-subset" }])).toBe(
      "改完的文字装不回原来的字节数；文件里的字体没带这个字，补不出来",
    );
    // 同一个理由出现两次只说一次。
    expect(pdfFailureText([{ target: tgt(), reason: "text-drifted" }, { target: tgt(), reason: "text-drifted" }])).toBe(
      "那段文字后来又被改过，不能拿旧内容覆盖",
    );
  });

  it("按页取 run：翻页只看这一页的", () => {
    expect(runsOfPage(model(), 1).map((item) => item.start)).toEqual([10, 40]);
    expect(runsOfPage(model(), 9)).toEqual([]);
  });
});
