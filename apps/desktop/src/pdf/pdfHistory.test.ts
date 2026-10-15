/**
 * pdf 那一层的记录形状（阶段 7 的 modlog 落盘）。
 *
 * 钉三条：① 只有带 `applied` 且每条有旧/新文字的记录才算 pdf 的——同一 docKey 下混着文档、表格、
 * 幻灯片三类记录，认错形状就会拿 pdf 的撤销去倒回幻灯片批次（§2.5）；② 挑最近一条没撤销的按 `ts`，
 * 不按数组顺序；③ 记录只存改动本身，不存整包字节。
 */
import { describe, expect, it } from "vitest";
import type { PdfAppliedEdit } from "@office-copilot/pdf-io/pure";
import { isPdfEditRecord, pdfEditRecordOf, pickLatestPdfEdit } from "./pdfHistory.js";

const EDIT: PdfAppliedEdit = {
  target: { pageIndex: 0, streamRef: "6 0 R", withinForm: null, start: 32 },
  fromText: "Quarterly Report 2026",
  toText: "Quarterly Report 2027",
};

describe("pdf 改动的历史记录形状", () => {
  it("pdf 记录认得出来，幻灯片与表格那种 ops 记录一律不算", () => {
    const record = pdfEditRecordOf([EDIT], "改标题");
    expect(isPdfEditRecord(record)).toBe(true);
    expect(record.label).toBe("PDF 改动");
    expect(isPdfEditRecord({ id: "x", ts: 1, label: "幻灯片改动", summary: "s", ops: [{ partPath: "a", beforeXml: "b", afterXml: "c" }] })).toBe(false);
    expect(isPdfEditRecord(undefined)).toBe(false);
    expect(isPdfEditRecord({ ...record, applied: [] })).toBe(false);
    expect(isPdfEditRecord({ ...record, applied: [{ ...EDIT, fromText: undefined as unknown as string }] })).toBe(false);
  });

  it("挑最近一条没撤销的按 ts，不是按数组顺序", () => {
    const older = { ...pdfEditRecordOf([EDIT], "早的一批"), id: "o", ts: 100 };
    const newer = { ...pdfEditRecordOf([EDIT], "晚的一批"), id: "n", ts: 200 };
    expect(pickLatestPdfEdit([newer, older])?.summary).toBe("晚的一批");
    expect(pickLatestPdfEdit([newer, { ...older, undone: true }])?.summary).toBe("晚的一批");
    // 晚的那批撤过了就退回早的那批——历史里逐条撤销正是这个语义（§2.5）。
    expect(pickLatestPdfEdit([{ ...newer, undone: true }, older])?.summary).toBe("早的一批");
    expect(pickLatestPdfEdit([{ ...newer, undone: true }, { ...older, undone: true }])).toBeNull();
    expect(pickLatestPdfEdit([])).toBeNull();
  });

  it("记录里不塞整包字节：撤销要的只有定位键与旧文字", () => {
    const record = pdfEditRecordOf([EDIT, { ...EDIT, target: { ...EDIT.target, start: 90 } }], "改两处");
    expect(record.applied).toHaveLength(2);
    expect(JSON.stringify(record)).not.toContain("bytes");
  });
});
