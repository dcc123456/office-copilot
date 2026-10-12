import { describe, expect, it } from "vitest";
import { runSheetOps, setCellTextOp, setColumnWidthOp, type SheetCell, type SheetPayload } from "@office-copilot/sheet-io";
import { isSheetEditRecord, pickLatestSheetEdit, sheetEditRecordOf, undoSheetEdit } from "./sheetHistory.js";

/**
 * 造一张最小表（与 IO 包的测试夹具同形状）。
 *
 * @param cells 已用的格。
 * @returns `SheetPayload`；`partPath` 固定，好和记录里的配对。
 */
function sheetOf(cells: SheetCell[]): SheetPayload {
  return {
    name: "销售",
    sheetId: 1,
    partPath: "xl/worksheets/sheet1.xml",
    lastColIndex: cells.reduce((acc, cell) => Math.max(acc, cell.colIndex), 0),
    lastRowIndex: cells.reduce((acc, cell) => Math.max(acc, cell.rowIndex), 0),
    cells: Object.fromEntries(cells.map((cell) => [cell.address, cell])),
    merges: [],
    columns: [{ minColIndex: 0, maxColIndex: 3, widthChars: 12, customWidth: true }],
    rows: [],
    warnings: [],
  };
}

/** 造一格：列/行号 + 值，地址按 0 基算出来。 */
function cellOf(colIndex: number, rowIndex: number, value: number | string): SheetCell {
  return {
    address: `${String.fromCharCode(65 + colIndex)}${rowIndex + 1}`,
    colIndex,
    rowIndex,
    kind: typeof value === "number" ? "number" : "string",
    value,
  };
}

describe("表格改动写进修改历史（spec 阶段 3 · C15/C17）", () => {
  it("记录里只留真正落地的 op，跳过的条目不进历史", () => {
    const sheet = sheetOf([cellOf(0, 0, 1)]);
    const otherSheetOp = setCellTextOp(sheet, "A1", "9")!;
    const result = runSheetOps(sheet, [
      setCellTextOp(sheet, "B2", "12")!,
      { ...otherSheetOp, partPath: "xl/worksheets/sheet2.xml" },
    ]);
    const record = sheetEditRecordOf(result, "补一行合计");
    expect(record.ops).toHaveLength(1);
    expect(record.inverses).toHaveLength(1);
    expect(record.partPath).toBe(sheet.partPath);
    expect(record.label).toBe("表格改动");
    expect(record.summary).toBe("补一行合计");
    expect(isSheetEditRecord(record)).toBe(true);
  });

  it("文档侧的记录与已撤销过的记录都不算表格的最近改动", () => {
    const sheet = sheetOf([cellOf(0, 0, 1)]);
    const record = sheetEditRecordOf(runSheetOps(sheet, [setCellTextOp(sheet, "A1", "2")!]), "改 A1");
    const docRecord = { id: "d", ts: record.ts + 10, label: "AI 修改", summary: "整篇", ops: [], before: [], after: [] };
    expect(pickLatestSheetEdit([docRecord, record])).toEqual(record);
    expect(pickLatestSheetEdit([{ ...record, undone: true }])).toBeNull();
    expect(pickLatestSheetEdit([])).toBeNull();
  });

  it("撤销把值倒回去；之后的手工编辑不牵连（§2.5）", () => {
    const sheet = sheetOf([cellOf(0, 0, 1), cellOf(0, 1, 300)]);
    const result = runSheetOps(sheet, [setCellTextOp(sheet, "A1", "999")!]);
    // 用户在那次改动之后自己把 B2 改了：撤销 A1 不许顺手把它也退回去
    const withManualEdit = {
      ...result.after,
      cells: { ...result.after.cells, A2: cellOf(0, 1, "手工改过") },
    };
    const record = sheetEditRecordOf(result, "改 A1");
    const outcome = undoSheetEdit(withManualEdit, record);
    expect(outcome.appliedCount).toBe(1);
    expect(outcome.failed).toEqual([]);
    expect(outcome.after.cells["A1"]?.value).toBe(1);
    expect(outcome.after.cells["A2"]?.value).toBe("手工改过");
  });

  it("现场被改过时报部分撤销（x/y），不猜位置", () => {
    const sheet = sheetOf([cellOf(0, 0, 1)]);
    const result = runSheetOps(sheet, [
      setCellTextOp(sheet, "A1", "999")!,
      setColumnWidthOp(sheet, 1, 103)!,
    ]);
    const record = sheetEditRecordOf(result, "改数值并加宽 B 列");
    const overwritten = { ...result.after, cells: { ...result.after.cells, A1: cellOf(0, 0, 50) } };
    const outcome = undoSheetEdit(overwritten, record);
    // 列宽那条能撤销（内容寻址认得出这一列），改值那条被现场核对拦下
    expect(outcome.appliedCount).toBe(1);
    expect(outcome.failed).toHaveLength(1);
    expect(outcome.failed[0]).toContain("认不出这一行");
    expect(outcome.after.cells["A1"]?.value).toBe(50);
    expect(outcome.after.columns).toEqual(sheet.columns);
  });

  it("记录属于另一张表时整条不落地", () => {
    const sheet = sheetOf([cellOf(0, 0, 1)]);
    const record = sheetEditRecordOf(runSheetOps(sheet, [setCellTextOp(sheet, "A1", "2")!]), "改 A1");
    const outcome = undoSheetEdit({ ...sheet, partPath: "xl/worksheets/sheet9.xml" }, record);
    expect(outcome.appliedCount).toBe(0);
    expect(outcome.reason).toContain("另一张表");
  });
});
