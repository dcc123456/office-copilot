import { describe, expect, it } from "vitest";
import { EDITABLE_CELL_LIMIT, readOnlyReason, usedCellCount } from "./sheet-dto.js";
import type { SheetCell, SheetPayload } from "./model.js";

/**
 * 造一张只有格子数的表，其余字段对本次判定无关。
 *
 * @param cellCount 要生成的已用格数量（0 基行号，地址是 `A1..A{cellCount}`）。
 * @returns 一个合法的 `SheetPayload`，`cells` 里每格都是数字 1。
 */
function sheetWithCells(cellCount: number): SheetPayload {
  const cells: Record<string, SheetCell> = {};
  for (let rowIndex = 0; rowIndex < cellCount; rowIndex++) {
    const address = `A${rowIndex + 1}`;
    cells[address] = { address, colIndex: 0, rowIndex, kind: "number", value: 1 };
  }
  return {
    name: `表${cellCount}`,
    sheetId: 1,
    partPath: "xl/worksheets/sheet1.xml",
    lastColIndex: cellCount === 0 ? -1 : 0,
    lastRowIndex: cellCount === 0 ? -1 : cellCount - 1,
    cells,
    merges: [],
    columns: [],
    rows: [],
    warnings: [],
  };
}

describe("跨界载荷的只读上限判定（spec D19 / §七 大表缓解）", () => {
  it("空工作簿是 0 格，可编辑", () => {
    expect(usedCellCount({ sheets: [] })).toBe(0);
    expect(readOnlyReason({ sheets: [] })).toBeNull();
  });

  it("格子数按所有表相加，含只有格式的格", () => {
    const workbook = { sheets: [sheetWithCells(3), sheetWithCells(4)] };
    expect(usedCellCount(workbook)).toBe(7);
    expect(readOnlyReason(workbook)).toBeNull();
  });

  it("正好在上限内仍可编辑，超一个格就只读", () => {
    expect(readOnlyReason({ sheets: [sheetWithCells(EDITABLE_CELL_LIMIT)] })).toBeNull();
    const reason = readOnlyReason({ sheets: [sheetWithCells(EDITABLE_CELL_LIMIT + 1)] });
    expect(reason).not.toBeNull();
    // 文案必须把实测数量与上限都报出来，用户才知道要拆多大
    expect(reason).toContain(String(EDITABLE_CELL_LIMIT + 1));
    expect(reason).toContain(String(EDITABLE_CELL_LIMIT));
  });

  it("只读判定不含千分位分组，跨 ICU 环境文案稳定", () => {
    const reason = readOnlyReason({ sheets: [sheetWithCells(EDITABLE_CELL_LIMIT + 1)] });
    expect(reason).not.toMatch(/,\d{3}\b/);
  });
});
