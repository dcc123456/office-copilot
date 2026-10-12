import { describe, expect, it } from "vitest";
import type { SheetPayload } from "@office-copilot/sheet-io";
import {
  ROW_HEADER_WIDTH_PX,
  addressAtPoint,
  buildSheetLayout,
  withColumnWidth,
  displayTextOf,
  moveAddress,
  selectionLabel,
  selectionRectOf,
  visibleSlotRange,
  withCellText,
} from "./sheetGridModel.js";

/** 手工摆一张最小表：A1 文字、B1 数字、C1 带日期样式，另有一列一行被隐藏。 */
function sampleSheet(): SheetPayload {
  return {
    name: "销售",
    sheetId: 1,
    partPath: "xl/worksheets/sheet1.xml",
    lastColIndex: 2,
    lastRowIndex: 1,
    cells: {
      A1: { address: "A1", colIndex: 0, rowIndex: 0, kind: "string", value: "月份", styleIndex: 4 },
      B1: { address: "B1", colIndex: 1, rowIndex: 0, kind: "number", value: 120, styleIndex: 1 },
      C1: {
        address: "C1",
        colIndex: 2,
        rowIndex: 0,
        kind: "number",
        value: 46053,
        styleIndex: 2,
        numberFormat: { numFmtId: 164, formatCode: "yyyy-mm-dd", isDateLike: true },
      },
      A2: { address: "A2", colIndex: 0, rowIndex: 1, kind: "blank", value: null, styleIndex: 0, styleOnly: true },
      D1: { address: "D1", colIndex: 3, rowIndex: 0, kind: "string", value: "没有格式" },
    },
    merges: [],
    columns: [
      { minColIndex: 0, maxColIndex: 0, widthChars: 12, customWidth: true },
      { minColIndex: 3, maxColIndex: 3, hidden: true },
    ],
    rows: [{ rowIndex: 1, heightPt: 22.5, customHeight: true }],
    warnings: [],
  };
}

describe("网格几何（阶段 2：位置与单位的对应关系）", () => {
  it("列宽按 IO 包换算：12 字符是 89px，未设宽度是 Excel 默认 64px", () => {
    const layout = buildSheetLayout(sampleSheet());
    expect(layout.colWidthPx[0]).toBe(89);
    expect(layout.colWidthPx[1]).toBe(64);
    expect(layout.colIndexes).not.toContain(3);
  });

  it("行高 15 磅是 20px，自定义 22.5 磅是 30px，隐藏行不进可见行序列", () => {
    const layout = buildSheetLayout(sampleSheet());
    expect(layout.rowHeightPx[0]).toBe(20);
    expect(layout.rowHeightPx[1]).toBe(30);
    const hiddenRowSheet = sampleSheet();
    hiddenRowSheet.rows = [{ rowIndex: 0, hidden: true }];
    expect(buildSheetLayout(hiddenRowSheet).rowIndexes).not.toContain(0);
  });

  it("滚动窗口按前缀和定位，行号夹在有效区间内", () => {
    const layout = buildSheetLayout(sampleSheet());
    const [startAtTop, endAtTop] = visibleSlotRange(layout.rowOffsetPx, 0, 400, 0);
    expect(startAtTop).toBe(0);
    expect(endAtTop).toBeGreaterThan(5);
    const [startScrolled] = visibleSlotRange(layout.rowOffsetPx, 400, 400, 0);
    expect(startScrolled).toBe(19);
    const [, endAtBottom] = visibleSlotRange(layout.rowOffsetPx, layout.totalHeightPx, 400, 0);
    expect(endAtBottom).toBe(layout.rowIndexes.length);
    expect(ROW_HEADER_WIDTH_PX).toBeGreaterThan(0);
  });
});

describe("输入落地（§2.3 不许拍平已有格式）", () => {
  it("数字输入保留原格的样式索引", () => {
    const next = withCellText(sampleSheet(), "B1", "300");
    expect(next.cells["B1"]).toMatchObject({ kind: "number", value: 300, styleIndex: 1 });
    expect(next.cells["B1"]?.numberFormat).toBeUndefined();
  });

  it("日期输入带出数字格式并放大已用区域", () => {
    const next = withCellText(sampleSheet(), "D9", "2026-02-01");
    expect(next.cells["D9"]).toMatchObject({ kind: "number", value: 46054, colIndex: 3, rowIndex: 8 });
    expect(next.cells["D9"]?.numberFormat).toMatchObject({ numFmtId: 14, isDateLike: true });
    expect(next.lastColIndex).toBe(3);
    expect(next.lastRowIndex).toBe(8);
  });

  it("清空只删内容不删格式：带样式的格退化成只有格式的格，无样式的格整格删掉", () => {
    const styled = withCellText(sampleSheet(), "A1", "");
    expect(styled.cells["A1"]).toMatchObject({ kind: "blank", styleIndex: 4, styleOnly: true });
    expect(displayTextOf(styled.cells["A1"])).toBe("");
    const plain = withCellText(sampleSheet(), "D1", "");
    expect(plain.cells["D1"]).toBeUndefined();
  });

  it("改过的格不许留着旧公式与旧原文", () => {
    const withFormula = sampleSheet();
    withFormula.cells["B1"] = { ...withFormula.cells["B1"]!, formula: "SUM(A1:A2)", cachedValue: 120, rawXml: "<c r=\"B1\"/>" };
    const next = withCellText(withFormula, "B1", "5");
    expect(next.cells["B1"]?.formula).toBeUndefined();
    expect(next.cells["B1"]?.rawXml).toBeUndefined();
  });
});

describe("显示与选区", () => {
  it("日期格显示成日期而不是序列号，空格显示成空串", () => {
    const sheet = sampleSheet();
    expect(displayTextOf(sheet.cells["C1"])).toBe("2026-01-31");
    expect(displayTextOf(sheet.cells["A2"])).toBe("");
    expect(displayTextOf(undefined)).toBe("");
  });

  it("选区两点归一化，状态栏标签单格不带冒号", () => {
    const range = selectionRectOf({ colIndex: 3, rowIndex: 5 }, { colIndex: 1, rowIndex: 2 });
    expect(range).toEqual({ start: { colIndex: 1, rowIndex: 2 }, end: { colIndex: 3, rowIndex: 5 } });
    expect(selectionLabel(range)).toBe("B3:D6");
    expect(selectionLabel(selectionRectOf({ colIndex: 1, rowIndex: 1 }, { colIndex: 1, rowIndex: 1 }))).toBe("B2");
  });

  it("方向键越界时夹在边上，不绕回另一端", () => {
    expect(moveAddress({ colIndex: 0, rowIndex: 0 }, -1, -1, 10, 20)).toEqual({ colIndex: 0, rowIndex: 0 });
    expect(moveAddress({ colIndex: 10, rowIndex: 20 }, 1, 1, 10, 20)).toEqual({ colIndex: 10, rowIndex: 20 });
    expect(moveAddress({ colIndex: 2, rowIndex: 3 }, 0, 1, 10, 20)).toEqual({ colIndex: 2, rowIndex: 4 });
  });
});

describe("点选与列宽拖动（都只认 IO 包给的 px）", () => {
  it("坐标落在哪格就选哪格，超出画布夹到最后一格", () => {
    const layout = buildSheetLayout(sampleSheet());
    // 第一列 89px 宽、第一行 20px 高、第二行 30px 高
    expect(addressAtPoint(layout, 0, 0)).toEqual({ colIndex: 0, rowIndex: 0 });
    expect(addressAtPoint(layout, 100, 25)).toEqual({ colIndex: 1, rowIndex: 1 });
    expect(addressAtPoint(layout, layout.totalWidthPx + 500, layout.totalHeightPx + 900)).toEqual({
      colIndex: layout.colIndexes[layout.colIndexes.length - 1],
      rowIndex: layout.rowIndexes[layout.rowIndexes.length - 1],
    });
  });

  it("拖动中间一列时把 <col> 范围劈成三段，相邻列宽度不受牵连", () => {
    const wide: SheetPayload = { ...sampleSheet(), columns: [{ minColIndex: 0, maxColIndex: 4, widthChars: 12, customWidth: true }] };
    const next = withColumnWidth(wide, 2, 64);
    const ranges = next.columns
      .map((column) => [column.minColIndex, column.maxColIndex, column.widthChars] as const)
      .sort((a, b) => a[0] - b[0]);
    expect(ranges).toEqual([
      [0, 1, 12],
      [2, 2, 2158 / 256], // 64px 换回字符正好是 Excel 默认的 8.43
      [3, 4, 12],
    ]);
  });
});
