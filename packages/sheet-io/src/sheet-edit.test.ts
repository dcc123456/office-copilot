import { describe, expect, it } from "vitest";
import { deleteTracksOp, insertTracksOp, revertSheetOps, runSheetOps, setCellTextOp, setColumnWidthOp } from "./sheet-edit.js";
import { parseRange } from "./address.js";
import type { SheetCell, SheetColumn, SheetLike as SheetPayload, SheetRowMeta } from "./model.js";

/** 表的部件路径；op 与载荷必须对得上才落地。 */
const PART = "xl/worksheets/sheet1.xml";

/**
 * 造一格。
 *
 * @param colIndex 列号（0 基）。
 * @param rowIndex 行号（0 基）。
 * @param value 格里的值。
 * @param styleIndex 原 `@s`；带上它就是一格"有格式"的格，撤销不许把格式拍平。
 * @returns 一个最小可用的格子。
 */
function cellOf(colIndex: number, rowIndex: number, value: number | string, styleIndex?: number): SheetCell {
  const letters = String.fromCharCode(65 + colIndex);
  return {
    address: `${letters}${rowIndex + 1}`,
    colIndex,
    rowIndex,
    kind: typeof value === "number" ? "number" : "string",
    value,
    ...(styleIndex === undefined ? {} : { styleIndex }),
  };
}

/**
 * 造一张表（载荷形状，不含 XML 原文）。
 *
 * @param cells 已用的格。
 * @param columns `<col>` 分段。
 * @param mergeRefs 合并区的 A1 写法，如 `A1:B2`。
 * @param rows 行元信息。
 * @returns `SheetPayload`；已用区域由格子的实际位置推出。
 */
function sheetOf(cells: SheetCell[], columns: SheetColumn[] = [], mergeRefs: string[] = [], rows: SheetRowMeta[] = []): SheetPayload {
  const lastColIndex = cells.reduce((acc, cell) => Math.max(acc, cell.colIndex), -1);
  const lastRowIndex = cells.reduce((acc, cell) => Math.max(acc, cell.rowIndex), -1);
  return {
    name: "销售",
    sheetId: 1,
    partPath: PART,
    lastColIndex,
    lastRowIndex,
    cells: Object.fromEntries(cells.map((cell) => [cell.address, cell])),
    merges: mergeRefs.map((ref) => ({ ref, range: parseRange(ref)! })),
    columns,
    rows,
    warnings: [],
  };
}

/**
 * 造一个带公式的格子（插删行列时要验证引用跟着走）。
 *
 * @param address 大写 A1 地址。
 * @param colIndex 列号（0 基）。
 * @param rowIndex 行号（0 基）。
 * @param formula 公式体，不含前导 `=`。
 * @returns 一个数字格；缓存值给 0，保存时会由 `fullCalcOnLoad` 让 Excel 重算。
 */
function formulaCell(address: string, colIndex: number, rowIndex: number, formula: string): SheetCell {
  return { address, colIndex, rowIndex, kind: "number", value: 0, formula };
}

describe("改一格的值（spec 阶段 3：op 自带逆操作）", () => {
  it("落地后撤销，格子与已用区域都回到原样", () => {
    const sheet = sheetOf([cellOf(0, 0, 1), cellOf(2, 1, 300)]);
    const op = setCellTextOp(sheet, "C2", "999");
    expect(op).not.toBeNull();
    const applied = runSheetOps(sheet, [op]);
    expect(applied.ok).toBe(true);
    expect(applied.after.cells["C2"]?.value).toBe(999);

    const reverted = revertSheetOps(applied.after, applied.inverses);
    expect(reverted.ok).toBe(true);
    expect(reverted.after.cells).toEqual(sheet.cells);
    expect(reverted.after.lastRowIndex).toBe(sheet.lastRowIndex);
  });

  it("新写的格被撤销后整格消失，已用区域缩回原大小", () => {
    const sheet = sheetOf([cellOf(0, 0, 1)]);
    const applied = runSheetOps(sheet, [setCellTextOp(sheet, "E9", "77")!]);
    expect(applied.after.cells["E9"]).toBeDefined();
    expect(applied.after.lastRowIndex).toBe(8);

    const reverted = revertSheetOps(applied.after, applied.inverses);
    expect(reverted.after.cells["E9"]).toBeUndefined();
    expect(reverted.after.lastRowIndex).toBe(0);
  });

  it("清空一格时，带格式的格退化成只有格式，无格式的格整格删掉", () => {
    const sheet = sheetOf([cellOf(0, 0, 5, 3), cellOf(1, 0, 6)]);
    const styled = runSheetOps(sheet, [setCellTextOp(sheet, "A1", "")!]);
    expect(styled.after.cells["A1"]).toMatchObject({ kind: "blank", styleIndex: 3, styleOnly: true });

    const plain = runSheetOps(sheet, [setCellTextOp(sheet, "B1", "")!]);
    expect(plain.after.cells["B1"]).toBeUndefined();
  });

  it("改数字与日期时样式索引原样带走，旧的公式与原文不许留下", () => {
    const sheet = sheetOf([cellOf(1, 0, 120, 1)]);
    const withHistory = {
      ...sheet,
      cells: { ...sheet.cells, B1: { ...sheet.cells["B1"]!, formula: "SUM(A1:A2)", cachedValue: 120, rawXml: '<c r="B1"/>' } },
    };
    // 日期串解析成序列号 + 数字格式（换算只在 cell-text/geometry 一处发生，§2.1）
    const applied = runSheetOps(withHistory, [setCellTextOp(withHistory, "B1", "2026-02-01")!]);
    expect(applied.after.cells["B1"]).toMatchObject({ kind: "number", value: 46054, styleIndex: 1 });
    expect(applied.after.cells["B1"]?.numberFormat).toMatchObject({ numFmtId: 14, isDateLike: true });
    expect(applied.after.cells["B1"]?.formula).toBeUndefined();
    expect(applied.after.cells["B1"]?.rawXml).toBeUndefined();
  });

  it("往表尾外写一格时已用区域放大，撤销后缩回去", () => {
    const sheet = sheetOf([cellOf(0, 0, 1)]);
    const applied = runSheetOps(sheet, [setCellTextOp(sheet, "D9", "2026-02-01")!]);
    expect(applied.after.cells["D9"]).toMatchObject({ colIndex: 3, rowIndex: 8, value: 46054 });
    expect(applied.after.lastColIndex).toBe(3);
    expect(applied.after.lastRowIndex).toBe(8);
    expect(revertSheetOps(applied.after, applied.inverses).after.lastRowIndex).toBe(0);
  });

  it("现场核对只看内容：格的样式被用户改过时撤销被拦下，不覆盖新格式", () => {
    const sheet = sheetOf([cellOf(0, 0, 1), cellOf(1, 0, 2)]);
    const applied = runSheetOps(sheet, [setCellTextOp(sheet, "A1", "2")!]);
    expect(applied.ok).toBe(true);
    // 撤销前用户把 A1 刷成了粗体（值没动，所以行指纹仍认得出这一行）：这时必须拦住，
    // 否则撤销会把用户的加粗连同旧值一起抹掉（§2.3）。
    const restyled = { ...applied.after, cells: { ...applied.after.cells, A1: { ...applied.after.cells["A1"]!, styleIndex: 5 } } };
    const reverted = revertSheetOps(restyled, applied.inverses);
    expect(reverted.ok).toBe(false);
    expect(reverted.skipped).toHaveLength(1);
    expect(reverted.skipped[0]?.reason).toContain("A1");
    expect(reverted.after.cells["A1"]?.styleIndex).toBe(5);
  });

  it("整行内容被用户改过时撤销认不出落点，宁可报部分撤销", () => {
    const sheet = sheetOf([cellOf(0, 0, 1)]);
    const applied = runSheetOps(sheet, [setCellTextOp(sheet, "A1", "2")!]);
    const editedByUser = { ...applied.after, cells: { ...applied.after.cells, A1: cellOf(0, 0, 50) } };
    const reverted = revertSheetOps(editedByUser, applied.inverses);
    expect(reverted.ok).toBe(false);
    expect(reverted.skipped[0]?.reason).toContain("认不出这一行");
    expect(reverted.after.cells["A1"]?.value).toBe(50);
  });

  it("插入行把内容整体下移后，撤销仍命中原来那一行", () => {
    const sheet = sheetOf([cellOf(0, 0, 7), cellOf(0, 1, 300)]);
    const op = setCellTextOp(sheet, "A2", "660")!;
    const applied = runSheetOps(sheet, [op]);
    // 模拟"用户在表头之上插了一行"：所有格下移一位，值不动（地址变了，内容没变）
    const shiftedCells = Object.values(applied.after.cells).map((cell) => cellOf(cell.colIndex, cell.rowIndex + 1, cell.value as number));
    const shifted = sheetOf(shiftedCells);
    const reverted = revertSheetOps(shifted, applied.inverses);
    expect(reverted.ok).toBe(true);
    expect(reverted.after.cells["A3"]?.value).toBe(300);
    expect(reverted.after.cells["A2"]?.value).toBe(7);
  });

  it("op 指向别的表就整条跳过，不许静默写进当前表", () => {
    const sheet = sheetOf([cellOf(0, 0, 1)]);
    const op = setCellTextOp(sheet, "A1", "9")!;
    const other = { ...sheet, partPath: "xl/worksheets/sheet2.xml" };
    const applied = runSheetOps(other, [op]);
    expect(applied.ok).toBe(false);
    expect(applied.skipped[0]?.reason).toContain("另一张表");
  });
});

describe("改列宽（`<col min max>` 劈段与复原）", () => {
  const wideRange: SheetColumn[] = [{ minColIndex: 0, maxColIndex: 3, widthChars: 12, customWidth: true }];

  it("只改中间一列时把范围劈成三段，相邻列的宽度不受牵连", () => {
    const sheet = sheetOf([cellOf(0, 0, 1)], wideRange);
    // 103px 是 14 个字符宽（`geometry.ts`：px = chars × 7 + 5）
    const applied = runSheetOps(sheet, [setColumnWidthOp(sheet, 1, 103)!]);
    expect(applied.after.columns).toHaveLength(3);
    // 分段顺序也要按列号排好，Excel 自己写出的 `<cols>` 就是这个顺序
    expect(applied.after.columns.map((column) => column.minColIndex)).toEqual([0, 1, 2]);
    expect(applied.after.columns).toContainEqual({ minColIndex: 0, maxColIndex: 0, widthChars: 12, customWidth: true });
    expect(applied.after.columns).toContainEqual({ minColIndex: 1, maxColIndex: 1, widthChars: 14, customWidth: true });
    expect(applied.after.columns).toContainEqual({ minColIndex: 2, maxColIndex: 3, widthChars: 12, customWidth: true });

    const reverted = revertSheetOps(applied.after, applied.inverses);
    expect(reverted.ok).toBe(true);
    expect(reverted.after.columns).toEqual(wideRange);
  });

  it("这一列本来就这么宽时不产生 op", () => {
    const sheet = sheetOf([cellOf(0, 0, 1)], [{ minColIndex: 1, maxColIndex: 1, widthChars: 14, customWidth: true }]);
    expect(setColumnWidthOp(sheet, 1, 103)).toBeNull();
  });

  it("用户重排过分段后撤销被拦下：不重排已经不属于我们的列宽", () => {
    const sheet = sheetOf([cellOf(0, 0, 1)], wideRange);
    const applied = runSheetOps(sheet, [setColumnWidthOp(sheet, 1, 103)!]);
    const mangled = { ...applied.after, columns: [{ minColIndex: 1, maxColIndex: 1, widthChars: 20, customWidth: true }] };
    const reverted = revertSheetOps(mangled, applied.inverses);
    expect(reverted.ok).toBe(false);
    expect(reverted.skipped[0]?.reason).toContain("列宽");
    expect(reverted.after.columns).toEqual(mangled.columns);
  });

  it("没有 `<col>` 时也能建出这一列的宽度段", () => {
    const sheet = sheetOf([cellOf(0, 0, 1)]);
    const op = setColumnWidthOp(sheet, 2, 103);
    expect(op?.afterColumns).toEqual([{ minColIndex: 2, maxColIndex: 2, widthChars: 14, customWidth: true }]);
    expect(op?.beforeColumns).toEqual([]);
    const applied = runSheetOps(sheet, [op!]);
    expect(applied.after.columns).toEqual([{ minColIndex: 2, maxColIndex: 2, widthChars: 14, customWidth: true }]);
  });
});

describe("插入与删除整行/整列（spec 阶段 3 C16）", () => {
  it("插入一行把下方内容整体下移，公式引用与已用区域一起跟上", () => {
    const sheet = sheetOf([cellOf(0, 0, 1), cellOf(0, 1, 2), formulaCell("B3", 1, 2, "SUM(A1:A2)")]);
    const applied = runSheetOps(sheet, [insertTracksOp(sheet, "row", 0)!]);
    expect(applied.ok).toBe(true);
    expect(Object.keys(applied.after.cells).sort()).toEqual(["A2", "A3", "B4"]);
    expect(applied.after.cells["B4"]?.formula).toBe("SUM(A2:A3)");
    expect(applied.after.lastRowIndex).toBe(3);

    const reverted = revertSheetOps(applied.after, applied.inverses);
    expect(reverted.ok).toBe(true);
    expect(reverted.after.cells).toEqual(sheet.cells);
    expect(reverted.after.lastRowIndex).toBe(2);
  });

  it("删掉一行后撤销，整带内容与行高原样插回原位", () => {
    const sheet = sheetOf([cellOf(0, 0, 1), cellOf(0, 1, 2), cellOf(0, 2, 3)], [], [], [
      { rowIndex: 1, heightPt: 30, customHeight: true, attrs: { "x14ac:dyDescent": "0.25" } },
    ]);
    const applied = runSheetOps(sheet, [deleteTracksOp(sheet, "row", 1)!]);
    expect(applied.ok).toBe(true);
    expect(Object.keys(applied.after.cells).sort()).toEqual(["A1", "A2"]);
    expect(applied.after.rows).toEqual([]);

    const reverted = revertSheetOps(applied.after, applied.inverses);
    expect(reverted.ok).toBe(true);
    expect(reverted.after.cells).toEqual(sheet.cells);
    expect(reverted.after.rows).toEqual(sheet.rows);
  });

  it("公式引用的格子正好在被删的区间里时整条拒绝，不产 #REF!", () => {
    const sheet = sheetOf([cellOf(0, 0, 1), cellOf(0, 1, 2), cellOf(0, 2, 3), formulaCell("B1", 1, 0, "SUM(A2:A3)")]);
    const applied = runSheetOps(sheet, [deleteTracksOp(sheet, "row", 1, 2)!]);
    expect(applied.ok).toBe(false);
    expect(applied.skipped[0]?.reason).toContain("#REF!");
    expect(applied.after).toBe(sheet);
  });

  it("撤销插入时那一带已被用户写了数据就拦住，不许连带删掉它", () => {
    const sheet = sheetOf([cellOf(0, 0, 1), cellOf(0, 1, 2)]);
    const applied = runSheetOps(sheet, [insertTracksOp(sheet, "row", 1)!]);
    expect(Object.keys(applied.after.cells).sort()).toEqual(["A1", "A3"]);

    const withUserInput = { ...applied.after, cells: { ...applied.after.cells, A2: cellOf(0, 1, 99) } };
    const reverted = revertSheetOps(withUserInput, applied.inverses);
    expect(reverted.ok).toBe(false);
    expect(reverted.skipped[0]?.reason).toContain("格子数与记录时不同");
    expect(reverted.after.cells["A2"]?.value).toBe(99);
  });

  it("插入一列时被跨过的列宽段与合并区变宽，撤销后收回去", () => {
    const sheet = sheetOf([cellOf(0, 0, 1), cellOf(1, 1, 2)], [{ minColIndex: 0, maxColIndex: 2, widthChars: 12, customWidth: true }], ["A1:B2"]);
    const applied = runSheetOps(sheet, [insertTracksOp(sheet, "col", 0)!]);
    expect(applied.ok).toBe(true);
    expect(applied.after.columns).toEqual([{ minColIndex: 1, maxColIndex: 3, widthChars: 12, customWidth: true }]);
    expect(applied.after.merges.map((merge) => merge.ref)).toEqual(["B1:C2"]);
    expect(Object.keys(applied.after.cells).sort()).toEqual(["B1", "C2"]);

    const reverted = revertSheetOps(applied.after, applied.inverses);
    expect(reverted.ok).toBe(true);
    expect(reverted.after.columns).toEqual(sheet.columns);
    expect(reverted.after.merges).toEqual(sheet.merges);
  });

  it("盖住删除带的列宽段只增减端点，来回都是同一段", () => {
    const sheet = sheetOf([cellOf(0, 0, 1), cellOf(1, 0, 2), cellOf(2, 0, 3)], [{ minColIndex: 1, maxColIndex: 3, widthChars: 12, customWidth: true }]);
    const applied = runSheetOps(sheet, [deleteTracksOp(sheet, "col", 2)!]);
    expect(applied.after.columns).toEqual([{ minColIndex: 1, maxColIndex: 2, widthChars: 12, customWidth: true }]);

    const reverted = revertSheetOps(applied.after, applied.inverses);
    expect(reverted.ok).toBe(true);
    expect(reverted.after.columns).toEqual(sheet.columns);
  });

  it("劈开列宽段时带内那截存进前像，撤销后写出相邻同宽的两段", () => {
    const sheet = sheetOf([cellOf(0, 0, 1), cellOf(1, 0, 2), cellOf(2, 0, 3)], [{ minColIndex: 1, maxColIndex: 2, widthChars: 12, customWidth: true }]);
    const applied = runSheetOps(sheet, [deleteTracksOp(sheet, "col", 2)!]);
    expect(applied.after.columns).toEqual([{ minColIndex: 1, maxColIndex: 1, widthChars: 12, customWidth: true }]);

    const reverted = revertSheetOps(applied.after, applied.inverses);
    expect(reverted.ok).toBe(true);
    // 宽度与属性一列都没丢，只是原来的 `<col min=2 max=3>` 写成两段相邻同宽的段
    expect(reverted.after.columns).toEqual([
      { minColIndex: 1, maxColIndex: 1, widthChars: 12, customWidth: true },
      { minColIndex: 2, maxColIndex: 2, widthChars: 12, customWidth: true },
    ]);
  });

  it("删除带把合并区劈开时整条拒绝：合不回去的改动不做", () => {
    const sheet = sheetOf([cellOf(0, 0, 1), cellOf(1, 0, 2)], [], ["A1:B1"]);
    const applied = runSheetOps(sheet, [deleteTracksOp(sheet, "col", 1)!]);
    expect(applied.ok).toBe(false);
    expect(applied.skipped[0]?.reason).toContain("A1:B1");
    expect(applied.after.merges).toEqual(sheet.merges);
  });

  it("整块落在删除带里的合并区随带删掉，撤销时原样回来", () => {
    const sheet = sheetOf([cellOf(0, 0, 1), cellOf(1, 0, 2), cellOf(4, 0, 5)], [], ["B1:C1"]);
    const applied = runSheetOps(sheet, [deleteTracksOp(sheet, "col", 1, 2)!]);
    expect(applied.ok).toBe(true);
    expect(applied.after.merges).toEqual([]);
    expect(Object.keys(applied.after.cells).sort()).toEqual(["A1", "C1"]);

    const reverted = revertSheetOps(applied.after, applied.inverses);
    expect(reverted.ok).toBe(true);
    expect(reverted.after.merges).toEqual(sheet.merges);
    expect(reverted.after.cells).toEqual(sheet.cells);
  });

  it("列轴插删之后每行的 spans 跟着改，行轴不动它", () => {
    const sheet = sheetOf([cellOf(0, 0, 1), cellOf(2, 0, 3)], [], [], [{ rowIndex: 0, attrs: { spans: "1:3" } }]);
    const widened = runSheetOps(sheet, [insertTracksOp(sheet, "col", 1)!]);
    expect(widened.after.rows[0]?.attrs?.spans).toBe("1:4");
    const taller = runSheetOps(sheet, [insertTracksOp(sheet, "row", 1)!]);
    expect(taller.after.rows[0]?.attrs?.spans).toBe("1:3");
  });
});
