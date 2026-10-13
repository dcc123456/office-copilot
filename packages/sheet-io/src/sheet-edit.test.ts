import { describe, expect, it } from "vitest";
import { deleteTracksOp, insertTracksOp, revertSheetOps, runSheetOps, setCellFormulaOp, setCellTextOp, setColumnWidthOp, sortRowsOp, type SortRowsOp, type SortRowsResult } from "./sheet-edit.js";
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
 * 逐列取现场生效的宽度（劈段之后"这列多宽"要看盖住它的那一段）。
 *
 * @param sheet 载荷。
 * @param colIndexes 要看的列号（0 基）。
 * @returns 每列的字符宽；这列没有自定义宽度时给 null。
 */
function widthsAt(sheet: SheetPayload, colIndexes: number[]): Array<number | null> {
  return colIndexes.map(
    (colIndex) => sheet.columns.find((column) => colIndex >= column.minColIndex && colIndex <= column.maxColIndex)?.widthChars ?? null,
  );
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

  it("用户改过这一列的宽度后撤销被拦下：不覆盖用户的列宽", () => {
    const sheet = sheetOf([cellOf(0, 0, 1)], wideRange);
    const applied = runSheetOps(sheet, [setColumnWidthOp(sheet, 1, 103)!]);
    const mangled = { ...applied.after, columns: [{ minColIndex: 1, maxColIndex: 1, widthChars: 20, customWidth: true }] };
    const reverted = revertSheetOps(mangled, applied.inverses);
    expect(reverted.ok).toBe(false);
    expect(reverted.skipped[0]?.reason).toContain("列宽");
    expect(reverted.after.columns).toEqual(mangled.columns);
  });

  it("用户只把分段劈开、宽度没动时撤销照样落回去：他劈出来的段留着", () => {
    const sheet = sheetOf([cellOf(0, 0, 1)], wideRange);
    const applied = runSheetOps(sheet, [setColumnWidthOp(sheet, 1, 103)!]);
    // 用户在撤销之前把 `{2–3}` 单独拆开看（宽度一位没动）：这属于表示法变化，不是数据改动
    const resplit = {
      ...applied.after,
      columns: [
        ...applied.after.columns.slice(0, 2),
        { minColIndex: 2, maxColIndex: 2, widthChars: 12, customWidth: true },
        { minColIndex: 3, maxColIndex: 3, widthChars: 12, customWidth: true },
      ],
    };
    const reverted = revertSheetOps(resplit, applied.inverses);
    expect(reverted.ok, reverted.skipped.map((item) => item.reason).join(" | ")).toBe(true);
    expect(widthsAt(reverted.after, [0, 1, 2, 3])).toEqual([12, 12, 12, 12]);
    expect(reverted.after.columns).toHaveLength(4);
  });

  it("一批里改同一覆盖带的两列宽度：两条都落地（阶段 4 的批内死角）", () => {
    const sheet = sheetOf([cellOf(0, 0, 1), cellOf(1, 0, 2), cellOf(2, 0, 3), cellOf(3, 0, 4)], wideRange);
    // 103px → 14 字符，110px → 15 字符
    const applied = runSheetOps(sheet, [setColumnWidthOp(sheet, 1, 103)!, setColumnWidthOp(sheet, 2, 110)!]);
    expect(applied.ok, applied.skipped.map((item) => item.reason).join(" | ")).toBe(true);
    expect(widthsAt(applied.after, [0, 1, 2, 3])).toEqual([12, 14, 15, 12]);

    const reverted = revertSheetOps(applied.after, applied.inverses);
    expect(reverted.ok, reverted.skipped.map((item) => item.reason).join(" | ")).toBe(true);
    // 精确路径能把 `<cols>` 还原成原来那一段，不必留四段等价写法
    expect(reverted.after.columns).toEqual(wideRange);
  });

  it("同一列在一批里被改两次：第二条以现场宽度对不上被拦下", () => {
    const sheet = sheetOf([cellOf(1, 0, 2)], wideRange);
    const applied = runSheetOps(sheet, [setColumnWidthOp(sheet, 1, 103)!, setColumnWidthOp(sheet, 1, 110)!]);
    expect(applied.ok).toBe(false);
    expect(applied.skipped).toHaveLength(1);
    expect(applied.skipped[0]?.reason).toContain("列宽");
    expect(widthsAt(applied.after, [1])).toEqual([14]);
  });

  it("一批里跨两张覆盖带改列宽：两条都落地，撤销逐条回到各自的前像", () => {
    const columns: SheetColumn[] = [{ minColIndex: 0, maxColIndex: 1, widthChars: 12, customWidth: true }, { minColIndex: 2, maxColIndex: 3, widthChars: 9, customWidth: true }];
    const sheet = sheetOf([cellOf(0, 0, 1), cellOf(1, 0, 2), cellOf(2, 0, 3), cellOf(3, 0, 4)], columns);
    const applied = runSheetOps(sheet, [setColumnWidthOp(sheet, 0, 103)!, setColumnWidthOp(sheet, 3, 110)!]);
    expect(applied.ok, applied.skipped.map((item) => item.reason).join(" | ")).toBe(true);
    expect(widthsAt(applied.after, [0, 1, 2, 3])).toEqual([14, 12, 9, 15]);
    expect(revertSheetOps(applied.after, applied.inverses).after.columns).toEqual(columns);
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

describe("写公式（spec D4：`<f>` 交给 Excel，缓存值 `<v>` 由本地算）", () => {
  it("白名单内的公式当场算出缓存值，撤销后新写的格整格消失", () => {
    const sheet = sheetOf([cellOf(0, 0, 10), cellOf(0, 1, 20), cellOf(0, 2, 30)]);
    const created = setCellFormulaOp(sheet, "D1", "=SUM(A1:A3)");
    expect(created).not.toBeNull();
    expect(created?.recalcNeeded).toBe(false);
    const applied = runSheetOps(sheet, [created!.op]);
    expect(applied.ok).toBe(true);
    expect(applied.after.cells["D1"]).toMatchObject({ kind: "number", value: 60, formula: "SUM(A1:A3)", cachedValue: 60 });

    const reverted = revertSheetOps(applied.after, applied.inverses);
    expect(reverted.ok).toBe(true);
    expect(reverted.after.cells["D1"]).toBeUndefined();
  });

  it("本地算不出的公式回报需要重算，值留空而不是顶上一个过期数字（验收 A4）", () => {
    const sheet = sheetOf([cellOf(0, 0, 10)]);
    const created = setCellFormulaOp(sheet, "B1", "TODAY()")!;
    expect(created.recalcNeeded).toBe(true);
    const applied = runSheetOps(sheet, [created.op]);
    expect(applied.after.cells["B1"]?.formula).toBe("TODAY()");
    expect(applied.after.cells["B1"]?.value).toBeNull();
    expect(applied.after.cells["B1"]?.cachedValue).toBeUndefined();
  });

  it("盖掉旧值时样式与未解释属性跟着走，旧的公式与原文不许留下（§2.3）", () => {
    const base = sheetOf([cellOf(0, 0, 10), cellOf(3, 0, "合计", 3)]);
    const sheet: SheetPayload = {
      ...base,
      cells: {
        ...base.cells,
        A1: { ...base.cells["A1"]!, attrs: { cm: "1" } },
        D1: { ...base.cells["D1"]!, formula: "SUM(A1:A3)", cachedValue: 40, rawXml: '<c r="D1" s="3"/>' },
      },
    };
    const applied = runSheetOps(sheet, [setCellFormulaOp(sheet, "D1", "A1+1")!.op]);
    expect(applied.after.cells["D1"]).toMatchObject({ kind: "number", value: 11, styleIndex: 3, formula: "A1+1", cachedValue: 11 });
    expect(applied.after.cells["D1"]?.rawXml).toBeUndefined();
    // 引用的那一格带着未解释属性，取数只看值，不许把 `attrs` 抄过来
    expect(applied.after.cells["A1"]).toMatchObject({ value: 10, attrs: { cm: "1" } });
  });

  it("自我引用判成循环：公式写进去、值留空，凭 recalcNeeded 提示重算", () => {
    const sheet = sheetOf([cellOf(0, 0, 10)]);
    const created = setCellFormulaOp(sheet, "A1", "A1*2")!;
    expect(created.recalcNeeded).toBe(true);
    expect(created.op.after?.formula).toBe("A1*2");
    expect(created.op.after?.value).toBeNull();
  });

  it("地址认不出或公式为空时返回 null，不产生空改动", () => {
    const sheet = sheetOf([cellOf(0, 0, 1)]);
    expect(setCellFormulaOp(sheet, "A", "=1+1")).toBeNull();
    expect(setCellFormulaOp(sheet, "A1", "   ")).toBeNull();
    expect(setCellFormulaOp(sheet, "A1", "=")).toBeNull();
  });
});

describe("一批 op 改到同一行（阶段 4：整列计算值要一次落地）", () => {
  it("同一行的两格在一批里都落地，撤销后两格一起回到原样", () => {
    const sheet = sheetOf([cellOf(0, 0, 1), cellOf(1, 0, 2)]);
    const applied = runSheetOps(sheet, [setCellFormulaOp(sheet, "C1", "A1*2")!.op, setCellFormulaOp(sheet, "D1", "B1+1")!.op]);
    expect(applied.ok, applied.skipped.map((item) => item.reason).join(" | ")).toBe(true);
    expect(applied.after.cells["C1"]?.value).toBe(2);
    expect(applied.after.cells["D1"]?.value).toBe(3);

    const reverted = revertSheetOps(applied.after, applied.inverses);
    expect(reverted.ok).toBe(true);
    expect(reverted.after.cells).toEqual(sheet.cells);
  });

  it("给整列加计算值：五行各一条 op，一批里全落地", () => {
    const values = [10, 20, 30, 40, 50];
    const sheet = sheetOf(values.map((value, index) => cellOf(0, index, value)));
    const ops = values.map((_, index) => setCellFormulaOp(sheet, `B${index + 1}`, `A${index + 1}^2`)!.op);
    const applied = runSheetOps(sheet, ops);
    expect(applied.ok, applied.skipped.map((item) => item.reason).join(" | ")).toBe(true);
    expect(applied.after.cells["B1"]?.value).toBe(100);
    expect(applied.after.cells["B5"]?.value).toBe(2500);
    expect(revertSheetOps(applied.after, applied.inverses).after.cells["B5"]).toBeUndefined();
  });

  it("同一格在一批里被改两次时第二条仍然拦下，不静默覆盖", () => {
    const sheet = sheetOf([cellOf(0, 0, 1)]);
    const applied = runSheetOps(sheet, [setCellTextOp(sheet, "A1", "2")!, setCellTextOp(sheet, "A1", "3")!]);
    expect(applied.ok).toBe(false);
    expect(applied.skipped[0]?.reason).toContain("现在装的不是");
    expect(applied.after.cells["A1"]?.value).toBe(2);
  });

  it("批内锚点不越过现场核对：落地后用户又改了那一行，撤销仍然认不出", () => {
    const sheet = sheetOf([cellOf(0, 0, 1), cellOf(1, 0, 2)]);
    const applied = runSheetOps(sheet, [setCellTextOp(sheet, "C1", "9")!, setCellTextOp(sheet, "D1", "8")!]);
    expect(applied.ok).toBe(true);
    // 用户把刚写进去的 C1 又改了：这一行的内容指纹不再是撤销所依据的那把
    const userEdited = { ...applied.after, cells: { ...applied.after.cells, C1: { ...applied.after.cells["C1"]!, value: 99 } } };
    const reverted = revertSheetOps(userEdited, applied.inverses);
    expect(reverted.ok).toBe(false);
    expect(reverted.skipped.length).toBeGreaterThan(0);
    expect(reverted.after.cells["C1"]?.value).toBe(99);
  });
});

describe("按一列重排整块行（spec 决策 D7：第五类 op sortRows）", () => {
  /**
   * 取排序 op，拒绝时把理由当测试失败抛出来。
   *
   * @param result `sortRowsOp` 的返回。
   * @returns 那条 op；`ok:false` 时抛 `reason`，让失败信息直接是引擎给的理由。
   */
  function mustOp(result: SortRowsResult): SortRowsOp {
    if (!result.ok) throw new Error(result.reason);
    return result.op;
  }

  /**
   * 造一张三行数据的销售表：表头 + 华东300/华北100/华南200。
   *
   * @param options.merges 合并区（A1 写法）。
   * @param options.extraCells 追加的格子（用于造"区域外还有内容"与带公式的行）。
   * @param options.rows 行元信息。
   * @returns 载荷；`A2`/`B2` 故意带 `styleIndex` 3，好验证格式跟着记录走。
   */
  function salesSheet(options: { merges?: string[]; extraCells?: SheetCell[]; rows?: SheetRowMeta[] } = {}): SheetPayload {
    const cells: SheetCell[] = [
      cellOf(0, 0, "区域"),
      cellOf(1, 0, "销售额"),
      cellOf(0, 1, "华东", 3),
      cellOf(1, 1, 300, 3),
      cellOf(0, 2, "华北"),
      cellOf(1, 2, 100),
      cellOf(0, 3, "华南"),
      cellOf(1, 3, 200),
      ...(options.extraCells ?? []),
    ];
    return sheetOf(cells, [], options.merges ?? [], options.rows ?? []);
  }

  it("按销售额升序把整行搬到新位次，撤销逐格回到原样", () => {
    const sheet = salesSheet();
    const applied = runSheetOps(sheet, [mustOp(sortRowsOp(sheet, { range: "A2:B4", column: "销售额" }))]);
    expect(applied.ok, applied.skipped.map((item) => item.reason).join(" | ")).toBe(true);
    expect(applied.after.cells["A2"]?.value).toBe("华北");
    expect(applied.after.cells["B3"]?.value).toBe(200);
    expect(applied.after.cells["A4"]?.value).toBe("华东");
    // 华东那行原本带 styleIndex 3：格式是记录的一部分，搬走时不许留在原位（§2.3）
    expect(applied.after.cells["B4"]?.styleIndex).toBe(3);
    expect(applied.after.cells["B2"]?.styleIndex).toBeUndefined();

    const reverted = revertSheetOps(applied.after, applied.inverses);
    expect(reverted.ok, reverted.skipped.map((item) => item.reason).join(" | ")).toBe(true);
    expect(reverted.after.cells).toEqual(sheet.cells);
  });

  it("降序排一次、再降序排一次时空操作会被拦下", () => {
    const sheet = salesSheet();
    const applied = runSheetOps(sheet, [mustOp(sortRowsOp(sheet, { range: "A2:B4", column: "销售额", order: "desc" }))]);
    expect(applied.after.cells["A2"]?.value).toBe("华东");
    expect(applied.after.cells["A4"]?.value).toBe("华北");
    const again = sortRowsOp(applied.after, { range: "A2:B4", column: "销售额", order: "desc" });
    expect(!again.ok && again.reason).toContain("已经是降序");
  });

  it("键值相同的行保持原次序：整块搬的是记录，不是把相同值随机打散", () => {
    const sheet = sheetOf([
      cellOf(0, 0, "区域"),
      cellOf(1, 0, "销售额"),
      cellOf(0, 1, "甲"),
      cellOf(1, 1, 1),
      cellOf(0, 2, "乙"),
      cellOf(1, 2, 5),
      cellOf(0, 3, "丙"),
      cellOf(1, 3, 1),
    ]);
    const applied = runSheetOps(sheet, [mustOp(sortRowsOp(sheet, { range: "A2:B4", column: "销售额" }))]);
    expect(applied.ok, applied.skipped.map((item) => item.reason).join(" | ")).toBe(true);
    expect(applied.after.cells["A2"]?.value).toBe("甲");
    expect(applied.after.cells["A3"]?.value).toBe("丙");
    expect(applied.after.cells["A4"]?.value).toBe("乙");
  });

  it("键列空白的行升序降序都排在带尾", () => {
    const sheet = sheetOf([
      cellOf(0, 0, "区域"),
      cellOf(1, 0, "销售额"),
      cellOf(0, 1, "五"),
      cellOf(1, 1, 5),
      cellOf(0, 2, "空"),
      cellOf(0, 3, "一"),
      cellOf(1, 3, 1),
    ]);
    const ascending = runSheetOps(sheet, [mustOp(sortRowsOp(sheet, { range: "A2:B4", column: "销售额" }))]);
    expect(ascending.after.cells["A4"]?.value).toBe("空");
    expect(ascending.after.cells["A2"]?.value).toBe("一");
    const descending = runSheetOps(sheet, [mustOp(sortRowsOp(sheet, { range: "A2:B4", column: "销售额", order: "desc" }))]);
    expect(descending.after.cells["A4"]?.value).toBe("空");
    expect(descending.after.cells["A2"]?.value).toBe("五");
  });

  it("行高与单行合并区跟着记录一起挪，不留在原位", () => {
    const sheet = salesSheet({ merges: ["A2:B2"], rows: [{ rowIndex: 1, heightPt: 30, customHeight: true }] });
    const applied = runSheetOps(sheet, [mustOp(sortRowsOp(sheet, { range: "A2:B4", column: "销售额" }))]);
    expect(applied.ok, applied.skipped.map((item) => item.reason).join(" | ")).toBe(true);
    expect(applied.after.merges.map((merge) => merge.ref)).toEqual(["A4:B4"]);
    expect(applied.after.rows.find((row) => row.rowIndex === 3)?.heightPt).toBe(30);
    expect(applied.after.rows.find((row) => row.rowIndex === 1)).toBeUndefined();
    const reverted = revertSheetOps(applied.after, applied.inverses);
    expect(reverted.ok, reverted.skipped.map((item) => item.reason).join(" | ")).toBe(true);
    expect(reverted.after.merges.map((merge) => merge.ref)).toEqual(["A2:B2"]);
    expect(reverted.after.rows).toEqual(sheet.rows);
  });

  it("带内有公式、区域外还有格子、合并区劈不开时三条都拒绝", () => {
    const withFormula = salesSheet({ extraCells: [formulaCell("C2", 2, 1, "B2/1000")] });
    const formulaRefusal = sortRowsOp(withFormula, { range: "A2:C4", column: "销售额" });
    expect(formulaRefusal.ok).toBe(false);
    if (!formulaRefusal.ok) expect(formulaRefusal.reason).toContain("带公式");

    const outside = salesSheet({ extraCells: [cellOf(2, 1, 0.3)] });
    const outsideRefusal = sortRowsOp(outside, { range: "A2:B4", column: "销售额" });
    expect(outsideRefusal.ok).toBe(false);
    if (!outsideRefusal.ok) expect(outsideRefusal.reason).toContain("在排序区域之外还有内容");

    const acrossRows = sortRowsOp(salesSheet({ merges: ["A2:A3"] }), { range: "A2:B4", column: "销售额" });
    expect(acrossRows.ok).toBe(false);
    if (!acrossRows.ok) expect(acrossRows.reason).toContain("合并区跨行或探出排序区域");

    const stickingOut = sortRowsOp(salesSheet({ merges: ["A2:C2"] }), { range: "A2:B4", column: "销售额" });
    expect(stickingOut.ok).toBe(false);
    if (!stickingOut.ok) expect(stickingOut.reason).toContain("合并区跨行或探出排序区域");
  });

  it("区域与键列写错时逐条给理由，不猜位置", () => {
    const sheet = salesSheet();
    const refusals = [
      sortRowsOp(sheet, { range: "A2:B2", column: "销售额" }),
      sortRowsOp(sheet, { range: "A1:B99", column: "销售额" }),
      sortRowsOp(sheet, { range: "A2:B4", column: "不存在的列" }),
      sortRowsOp(sheet, { range: "A2:A4", column: "销售额" }),
      sortRowsOp(sheet, { range: "A:C", column: "销售额" }),
    ];
    expect(refusals.map((refusal) => refusal.ok)).toEqual([false, false, false, false, false]);
    expect(refusals[0]).toEqual({ ok: false, reason: expect.stringContaining("不足两行") });
    expect((refusals[1] as { reason: string }).reason).toContain("超出了这张表的数据区");
    expect((refusals[2] as { reason: string }).reason).toContain("认不出键列");
    expect((refusals[3] as { reason: string }).reason).toContain("不在排序区域");
  });

  it("排完用户又改了带里一格：撤销报部分撤销，不许按位置硬盖", () => {
    const sheet = salesSheet();
    const applied = runSheetOps(sheet, [mustOp(sortRowsOp(sheet, { range: "A2:B4", column: "销售额" }))]);
    // 用户在排好的表上把中间那一行（华南）的销售额改成 999：现场不再是撤销所依据的那份镜像
    const userEdited = { ...applied.after, cells: { ...applied.after.cells, B3: { ...applied.after.cells["B3"]!, value: 999 } } };
    const reverted = revertSheetOps(userEdited, applied.inverses);
    expect(reverted.ok).toBe(false);
    expect(reverted.skipped[0]?.reason).toContain("B3");
    expect(reverted.after.cells["B3"]?.value).toBe(999);

    // 改的是锚点那一行（新的第一行）时内容寻址直接认不出，同样一格都不许动
    const anchorEdited = { ...applied.after, cells: { ...applied.after.cells, B2: { ...applied.after.cells["B2"]!, value: 999 } } };
    const anchorRevert = revertSheetOps(anchorEdited, applied.inverses);
    expect(anchorRevert.ok).toBe(false);
    expect(anchorRevert.skipped[0]?.reason).toContain("认不出要重排的第一行");
    expect(anchorRevert.after.cells["B2"]?.value).toBe(999);
  });

  it("排序与改值同批时，改值按内容寻址落到那条记录的新行上", () => {
    const sheet = salesSheet();
    const applied = runSheetOps(sheet, [
      mustOp(sortRowsOp(sheet, { range: "A2:B4", column: "销售额" })),
      // 这条 op 造在排序前的表上：它认的是"华南/200 那一行"，排序把它搬到了第 3 行
      setCellTextOp(sheet, "B4", "999")!,
    ]);
    expect(applied.ok, applied.skipped.map((item) => item.reason).join(" | ")).toBe(true);
    expect(applied.after.cells["A3"]?.value).toBe("华南");
    expect(applied.after.cells["B3"]?.value).toBe(999);
    expect(applied.after.cells["B4"]?.value).toBe(300);
  });

  it("升序排到一半发现超过一次能安全记账的格数就整条拒", () => {
    const bigCells: SheetCell[] = [];
    for (let rowIndex = 0; rowIndex < 16_668; rowIndex += 1) {
      for (let colIndex = 0; colIndex < 3; colIndex += 1) bigCells.push(cellOf(colIndex, rowIndex, rowIndex * 3 + colIndex));
    }
    const refusal = sortRowsOp(sheetOf(bigCells), { range: "A1:C16668", column: "C" });
    expect(refusal.ok).toBe(false);
    if (!refusal.ok) expect(refusal.reason).toContain("50000");
  });
});
