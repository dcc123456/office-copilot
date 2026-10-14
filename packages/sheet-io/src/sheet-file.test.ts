import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { buildSampleWorkbookBytes } from "../scripts/make-sample-sheet.js";
import { cellDisplayText, parseUserCellText } from "./cell-text.js";
import { SheetFile, SheetOpenError } from "./sheet-file.js";
import { applySheetPayload, sheetPayloadOf, workbookPayload } from "./sheet-dto.js";
import { deleteTracksOp, insertTracksOp, revertSheetOps, runSheetOps, setCellFormulaOp, setCellTextOp, setColumnWidthOp, sortRowsOp } from "./sheet-edit.js";
import { describeSheet, findRows, readRange } from "./analysis/sheet-table.js";
import type { SheetModel, SheetLike as SheetPayload } from "./model.js";

const SAMPLE_PART_PATHS = [
  "[Content_Types].xml",
  "_rels/.rels",
  "docProps/core.xml",
  "docProps/app.xml",
  "xl/workbook.xml",
  "xl/_rels/workbook.xml.rels",
  "xl/theme/theme1.xml",
  "xl/styles.xml",
  "xl/sharedStrings.xml",
  "xl/worksheets/sheet1.xml",
  "xl/worksheets/sheet2.xml",
  "xl/externalLinks/externalLink1.xml",
  "xl/externalLinks/_rels/externalLink1.xml.rels",
];

async function partTexts(bytes: Uint8Array): Promise<Record<string, string>> {
  const zip = await JSZip.loadAsync(bytes);
  const acc: Record<string, string> = {};
  for (const path of SAMPLE_PART_PATHS) {
    const entry = zip.file(path);
    acc[path] = entry ? await entry.async("string") : "";
  }
  return acc;
}

/** 取 `cellXfs` 里的 xf 标签序列，用于断言"只追加不重排"。 */
function xfsOf(stylesXml: string): string[] {
  const body = stylesXml.slice(stylesXml.indexOf("<cellXfs"), stylesXml.indexOf("</cellXfs>"));
  return body.match(/<xf[^>]*\/>/g) ?? [];
}

async function openSample(): Promise<SheetFile> {
  return SheetFile.fromBytes(await buildSampleWorkbookBytes(), { fileName: "sample.xlsx" });
}

/**
 * 走一遍桌面的保存路径：重新打开夹具、按 partPath 配对写回载荷、存盘再读回来。
 *
 * @param payload op 落地后的表载荷（不含 XML 原文，渲染层能碰到的就是它）。
 * @param sheetIndex 写第几张表，默认第一张。
 * @returns 读回来的表模型，外加它那份部件文本（断言未建模结构没被牵连时用）。
 */
async function writeAndReopen(payload: SheetPayload, sheetIndex = 0): Promise<{ sheet: SheetModel; partXml: string }> {
  const file = await openSample();
  const model = (await file.sheetAt(sheetIndex)) as SheetModel;
  applySheetPayload(model, payload);
  file.writeSheet(model);
  const bytes = await file.toBytes();
  const reopened = await SheetFile.fromBytes(bytes, { fileName: "sample.xlsx" });
  return {
    sheet: (await reopened.sheetAt(sheetIndex)) as SheetModel,
    partXml: (await partTexts(bytes))[`xl/worksheets/sheet${sheetIndex + 1}.xml`] ?? "",
  };
}

/** 列宽分段的 `[起始列, 结束列, 字符宽]` 三元组：断言时比整对象好读，也不会因 `attrs` 键序假红。 */
function widthSegments(sheet: SheetModel): Array<[number, number, number | undefined]> {
  return sheet.columns.map((column) => [column.minColIndex, column.maxColIndex, column.widthChars] as [number, number, number | undefined]);
}

describe("打开真实形态的 .xlsx（spec 阶段 1 / 验收 A2–A4）", () => {
  it("读出两张表，字符串/数字/布尔/日期各归其位", async () => {
    const file = await openSample();
    expect(file.sheetCount).toBe(2);
    expect(file.sheetRefAt(0)?.name).toBe("销售");
    const sheet = await file.sheetAt(0);
    if (!sheet) throw new Error("没有读到第一张表");

    expect(sheet.cells["A1"]?.value).toBe("月份");
    expect(sheet.cells["C2"]).toMatchObject({ kind: "number", value: 120, styleIndex: 1 });
    expect(sheet.cells["E2"]?.numberFormat).toMatchObject({ numFmtId: 10, isDateLike: false });
  });

  it("日期只由样式决定：自定义 164 与内建 14 都判成日期，值仍是序列号", async () => {
    const sheet = (await (await openSample()).sheetAt(0)) as SheetModel;
    expect(sheet.cells["D2"]).toMatchObject({ kind: "number", value: 46053 });
    expect(sheet.cells["D2"]?.numberFormat).toMatchObject({ numFmtId: 164, formatCode: "yyyy-mm-dd", isDateLike: true });
    expect(sheet.cells["D3"]?.numberFormat).toMatchObject({ numFmtId: 14, isDateLike: true });
  });

  it("公式与缓存值一起读回来（spec D4 的读侧）", async () => {
    const sheet = (await (await openSample()).sheetAt(0)) as SheetModel;
    expect(sheet.cells["G2"]).toMatchObject({ formula: "SUM(C2:C4)", value: 360, cachedValue: 360, kind: "number" });
  });

  it("富文本 run 保留但纯文本无损；只有格式的格不丢（实测 §二 P2）", async () => {
    const sheet = (await (await openSample()).sheetAt(0)) as SheetModel;
    expect(sheet.cells["A2"]?.value).toBe("粗蓝");
    expect(sheet.cells["A2"]?.runs).toEqual([
      { text: "粗", bold: true },
      { text: "蓝", color: { hex: "#0000ff" } },
    ]);
    // 只有格式没有值的格（实测 `<c r="H2"/>` / `<c r="E3" s="3"/>`）必须留在模型里
    expect(sheet.cells["E3"]).toMatchObject({ kind: "blank", value: null, styleIndex: 3, styleOnly: true });
    expect(sheet.cells["I1"]).toMatchObject({ kind: "blank", value: null });
    expect(sheet.cells["H2"]?.kind).toBe("blank");
  });

  it("合并、冻结、以及未建模结构都如实报出来", async () => {
    const sheet = (await (await openSample()).sheetAt(0)) as SheetModel;
    expect(sheet.merges.map((m) => m.ref)).toEqual(["H1:I2"]);
    expect(sheet.freeze).toEqual({ colCount: 1, rowCount: 1 });
    expect(sheet.warnings).toContain("unmodelled:conditionalFormatting");
    expect(sheet.warnings).toContain("unmodelled:dataValidations");
    expect(sheet.warnings).toContain("unmodelled:pageMargins");
    expect(sheet.lastColIndex).toBe(9);
    expect(sheet.lastRowIndex).toBe(3);
  });

  it("列宽按字符数原样存，不做第二处换算", async () => {
    const sheet = (await (await openSample()).sheetAt(0)) as SheetModel;
    expect(sheet.columns[0]).toMatchObject({ minColIndex: 0, maxColIndex: 2, widthChars: 12, customWidth: true });
  });
});

describe("回写保真（验收 A1：空改动不许动任何部件）", () => {
  it("打开后不改任何东西，写出的每个部件与原文件一致", async () => {
    const original = await buildSampleWorkbookBytes();
    const file = await SheetFile.fromBytes(original, { fileName: "sample.xlsx" });
    const written = await file.toBytes();
    expect(await partTexts(written)).toEqual(await partTexts(original));
  });

  it("像桌面那样把每张表都写一遍：富文本格仍指原条目，颜色一个字都不许丢", async () => {
    const original = await buildSampleWorkbookBytes();
    const file = await SheetFile.fromBytes(original, { fileName: "sample.xlsx" });
    for (let index = 0; index < file.sheetCount; index += 1) {
      const sheet = await file.sheetAt(index);
      if (sheet) file.writeSheet(sheet);
    }
    const parts = await partTexts(await file.toBytes());
    // A2 是夹具里唯一的富文本格：换索引就是换条目，格式只剩能重新序列化的那部分，所以 `<v>` 必须还是 6
    expect(parts["xl/worksheets/sheet1.xml"]).toContain('<c r="A2" t="s"><v>6</v></c>');
    expect(parts["xl/sharedStrings.xml"]).toContain('<si><r><rPr><b/></rPr><t>粗</t></r><r><rPr><color rgb="FF0000FF"/></rPr><t>蓝</t></r></si>');
    expect(parts["xl/sharedStrings.xml"]).toContain('uniqueCount="11"');
    expect((parts["xl/sharedStrings.xml"].match(/<si>/g) ?? []).length).toBe(11);
  });

  it("桌面保存走的是载荷那条路，载荷里没有 sharedStringIndex——仍必须复用原富文本条目", async () => {
    const file = await openSample();
    for (let index = 0; index < file.sheetCount; index += 1) {
      const model = (await file.sheetAt(index)) as SheetModel;
      applySheetPayload(model, sheetPayloadOf(model));
      file.writeSheet(model);
    }
    const parts = await partTexts(await file.toBytes());
    expect(parts["xl/worksheets/sheet1.xml"]).toContain('<c r="A2" t="s"><v>6</v></c>');
    expect(parts["xl/sharedStrings.xml"]).toContain('uniqueCount="11"');
    expect((parts["xl/sharedStrings.xml"].match(/<si>/g) ?? []).length).toBe(11);
  });

  it("富文本真被改过时，重新序列化的条目必须把 run 颜色写回去", async () => {
    const file = await openSample();
    const sheet = (await file.sheetAt(0)) as SheetModel;
    const rich = sheet.cells["A2"];
    if (!rich?.runs) throw new Error("夹具缺少富文本格 A2");
    rich.runs = [{ ...rich.runs[0]!, text: "浓" }, ...rich.runs.slice(1)];
    rich.value = "浓蓝";
    file.writeSheet(sheet);

    const parts = await partTexts(await file.toBytes());
    expect(parts["xl/sharedStrings.xml"]).toContain('<si><r><rPr><b/></rPr><t>浓</t></r><r><rPr><color rgb="FF0000FF"/></rPr><t>蓝</t></r></si>');
  });

  it("改一格之后，未建模结构与其它部件仍然原样", async () => {
    const file = await openSample();
    const sheet = (await file.sheetAt(0)) as SheetModel;
    const target = sheet.cells["C2"];
    if (!target) throw new Error("夹具缺少 C2");
    target.value = 999;
    file.writeSheet(sheet);
    const written = await file.toBytes();

    const parts = await partTexts(written);
    expect(parts["xl/worksheets/sheet1.xml"]).toContain("<conditionalFormatting");
    expect(parts["xl/worksheets/sheet1.xml"]).toContain("<dataValidations");
    expect(parts["xl/worksheets/sheet1.xml"]).toContain('x14ac:dyDescent="0.25"');
    expect(parts["xl/worksheets/sheet1.xml"]).toContain("<mergeCell ref=\"H1:I2\"/>");
    expect(parts["xl/worksheets/sheet1.xml"]).toContain('<c r="E3" s="3"/>');
    expect(parts["xl/theme/theme1.xml"]).toBe((await partTexts(await buildSampleWorkbookBytes()))["xl/theme/theme1.xml"]);
    expect(parts["xl/styles.xml"]).toBe((await partTexts(await buildSampleWorkbookBytes()))["xl/styles.xml"]);
    expect(parts["xl/workbook.xml"]).toContain('<definedName name="合计">销售!$G$2</definedName>');

    const reopened = await SheetFile.fromBytes(written, { fileName: "sample.xlsx" });
    const reopenedSheet = (await reopened.sheetAt(0)) as SheetModel;
    expect(reopenedSheet.cells["C2"]?.value).toBe(999);
    expect(reopenedSheet.cells["D2"]?.numberFormat?.isDateLike).toBe(true);
  });

  it("不建模的外部链接部件原样带着走，引用它的公式也不许被改写（验收 A6）", async () => {
    const original = await buildSampleWorkbookBytes();
    const file = await SheetFile.fromBytes(original, { fileName: "sample.xlsx" });
    expect(file.unsupported.map((item) => item.reason)).toContain("externalLinks");

    // 改一格会重写整个 <sheetData>，跨簿引用必须逐字活下来
    const sheet = (await file.sheetAt(0)) as SheetModel;
    const target = sheet.cells["C3"];
    if (!target) throw new Error("夹具缺少 C3");
    target.value = 200;
    file.writeSheet(sheet);
    const parts = await partTexts(await file.toBytes());

    expect(parts["xl/externalLinks/externalLink1.xml"]).toBe((await partTexts(original))["xl/externalLinks/externalLink1.xml"]);
    expect(parts["xl/externalLinks/_rels/externalLink1.xml.rels"]).toContain('TargetMode="External"');
    expect(parts["xl/worksheets/sheet1.xml"]).toContain(`<f>'[1]Sheet1'!$A$2</f>`);
    expect(parts["xl/workbook.xml"]).toContain('<externalReference r:id="rId5"/>');

    const reopened = await SheetFile.fromBytes(await file.toBytes(), { fileName: "sample.xlsx" });
    const reopenedSheet = (await reopened.sheetAt(0)) as SheetModel;
    expect(reopenedSheet.cells["J2"]?.value).toBe(3);
  });

  it("手改的日期复用已有样式，自定义格式才往 cellXfs 追加（§2.1 落盘侧 + 实测 P3 的位置索引）", async () => {
    const before = await partTexts(await buildSampleWorkbookBytes());
    const file = await openSample();
    const sheet = (await file.sheetAt(0)) as SheetModel;
    const parsed = parseUserCellText("2026-02-01");
    // 内建日期格式 14 在夹具里已有（`s="5"`），复用即可，不该新建样式
    sheet.cells["F2"] = { address: "F2", colIndex: 5, rowIndex: 1, kind: parsed.kind, value: parsed.value, numberFormat: parsed.numberFormat };
    // 自定义中文年月是夹具里没有的格式：追加一条 numFmt + 一条 xf，落在末尾
    sheet.cells["F3"] = {
      address: "F3",
      colIndex: 5,
      rowIndex: 2,
      kind: "number",
      value: parsed.value,
      numberFormat: { numFmtId: 165, formatCode: 'yyyy"年"m"月"', isDateLike: true },
    };
    file.writeSheet(sheet);
    const written = await file.toBytes();
    const parts = await partTexts(written);

    // 前 6 条 xf 必须逐字不变，否则用户已有格式会整体错位
    expect(xfsOf(parts["xl/styles.xml"]).slice(0, 6)).toEqual(xfsOf(before["xl/styles.xml"]));
    expect(xfsOf(parts["xl/styles.xml"]).length).toBe(7);
    expect(parts["xl/worksheets/sheet1.xml"]).toContain('<c r="F2" s="5"><v>46054</v></c>');
    expect(parts["xl/worksheets/sheet1.xml"]).toContain('<c r="F3" s="6"><v>46054</v></c>');
    expect(parts["xl/styles.xml"]).toContain('<numFmt numFmtId="165" formatCode="yyyy&quot;年&quot;m&quot;月&quot;"/>');

    const reopened = await SheetFile.fromBytes(written, { fileName: "sample.xlsx" });
    const reopenedSheet = (await reopened.sheetAt(0)) as SheetModel;
    expect(reopenedSheet.cells["F2"]?.numberFormat).toMatchObject({ numFmtId: 14, isDateLike: true });
    expect(cellDisplayText(reopenedSheet.cells["F3"] as SheetModel["cells"][string])).toBe("2026年2月");
  });

  it("给带字体的格改数字格式：只换 numFmtId，粗体/红字一条都不许丢（§2.3 不许拍平局部格式）", async () => {
    const before = await partTexts(await buildSampleWorkbookBytes());
    const file = await openSample();
    const sheet = (await file.sheetAt(0)) as SheetModel;
    // 夹具里 s="4" 是"粗体红字 + General"那一格；这里只把它的显示改成百分比
    sheet.cells["A3"] = {
      address: "A3",
      colIndex: 0,
      rowIndex: 2,
      kind: "number",
      value: 0.5,
      styleIndex: 4,
      numberFormat: { numFmtId: 10, isDateLike: false },
    };
    file.writeSheet(sheet);
    const written = await file.toBytes();
    const parts = await partTexts(written);
    const xfs = xfsOf(parts["xl/styles.xml"]);

    // 已有 xf 逐字不变（§二 P3 的位置索引红线），新增的那条必须还带着 fontId="1"
    expect(xfs.slice(0, 6)).toEqual(xfsOf(before["xl/styles.xml"]));
    expect(xfs.length).toBe(7);
    expect(xfs[6]).toContain('numFmtId="10"');
    expect(xfs[6]).toContain('fontId="1"');

    const reopened = await SheetFile.fromBytes(written, { fileName: "sample.xlsx" });
    const cell = (await reopened.sheetAt(0)) as SheetModel;
    const a3 = cell.cells["A3"];
    expect(a3).toMatchObject({ kind: "number", value: 0.5 });
    expect(a3?.numberFormat).toMatchObject({ numFmtId: 10, isDateLike: false });
    expect(reopened.styleTable.fontFor(a3?.styleIndex)).toMatchObject({ bold: true });
    expect(cellDisplayText(a3!)).toBe("50.00%");
  });

  it("新增字符串只追加、样式表前 6 条一个字节都不许变（实测 §二 P3 的位置索引风险）", async () => {
    const file = await openSample();
    const sheet = (await file.sheetAt(0)) as SheetModel;
    const beforeStyles = (await partTexts(await buildSampleWorkbookBytes()))["xl/styles.xml"];

    sheet.cells["A6"] = {
      address: "A6",
      colIndex: 0,
      rowIndex: 5,
      kind: "string",
      value: "新加的项目名 带前后空格 ",
    };
    const dateFmtId = file.styleTable.appendNumberFormat("yyyy\"年\"m\"月\"");
    const dateStyleIndex = file.styleTable.appendCellStyle({ numFmtId: dateFmtId });
    const boldIndex = file.styleTable.appendCellStyle({ font: { bold: true, sizePt: 12, color: { hex: "#008000" } } });
    sheet.cells["C6"] = { address: "C6", colIndex: 2, rowIndex: 5, kind: "number", value: 1, styleIndex: dateStyleIndex };
    sheet.cells["E6"] = {
      address: "E6",
      colIndex: 4,
      rowIndex: 5,
      kind: "number",
      value: 46053,
      styleIndex: dateStyleIndex,
      numberFormat: { numFmtId: dateFmtId, formatCode: "yyyy\"年\"m\"月\"", isDateLike: true },
    };
    sheet.cells["F6"] = { address: "F6", colIndex: 5, rowIndex: 5, kind: "number", value: 2, styleIndex: boldIndex };
    // 公式必须同时写缓存值（验收 A4）。
    sheet.cells["G6"] = {
      address: "G6",
      colIndex: 6,
      rowIndex: 5,
      kind: "number",
      value: 3,
      formula: "SUM(C2:C6)",
      cachedValue: 3,
    };
    file.writeSheet(sheet);

    const writtenBytes = await file.toBytes();
    const reopened = await SheetFile.fromBytes(writtenBytes, { fileName: "sample.xlsx" });
    const reopenedSheet = (await reopened.sheetAt(0)) as SheetModel;
    const afterStyles = (await partTexts(writtenBytes))["xl/styles.xml"];

    expect(reopenedSheet.cells["A6"]?.value).toBe("新加的项目名 带前后空格 ");
    expect(reopenedSheet.cells["E6"]?.numberFormat?.isDateLike).toBe(true);
    expect(reopenedSheet.cells["G6"]).toMatchObject({ formula: "SUM(C2:C6)", value: 3 });
    // 原来那 6 条 xf 必须一字不变：位置索引一动，用户已有格式就整体错位。
    expect(xfsOf(afterStyles).slice(0, 6)).toEqual(xfsOf(beforeStyles));
    expect(xfsOf(afterStyles).length).toBe(8);
    expect(reopenedSheet.cells["C2"]?.styleIndex).toBe(1);
    expect(reopenedSheet.cells["D3"]?.styleIndex).toBe(5);
    expect(reopened.styleTable.cellXfCount).toBeGreaterThan(6);
  });

  it("快照可以整批回退到打开时的内容", async () => {
    const file = await openSample();
    const snapshot = file.getSnapshot();
    const sheet = (await file.sheetAt(0)) as SheetModel;
    delete sheet.cells["C2"];
    file.writeSheet(sheet);
    file.restoreSnapshot(snapshot);
    const parts = await partTexts(await file.toBytes());
    expect(parts["xl/worksheets/sheet1.xml"]).toContain('<c r="C2" s="1"><v>120</v></c>');
  });
});

describe("保存时让 Excel 自己重算公式（`calcPr/@fullCalcOnLoad`）", () => {
  /** 打开夹具、用 op 改一格、保存，返回 `xl/workbook.xml` 的部件文本。 */
  async function saveAndReadWorkbookXml(): Promise<string> {
    const file = await openSample();
    const model = (await file.sheetAt(0)) as SheetModel;
    const applied = runSheetOps(sheetPayloadOf(model), [setCellTextOp(sheetPayloadOf(model), "C2", "42")!]);
    applySheetPayload(model, applied.after);
    file.writeSheet(model);
    return (await partTexts(await file.toBytes()))["xl/workbook.xml"];
  }

  it("改过表之后保存才补上标记，已有 `<calcPr>` 时只加这一个属性", async () => {
    const workbookXml = await saveAndReadWorkbookXml();
    expect(workbookXml).toContain('<calcPr calcId="191029" fullCalcOnLoad="1"/>');
    expect(workbookXml.match(/fullCalcOnLoad/g)).toHaveLength(1);
  });

  it("对已带标记的文件再存一遍不叠加（幂等，与 §2.6 同族的约束）", async () => {
    const marked = await saveAndReadWorkbookXml();
    const zip = await JSZip.loadAsync(await buildSampleWorkbookBytes());
    zip.file("xl/workbook.xml", marked);
    const file = await SheetFile.fromBytes(await zip.generateAsync({ type: "uint8array" }), { fileName: "sample.xlsx" });
    const model = (await file.sheetAt(0)) as SheetModel;
    applySheetPayload(model, runSheetOps(sheetPayloadOf(model), [setCellTextOp(sheetPayloadOf(model), "C2", "43")!]).after);
    file.writeSheet(model);
    const second = (await partTexts(await file.toBytes()))["xl/workbook.xml"];
    expect(second.match(/fullCalcOnLoad/g)).toHaveLength(1);
    expect(second).toContain('<calcPr calcId="191029" fullCalcOnLoad="1"/>');
  });

  it("原文件没有 `<calcPr>` 时，标记按 schema 顺序插在 `</workbook>` 之前", async () => {
    const zip = await JSZip.loadAsync(await buildSampleWorkbookBytes());
    zip.file("xl/workbook.xml", (await zip.file("xl/workbook.xml")!.async("string")).replace('<calcPr calcId="191029"/>', ""));
    const file = await SheetFile.fromBytes(await zip.generateAsync({ type: "uint8array" }), { fileName: "book.xlsx" });
    const model = (await file.sheetAt(0)) as SheetModel;
    applySheetPayload(model, runSheetOps(sheetPayloadOf(model), [setCellTextOp(sheetPayloadOf(model), "C2", "42")!]).after);
    file.writeSheet(model);
    expect((await partTexts(await file.toBytes()))["xl/workbook.xml"]).toContain(
      '</definedNames><calcPr fullCalcOnLoad="1"/></workbook>',
    );
  });

  it("没改过任何表时 workbook.xml 逐字节不动（验收 A1）", async () => {
    const file = await openSample();
    const original = (await partTexts(await buildSampleWorkbookBytes()))["xl/workbook.xml"];
    expect((await partTexts(await file.toBytes()))["xl/workbook.xml"]).toBe(original);
  });
});

describe("打开失败时的错误码（验收 A6）", () => {
  it("含宏的包直接拒绝，不尝试解析", async () => {
    const zip = await JSZip.loadAsync(await buildSampleWorkbookBytes());
    zip.file("xl/vbaProject.bin", new Uint8Array([1, 2, 3]));
    await expect(
      SheetFile.fromBytes(await zip.generateAsync({ type: "uint8array" }), { fileName: "book.xlsm" }),
    ).rejects.toMatchObject({ reason: "macro-enabled" });
  });

  it("不是 zip 的输入给出 not-a-zip，而不是崩溃堆栈", async () => {
    const error = await SheetFile.fromBytes(new Uint8Array([1, 2, 3, 4]), { fileName: "bad.xlsx" }).catch((e: unknown) => e);
    expect(error).toBeInstanceOf(SheetOpenError);
    expect((error as SheetOpenError).reason).toBe("not-a-zip");
  });
});

describe("跨界载荷（阶段 2：渲染层只拿结构化格子，保存时按 partPath 写回）", () => {
  it("载荷里没有 sourceXml，按载荷写回的那格读回一致", async () => {
    const file = await openSample();
    const sheet = (await file.sheetAt(0)) as SheetModel;
    const payload = sheetPayloadOf(sheet);
    expect("sourceXml" in payload).toBe(false);

    // 渲染层能碰到的只有这个对象：改一格数字，再追加一格纯文本
    const edited = payload.cells["C2"];
    if (!edited) throw new Error("夹具缺少 C2");
    edited.value = 777;
    payload.cells["A7"] = { address: "A7", colIndex: 0, rowIndex: 6, kind: "string", value: "载荷写入" };
    payload.lastRowIndex = 6;

    // 保存的那一步：主进程重新从磁盘打开同一份包，把载荷按 partPath 配对写回
    const targetFile = await openSample();
    const target = (await targetFile.sheetAt(0)) as SheetModel;
    applySheetPayload(target, payload);
    targetFile.writeSheet(target);
    const bytes = await targetFile.toBytes();
    const parts = await partTexts(bytes);
    expect(parts["xl/worksheets/sheet1.xml"]).toContain('<c r="C2" s="1"><v>777</v></c>');

    const reopened = await SheetFile.fromBytes(bytes, { fileName: "sample.xlsx" });
    const readBack = (await reopened.sheetAt(0)) as SheetModel;
    expect(readBack.cells["C2"]?.value).toBe(777);
    expect(readBack.cells["A7"]?.value).toBe("载荷写入");
    expect(readBack.lastRowIndex).toBe(6);
  });

  it("op 改完再撤销：写出去的文件读回来与原来一致（阶段 3 · C15/C17 的落盘证据）", async () => {
    const file = await openSample();
    const sheet = (await file.sheetAt(0)) as SheetModel;
    const payload = sheetPayloadOf(sheet);
    const originalValue = payload.cells["C2"]?.value;
    const applied = runSheetOps(payload, [setCellTextOp(payload, "C2", "9999")!, setColumnWidthOp(payload, 1, 103)!]);
    expect(applied.ok).toBe(true);

    // 保存那一步与桌面主进程一致：重新从磁盘打开、按 partPath 写回载荷
    const editedFile = await openSample();
    const editedModel = (await editedFile.sheetAt(0)) as SheetModel;
    applySheetPayload(editedModel, applied.after);
    editedFile.writeSheet(editedModel);
    const editedBytes = await editedFile.toBytes();
    const reopened = await SheetFile.fromBytes(editedBytes, { fileName: "sample.xlsx" });
    expect((await reopened.sheetAt(0))?.cells["C2"]?.value).toBe(9999);

    // 撤销后第二次保存：磁盘上已经是改过的那份，从它出发写回逆操作的结果
    const reverted = revertSheetOps(applied.after, applied.inverses);
    expect(reverted.ok).toBe(true);
    const undoneFile = await SheetFile.fromBytes(editedBytes, { fileName: "sample.xlsx" });
    const undoneModel = (await undoneFile.sheetAt(0)) as SheetModel;
    applySheetPayload(undoneModel, reverted.after);
    undoneFile.writeSheet(undoneModel);
    const backAgain = await SheetFile.fromBytes(await undoneFile.toBytes(), { fileName: "sample.xlsx" });
    const readBack = await backAgain.sheetAt(0);
    expect(readBack?.cells["C2"]?.value).toBe(originalValue);
    expect(readBack?.columns).toEqual(payload.columns);
  });

  it("拖出来的列宽真的存进 `<cols>`，撤销后又是原来那两段", async () => {
    const payload = sheetPayloadOf((await (await openSample()).sheetAt(0)) as SheetModel);
    // 把 B 列拖成 103px（= 14 个字符宽）：原来 1–3 列那一段必须被劈开
    const applied = runSheetOps(payload, [setColumnWidthOp(payload, 1, 103)!]);
    expect(applied.ok).toBe(true);

    const edited = await writeAndReopen(applied.after);
    expect(widthSegments(edited.sheet)).toEqual([
      [0, 0, 12],
      [1, 1, 14],
      [2, 2, 12],
      [3, 3, 14],
    ]);
    // 部件里的属性也得是劈开后的样子：劈出来的段不许留着旧 `min`/`max`/`width`
    expect(edited.partXml).toContain(
      '<cols><col min="1" max="1" width="12" customWidth="1"/><col min="2" max="2" width="14" customWidth="1"/><col min="3" max="3" width="12" customWidth="1"/><col min="4" max="4" width="14" customWidth="1"/></cols>',
    );
    // 改写 `<cols>` 不许牵连同一张表里其它没建模的结构
    expect(edited.partXml).toContain('<mergeCell ref="H1:I2"/>');
    expect(edited.partXml).toContain("<conditionalFormatting");

    const reverted = revertSheetOps(applied.after, applied.inverses);
    expect(reverted.ok).toBe(true);
    expect(widthSegments((await writeAndReopen(reverted.after)).sheet)).toEqual([
      [0, 2, 12],
      [3, 3, 14],
    ]);
  });

  it("一批里改两列列宽：两条都写进 `<cols>`，撤销后又是原来那两段（阶段 4 的批内死角）", async () => {
    const payload = sheetPayloadOf((await (await openSample()).sheetAt(0)) as SheetModel);
    // B 列拖成 14 字符、C 列拖成 15 字符：两条 op 都锚在同一段 `{1–3}`（0 基 0–2）上
    const applied = runSheetOps(payload, [setColumnWidthOp(payload, 1, 103)!, setColumnWidthOp(payload, 2, 110)!]);
    expect(applied.ok, applied.skipped.map((item) => item.reason).join(" | ")).toBe(true);

    const edited = await writeAndReopen(applied.after);
    expect(widthSegments(edited.sheet)).toEqual([
      [0, 0, 12],
      [1, 1, 14],
      [2, 2, 15],
      [3, 3, 14],
    ]);
    expect(edited.partXml).toContain('<col min="3" max="3" width="15" customWidth="1"/>');

    const reverted = revertSheetOps(applied.after, applied.inverses);
    expect(reverted.ok, reverted.skipped.map((item) => item.reason).join(" | ")).toBe(true);
    expect(widthSegments((await writeAndReopen(reverted.after)).sheet)).toEqual(widthSegments(payload));
  });

  it("表里本来没有 `<cols>` 时，第一次设列宽把区间插在 `<sheetData>` 之前", async () => {
    // 夹具的第二张表是空表，`<sheetFormatPr/>` 后面直接就是 `<sheetData/>`
    const blank = sheetPayloadOf((await (await openSample()).sheetAt(1)) as SheetModel);
    const applied = runSheetOps(blank, [setColumnWidthOp(blank, 0, 103)!]);
    expect(applied.ok).toBe(true);

    const written = await writeAndReopen(applied.after, 1);
    expect(written.partXml).toContain('<cols><col min="1" max="1" width="14" customWidth="1"/></cols><sheetData/>');
    expect(widthSegments(written.sheet)).toEqual([[0, 0, 14]]);
  });

  it("警告与未建模部件清单必须一路带进载荷，不许在跨界处丢（spec §零.7）", async () => {
    const file = await openSample();
    const workbook = await file.toWorkbook();
    const payload = workbookPayload(workbook.fileName, workbook.sheets, workbook.warnings, workbook.unsupported);
    expect(payload.fileName).toBe("sample.xlsx");
    expect(payload.sheets.map((sheet) => sheet.name)).toEqual(["销售", "空表"]);
    expect(payload.unsupported.map((item) => item.reason)).toContain("externalLinks");
  });
});

describe("插删整行整列的落盘证据（阶段 3 · C16）", () => {
  it("插一行写进文件后维度变大，撤销再存回去与插入前逐字一致", async () => {
    const payload = sheetPayloadOf((await (await openSample()).sheetAt(0)) as SheetModel);
    // J 列的公式带外部引用，它让本表任何插删都被拒绝；先删掉 J 列，再在插入点上做来回
    const dropped = runSheetOps(payload, [deleteTracksOp(payload, "col", 9)!]);
    expect(dropped.ok).toBe(true);
    const baseline = (await writeAndReopen(dropped.after)).partXml;

    const applied = runSheetOps(dropped.after, [insertTracksOp(dropped.after, "row", 2)!]);
    expect(applied.ok, applied.skipped.map((item) => item.reason).join(" | ")).toBe(true);
    expect(applied.after.cells["A5"]?.value).toBe("3月");
    const edited = await writeAndReopen(applied.after);
    expect(edited.partXml).toContain('<dimension ref="A1:I5"/>');
    // 没建模的条件格式停在原处不跟着挪——这是插删的已知代价，spec 阶段 3 里明写着
    expect(edited.partXml).toContain('<conditionalFormatting sqref="C2:C4">');

    const reverted = revertSheetOps(applied.after, applied.inverses);
    expect(reverted.ok).toBe(true);
    expect((await writeAndReopen(reverted.after)).partXml).toBe(baseline);
  });

  it("表里只有外部工作簿引用时不再否决插列，那段原样带过（spec D9 第 1 条）", async () => {
    const payload = sheetPayloadOf((await (await openSample()).sheetAt(0)) as SheetModel);
    const applied = runSheetOps(payload, [insertTracksOp(payload, "col", 3)!]);
    expect(applied.ok, applied.skipped.map((item) => item.reason).join(" | ")).toBe(true);
    // 外部引用的坐标系在别的文件里：格子随本地列位移挪位，公式文本一个字都不改。
    expect(applied.after.cells["K2"]?.formula).toBe("'[1]Sheet1'!$A$2");

    const edited = await writeAndReopen(applied.after);
    expect(edited.partXml).toContain("'[1]Sheet1'!$A$2");
    const reverted = revertSheetOps(edited.sheet, applied.inverses.map((op) => op));
    expect(reverted.skipped).toEqual([]);
  });

  it("删除带把合并区劈成两半时拒绝：合不回去的改动不做", async () => {
    const payload = sheetPayloadOf((await (await openSample()).sheetAt(0)) as SheetModel);
    const applied = runSheetOps(payload, [deleteTracksOp(payload, "col", 7)!]);
    expect(applied.ok).toBe(false);
    expect(applied.skipped[0]?.reason).toContain("H1:I2");
    expect(applied.skipped[0]?.reason).toContain("劈开");
    expect(applied.after.merges).toEqual(payload.merges);
  });

  it("合并区跟着插列挪位：写进文件读回来是 I1:J2，撤销后回到 H1:I2", async () => {
    const payload = sheetPayloadOf((await (await openSample()).sheetAt(0)) as SheetModel);
    // J 列的公式带外部引用，任何让它位移的插列都会被拒；先删掉 J 列再插
    const dropped = runSheetOps(payload, [deleteTracksOp(payload, "col", 9)!]);
    expect(dropped.ok).toBe(true);

    const applied = runSheetOps(dropped.after, [insertTracksOp(dropped.after, "col", 3)!]);
    expect(applied.ok).toBe(true);
    const edited = await writeAndReopen(applied.after);
    expect(edited.partXml).toContain('<mergeCells count="1"><mergeCell ref="I1:J2"/></mergeCells>');
    expect(edited.sheet.merges.map((merge) => merge.ref)).toEqual(["I1:J2"]);
    expect(widthSegments(edited.sheet)).toEqual([
      [0, 2, 12],
      [4, 4, 14],
    ]);
    // 改写 `<mergeCells>` 不许牵连同一张表里其它没建模的结构
    expect(edited.partXml).toContain("<dataValidation");
    expect(edited.partXml).toContain("<pane ");

    const reverted = revertSheetOps(applied.after, applied.inverses);
    expect(reverted.ok).toBe(true);
    const back = await writeAndReopen(reverted.after);
    expect(back.partXml).toContain('<mergeCells count="1"><mergeCell ref="H1:I2"/></mergeCells>');
    expect(widthSegments(back.sheet)).toEqual([
      [0, 2, 12],
      [3, 3, 14],
    ]);
  });
});

describe("按一列重排整行落盘（spec 决策 D7：格式跟着记录走）", () => {
  it("真实夹具里能排的区域排完写回文件，读回来整行连格式一起换位", async () => {
    const payload = sheetPayloadOf((await (await openSample()).sheetAt(0)) as SheetModel);
    // 第 3~4 行是夹具里唯一"整条记录都在区域内、又不带公式"的一带：C 列 150 与 90
    const applied = runSheetOps(payload, [sortRowsOp(payload, { range: "A3:E4", column: "C" })!.op]);
    expect(applied.ok, applied.skipped.map((item) => item.reason).join(" | ")).toBe(true);

    const edited = await writeAndReopen(applied.after);
    expect(edited.sheet.cells["A3"]?.value).toBe("3月");
    expect(edited.sheet.cells["C3"]?.value).toBe(90);
    expect(edited.sheet.cells["A4"]?.value).toBe("2月");
    expect(edited.sheet.cells["C4"]?.value).toBe(150);
    // 日期样式（内建 14）与百分比样式（内建 10）各随各的记录走；A4 的粗体红字不许被拍平
    expect(edited.sheet.cells["D4"]?.numberFormat).toMatchObject({ numFmtId: 14, isDateLike: true });
    expect(edited.sheet.cells["E4"]?.numberFormat).toMatchObject({ numFmtId: 10, isDateLike: false });
    expect(edited.sheet.cells["E4"]?.styleOnly).toBe(true);
    expect(edited.sheet.cells["E3"]).toBeUndefined();
    expect(edited.sheet.cells["A4"]?.styleIndex).toBe(4);

    const reverted = revertSheetOps(applied.after, applied.inverses);
    expect(reverted.ok, reverted.skipped.map((item) => item.reason).join(" | ")).toBe(true);
    const baseline = (await writeAndReopen(payload)).partXml;
    expect((await writeAndReopen(reverted.after)).partXml).toBe(baseline);
  });

  it("真实夹具里探出区域的格子与带公式的行一律拒绝", async () => {
    const payload = sheetPayloadOf((await (await openSample()).sheetAt(0)) as SheetModel);
    const outside = sortRowsOp(payload, { range: "A2:E4", column: "C" });
    expect(outside.ok).toBe(false);
    if (!outside.ok) expect(outside.reason).toContain("G2 在排序区域之外还有内容");

    const withFormula = sortRowsOp(payload, { range: "A2:J4", column: "C" });
    expect(withFormula.ok).toBe(false);
    if (!withFormula.ok) expect(withFormula.reason).toContain("带公式");
  });
});

describe("公式落盘（spec 阶段 4 D4：`<f>` 交给 Excel，缓存值 `<v>` 由本地算）", () => {
  it("写一条 AVERAGE：文件里 `<f>` 与算出的 `<v>` 并存，读回来两者都在", async () => {
    const payload = sheetPayloadOf((await (await openSample()).sheetAt(0)) as SheetModel);
    const created = setCellFormulaOp(payload, "F2", "=AVERAGE(C2:C4)");
    expect(created?.recalcNeeded).toBe(false);
    const applied = runSheetOps(payload, [created!.op]);
    expect(applied.ok, applied.skipped.map((item) => item.reason).join(" | ")).toBe(true);

    const edited = await writeAndReopen(applied.after);
    expect(edited.partXml).toContain('<c r="F2"><f>AVERAGE(C2:C4)</f><v>120</v></c>');
    expect(edited.sheet.cells["F2"]?.formula).toBe("AVERAGE(C2:C4)");
    expect(edited.sheet.cells["F2"]?.value).toBe(120);
    expect(edited.sheet.cells["F2"]?.cachedValue).toBe(120);
    // 已有的公式格与未建模结构不许被这次写入牵连
    expect(edited.partXml).toContain('<c r="G2"><f>SUM(C2:C4)</f><v>360</v></c>');
    expect(edited.partXml).toContain("<conditionalFormatting");

    const reverted = revertSheetOps(applied.after, applied.inverses);
    expect(reverted.ok).toBe(true);
    const back = await writeAndReopen(reverted.after);
    expect(back.partXml).not.toContain("AVERAGE(C2:C4)");
    expect(back.sheet.cells["F2"]).toBeUndefined();
  });

  it("本地算不出的公式只写 `<f>`，绝不留一个过期数字（验收 A4）", async () => {
    const payload = sheetPayloadOf((await (await openSample()).sheetAt(0)) as SheetModel);
    const created = setCellFormulaOp(payload, "F3", "TODAY()");
    expect(created?.recalcNeeded).toBe(true);
    const applied = runSheetOps(payload, [created!.op]);
    expect(applied.ok).toBe(true);

    const edited = await writeAndReopen(applied.after);
    expect(edited.partXml).toContain('<c r="F3"><f>TODAY()</f></c>');
    expect(edited.partXml).not.toContain('<c r="F3"><f>TODAY()</f><v>');
    expect(edited.sheet.cells["F3"]?.formula).toBe("TODAY()");
    expect(edited.sheet.cells["F3"]?.value).toBeNull();
  });

  it("把带格式的格改成公式：`@s` 原样带走，显示仍按它的数字格式（§2.3）", async () => {
    const payload = sheetPayloadOf((await (await openSample()).sheetAt(0)) as SheetModel);
    const applied = runSheetOps(payload, [setCellFormulaOp(payload, "C2", "C3+C4")!.op]);
    expect(applied.ok).toBe(true);
    expect(applied.after.cells["C2"]).toMatchObject({ styleIndex: payload.cells["C2"]?.styleIndex, value: 240 });

    const edited = await writeAndReopen(applied.after);
    expect(edited.partXml).toContain(`<c r="C2" s="${payload.cells["C2"]?.styleIndex}"><f>C3+C4</f><v>240</v></c>`);
    // 缓存值仍按这一格自己的数字格式显示（实测夹具 `@s=1` 是内建 3：千分位、不带小数）
    expect(cellDisplayText(edited.sheet.cells["C2"]!)).toBe("240");
  });

  it("改了公式引用的格子后存盘：`<v>` 是新和，撤销再存回到 Excel 留下的那份缓存（spec D8 的落盘证据）", async () => {
    const payload = sheetPayloadOf((await (await openSample()).sheetAt(0)) as SheetModel);
    const originalSum = Number(payload.cells["G2"]?.cachedValue);
    const increaseBy = 100;
    const applied = runSheetOps(payload, [setCellTextOp(payload, "C2", String(Number(payload.cells["C2"]?.value) + increaseBy))!]);
    expect(applied.ok, applied.skipped.map((item) => item.reason).join(" | ")).toBe(true);

    const written = await writeAndReopen(applied.after);
    expect(written.partXml).toContain(`<c r="G2"><f>SUM(C2:C4)</f><v>${originalSum + increaseBy}</v></c>`);
    expect(cellDisplayText(written.sheet.cells["G2"]!)).toBe(String(originalSum + increaseBy));

    const reverted = revertSheetOps(applied.after, applied.inverses);
    expect(reverted.ok, reverted.skipped.map((item) => item.reason).join(" | ")).toBe(true);
    const undone = await writeAndReopen(reverted.after);
    expect(undone.partXml).toContain(`<c r="G2"><f>SUM(C2:C4)</f><v>${originalSum}</v></c>`);
  });
});

describe("分析读工具跑在真实 .xlsx 上（阶段 4 · 验收 B10/B11 的落盘证据）", () => {
  /** 打开夹具第一张表并换成分析函数吃的载荷。 */
  async function samplePayload(): Promise<SheetPayload> {
    return sheetPayloadOf((await (await openSample()).sheetAt(0)) as SheetModel);
  }

  it("表画像的行列数与逐列数字都来自文件，合计与表里那条公式的缓存值对上", async () => {
    const payload = await samplePayload();
    const profile = describeSheet(payload);
    expect(profile).toMatchObject({ sheetName: "销售", usedRange: "A1:J4", rowCount: 3, columnCount: 10, columnsTruncated: false });
    const sales = profile.columns[2]!;
    expect(sales).toMatchObject({ header: "销量", type: "number", populated: 3 });
    expect(sales.numeric).toMatchObject({ count: 3, sum: 360, min: 90, max: 150 });
    // G2 存着 Excel 留下的 `SUM(C2:C4)` 缓存值：引擎自己算出的合计必须与它一致
    expect(payload.cells.G2?.cachedValue).toBe(sales.numeric?.sum);
    expect(profile.columns[3]!.type).toBe("date");
    expect(profile.columns[6]!.header).toBe("G");
  });

  it("readRange 的显示串与网格一致：日期按单元格格式、百分比带百分号、富文本合成一句", async () => {
    const read = readRange(await samplePayload(), { range: "A1:F4" });
    expect(read).toMatchObject({ ok: true, ref: "A1:F4", rowCount: 4, colCount: 6, truncated: false });
    expect(read.rows.map((row) => row.rowNumber)).toEqual([1, 2, 3, 4]);
    expect(read.rows[0]!.cells).toEqual(["月份", "地区", "销量", "日期", "增长", "备注"]);
    expect(read.rows[1]!.cells[0]).toBe("粗蓝");
    expect(read.rows[1]!.cells[2]).toBe("120");
    // D2 是序列号 46053 + 自定义格式 `yyyy-mm-dd`：显示串由引擎换算，不是我们手算的日期（§2.1）
    expect(read.rows[1]!.cells[3]).toBe("2026-01-31");
    expect(read.rows[1]!.cells[4]).toBe("12.00%");
    expect(read.rows[1]!.cells[5]).toBe("");
  });

  it("findRows 的条件口径在真实表上成立，命中给的是能直接引用的地址", async () => {
    const payload = await samplePayload();
    const hits = findRows(payload, { column: "销量", criteria: ">100" });
    expect(hits).toMatchObject({ ok: true, matchedCount: 2, truncated: false });
    expect(hits.hits.map((hit) => `${hit.address}:${hit.display}`)).toEqual(["C2:120", "C3:150"]);
    expect(findRows(payload, { column: "地区", criteria: "<>华东" }).hits.map((hit) => hit.rowNumber)).toEqual([4]);
    expect(findRows(payload, { column: "月份", criteria: "3月" }).hits.map((hit) => hit.address)).toEqual(["A4"]);
  });
});

/**
 * 只有 `t="inlineStr"`、连 `xl/sharedStrings.xml` 都没有的一份包（openpyxl 这类导出器的实际形状）。
 *
 * 单独造这份夹具是为了钉住一个实测缺陷：我们把文本一律写成 `t="s"`，就必须凭空新建这张表，
 * 而**光写部件不算登记**——读方是按 `xl/workbook.xml` 的关系与内容类型找表的。
 *
 * @param sheetDataInner `<sheetData>` 里要放的内容；默认是纯 inline 文本，传富文本那份可以钉住 run 的保留。
 * @returns `.xlsx` 字节：一张两行两列的小表，A 列文本、B 列数值。
 */
async function buildInlineOnlyBytes(sheetDataInner = `<row r="1"><c r="A1" t="inlineStr"><is><t>月份</t></is></c><c r="B1" t="inlineStr"><is><t>销售额</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>1月</t></is></c><c r="B2"><v>100</v></c></row>`): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`,
  );
  zip.file(
    "_rels/.rels",
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
  );
  zip.file(
    "xl/workbook.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="销售" sheetId="1" r:id="rId1"/></sheets></workbook>`,
  );
  zip.file(
    "xl/_rels/workbook.xml.rels",
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`,
  );
  zip.file(
    "xl/styles.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>`,
  );
  zip.file(
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:B2"/><sheetData>${sheetDataInner}</sheetData></worksheet>`,
  );
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}

describe("原文件没有字符串表时把新表登记齐全（实测：openpyxl 找不到表就报错）", () => {
  it("写一个 inlineStr 格要连部件、workbook 关系、内容类型一起补上", async () => {
    const file = await SheetFile.fromBytes(await buildInlineOnlyBytes(), { fileName: "inline.xlsx" });
    const sheet = (await file.sheetAt(0)) as SheetModel;
    expect(sheet.cells["A1"]?.value).toBe("月份");
    const op = setCellTextOp(sheet, "C1", "备注");
    if (!op) throw new Error("这一格建不出 op");
    const result = runSheetOps(sheet, [op]);
    expect(result.ops).toHaveLength(1);
    file.writeSheet(result.after as SheetModel);

    const parts = await partTexts(await file.toBytes());
    expect(parts["xl/sharedStrings.xml"]).toContain("<si><t>备注</t></si>");
    expect(parts["xl/_rels/workbook.xml.rels"]).toContain('Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"');
    expect(parts["[Content_Types].xml"]).toContain('PartName="/xl/sharedStrings.xml"');

    const reopened = await SheetFile.fromBytes(await file.toBytes(), { fileName: "inline.xlsx" });
    const back = (await reopened.sheetAt(0)) as SheetModel;
    expect(back.cells["C1"]?.value).toBe("备注");
    // inline 文本被改成共享字符串之后，原来那两句还得读得回来（表是新建的，条目按出现顺序排）
    expect(back.cells["A1"]?.value).toBe("月份");
    expect(back.cells["B2"]?.value).toBe(100);
  });

  it("原文件已经有字符串表时，关系与内容类型一个字都不许多写", async () => {
    const before = await partTexts(await buildSampleWorkbookBytes());
    const file = await openSample();
    const sheet = (await file.sheetAt(0)) as SheetModel;
    const op = setCellTextOp(sheet, "C5", "四月");
    if (!op) throw new Error("这一格建不出 op");
    file.writeSheet(runSheetOps(sheet, [op]).after as SheetModel);

    const parts = await partTexts(await file.toBytes());
    expect(parts["xl/_rels/workbook.xml.rels"]).toBe(before["xl/_rels/workbook.xml.rels"]);
    expect(parts["[Content_Types].xml"]).toBe(before["[Content_Types].xml"]);
    expect(parts["xl/sharedStrings.xml"]).toContain("<si><t>四月</t></si>");
  });
});

/**
 * `<is>` 里带 `<r>` run 的 inline 富文本——openpyxl 的 `CellRichText`、部分导出器的实际产出。
 * A1 三段：粗体的"月"、普通的"份"、红字的"表"。
 */
const RICH_INLINE_SHEET_DATA =
  `<row r="1"><c r="A1" t="inlineStr"><is><r><rPr><b val="1"/></rPr><t>月</t></r><r><t>份</t></r><r><rPr><color rgb="FFFF0000"/></rPr><t>表</t></r></is></c><c r="B1" t="inlineStr"><is><t>销售额</t></is></c></row>` +
  `<row r="2"><c r="A2" t="inlineStr"><is><t>1月</t></is></c><c r="B2"><v>100</v></c></row>`;

describe("inline 富文本原样回写（实测缺陷：保存一次就把粗体红字拍平成纯文本）", () => {
  it("一个字都没改的表，run 与整体形状原封不动回到文件里", async () => {
    const bytes = await buildInlineOnlyBytes(RICH_INLINE_SHEET_DATA);
    const file = await SheetFile.fromBytes(bytes, { fileName: "rich.xlsx" });
    const sheet = (await file.sheetAt(0)) as SheetModel;
    expect(sheet.cells["A1"]?.value).toBe("月份表");

    // 桌面保存路径：载荷过一遍跨界层，再按 partPath 配对写回
    const target = await SheetFile.fromBytes(bytes, { fileName: "rich.xlsx" });
    const model = (await target.sheetAt(0)) as SheetModel;
    applySheetPayload(model, sheetPayloadOf(sheet));
    target.writeSheet(model);
    const out = await target.toBytes();

    const parts = await partTexts(out);
    expect(parts["xl/worksheets/sheet1.xml"]).toContain(
      `<c r="A1" t="inlineStr"><is><r><rPr><b val="1"/></rPr><t>月</t></r><r><t>份</t></r><r><rPr><color rgb="FFFF0000"/></rPr><t>表</t></r></is></c>`,
    );
    // 拍平的证据就是这张凭空长出来的表：inline 文本被搬进 `xl/sharedStrings.xml` 时 run 已经没了
    expect(parts["xl/sharedStrings.xml"]).toBe("");

    const reopened = await SheetFile.fromBytes(out, { fileName: "rich.xlsx" });
    expect((await reopened.sheetAt(0))?.cells["A1"]?.value).toBe("月份表");
  });

  it("改过的那一格照常写成共享字符串，没改的那格仍带着 run", async () => {
    const bytes = await buildInlineOnlyBytes(RICH_INLINE_SHEET_DATA);
    const file = await SheetFile.fromBytes(bytes, { fileName: "rich.xlsx" });
    const sheet = (await file.sheetAt(0)) as SheetModel;
    const op = setCellTextOp(sheet, "A1", "改过了");
    if (!op) throw new Error("这一格建不出 op");
    const result = runSheetOps(sheet, [op]);
    expect(result.after.cells["A1"]?.rawXml).toBeUndefined();
    file.writeSheet(result.after as SheetModel);

    const parts = await partTexts(await file.toBytes());
    expect(parts["xl/worksheets/sheet1.xml"]).toContain(`<c r="A1" t="s"><v>0</v></c>`);
    expect(parts["xl/sharedStrings.xml"]).toContain("<si><t>改过了</t></si>");
    expect(parts["xl/worksheets/sheet1.xml"]).toContain('<c r="B1" t="inlineStr"><is><t>销售额</t></is></c>');

    const reopened = await SheetFile.fromBytes(await file.toBytes(), { fileName: "rich.xlsx" });
    expect((await reopened.sheetAt(0))?.cells["A1"]?.value).toBe("改过了");
  });
});
