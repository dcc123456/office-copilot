import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { buildSampleWorkbookBytes } from "../scripts/make-sample-sheet.js";
import { cellDisplayText, parseUserCellText } from "./cell-text.js";
import { SheetFile, SheetOpenError } from "./sheet-file.js";
import { applySheetPayload, sheetPayloadOf, workbookPayload } from "./sheet-dto.js";
import { revertSheetOps, runSheetOps, setCellTextOp, setColumnWidthOp } from "./sheet-edit.js";
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
