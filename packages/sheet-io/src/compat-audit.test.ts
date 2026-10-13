import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { buildSampleWorkbookBytes } from "../scripts/make-sample-sheet.js";
import { auditXlsxPackage } from "./compat-audit.js";
import { resolveChartSpec } from "./chart-spec.js";
import { SheetFile } from "./sheet-file.js";
import { sheetPayloadOf } from "./sheet-dto.js";
import { runSheetOps, setCellTextOp, sortRowsOp } from "./sheet-edit.js";
import type { SheetModel } from "./model.js";

/**
 * 造一个最小 `.xlsx` 包字节（只放审计会看的部件，别的一个都不带）。
 *
 * @param parts 部件路径 → 文本；`[Content_Types].xml` 与关系表由调用方自己给。
 * @returns zip 字节，可直接喂 `auditXlsxPackage`。
 */
async function packageBytes(parts: Record<string, string>): Promise<Uint8Array> {
  const zip = new JSZip();
  for (const [path, text] of Object.entries(parts)) zip.file(path, text);
  return zip.generateAsync({ type: "uint8array" });
}

/** 一份"结构没问题"的包：内容类型表覆盖全部部件，关系指向存在，表部件顺序合规。 */
const OK_PARTS: Record<string, string> = {
  "[Content_Types].xml":
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
    '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
    "</Types>",
  "_rels/.rels":
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
    "</Relationships>",
  "xl/workbook.xml":
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    '<sheets><sheet name="销售" sheetId="1" r:id="rId1"/></sheets></workbook>',
  "xl/_rels/workbook.xml.rels":
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
    "</Relationships>",
  "xl/worksheets/sheet1.xml":
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:B2"/><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><pageMargins left="0.5"/></worksheet>',
};

describe("auditXlsxPackage：办公软件会判损坏的几类结构问题", () => {
  it("结构合规的包返回空数组（不许误报）", async () => {
    expect(await auditXlsxPackage(await packageBytes(OK_PARTS))).toEqual([]);
  });

  it("部件没被 [Content_Types] 覆盖时报一条，detail 是它的扩展名", async () => {
    const parts = { ...OK_PARTS, "xl/printerSettings.bin": "binary-blob" };
    const issues = await auditXlsxPackage(await packageBytes(parts));
    expect(issues).toEqual([{ code: "part-not-declared", part: "xl/printerSettings.bin", detail: "bin" }]);
  });

  it("Override 指向不存在的部件也算说谎，Excel 打不开", async () => {
    const parts = {
      ...OK_PARTS,
      "[Content_Types].xml": OK_PARTS["[Content_Types].xml"]!.replace(
        "</Types>",
        '<Override PartName="/xl/chartColors1.xml" ContentType="application/vnd.ms-office.chartcolorstyle+xml"/></Types>',
      ),
    };
    const issues = await auditXlsxPackage(await packageBytes(parts));
    expect(issues).toEqual([{ code: "declared-part-missing", part: "xl/chartColors1.xml", detail: "" }]);
  });

  it("关系指向缺失部件、同一关系表里 Id 重复、部件引用了没有的 r:id，三种都抓得到", async () => {
    const brokenRels =
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet9.xml"/>' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
      "</Relationships>";
    const codes = await (async () => {
      const issues = await auditXlsxPackage(
        await packageBytes({ ...OK_PARTS, "xl/_rels/workbook.xml.rels": brokenRels }),
      );
      return issues.map((issue) => `${issue.code}@${issue.part}`);
    })();
    expect(codes).toContain("relationship-target-missing@xl/workbook.xml");
    expect(codes).toContain("relationship-id-duplicate@xl/workbook.xml");

    const dangling = await auditXlsxPackage(
      await packageBytes({
        ...OK_PARTS,
        "xl/worksheets/sheet1.xml": OK_PARTS["xl/worksheets/sheet1.xml"]!.replace(
          "<pageMargins left=\"0.5\"/>",
          "<drawing r:id=\"rId7\"/>",
        ),
      }),
    );
    expect(dangling).toEqual([{ code: "unresolved-relationship-id", part: "xl/worksheets/sheet1.xml", detail: "rId7" }]);
  });

  it("`<worksheet>` 子元素顺序违反 §18.3.1.100 时报出前后两个元素名", async () => {
    const issues = await auditXlsxPackage(
      await packageBytes({
        ...OK_PARTS,
        "xl/worksheets/sheet1.xml":
          '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
          "<pageMargins left=\"0.5\"/><dimension ref=\"A1:B2\"/>" +
          "<sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c></row></sheetData></worksheet>",
      }),
    );
    expect(issues).toEqual([{ code: "worksheet-child-order", part: "xl/worksheets/sheet1.xml", detail: "pageMargins -> dimension" }]);
  });

  it("`<sheetData>` 是 worksheet 唯一的必选孩子，缺了单独报一条", async () => {
    const issues = await auditXlsxPackage(
      await packageBytes({
        ...OK_PARTS,
        "xl/worksheets/sheet1.xml": '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1"/></worksheet>',
      }),
    );
    expect(issues.map((issue) => issue.code)).toContain("worksheet-missing-sheet-data");
  });

  it("图表部件里必须有 `<c:chart>`，只有 chartSpace 空壳时报出来", async () => {
    const parts = {
      ...OK_PARTS,
      "xl/charts/chart1.xml":
        '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:spPr/></c:chartSpace>',
      "[Content_Types].xml": OK_PARTS["[Content_Types].xml"]!.replace(
        "</Types>",
        '<Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>',
      ),
    };
    const issues = await auditXlsxPackage(await packageBytes(parts));
    expect(issues).toEqual([{ code: "chart-missing-chart", part: "xl/charts/chart1.xml", detail: "" }]);
  });
});

/**
 * 我们自己写出的包必须过审计——这条是"导出能在 Excel 里直接打开"的机检面。
 *
 * 本机装不了桌面 Excel（`AGENTS.md` §8），openpyxl 只读值、Numbers 只会说"文件已损坏"，
 * 所以拿同一套判定回头看自己的三种产物：原始夹具、改过格子的写回、加了原生图表的写回。
 */
describe("我们自己的写入产物过审计（不许给自己判损坏）", () => {
  /** 打开夹具并改一个格子，走桌面的保存路径出字节。 */
  async function bytesAfterCellWrite(): Promise<Uint8Array> {
    const file = await SheetFile.fromBytes(await buildSampleWorkbookBytes(), { fileName: "sample.xlsx" });
    const sheet = (await file.sheetAt(0)) as SheetModel;
    const op = setCellTextOp(sheet, "B2", "改过的值");
    if (!op) throw new Error("夹具里没有 B2");
    const after = runSheetOps(sheet, [op]);
    file.writeSheet(after.after as SheetModel);
    return file.toBytes();
  }

  /** 给夹具的第一张表挂一个原生折线图，出字节。 */
  async function bytesAfterChartAdd(): Promise<Uint8Array> {
    const file = await SheetFile.fromBytes(await buildSampleWorkbookBytes(), { fileName: "sample.xlsx" });
    const sheet = sheetPayloadOf((await file.sheetAt(0)) as SheetModel);
    const resolved = resolveChartSpec(sheet, {
      kind: "line",
      title: "销量对比",
      categories: "A2:A4",
      series: [{ values: "C2:C4", name: "C1" }],
      anchor: "E2",
    });
    if (!resolved.ok) throw new Error(JSON.stringify(resolved));
    const added = await file.addChart(sheet.partPath, resolved.chart);
    if (!added.ok) throw new Error(JSON.stringify(added));
    return file.toBytes();
  }

  /** 按 C 列把夹具的第 3~4 行整体换位（spec 决策 D7），走桌面的保存路径出字节。 */
  async function bytesAfterSortWrite(): Promise<Uint8Array> {
    const file = await SheetFile.fromBytes(await buildSampleWorkbookBytes(), { fileName: "sample.xlsx" });
    const sheet = (await file.sheetAt(0)) as SheetModel;
    const made = sortRowsOp(sheet, { range: "A3:E4", column: "C" });
    if (!made.ok) throw new Error(made.reason);
    const after = runSheetOps(sheet, [made.op]);
    file.writeSheet(after.after as SheetModel);
    return file.toBytes();
  }

  it.each([
    ["原始夹具", async () => buildSampleWorkbookBytes()],
    ["改过格子的写回", bytesAfterCellWrite],
    ["重排过整行的写回", bytesAfterSortWrite],
    ["加过图表的写回", bytesAfterChartAdd],
  ] as const)("%s：审计返回空数组", async (_label, produce) => {
    expect(await auditXlsxPackage(await produce())).toEqual([]);
  });
});
