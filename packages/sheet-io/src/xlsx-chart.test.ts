import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { buildSampleWorkbookBytes } from "../scripts/make-sample-sheet.js";
import { SheetFile, type ChartAddResult } from "./sheet-file.js";
import { chartExtentEmu, resolveChartSpec } from "./chart-spec.js";
import { sheetPayloadOf } from "./sheet-dto.js";
import { CHART_CONTENT_TYPE, DRAWING_CONTENT_TYPE } from "./xlsx-chart.js";
import type { SheetModel } from "./model.js";

/**
 * 造一份带图表的 `.xlsx` 字节：解析夹具表里的区域 → 走 `SheetFile.addChart` → 出字节。
 *
 * @param overrides 覆盖默认请求的字段，用于测拒绝分支。
 * @returns 解析结果、落地结果与字节；解析或落地没过时 `bytes` 是 null。
 */
async function writeChart(overrides: Partial<Parameters<typeof resolveChartSpec>[1]> = {}) {
  const file = await SheetFile.fromBytes(await buildSampleWorkbookBytes(), { fileName: "sample.xlsx" });
  const sheet = sheetPayloadOf((await file.sheetAt(0)) as SheetModel);
  const resolved = resolveChartSpec(sheet, {
    kind: "bar",
    title: "销量对比",
    categories: "A2:A4",
    series: [{ values: "C2:C4", name: "C1" }],
    anchor: "E2",
    ...overrides,
  });
  if (!resolved.ok) return { resolved, added: undefined as ChartAddResult | undefined, bytes: null };
  const added = await file.addChart(sheet.partPath, resolved.chart);
  return { resolved, added, bytes: added.ok ? await file.toBytes() : null };
}

/** 把包里的部件名与两个关键部件正文取出来，断言时不必到处写 zip 代码。 */
async function inspect(bytes: Uint8Array) {
  const zip = await JSZip.loadAsync(bytes);
  const text = async (path: string) => (await zip.file(path)?.async("string")) ?? "";
  return {
    names: Object.keys(zip.files),
    chartXml: await text("xl/charts/chart1.xml"),
    drawingXml: await text("xl/drawings/drawing1.xml"),
    drawingRels: await text("xl/drawings/_rels/drawing1.xml.rels"),
    sheetRels: await text("xl/worksheets/_rels/sheet1.xml.rels"),
    contentTypes: await text("[Content_Types].xml"),
    sheetXml: await text("xl/worksheets/sheet1.xml"),
  };
}

describe("图表规格解析（spec 决策 D5 · 验收 B10 的图表侧）", () => {
  it("数值与引用串都由引擎给出，模型不参与取数", async () => {
    const file = await SheetFile.fromBytes(await buildSampleWorkbookBytes(), { fileName: "sample.xlsx" });
    const sheet = sheetPayloadOf((await file.sheetAt(0)) as SheetModel);
    const resolved = resolveChartSpec(sheet, {
      kind: "bar",
      title: "销量对比",
      categories: "A2:A4",
      series: [{ values: "C2:C4", name: "C1" }],
      anchor: "E2",
    });
    if (!resolved.ok) throw new Error(resolved.reason);
    const { chart } = resolved;
    expect(chart.series[0]!.values).toEqual([120, 150, 90]);
    expect(chart.series[0]!.name).toBe("销量");
    expect(chart.categoryRef).toBe(`'${sheet.name}'!$A$2:$A$4`);
    expect(chart.series[0]!.valueRef).toBe(`'${sheet.name}'!$C$2:$C$4`);
    // 类别里 A2 是富文本"粗蓝"，显示串由引擎合成，图表不许自己再拼一遍
    expect(chart.categories).toEqual(["粗蓝", "2月", "3月"]);
    expect(chart.categoriesAreNumeric).toBe(false);
    // px → EMU 只在这一处发生（§2.1）
    expect(chartExtentEmu(chart)).toEqual({ cx: 4_572_000, cy: 2_743_200 });
  });

  it("二维区域、未知图型、条数不对齐、非数值格各自给一句理由", async () => {
    const file = await SheetFile.fromBytes(await buildSampleWorkbookBytes(), { fileName: "sample.xlsx" });
    const sheet = sheetPayloadOf((await file.sheetAt(0)) as SheetModel);
    const base = { categories: "A2:A4", anchor: "E2" };
    expect(reasonOf(resolveChartSpec(sheet, { ...base, kind: "radar", series: [{ values: "C2:C4" }] }))).toContain("radar");
    expect(reasonOf(resolveChartSpec(sheet, { ...base, kind: "bar", series: [{ values: "C2:D4" }] }))).toContain("二维");
    expect(reasonOf(resolveChartSpec(sheet, { ...base, kind: "bar", series: [{ values: "C2:C3" }] }))).toContain("对不上");
    // 地区列是文本，画不出数值系列
    expect(reasonOf(resolveChartSpec(sheet, { ...base, kind: "bar", series: [{ values: "B2:B4" }] }))).toContain("不是数值");
    expect(reasonOf(resolveChartSpec(sheet, { ...base, kind: "bar", series: [], categories: "A2:A4" }))).toContain("至少");
  });

  it("公式格按缓存值入图：G2 的 SUM 有缓存 360，能当一点用", async () => {
    const file = await SheetFile.fromBytes(await buildSampleWorkbookBytes(), { fileName: "sample.xlsx" });
    const sheet = sheetPayloadOf((await file.sheetAt(0)) as SheetModel);
    const resolved = resolveChartSpec(sheet, {
      kind: "line",
      categories: "A2:A2",
      series: [{ values: "G2:G2", name: "C1" }],
      anchor: "L2",
    });
    if (!resolved.ok) throw new Error(resolved.reason);
    expect(resolved.chart.series[0]!.values).toEqual([360]);
  });
});

describe("把图表写进 .xlsx（spec 决策 D5 · 阶段 5 的落盘证据）", () => {
  it("五个部件一起落地，引用链完整", async () => {
    const { resolved, added, bytes } = await writeChart();
    if (!added?.ok || !bytes) throw new Error(resolved.ok ? "没落地" : resolved.reason);
    const seen = await inspect(bytes);
    expect(seen.names).toContain("xl/charts/chart1.xml");
    expect(seen.names).toContain("xl/drawings/drawing1.xml");
    expect(seen.drawingRels).toContain(`Target="../charts/chart1.xml"`);
    expect(seen.sheetRels).toContain(`Target="../drawings/drawing1.xml"`);
    expect(seen.sheetXml).toMatch(/<drawing r:id="rId1"\/>/);
    expect(seen.contentTypes).toContain(`PartName="/xl/charts/chart1.xml" ContentType="${CHART_CONTENT_TYPE}"`);
    expect(seen.contentTypes).toContain(`PartName="/xl/drawings/drawing1.xml" ContentType="${DRAWING_CONTENT_TYPE}"`);
    // 锚点 E2 → col 4 / row 1（0 基），尺寸按 px 换算
    expect(seen.drawingXml).toContain("<xdr:col>4</xdr:col>");
    expect(seen.drawingXml).toContain("<xdr:row>1</xdr:row>");
    expect(seen.drawingXml).toContain('cx="4572000"');
  });

  it("图表正文带引用与缓存值，标题与图型都在", async () => {
    const { resolved, bytes } = await writeChart();
    if (!bytes) throw new Error(resolved.ok ? "没落地" : resolved.reason);
    const { chartXml } = await inspect(bytes);
    expect(chartXml).toContain("<c:barChart>");
    expect(chartXml).toContain('<c:barDir val="col"/>');
    expect(chartXml).toContain(`<c:f>'销售'!$C$2:$C$4</c:f>`);
    expect(chartXml).toContain('<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="3"/>');
    expect(chartXml).toContain('<c:pt idx="2"><c:v>90</c:v></c:pt>');
    expect(chartXml).toContain("<c:strRef>");
    expect(chartXml).toContain("<a:t>销量对比</a:t>");
    expect(chartXml).toContain("<c:catAx>");
  });

  it("保存后重开：图表活着，未建模部件仍然逐条活着，读回的格子一字未改", async () => {
    const before = await SheetFile.fromBytes(await buildSampleWorkbookBytes(), { fileName: "sample.xlsx" });
    const beforeSheet = sheetPayloadOf((await before.sheetAt(0)) as SheetModel);
    const { resolved, added, bytes } = await writeChart();
    if (!added?.ok || !bytes) throw new Error(resolved.ok ? "没落地" : resolved.reason);
    const after = await SheetFile.fromBytes(bytes, { fileName: "sample.xlsx" });
    const afterSheet = sheetPayloadOf((await after.sheetAt(0)) as SheetModel);
    expect(JSON.stringify(afterSheet.cells)).toBe(JSON.stringify(beforeSheet.cells));
    expect(afterSheet.lastColIndex).toBe(beforeSheet.lastColIndex);
    // 现在文件里有图表与画图部件了，按 spec 的口径它们进 `unsupported`（原样保留，不解释）
    expect(after.unsupported.map((item) => item.reason)).toEqual(expect.arrayContaining(["chart", "drawing"]));
    const { sheetXml } = await inspect(bytes);
    expect(sheetXml).toContain('<mergeCell ref="H1:I2"/>');
    expect(sheetXml).toContain('<conditionalFormatting sqref="C2:C4">');
    expect(sheetXml).toContain("<dataValidation");
  });

  it("一张表只能挂一个画图部件：第二次请求被明确拒绝，不静默改坏已有图表", async () => {
    const first = await writeChart();
    expect(first.added?.ok).toBe(true);
    const bytes = first.bytes as Uint8Array;
    const file = await SheetFile.fromBytes(bytes, { fileName: "sample.xlsx" });
    const sheet = sheetPayloadOf((await file.sheetAt(0)) as SheetModel);
    const resolved = resolveChartSpec(sheet, { kind: "pie", categories: "A2:A4", series: [{ values: "C2:C4" }], anchor: "H8" });
    if (!resolved.ok) throw new Error(resolved.reason);
    expect(await file.addChart(sheet.partPath, resolved.chart)).toEqual({ ok: false, code: "drawing-exists" });
  });
});

/** 取解析失败的理由串，便于上面逐条 `toContain`。 */
function reasonOf(result: ReturnType<typeof resolveChartSpec>): string {
  return result.ok ? "" : result.reason;
}
