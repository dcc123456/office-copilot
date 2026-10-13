/**
 * 把 `ResolvedChart` 拼成 `.xlsx` 里的原生图表部件（spec 决策 D5 · 提交 15）。
 *
 * 形状来自本机实测的一手参照：用 openpyxl 3.1.5 往同样的数据上插一张柱状图（`tmp/chart-ref.xlsx`），
 * 拆出它写的 `xl/charts/chart1.xml` + `xl/drawings/drawing1.xml` + 两层关系，确认"最小可用集"就是这些元素；
 * 在它之上再补 Excel 自己一定会写的 `strCache`/`numCache`（没有缓存值时，只看缓存的查看器会画空图）。
 * 这里是**纯字符串拼装**：部件之间的引用关系由 `SheetFile.addChart` 负责登记（`sheet-file.ts`）。
 */
import { chartExtentEmu, type ResolvedChart } from "./chart-spec.js";
import { relativeTarget } from "./opc.js";
import { escapeAttr, escapeText } from "./xml.js";

/** 图表部件与画图部件的产物集合，路径都是包内正斜杠路径。 */
export interface ChartPartSet {
  chartPartPath: string;
  drawingPartPath: string;
  drawingRelsPath: string;
  chartXml: string;
  drawingXml: string;
  drawingRelsXml: string;
}

/** 命名空间常量，写法与 ECMA-376 Part 1 §21.2 一致。 */
const C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const XDR_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const REL_CHART = `${R_NS}/chart`;
const REL_DRAWING = `${R_NS}/drawing`;

/** `[Content_Types].xml` 里这两个部件的声明（写错类型 Excel 直接判定文件损坏）。 */
export const CHART_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
export const DRAWING_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.drawing+xml";

/**
 * 生成 `<c:strRef>` / `<c:numRef>` 里的那一段引用 + 缓存。
 *
 * @param tag `str` 或 `num`，决定元素名与缓存元素名。
 * @param ref 已带表名前缀的 A1 引用串，如 `'销售'!$C$2:$C$4`。
 * @param points 逐点的显示值（文本类别）或数值（数据系列）。
 * @returns 可直接嵌进 `<c:cat>` / `<c:val>` / `<c:tx>` 的片段。
 */
function refWithCache(tag: "str" | "num", ref: string, points: Array<string | number>): string {
  const cacheName = tag === "str" ? "strCache" : "numCache";
  const format = tag === "str" ? "" : "<c:formatCode>General</c:formatCode>";
  const pts = points
    .map((value, index) => `<c:pt idx="${index}"><c:v>${escapeText(String(value))}</c:v></c:pt>`)
    .join("");
  return (
    `<c:${tag}Ref><c:f>${escapeText(ref)}</c:f>` +
    `<c:${cacheName}>${format}<c:ptCount val="${points.length}"/>${pts}</c:${cacheName}></c:${tag}Ref>`
  );
}

/**
 * 拼一条系列（`<c:ser>`）。
 *
 * @param chart 已解析的图表。
 * @param index 系列序号，`idx`/`order` 都用它。
 * @returns `<c:ser>` 片段；类别轴按现场判定决定写 `strRef` 还是 `numRef`。
 */
function seriesXml(chart: ResolvedChart, index: number): string {
  const series = chart.series[index]!;
  const name = series.nameRef
    ? `<c:tx>${refWithCache("str", series.nameRef, [series.name])}</c:tx>`
    : "";
  const cat = `<c:cat>${chart.categoriesAreNumeric
    ? refWithCache("num", chart.categoryRef, chart.categoryNumbers)
    : refWithCache("str", chart.categoryRef, chart.categories)}</c:cat>`;
  return `<c:ser><c:idx val="${index}"/><c:order val="${index}"/>${name}${cat}<c:val>${refWithCache(
    "num",
    series.valueRef,
    series.values,
  )}</c:val></c:ser>`;
}

/**
 * 拼坐标轴（柱/线用，饼图没有轴）。
 *
 * @returns `<c:catAx>` + `<c:valAx>` 片段，`axId` 与图型块里的两条 `axId` 对应。
 */
function axesXml(): string {
  return (
    '<c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/>' +
    '<c:axPos val="l"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:crossAx val="100"/><c:lblOffset val="100"/></c:catAx>' +
    '<c:valAx><c:axId val="100"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/>' +
    '<c:axPos val="l"/><c:majorGridlines/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:crossAx val="10"/></c:valAx>'
  );
}

/**
 * 图表部件正文。
 *
 * @param chart 已解析的图表。
 * @returns 带 XML 声明的 `xl/charts/chartN.xml` 文本。
 */
export function chartXmlOf(chart: ResolvedChart): string {
  const series = chart.series.map((_, index) => seriesXml(chart, index)).join("");
  const body =
    chart.kind === "pie"
      ? `<c:pieChart>${series}<c:firstSliceAng val="0"/></c:pieChart>`
      : chart.kind === "line"
        ? `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${series}<c:marker val="1"/><c:axId val="10"/><c:axId val="100"/></c:lineChart>${axesXml()}`
        : `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>${series}<c:gapWidth val="150"/><c:axId val="10"/><c:axId val="100"/></c:barChart>${axesXml()}`;
  const title = `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${escapeText(
    chart.title,
  )}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>`;
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<c:chartSpace xmlns:c="${C_NS}" xmlns:a="${A_NS}" xmlns:r="${R_NS}">` +
    `<c:chart>${title}<c:plotArea>${body}<c:spPr><a:noFill/></c:spPr></c:plotArea>` +
    `<c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend>` +
    `<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>`
  );
}

/**
 * 画图部件正文：把图表按一个"锚在格子上"的框放进工作表。
 *
 * @param chart 已解析的图表。
 * @returns `xl/drawings/drawingN.xml` 文本，引用 `rId1`（由图表自己的关系解析）。
 */
export function drawingXmlOf(chart: ResolvedChart): string {
  const ext = chartExtentEmu(chart);
  const from =
    `<xdr:from><xdr:col>${chart.anchor.colIndex}</xdr:col><xdr:colOff>0</xdr:colOff>` +
    `<xdr:row>${chart.anchor.rowIndex}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>`;
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<xdr:wsDr xmlns:xdr="${XDR_NS}" xmlns:a="${A_NS}" xmlns:r="${R_NS}">` +
    `<xdr:oneCellAnchor>${from}<xdr:ext cx="${ext.cx}" cy="${ext.cy}"/>` +
    `<xdr:graphicFrame macro="">` +
    `<xdr:nvGraphicFramePr><xdr:cNvPr id="2000" name="${escapeAttr(chart.title)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
    `<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>` +
    `<a:graphic><a:graphicData uri="${C_NS}">` +
    `<c:chart xmlns:c="${C_NS}" xmlns:r="${R_NS}" r:id="rId1"/>` +
    `</a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:oneCellAnchor></xdr:wsDr>`
  );
}

/**
 * 一组新部件的路径与正文。
 *
 * @param chart 已解析的图表。
 * @param indexes `{ chart: N, drawing: M }`，用于避开文件里已有的同名部件。
 * @returns `ChartPartSet`；调用方负责把这些路径登记进包里并补关系与 content types。
 */
export function buildChartParts(chart: ResolvedChart, indexes: { chart: number; drawing: number }): ChartPartSet {
  const chartPartPath = `xl/charts/chart${indexes.chart}.xml`;
  const drawingPartPath = `xl/drawings/drawing${indexes.drawing}.xml`;
  return {
    chartPartPath,
    drawingPartPath,
    drawingRelsPath: `xl/drawings/_rels/drawing${indexes.drawing}.xml.rels`,
    chartXml: chartXmlOf(chart),
    drawingXml: drawingXmlOf(chart),
    drawingRelsXml:
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="${escapeAttr(REL_CHART)}" Target="${escapeAttr(relativeTarget(drawingPartPath, chartPartPath))}"/>` +
      `</Relationships>`,
  };
}

/** 画图部件的关系类型 URI（上层往 sheet 的关系里追加 `<c:drawing>` 时用）。 */
export const DRAWING_REL_TYPE = REL_DRAWING;
