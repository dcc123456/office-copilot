/**
 * 阶段 1 的夹具生成器（对齐 `packages/docx-io/scripts/make-sample.ts` 的先例）。
 *
 * 为什么不直接把探针里的 `sample.xlsx` 存进仓库：`AGENTS.md` §7.3 禁止测试产物入库。
 * 这里的 XML **逐条抄自实测结构**（spec §二 P1/P2/P3/P5），包含 `x14ac` 扩展属性、
 * `conditionalFormatting`、`dataValidations`、`pageMargins` 这些我们不建模的东西，
 * 用来钉住"改写一张表不许碰其它结构"这条红线。
 *
 * 单独运行：`node --experimental-strip-types scripts/make-sample-sheet.ts /tmp/sample.xlsx`
 */
import JSZip from "jszip";
import { writeFile } from "node:fs/promises";

/** 实测的字符串表：`uniqueCount=11`，其中一条是富文本两个 run。 */
const SHARED_STRINGS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="12" uniqueCount="11"><si><t>月份</t></si><si><t>地区</t></si><si><t>销量</t></si><si><t>日期</t></si><si><t>增长</t></si><si><t>备注</t></si><si><r><rPr><b/></rPr><t>粗</t></r><r><rPr><color rgb="FF0000FF"/></rPr><t>蓝</t></r></si><si><t>华东</t></si><si><t>2月</t></si><si><t>3月</t></si><si><t>华北</t></si></sst>`;

/** cellXfs：0=General、1=numFmt 3、2=自定义 164(yyyy-mm-dd)、3=10(百分比)、4=粗体红字、5=内建 14(日期)。 */
const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"><numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/></numFmts><fonts count="2" x14ac:knownFonts="1"><font><sz val="11"/><name val="Calibri"/></font><font><b/><color rgb="FFFF0000"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="6"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="3" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="10" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs><dxfs count="1"><dxf><font><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf></dxfs></styleSheet>`;

/** 第一张表：形态尽量齐全（冻结、列宽、合并、公式带缓存值、只有格式的格、日期三种写法）。 */
const SHEET1 = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"><dimension ref="A1:J4"/><sheetViews><sheetView workbookViewId="0"><pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/><selection pane="bottomRight"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15" outlineLevelRow="0" outlineLevelCol="0" x14ac:dyDescent="55"/><cols><col min="1" max="3" width="12" customWidth="1"/><col min="4" max="4" width="14" customWidth="1"/></cols><sheetData><row r="1" spans="1:10" x14ac:dyDescent="0.25"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="s"><v>2</v></c><c r="D1" t="s"><v>3</v></c><c r="E1" t="s"><v>4</v></c><c r="F1" t="s"><v>5</v></c><c r="I1"/></row><row r="2" spans="1:9" x14ac:dyDescent="0.25"><c r="A2" t="s"><v>6</v></c><c r="B2" t="s"><v>7</v></c><c r="C2" s="1"><v>120</v></c><c r="D2" s="2"><v>46053</v></c><c r="E2" s="3"><v>0.12</v></c><c r="G2"><f>SUM(C2:C4)</f><v>360</v></c><c r="H2"/><c r="I2"/><c r="J2"><f>'[1]Sheet1'!$A$2</f><v>3</v></c></row><row r="3" spans="1:5" x14ac:dyDescent="0.25"><c r="A3" s="4" t="s"><v>8</v></c><c r="B3" t="s"><v>7</v></c><c r="C3"><v>150</v></c><c r="D3" s="5"><v>46081</v></c><c r="E3" s="3"/></row><row r="4" spans="1:5" x14ac:dyDescent="0.25"><c r="A4" t="s"><v>9</v></c><c r="B4" t="s"><v>10</v></c><c r="C4"><v>90</v></c><c r="D4" s="5"><v>46112</v></c></row></sheetData><mergeCells count="1"><mergeCell ref="H1:I2"/></mergeCells><conditionalFormatting sqref="C2:C4"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>100</formula></cfRule></conditionalFormatting><dataValidations count="1"><dataValidation type="list" allowBlank="1" sqref="B2:B4"><formula1>"华东,华北"</formula1></dataValidation></dataValidations><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><pageSetup orientation="portrait" scale="100"/></worksheet>`;

/** 第二张表：完全空的表（往返时必须保持空，不许被写成 A1 区域）。 */
const SHEET2 = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><dimension ref="A1"/><sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData/></worksheet>`;

const WORKBOOK = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="销售" sheetId="1" r:id="rId1"/><sheet name="空表" sheetId="2" r:id="rId2"/></sheets><externalReferences><externalReference r:id="rId5"/></externalReferences><definedNames><definedName name="合计">销售!$G$2</definedName></definedNames></workbook>`;

const WORKBOOK_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink1.xml"/></Relationships>`;

const ROOT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>`;

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/externalLinks/externalLink1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/><Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>`;

const CORE_PROPS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>Office Copilot</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">2026-01-01T00:00:00Z</dcterms:created></cp:coreProperties>`;

const APP_PROPS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Microsoft Excel</Application></Properties>`;

/** 主题部件：我们不建模它，但它必须在包里，用来验证"未建模部件逐字节带过"。 */
const THEME = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1></a:clrScheme></a:themeElements></a:theme>`;

/**
 * 外部链接部件（Excel 里"引用了另一个工作簿"时产生）。`sheet-io` 明确**不建模**它，
 * 但必须原样带着走：引用它的公式形如 `'[1]Sheet1'!A2`，索引 `[1]` 就指向这个部件，
 * 部件被丢掉后这些公式在 Excel 里全部取不到值（spec 验收 A6）。
 */
const EXTERNAL_LINK = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><externalBook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"><sheetNames><sheetName val="Sheet1"/></sheetNames><sheetDataSet><sheetData sheetId="0"><row r="2"><cell r="A2"><v>3</v></cell></row></sheetData></sheetDataSet></externalBook></externalLink>`;

/** 外部链接的关系：`TargetMode="External"` + 绝对路径，这类关系不能按包内路径解析。 */
const EXTERNAL_LINK_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="file:///C:\\Users\\Public\\book1.xlsx" TargetMode="External"/></Relationships>`;

/**
 * 生成夹具工作簿。
 *
 * @returns `.xlsx` 的字节；每次调用内容一致（夹具必须可重复）。
 */
export async function buildSampleWorkbookBytes(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", CONTENT_TYPES);
  zip.file("_rels/.rels", ROOT_RELS);
  zip.file("docProps/core.xml", CORE_PROPS);
  zip.file("docProps/app.xml", APP_PROPS);
  zip.file("xl/workbook.xml", WORKBOOK);
  zip.file("xl/_rels/workbook.xml.rels", WORKBOOK_RELS);
  zip.file("xl/theme/theme1.xml", THEME);
  zip.file("xl/styles.xml", STYLES);
  zip.file("xl/sharedStrings.xml", SHARED_STRINGS);
  zip.file("xl/worksheets/sheet1.xml", SHEET1);
  zip.file("xl/worksheets/sheet2.xml", SHEET2);
  zip.file("xl/externalLinks/externalLink1.xml", EXTERNAL_LINK);
  zip.file("xl/externalLinks/_rels/externalLink1.xml.rels", EXTERNAL_LINK_RELS);
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}

// 单独运行产夹具：`node --experimental-strip-types scripts/make-sample-sheet.ts /tmp/sample-sheet.xlsx`
if (process.argv[1]?.includes("make-sample-sheet")) {
  const target = process.argv[2] ?? "/tmp/sample-sheet.xlsx";
  const bytes = await buildSampleWorkbookBytes();
  await writeFile(target, bytes);
  console.log(`夹具已写入 ${target}（${bytes.byteLength} 字节）`);
}
