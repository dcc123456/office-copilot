/**
 * 一份最小可用 `.pptx` 的构造器，给测试用（`pptx-io` 自己的单测与桌面主进程出入口的单测共用，
 * `AGENTS.md` §3.2：同一份夹具不许长两处）。
 *
 * 对外只经 `@office-copilot/pptx-io/testing` 这个子路径露出（见 `package.json` 的 `exports`），
 * 不进 `index.ts`——生产入口带一份假幻灯片是纯死重。
 *
 * 它只放"读侧与审计要看的那几件"：内容类型表、包关系、`presentation.xml` 与它的关系表、
 * 一张带 `cSld > spTree` 与版式关系的幻灯片、那张版式，外加一个二进制部件用来盯住"二进制不许被改坏"。
 * {@link minimalDeckParts} 没有主题/母版/备注——读侧对缺失部件的容忍由 `pptx.test.ts` 另测；
 * 要测母版/版式那条继承链用 {@link inheritanceDeckParts}（它带一张母版）。
 */
import JSZip from "jszip";

/** 部件路径 → 文本；想**留**某个部件（例如故意不给幻灯片关系表）就把它的值传 `undefined`。 */
export type DeckParts = Record<string, string | undefined>;

/**
 * 造一份 `.pptx` 字节。
 *
 * @param parts 要放进包里的文本部件；省略时用 {@link minimalDeckParts}，即一份结构合规的最小 deck。
 * @returns zip 字节，可直接喂 `PptxFile.open` / `auditPptxPackage`。
 */
export async function buildDeckBytes(parts: DeckParts = minimalDeckParts()): Promise<Uint8Array> {
  const zip = new JSZip();
  for (const [path, text] of Object.entries(parts)) {
    if (text !== undefined) zip.file(path, text);
  }
  zip.file("ppt/printerSettings/printerSettings1.bin", new Uint8Array([1, 2, 3, 4, 5]));
  return zip.generateAsync({ type: "uint8array" });
}

/** 一份"结构没问题"的夹具：内容类型表覆盖全部部件，幻灯片有 `cSld > spTree` 与版式关系，标题形状 `id=2` 文字"标题甲"。 */
export function minimalDeckParts(): DeckParts {
  return {
    "[Content_Types].xml":
      '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.presentationml.printerSettings"/>' +
      '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>' +
      '<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>' +
      '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>' +
      "</Types>",
    "_rels/.rels":
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>' +
      "</Relationships>",
    "ppt/presentation.xml":
      '<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
      '<p:sldIdLst><p:sldId id="256" r:id="rId7"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/></p:presentation>',
    "ppt/_rels/presentation.xml.rels":
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>' +
      "</Relationships>",
    "ppt/slides/slide1.xml":
      '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
      '<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>' +
      '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/>' +
      "<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>标题甲</a:t></a:r></a:p></p:txBody></p:sp>" +
      "</p:spTree></p:cSld></p:sld>",
    "ppt/slides/_rels/slide1.xml.rels":
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>' +
      "</Relationships>",
    "ppt/slideLayouts/slideLayout1.xml":
      '<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
      '<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sldLayout>',
  };
}

/**
 * 一页里放两个形状（`id=2` 标题在下、`id=3` 文本框在上）的 deck，给叠放次序的用例用。
 *
 * 最小夹具只有一形状，"最顶/最底"永远重合，看不出重排有没有落地，所以另开一份而不是改那份
 * （改它会把字节级往返那几条用例一起牵动）。
 *
 * @returns 直接传给 {@link buildDeckBytes} 的部件表。
 */
export function twoShapeDeckParts(): DeckParts {
  return {
    ...minimalDeckParts(),
    "ppt/slides/slide1.xml":
      '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
      '<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>' +
      '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/>' +
      "<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>标题甲</a:t></a:r></a:p></p:txBody></p:sp>" +
      '<p:sp><p:nvSpPr><p:cNvPr id="3" name="TextBox 2"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr/>' +
      "<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>正文乙</a:t></a:r></a:p></p:txBody></p:sp>" +
      "</p:spTree></p:cSld></p:sld>",
  };
}

/**
 * 一份带**完整继承链**的 deck（幻灯片 → 版式 → 母版），给 `inherit.ts` 与读模型的 `effectiveBoxPx` 用。
 *
 * 三格各占一条链上的不同情形，EMU 值都挑成整 px（9525 EMU/px，见 `units.ts`）方便断言：
 * - 标题 `ctrTitle`（`id=2`）：幻灯片与版式都写 `<p:spPr/>`，只有母版的 `title` 写了四数 → 四轴全部来自母版，
 *   顺便测版式→母版的类型映射（`ctrTitle` 归到 `title`）。
 * - 正文 `body idx="1"`（`id=3`）：幻灯片自己写全四数（100/120/300/80），版式那格也写了另一套 → 自己写的必须赢。
 * - 文本框（`id=4`，无 `<p:ph>`）：写全四数（50/50/200/100），版式根本没有对应格子 → 非占位符不参与继承。
 * 母版里有 `title`（96/48/768/120）与 `body`（80/160/400/200）两格：前者给 `ctrTitle` 映射用，
 * 后者给 `subTitle`/`obj` → `body` 那档映射用（用例是改版式那一格的类型来走这两条）。
 *
 * 为什么每层的 `<a:xfrm>` 都是 `off` + `ext` 一整对：`CT_Transform2D` 把两个子元素都定为必选，
 * 本仓的审计（`audit.ts`）会报 `slide-incomplete-transform`，夹具不该给自己造一份"打开就判损坏"的输入。
 *
 * @returns 直接传给 {@link buildDeckBytes} 的部件表。
 */
export function inheritanceDeckParts(): DeckParts {
  return {
    ...minimalDeckParts(),
    "[Content_Types].xml":
      '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.presentationml.printerSettings"/>' +
      '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>' +
      '<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>' +
      '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>' +
      '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>' +
      "</Types>",
    "ppt/_rels/presentation.xml.rels":
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>' +
      "</Relationships>",
    "ppt/presentation.xml":
      '<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
      '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>' +
      '<p:sldIdLst><p:sldId id="256" r:id="rId7"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/></p:presentation>',
    "ppt/slides/slide1.xml":
      '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
      '<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>' +
      '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr/><p:nvPr><p:ph type="ctrTitle"/></p:nvPr></p:nvSpPr><p:spPr/>' +
      "<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>标题甲</a:t></a:r></a:p></p:txBody></p:sp>" +
      '<p:sp><p:nvSpPr><p:cNvPr id="3" name="Body 2"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>' +
      '<p:spPr><a:xfrm><a:off x="952500" y="1143000"/><a:ext cx="2857500" cy="762000"/></a:xfrm></p:spPr>' +
      "<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>正文乙</a:t></a:r></a:p></p:txBody></p:sp>" +
      '<p:sp><p:nvSpPr><p:cNvPr id="4" name="TextBox 3"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>' +
      '<p:spPr><a:xfrm><a:off x="476250" y="476250"/><a:ext cx="1905000" cy="952500"/></a:xfrm></p:spPr>' +
      "<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>正文丙</a:t></a:r></a:p></p:txBody></p:sp>" +
      "</p:spTree></p:cSld></p:sld>",
    "ppt/slideLayouts/slideLayout1.xml":
      '<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="titleOnly" name="Title Only">' +
      '<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>' +
      '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr/><p:nvPr><p:ph type="ctrTitle"/></p:nvPr></p:nvSpPr><p:spPr/>' +
      "<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody></p:sp>" +
      '<p:sp><p:nvSpPr><p:cNvPr id="3" name="Content Placeholder 2"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>' +
      '<p:spPr><a:xfrm><a:off x="685800" y="1828800"/><a:ext cx="4572000" cy="2743200"/></a:xfrm></p:spPr>' +
      "<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody></p:sp>" +
      "</p:spTree></p:cSld>" +
      '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>',
    "ppt/slideLayouts/_rels/slideLayout1.xml.rels":
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>' +
      "</Relationships>",
    "ppt/slideMasters/slideMaster1.xml":
      '<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
      '<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>' +
      '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>' +
      '<p:spPr><a:xfrm><a:off x="914400" y="457200"/><a:ext cx="7315200" cy="1143000"/></a:xfrm></p:spPr>' +
      "<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody></p:sp>" +
      '<p:sp><p:nvSpPr><p:cNvPr id="3" name="Body 2"/><p:cNvSpPr/><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr>' +
      '<p:spPr><a:xfrm><a:off x="762000" y="1524000"/><a:ext cx="3810000" cy="1905000"/></a:xfrm></p:spPr>' +
      "<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody></p:sp>" +
      "</p:spTree></p:cSld>" +
      '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>' +
      "<p:sldLayoutIdLst><p:sldLayoutId id=\"2147483649\" r:id=\"rId1\"/></p:sldLayoutIdLst></p:sldMaster>",
    "ppt/slideMasters/_rels/slideMaster1.xml.rels":
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>' +
      "</Relationships>",
  };
}

/**
 * 一份把**字号继承链五档**各占一格的 deck（在 {@link inheritanceDeckParts} 的链上加母版 `p:txStyles`），
 * 给 `text-styles.ts` 与读模型的 `effectiveSizePt` 用。
 *
 * 每一档挑一个互不相同的数，且**下层也全都写了值**——这样"少问一层"或"问错层"都会答出一个看得出错在哪的数字，
 * 而不是恰好撞对：
 * - `id=2` 标题（`ph type="ctrTitle"`）两段：段一自己写了 `a:pPr/a:defRPr sz="2000"` → **paragraph** 20pt；
 *   段二什么都没写，落到版式那一格的 `a:lstStyle` `sz="4000"` → **layout** 40pt（母版 `titleStyle` 写的 44pt 必须输）。
 * - `id=3` 正文（`body idx="1"`）：版式那一格 `lstStyle` 是空的，母版 `body` 那格写了 `sz="2400"` → **master-list** 24pt
 *   （母版 `bodyStyle` 的 32pt 必须输）；另一段 run 自己写了 `sz="1000"` → **run** 10pt，`sizePt` 同时有值（写侧靠它）。
 * - `id=4` 文本框（无 `<p:ph>`）自己写了 `a:lstStyle` `sz="1400"` → **shape-list** 14pt；
 *   它第二段是 `lvl="1"`，而链上只写了 `lvl1pPr` 那一档 → 两格都必须 undefined（**不许拿 18pt 凑**）。
 * - `id=5` 文本框（无 `<p:ph>`、没写 `lstStyle`）整条链只剩母版 `p:txStyles/otherStyle` `sz="1800"` → **master-text-styles** 18pt。
 *
 * `p:txStyles` 的元素名与位置按真文件核过（`AGENTS.md` §6.2）：本机 python-pptx 默认模板与 WPS 模板的母版
 * 顶层子元素顺序都是 `cSld`、`clrMap`、`sldLayoutIdLst`、`txStyles`，三档都**不带** `@phType` / `@styleId`。
 *
 * @returns 直接传给 {@link buildDeckBytes} 的部件表。
 */
export function textStyleDeckParts(): DeckParts {
  return {
    ...inheritanceDeckParts(),
    "ppt/slides/slide1.xml":
      '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
      '<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>' +
      '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr/><p:nvPr><p:ph type="ctrTitle"/></p:nvPr></p:nvSpPr><p:spPr/>' +
      "<p:txBody><a:bodyPr/><a:lstStyle/>" +
      '<a:p><a:pPr><a:defRPr sz="2000"/></a:pPr><a:r><a:t>段一</a:t></a:r></a:p>' +
      "<a:p><a:r><a:t>段二</a:t></a:r></a:p>" +
      "</p:txBody></p:sp>" +
      '<p:sp><p:nvSpPr><p:cNvPr id="3" name="Body 2"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>' +
      '<p:spPr><a:xfrm><a:off x="952500" y="1143000"/><a:ext cx="2857500" cy="762000"/></a:xfrm></p:spPr>' +
      "<p:txBody><a:bodyPr/><a:lstStyle/>" +
      "<a:p><a:r><a:t>正文甲</a:t></a:r></a:p>" +
      '<a:p><a:r><a:rPr sz="1000"/><a:t>正文乙</a:t></a:r></a:p>' +
      "</p:txBody></p:sp>" +
      '<p:sp><p:nvSpPr><p:cNvPr id="4" name="TextBox 3"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>' +
      '<p:spPr><a:xfrm><a:off x="476250" y="476250"/><a:ext cx="1905000" cy="952500"/></a:xfrm></p:spPr>' +
      "<p:txBody><a:bodyPr/>" +
      '<a:lstStyle><a:lvl1pPr><a:defRPr sz="1400"/></a:lvl1pPr></a:lstStyle>' +
      "<a:p><a:r><a:t>文本一</a:t></a:r></a:p>" +
      '<a:p><a:pPr lvl="1"/><a:r><a:t>文本二</a:t></a:r></a:p>' +
      "</p:txBody></p:sp>" +
      '<p:sp><p:nvSpPr><p:cNvPr id="5" name="TextBox 4"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>' +
      '<p:spPr><a:xfrm><a:off x="476250" y="1714500"/><a:ext cx="1905000" cy="952500"/></a:xfrm></p:spPr>' +
      "<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>文本三</a:t></a:r></a:p></p:txBody></p:sp>" +
      "</p:spTree></p:cSld></p:sld>",
    // 版式：`ctrTitle` 那一格写 40pt（layout 档），`body idx="1"` 那格留空（让链继续问母版那一格）。
    "ppt/slideLayouts/slideLayout1.xml":
      '<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="titleOnly" name="Title Only">' +
      '<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>' +
      '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr/><p:nvPr><p:ph type="ctrTitle"/></p:nvPr></p:nvSpPr><p:spPr/>' +
      "<p:txBody><a:bodyPr/>" +
      '<a:lstStyle><a:lvl1pPr><a:defRPr sz="4000"/></a:lvl1pPr></a:lstStyle>' +
      "<a:p><a:endParaRPr/></a:p></p:txBody></p:sp>" +
      '<p:sp><p:nvSpPr><p:cNvPr id="3" name="Content Placeholder 2"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>' +
      '<p:spPr><a:xfrm><a:off x="685800" y="1828800"/><a:ext cx="4572000" cy="2743200"/></a:xfrm></p:spPr>' +
      "<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody></p:sp>" +
      "</p:spTree></p:cSld>" +
      '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>',
    // 母版：`body` 那格的 `lstStyle` 写 24pt（master-list 档），末尾 `p:txStyles` 三档各写一个数（最后一档）。
    "ppt/slideMasters/slideMaster1.xml":
      '<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
      '<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>' +
      '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>' +
      '<p:spPr><a:xfrm><a:off x="914400" y="457200"/><a:ext cx="7315200" cy="1143000"/></a:xfrm></p:spPr>' +
      "<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody></p:sp>" +
      '<p:sp><p:nvSpPr><p:cNvPr id="3" name="Body 2"/><p:cNvSpPr/><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr>' +
      '<p:spPr><a:xfrm><a:off x="762000" y="1524000"/><a:ext cx="3810000" cy="1905000"/></a:xfrm></p:spPr>' +
      "<p:txBody><a:bodyPr/>" +
      '<a:lstStyle><a:lvl1pPr><a:defRPr sz="2400"/></a:lvl1pPr></a:lstStyle>' +
      "<a:p><a:endParaRPr/></a:p></p:txBody></p:sp>" +
      "</p:spTree></p:cSld>" +
      '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>' +
      "<p:sldLayoutIdLst><p:sldLayoutId id=\"2147483649\" r:id=\"rId1\"/></p:sldLayoutIdLst>" +
      "<p:txStyles>" +
      '<p:titleStyle><a:lvl1pPr><a:defRPr sz="4400"/></a:lvl1pPr></p:titleStyle>' +
      '<p:bodyStyle><a:lvl1pPr><a:defRPr sz="3200"/></a:lvl1pPr><a:lvl2pPr><a:defRPr sz="2800"/></a:lvl2pPr></p:bodyStyle>' +
      '<p:otherStyle><a:defPPr/><a:lvl1pPr><a:defRPr sz="1800"/></a:lvl1pPr></p:otherStyle>' +
      "</p:txStyles></p:sldMaster>",
  };
}

/**
 * 一份带**主题色**的 deck（在 {@link inheritanceDeckParts} 的继承链上再接一层 `theme1.xml`），给 `theme.ts` 的解色用例用。
 *
 * 色板刻意挑成一眼能核对的三个值，并按 `p:clrMap` 的映射错位一次，防止"按同名回退"蒙对：
 * - `lt1` = `#ffffff`、`dk1` = `#000000`（且用 `a:sysClr @lastClr` 写，测 `sysClr` 这条路）、`accent1` = `#ff0000`。
 * - 母版写的是 `bg1="lt1" tx1="dk1" accent1="accent3"`：所以 `schemeClr val="accent1"` 实际落在 `#ffffff`，
 *   `val="tx1"` 落在 `#000000`。
 * - 四条 run 各占一种情形：`accent1`（无变换，能解）、`tx1`（走 `dk1`）、带 `a:lumMod`/`a:lumOff` 的 `accent1`
 *   与 `val="phClr"`（都**不许**给 hex）、`srgbClr`（不经主题，直接要）。
 *
 * 主题部件是新部件，所以本体、母版的关系、`[Content_Types].xml` 的 Override 三处一起登记（§2.10）。
 *
 * @param withTheme 传 false 时整条主题链都不给（母版关系与部件都缺），用来断言"链路断了就不猜色"。
 * @returns 直接传给 {@link buildDeckBytes} 的部件表。
 */
export function colorDeckParts(withTheme = true): DeckParts {
  const parts: DeckParts = {
    ...inheritanceDeckParts(),
    "ppt/slides/slide1.xml":
      '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
      '<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>' +
      '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr/><p:nvPr><p:ph type="ctrTitle"/></p:nvPr></p:nvSpPr><p:spPr/>' +
      "<p:txBody><a:bodyPr/><a:lstStyle/><a:p>" +
      '<a:r><a:rPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:rPr><a:t>标题甲</a:t></a:r>' +
      '<a:r><a:rPr><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:rPr><a:t>标题乙</a:t></a:r>' +
      "</a:p></p:txBody></p:sp>" +
      '<p:sp><p:nvSpPr><p:cNvPr id="3" name="Body 2"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>' +
      '<p:spPr><a:xfrm><a:off x="952500" y="1143000"/><a:ext cx="2857500" cy="762000"/></a:xfrm></p:spPr>' +
      "<p:txBody><a:bodyPr/><a:lstStyle/><a:p>" +
      '<a:r><a:rPr><a:solidFill><a:schemeClr val="accent1"><a:lumMod val="60000"/><a:lumOff val="40000"/></a:schemeClr></a:solidFill></a:rPr><a:t>正文甲</a:t></a:r>' +
      '<a:r><a:rPr><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:rPr><a:t>正文乙</a:t></a:r>' +
      '<a:r><a:rPr><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:rPr><a:t>正文丙</a:t></a:r>' +
      "</a:p></p:txBody></p:sp>" +
      "</p:spTree></p:cSld></p:sld>",
  };
  // 把 `accent1` 这一档故意映射到 `accent3`（#ffffff），这样"按同名回退"会解成 #ff0000，一眼看出映射有没有真的走。
  parts["ppt/slideMasters/slideMaster1.xml"] = (parts["ppt/slideMasters/slideMaster1.xml"] ?? "").replace(
    'accent1="accent1"',
    'accent1="accent3"',
  );
  if (!withTheme) return parts;
  parts["[Content_Types].xml"] = (parts["[Content_Types].xml"] ?? "").replace(
    "</Types>",
    '<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>',
  );
  parts["ppt/slideMasters/_rels/slideMaster1.xml.rels"] =
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>' +
    '<Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>' +
    "</Relationships>";
  parts["ppt/theme/theme1.xml"] =
    '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office">' +
    '<a:themeElements><a:clrScheme name="Office">' +
    '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>' +
    '<a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>' +
    '<a:dk2><a:srgbClr val="44546A"/></a:dk2>' +
    '<a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>' +
    '<a:accent1><a:srgbClr val="FF0000"/></a:accent1>' +
    '<a:accent2><a:srgbClr val="ED7D31"/></a:accent2>' +
    '<a:accent3><a:srgbClr val="FFFFFF"/></a:accent3>' +
    '<a:accent4><a:srgbClr val="70AD47"/></a:accent4>' +
    '<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>' +
    '<a:accent6><a:srgbClr val="7030A0"/></a:accent6>' +
    '<a:hlink><a:srgbClr val="0563C1"/></a:hlink>' +
    '<a:folHlink><a:srgbClr val="954F72"/></a:folHlink>' +
    "</a:clrScheme></a:themeElements></a:theme>";
  return parts;
}
