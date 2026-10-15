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
