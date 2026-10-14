/**
 * 一份最小可用 `.pptx` 的构造器，给测试用（`pptx-io` 自己的单测与桌面主进程出入口的单测共用，
 * `AGENTS.md` §3.2：同一份夹具不许长两处）。
 *
 * 对外只经 `@office-copilot/pptx-io/testing` 这个子路径露出（见 `package.json` 的 `exports`），
 * 不进 `index.ts`——生产入口带一份假幻灯片是纯死重。
 *
 * 它只放"读侧与审计要看的那几件"：内容类型表、包关系、`presentation.xml` 与它的关系表、
 * 一张带 `cSld > spTree` 与版式关系的幻灯片、那张版式，外加一个二进制部件用来盯住"二进制不许被改坏"。
 * 没有主题/母版/备注——读侧对缺失部件的容忍由 `pptx.test.ts` 另测。
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
