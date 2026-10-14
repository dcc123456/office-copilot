import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { auditPptxPackage } from "./audit.js";
import { PptxFile } from "./pptx-file.js";
import { applySlideEdit, readPresentation, setShapeRunText } from "./index.js";

/**
 * 造一份最小 `.pptx` 字节（只放审计会看的部件，别的一个都不带）。
 *
 * @param parts 部件路径 → 文本；想留某个部件（例如故意不给幻灯片关系表）就传 `undefined`。
 * @returns zip 字节，可直接喂 `auditPptxPackage`。
 */
async function deckBytes(parts: Record<string, string | undefined>): Promise<Uint8Array> {
  const zip = new JSZip();
  for (const [path, text] of Object.entries(parts)) {
    if (text !== undefined) zip.file(path, text);
  }
  zip.file("ppt/printerSettings/printerSettings1.bin", new Uint8Array([1, 2, 3, 4, 5]));
  return zip.generateAsync({ type: "uint8array" });
}

/** 一份"结构没问题"的夹具：内容类型表覆盖全部部件，幻灯片有 `cSld > spTree` 与版式关系。 */
function okParts(): Record<string, string | undefined> {
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
      "<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>" +
      '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/>' +
      "<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>标题甲</a:t></a:r></a:p></p:txBody></p:sp>" +
      "</p:spTree></p:cSld></p:sld>",
    "ppt/slides/_rels/slide1.xml.rels":
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>' +
      "</Relationships>",
    "ppt/slideLayouts/slideLayout1.xml":
      '<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
      "<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sldLayout>",
  };
}

describe("auditPptxPackage：PowerPoint 会判损坏的几类结构问题", () => {
  it("结构合规的包返回空数组（不许误报）", async () => {
    expect(await auditPptxPackage(await deckBytes(okParts()))).toEqual([]);
  });

  it("通用那条确实在跑：部件没被 [Content_Types] 覆盖时报一条", async () => {
    const parts = { ...okParts() };
    parts["ppt/media/picture.png"] = "not-a-real-png";
    const issues = await auditPptxPackage(await deckBytes(parts));
    expect(issues).toEqual([{ code: "part-not-declared", part: "ppt/media/picture.png", detail: "png" }]);
  });

  it("幻灯片只有 cSld 空壳、没有 spTree 时报出", async () => {
    const parts = { ...okParts(), "ppt/slides/slide1.xml": '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld/></p:sld>' };
    expect(await auditPptxPackage(await deckBytes(parts))).toEqual([
      { code: "slide-missing-shape-tree", part: "ppt/slides/slide1.xml", detail: "" },
    ]);
  });

  it("幻灯片连关系表都没有时，报缺版式关系（PowerPoint 侧的必须关联）", async () => {
    const parts = { ...okParts(), "ppt/slides/_rels/slide1.xml.rels": undefined };
    expect(await auditPptxPackage(await deckBytes(parts))).toEqual([
      { code: "slide-missing-layout-relationship", part: "ppt/slides/slide1.xml", detail: "" },
    ]);
  });

  it("版式关系指向不存在的部件时，由通用那条报 relationship-target-missing", async () => {
    const parts = {
      ...okParts(),
      "ppt/slides/_rels/slide1.xml.rels":
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout9.xml"/>' +
        "</Relationships>",
    };
    expect(await auditPptxPackage(await deckBytes(parts))).toEqual([
      { code: "relationship-target-missing", part: "ppt/slides/slide1.xml", detail: "rId1 -> ppt/slideLayouts/slideLayout9.xml" },
    ]);
  });
});

/**
 * 我们自己写出的包必须过审计——这条是"改过的 `.pptx` 能在 PowerPoint 里直接打开"的机检面。
 *
 * 本机装不了 PowerPoint（`AGENTS.md` §8），python-pptx 只读值、Keynote 只会说打不开，
 * 所以拿同一套判定回头看两种产物：原始夹具、改过幻灯片文字后的写回。
 */
describe("我们自己的写入产物过审计（不许给自己判损坏）", () => {
  /** 打开夹具、把标题那格改成"改过的标题"，走桌面的保存路径出字节。 */
  async function bytesAfterTextEdit(): Promise<Uint8Array> {
    const bytes = await deckBytes(okParts());
    const file = await PptxFile.open(bytes);
    const { model } = await readPresentation(bytes);
    const partPath = model.slides[0]!.partPath;
    const made = await setShapeRunText(file, { partPath, shapeId: 2, paraIndex: 0, runIndex: 0 }, "改过的标题");
    if (!made.ok) throw new Error(JSON.stringify(made));
    applySlideEdit(file, made.op);
    return file.toBytes();
  }

  it.each([
    ["原始夹具", async () => deckBytes(okParts())],
    ["改过幻灯片文字的写回", bytesAfterTextEdit],
  ] as const)("%s：审计返回空数组", async (_label, produce) => {
    expect(await auditPptxPackage(await produce())).toEqual([]);
  });
});
