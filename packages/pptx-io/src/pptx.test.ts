import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { PptxFile } from "./pptx-file.js";
import { readPresentation } from "./read.js";

/**
 * 造一份最小 `.pptx` 字节，部件形状照抄本机 python-pptx 产物（spec §取证）。
 *
 * @returns zip 的 Uint8Array。含一张 4:3 幻灯片、三种形状（标题占位符、正文占位符、自带几何的文本框）
 *   与一个二进制部件，用来同时验读模型与保真回写。
 */
async function minimalDeck(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>',
  );
  zip.file(
    "_rels/.rels",
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>',
  );
  zip.file(
    "ppt/presentation.xml",
    '<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldIdLst><p:sldId id="256" r:id="rId7"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000" type="screen4x3"/></p:presentation>',
  );
  zip.file(
    "ppt/_rels/presentation.xml.rels",
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>',
  );
  zip.file(
    "ppt/slides/slide1.xml",
    '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>标题甲</a:t></a:r></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Content Placeholder 2"/><p:cNvSpPr/><p:nvPr><p:ph idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>副标题乙</a:t></a:r></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="4" name="TextBox 3"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="914400" y="3657600"/><a:ext cx="2743200" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>正文丙</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>',
  );
  zip.file("ppt/printerSettings/printerSettings1.bin", new Uint8Array([1, 2, 3, 4, 5]));
  return zip.generateAsync({ type: "uint8array" });
}

describe("pptx-io：读模型", () => {
  it("读页面尺寸（EMU 原样，不换算）与按 sldIdLst 顺序的幻灯片", async () => {
    const { model } = await readPresentation(await minimalDeck());
    expect(model.slideSizeEmu).toEqual({ cx: 9144000, cy: 6858000 });
    expect(model.slides).toHaveLength(1);
    expect(model.slides[0]!.order).toBe(0);
    expect(model.slides[0]!.slideId).toBe(256);
    expect(model.slides[0]!.partPath).toBe("ppt/slides/slide1.xml");
  });

  it("占位符继承版式几何（hasOwnGeometry=false），文本框自带几何（true）", async () => {
    const { model } = await readPresentation(await minimalDeck());
    const shapes = model.slides[0]!.shapes;
    expect(shapes.map((s) => s.name)).toEqual(["Title 1", "Content Placeholder 2", "TextBox 3"]);
    const [title, body, textbox] = shapes;
    expect(title!.isPlaceholder).toBe(true);
    expect(title!.placeholderRef).toBe("title");
    expect(title!.hasOwnGeometry).toBe(false); // 空 <p:spPr/>：几何从版式继承，不许烘焙（D3）
    expect(title!.paragraphTexts).toEqual(["标题甲"]);
    expect(body!.isPlaceholder).toBe(true);
    expect(body!.placeholderRef).toBe("1"); // idx="1" 而非 type
    expect(textbox!.isPlaceholder).toBe(false);
    expect(textbox!.hasOwnGeometry).toBe(true);
    expect(textbox!.paragraphTexts).toEqual(["正文丙"]);
  });
});

describe("pptx-io：保真回写", () => {
  it("没改任何部件时，解包后每个部件内容与原文件逐字一致（含二进制）", async () => {
    const bytes = await minimalDeck();
    const out = await (await PptxFile.open(bytes)).toBytes();
    const before = await JSZip.loadAsync(bytes);
    const after = await JSZip.loadAsync(out);
    expect(after.file("ppt/slides/slide1.xml") && before.file("ppt/slides/slide1.xml")).toBeTruthy();
    expect(await after.file("ppt/slides/slide1.xml")!.async("string")).toBe(await before.file("ppt/slides/slide1.xml")!.async("string"));
    expect(await after.file("ppt/printerSettings/printerSettings1.bin")!.async("uint8array")).toEqual(await before.file("ppt/printerSettings/printerSettings1.bin")!.async("uint8array"));
  });

  it("改一张幻灯片只标记账号那一个部件，空 <p:spPr/> 仍是空的", async () => {
    const file = await PptxFile.open(await minimalDeck());
    const slidePath = file.model.slides[0]!.partPath;
    const xml = (await file.partText(slidePath))!;
    file.writePartText(slidePath, xml.replace("<a:t>标题甲</a:t>", "<a:t>标题乙</a:t>"));
    expect(file.changedParts).toEqual([slidePath]);
    const written = await file.partText(slidePath);
    expect(written).toContain("<a:t>标题乙</a:t>");
    expect(written).toContain("<p:spPr/>"); // 回写没把占位符的空几何烘成实数（D3）
  });
});
