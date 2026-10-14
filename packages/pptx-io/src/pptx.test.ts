import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { PptxFile } from "./pptx-file.js";
import { applySlideEdit, readPresentation, revertSlideEdit, setRunColor, setRunFont, setShapeRunText } from "./index.js";

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

/**
 * 造一份标题段里有**两个不同格式 run** 的幻灯片，用来验编辑只动目标 run、不拍平邻居（§2.3）。
 *
 * @returns zip 字节；标题形状 id=2，段内 run0 粗体"标题甲"、run1 sz=1800"副"。
 */
async function twoRunDeck(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>');
  zip.file("_rels/.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>');
  zip.file("ppt/presentation.xml", '<p:presentation xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldIdLst><p:sldId id="256" r:id="rId7"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/></p:presentation>');
  zip.file("ppt/_rels/presentation.xml.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>');
  zip.file("ppt/slides/slide1.xml", '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr b="1"/><a:t>标题甲</a:t></a:r><a:r><a:rPr sz="1800"/><a:t>副</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>');
  return zip.generateAsync({ type: "uint8array" });
}

describe("pptx-io：结构化文字编辑 op + 撤销", () => {
  it("产出的 op 带前后像，apply 后部件变、revert 后逐字回到原样", async () => {
    const file = await PptxFile.open(await minimalDeck());
    const partPath = file.model.slides[0]!.partPath;
    const original = (await file.partText(partPath))!;
    const result = await setShapeRunText(file, { partPath, shapeId: 2, paraIndex: 0, runIndex: 0 }, "新标题");
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.op.beforeXml).toBe(original);
    expect(result.op.afterXml).toContain("<a:t>新标题</a:t>");
    // 提案阶段不落地：部件文本仍未变
    expect((await file.partText(partPath))! ).toBe(original);

    applySlideEdit(file, result.op);
    expect((await file.partText(partPath))!).toContain("<a:t>新标题</a:t>");
    expect(file.changedParts).toEqual([partPath]);

    revertSlideEdit(file, result.op);
    expect((await file.partText(partPath))!).toBe(original); // 撤销逐字回原
  });

  it("改 run0 不动 run1 的 rPr，局部格式不被拍平（§2.3）", async () => {
    const file = await PptxFile.open(await twoRunDeck());
    const partPath = file.model.slides[0]!.partPath;
    const result = await setShapeRunText(file, { partPath, shapeId: 2, paraIndex: 0, runIndex: 0 }, "改后");
    if (!result.ok) throw new Error(result.reason);
    applySlideEdit(file, result.op);
    const after = (await file.partText(partPath))!;
    expect(after).toContain("<a:rPr b=\"1\"/><a:t>改后</a:t>"); // 目标 run 文本变、其 rPr 留住
    expect(after).toContain("<a:rPr sz=\"1800\"/><a:t>副</a:t>"); // 邻居 run 的格式与文字一个没动
  });

  it("越界或图片形状回带原因的失败而不是抛异常", async () => {
    const file = await PptxFile.open(await minimalDeck());
    const partPath = file.model.slides[0]!.partPath;
    const badPara = await setShapeRunText(file, { partPath, shapeId: 2, paraIndex: 5, runIndex: 0 }, "x");
    expect(badPara.ok).toBe(false);
    const badShape = await setShapeRunText(file, { partPath, shapeId: 999, paraIndex: 0, runIndex: 0 }, "x");
    expect(badShape.ok).toBe(false);
    if (!badShape.ok) expect(badShape.reason).toContain("找不到形状");
  });
});

describe("pptx-io：字体编辑 op（粗体 / 字号）", () => {
  it("给已有 rPr 的 run 加字号，pt×100 落进 sz，邻居 run 的 b 不受影响（§2.1/§2.3）", async () => {
    const file = await PptxFile.open(await twoRunDeck());
    const partPath = file.model.slides[0]!.partPath;
    const result = await setRunFont(file, { partPath, shapeId: 2, paraIndex: 0, runIndex: 0 }, { sizePt: 24 });
    if (!result.ok) throw new Error(result.reason);
    applySlideEdit(file, result.op);
    const after = (await file.partText(partPath))!;
    expect(after).toContain("<a:rPr b=\"1\" sz=\"2400\"/><a:t>标题甲</a:t>"); // 24pt → sz=2400，原 b=1 留住
    expect(after).toContain("<a:rPr sz=\"1800\"/><a:t>副</a:t>"); // 邻居 run 一点没动
  });

  it("取消粗体写 b=0 而不是删属性，apply 后 revert 逐字回原", async () => {
    const file = await PptxFile.open(await twoRunDeck());
    const partPath = file.model.slides[0]!.partPath;
    const original = (await file.partText(partPath))!;
    const result = await setRunFont(file, { partPath, shapeId: 2, paraIndex: 0, runIndex: 0 }, { bold: false });
    if (!result.ok) throw new Error(result.reason);
    applySlideEdit(file, result.op);
    expect((await file.partText(partPath))!).toContain("<a:rPr b=\"0\"/><a:t>标题甲</a:t>");
    revertSlideEdit(file, result.op);
    expect((await file.partText(partPath))!).toBe(original);
  });

  it("run 没有 rPr 时新建 <a:rPr/> 插到 <a:t> 之前", async () => {
    const file = await PptxFile.open(await minimalDeck()); // 标题 run 是 <a:r><a:t>标题甲</a:t></a:r>，无 rPr
    const partPath = file.model.slides[0]!.partPath;
    const result = await setRunFont(file, { partPath, shapeId: 2, paraIndex: 0, runIndex: 0 }, { bold: true, sizePt: 18 });
    if (!result.ok) throw new Error(result.reason);
    applySlideEdit(file, result.op);
    expect((await file.partText(partPath))!).toContain("<a:r><a:rPr b=\"1\" sz=\"1800\"/><a:t>标题甲</a:t></a:r>");
  });

  it("没给任何字体属性或字号非正数时带原因拒绝", async () => {
    const file = await PptxFile.open(await minimalDeck());
    const partPath = file.model.slides[0]!.partPath;
    const empty = await setRunFont(file, { partPath, shapeId: 2, paraIndex: 0, runIndex: 0 }, {});
    expect(empty.ok).toBe(false);
    const badSize = await setRunFont(file, { partPath, shapeId: 2, paraIndex: 0, runIndex: 0 }, { sizePt: 0 });
    expect(badSize.ok).toBe(false);
    if (!badSize.ok) expect(badSize.reason).toContain("字号");
  });
});

describe("pptx-io：字体颜色 op", () => {
  it("给已有 rPr 的 run 上色，solidFill 落进 rPr，b 属性与邻居 run 都不动（§2.2/§2.3）", async () => {
    const file = await PptxFile.open(await twoRunDeck());
    const partPath = file.model.slides[0]!.partPath;
    const result = await setRunColor(file, { partPath, shapeId: 2, paraIndex: 0, runIndex: 0 }, "#ff8800");
    if (!result.ok) throw new Error(result.reason);
    applySlideEdit(file, result.op);
    const after = (await file.partText(partPath))!;
    expect(after).toContain("<a:rPr b=\"1\"><a:solidFill><a:srgbClr val=\"FF8800\"/></a:solidFill></a:rPr><a:t>标题甲</a:t>");
    expect(after).toContain("<a:rPr sz=\"1800\"/><a:t>副</a:t>"); // 邻居 run 一点没动
  });

  it("重复上色只留一个 solidFill，不叠加第二个填充", async () => {
    const file = await PptxFile.open(await twoRunDeck());
    const partPath = file.model.slides[0]!.partPath;
    const first = await setRunColor(file, { partPath, shapeId: 2, paraIndex: 0, runIndex: 0 }, "#000000");
    if (!first.ok) throw new Error(first.reason);
    applySlideEdit(file, first.op);
    const second = await setRunColor(file, { partPath, shapeId: 2, paraIndex: 0, runIndex: 0 }, "#ffffff");
    if (!second.ok) throw new Error(second.reason);
    applySlideEdit(file, second.op);
    const after = (await file.partText(partPath))!;
    expect(after.match(/<a:solidFill>/g)?.length).toBe(1); // 只有一个填充
    expect(after).toContain("<a:srgbClr val=\"FFFFFF\"/>");
  });

  it("run 没有 rPr 时新建 rPr 再放 solidFill", async () => {
    const file = await PptxFile.open(await minimalDeck()); // 标题 run 无 rPr
    const partPath = file.model.slides[0]!.partPath;
    const result = await setRunColor(file, { partPath, shapeId: 2, paraIndex: 0, runIndex: 0 }, "AABBCC");
    if (!result.ok) throw new Error(result.reason); // 不带 # 也接受
    applySlideEdit(file, result.op);
    expect((await file.partText(partPath))!).toContain("<a:rPr><a:solidFill><a:srgbClr val=\"AABBCC\"/></a:solidFill></a:rPr><a:t>标题甲</a:t>");
  });

  it("非法颜色串带原因拒绝而不是抛异常", async () => {
    const file = await PptxFile.open(await minimalDeck());
    const partPath = file.model.slides[0]!.partPath;
    const bad = await setRunColor(file, { partPath, shapeId: 2, paraIndex: 0, runIndex: 0 }, "红色");
    expect(bad.ok).toBe(false);
    if (!bad.ok) expect(bad.reason).toContain("颜色");
  });
});
