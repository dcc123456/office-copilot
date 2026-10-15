import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { PptxFile } from "./pptx-file.js";
import { addTextBox, applySlideEdit, moveShapeLayer, readPresentation, removeShape, revertSlideEdit, revertSlideEdits, setRunColor, setRunFont, setShapeGeometry, setShapeRunText } from "./index.js";

/**
 * 造一份最小 `.pptx` 字节，部件形状照抄本机 python-pptx 产物（spec §取证）。
 *
 * @param slideXml 要写进 `ppt/slides/slide1.xml` 的文本；传 `null` 表示**这个部件干脆不存在**
 *   （关系表仍指向它），用来验读侧的失败码；不传用默认那份三形状幻灯片。
 * @param notesText 给定时额外挂一个备注部件（`ppt/notesSlides/notesSlide1.xml`，带关系与内容类型 Override，
 *   照 §2.10 那三处一起登记），段落用它拼出来；不给就没有备注页。
 * @returns zip 的 Uint8Array。含一张 4:3 幻灯片、三种形状（标题占位符、正文占位符、自带几何的文本框）
 *   与一个二进制部件，用来同时验读模型与保真回写。
 */
async function minimalDeck(slideXml?: string | null, notesText?: string): Promise<Uint8Array> {
  const zip = new JSZip();
  const notesOverride =
    '<Override PartName="/ppt/notesSlides/notesSlide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>';
  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>${notesText === undefined ? "" : notesOverride}</Types>`,
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
  const defaultSlideXml =
    '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>标题甲</a:t></a:r></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Content Placeholder 2"/><p:cNvSpPr/><p:nvPr><p:ph idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>副标题乙</a:t></a:r></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="4" name="TextBox 3"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="914400" y="3657600"/><a:ext cx="2743200" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>正文丙</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>';
  if (slideXml !== null) zip.file("ppt/slides/slide1.xml", slideXml ?? defaultSlideXml);
  if (notesText !== undefined) {
    // 备注部件照 WPS 自带模板里那份的真实形状：sldImg + body + sldNum 三个占位符，页码那格正文是 `<a:fld>`，
    // 读侧只许取 body 那一格的文字（把页码念成备注就是编内容）。
    zip.file(
      "ppt/slides/_rels/slide1.xml.rels",
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide1.xml"/></Relationships>',
    );
    zip.file(
      "ppt/notesSlides/notesSlide1.xml",
      `<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="幻灯片图像占位符 1"/><p:cNvSpPr/><p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>缩略图里不该被读到的字</a:t></a:r></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="备注占位符 2"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>${notesText
        .split("⏎")
        .map((line) => `<a:p><a:r><a:t>${line}</a:t></a:r></a:p>`)
        .join("")}</p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="4" name="灯片编号占位符 3"/><p:cNvSpPr/><p:nvPr><p:ph type="sldNum" sz="quarter" idx="10"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:fld id="{85D0DACE-38E0-42D2-9336-2B707D34BC6D}" type="slidenum"><a:r><a:t>1</a:t></a:r></a:fld></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:notes>`,
    );
  }
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
    expect(title!.boxPx).toBeUndefined(); // 继承来的坐标不给数字——视图据此把几何输入框换成一句说明
    expect(title!.paragraphTexts).toEqual(["标题甲"]);
    expect(title!.paragraphRunCounts).toEqual([1]); // 一段一条 run：视图允许整段改
    expect(title!.paragraphRunFormats).toEqual([[{}]]); // 那条 run 没写 rPr：字号/粗体/颜色全是继承，不许猜
    expect(body!.isPlaceholder).toBe(true);
    expect(body!.placeholderRef).toBe("1"); // idx="1" 而非 type
    expect(textbox!.isPlaceholder).toBe(false);
    expect(textbox!.hasOwnGeometry).toBe(true);
    expect(textbox!.paragraphTexts).toEqual(["正文丙"]);
    expect(textbox!.paragraphRunCounts).toEqual([1]);
    // EMU → CSS px 只在 IO 边发生一次（§2.1）：914400/9525=96，3657600/9525=384，2743200/9525=288。
    expect(textbox!.boxPx).toEqual({ leftPx: 96, topPx: 384, widthPx: 288, heightPx: 96 });
  });

  it("一段里两条不同格式 run 时如实报 2，视图据此不许整段改（§2.3）", async () => {
    const { model } = await readPresentation(await twoRunDeck());
    const title = model.slides[0]!.shapes.find((s) => s.shapeId === 2)!;
    // 文字是拼接串，看不出这段里其实有粗体与非粗体两截——条数才是视图要的那个信息。
    expect(title.paragraphTexts).toEqual(["标题甲副"]);
    expect(title.paragraphRunCounts).toEqual([2]);
    // 逐 run 的格式照实给：run0 只写了 b，run1 只写了 sz（半磅 1800 → 18pt），都没写颜色。
    expect(title.paragraphRunFormats).toEqual([[{ bold: true }, { sizePt: 18 }]]);
  });
});

describe("pptx-io：一页读不懂时的失败码（spec line 439）", () => {
  it("部件缺失、XML 解析失败、没有形状树——三种各有各的码，不折成\"这一页没有形状\"", async () => {
    const missing = (await readPresentation(await minimalDeck(null))).model.slides[0]!;
    expect(missing.readFailure).toBe("missing-part");
    expect(missing.shapes).toHaveLength(0);
    const broken = (await readPresentation(await minimalDeck('<p:sld xmlns:p="p"><p:cSld><p:spTree><p:sp>'))).model.slides[0]!;
    expect(broken.readFailure).toBe("invalid-xml");
    const noTree = (await readPresentation(await minimalDeck('<p:sld xmlns:p="p"><p:cSld/></p:sld>'))).model.slides[0]!;
    expect(noTree.readFailure).toBe("no-sp-tree");
  });

  it("读得动的那一页不给失败码——这时空数组才真的是\"这一页没有形状\"", async () => {
    const emptyTree = (
      await readPresentation(
        await minimalDeck('<p:sld xmlns:p="p"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sld>'),
      )
    ).model.slides[0]!;
    expect(emptyTree.readFailure).toBeUndefined();
    expect(emptyTree.shapes).toHaveLength(0);
    const normal = (await readPresentation(await minimalDeck())).model.slides[0]!;
    expect(normal.readFailure).toBeUndefined();
    expect(normal.shapes).toHaveLength(3);
  });
});

describe("pptx-io：演讲者备注的读侧（spec 阶段 2 那条「notes 仍缺」）", () => {
  it("只取 type=\"body\" 那一格：缩略图占位符与页码字段都不算备注正文", async () => {
    const slide = (await readPresentation(await minimalDeck(undefined, "第一句备注⏎第二句备注"))).model.slides[0]!;
    expect(slide.notes?.partPath).toBe("ppt/notesSlides/notesSlide1.xml");
    expect(slide.notes?.paragraphTexts).toEqual(["第一句备注", "第二句备注"]);
  });

  it("没有备注部件与有部件但没写字是两件事，不许混着报", async () => {
    const none = (await readPresentation(await minimalDeck())).model.slides[0]!;
    expect(none.notes).toBeUndefined();
    const empty = (await readPresentation(await minimalDeck(undefined, ""))).model.slides[0]!;
    expect(empty.notes?.partPath).toBe("ppt/notesSlides/notesSlide1.xml");
    expect(empty.notes?.paragraphTexts).toEqual([""]);
  });

  it("备注部件是未建模部件：一条都不改地存回去，字节与段落都原样", async () => {
    const bytes = await minimalDeck(undefined, "原样带过的备注");
    const out = await (await PptxFile.open(bytes)).toBytes();
    const zip = await JSZip.loadAsync(out);
    expect(await zip.file("ppt/notesSlides/notesSlide1.xml")!.async("string")).toBe(
      await (await JSZip.loadAsync(bytes)).file("ppt/notesSlides/notesSlide1.xml")!.async("string"),
    );
    expect((await readPresentation(out)).model.slides[0]!.notes?.paragraphTexts).toEqual(["原样带过的备注"]);
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

describe("pptx-io：撤销的现场核对（内容寻址，§2.5）", () => {
  it("现场正是改完后的那个样子时，前像写回、这一页逐字回到改前", async () => {
    const file = await PptxFile.open(await minimalDeck());
    const partPath = file.model.slides[0]!.partPath;
    const original = (await file.partText(partPath))!;
    const made = await setShapeRunText(file, { partPath, shapeId: 2, paraIndex: 0, runIndex: 0 }, "新标题");
    if (!made.ok) throw new Error(made.reason);
    applySlideEdit(file, made.op);

    const outcome = await revertSlideEdits(file, [made.op]);
    expect(outcome.applied).toEqual([made.op]);
    expect(outcome.skipped).toEqual([]);
    expect((await file.partText(partPath))!).toBe(original);
  });

  it("同批两条改同一页：倒序送回时逐条现场对得上，整页一路回到改前", async () => {
    const file = await PptxFile.open(await minimalDeck());
    const partPath = file.model.slides[0]!.partPath;
    const original = (await file.partText(partPath))!;
    const first = await setShapeRunText(file, { partPath, shapeId: 2, paraIndex: 0, runIndex: 0 }, "第一次");
    if (!first.ok) throw new Error(first.reason);
    applySlideEdit(file, first.op);
    const second = await setShapeRunText(file, { partPath, shapeId: 3, paraIndex: 0, runIndex: 0 }, "第二次");
    if (!second.ok) throw new Error(second.reason);
    applySlideEdit(file, second.op);

    const outcome = await revertSlideEdits(file, [second.op, first.op]);
    expect(outcome.applied).toEqual([second.op, first.op]);
    expect((await file.partText(partPath))!).toBe(original);
  });

  it("这一页后来又被手工改过：那一条跳过并给理由，旧前像绝不覆盖后来的编辑", async () => {
    const file = await PptxFile.open(await minimalDeck());
    const partPath = file.model.slides[0]!.partPath;
    const made = await setShapeRunText(file, { partPath, shapeId: 2, paraIndex: 0, runIndex: 0 }, "AI 改的");
    if (!made.ok) throw new Error(made.reason);
    applySlideEdit(file, made.op);
    const manual = await setShapeRunText(file, { partPath, shapeId: 3, paraIndex: 0, runIndex: 0 }, "用户自己改的");
    if (!manual.ok) throw new Error(manual.reason);
    applySlideEdit(file, manual.op);
    const beforeRevertAttempt = (await file.partText(partPath))!;

    const outcome = await revertSlideEdits(file, [made.op]);
    expect(outcome.applied).toEqual([]);
    expect(outcome.skipped[0]!.reason).toContain("后来又被改过");
    // op 记的是整页后像，而这一页后来又被改过，所以整条撤不动，用户那句"自己改的"一个字没丢。
    expect((await file.partText(partPath))!).toBe(beforeRevertAttempt);
  });

  it("部件已经不在包里时给理由而不是抛异常", async () => {
    const file = await PptxFile.open(await minimalDeck());
    const ghost = { partPath: "ppt/slides/slide9.xml", beforeXml: "<p:sld/>", afterXml: "<p:sld>x</p:sld>" };
    const outcome = await revertSlideEdits(file, [ghost]);
    expect(outcome.applied).toEqual([]);
    expect(outcome.skipped[0]!.reason).toContain("已经找不到");
  });
});

describe("pptx-io：叠放次序 op（结构 op，§2.5）", () => {
  /** 从幻灯片 XML 里按文档顺序取出形状名——`<p:spTree>` 的文档顺序就是叠放次序（第一个最底）。 */
  function shapeOrder(xml: string): string[] {
    return [...xml.matchAll(/<p:cNvPr id="\d+" name="([^"]*)"/g)].map((m) => m[1]!).filter((name) => name !== "");
  }

  it("上移一层：目标与后一个形状交换位置，撤销逐字回原样", async () => {
    const file = await PptxFile.open(await minimalDeck());
    const partPath = file.model.slides[0]!.partPath;
    const original = (await file.partText(partPath))!;
    const result = await moveShapeLayer(file, { partPath, shapeId: 2 }, "up");
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(shapeOrder(result.op.afterXml)).toEqual(["Content Placeholder 2", "Title 1", "TextBox 3"]);
    // 形状之前的两个非形状子元素（nvGrpSpPr / grpSpPr）不许被卷进重排
    // （op 的后像带 XML 声明、部件原文不带，所以只比第一个形状之前的那一段）
    const head = (xml: string) => xml.slice(xml.indexOf("<p:cSld"), xml.indexOf("<p:sp>"));
    expect(head(result.op.afterXml)).toBe(head(original));

    applySlideEdit(file, result.op);
    expect(shapeOrder((await file.partText(partPath))!)).toEqual(["Content Placeholder 2", "Title 1", "TextBox 3"]);
    revertSlideEdit(file, result.op);
    expect((await file.partText(partPath))!).toBe(original);
  });

  it("置于底层：最后一个形状挪到最前，其它形状的相对顺序不变", async () => {
    const file = await PptxFile.open(await minimalDeck());
    const partPath = file.model.slides[0]!.partPath;
    const result = await moveShapeLayer(file, { partPath, shapeId: 4 }, "bottom");
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(shapeOrder(result.op.afterXml)).toEqual(["TextBox 3", "Title 1", "Content Placeholder 2"]);
  });

  it("已经在要去的尽头时回带理由的空提案：不产出 before==after 的假改动", async () => {
    const file = await PptxFile.open(await minimalDeck());
    const partPath = file.model.slides[0]!.partPath;
    const original = (await file.partText(partPath))!;
    const alreadyBottom = await moveShapeLayer(file, { partPath, shapeId: 2 }, "down");
    expect(alreadyBottom.ok).toBe(false);
    if (!alreadyBottom.ok) expect(alreadyBottom.reason).toContain("最底层");
    const alreadyTop = await moveShapeLayer(file, { partPath, shapeId: 4 }, "top");
    expect(alreadyTop.ok).toBe(false);
    if (!alreadyTop.ok) expect(alreadyTop.reason).toContain("最顶层");
    expect((await file.partText(partPath))!).toBe(original); // 被拒的提案一个字都不落
  });

  it("形状号越界与部件缺失各回一条理由，不抛异常", async () => {
    const file = await PptxFile.open(await minimalDeck());
    const partPath = file.model.slides[0]!.partPath;
    const badShape = await moveShapeLayer(file, { partPath, shapeId: 999 }, "up");
    expect(badShape.ok).toBe(false);
    if (!badShape.ok) expect(badShape.reason).toContain("找不到形状");
    const badPart = await moveShapeLayer(file, { partPath: "ppt/slides/slide9.xml", shapeId: 2 }, "up");
    expect(badPart.ok).toBe(false);
    if (!badPart.ok) expect(badPart.reason).toContain("部件不存在");
  });
});

describe("pptx-io：增删形状 op（结构 op，新建要分配形状号）", () => {
  /** 按文档顺序取出这一页的形状号（spTree 自己的 `id="1"` 不算形状）。 */
  function shapeIds(xml: string): number[] {
    return [...xml.matchAll(/<p:cNvPr id="(\d+)" name="([^"]*)"/g)]
      .map((m) => Number(m[1]))
      .filter((id) => id > 1);
  }

  /** 第一个 `<p:sp>` 的原文，用来证明加/删都只碰自己那一个节点。 */
  function firstShapeXml(xml: string): string {
    const from = xml.indexOf("<p:sp>");
    return xml.slice(from, xml.indexOf("</p:sp>", from) + "</p:sp>".length);
  }

  it("加文本框：形状号是这一页已有 @id 的最大值 +1，新形状挂在形状子序列末尾（= 最顶层）", async () => {
    const file = await PptxFile.open(await minimalDeck());
    const partPath = file.model.slides[0]!.partPath;
    const original = (await file.partText(partPath))!;
    const result = await addTextBox(file, { partPath }, { leftPx: 96, topPx: 96, widthPx: 480, heightPx: 144 }, "新文本框");
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(shapeIds(result.op.afterXml)).toEqual([2, 3, 4, 5]);
    expect(result.op.afterXml).toContain(firstShapeXml(original)); // 已有形状逐字不动（§2.3）

    applySlideEdit(file, result.op);
    const reread = await readPresentation(await file.toBytes());
    const added = reread.model.slides[0]!.shapes[3]!;
    expect(added.shapeId).toBe(5);
    expect(added.name).toBe("TextBox 4");
    expect(added.paragraphTexts).toEqual(["新文本框"]);
    expect(added.paragraphRunCounts).toEqual([1]); // 一条 run：视图允许整段改
    expect(added.hasOwnGeometry).toBe(true);
    expect(added.boxPx).toEqual({ leftPx: 96, topPx: 96, widthPx: 480, heightPx: 144 }); // px 交出什么读回什么（§2.1）
    revertSlideEdit(file, result.op);
    expect((await file.partText(partPath))!).toBe(original);
  });

  it("加文本框：px→EMU 只在 units.ts 那一处，形状骨架照抄 PowerPoint 插入文本框的那一段", async () => {
    const file = await PptxFile.open(await minimalDeck());
    const partPath = file.model.slides[0]!.partPath;
    const result = await addTextBox(file, { partPath }, { leftPx: 96, topPx: 96, widthPx: 480, heightPx: 144 }, "字");
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    // 96px = 914400 EMU，480px = 4572000 EMU，144px = 1371600 EMU
    expect(result.op.afterXml).toContain('<a:off x="914400" y="914400"/><a:ext cx="4572000" cy="1371600"/>');
    expect(result.op.afterXml).toContain('<p:cNvSpPr txBox="1"/>');
    expect(result.op.afterXml).toContain('<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>');
    expect(result.op.afterXml).toContain('<a:bodyPr wrap="none"><a:spAutoFit/></a:bodyPr><a:lstStyle/>');
  });

  it("加文本框：文字里的尖括号与 & 按 XML 文本转义，读回来是原样那几个字", async () => {
    const file = await PptxFile.open(await minimalDeck());
    const partPath = file.model.slides[0]!.partPath;
    const result = await addTextBox(file, { partPath }, { leftPx: 0, topPx: 0, widthPx: 100, heightPx: 50 }, "a<b>&c");
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.op.afterXml).toContain("<a:t>a&lt;b&gt;&amp;c</a:t>");
    applySlideEdit(file, result.op);
    const reread = await readPresentation(await file.toBytes());
    expect(reread.model.slides[0]!.shapes[3]!.paragraphTexts).toEqual(["a<b>&c"]);
  });

  it("加文本框：空串文字保留模板里那个空的 <a:p/>，不造空 run", async () => {
    const file = await PptxFile.open(await minimalDeck());
    const partPath = file.model.slides[0]!.partPath;
    const empty = await addTextBox(file, { partPath }, { leftPx: 0, topPx: 0, widthPx: 100, heightPx: 50 }, "");
    expect(empty.ok).toBe(true);
    if (!empty.ok) return;
    expect(empty.op.afterXml).toContain("<a:p/>");
    expect(empty.op.afterXml.slice(empty.op.afterXml.lastIndexOf("<p:sp>"))).not.toContain("<a:t>");
  });

  it("加文本框：宽高非正与部件缺失各回一条理由，被拒的提案一个字都不落", async () => {
    const file = await PptxFile.open(await minimalDeck());
    const partPath = file.model.slides[0]!.partPath;
    const original = (await file.partText(partPath))!;
    const badWidth = await addTextBox(file, { partPath }, { leftPx: 0, topPx: 0, widthPx: 0, heightPx: 50 }, "字");
    expect(badWidth.ok).toBe(false);
    if (!badWidth.ok) expect(badWidth.reason).toContain("宽度");
    const badHeight = await addTextBox(file, { partPath }, { leftPx: 0, topPx: 0, widthPx: 10, heightPx: -1 }, "字");
    expect(badHeight.ok).toBe(false);
    if (!badHeight.ok) expect(badHeight.reason).toContain("高度");
    const badPart = await addTextBox(file, { partPath: "ppt/slides/slide9.xml" }, { leftPx: 0, topPx: 0, widthPx: 10, heightPx: 10 }, "字");
    expect(badPart.ok).toBe(false);
    if (!badPart.ok) expect(badPart.reason).toContain("部件不存在");
    expect((await file.partText(partPath))!).toBe(original);
  });

  it("删形状：只摘掉那一个 <p:sp>，其余形状逐字不动，撤销逐字回原样", async () => {
    const file = await PptxFile.open(await minimalDeck());
    const partPath = file.model.slides[0]!.partPath;
    const original = (await file.partText(partPath))!;
    const result = await removeShape(file, { partPath, shapeId: 3 });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(shapeIds(result.op.afterXml)).toEqual([2, 4]);
    expect(result.op.afterXml).not.toContain("副标题乙");
    expect(result.op.afterXml).toContain(firstShapeXml(original));
    expect(result.op.afterXml).toContain("正文丙");

    applySlideEdit(file, result.op);
    const reread = await readPresentation(await file.toBytes());
    expect(reread.model.slides[0]!.shapes.map((s) => s.name)).toEqual(["Title 1", "TextBox 3"]);
    revertSlideEdit(file, result.op);
    expect((await file.partText(partPath))!).toBe(original);
  });

  it("删掉这一页最后一个形状是合法操作（页面变空白，读侧不炸），越界与部件缺失回理由不抛异常", async () => {
    const file = await PptxFile.open(await minimalDeck());
    const partPath = file.model.slides[0]!.partPath;
    for (const shapeId of [2, 3, 4]) {
      const removed = await removeShape(file, { partPath, shapeId });
      expect(removed.ok).toBe(true);
      if (!removed.ok) return;
      applySlideEdit(file, removed.op);
    }
    const reread = await readPresentation(await file.toBytes());
    expect(reread.model.slides[0]!.shapes).toEqual([]);

    const badShape = await removeShape(file, { partPath, shapeId: 999 });
    expect(badShape.ok).toBe(false);
    if (!badShape.ok) expect(badShape.reason).toContain("找不到形状");
    const badPart = await removeShape(file, { partPath: "ppt/slides/slide9.xml", shapeId: 2 });
    expect(badPart.ok).toBe(false);
    if (!badPart.ok) expect(badPart.reason).toContain("部件不存在");
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

describe("pptx-io：形状几何 op（移动 / 缩放）", () => {
  it("只改宽：目标形状的 off 与 cy 原样留住，只有 cx 按 px×9525 换成 EMU（§2.1）", async () => {
    const file = await PptxFile.open(await minimalDeck()); // 文本框 shapeId=4，已有 xfrm
    const partPath = file.model.slides[0]!.partPath;
    const result = await setShapeGeometry(file, { partPath, shapeId: 4 }, { widthPx: 100 });
    if (!result.ok) throw new Error(result.reason);
    applySlideEdit(file, result.op);
    const after = (await file.partText(partPath))!;
    expect(after).toContain("<a:off x=\"914400\" y=\"3657600\"/><a:ext cx=\"952500\" cy=\"914400\"/>"); // 100px→952500 EMU，off/cy 没动
  });

  it("占位符空 spPr 上新建 xfrm，off 先于 ext（D3：用户明示覆盖才写几何）", async () => {
    const file = await PptxFile.open(await minimalDeck()); // 标题 shapeId=2，spPr 是空的
    const partPath = file.model.slides[0]!.partPath;
    const result = await setShapeGeometry(file, { partPath, shapeId: 2 }, { leftPx: 100, topPx: 100, widthPx: 200, heightPx: 200 });
    if (!result.ok) throw new Error(result.reason);
    applySlideEdit(file, result.op);
    const after = (await file.partText(partPath))!;
    expect(after).toContain("<p:spPr><a:xfrm><a:off x=\"952500\" y=\"952500\"/><a:ext cx=\"1905000\" cy=\"1905000\"/></a:xfrm></p:spPr>");
  });

  it("占位符那种自己没写几何的形状，只给一轴时带原因拒绝（OOXML 里 x、y 与 cx、cy 都是必填）", async () => {
    const file = await PptxFile.open(await minimalDeck()); // 标题 shapeId=2，<p:spPr/> 是空的
    const partPath = file.model.slides[0]!.partPath;
    const original = (await file.partText(partPath))!;
    const onlyTop = await setShapeGeometry(file, { partPath, shapeId: 2 }, { topPx: 60 });
    expect(onlyTop.ok).toBe(false);
    if (!onlyTop.ok) expect(onlyTop.reason).toContain("leftPx");
    const onlyWidth = await setShapeGeometry(file, { partPath, shapeId: 2 }, { widthPx: 200 });
    expect(onlyWidth.ok).toBe(false);
    if (!onlyWidth.ok) expect(onlyWidth.reason).toContain("heightPx");
    expect(await file.partText(partPath)).toBe(original); // 拒了就不许留下半个 <a:off>
  });

  it("apply 后部件变、revert 逐字回到原样", async () => {
    const file = await PptxFile.open(await minimalDeck());
    const partPath = file.model.slides[0]!.partPath;
    const original = (await file.partText(partPath))!;
    const result = await setShapeGeometry(file, { partPath, shapeId: 4 }, { leftPx: 50, heightPx: 150 });
    if (!result.ok) throw new Error(result.reason);
    applySlideEdit(file, result.op);
    expect((await file.partText(partPath))!).not.toBe(original);
    revertSlideEdit(file, result.op);
    expect((await file.partText(partPath))!).toBe(original);
  });

  it("没给几何属性或宽/高非正数时带原因拒绝", async () => {
    const file = await PptxFile.open(await minimalDeck());
    const partPath = file.model.slides[0]!.partPath;
    const empty = await setShapeGeometry(file, { partPath, shapeId: 4 }, {});
    expect(empty.ok).toBe(false);
    const badW = await setShapeGeometry(file, { partPath, shapeId: 4 }, { widthPx: 0 });
    expect(badW.ok).toBe(false);
    if (!badW.ok) expect(badW.reason).toContain("宽度");
  });
});

describe("pptx-io：写进去再读回来，单位与颜色形态原样对上（§2.1/§2.2 全程只有一处换算）", () => {
  it("字号 pt、颜色 hex、几何 px 交出什么读回什么，视图不会看到第二次换算的数", async () => {
    const file = await PptxFile.open(await minimalDeck()); // 标题 run 原本没有 rPr，spPr 是空的
    const partPath = file.model.slides[0]!.partPath;
    const target = { partPath, shapeId: 2, paraIndex: 0, runIndex: 0 };
    // 逐条算、逐条落：后一条的前像必须包含前一条的结果（视图那批改动在主进程里也是这个顺序）。
    for (const produce of [
      () => setRunFont(file, target, { sizePt: 24.5, bold: true }),
      () => setRunColor(file, target, "#3cb371"),
      () => setShapeGeometry(file, { partPath, shapeId: 2 }, { leftPx: 120, topPx: 60, widthPx: 500, heightPx: 250 }),
    ]) {
      const result = await produce();
      if (!result.ok) throw new Error(result.reason);
      applySlideEdit(file, result.op);
    }
    const { model } = await readPresentation(await file.toBytes());
    const title = model.slides[0]!.shapes.find((s) => s.shapeId === 2)!;
    // 24.5pt → sz=2450 → 24.5pt；颜色大写落盘、小写读回（编辑器内一律 #rrggbb，§2.2）。
    expect(title.paragraphRunFormats).toEqual([[{ sizePt: 24.5, bold: true, hex: "#3cb371" }]]);
    expect(title.boxPx).toEqual({ leftPx: 120, topPx: 60, widthPx: 500, heightPx: 250 });
    expect(title.hasOwnGeometry).toBe(true); // 用户明示覆盖后，这份几何就是形状自己的了
  });
});
