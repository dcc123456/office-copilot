import { describe, expect, it } from "vitest";
import { auditPptxPackage } from "./audit.js";
import { PptxFile } from "./pptx-file.js";
import { applySlideEdit, readPresentation, setShapeGeometry, setShapeRunText } from "./index.js";
import { buildDeckBytes, minimalDeckParts, type DeckParts } from "./test-deck.js";

// 夹具的构造器在 `test-deck.ts`：桌面主进程出入口的单测用同一份，不许长两处（`AGENTS.md` §3.2）。

/**
 * 把夹具那张幻灯片里标题形状的 `<p:spPr/>` 换成指定内容，其余部件不动。
 *
 * @param spPrInner 要塞进 `<p:spPr>` 的内部 XML（如整段 `<a:xfrm>`）。
 * @returns 直接传给 {@link buildDeckBytes} 的部件表。
 */
function partsWithSpPr(spPrInner: string): DeckParts {
  const slide = minimalDeckParts()["ppt/slides/slide1.xml"]!;
  return { ...minimalDeckParts(), "ppt/slides/slide1.xml": slide.replace("<p:spPr/>", `<p:spPr>${spPrInner}</p:spPr>`) };
}

describe("auditPptxPackage：PowerPoint 会判损坏的几类结构问题", () => {
  it("结构合规的包返回空数组（不许误报）", async () => {
    expect(await auditPptxPackage(await buildDeckBytes(minimalDeckParts()))).toEqual([]);
  });

  it("通用那条确实在跑：部件没被 [Content_Types] 覆盖时报一条", async () => {
    const parts = { ...minimalDeckParts() };
    parts["ppt/media/picture.png"] = "not-a-real-png";
    const issues = await auditPptxPackage(await buildDeckBytes(parts));
    expect(issues).toEqual([{ code: "part-not-declared", part: "ppt/media/picture.png", detail: "png" }]);
  });

  it("幻灯片只有 cSld 空壳、没有 spTree 时报出", async () => {
    const parts = { ...minimalDeckParts(), "ppt/slides/slide1.xml": '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld/></p:sld>' };
    expect(await auditPptxPackage(await buildDeckBytes(parts))).toEqual([
      { code: "slide-missing-shape-tree", part: "ppt/slides/slide1.xml", detail: "" },
    ]);
  });

  it("幻灯片连关系表都没有时，报缺版式关系（PowerPoint 侧的必须关联）", async () => {
    const parts = { ...minimalDeckParts(), "ppt/slides/_rels/slide1.xml.rels": undefined };
    expect(await auditPptxPackage(await buildDeckBytes(parts))).toEqual([
      { code: "slide-missing-layout-relationship", part: "ppt/slides/slide1.xml", detail: "" },
    ]);
  });

  // 下面三条盯的是 `<a:off>`/`<a:ext>` 的必填属性只写了一半：§20.1.1.15 与 §20.1.1.14 把 x、y、cx、cy 全定为
  // required attribute，python-pptx 读到半个 `<a:off y>` 直接抛 InvalidXmlError（本机实测）。写侧那条守卫在 `edit.ts`。
  it("几何元素只写了一半必填属性时报出，缺哪个属性念进 detail", async () => {
    const parts = partsWithSpPr('<a:xfrm><a:off y="571500"/><a:ext cx="5486400" cy="685800"/></a:xfrm>');
    expect(await auditPptxPackage(await buildDeckBytes(parts))).toEqual([
      { code: "slide-incomplete-transform", part: "ppt/slides/slide1.xml", detail: "a:off 缺 x" },
    ]);
  });

  it("四个必填属性写全的几何不报（不许误报）", async () => {
    const parts = partsWithSpPr('<a:xfrm><a:off x="838200" y="571500"/><a:ext cx="5486400" cy="685800"/></a:xfrm>');
    expect(await auditPptxPackage(await buildDeckBytes(parts))).toEqual([]);
  });

  it("组合形状那对 chOff/chExt 一起查，且按本地名比对，换个命名空间前缀写法也抓得到", async () => {
    const parts = partsWithSpPr('<xfrm xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"><chOff x="0" y="0"/><ext cx="5486400"/></xfrm>');
    expect(await auditPptxPackage(await buildDeckBytes(parts))).toEqual([
      { code: "slide-incomplete-transform", part: "ppt/slides/slide1.xml", detail: "ext 缺 cy" },
    ]);
  });

  it("版式关系指向不存在的部件时，由通用那条报 relationship-target-missing", async () => {
    const parts = {
      ...minimalDeckParts(),
      "ppt/slides/_rels/slide1.xml.rels":
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout9.xml"/>' +
        "</Relationships>",
    };
    expect(await auditPptxPackage(await buildDeckBytes(parts))).toEqual([
      { code: "relationship-target-missing", part: "ppt/slides/slide1.xml", detail: "rId1 -> ppt/slideLayouts/slideLayout9.xml" },
    ]);
  });
});

/**
 * 我们自己写出的包必须过审计——这条是"改过的 `.pptx` 能在 PowerPoint 里直接打开"的机检面。
 *
 * 本机装不了 PowerPoint（`AGENTS.md` §8），python-pptx 只读值、Keynote 只会说打不开，
 * 所以拿同一套判定回头看三种产物：原始夹具、改过幻灯片文字的写回、写全四个几何值的写回。
 */
describe("我们自己的写入产物过审计（不许给自己判损坏）", () => {
  /** 打开夹具、把标题那格改成"改过的标题"，走桌面的保存路径出字节。 */
  async function bytesAfterTextEdit(): Promise<Uint8Array> {
    const bytes = await buildDeckBytes(minimalDeckParts());
    const file = await PptxFile.open(bytes);
    const { model } = await readPresentation(bytes);
    const partPath = model.slides[0]!.partPath;
    const made = await setShapeRunText(file, { partPath, shapeId: 2, paraIndex: 0, runIndex: 0 }, "改过的标题");
    if (!made.ok) throw new Error(JSON.stringify(made));
    applySlideEdit(file, made.op);
    return file.toBytes();
  }

  /** 打开夹具、给标题形状写全四个几何值（px → EMU 由引擎换算），走同一条保存路径出字节。 */
  async function bytesAfterGeometryEdit(): Promise<Uint8Array> {
    const file = await PptxFile.open(await buildDeckBytes(minimalDeckParts()));
    const made = await setShapeGeometry(file, { partPath: "ppt/slides/slide1.xml", shapeId: 2 }, {
      leftPx: 88,
      topPx: 60,
      widthPx: 576,
      heightPx: 72,
    });
    if (!made.ok) throw new Error(JSON.stringify(made));
    applySlideEdit(file, made.op);
    return file.toBytes();
  }

  it.each([
    ["原始夹具", async () => buildDeckBytes(minimalDeckParts())],
    ["改过幻灯片文字的写回", bytesAfterTextEdit],
    ["改过形状几何的写回", bytesAfterGeometryEdit],
  ] as const)("%s：审计返回空数组", async (_label, produce) => {
    expect(await auditPptxPackage(await produce())).toEqual([]);
  });
});
