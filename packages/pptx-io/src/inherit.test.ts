import { describe, expect, it } from "vitest";
import { loadPackage, xml } from "@office-copilot/opc";
import { auditPptxPackage } from "./audit.js";
import { applySlideEdit, setShapeRunText } from "./edit.js";
import { createGeometryInheritance, readExplicitBoxAxes } from "./inherit.js";
import { PptxFile } from "./pptx-file.js";
import { readPresentation, type SlideShape } from "./read.js";
import { buildDeckBytes, inheritanceDeckParts, minimalDeckParts, type DeckParts } from "./test-deck.js";
import type { ShapeGeometryPatch } from "./units.js";

// 夹具在 `test-deck.ts` 的 `inheritanceDeckParts`：三格各占继承链上的一种情形，别处不许再造一份（§3.2）。

/**
 * 在夹具上跑一遍读模型，按形状号点名列出那张幻灯片里的形状。
 *
 * @param parts 夹具部件表（想改某一格就复制一份覆盖对应部件）。
 * @returns `shapeId → 形状快照`；夹具里形状号不重复。
 */
async function shapesById(parts: DeckParts): Promise<Map<number, SlideShape>> {
  const { model } = await readPresentation(await buildDeckBytes(parts));
  return new Map(model.slides[0]!.shapes.map((shape) => [shape.shapeId, shape]));
}

describe("readExplicitBoxAxes：逐轴读 `<a:xfrm>` 里显式写了的轴", () => {
  const NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';

  /**
   * 解析一段 `<p:spPr>` 并抽出它显式写了的几何轴。
   *
   * @param spPrXml `<p:spPr>` 的完整 XML（带命名空间声明）。
   * @returns 逐轴的值，单位 CSS px @96DPI；没写的轴不在表里。
   */
  function axesOf(spPrXml: string): ShapeGeometryPatch {
    // 这段 XML 的根就是 `<p:spPr>` 自己，直接把它交给被测函数（不用再找一层子元素）。
    return readExplicitBoxAxes(xml.parseDocument(spPrXml));
  }

  it("四数齐全时四个都给，EMU 按 9525 换成整 px（换算只在 units.ts，§2.1）", () => {
    expect(axesOf(`<p:spPr ${NS}><a:xfrm><a:off x="952500" y="1143000"/><a:ext cx="2857500" cy="762000"/></a:xfrm></p:spPr>`)).toEqual({
      leftPx: 100,
      topPx: 120,
      widthPx: 300,
      heightPx: 80,
    });
  });

  it("只写了 `a:off` 时只给左/上两轴——python-pptx 的逐轴回退就靠这一条（`_effective_value`）", () => {
    expect(axesOf(`<p:spPr ${NS}><a:xfrm><a:off x="952500" y="1143000"/></a:xfrm></p:spPr>`)).toEqual({ leftPx: 100, topPx: 120 });
  });

  it("没有 `<p:spPr>` 或没有 `<a:xfrm>` 时给空对象，不是补零（D3 不许造值）", () => {
    expect(axesOf(`<p:spPr ${NS}/>`)).toEqual({});
    expect(readExplicitBoxAxes(undefined)).toEqual({});
  });

  it("属性写的是非数字时那一轴直接没有，其余轴照给", () => {
    expect(axesOf(`<p:spPr ${NS}><a:xfrm><a:off x="abc" y="1143000"/><a:ext cx="2857500" cy="762000"/></a:xfrm></p:spPr>`)).toEqual({
      topPx: 120,
      widthPx: 300,
      heightPx: 80,
    });
  });
});

describe("读模型的 effectiveBoxPx：形状自己写的轴赢，缺的按 版式 → 母版 补（spec D3）", () => {
  it("幻灯片与版式都写空 `<p:spPr/>` 的标题：四轴全来自母版，且版式的 `ctrTitle` 按 python-pptx 那张表归到母版的 `title`", async () => {
    const title = (await shapesById(inheritanceDeckParts())).get(2)!;
    expect(title.boxPx).toBeUndefined(); // 自己没写这条事实必须留着：写侧靠它决定"不许烘焙"（D3）
    expect(title.effectiveBoxPx).toEqual({ leftPx: 96, topPx: 48, widthPx: 768, heightPx: 120 });
    expect(title.boxPxSources).toEqual({ leftPx: "master", topPx: "master", widthPx: "master", heightPx: "master" });
  });

  it("幻灯片那格自己写全了四数：取自己的值，来源标 shape（继承值不许盖掉它）", async () => {
    const body = (await shapesById(inheritanceDeckParts())).get(3)!;
    expect(body.placeholderRef).toBe("body");
    expect(body.effectiveBoxPx).toEqual({ leftPx: 100, topPx: 120, widthPx: 300, heightPx: 80 });
    expect(body.boxPxSources).toEqual({ leftPx: "shape", topPx: "shape", widthPx: "shape", heightPx: "shape" });
  });

  it("版式写了几何、幻灯片那格空着：四轴来自版式，来源必须是 layout（视图据此显示「继承」）", async () => {
    const parts = inheritanceDeckParts();
    parts["ppt/slides/slide1.xml"] = parts["ppt/slides/slide1.xml"]!.replace(
      '<p:spPr><a:xfrm><a:off x="952500" y="1143000"/><a:ext cx="2857500" cy="762000"/></a:xfrm></p:spPr>',
      "<p:spPr/>",
    );
    const body = (await shapesById(parts)).get(3)!;
    expect(body.effectiveBoxPx).toEqual({ leftPx: 72, topPx: 192, widthPx: 480, heightPx: 288 });
    expect(body.boxPxSources).toEqual({ leftPx: "layout", topPx: "layout", widthPx: "layout", heightPx: "layout" });
  });

  it("没有 `<p:ph>` 的文本框不参与继承：自己没写就一个轴都没有，哪怕版式第一格写着数", async () => {
    const parts = inheritanceDeckParts();
    parts["ppt/slides/slide1.xml"] = parts["ppt/slides/slide1.xml"]!.replace(
      '<p:spPr><a:xfrm><a:off x="476250" y="476250"/><a:ext cx="1905000" cy="952500"/></a:xfrm></p:spPr>',
      "<p:spPr/>",
    );
    const textBox = (await shapesById(parts)).get(4)!;
    expect(textBox.effectiveBoxPx).toEqual({});
    expect(textBox.boxPxSources).toEqual({});
  });

  it("幻灯片那一格的 `idx` 在版式里找不到对应格子时不继承（匹配键是 idx，python-pptx 的 `placeholders.get(idx=)`）", async () => {
    const parts = inheritanceDeckParts();
    parts["ppt/slides/slide1.xml"] = parts["ppt/slides/slide1.xml"]!.replace('<p:ph type="ctrTitle"/>', '<p:ph type="ctrTitle" idx="9"/>');
    expect((await shapesById(parts)).get(2)!.effectiveBoxPx).toEqual({});
  });

  it("版式与母版之间那条关系断了时只走到版式那一层，不越过它去猜母版", async () => {
    const parts = inheritanceDeckParts();
    parts["ppt/slideLayouts/_rels/slideLayout1.xml.rels"] = undefined;
    const shapes = await shapesById(parts);
    expect(shapes.get(2)!.effectiveBoxPx).toEqual({}); // 版式那一格自己没写，母版又查不到 → 没有就是不没有
    expect(shapes.get(3)!.boxPxSources.widthPx).toBe("shape"); // 自己写满的那格不受影响
  });

  it("最小夹具（版式里没有占位符格子）继承不出东西：`effectiveBoxPx` 是空对象而不是四个 0", async () => {
    const title = (await shapesById(minimalDeckParts())).get(2)!;
    expect(title.effectiveBoxPx).toEqual({});
    expect(title.boxPxSources).toEqual({});
  });
});

describe("createGeometryInheritance：母版那一格按版式的类型找，查不到就停在上一层", () => {
  /**
   * 在夹具上直接建解析器。
   *
   * @param parts 夹具部件表。
   * @returns `resolve` 入口（一份包一个，版式/母版只解析一次）。
   */
  async function inheritanceOf(parts: DeckParts) {
    return createGeometryInheritance(await loadPackage(await buildDeckBytes(parts)));
  }

  it("版式那一格是 `subTitle` 时归到母版的 `body`：数从母版的 body 格子来（`LayoutPlaceholder._base_placeholder` 那张表）", async () => {
    const parts = inheritanceDeckParts();
    parts["ppt/slideLayouts/slideLayout1.xml"] = parts["ppt/slideLayouts/slideLayout1.xml"]!.replace('<p:ph type="ctrTitle"/>', '<p:ph type="subTitle"/>');
    const resolved = await (await inheritanceOf(parts)).resolve("ppt/slides/slide1.xml", { idx: 0, type: "ctrTitle" }, {});
    expect(resolved.box).toEqual({ leftPx: 80, topPx: 160, widthPx: 400, heightPx: 200 });
    expect(resolved.sources.widthPx).toBe("master");
  });

  it("表里没有的类型原样往母版比对（python-pptx 这种情况直接抛 KeyError，我们只当查不到）", async () => {
    const parts = inheritanceDeckParts();
    parts["ppt/slideLayouts/slideLayout1.xml"] = parts["ppt/slideLayouts/slideLayout1.xml"]!.replace('<p:ph type="ctrTitle"/>', '<p:ph type="vertTitle"/>');
    expect(await (await inheritanceOf(parts)).resolve("ppt/slides/slide1.xml", { idx: 0, type: "vertTitle" }, {})).toEqual({ box: {}, sources: {} });
  });

  it("幻灯片路径不在包里（关系查不到版式）时只回自己写的轴，不抛", async () => {
    const resolved = await (await inheritanceOf(inheritanceDeckParts())).resolve("ppt/slides/slide404.xml", { idx: 0, type: "ctrTitle" }, { leftPx: 10 });
    expect(resolved).toEqual({ box: { leftPx: 10 }, sources: { leftPx: "shape" } });
  });
});

describe("继承链不许改变落盘产物（spec D3 的验收面）", () => {
  it("带完整继承链的夹具自己过审计返回空数组（不许误报）", async () => {
    expect(await auditPptxPackage(await buildDeckBytes(inheritanceDeckParts()))).toEqual([]);
  });

  it("模型算得出标题的继承几何，改完文字落盘后 `<p:spPr/>` 仍是空的（有效值只读，绝不烤进 slide）", async () => {
    const bytes = await buildDeckBytes(inheritanceDeckParts());
    const { model } = await readPresentation(bytes);
    const partPath = model.slides[0]!.partPath;
    const title = model.slides[0]!.shapes.find((shape) => shape.shapeId === 2)!;
    expect(title.effectiveBoxPx.widthPx).toBe(768); // 先确认"算得出"，否则这条验收是在测一份没继承的夹具

    const file = await PptxFile.open(bytes);
    const made = await setShapeRunText(file, { partPath, shapeId: 2, paraIndex: 0, runIndex: 0 }, "改过的标题");
    if (!made.ok) throw new Error(JSON.stringify(made));
    applySlideEdit(file, made.op);
    const written = await file.partText(partPath);
    expect(written).toContain("<p:spPr/>");
    expect(written).not.toContain("7315200"); // 母版那个 cx 不许被写进幻灯片（写进去就切断母版联动）
    expect(await auditPptxPackage(file.toBytes())).toEqual([]);
  });
});
