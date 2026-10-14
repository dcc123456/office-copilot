import { describe, expect, it } from "vitest";
import { auditPptxPackage } from "./audit.js";
import { PptxFile } from "./pptx-file.js";
import { applySlideEdit, readPresentation, setShapeRunText } from "./index.js";
import { buildDeckBytes, minimalDeckParts } from "./test-deck.js";

// 夹具的构造器在 `test-deck.ts`：桌面主进程出入口的单测用同一份，不许长两处（`AGENTS.md` §3.2）。

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
 * 所以拿同一套判定回头看两种产物：原始夹具、改过幻灯片文字后的写回。
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

  it.each([
    ["原始夹具", async () => buildDeckBytes(minimalDeckParts())],
    ["改过幻灯片文字的写回", bytesAfterTextEdit],
  ] as const)("%s：审计返回空数组", async (_label, produce) => {
    expect(await auditPptxPackage(await produce())).toEqual([]);
  });
});
