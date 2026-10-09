import { describe, it, expect } from "vitest";
import JSZip from "jszip";
import { ooxmlToPm, pmToOoxml } from "./ooxml-pm.js";
import { DocxFile } from "./docx.js";
import type { PMDoc, PMNode } from "./pm-schema.js";

/**
 * Fixtures mirror the shape Word writes for a text-box document: an
 * `mc:AlternateContent` whose `mc:Choice` branch is a `wps:` group of boxes and
 * whose `mc:Fallback` branch is the same boxes as VML. Two boxes, each carrying
 * a `w14:paraId` we can watch for, and each appearing twice per branch.
 */
const NS = [
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"',
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"',
  'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"',
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"',
  'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"',
  'xmlns:v="urn:schemas-microsoft-com:vml"',
  'xmlns:o="urn:schemas-microsoft-com:office:office"',
  'xmlns:w10="urn:schemas-microsoft-com:office:word"',
].join(" ");

function doc(body: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${NS}><w:body>${body}<w:sectPr/></w:body></w:document>`;
}

/**
 * One anchor at (0, 10px) holding a group of two boxes: 姓名 (no fill, centre
 * anchored) at the group origin and 电话 (grey fill, black border) 10px right.
 */
const WSP_NAME = `<wps:wsp><wps:cNvPr id="3" name="文本框 3"/><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="95250" cy="95250"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></wps:spPr><wps:txbx><w:txbxContent><w:p w14:paraId="11111111"><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="36"/></w:rPr><w:t>姓名</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr rot="0" vert="horz" wrap="square" lIns="38100" tIns="19050" rIns="38100" bIns="19050" anchor="ctr" anchorCtr="0"><a:noAutofit/></wps:bodyPr></wps:wsp>`;

const WSP_PHONE = `<wps:wsp><wps:cNvPr id="4" name="文本框 4"/><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="95250" y="0"/><a:ext cx="95250" cy="95250"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="F0F0F0"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></wps:spPr><wps:txbx><w:txbxContent><w:p w14:paraId="22222222"><w:r><w:t>电话</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr rot="0" vert="horz" wrap="square" anchor="t" anchorCtr="0"><a:noAutofit/></wps:bodyPr></wps:wsp>`;

const VML = `<w:pict><v:shape id="文本框 3" style="position:absolute;left:0;top:10pt;width:10pt;height:10pt" fillcolor="white" stroked="f" inset="0" alt=""><v:textbox inset="0,0,0,0"><w:txbxContent><w:p><w:r><w:t>姓名</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape><v:shape id="文本框 4" style="position:absolute;left:100pt;top:10pt;width:10pt;height:10pt" fillcolor="#F0F0F0" stroked="t" strokecolor="black #000000" inset="0" alt=""><v:textbox inset="0,0,0,0"><w:txbxContent><w:p><w:r><w:t>电话</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict>`;

const GROUP = `<mc:AlternateContent><mc:Choice Requires="wpg"><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="251661312" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>95250</wp:posOffset></wp:positionV><wp:extent cx="190500" cy="95250"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="2" name="分组 2"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"><wpg:grpSp><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="190500" cy="95250"/><a:chOff x="0" y="0"/><a:chExt cx="190500" cy="95250"/></a:xfrm></wpg:grpSpPr>${WSP_NAME}${WSP_PHONE}</wpg:grpSp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice><mc:Fallback>${VML}</mc:Fallback></mc:AlternateContent>`;

/** A paragraph that exists only to anchor the group — the resume-template case. */
const ANCHOR_PARA = `<w:p><w:pPr><w:rPr><w:noProof/></w:rPr></w:pPr><w:r><w:rPr><w:noProof/></w:rPr>${GROUP}</w:r></w:p>`;

const count = (hay: string, needle: string): number =>
  hay.split(needle).length - 1;

function boxes(group: PMNode | undefined): Extract<PMNode, { type: "floatBox" }>[] {
  expect(group?.type).toBe("floatGroup");
  const content = (group as Extract<PMNode, { type: "floatGroup" }>).content ?? [];
  return content.filter(
    (b): b is Extract<PMNode, { type: "floatBox" }> => b.type === "floatBox",
  );
}

function pmText(node: PMNode): string {
  if (node.type === "text") return node.text;
  return (node.content ?? []).map(pmText).join("");
}

describe("ooxmlToPm: floating text boxes", () => {
  it("hoists boxes out of a paragraph that only anchors them", () => {
    const pm = ooxmlToPm(doc(ANCHOR_PARA));
    // The boxes come first, then the paragraph that carries them: Word measures
    // every anchor offset from that paragraph's own top, so it has to keep its
    // line in the flow even when all it shows is the anchor.
    expect(pm.content.map((n) => n.type)).toEqual(["floatGroup", "paragraph"]);
    const bs = boxes(pm.content[0]);
    expect(bs).toHaveLength(2);
    expect(pmText(bs[0]!)).toBe("姓名");
    expect(pmText(bs[1]!)).toBe("电话");
  });

  it("keeps each of two anchor paragraphs at its own line", () => {
    // Two groups stacked one after the other: if the host paragraphs were
    // dropped they would share a single flow position and overlap on the page.
    const pm = ooxmlToPm(doc(ANCHOR_PARA + ANCHOR_PARA));
    expect(pm.content.map((n) => n.type)).toEqual([
      "floatGroup",
      "paragraph",
      "floatGroup",
      "paragraph",
    ]);
  });

  it("maps anchor and shape geometry into px", () => {
    const pm = ooxmlToPm(doc(ANCHOR_PARA));
    const group = pm.content[0] as Extract<PMNode, { type: "floatGroup" }>;
    expect(group.attrs).toMatchObject({ x: 0, y: 10, relH: "column", relV: "paragraph" });
    const bs = boxes(group);
    // 95250 EMU = 10 px at 96 dpi; the second box sits one box-width right.
    expect(bs[0]!.attrs).toMatchObject({ x: 0, y: 0, w: 10, h: 10 });
    expect(bs[1]!.attrs).toMatchObject({ x: 10, y: 0, w: 10, h: 10 });
  });

  it("carries fill, stroke and vertical alignment", () => {
    const bs = boxes(ooxmlToPm(doc(ANCHOR_PARA)).content[0]);
    expect(bs[0]!.attrs).toMatchObject({ fill: null, stroke: null, vAlign: "center" });
    expect(bs[1]!.attrs).toMatchObject({
      fill: "#F0F0F0",
      stroke: "#000000",
      strokeWidth: 1,
      vAlign: "flex-start",
    });
  });

  it("keeps the host paragraph and its text when a paragraph mixes both", () => {
    const pm = ooxmlToPm(
      doc(
        `<w:p><w:r><w:t>正文</w:t></w:r><w:r><w:rPr><w:noProof/></w:rPr>${GROUP}</w:r></w:p>`,
      ),
    );
    expect(pm.content).toHaveLength(2);
    expect(pm.content[0]!.type).toBe("floatGroup");
    expect(pmText(pm.content[1]!)).toBe("正文");
  });

  it("keeps a page break that trails the anchors behind them", () => {
    // The resume template parks the break at the end of the paragraph that
    // carries a page's worth of boxes. Hoisting it ahead of them shifts every
    // following page by one.
    const pm = ooxmlToPm(
      doc(
        `<w:p><w:r><w:rPr><w:noProof/></w:rPr>${GROUP}</w:r><w:r><w:br w:type="page"/></w:r></w:p>`,
      ),
    );
    expect(pm.content.map((n) => n.type)).toEqual(["floatGroup", "paragraph"]);
    const para = pm.content[1] as Extract<PMNode, { type: "paragraph" }>;
    expect((para.content ?? []).map((c) => c.type)).toEqual(["pageBreak"]);
  });

  it("reads the legacy w:pict form from its VML style", () => {
    const pm = ooxmlToPm(
      doc(`<w:p><w:r><w:rPr><w:noProof/></w:rPr>${VML}</w:r></w:p>`),
    );
    const bs = boxes(pm.content[0]);
    expect(bs).toHaveLength(2);
    // 100pt → 133px.
    expect(bs[1]!.attrs.x).toBe(133);
    expect(bs[0]!.attrs.fill).toBe("white");
  });

  it("leaves a picture anchor to the image path instead of a float group", () => {
    const pic = `<w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1"><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="95250" cy="95250"/><wp:wrapNone/><wp:docPr id="9" name="图片 9"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="9" name="a.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId5"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="95250" cy="95250"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing>`;
    const pm = ooxmlToPm(doc(`<w:p><w:r><w:rPr><w:noProof/></w:rPr>${pic}</w:r></w:p>`));
    expect(pm.content[0]!.type).toBe("paragraph");
    const inline = (pm.content[0] as { content?: PMNode[] }).content ?? [];
    expect(inline[0]).toMatchObject({
      type: "image",
      attrs: { layout: "anchor", src: "relationship:rId5" },
    });
  });
});

describe("pmToOoxml: floating text boxes", () => {
  it("writes an untouched document back byte-for-byte", () => {
    const pm = ooxmlToPm(doc(ANCHOR_PARA));
    const xml = pmToOoxml(pm);
    // Both boxes keep their source bodies, so both paraIds survive…
    expect(xml).toContain('w14:paraId="11111111"');
    expect(xml).toContain('w14:paraId="22222222"');
    // …and so does the branch Word uses when it cannot draw a group.
    expect(xml).toContain("<mc:Fallback>");
    expect(xml).toContain("<wpg:grpSp>");
    // Each box appears once per branch, never a third time.
    expect(count(xml, "姓名")).toBe(2);
    expect(count(xml, "电话")).toBe(2);
    expect(xml).toContain("<w:noProof/>");
  });

  it("rebuilds only the box whose text was edited", () => {
    const pm = ooxmlToPm(doc(ANCHOR_PARA));
    const group = pm.content[0] as Extract<PMNode, { type: "floatGroup" }>;
    const edited: PMNode = {
      type: "floatBox",
      attrs: { ...group.content![0]!.attrs, dirty: true },
      content: [{ type: "paragraph", content: [{ type: "text", text: "丁长长" }] }],
    };
    const next: PMDoc = {
      type: "doc",
      content: [{ ...group, content: [edited, group.content![1]!] }],
    };
    const xml = pmToOoxml(next);
    expect(count(xml, "丁长长")).toBe(2); // Choice and Fallback stay in step
    expect(xml).not.toContain('w14:paraId="11111111"');
    expect(xml).toContain('w14:paraId="22222222"');
    expect(count(xml, "电话")).toBe(2);
  });

  it("declares every drawing prefix it may emit", () => {
    const xml = pmToOoxml(ooxmlToPm(doc(ANCHOR_PARA)));
    for (const prefix of ["mc", "wpg", "wps", "v", "o", "w10", "w14"]) {
      expect(xml).toContain(`xmlns:${prefix}=`);
    }
  });
});

describe("ooxmlToPm: decorative shapes", () => {
  /** The resume's section headers underline themselves with a plain line. */
  const LINE = `<mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="476250" cy="285750"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="9" name="直接连接符 9"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvPr id="9" name="直接连接符 9"/><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="476250" cy="0"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:ln w="12700"><a:solidFill><a:srgbClr val="0B70BD"/></a:solidFill></a:ln></wps:spPr><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice></mc:AlternateContent>`;

  it("keeps a connector line at its own zero height", () => {
    // A `cy` of zero is the shape saying "I am a line", not "I have no size",
    // so it must not borrow the anchor's 30px extent — that turned every
    // underline into a box drawn across the text below it.
    const pm = ooxmlToPm(doc(`<w:p><w:r>${LINE}</w:r></w:p>`));
    const bs = boxes(pm.content[0]);
    expect(bs).toHaveLength(1);
    expect(bs[0]!.attrs.h).toBe(0);
    expect(bs[0]!.attrs.w).toBe(50);
    expect(bs[0]!.attrs.stroke).toBe("#0B70BD");
  });
});

describe("DocxFile round-trip with text boxes", () => {
  it("preserves box text without duplicating the fallback branch", async () => {
    const zip = new JSZip();
    zip.file(
      "[Content_Types].xml",
      `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
    );
    zip.file(
      "_rels/.rels",
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="word/document.xml"/></Relationships>`,
    );
    zip.file(
      "word/_rels/document.xml.rels",
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`,
    );
    zip.file(
      "word/document.xml",
      doc(`${ANCHOR_PARA}<w:p><w:r><w:t>正文</w:t></w:r></w:p>`),
    );
    const bytes = await zip.generateAsync({ type: "uint8array" });
    const docx = await DocxFile.fromBytes(bytes);
    const { doc: pm } = await docx.toProseMirror();

    expect(boxes(pm.content[0])).toHaveLength(2);
    await docx.replaceProseMirror(pm);
    const reopen = await DocxFile.fromBytes(await docx.toBytes());

    // The saved file reopens with the same two boxes and the same text.
    const again = (await reopen.toProseMirror()).doc;
    expect(boxes(again.content[0]).map(pmText)).toEqual(["姓名", "电话"]);

    // The listing sees the host paragraph once and no VML duplicate, so the
    // outline cannot show every box twice.
    const texts = reopen.listParagraphs().map((p) => p.text);
    expect(texts).toContain("正文");
    expect(count(texts.join("|"), "姓名")).toBe(1);
    expect(texts).toContain("姓名");
  });
});
