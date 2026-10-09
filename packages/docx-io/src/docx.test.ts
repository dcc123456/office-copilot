import { describe, it, expect, beforeEach } from "vitest";
import JSZip from "jszip";
import { DocxFile } from "./docx.js";

async function makeDocx(
  paragraphs: string[],
  extra: { sectPr?: string; stylesXml?: string } = {},
): Promise<Uint8Array> {
  const zip = new JSZip();
  const body = paragraphs
    .map((t) => `<w:p><w:r><w:t>${escape(t)}</w:t></w:r></w:p>`)
    .join("");
  zip.file(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
  <w:body>
    ${body}
    <w:sectPr>${extra.sectPr ?? '<w:pgSz w:w="12240" w:h="15840"/>'}</w:sectPr>
  </w:body>
</w:document>`,
  );
  if (extra.stylesXml) zip.file("word/styles.xml", extra.stylesXml);
  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`,
  );
  zip.file(
    "_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`,
  );
  return zip.generateAsync({ type: "uint8array" });
}

function escape(s: string) {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

describe("DocxFile", () => {
  let bytes: Uint8Array;

  beforeEach(async () => {
    bytes = await makeDocx([
      "甲方应在约定期限内交付产品。",
      "甲方对产品质量负责。",
      "本合同一式两份。",
    ]);
  });

  it("lists paragraphs in order", async () => {
    const d = await DocxFile.fromBytes(bytes);
    const paras = d.listParagraphs();
    expect(paras).toHaveLength(3);
    expect(paras[0]!.text).toContain("甲方");
  });

  it("replaces literal text and counts matches", async () => {
    const d = await DocxFile.fromBytes(bytes);
    const res = d.replaceText({ find: "甲方", replace: "客户" });
    expect(res.count).toBe(2);
    expect(res.previews).toHaveLength(2);
    const out = d.listParagraphs();
    expect(out[0]!.text).toContain("客户");
    expect(out[1]!.text).toContain("客户");
    expect(out[0]!.text).not.toContain("甲方");
  });

  it("returns 0 replacements when nothing matches", async () => {
    const d = await DocxFile.fromBytes(bytes);
    const res = d.replaceText({ find: "不存在", replace: "x" });
    expect(res.count).toBe(0);
  });

  it("supports regex replacement", async () => {
    const d = await DocxFile.fromBytes(bytes);
    const res = d.replaceText({ find: "甲.?", replace: "客户", regex: true });
    expect(res.count).toBeGreaterThan(0);
    const text = d.listParagraphs().map((p) => p.text).join("\n");
    expect(text).not.toContain("甲方");
  });

  it("appends a paragraph after sectPr", async () => {
    const d = await DocxFile.fromBytes(bytes);
    d.insertParagraph({ text: "附加条款", position: "end" });
    const paras = d.listParagraphs();
    expect(paras).toHaveLength(4);
    expect(paras.at(-1)!.text).toBe("附加条款");
  });

  it("inserts a paragraph at the start", async () => {
    const d = await DocxFile.fromBytes(bytes);
    d.insertParagraph({ text: "前言", position: "start" });
    const paras = d.listParagraphs();
    expect(paras[0]!.text).toBe("前言");
  });

  it("inserts before/after a specific anchor", async () => {
    const d = await DocxFile.fromBytes(bytes);
    d.insertParagraph({ text: "BEFORE", position: "before", anchor: 1 });
    d.insertParagraph({ text: "AFTER", position: "after", anchor: 2 });
    const paras = d.listParagraphs().map((p) => p.text);
    expect(paras[1]).toBe("BEFORE");
    // Original paragraph at index 1 ("甲方对...") has shifted to index 2;
    // AFTER is inserted after original index 2 ("本合同..."), which is now 3.
    expect(paras[3]).toBe("AFTER");
  });

  it("inserts a table with bold header", async () => {
    const d = await DocxFile.fromBytes(bytes);
    d.insertTable({ rows: [["A", "B"], ["1", "2"]], headerRow: true });
    const out = await d.toBytes();
    // Reopen and confirm the table content survived the round-trip.
    const d2 = await DocxFile.fromBytes(out);
    const text = d2.listParagraphs().map((p) => p.text).join("\n");
    expect(text).toContain("A");
    expect(text).toContain("1");
  });

  it("snapshot and restore undoes an edit", async () => {
    const d = await DocxFile.fromBytes(bytes);
    const snap = d.getSnapshot();
    d.replaceText({ find: "甲方", replace: "客户" });
    expect(d.listParagraphs()[0]!.text).toContain("客户");
    d.restoreSnapshot(snap);
    expect(d.listParagraphs()[0]!.text).toContain("甲方");
  });

  it("round-trips an edited document through toBytes", async () => {
    const d = await DocxFile.fromBytes(bytes);
    d.replaceText({ find: "甲方", replace: "客户" });
    d.insertParagraph({ text: "再见", position: "end" });
    const out = await d.toBytes();
    const d2 = await DocxFile.fromBytes(out);
    const paras = d2.listParagraphs();
    expect(paras[0]!.text).toContain("客户");
    expect(paras.at(-1)!.text).toBe("再见");
  });

  it("exports the line grid and default paragraph spacing", async () => {
    const grid = await DocxFile.fromBytes(
      await makeDocx(["正文"], {
        sectPr: `<w:pgSz w:w="11906" w:h="16838"/><w:docGrid w:type="lines" w:linePitch="312"/>`,
        stylesXml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:default="1" w:styleId="1"><w:name w:val="Normal"/><w:pPr><w:spacing w:after="120"/></w:pPr></w:style></w:styles>`,
      }),
    );
    // 312 twips of grid = 20.8px at 96dpi; 120 twentieths of a point = 6pt = 8px.
    expect((await grid.toProseMirror()).section).toMatchObject({
      linePitch: 20.8,
      paraSpacingAfter: 8,
    });
  });

  it("leaves the flow metrics unset without a line grid", async () => {
    const plain = await DocxFile.fromBytes(await makeDocx(["正文"]));
    const { section } = await plain.toProseMirror();
    expect(section?.linePitch).toBeUndefined();
    expect(section?.paraSpacingAfter).toBe(0);
  });
});
