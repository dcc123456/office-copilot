import { describe, it, expect, beforeEach } from "vitest";
import JSZip from "jszip";
import { DocxFile } from "./docx.js";
import type { PMDoc } from "./pm-schema.js";

async function makeDocx(
  paragraphs: string[],
  extra: {
    sectPr?: string;
    stylesXml?: string;
    /** 原样插进 body 的 <w:p>…</w:p>，用来造带 pPr 的真实段落。 */
    rawParagraphs?: string[];
  } = {},
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
    ${body}${(extra.rawParagraphs ?? []).join("")}
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

  /**
   * 真实文件层面的保真验收：打开带缩进的 docx → 走一遍编辑器保存路径
   * （toProseMirror / replaceProseMirror / toBytes）→ 重开读 document.xml。
   *
   * @param paragraphs 带 pPr 的原始 <w:p> 片段。
   * @returns 保存一次与保存两次后的 document.xml 文本。
   */
  async function saveTwice(
    paragraphs: string[],
  ): Promise<{ once: string; twice: string }> {
    const readBody = async (bytes: Uint8Array) => {
      const zip = await JSZip.loadAsync(bytes);
      return (await zip.file("word/document.xml")!.async("string"));
    };
    const reopen = async (bytes: Uint8Array): Promise<Uint8Array> => {
      const docx = await DocxFile.fromBytes(bytes);
      const loaded = await docx.toProseMirror();
      await docx.replaceProseMirror(loaded.doc);
      return docx.toBytes();
    };
    const source = await makeDocx([], { rawParagraphs: paragraphs });
    const onceBytes = await reopen(source);
    return {
      once: await readBody(onceBytes),
      twice: await readBody(await reopen(onceBytes)),
    };
  }

  it("keeps every indentation spelling through a real save", async () => {
    const { once, twice } = await saveTwice([
      `<w:p><w:pPr><w:ind w:left="300" w:right="200" w:leftChars="200"/><w:spacing w:line="360" w:lineRule="auto" w:beforeLines="50"/></w:pPr><w:r><w:t>合同右侧缩进</w:t></w:r></w:p>`,
      `<w:p><w:pPr><w:ind w:start="400" w:end="200"/></w:pPr><w:r><w:t>2010 写法</w:t></w:r></w:p>`,
      `<w:p><w:pPr><w:ind w:left="300" w:hanging="420"/></w:pPr><w:r><w:t>参考文献悬挂缩进</w:t></w:r></w:p>`,
    ]);
    expect(once).toContain('w:right="200"');
    expect(once).toContain('w:leftChars="200"');
    expect(once).toContain('w:beforeLines="50"');
    // 第二次保存只可能因为"我们把它归一化了"而少掉写法，不该少掉数值。
    expect(once).toContain('w:left="400"');
    expect(once).toContain('w:hanging="420"');
    expect(twice).toBe(once);
  });

  it("keeps the pagination switches and the paragraph mark run props through a save", async () => {
    const { once, twice } = await saveTwice([
      `<w:p><w:pPr><w:keepNext/><w:widowControl w:val="0"/><w:outlineLvl w:val="9"/><w:spacing w:after="120"/><w:rPr><w:sz w:val="28"/></w:rPr></w:pPr><w:r><w:t>标题候选</w:t></w:r></w:p>`,
    ]);
    expect(once).toContain("<w:keepNext/>");
    expect(once).toContain('<w:widowControl w:val="0"/>');
    expect(once).toContain('<w:outlineLvl w:val="9"/>');
    // 段落标记自己的字符属性排在最后，否则 CT_PPr 顺序不合法。
    expect(once.indexOf("<w:rPr")).toBeGreaterThan(once.indexOf("<w:spacing"));
    expect(twice).toBe(once);
  });

  it("keeps an imported paragraph border through two real saves", async () => {
    const { once, twice } = await saveTwice([
      `<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr></w:pPr><w:r><w:t>Word 里插入的分隔线</w:t></w:r></w:p>`,
    ]);
    expect(once).toContain("<w:pBdr>");
    expect(once).toContain('w:sz="6"');
    expect(twice).toBe(once);
  });

  it("keeps the named style a paragraph already carries through two real saves", async () => {
    const { once, twice } = await saveTwice([
      `<w:p><w:pPr><w:pStyle w:val="af3"/><w:ind w:left="300" w:right="200"/></w:pPr><w:r><w:t>公文正文</w:t></w:r></w:p>`,
    ]);
    // 只能有一条 pStyle，而且必须还是源文档那条：写侧不许把它换成推断的样式名。
    expect(once.match(/<w:pStyle\b[^>]*>/g)).toEqual([`<w:pStyle w:val="af3"/>`]);
    expect(once).toContain('w:right="200"');
    expect(twice).toBe(once);
  });

  /**
   * 编辑器里改完"这段落是什么"再存一遍，看样式引用有没有跟着改对。
   *
   * @param raw 源文档里的 `<w:p>` 片段。
   * @param edit 对加载出来的 PM JSON 的原地修改，模拟用户在编辑器里的操作。
   * @returns 保存后 document.xml 里全部 `<w:pStyle>` 片段；空数组表示没有样式引用。
   */
  async function styleRefsAfterEdit(
    raw: string,
    edit: (doc: PMDoc) => void,
  ): Promise<string[]> {
    const source = await makeDocx([], {
      rawParagraphs: [raw],
      stylesXml: `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:style w:type="paragraph" w:default="1" w:styleId="1"><w:name w:val="Normal"/></w:style>
        <w:style w:type="paragraph" w:styleId="2"><w:name w:val="heading 1"/></w:style>
        <w:style w:type="paragraph" w:styleId="3"><w:name w:val="heading 2"/></w:style>
      </w:styles>`,
    });
    const docx = await DocxFile.fromBytes(source);
    const loaded = await docx.toProseMirror();
    edit(loaded.doc);
    await docx.replaceProseMirror(loaded.doc);
    const saved = await JSZip.loadAsync(await docx.toBytes());
    const body = await saved.file("word/document.xml")!.async("string");
    return body.match(/<w:pStyle\b[^>]*>/g) ?? [];
  }

  it("re-points the heading style when the user changes its level", async () => {
    const heading = `<w:p><w:pPr><w:pStyle w:val="2"/></w:pPr><w:r><w:t>章节标题</w:t></w:r></w:p>`;
    // id "2" 在这份文件里声明的是"heading 1"；改成 2 级后 Word 必须收到 "3"，
    // 既不能继续用 "2"（段落会退回一级标题），也不能写文档里根本没有的 Heading2。
    expect(
      await styleRefsAfterEdit(heading, (doc) => {
        (doc.content![0]!.attrs as Record<string, unknown>).level = 2;
      }),
    ).toEqual([`<w:pStyle w:val="3"/>`]);
    expect(await styleRefsAfterEdit(heading, () => {})).toEqual([`<w:pStyle w:val="2"/>`]);
  });

  it("drops the heading style when the user converts a heading to body text", async () => {
    const refs = await styleRefsAfterEdit(
      `<w:p><w:pPr><w:pStyle w:val="2"/><w:spacing w:after="120"/></w:pPr><w:r><w:t>降级成正文</w:t></w:r></w:p>`,
      (doc) => {
        const first = doc.content![0]!;
        doc.content![0] = { ...first, type: "paragraph" };
      },
    );
    expect(refs).toEqual([]);
  });

  it("hands the style pane the paragraph styles this document declares", async () => {
    const withStyles = await DocxFile.fromBytes(
      await makeDocx(["正文"], {
        stylesXml: `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:style w:type="paragraph" w:default="1" w:styleId="1"><w:name w:val="Normal"/></w:style>
          <w:style w:type="paragraph" w:styleId="2"><w:name w:val="heading 1"/></w:style>
          <w:style w:type="character" w:styleId="3"><w:name w:val="Strong"/></w:style>
        </w:styles>`,
      }),
    );
    // 窗格要的是"这份文档里能套的样式"，不是 Word 的预置清单：字符样式不在段落窗格里。
    expect((await withStyles.toProseMirror()).styles).toEqual([
      { styleId: "1", name: "Normal", isDefault: true },
      { styleId: "2", name: "heading 1", headingLevel: 1, isDefault: false },
    ]);
    // 没带 styles.xml 的文件也要给数组，渲染层不用区分"读不到"和"读到空"。
    const bare = await DocxFile.fromBytes(await makeDocx(["正文"]));
    expect((await bare.toProseMirror()).styles).toEqual([]);
  });

  it("hands the renderer the resolved formatting of every declared paragraph style", async () => {
    const docx = await DocxFile.fromBytes(
      await makeDocx(["正文"], {
        stylesXml: `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:eastAsia="宋体"/></w:rPr></w:rPrDefault></w:docDefaults>
          <w:style w:type="paragraph" w:default="1" w:styleId="1"><w:name w:val="Normal"/>
            <w:pPr><w:jc w:val="both"/></w:pPr><w:rPr><w:sz w:val="21"/></w:rPr></w:style>
          <w:style w:type="paragraph" w:styleId="2"><w:name w:val="heading 1"/><w:basedOn w:val="1"/>
            <w:pPr><w:spacing w:before="340" w:line="576" w:lineRule="auto"/></w:pPr>
            <w:rPr><w:b/><w:sz w:val="44"/></w:rPr></w:style>
        </w:styles>`,
      }),
    );
    // 样式链在 docx-io 里合完再交给渲染层：单位仍是编辑器那套（pt / 倍数 / CSS px），
    // 渲染层只管把它们拼成选择器，不再出现第二处换算（AGENTS §2.1）。
    expect((await docx.toProseMirror()).styleFormats).toEqual([
      // 默认那条样式出现两次是刻意的：一次给"没有引用"的段落（styleId: null），
      // 一次给它自己的 id——Word 文档里大量段落就是显式写着 Normal 的。
      { styleId: null, format: { textAlign: "justify", fontFamily: "宋体", fontSize: "14px" } },
      { styleId: "1", format: { textAlign: "justify", fontFamily: "宋体", fontSize: "14px" } },
      {
        styleId: "2",
        format: {
          textAlign: "justify",
          spacingBefore: 17,
          lineHeight: "2.4",
          fontFamily: "宋体",
          fontSize: "29.3px",
          bold: true,
        },
      },
    ]);
    const bare = await DocxFile.fromBytes(await makeDocx(["正文"]));
    expect((await bare.toProseMirror()).styleFormats).toEqual([]);
  });

  /**
   * 编辑器里插入的分隔线必须落成段落底部边框：本仓不写 styles.xml，
   * 旧实现引用的 HorizontalRule 样式在文档里不存在，Word 只会排出空段落。
   */
  it("saves the editor's horizontal rule as a paragraph bottom border", async () => {
    const docx = await DocxFile.fromBytes(await makeDocx(["上文"]));
    const loaded = await docx.toProseMirror();
    await docx.replaceProseMirror({
      type: "doc",
      content: [...loaded.doc.content, { type: "horizontalRule" }],
    });
    const zip = await JSZip.loadAsync(docx.toBytes());
    const body = await zip.file("word/document.xml")!.async("string");
    expect(body).toContain(
      '<w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/>',
    );
    expect(body).not.toContain("HorizontalRule");
  });
});
