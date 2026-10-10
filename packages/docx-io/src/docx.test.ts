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
    /** 额外部件：zip 路径 → 内容。脚注/尾注这类旁路部件由调用方自己写全。 */
    parts?: Record<string, string>;
    /** 插进 `word/_rels/document.xml.rels` 的 `<Relationship>` 片段；给了才会生成该部件。 */
    documentRels?: string;
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
  // 真实 Word 包里这些部件在 [Content_Types].xml 里都有 Override；夹具省掉，
  // 正好让用例能断言"保存时登记被补上"，第二次保存再断言不重复登记。
  if (extra.documentRels) {
    zip.file(
      "word/_rels/document.xml.rels",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${extra.documentRels}</Relationships>`,
    );
  }
  for (const [path, content] of Object.entries(extra.parts ?? {})) {
    zip.file(path, content);
  }
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

  it("keeps a right tab stop with a dot leader through a real save", async () => {
    // 目录条目的"标题⇥页码"靠这条制表位对齐（排名 3 与排名 5 的共同前提）。
    // attrs 已经建模了它，第二次保存必须写出同一条，而不是从 rawPpr 再透传一份变成两条。
    const { once, twice } = await saveTwice([
      `<w:p><w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="8306"/></w:tabs>` +
        `<w:spacing w:after="60"/></w:pPr><w:r><w:t>第一章</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>3</w:t></w:r></w:p>`,
    ]);
    expect(once.match(/<w:tabs>/g)).toHaveLength(1);
    expect(once).toContain('<w:tab w:val="right" w:leader="dot" w:pos="8306"/>');
    expect(twice).toBe(once);
  });

  it("keeps a bookmark that spans two paragraphs through a real save", async () => {
    // 目录与交叉引用都靠书签定位（排名 3 的前提），所以这一对必须经
    // fromBytes → toProseMirror → replaceProseMirror → toBytes 整条路原样回来。
    const { once, twice } = await saveTwice([
      `<w:p><w:r><w:t>第一章</w:t></w:r><w:bookmarkStart w:id="0" w:name="_Toc1"/></w:p>`,
      `<w:p><w:bookmarkEnd w:id="0"/><w:r><w:t>正文</w:t></w:r></w:p>`,
    ]);
    expect(once).toContain(`<w:bookmarkStart w:id="0" w:name="_Toc1"/>`);
    expect(once).toContain(`<w:bookmarkEnd w:id="0"/>`);
    expect(twice).toBe(once);
  });

  it("keeps a TOC field with a nested PAGEREF through a real save", async () => {
    // Word 生成的目录条目里嵌着 PAGEREF 子域，缓存结果还带着制表符；
    // 这一整段必须经 fromBytes → toProseMirror → replaceProseMirror → toBytes 稳定回来。
    const toc =
      `<w:p>` +
      `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
      `<w:r><w:instrText xml:space="preserve"> TOC \\o &quot;1-3&quot; \\h </w:instrText></w:r>` +
      `<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
      `<w:hyperlink w:anchor="_Toc1">` +
      `<w:r><w:t>第一章</w:t></w:r><w:r><w:tab/></w:r>` +
      `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
      `<w:r><w:instrText xml:space="preserve"> PAGEREF _Toc1 </w:instrText></w:r>` +
      `<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>3</w:t></w:r>` +
      `<w:r><w:fldChar w:fldCharType="end"/></w:r></w:hyperlink>` +
      `<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`;
    const { once, twice } = await saveTwice([toc]);
    // 指令里的引号在文本位置（不是属性值），docx.ts 落盘前会重新序列化一遍，
    // `&quot;` 于是回到 `"` —— 这在 XML 文本里是合法的，别按属性位置的写法去断言。
    expect(once).toContain(`TOC \\o "1-3" \\h`);
    expect(once).toContain(`<w:t xml:space="preserve">第一章</w:t>`);
    expect(twice).toBe(once);
  });

  it("keeps a Word-generated multi-paragraph TOC field through a real save", async () => {
    // 真实 Word 存出来的目录：begin+instr+separate 在首段，end 在末段之后，
    // 每条条目自己的 PAGEREF 是段内域。跨段边界必须经
    // fromBytes → toProseMirror → replaceProseMirror → toBytes 保持在原来的段落里。
    const entry = (title: string, anchor: string, page: string) =>
      `<w:r><w:t>${title}</w:t></w:r><w:r><w:tab/></w:r>` +
      `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
      `<w:r><w:instrText xml:space="preserve"> PAGEREF ${anchor} </w:instrText></w:r>` +
      `<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>${page}</w:t></w:r>` +
      `<w:r><w:fldChar w:fldCharType="end"/></w:r>`;
    const { once, twice } = await saveTwice([
      `<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr>` +
        `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
        `<w:r><w:instrText xml:space="preserve"> TOC \\o &quot;1-3&quot; </w:instrText></w:r>` +
        `<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
        entry("第一章", "_Toc1", "3") +
        `</w:p>`,
      `<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr>` +
        entry("第二章", "_Toc2", "5") +
        `<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`,
    ]);
    expect(once).toContain(`TOC \\o "1-3"`);
    expect(once).toContain(`<w:pStyle w:val="TOC1"/>`);
    // 条目的制表符要回到 run 元素 `<w:tab/>`：留在 `<w:t>` 里只是一个空白字符，
    // Word 不会按制表位对齐，目录页码就贴不上右边界。
    expect(once).toContain(`<w:t xml:space="preserve">第一章</w:t><w:tab/></w:r>`);
    expect(once.match(/<w:tab\/>/g)).toHaveLength(2);
    // 三条子域 + 一条外层域 = 4 对 begin/end，且 end 记号在两段里各留一个（子域 2 + 外层 1）。
    expect(once.match(/fldCharType="begin"/g)).toHaveLength(3);
    expect(once.match(/fldCharType="end"/g)).toHaveLength(3);
    expect(once.indexOf(`fldCharType="end"`)).toBeLessThan(once.lastIndexOf(`fldCharType="end"`));
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
   * 编辑器里插入的分隔线必须落成段落底部边框：旧实现引用了一条文档里根本不存在的
   * HorizontalRule 样式，Word 只会排出空段落。
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

  /**
   * 排名 2 ⑥/⑦：窗格里"新建样式"要能真的存进文件。这里走完整往返——写 styles.xml →
   * 重新读回来 → 样式链解析，等价于保存后重开一次编辑器。
   */
  it("saves a created style into word/styles.xml and reads it back", async () => {
    const docx = await DocxFile.fromBytes(await makeDocx(["正文"]));
    const { styles, styleFormats } = await docx.applyStyleEdits([
      {
        name: "重点段落",
        paragraph: { indent: 36, spacingBefore: 12 },
        marks: [{ type: "textStyle", attrs: { fontSize: "24px", color: "#c00000" } }],
      },
    ]);
    expect(styles.map((s) => s.name)).toContain("重点段落");
    const out = await docx.toBytes();
    const zip = await JSZip.loadAsync(out);
    const stylesXml = await zip.file("word/styles.xml")?.async("string");
    expect(stylesXml).toContain('w:customStyle="1"');
    // 这份文档原本没有样式表，部件登记必须补全，否则 Word 判"文档损坏"。
    const contentTypes = await zip.file("[Content_Types].xml")!.async("string");
    expect(contentTypes).toContain('PartName="/word/styles.xml"');
    expect(contentTypes).toContain("wordprocessingml.styles+xml");
    const rels = await zip.file("word/_rels/document.xml.rels")!.async("string");
    expect(rels).toContain('Target="styles.xml"');
    const reopened = await DocxFile.fromBytes(out);
    const reloaded = await reopened.toProseMirror();
    expect(reloaded.styles.map((s) => s.name)).toContain("重点段落");
    expect(reloaded.styleFormats).toEqual(styleFormats);
  });

  /**
   * 主进程 `saveDocx` 的写入顺序（样式表先落地，再写正文）钉在这条用例里：
   * 正文引用一条刚刚新建的样式时，**声明与引用必须同时出现在保存出来的文件里**。
   * 只有 `w:pStyle` 没有 `w:style`，Word 会把这段排成本文档的默认格式（排名 2 ⑥/⑦）。
   */
  it("saves both the declaration and the body reference of a created style", async () => {
    const docx = await DocxFile.fromBytes(await makeDocx(["正文"]));
    await docx.applyStyleEdits([
      { styleId: "Emphasis", name: "重点段落", paragraph: { indent: 36 }, marks: [] },
    ]);
    await docx.replaceProseMirror({
      type: "doc",
      content: [
        {
          type: "paragraph",
          attrs: { styleId: "Emphasis" },
          content: [{ type: "text", text: "重点" }],
        },
      ],
    } as PMDoc);
    const zip = await JSZip.loadAsync(await docx.toBytes());
    expect(await zip.file("word/document.xml")!.async("string")).toContain(
      '<w:pStyle w:val="Emphasis"/>',
    );
    expect(await zip.file("word/styles.xml")!.async("string")).toContain(
      'w:styleId="Emphasis"',
    );
  });

  /** 已有样式表时按 splice-and-keep 合并：docDefaults 那条链不能被保存一次就抹平。 */
  it("keeps docDefaults alive when a style is edited", async () => {
    const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:pPrDefault><w:pPr><w:spacing w:after="120"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="1"><w:name w:val="Normal"/></w:style></w:styles>`;
    const docx = await DocxFile.fromBytes(await makeDocx(["正文"], { stylesXml }));
    await docx.applyStyleEdits([
      { name: "Loose", paragraph: { indent: 36 }, marks: [] },
    ]);
    const reopened = await DocxFile.fromBytes(await docx.toBytes());
    const { styleFormats } = await reopened.toProseMirror();
    const created = styleFormats.find((entry) => entry.styleId === "Loose");
    expect(created?.format).toMatchObject({ indent: 36, spacingAfter: 6 });
  });

  /** 一次编辑都没产出时不该凭空造一份 styles.xml。 */
  it("leaves the package untouched when no style edit is valid", async () => {
    const docx = await DocxFile.fromBytes(await makeDocx(["正文"]));
    await docx.applyStyleEdits([{ name: "   ", paragraph: {}, marks: [] }]);
    const zip = await JSZip.loadAsync(await docx.toBytes());
    expect(zip.file("word/styles.xml")).toBeNull();
    // 关系与内容类型也不许出现悬空引用：Word 会因为指向不存在的部件而判文档损坏。
    expect(await zip.file("[Content_Types].xml")!.async("string")).not.toContain("/word/styles.xml");
  });
});

/**
 * 排名 3 ④：脚注 / 尾注部件。
 *
 * 正文里只有一个引用点 `<w:footnoteReference w:id/>`，脚注正文住在
 * `word/footnotes.xml`，靠同一个 id 对应；部件里还带着分隔符那几条 Word 自己
 * 用的条目，我们没有建模，必须原样带走（§2.7：判不了角色的不产生改动）。
 */
const FN_REL =
  '<Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>';
const EN_REL =
  '<Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/>';

/** Word 写出来的脚注部件：分隔符条目在前，正文条目开头带编号记号。 */
const FOOTNOTES_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
  `<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>` +
  `<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>` +
  `<w:footnote w:id="1"><w:p>` +
  `<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r>` +
  `<w:r><w:t xml:space="preserve">参见《民法典》第五百零九条。</w:t></w:r></w:p></w:footnote>` +
  `</w:footnotes>`;

const ENDNOTES_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
  `<w:endnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:endnote>` +
  `<w:endnote w:id="2"><w:p>` +
  `<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteRef/></w:r>` +
  `<w:r><w:t xml:space="preserve">尾注正文。</w:t></w:r></w:p></w:endnote>` +
  `</w:endnotes>`;

/** 正文两段，各带一个引用点。 */
const REF_PARAGRAPHS = [
  `<w:p><w:r><w:t>正文甲</w:t></w:r>` +
    `<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="1"/></w:r></w:p>`,
  `<w:p><w:r><w:t>正文乙</w:t></w:r>` +
    `<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteReference w:id="2"/></w:r></w:p>`,
];

function footnotedDocx(): Promise<Uint8Array> {
  return makeDocx([], {
    rawParagraphs: REF_PARAGRAPHS,
    documentRels: FN_REL + EN_REL,
    parts: { "word/footnotes.xml": FOOTNOTES_XML, "word/endnotes.xml": ENDNOTES_XML },
  });
}

describe("脚注与尾注部件往返", () => {
  /** 打开 → 走一遍编辑器保存路径 → 重开，返回保存出来的包。 */
  async function savedOnce(
    bytes: Uint8Array,
    edit?: (doc: PMDoc) => PMDoc,
  ): Promise<Uint8Array> {
    const docx = await DocxFile.fromBytes(bytes);
    const loaded = await docx.toProseMirror();
    await docx.replaceProseMirror(edit ? edit(loaded.doc) : loaded.doc);
    return docx.toBytes();
  }

  const part = async (bytes: Uint8Array, path: string): Promise<string> => {
    const zip = await JSZip.loadAsync(bytes);
    return (await zip.file(path)?.async("string")) ?? "";
  };

  it("读回引用点，脚注正文与分隔符条目一起进 store", async () => {
    const docx = await DocxFile.fromBytes(await footnotedDocx());
    const loaded = await docx.toProseMirror();
    expect(loaded.doc.content[0]!.content).toEqual([
      { type: "text", text: "正文甲" },
      { type: "footnoteReference", attrs: { id: "1" } },
    ]);
    expect(loaded.doc.content[1]!.content?.at(-1)).toMatchObject({
      type: "endnoteReference",
      attrs: { id: "2" },
    });
    expect(loaded.footnotes.footnote.map((e) => e.id)).toEqual(["-1", "0", "1"]);
    expect(loaded.footnotes.footnote[0]!.role).toBe("separator");
    expect(loaded.footnotes.endnote.map((e) => e.id)).toEqual(["-1", "2"]);
    expect(loaded.footnotes.unmodeled).toEqual([]);
  });

  it("保存一次就补齐 Content_Types 与关系，且分隔符与正文条目原样带走", async () => {
    const out = await savedOnce(await footnotedDocx());
    const body = await part(out, "word/document.xml");
    expect(body).toContain('<w:footnoteReference w:id="1"/>');
    expect(body).toContain('<w:endnoteReference w:id="2"/>');
    const fn = await part(out, "word/footnotes.xml");
    expect(fn).toContain('<w:footnote w:type="separator" w:id="-1">');
    expect(fn).toContain("参见《民法典》第五百零九条。");
    expect(fn).toContain("<w:footnoteRef/>");
    // 保存时登记被补上：夹具里没写 Override，Word 认的是包里的登记。
    const zip = await JSZip.loadAsync(out);
    const ct = await zip.file("[Content_Types].xml")!.async("string");
    expect(ct).toContain('PartName="/word/footnotes.xml"');
    expect(ct).toContain("wordprocessingml.footnotes+xml");
    expect(ct).toContain("wordprocessingml.endnotes+xml");
    const rels = await zip.file("word/_rels/document.xml.rels")!.async("string");
    expect(rels.match(/relationships\/footnotes/g)).toHaveLength(1);
    expect(rels.match(/relationships\/endnotes/g)).toHaveLength(1);
  });

  it("保存两次不再多登记，部件内容逐字节稳定", async () => {
    const once = await savedOnce(await footnotedDocx());
    const twice = await savedOnce(once);
    expect(await part(twice, "word/footnotes.xml")).toBe(await part(once, "word/footnotes.xml"));
    const zip = await JSZip.loadAsync(twice);
    const ct = await zip.file("[Content_Types].xml")!.async("string");
    expect(ct.match(/\/word\/footnotes\.xml/g)).toHaveLength(1);
  });

  it("删掉正文里的引用点，部件里那条脚注也一起消失，分隔符留下", async () => {
    const out = await savedOnce(await footnotedDocx(), (doc) => ({
      ...doc,
      content: doc.content.map((p) => ({
        ...p,
        content: (p.content ?? []).filter((i) => i.type !== "footnoteReference"),
      })),
    }));
    const fn = await part(out, "word/footnotes.xml");
    expect(fn).not.toContain("参见《民法典》");
    expect(fn).toContain('<w:footnote w:type="separator" w:id="-1">');
    // 尾注没被牵连。
    expect(await part(out, "word/endnotes.xml")).toContain("尾注正文。");
  });

  it("正文插入一个新引用点，部件补一条空正文条目并带上编号记号", async () => {
    const out = await savedOnce(await footnotedDocx(), (doc) => ({
      ...doc,
      content: [
        ...doc.content,
        {
          type: "paragraph",
          content: [
            { type: "text", text: "新脚注" },
            { type: "footnoteReference", attrs: { id: "9" } },
          ],
        },
      ],
    }));
    const fn = await part(out, "word/footnotes.xml");
    expect(fn).toContain('<w:footnote w:id="9">');
    expect(fn).toContain("<w:footnoteRef/>");
    expect(await part(out, "word/document.xml")).toContain('<w:footnoteReference w:id="9"/>');
  });

  it("部件里没有建模不了的条目时，整个部件原样透传不重写", async () => {
    const out = await savedOnce(
      await makeDocx([], {
        rawParagraphs: [REF_PARAGRAPHS[0]!],
        documentRels: FN_REL,
        parts: {
          "word/footnotes.xml": FOOTNOTES_XML.replace('w:id="1"', 'w:id="not-a-number"'),
        },
      }),
    );
    // id 不是十进制 → 这个部件不能按条目重建，否则那一条就丢了。
    expect(await part(out, "word/footnotes.xml")).toContain('w:id="not-a-number"');
  });

  it("本来没有脚注的文档，保存后不会凭空多出部件", async () => {
    const out = await savedOnce(await makeDocx(["只有一段"]));
    expect(await part(out, "word/footnotes.xml")).toBe("");
    const zip = await JSZip.loadAsync(out);
    expect(await zip.file("[Content_Types].xml")!.async("string")).not.toContain("footnotes+xml");
  });

  it("正文段落顺序与文字都不因脚注改动", async () => {
    const docx = await DocxFile.fromBytes(await footnotedDocx());
    const loaded = await docx.toProseMirror();
    expect(loaded.doc.content.map((p) => p.content?.map((i) => (i.type === "text" ? i.text : i.type)))).toEqual([
      ["正文甲", "footnoteReference"],
      ["正文乙", "endnoteReference"],
    ]);
  });
});
