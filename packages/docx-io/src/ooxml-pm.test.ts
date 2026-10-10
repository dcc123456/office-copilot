import { describe, it, expect } from "vitest";
import JSZip from "jszip";
import {
  ooxmlToPm,
  pmToOoxml,
  ptToCssFontSize,
  cssFontSizeToPt,
  ptToCssLetterSpacing,
  cssPxToPt,
  cmToPt,
  ptToCm,
  cssColorToOoxmlHex,
  canonColor,
} from "./ooxml-pm.js";
import { StyleMap } from "./styles.js";
import type { PPrAttrs, PMDoc, TabStop } from "./pm-schema.js";
import { DocxFile } from "./docx.js";

function bodyXml(...inner: string[]): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${inner.join("")}<w:sectPr/></w:body></w:document>`;
}

/**
 * 打出一个能被 `DocxFile` 读回来的最小 .docx 字节。
 *
 * @param documentXml `word/document.xml` 的完整内容（含 XML 声明）。
 * @returns zip 字节，交给 `DocxFile.fromBytes`。
 */
async function minimalDocx(documentXml: string): Promise<Uint8Array> {
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
  zip.file("word/document.xml", documentXml);
  return zip.generateAsync({ type: "uint8array" });
}

describe("ooxmlToPm", () => {
  it("parses a plain paragraph", () => {
    const pm = ooxmlToPm(bodyXml(`<w:p><w:r><w:t>Hello</w:t></w:r></w:p>`));
    expect(pm.content).toHaveLength(1);
    const p = pm.content[0]!;
    expect(p.type).toBe("paragraph");
    expect(p.content?.[0]).toMatchObject({ type: "text", text: "Hello" });
  });

  it("parses headings by pStyle", () => {
    const pm = ooxmlToPm(
      bodyXml(
        `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Title</w:t></w:r></w:p>`,
      ),
    );
    expect(pm.content[0]).toMatchObject({
      type: "heading",
      attrs: { level: 1 },
    });
  });

  it("keeps heading levels 4 to 6 both ways", () => {
    for (const level of [4, 5, 6]) {
      const pm = ooxmlToPm(
        bodyXml(
          `<w:p><w:pPr><w:pStyle w:val="Heading${level}"/></w:pPr><w:r><w:t>Deep</w:t></w:r></w:p>`,
        ),
      );
      expect(pm.content[0]).toMatchObject({ type: "heading", attrs: { level } });
      expect(pmToOoxml(pm)).toContain(`<w:pStyle w:val="Heading${level}"/>`);
    }
  });

  it("reads a heading level styles.xml declares under its own name", () => {
    const styles = StyleMap.fromXml(
      `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="a3"><w:name w:val="heading 5"/></w:style></w:styles>`,
    );
    const pm = ooxmlToPm(
      bodyXml(`<w:p><w:pPr><w:pStyle w:val="a3"/></w:pPr><w:r><w:t>Five</w:t></w:r></w:p>`),
      undefined,
      styles,
    );
    expect(pm.content[0]).toMatchObject({ type: "heading", attrs: { level: 5 } });
  });

  it("still refuses to promote an undeclared numeric style above level 3", () => {
    // A bare digit says nothing about depth — the WPS resume that made 174 body
    // paragraphs into 26px titles is why this guess stays capped at 3 while the
    // named-heading path widened to 6.
    const pm = ooxmlToPm(
      bodyXml(`<w:p><w:pPr><w:pStyle w:val="5"/></w:pPr><w:r><w:t>Body</w:t></w:r></w:p>`),
    );
    expect(pm.content[0]).toMatchObject({ type: "heading", attrs: { level: 3 } });
  });

  it("does not guess a heading from a style styles.xml declares", () => {
    // WPS names its list paragraph style "1", and the resume runs 174
    // paragraphs through it. Guessing a heading from the digits alone turned
    // every body line into 26px title text piled over the next one.
    const styles = StyleMap.fromXml(
      `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="1"><w:name w:val="列表段落1"/><w:basedOn w:val="a"/></w:style></w:styles>`,
    );
    const pm = ooxmlToPm(
      bodyXml(
        `<w:p><w:pPr><w:pStyle w:val="1"/></w:pPr><w:r><w:t>Skill</w:t></w:r></w:p>`,
      ),
      undefined,
      styles,
    );
    expect(pm.content[0]!.type).toBe("paragraph");
  });

  it("captures bold/italic/underline/strike marks", () => {
    const pm = ooxmlToPm(
      bodyXml(
        `<w:p><w:r><w:rPr><w:b/><w:i/><w:u w:val="single"/><w:strike/></w:rPr><w:t>x</w:t></w:r></w:p>`,
      ),
    );
    const text = (pm.content[0] as { content: Array<{ marks?: unknown[] }> }).content[0]!;
    const marks = (text.marks ?? []).map((m) => (m as { type: string }).type);
    expect(marks.sort()).toEqual(["bold", "italic", "strike", "underline"]);
  });

  it("parses a simple table", () => {
    const tbl = `<w:tbl><w:tr><w:tc><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>b</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`;
    const pm = ooxmlToPm(bodyXml(tbl));
    const table = pm.content[0]!;
    expect(table.type).toBe("table");
    const row = (table as { content: unknown[] }).content[0] as { type: string; content: unknown[] };
    expect(row.type).toBe("tableRow");
    expect(row.content).toHaveLength(2);
  });

  it("wraps unknown elements as unmanaged blocks", () => {
    const pm = ooxmlToPm(bodyXml(`<w:p><w:r><w:t>ok</w:t></w:r></w:p><w:tblBorders/>`));
    const last = pm.content[pm.content.length - 1]!;
    expect(last.type).toBe("unmanagedBlock");
  });

  it("groups consecutive list paragraphs into bulletList", () => {
    const p = (id: string) =>
      `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="${id}"/></w:numPr></w:pPr><w:r><w:t>item</w:t></w:r></w:p>`;
    const pm = ooxmlToPm(bodyXml(p("1") + p("1")));
    expect(pm.content[0]!.type).toBe("bulletList");
    expect((pm.content[0] as { content: unknown[] }).content).toHaveLength(2);
  });
});

describe("character size units", () => {
  const runWithSz = (half: number) =>
    ooxmlToPm(
      bodyXml(
        `<w:p><w:r><w:rPr><w:sz w:val="${half}"/><w:szCs w:val="${half}"/></w:rPr><w:t>x</w:t></w:r></w:p>`,
      ),
    );
  const sizeOf = (pm: ReturnType<typeof ooxmlToPm>) =>
    ((pm.content[0] as { content: Array<{ marks?: Array<{ attrs?: Record<string, unknown> }> }> })
      .content[0]?.marks?.[0]?.attrs?.fontSize) as string | undefined;
  const withSize = (fontSize: string | undefined) =>
    pmToOoxml({
      type: "doc",
      content: [
        {
          type: "paragraph",
          content: [
            { type: "text", text: "x", marks: [{ type: "textStyle", attrs: { fontSize } }] },
          ],
        },
      ],
    });

  it("reads Word points as CSS px so the browser honours them", () => {
    expect(sizeOf(runWithSz(24))).toBe("16px"); // 小四 = 12pt
    expect(sizeOf(runWithSz(21))).toBe("14px"); // 五号 = 10.5pt
    expect(sizeOf(runWithSz(32))).toBe("21.3px"); // 16pt
  });

  it("writes px back as points, not as if px were points", () => {
    expect(withSize("14px")).toContain('<w:sz w:val="21"/>');
  });

  it("keeps working for bare numbers, which older versions wrote as points", () => {
    expect(withSize("12")).toContain('<w:sz w:val="24"/>');
  });

  it("drops an unparsable size instead of emitting w:sz NaN", () => {
    expect(withSize("小四")).not.toContain("<w:sz");
  });

  it("survives a round trip", () => {
    expect(withSize(sizeOf(runWithSz(24)))).toContain('<w:sz w:val="24"/>');
  });
});

describe("colour values", () => {
  const withMark = (attrs: Record<string, string | undefined>) =>
    pmToOoxml({
      type: "doc",
      content: [
        {
          type: "paragraph",
          content: [{ type: "text", text: "x", marks: [{ type: "textStyle", attrs }] }],
        },
      ],
    });

  const markOf = (pm: ReturnType<typeof ooxmlToPm>) =>
    (
      pm.content[0] as {
        content: Array<{ marks?: Array<{ attrs?: Record<string, unknown> }> }>;
      }
    ).content[0]?.marks?.[0];

  it("reads hex and rgb() as OOXML hex", () => {
    expect(cssColorToOoxmlHex("#ff0000")).toBe("FF0000");
    expect(cssColorToOoxmlHex("#f00")).toBe("FF0000");
    expect(cssColorToOoxmlHex("rgb(16, 124, 208)")).toBe("107CD0");
    expect(cssColorToOoxmlHex("rgba(0, 255, 0, 1)")).toBe("00FF00");
  });

  it("refuses anything Word cannot read", () => {
    expect(cssColorToOoxmlHex("red")).toBeNull();
    expect(cssColorToOoxmlHex("#gg0000")).toBeNull();
    expect(cssColorToOoxmlHex(undefined)).toBeNull();
  });

  it("normalises a pasted rgb() colour before writing it", () => {
    expect(withMark({ color: "rgb(255, 0, 0)" })).toContain('<w:color w:val="FF0000"/>');
    expect(withMark({ backgroundColor: "rgb(255, 255, 0)" })).toContain(
      '<w:highlight w:val="yellow"/>',
    );
  });

  it("drops a garbage colour instead of corrupting the file", () => {
    const xml = withMark({ color: "not-a-colour", backgroundColor: "also-bad" });
    expect(xml).not.toContain("<w:color");
    expect(xml).not.toContain("<w:shd");
  });

  it("round-trips a Word colour", () => {
    const pm = ooxmlToPm(
      bodyXml(`<w:p><w:r><w:rPr><w:color w:val="A4262C"/></w:rPr><w:t>x</w:t></w:r></w:p>`),
    );
    expect(markOf(pm)?.attrs?.color).toBe("#a4262c");
    expect(pmToOoxml(pm)).toContain('<w:color w:val="A4262C"/>');
  });

  it("reads highlight names and shading as canonical hex", () => {
    const pm = ooxmlToPm(
      bodyXml(
        `<w:p>` +
          `<w:r><w:rPr><w:highlight w:val="darkYellow"/></w:rPr><w:t>x</w:t></w:r>` +
          `<w:r><w:rPr><w:shd w:val="clear" w:color="auto" w:fill="C00000"/></w:rPr><w:t>y</w:t></w:r>` +
          `<w:r><w:rPr><w:highlight w:val="none"/></w:rPr><w:t>z</w:t></w:r>` +
          `</w:p>`,
      ),
    );
    const marks = (
      pm.content[0] as {
        content: Array<{ marks?: Array<{ attrs?: Record<string, unknown> }> }>;
      }
    ).content;
    expect(marks[0]?.marks?.[0]?.attrs?.backgroundColor).toBe("#808000");
    expect(marks[1]?.marks?.[0]?.attrs?.backgroundColor).toBe("#c00000");
    expect(marks[2]?.marks ?? []).toHaveLength(0);
  });
});

describe("字符样式维度", () => {
  const runs = (...inner: string[]) => ooxmlToPm(bodyXml(`<w:p>${inner.join("")}</w:p>`));
  type Run = { marks?: Array<{ type: string; attrs?: Record<string, unknown> }> };
  const runsOf = (pm: ReturnType<typeof runs>) =>
    (pm.content[0] as { content?: Run[] }).content ?? [];
  const marksOf = (pm: ReturnType<typeof runs>, i: number) =>
    (runsOf(pm)[i]?.marks ?? []).map((m) => m.type);
  const attrsOf = (pm: ReturnType<typeof runs>, i: number) =>
    runsOf(pm)[i]?.marks?.find((m) => m.attrs)?.attrs;

  it("reads an explicit off switch as off", () => {
    // `<w:b w:val="0"/>` is how Word writes "not bold, inherited reset" — the
    // element being present says nothing on its own.
    const pm = runs(
      `<w:r><w:rPr><w:b w:val="0"/><w:i w:val="false"/><w:u w:val="none"/><w:strike w:val="off"/></w:rPr><w:t>a</w:t></w:r>`,
      `<w:r><w:rPr><w:b/><w:i/><w:u w:val="single"/><w:strike/></w:rPr><w:t>b</w:t></w:r>`,
    );
    expect(marksOf(pm, 0)).toEqual([]);
    expect(marksOf(pm, 1).sort()).toEqual(["bold", "italic", "strike", "underline"].sort());
  });

  it("writes italic and underline back as marks", () => {
    const xml = pmToOoxml({
      type: "doc",
      content: [
        {
          type: "paragraph",
          content: [
            { type: "text", text: "a", marks: [{ type: "italic" }, { type: "underline" }] },
          ],
        },
      ],
    });
    expect(xml).toContain("<w:i/>");
    expect(xml).toContain('<w:u w:val="single"/>');
  });

  it("round-trips 字间距 through twentieths of a point", () => {
    const pm = runs(`<w:r><w:rPr><w:spacing w:val="10"/></w:rPr><w:t>a</w:t></w:r>`);
    expect(attrsOf(pm, 0)?.letterSpacing).toBe(ptToCssLetterSpacing(0.5));
    // The value a template stores is the one that produces 10 twips again, so
    // save → reload → save leaves the file byte-identical.
    const xml = pmToOoxml(pm);
    expect(xml).toContain('<w:spacing w:val="10"/>');
    expect(attrsOf(ooxmlToPm(xml), 0)?.letterSpacing).toBe(ptToCssLetterSpacing(0.5));
  });

  it("keeps 中文字体 and 西文字体 in their own slots", () => {
    const pm = runs(
      `<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="宋体"/></w:rPr><w:t>汉字</w:t></w:r>`,
      `<w:r><w:rPr><w:rFonts w:ascii="仿宋" w:hAnsi="仿宋" w:eastAsia="仿宋" w:cs="仿宋"/></w:rPr><w:t>x</w:t></w:r>`,
    );
    expect(attrsOf(pm, 0)).toMatchObject({
      fontFamily: "宋体",
      fontFamilyWest: "Times New Roman",
    });
    expect(attrsOf(pm, 1)?.fontFamilyWest).toBeUndefined();

    const xml = pmToOoxml(pm);
    expect(xml).toContain(
      '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman" w:eastAsia="宋体"/>',
    );
    // One writer per element: a second `<w:rFonts>` here would be a real bug.
    expect(xml.match(/<w:rFonts /g)).toHaveLength(2);
  });

  it("serialises run properties in a fixed order", () => {
    const withBoth = pmToOoxml({
      type: "doc",
      content: [
        {
          type: "paragraph",
          content: [
            {
              type: "text",
              text: "a",
              marks: [
                { type: "underline" },
                { type: "textStyle", attrs: { color: "#c00000", fontSize: "21px" } },
              ],
            },
          ],
        },
      ],
    });
    const rPr = /<w:rPr>(.*)<\/w:rPr>/.exec(withBoth)?.[1] ?? "";
    expect(rPr.indexOf("<w:color")).toBeLessThan(rPr.indexOf("<w:sz"));
    expect(rPr.indexOf("<w:sz")).toBeLessThan(rPr.indexOf("<w:u"));
  });

  /**
   * 上标 / 下标是 `<w:vertAlign w:val>` 这一个元素的三档（ECMA-376 §17.3.2.27 的
   * ST_VerticalAlignRun），所以读出来也只能是一个标记带一个值。
   */
  it("上标与下标读成同一个标记的两个值，baseline 读成没有标记", () => {
    const pm = runs(
      `<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t>1</w:t></w:r>`,
      `<w:r><w:rPr><w:vertAlign w:val="subscript"/></w:rPr><w:t>2</w:t></w:r>`,
      // 缺省位置也是"写过"：Word 用它显式取消样式继承来的上标。
      `<w:r><w:rPr><w:vertAlign w:val="baseline"/></w:rPr><w:t>3</w:t></w:r>`,
      // 认不出来的档位整个不写，不能把一个 Word 不认的值带到下一份文件里。
      `<w:r><w:rPr><w:vertAlign w:val="raised"/></w:rPr><w:t>4</w:t></w:r>`,
    );
    expect(marksOf(pm, 0)).toEqual(["vertAlign"]);
    expect(attrsOf(pm, 0)).toEqual({ val: "superscript" });
    expect(attrsOf(pm, 1)).toEqual({ val: "subscript" });
    expect(marksOf(pm, 2)).toEqual([]);
    expect(marksOf(pm, 3)).toEqual([]);
  });

  it("上标读进来存回去是同一份 XML，两档不会互相顶掉", () => {
    const xml = pmToOoxml(
      runs(
        `<w:r><w:rPr><w:vertAlign w:val="subscript"/></w:rPr><w:t>H</w:t></w:r>`,
        `<w:r><w:rPr><w:b/><w:vertAlign w:val="superscript"/></w:rPr><w:t>2</w:t></w:r>`,
      ),
    );
    expect(xml).toContain('<w:vertAlign w:val="subscript"/>');
    expect(xml).toContain("<w:b/><w:vertAlign w:val=\"superscript\"/>");
    // 一个 run 只有一个 `<w:vertAlign>`：写两条 Word 只认第一条，第二条是静默丢失。
    expect(xml.match(/<w:vertAlign /g)).toHaveLength(2);
    expect(pmToOoxml(ooxmlToPm(xml))).toBe(xml);
  });

  it("`w:vertAlign` 排在 rPr 末尾，与 CT_RPr 的元素顺序一致", () => {
    const xml = pmToOoxml({
      type: "doc",
      content: [
        {
          type: "paragraph",
          content: [
            {
              type: "text",
              text: "a",
              marks: [{ type: "vertAlign", attrs: { val: "superscript" } }, { type: "underline" }],
            },
          ],
        },
      ],
    });
    const rPr = /<w:rPr>(.*)<\/w:rPr>/.exec(xml)?.[1] ?? "";
    expect(rPr.indexOf("<w:u")).toBeLessThan(rPr.indexOf("<w:vertAlign"));
  });
});

describe("canonColor", () => {
  const cases: [string, string | null][] = [
    ["#C00000", "#c00000"],
    ["#c00000", "#c00000"],
    ["#ccc", "#cccccc"],
    ["rgb(192, 0, 0)", "#c00000"],
    ["rgba(192, 0, 0, 0.5)", "#c00000"],
    ["  #107CD0  ", "#107cd0"],
    ["black", "#000000"],
    ["white", "#ffffff"],
    ["grey", "#808080"],
    ["auto", null],
    ["transparent", null],
    ["none", null],
    ["", null],
    ["not-a-colour", null],
    ["#gg0000", null],
  ];
  for (const [input, expected] of cases) {
    it(`canonicalises ${JSON.stringify(input)}`, () => {
      expect(canonColor(input)).toBe(expected);
    });
  }

  it("is stable — a canonical colour re-reads as itself", () => {
    for (const [input, expected] of cases) {
      if (expected === null) continue;
      expect(canonColor(canonColor(expected))).toBe(expected);
    }
  });

  it("makes the DOM round trip compare equal", () => {
    // The template writes #c00000, the browser reads it back as rgb(...).
    expect(canonColor("rgb(192, 0, 0)")).toBe(canonColor("#C00000"));
  });
});

describe("pmToOoxml round-trip", () => {
  it("re-serializes a paragraph with bold text", () => {
    const xml = pmToOoxml({
      type: "doc",
      content: [
        {
          type: "paragraph",
          content: [
            { type: "text", text: "hi " },
            { type: "text", text: "bold", marks: [{ type: "bold" }] },
          ],
        },
      ],
    });
    expect(xml).toContain("<w:b/>");
    expect(xml).toContain("bold");
  });

  it("round-trips a table through PMDoc", async () => {
    const pm = ooxmlToPm(
      bodyXml(
        `<w:tbl><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`,
      ),
    );
    const xml = pmToOoxml(pm);
    expect(xml).toContain("<w:tbl>");
    expect(xml).toContain(">cell<");
  });
});

describe("分隔线导出为段落底部边框", () => {
  const hrDoc: PMDoc = {
    type: "doc",
    content: [{ type: "horizontalRule" }],
  };

  it("写 w:pBdr/w:bottom，不再引用本仓没定义的 HorizontalRule 样式", () => {
    const xml = pmToOoxml(hrDoc);
    expect(xml).toContain(
      '<w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/>',
    );
    expect(xml).not.toContain("HorizontalRule");
  });

  it("带底部边框的空段落原样读回，重存不丢边框", () => {
    const bordered = `<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr></w:pPr></w:p>`;
    expect(pmToOoxml(ooxmlToPm(bodyXml(bordered)))).toContain(bordered);
  });
});

type MarkedNode = { marks?: Array<{ type: string }> };

describe("文档级超链接样式", () => {
  const linkDoc: PMDoc = {
    type: "doc",
    content: [
      {
        type: "paragraph",
        content: [
          {
            type: "text",
            text: "官网",
            marks: [{ type: "link", attrs: { href: "#top" } }],
          },
        ],
      },
    ],
  };

  const withOwnColor: PMDoc = {
    type: "doc",
    content: [
      {
        type: "paragraph",
        content: [
          {
            type: "text",
            text: "官网",
            marks: [
              { type: "link", attrs: { href: "#top" } },
              { type: "textStyle", attrs: { color: "#c00000" } },
            ],
          },
        ],
      },
    ],
  };

  const inlinesOf = (doc: PMDoc): MarkedNode[] =>
    ((doc.content[0] as { content?: MarkedNode[] })?.content ?? []) as MarkedNode[];

  it("keeps a hyperlink's own decoration out of the run style", () => {
    const pm = ooxmlToPm(
      bodyXml(
        `<w:p><w:r><w:t>正文</w:t></w:r>` +
          `<w:hyperlink r:id="rId5"><w:r>` +
          `<w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr><w:t>链接</w:t>` +
          `</w:r></w:hyperlink></w:p>`,
      ),
    );
    const inlines = inlinesOf(pm);
    expect(inlines[1]?.marks?.some((m) => m.type === "link")).toBe(true);
    // The blue and the rule belong to the document, not to this run: read as
    // the run's own colour they disagree with the sentence around them, and
    // every paragraph holding a link becomes unstyleable.
    expect(inlines[1]?.marks?.some((m) => m.type === "underline")).toBe(false);
    expect(inlines[1]?.marks?.some((m) => m.type === "textStyle")).toBe(false);
    expect(inlines[0]?.marks ?? []).toHaveLength(0);
  });

  it("writes the document link style into every link run", () => {
    const xml = pmToOoxml(linkDoc, {
      link: { color: "#000000", underline: false },
    });
    expect(xml).toContain('<w:color w:val="000000"/>');
    expect(xml).not.toContain("<w:u");

    const underlined = pmToOoxml(linkDoc, {
      link: { color: "rgb(0, 105, 92)", underline: true },
    });
    expect(underlined).toContain('<w:color w:val="00695C"/>');
    expect(underlined).toContain('<w:u w:val="single"/>');
  });

  it("falls back to Word's blue when no template spoke", () => {
    const xml = pmToOoxml(linkDoc);
    expect(xml).toContain('<w:color w:val="0563C1"/>');
    expect(xml).toContain('<w:u w:val="single"/>');
  });

  it("prefers the user's colour and never writes two of them", () => {
    const xml = pmToOoxml(withOwnColor, { link: { color: "#000000" } });
    expect(xml).toContain('<w:color w:val="C00000"/>');
    expect(xml.match(/<w:color /g)).toHaveLength(1);
  });

  it("survives save/reload without gaining a style of its own", () => {
    const reloaded = ooxmlToPm(
      pmToOoxml(linkDoc, { link: { color: "#000000" } }),
    );
    expect(inlinesOf(reloaded)[0]?.marks?.some((m) => m.type === "textStyle")).toBe(
      false,
    );
  });
});

describe("lineHeight 固定值 / 最小值", () => {
  const withLine = (lineHeight: string) =>
    pmToOoxml({
      type: "doc",
      content: [
        {
          type: "paragraph",
          attrs: { lineHeight },
          content: [{ type: "text", text: "x" }],
        },
      ],
    });
  const readBack = (xml: string) =>
    ((ooxmlToPm(xml).content[0] as { attrs?: Record<string, unknown> }).attrs
      ?.lineHeight as string | undefined);

  it("writes a bare multiplier as 240ths with lineRule=auto", () => {
    const xml = withLine("1.5");
    expect(xml).toContain('w:line="360"');
    expect(xml).toContain('w:lineRule="auto"');
  });

  it("writes pt:29 as twentieths of a point with lineRule=exact", () => {
    const xml = withLine("pt:29");
    expect(xml).toContain('w:line="580"');
    expect(xml).toContain('w:lineRule="exact"');
  });

  it("reads exact spacing back as pt:<point>", () => {
    expect(
      readBack(
        bodyXml(
          `<w:p><w:pPr><w:spacing w:line="580" w:lineRule="exact"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>`,
        ),
      ),
    ).toBe("pt:29");
  });

  it("reads 最小值 as atLeast:<point>, distinct from 固定值", () => {
    expect(
      readBack(
        bodyXml(
          `<w:p><w:pPr><w:spacing w:line="180" w:lineRule="atLeast"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>`,
        ),
      ),
    ).toBe("atLeast:9");
  });

  it("writes atLeast back with lineRule=atLeast, not exact", () => {
    const xml = withLine("atLeast:9");
    expect(xml).toContain('w:line="180"');
    expect(xml).toContain('w:lineRule="atLeast"');
    expect(readBack(xml)).toBe("atLeast:9");
  });

  it("replaces a paired <w:spacing> kept in rawPpr", () => {
    // An imported document can carry <w:spacing …></w:spacing> rather than the
    // self-closing form; leaving it in place would emit two spacing elements and
    // Word would obey the stale first one, so the 最小值 never reaches the page.
    const xml = pmToOoxml({
      type: "doc",
      content: [
        {
          type: "paragraph",
          attrs: {
            rawPpr: '<w:spacing w:before="120" w:line="240"></w:spacing>',
            lineHeight: "atLeast:20",
          },
          content: [{ type: "text", text: "x" }],
        },
      ],
    });
    expect(xml.match(/<w:spacing/g)).toHaveLength(1);
    expect(xml).toContain('w:line="400"');
    expect(xml).toContain('w:lineRule="atLeast"');
  });

  it("survives save/reload so a re-applied template stays a no-op", () => {
    expect(readBack(withLine("pt:29"))).toBe("pt:29");
    expect(readBack(withLine("1.5"))).toBe("1.5");
  });

  it("keeps the point size stable through the CSS px font size", () => {
    // 三号 = 16pt; templates write the px form, so it must convert exactly.
    expect(ptToCssFontSize(16)).toBe("21.3px");
    expect(cssFontSizeToPt("21.3px")).toBeCloseTo(16, 1);
  });
});

/**
 * 段落缩进/间距的**属性级合并**。以前的写法是"整元素重建"：cleanRaw 把源文档的
 * <w:ind>/<w:spacing> 删掉，再只按我们建模过的属性重新生成，于是右缩进、
 * w:*Chars（中文版 Word 的主写法）、w:*Lines、2010 版的 w:start/w:end
 * 保存一次就永久消失。下面每条都是那批丢失场景的回归用例。
 */
describe("段落缩进与间距的属性级合并", () => {
  const PARA = (pPr: string) =>
    bodyXml(`<w:p><w:pPr>${pPr}</w:pPr><w:r><w:t>x</w:t></w:r></w:p>`);
  /** OOXML → PM → OOXML，返回 attrs 与写出的 pPr 片段。 */
  function roundTripInd(pPr: string): { attrs: Record<string, unknown>; ind: string } {
    const pm = ooxmlToPm(PARA(pPr));
    const out = pmToOoxml(pm);
    return {
      attrs: (pm.content?.[0]?.attrs ?? {}) as Record<string, unknown>,
      ind: /<w:ind\b[^>]*>/.exec(out)?.[0] ?? "",
    };
  }
  function roundTripSpacing(pPr: string): { attrs: Record<string, unknown>; spacing: string } {
    const pm = ooxmlToPm(PARA(pPr));
    const out = pmToOoxml(pm);
    return {
      attrs: (pm.content?.[0]?.attrs ?? {}) as Record<string, unknown>,
      spacing: /<w:spacing\b[^>]*>/.exec(out)?.[0] ?? "",
    };
  }

  it("keeps a right indent that the editor never touched", () => {
    // 右缩进以前没有对应属性，重建 <w:ind> 时直接被抹平——用户右侧缩进全丢。
    const { attrs, ind } = roundTripInd('<w:ind w:left="300" w:right="200"/>');
    expect(attrs.indent).toBe(15);
    expect(attrs.indentRight).toBe(10);
    expect(ind).toContain('w:left="300"');
    expect(ind).toContain('w:right="200"');
  });

  it("reads the Office 2010 w:start / w:end spellings", () => {
    const { attrs, ind } = roundTripInd('<w:ind w:start="400" w:end="200"/>');
    expect(attrs.indent).toBe(20);
    expect(attrs.indentRight).toBe(10);
    // 归一化成 left/right，但不能把值换成 0。
    expect(ind).toContain('w:left="400"');
    expect(ind).toContain('w:right="200"');
    expect(ind).not.toContain("w:start");
    expect(ind).not.toContain("w:end");
  });

  it("carries indentation formats we do not model back out", () => {
    // 中文版 Word 用字符单位写缩进（w:leftChars 会压过 w:left），必须原样带回。
    const { ind } = roundTripInd('<w:ind w:left="300" w:leftChars="200"/>');
    expect(ind).toContain('w:leftChars="200"');
    expect(ind.match(/<w:ind\b/g)).toHaveLength(1);
  });

  it("drops the synonym that would override an indentation the user changed", () => {
    // 改了左缩进却留着 w:leftChars，Word 会照旧按 2 字符排版，改动看不见。
    const xml = pmToOoxml({
      type: "doc",
      content: [
        {
          type: "paragraph",
          attrs: { rawPpr: '<w:ind w:left="300" w:leftChars="200"/>', indent: 30 },
          content: [{ type: "text", text: "x" }],
        },
      ],
    });
    const ind = /<w:ind\b[^>]*>/.exec(xml)?.[0] ?? "";
    expect(ind).toContain('w:left="600"');
    expect(ind).not.toContain("w:leftChars");
  });

  it("round-trips a hanging indent and lets 0 clear it", () => {
    const hanging = roundTripInd('<w:ind w:left="300" w:hanging="420"/>');
    expect(hanging.attrs.indentFirst).toBe(-21);
    expect(hanging.ind).toContain('w:hanging="420"');
    // 段落面板把"特殊格式"改成"首行缩进 0"时，悬挂必须被清掉而不是留在文件里。
    const cleared = pmToOoxml({
      type: "doc",
      content: [
        {
          type: "paragraph",
          attrs: { rawPpr: '<w:ind w:left="300" w:hanging="420"/>', indentFirst: 0 },
          content: [{ type: "text", text: "x" }],
        },
      ],
    });
    const ind = /<w:ind\b[^>]*>/.exec(cleared)?.[0] ?? "";
    expect(ind).toContain('w:firstLine="0"');
    expect(ind).not.toContain("w:hanging");
  });

  it("replaces a per-line spacing setting when 段前 is set in points", () => {
    // w:beforeLines（按行计的段前）优先级高于 w:before，两者并存会让改动失效。
    const { spacing } = roundTripSpacing('<w:spacing w:beforeLines="100"/>');
    expect(spacing).not.toBe("");
    const xml = pmToOoxml({
      type: "doc",
      content: [
        {
          type: "paragraph",
          attrs: { rawPpr: '<w:spacing w:beforeLines="100" w:after="60"/>', spacingBefore: 12 },
          content: [{ type: "text", text: "x" }],
        },
      ],
    });
    const spacing2 = /<w:spacing\b[^>]*>/.exec(xml)?.[0] ?? "";
    expect(spacing2).toContain('w:before="240"');
    expect(spacing2).toContain('w:after="60"');
    expect(spacing2).not.toContain("w:beforeLines");
  });

  it("emits spacing and ind before jc, the order CT_PPr declares", () => {
    const xml = pmToOoxml({
      type: "doc",
      content: [
        {
          type: "paragraph",
          attrs: { rawPpr: '<w:pBdr><w:bottom w:val="single"/></w:pBdr>', spacingAfter: 6, indentRight: 10, textAlign: "center" },
          content: [{ type: "text", text: "x" }],
        },
      ],
    });
    const pPr = /<w:pPr>[\s\S]*?<\/w:pPr>/.exec(xml)?.[0] ?? "";
    expect(pPr.indexOf("<w:pBdr")).toBeLessThan(pPr.indexOf("<w:spacing"));
    expect(pPr.indexOf("<w:spacing")).toBeLessThan(pPr.indexOf("<w:ind"));
    expect(pPr.indexOf("<w:ind")).toBeLessThan(pPr.indexOf("<w:jc"));
  });

  it("makes both pPr builders agree on the same attrs", () => {
    // 新建段落与带 rawPpr 的导入段落现在共用 paragraphPprXml 一条路径；
    // 以前两条路径各写一份换算，负的 indentFirst 在干净路径里会被写成
    // w:firstLine="-420"。
    const attrs = { spacingBefore: 6, spacingAfter: 6, indent: 21, indentFirst: -21, lineHeight: "1.5" };
    const clean = pmToOoxml({
      type: "doc",
      content: [{ type: "paragraph", attrs, content: [{ type: "text", text: "x" }] }],
    });
    const imported = pmToOoxml({
      type: "doc",
      content: [
        {
          type: "paragraph",
          attrs: { ...attrs, rawPpr: "<w:snapToGrid/>" },
          content: [{ type: "text", text: "x" }],
        },
      ],
    });
    const grab = (xml: string, tag: string) =>
      [...xml.matchAll(new RegExp(`<w:${tag}\\b[^>]*>`, "g"))].map((m) => m[0]).join("");
    expect(grab(imported, "spacing")).toBe(grab(clean, "spacing"));
    expect(grab(imported, "ind")).toBe(grab(clean, "ind"));
    expect(grab(clean, "ind")).toContain('w:hanging="420"');
  });
});

describe("段落开关与大纲级别（Word 段落对话框的其余字段）", () => {
  /** PM attrs → 写出的 `<w:pPr>` 片段，用于断言元素本身与它们的先后。 */
  function pPrOf(attrs: PPrAttrs): string {
    const xml = pmToOoxml({
      type: "doc",
      content: [{ type: "paragraph", attrs, content: [{ type: "text", text: "x" }] }],
    });
    return /<w:pPr>[\s\S]*?<\/w:pPr>/.exec(xml)?.[0] ?? "";
  }
  /** OOXML → PM → OOXML，返回段落 attrs 与写出的 `<w:pPr>` 片段。 */
  function roundTripPpr(pPr: string): { attrs: PPrAttrs; pPr: string } {
    const pm = ooxmlToPm(bodyXml(`<w:p><w:pPr>${pPr}</w:pPr><w:r><w:t>x</w:t></w:r></w:p>`));
    const out = pmToOoxml(pm);
    return {
      attrs: (pm.content?.[0]?.attrs ?? {}) as PPrAttrs,
      pPr: /<w:pPr>[\s\S]*?<\/w:pPr>/.exec(out)?.[0] ?? "",
    };
  }

  it("reads 与下段同页 / 段中不分页 / 孤行控制 / 对齐网格 / 大纲级别", () => {
    const { attrs } = roundTripPpr(
      '<w:keepNext/><w:keepLines/><w:widowControl w:val="0"/><w:snapToGrid w:val="0"/><w:outlineLvl w:val="2"/>',
    );
    expect(attrs).toMatchObject({
      keepWithNext: true,
      keepLines: true,
      widowControl: false,
      snapToGrid: false,
      outlineLevel: 2,
    });
  });

  it("writes the switches back, 关闭的一侧带 w:val=0", () => {
    const pPr = pPrOf({
      keepWithNext: true,
      keepLines: true,
      widowControl: false,
      snapToGrid: false,
      pageBreakBefore: true,
      outlineLevel: 0,
    });
    expect(pPr).toContain("<w:keepNext/>");
    expect(pPr).toContain("<w:keepLines/>");
    expect(pPr).toContain("<w:pageBreakBefore/>");
    expect(pPr).toContain('<w:widowControl w:val="0"/>');
    expect(pPr).toContain('<w:snapToGrid w:val="0"/>');
    expect(pPr).toContain('<w:outlineLvl w:val="0"/>');
  });

  it("keeps the CT_PPr sequence: 开关在前，缩进居中，大纲级别在最后", () => {
    const pPr = pPrOf({
      keepWithNext: true,
      widowControl: false,
      spacingAfter: 6,
      indent: 21,
      textAlign: "center",
      outlineLevel: 1,
    });
    const rank = (tag: string) => pPr.indexOf(`<${tag}`);
    for (const tag of ["w:keepNext", "w:widowControl", "w:spacing", "w:ind", "w:jc", "w:outlineLvl"]) {
      expect(rank(tag)).toBeGreaterThan(-1);
    }
    expect(rank("w:keepNext")).toBeLessThan(rank("w:widowControl"));
    expect(rank("w:widowControl")).toBeLessThan(rank("w:spacing"));
    expect(rank("w:spacing")).toBeLessThan(rank("w:ind"));
    expect(rank("w:ind")).toBeLessThan(rank("w:jc"));
    expect(rank("w:jc")).toBeLessThan(rank("w:outlineLvl"));
  });

  it("keeps the paragraph mark's <w:rPr> behind the properties we set", () => {
    // Word 把段落标记的字符属性写在 pPr 末尾。以前整串 rawPpr 排在最前，
    // 我们补的 spacing/jc 就落到 rPr 之后——顺序不合 schema，Word 会提示文档有问题。
    const { pPr } = roundTripPpr(
      '<w:spacing w:after="120"/><w:rPr><w:sz w:val="24"/></w:rPr>',
    );
    expect(pPr).toContain('<w:spacing w:after="120"');
    expect(pPr.indexOf("<w:rPr")).toBeGreaterThan(pPr.indexOf("<w:spacing"));
  });

  it("is idempotent — 没动过的段落原样写回，也不会补上它没写过的开关", () => {
    const source =
      '<w:keepNext/><w:widowControl w:val="0"/><w:spacing w:after="120"/><w:jc w:val="center"/><w:outlineLvl w:val="9"/>';
    expect(roundTripPpr(source).pPr).toBe(`<w:pPr>${source}</w:pPr>`);
    // 一段纯文字不该长出 pPr。
    expect(pPrOf({})).toBe("");
  });

  it("reads 制表位：位置换算成 pt，对齐与前导符原样带走", () => {
    // Word 的目录条目就是这么写的：右对齐 + 点前导符，位置在右页边（twips）。
    const tabs = '<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="8306"/></w:tabs>';
    const { attrs, pPr } = roundTripPpr(tabs);
    expect(attrs.tabStops).toEqual([{ positionPt: 415.3, alignment: "right", leader: "dot" }]);
    expect(pPr).toBe(`<w:pPr>${tabs}</w:pPr>`);
  });

  it("keeps 缺省写法与 clear：没 w:val 是 left，没 w:leader 就不写", () => {
    const tabs =
      '<w:tabs><w:tab w:val="center" w:pos="4000"/><w:tab w:pos="200"/><w:tab w:val="clear" w:pos="9016"/></w:tabs>';
    const { attrs, pPr } = roundTripPpr(tabs);
    expect(attrs.tabStops).toEqual([
      { positionPt: 200, alignment: "center" },
      { positionPt: 10, alignment: "left" },
      { positionPt: 450.8, alignment: "clear" },
    ]);
    // 没写过的 w:leader 不补；缺省的 w:val 由 Word 省略，我们写出显式的 left。
    expect(pPr).toContain('<w:tab w:val="center" w:pos="4000"/>');
    expect(pPr).toContain('<w:tab w:val="clear" w:pos="9016"/>');
    expect(pPr).not.toContain("w:leader");
  });

  it("编辑器里清空制表位时，源文档那条 w:tabs 不再透传回来", () => {
    // Word 的"全部清除"就是删掉这一段自己的制表位定义（样式给的仍然生效），
    // 所以写出的是不带 w:tabs 的 pPr，而不是补一条空的 tabs。
    const pm = ooxmlToPm(
      bodyXml(
        '<w:p><w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="8306"/></w:tabs>' +
          '<w:spacing w:after="120"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>',
      ),
    );
    const attrs = { ...(pm.content?.[0]?.attrs as PPrAttrs), tabStops: [] };
    const out = pmToOoxml({
      type: "doc",
      content: [{ type: "paragraph", attrs, content: [{ type: "text", text: "x" }] }],
    });
    const pPr = /<w:pPr>[\s\S]*?<\/w:pPr>/.exec(out)?.[0] ?? "";
    expect(pPr).not.toContain("w:tabs");
    expect(pPr).toContain('<w:spacing w:after="120"');
  });

  it("制表位读不懂时整组当作没建模过，原文原样带走", () => {
    // 只留认得出的那几条 = 静默删掉用户的制表位；w:val 是表外值或没有 w:pos 都算读不懂。
    for (const tabs of [
      '<w:tabs><w:tab w:val="bogus" w:pos="100"/><w:tab w:val="right" w:pos="200"/></w:tabs>',
      '<w:tabs><w:tab w:val="right"/></w:tabs>',
      '<w:tabs><w:tab w:val="right" w:leader="stars" w:pos="100"/></w:tabs>',
    ]) {
      const { attrs, pPr } = roundTripPpr(tabs);
      expect(attrs.tabStops).toBeUndefined();
      expect(pPr).toBe(`<w:pPr>${tabs}</w:pPr>`);
    }
  });

  it("厘米录入与磅来回换算，位置落在 Word 的显示精度上", () => {
    // 界面按厘米录入（中文版 Word 的制表位对话框就是厘米），换算只能出自这里。
    expect(cmToPt(2)).toBe(56.7);
    expect(cmToPt(14.65)).toBe(415.3);
    expect(ptToCm(415.3)).toBe(14.65);
    expect(ptToCm(cmToPt(0.5))).toBe(0.5);
    // 读侧给的 0.1pt 位置显示成厘米、再录回去，最多差一个录入格（0.01cm ≈ 0.28pt），
    // 这就是 Word 自己按厘米定位置时会有的误差。
    const positionPt = 415.3;
    expect(Math.abs(cmToPt(ptToCm(positionPt)) - positionPt)).toBeLessThan(0.5);
  });
});

describe("首字下沉（w:pPr/w:framePr 的 dropCap 与 lines）", () => {
  /**
   * 把源文档的一段 pPr 读成 PM，再用给定 attrs 覆盖回去写出的 `<w:pPr>`。
   *
   * @param source 源文档 pPr 的内部 XML；传 null 表示这一段本来什么都没有（编辑器新建的段落）。
   * @param dropCap 用户在面板里选的下沉设置（undefined=没动过，null=选了"无"）。
   * @returns 写出的完整 `<w:pPr>` 串；段落没有 pPr 时是空串。
   */
  function rewriteFramePr(source: string | null, dropCap: PPrAttrs["dropCap"]): string {
    const loaded: PMDoc =
      source === null
        ? { type: "doc", content: [{ type: "paragraph", attrs: {}, content: [{ type: "text", text: "x" }] }] }
        : ooxmlToPm(
            bodyXml(`<w:p><w:pPr>${source}</w:pPr><w:r><w:t>x</w:t></w:r></w:p>`),
          );
    const attrs = { ...(loaded.content?.[0]?.attrs as PPrAttrs), dropCap };
    const xml = pmToOoxml({
      type: "doc",
      content: [{ type: "paragraph", attrs, content: [{ type: "text", text: "x" }] }],
    });
    return /<w:pPr>[\s\S]*?<\/w:pPr>/.exec(xml)?.[0] ?? "";
  }

  it("读出位置与行数，没改动的段落字节一致", () => {
    // Word 的"下沉 3 行"就是这个形状（§17.3.1.11：dropCap 是 ST_DropCap、lines 是整数）。
    const frame =
      '<w:framePr w:wrap="around" w:vAnchor="text" w:hAnchor="text" w:dropCap="drop" w:lines="3"/>';
    const pm = ooxmlToPm(bodyXml(`<w:p><w:pPr>${frame}</w:pPr><w:r><w:t>x</w:t></w:r></w:p>`));
    expect(pm.content?.[0]).toMatchObject({
      type: "paragraph",
      attrs: { dropCap: { position: "drop", lines: 3 } },
    });
    expect(rewriteFramePr(frame, undefined)).toBe(`<w:pPr>${frame}</w:pPr>`);
  });

  it("不是首字下沉的框架原样带走：真文本框、none 与表外值都算没建模过", () => {
    // 只留认得出的那几条 = 静默删掉用户的文本框；存在 `w:framePr` 本身就说"这段是框"。
    for (const frame of [
      '<w:framePr w:w="2191" w:h="811" w:hRule="exact" w:hSpace="180" w:wrap="around" w:vAnchor="text" w:hAnchor="page" w:x="1921"/>',
      '<w:framePr w:wrap="around" w:vAnchor="text" w:hAnchor="text" w:dropCap="none"/>',
      '<w:framePr w:dropCap="raised"/>',
    ]) {
      const pm = ooxmlToPm(bodyXml(`<w:p><w:pPr>${frame}</w:pPr><w:r><w:t>x</w:t></w:r></w:p>`));
      expect((pm.content?.[0]?.attrs as PPrAttrs).dropCap).toBeUndefined();
      expect(rewriteFramePr(frame, undefined)).toBe(`<w:pPr>${frame}</w:pPr>`);
    }
  });

  it("源文档没写 w:lines 时不猜行数，写回也就不会补上它", () => {
    const frame = '<w:framePr w:wrap="around" w:vAnchor="text" w:hAnchor="text" w:dropCap="margin"/>';
    const pm = ooxmlToPm(bodyXml(`<w:p><w:pPr>${frame}</w:pPr><w:r><w:t>x</w:t></w:r></w:p>`));
    // `margin` = 首字突出到页边距外（Word UI 的"文字环绕"档），行数留给 Word 自己按高度定。
    expect(pm.content?.[0]).toMatchObject({ attrs: { dropCap: { position: "margin" } } });
    expect(rewriteFramePr(frame, undefined)).toBe(`<w:pPr>${frame}</w:pPr>`);
  });

  it("新建的下沉补出锚定：缺省的 hAnchor/vAnchor 是页面，会把首字钉到纸边上去", () => {
    expect(rewriteFramePr(null, { position: "drop", lines: 2 })).toBe(
      '<w:pPr><w:framePr w:wrap="around" w:vAnchor="text" w:hAnchor="text" w:dropCap="drop" w:lines="2"/></w:pPr>',
    );
    // 行数没定就只写位置，不替用户编一个。
    expect(rewriteFramePr(null, { position: "margin" })).toBe(
      '<w:pPr><w:framePr w:wrap="around" w:vAnchor="text" w:hAnchor="text" w:dropCap="margin"/></w:pPr>',
    );
  });

  it("改行数只动 dropCap 与 lines，源文档那条框架的其余属性留着", () => {
    // 用户在下沉段落里只把 3 行改成 4 行，`w:hSpace` 这类没建模的属性不该被洗掉。
    const source =
      '<w:framePr w:wrap="around" w:vAnchor="text" w:hAnchor="text" w:hSpace="180" w:dropCap="drop" w:lines="3"/>';
    const expected =
      '<w:pPr><w:framePr w:wrap="around" w:vAnchor="text" w:hAnchor="text" w:dropCap="drop" w:lines="4" w:hSpace="180"/></w:pPr>';
    expect(rewriteFramePr(source, { position: "drop", lines: 4 })).toBe(expected);
    // 再存一次：顺序与值都不该再变（AGENTS §2.6 的幂等判据）。
    expect(
      rewriteFramePr(
        '<w:framePr w:wrap="around" w:vAnchor="text" w:hAnchor="text" w:dropCap="drop" w:lines="4" w:hSpace="180"/>',
        { position: "drop", lines: 4 },
      ),
    ).toBe(expected);
  });

  it("关掉下沉：纯下沉框架整条删掉，真文本框只把 dropCap 写成 none", () => {
    // `w:framePr` 的存在本身就意味着"这段是文本框"，留着空元素会把正文段变成匿名框。
    expect(
      rewriteFramePr(
        '<w:framePr w:wrap="around" w:vAnchor="text" w:hAnchor="text" w:dropCap="drop" w:lines="3"/>',
        null,
      ),
    ).toBe("");
    const realFrame =
      '<w:framePr w:w="2191" w:h="811" w:hRule="exact" w:wrap="around" w:vAnchor="text" w:hAnchor="page" w:x="1921" w:dropCap="drop" w:lines="3"/>';
    const off = rewriteFramePr(realFrame, null);
    expect(off).toContain('w:hRule="exact"');
    expect(off).toContain('w:dropCap="none"');
    expect(off).not.toContain("w:lines");
  });
});

describe("具名样式引用（w:pStyle 走 attrs.styleId）", () => {
  const WNS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const PARA = (pPr: string) =>
    bodyXml(`<w:p><w:pPr>${pPr}</w:pPr><w:r><w:t>x</w:t></w:r></w:p>`);
  /** 直接构造编辑器状态的段落，用于测"用户刚刚改了样式"的写侧。 */
  const paraDoc = (attrs: PPrAttrs): PMDoc => ({
    type: "doc",
    content: [
      { type: "paragraph", attrs, content: [{ type: "text", text: "x", marks: [] }] },
    ],
  });
  const styleRefs = (xml: string): string[] => xml.match(/<w:pStyle\b[^>]*>/g) ?? [];

  it("源文档的样式 id 进 attrs，写回只有一条 pStyle 且二次保存不变", () => {
    const pm = ooxmlToPm(PARA(`<w:pStyle w:val="af3"/><w:ind w:left="300"/>`));
    expect(pm.content?.[0]).toMatchObject({ type: "paragraph", attrs: { styleId: "af3" } });
    const out = pmToOoxml(pm);
    expect(styleRefs(out)).toEqual([`<w:pStyle w:val="af3"/>`]);
    expect(out).toContain('w:left="300"');
    expect(pmToOoxml(ooxmlToPm(out))).toBe(out);
  });

  it("编辑器里换了样式时，新的 id 顶掉源文档那条而不是并排两条", () => {
    // 同一份 pPr 里出现两个 w:pStyle 时 Word 取第一个，套哪个样式变得不可预测。
    const imported = ooxmlToPm(PARA(`<w:pStyle w:val="af3"/><w:ind w:left="300"/>`));
    const attrs = { ...(imported.content![0]!.attrs as PPrAttrs), styleId: "af9" };
    const out = pmToOoxml(paraDoc(attrs));
    expect(styleRefs(out)).toEqual([`<w:pStyle w:val="af9"/>`]);
    expect(out).not.toContain("af3");
    expect(out).toContain('w:left="300"');
  });

  it("styleId 为 null 是「清除样式」：引用删掉，其余段落属性一个都不动", () => {
    const out = pmToOoxml(
      paraDoc({ styleId: null, rawPpr: `<w:pStyle w:val="af3"/><w:ind w:left="300"/>` }),
    );
    expect(styleRefs(out)).toEqual([]);
    expect(out).toContain('w:left="300"');
  });

  it("没建模样式的段落不写 pStyle，编辑器新建的标题仍写推断的 Heading 级别", () => {
    expect(styleRefs(pmToOoxml(paraDoc({})))).toEqual([]);
    const fresh: PMDoc = {
      type: "doc",
      content: [
        {
          type: "heading",
          attrs: { level: 2 },
          content: [{ type: "text", text: "小标题", marks: [] }],
        },
      ],
    };
    expect(styleRefs(pmToOoxml(fresh))).toEqual([`<w:pStyle w:val="Heading2"/>`]);
  });

  it("WPS 那种数字 id 的标题不再被换成文档里不存在的 Heading1", () => {
    const styles = StyleMap.fromXml(
      `<w:styles xmlns:w="${WNS}"><w:style w:type="paragraph" w:styleId="2"><w:name w:val="heading 1"/></w:style></w:styles>`,
    );
    const pm = ooxmlToPm(PARA(`<w:pStyle w:val="2"/>`), undefined, styles);
    expect(pm.content?.[0]).toMatchObject({ type: "heading", attrs: { level: 1, styleId: "2" } });
    expect(styleRefs(pmToOoxml(pm))).toEqual([`<w:pStyle w:val="2"/>`]);
  });

  /** 声明了 1 级与 2 级标题样式的样式表，id 故意用 WPS 风格的数字。 */
  const TWO_HEADINGS = StyleMap.fromXml(
    `<w:styles xmlns:w="${WNS}">` +
      `<w:style w:type="paragraph" w:default="1" w:styleId="1"><w:name w:val="Normal"/></w:style>` +
      `<w:style w:type="paragraph" w:styleId="2"><w:name w:val="heading 1"/></w:style>` +
      `<w:style w:type="paragraph" w:styleId="3"><w:name w:val="heading 2"/></w:style>` +
      `</w:styles>`,
  );
  /** 编辑器里的标题段落：level 是用户刚改出来的，styleId 还是导入时那条。 */
  const headingDoc = (level: number, styleId?: string): PMDoc => ({
    type: "doc",
    content: [
      {
        type: "heading",
        attrs: { level, ...(styleId ? { styleId } : {}) },
        content: [{ type: "text", text: "x", marks: [] }],
      },
    ],
  });

  it("改级别后写出这一级的样式 id，而不是继续用旧的那一级", () => {
    // id "2" 在这份文档里就是"标题 1"；用户把它改成 2 级后仍写 "2"，Word 里段落
    // 会退回一级标题，用户的操作等于没生效。
    expect(styleRefs(pmToOoxml(headingDoc(2, "2"), { styles: TWO_HEADINGS }))).toEqual([
      `<w:pStyle w:val="3"/>`,
    ]);
  });

  it("级别没变时样式 id 一字不动", () => {
    expect(styleRefs(pmToOoxml(headingDoc(1, "2"), { styles: TWO_HEADINGS }))).toEqual([
      `<w:pStyle w:val="2"/>`,
    ]);
  });

  it("标题转成正文后，标题样式的引用必须消失", () => {
    expect(styleRefs(pmToOoxml(paraDoc({ styleId: "2" }), { styles: TWO_HEADINGS }))).toEqual([]);
  });

  it("样式表里没有这一级标题样式时退回推断 id，不留下对不上的引用", () => {
    expect(styleRefs(pmToOoxml(headingDoc(4, "2"), { styles: TWO_HEADINGS }))).toEqual([
      `<w:pStyle w:val="Heading4"/>`,
    ]);
  });

  it("编辑器里新建的标题也先问样式表，别写本文档没声明的 Heading2", () => {
    // WPS 写的文档里 heading 2 的 styleId 是 "3"；照节点类型推断会写成 "Heading2"，
    // 而这个名字在这份 styles.xml 里根本不存在，Word 只会当普通段落排。
    expect(styleRefs(pmToOoxml(headingDoc(2), { styles: TWO_HEADINGS }))).toEqual([
      `<w:pStyle w:val="3"/>`,
    ]);
    // 样式表没有这一级时才退回推断——新建文档连 styles.xml 都没有，行为照旧。
    expect(styleRefs(pmToOoxml(headingDoc(4), { styles: TWO_HEADINGS }))).toEqual([
      `<w:pStyle w:val="Heading4"/>`,
    ]);
  });

  it("没有样式表可对照时不猜：非标题样式一律放行", () => {
    // 只有样式表知道这个 id 是不是标题样式，无从判断就不能删用户的引用。
    expect(styleRefs(pmToOoxml(paraDoc({ styleId: "af3" }), { styles: TWO_HEADINGS }))).toEqual([
      `<w:pStyle w:val="af3"/>`,
    ]);
    expect(styleRefs(pmToOoxml(headingDoc(2, "2")))).toEqual([`<w:pStyle w:val="2"/>`]);
  });
});

describe("域（w:fldChar / w:instrText / w:fldSimple）", () => {
  /** Word 自己生成的目录条目：外层 TOC 域的结果里套超链接 + PAGEREF 子域。 */
  const TOC_ENTRY = bodyXml(
    `<w:p>` +
      `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
      `<w:r><w:instrText xml:space="preserve"> TOC \\o &quot;1-3&quot; \\h \\z \\u </w:instrText></w:r>` +
      `<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
      `<w:hyperlink w:anchor="_Toc1">` +
      `<w:r><w:t>第一章</w:t></w:r><w:r><w:tab/></w:r>` +
      `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
      `<w:r><w:instrText xml:space="preserve"> PAGEREF _Toc1 \\h </w:instrText></w:r>` +
      `<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
      `<w:r><w:t>3</w:t></w:r>` +
      `<w:r><w:fldChar w:fldCharType="end"/></w:r>` +
      `</w:hyperlink>` +
      `<w:r><w:fldChar w:fldCharType="end"/></w:r>` +
      `</w:p>`,
  );

  it("嵌套域不再劫持外层：一条 TOC 只产出一个 field 节点", () => {
    const para = ooxmlToPm(TOC_ENTRY).content[0]!;
    expect(para.content).toEqual([
      {
        type: "field",
        attrs: {
          kind: "TOC",
          instr: `TOC \\o "1-3" \\h \\z \\u`,
          display: "第一章\t3",
        },
      },
    ]);
  });

  it("写回时把 \\t 还原成 <w:tab/>，再读一遍还是同一个域", () => {
    const once = pmToOoxml(ooxmlToPm(TOC_ENTRY));
    expect(once).toContain(
      `<w:r><w:t xml:space="preserve">第一章</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t xml:space="preserve">3</w:t></w:r>`,
    );
    // 子域的指令不能混进外层：外层 instr 只有一句 TOC。
    expect(once.match(/PAGEREF/g)).toBeNull();
    expect(pmToOoxml(ooxmlToPm(once))).toBe(once);
  });

  /** begin 后直接 end：Word 插了域但还没算过结果。 */
  const UNCALCULATED = bodyXml(
    `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> DATE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`,
  );

  it("Word 没算过结果的域不编占位文字", () => {
    expect(ooxmlToPm(UNCALCULATED).content[0]!.content?.[0]).toMatchObject({
      type: "field",
      attrs: { instr: "DATE", display: "" },
    });
    // 写出去也不能凭空长出 "{DATE}" 这种用户没写过的文字。
    const xml = pmToOoxml(ooxmlToPm(UNCALCULATED));
    expect(xml).not.toContain("<w:t");
    expect(pmToOoxml(ooxmlToPm(xml))).toBe(xml);
  });

  it("简单域的缓存结果按字面保留，首尾空格不 trim", () => {
    const para = ooxmlToPm(
      bodyXml(
        `<w:p><w:fldSimple w:instr=" PAGE   \\* MERGEFORMAT  "><w:r><w:t xml:space="preserve"> 12 </w:t></w:r></w:fldSimple></w:p>`,
      ),
    ).content[0]!;
    expect(para.content?.[0]).toMatchObject({
      type: "field",
      attrs: { kind: "PAGE", instr: `PAGE   \\* MERGEFORMAT`, display: " 12 " },
    });
  });

  /**
   * Word 真正写出的多行目录：begin+instr+separate 在首段，end 在末段之后，
   * 每条条目里的页码是各自独立的一对 begin/end（PAGEREF 子域）。
   */
  const MULTI_PARA_TOC = bodyXml(
    `<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr>` +
      `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
      `<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" </w:instrText></w:r>` +
      `<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
      `<w:r><w:t>第一章</w:t></w:r><w:r><w:tab/></w:r>` +
      `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
      `<w:r><w:instrText xml:space="preserve"> PAGEREF _Toc1 </w:instrText></w:r>` +
      `<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
      `<w:r><w:t>3</w:t></w:r>` +
      `<w:r><w:fldChar w:fldCharType="end"/></w:r>` +
      `</w:p>`,
    `<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr>` +
      `<w:r><w:t>第二章</w:t></w:r><w:r><w:tab/></w:r>` +
      `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
      `<w:r><w:instrText xml:space="preserve"> PAGEREF _Toc2 </w:instrText></w:r>` +
      `<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
      `<w:r><w:t>5</w:t></w:r>` +
      `<w:r><w:fldChar w:fldCharType="end"/></w:r>` +
      `<w:r><w:fldChar w:fldCharType="end"/></w:r>` +
      `</w:p>`,
  );

  it("跨段落的域用 fieldBegin/fieldEnd 建模，条目内容不再被吞成缓存文本", () => {
    const pm = ooxmlToPm(MULTI_PARA_TOC);
    expect(pm.content[0]!.content?.[0]).toMatchObject({
      type: "fieldBegin",
      attrs: { kind: "TOC", instr: `TOC \\o "1-3"`, isSeparated: true },
    });
    // 边界之后是正文结构：标题文字、制表符、以及条目自己的 PAGEREF 芯片。
    expect(pm.content[0]!.content?.slice(1)).toEqual([
      { type: "text", text: "第一章\t" },
      { type: "field", attrs: { kind: "PAGEREF", instr: "PAGEREF _Toc1", display: "3" } },
    ]);
    const last = pm.content[1]!.content!;
    expect(last[last.length - 1]).toEqual({ type: "fieldEnd" });
  });

  it("跨段的域写回去还是跨段的：指令只进 instrText、末尾补 end 记号", () => {
    const once = pmToOoxml(ooxmlToPm(MULTI_PARA_TOC));
    // 写侧按属性式转义，`"` 变成 `&quot;`（读回来仍是原字符，往返一致即可）。
    expect(once).toContain(
      `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o &quot;1-3&quot; </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>`,
    );
    // 目录指令绝不能变成看得见的文字。
    expect(once).not.toContain(`<w:t xml:space="preserve"> TOC`);
    expect(once.match(/PAGEREF _Toc1/g)).toHaveLength(1);
    expect(once.match(/fldCharType="end"/g)).toHaveLength(3);
    expect(pmToOoxml(ooxmlToPm(once))).toBe(once);
  });

  /** 有 begin 有指令，但既没 separate 也没 end：Word 标成 dirty、等着自己重算。 */
  const DIRTY_TOC = bodyXml(
    `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
      `<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" </w:instrText></w:r>` +
      `<w:r><w:t>旧目录第一条</w:t></w:r></w:p>`,
  );

  it("没写到 separate 的起始边界也建模，域内容退回普通文字", () => {
    const para = ooxmlToPm(DIRTY_TOC).content[0]!;
    expect(para.content).toEqual([
      { type: "fieldBegin", attrs: { kind: "TOC", instr: `TOC \\o "1-3"`, isSeparated: false } },
      { type: "text", text: "旧目录第一条" },
    ]);
    const xml = pmToOoxml(ooxmlToPm(DIRTY_TOC));
    // 不凭空补 separate：Word 重算时才需要结果区，写侧不该替它决定。
    expect(xml).not.toContain(`fldCharType="separate"`);
    expect(pmToOoxml(ooxmlToPm(xml))).toBe(xml);
  });

  it("只有 end 的段落不造出幽灵起始边界", () => {
    const para = ooxmlToPm(bodyXml(`<w:p><w:r><w:t>正文</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`))
      .content[0]!;
    expect(para.content).toEqual([{ type: "text", text: "正文" }, { type: "fieldEnd" }]);
  });

  /**
   * Word「插入题注」写出的那一段：`SEQ` 域给编号，编号外面套一条 `_Ref…` 书签，
   * 交叉引用（⑤.2 的 `REF`/`PAGEREF`）点的就是这条书签。
   */
  const CAPTION = bodyXml(
    `<w:p>` +
      `<w:pPr><w:pStyle w:val="Caption"/></w:pPr>` +
      `<w:r><w:t xml:space="preserve">图 </w:t></w:r>` +
      `<w:bookmarkStart w:id="3" w:name="_Ref123456"/>` +
      `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
      `<w:r><w:instrText xml:space="preserve"> SEQ 图 \\* ARABIC </w:instrText></w:r>` +
      `<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
      `<w:r><w:t>2</w:t></w:r>` +
      `<w:r><w:fldChar w:fldCharType="end"/></w:r>` +
      `<w:bookmarkEnd w:id="3"/>` +
      `<w:r><w:t xml:space="preserve"> 系统架构</w:t></w:r>` +
      `</w:p>`,
  );

  it("题注：SEQ 域读成一条域，外面的书签留在段落这一层", () => {
    const para = ooxmlToPm(CAPTION).content[0]!;
    expect(para.content).toEqual([
      { type: "text", text: "图 " },
      { type: "bookmarkStart", attrs: { id: "3", name: "_Ref123456" } },
      { type: "field", attrs: { kind: "SEQ", instr: `SEQ 图 \\* ARABIC`, display: "2" } },
      { type: "bookmarkEnd", attrs: { id: "3" } },
      { type: "text", text: " 系统架构" },
    ]);
  });

  it("题注写回去：指令进 instrText、书签裹着域，逐字节回来", () => {
    const once = pmToOoxml(ooxmlToPm(CAPTION));
    expect(once).toContain(
      `<w:r><w:instrText xml:space="preserve"> SEQ 图 \\* ARABIC </w:instrText></w:r>`,
    );
    // 书签包进 run 就成了 run 的内容，Word 读不认，书签等于没写（§17.5.2.1）。
    expect(once).not.toContain(`<w:r><w:bookmarkStart`);
    expect(pmToOoxml(ooxmlToPm(once))).toBe(once);
  });

  /**
   * ⑤.2 交叉引用：本 app 写出的 `REF` / `PAGEREF` 域。
   * `display` 是缓存结果——刚插入时是空串（还没让 Word 算过），指向题注时是编号或页码。
   */
  const CROSSREF_PM: PMDoc = {
    type: "doc",
    content: [
      {
        type: "paragraph",
        content: [
          { type: "text", text: "见" },
          { type: "field", attrs: { kind: "REF", instr: `REF _Ref2 \\h`, display: "" } },
          { type: "text", text: "，页码在" },
          { type: "field", attrs: { kind: "PAGEREF", instr: `PAGEREF _Ref2 \\h`, display: "5" } },
        ],
      },
    ],
  };

  it("交叉引用写出去：指令只在 instrText，页码才是缓存结果", () => {
    const xml = pmToOoxml(CROSSREF_PM);
    expect(xml).toContain(
      `<w:r><w:instrText xml:space="preserve"> REF _Ref2 \\h </w:instrText></w:r>`,
    );
    // 指令变成看得见的正文，Word 里就会真的印出 `{ REF _Ref2 \h }`。
    expect(xml).not.toContain(`<w:t xml:space="preserve"> REF`);
    expect(xml).toContain(`<w:t xml:space="preserve">5</w:t>`);
    expect(pmToOoxml(ooxmlToPm(xml))).toBe(xml);
  });

  it("交叉引用读回来：域还是域，书签名与 \\h 开关都不丢", () => {
    const para = ooxmlToPm(pmToOoxml(CROSSREF_PM)).content[0]!;
    expect(para.content?.[1]).toEqual({
      type: "field",
      attrs: { kind: "REF", instr: `REF _Ref2 \\h`, display: "" },
    });
    expect(para.content?.[3]).toMatchObject({
      type: "field",
      attrs: { kind: "PAGEREF", display: "5" },
    });
  });

  /**
   * Word 自己写的交叉引用：整条域被 `w:hyperlink w:anchor` 包着（`\h` 的产物）。
   * 链接里的域也得走域状态机——只按 run 抽文字的话指令丢掉，存回去就是死文字。
   */
  const WORD_CROSSREF = bodyXml(
    `<w:p><w:r><w:t>如</w:t></w:r>` +
      `<w:hyperlink w:anchor="_Ref10">` +
      `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
      `<w:r><w:instrText xml:space="preserve"> PAGEREF _Ref10 \\h </w:instrText></w:r>` +
      `<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
      `<w:r><w:t>2</w:t></w:r>` +
      `<w:r><w:fldChar w:fldCharType="end"/></w:r>` +
      `</w:hyperlink>` +
      `<w:r><w:t>所示</w:t></w:r></w:p>`,
  );

  it("Word 写在超链接里的交叉引用：指令留着，不折成死文字", () => {
    const para = ooxmlToPm(WORD_CROSSREF).content[0]!;
    expect(para.content).toEqual([
      { type: "text", text: "如" },
      {
        type: "field",
        attrs: { kind: "PAGEREF", instr: `PAGEREF _Ref10 \\h`, display: "2" },
      },
      { type: "text", text: "所示" },
    ]);
    // 链接只落在文字上（addMark 的既有约定），这条 `w:hyperlink` 里只有域、没有文字，
    // 所以存回去时链接不重发：`\h` 写在指令里，Word 更新域时自己把超链接补上。
    const once = pmToOoxml(ooxmlToPm(WORD_CROSSREF));
    expect(once).toContain(
      `<w:r><w:instrText xml:space="preserve"> PAGEREF _Ref10 \\h </w:instrText></w:r>`,
    );
    expect(once).not.toContain(`<w:t xml:space="preserve"> PAGEREF`);
    expect(pmToOoxml(ooxmlToPm(once))).toBe(once);
  });

  it("站内锚点往返一致：w:anchor 装的是书签名，`#` 不能越存越多", () => {
    const xml = pmToOoxml(
      ooxmlToPm(
        bodyXml(`<w:p><w:hyperlink w:anchor="_Toc9"><w:r><w:t>第一章</w:t></w:r></w:hyperlink></w:p>`),
      ),
    );
    expect(xml).toContain(`<w:hyperlink w:anchor="_Toc9">`);
    expect(pmToOoxml(ooxmlToPm(xml))).toBe(xml);
  });
});

describe("书签（w:bookmarkStart / w:bookmarkEnd）", () => {
  /** 一段里夹一个起点、下一段收尾：正是 ECMA-376 §17.5.2.1 例子那种跨段书签。 */
  const SPANNING = bodyXml(
    `<w:p><w:r><w:t>前半</w:t></w:r><w:bookmarkStart w:id="0" w:name="testing123"/><w:r><w:t>后半</w:t></w:r></w:p>`,
    `<w:p><w:r><w:t>下段</w:t></w:r><w:bookmarkEnd w:id="0"/></w:p>`,
  );

  it("reads bookmarks as inline nodes in document order", () => {
    const pm = ooxmlToPm(SPANNING);
    const first = pm.content[0]!;
    expect(first.content?.map((n) => n.type)).toEqual([
      "text",
      "bookmarkStart",
      "text",
    ]);
    expect(first.content?.[1]).toEqual({
      type: "bookmarkStart",
      attrs: { id: "0", name: "testing123" },
    });
    expect(pm.content[1]!.content?.[1]).toEqual({
      type: "bookmarkEnd",
      attrs: { id: "0" },
    });
  });

  it("writes bookmarks as direct children of w:p, never inside w:r", () => {
    const xml = pmToOoxml(ooxmlToPm(SPANNING));
    expect(xml).toContain(
      `<w:t xml:space="preserve">前半</w:t></w:r><w:bookmarkStart w:id="0" w:name="testing123"/><w:r>`,
    );
    expect(xml).toContain(`<w:bookmarkEnd w:id="0"/></w:p>`);
  });

  it("is idempotent — 书签不参与文字，两次写出一致", () => {
    const once = pmToOoxml(ooxmlToPm(SPANNING));
    expect(pmToOoxml(ooxmlToPm(once))).toBe(once);
  });

  it("drops a bookmark whose other half the editor deleted", () => {
    // 起点在、终点被删：写出去就是 non-conformant 的文档，整对不写。
    const orphanStart = bodyXml(
      `<w:p><w:r><w:t>x</w:t></w:r><w:bookmarkStart w:id="7" w:name="gone"/></w:p>`,
    );
    const xml = pmToOoxml(ooxmlToPm(orphanStart));
    expect(xml).not.toContain("bookmark");
    // 只剩终点（起点在别的段落被删）同样不写。
    const orphanEnd = bodyXml(`<w:p><w:r><w:t>y</w:t></w:r><w:bookmarkEnd w:id="9"/></w:p>`);
    expect(pmToOoxml(ooxmlToPm(orphanEnd))).not.toContain("bookmarkEnd");
  });

  it("keeps bookmarks that live inside a heading paragraph", () => {
    // 目录的 _Toc 书签就挂在标题段里，标题丢了引用目录就没了锚点。
    const pm = ooxmlToPm(
      bodyXml(
        `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:bookmarkStart w:id="2" w:name="_Toc5"/><w:r><w:t>标题</w:t></w:r><w:bookmarkEnd w:id="2"/></w:p>`,
      ),
    );
    const h = pm.content[0]!;
    expect(h.type).toBe("heading");
    expect(h.content?.map((n) => n.type)).toEqual(["bookmarkStart", "text", "bookmarkEnd"]);
    const xml = pmToOoxml(pm);
    expect(xml).toContain(`<w:bookmarkStart w:id="2" w:name="_Toc5"/>`);
    expect(xml).toContain(`<w:bookmarkEnd w:id="2"/>`);
  });

  it("escapes a bookmark name that carries markup", () => {
    const pm: PMDoc = {
      type: "doc",
      content: [
        {
          type: "paragraph",
          content: [
            { type: "bookmarkStart", attrs: { id: "1", name: `a<b&"c` } },
            { type: "text", text: "x" },
            { type: "bookmarkEnd", attrs: { id: "1" } },
          ],
        },
      ],
    };
    expect(pmToOoxml(pm)).toContain(`<w:bookmarkStart w:id="1" w:name="a&lt;b&amp;&quot;c"/>`);
  });

  it("ignores a bookmark id that is not a number — 读不懂就不建模", () => {
    const pm = ooxmlToPm(
      bodyXml(
        `<w:p><w:bookmarkStart w:id="abc" w:name="odd"/><w:r><w:t>x</w:t></w:r><w:bookmarkEnd w:id="abc"/></w:p>`,
      ),
    );
    expect(pm.content[0]!.content?.map((n) => n.type)).toEqual(["text"]);
  });
});

/**
 * 排名 4 ①：批注的三个锚点。
 *
 * 内容模型给出的位置不一样：`w:commentRangeStart`/`w:commentRangeEnd` 属于
 * EG_RangeMarkupElements（`w:p` 的直接子元素，与 `w:r` 平级），而 `w:commentReference`
 * 属于 EG_RunInnerContent（必须装在 `w:r` 里面）。写出时必须保持这个差别。
 */
describe("批注锚点（w:commentRangeStart / End / Reference）", () => {
  /** 规范样例的形状：同段一个完整区间，引用记号紧跟在区间末尾。 */
  const INLINE = bodyXml(
    `<w:p><w:r><w:t>Some </w:t></w:r><w:commentRangeStart w:id="0"/>` +
      `<w:r><w:t>text.</w:t></w:r><w:commentRangeEnd w:id="0"/>` +
      `<w:r><w:rPr><w:rStyle w:val="af"/></w:rPr><w:commentReference w:id="0"/></w:r></w:p>`,
  );

  it("reads the three anchors as inline nodes in document order", () => {
    const pm = ooxmlToPm(INLINE);
    expect(pm.content[0]!.content?.map((n) => n.type)).toEqual([
      "text",
      "commentRangeStart",
      "text",
      "commentRangeEnd",
      "commentReference",
    ]);
    expect(pm.content[0]!.content?.[1]).toEqual({
      type: "commentRangeStart",
      attrs: { id: "0" },
    });
    expect(pm.content[0]!.content?.[4]).toEqual({ type: "commentReference", attrs: { id: "0" } });
  });

  it("keeps a range that spans paragraphs, one anchor per paragraph", () => {
    const pm = ooxmlToPm(
      bodyXml(
        `<w:p><w:r><w:t>前半</w:t></w:r><w:commentRangeStart w:id="3"/></w:p>`,
        `<w:p><w:commentRangeEnd w:id="3"/><w:r><w:t>后半</w:t></w:r></w:p>`,
      ),
    );
    expect(pm.content[0]!.content?.at(-1)).toEqual({ type: "commentRangeStart", attrs: { id: "3" } });
    expect(pm.content[1]!.content?.[0]).toEqual({ type: "commentRangeEnd", attrs: { id: "3" } });
  });

  it("writes range markers as direct children of w:p, never inside w:r", () => {
    const xml = pmToOoxml(ooxmlToPm(INLINE));
    expect(xml).toContain(
      `<w:t xml:space="preserve">Some </w:t></w:r><w:commentRangeStart w:id="0"/><w:r>`,
    );
    expect(xml).toContain(`</w:r><w:commentRangeEnd w:id="0"/><w:r>`);
    expect(xml).not.toMatch(/<w:r>[^<]*<w:commentRange/);
  });

  it("drops the source run's rPr on the reference — 记号形状由我们固定", () => {
    // Word 给这个记号排版的是文档里那条批注引用字符样式，它的 styleId 是 `af` 这类
    // 逐文档随机的值，抄不进另一份文档；宁可不写也不原样塞进去（§2.2 同一个道理）。
    const xml = pmToOoxml(ooxmlToPm(INLINE));
    expect(xml).toContain(`<w:r><w:commentReference w:id="0"/></w:r>`);
    expect(xml).not.toContain(`<w:rPr><w:rStyle w:val="af"/></w:rPr><w:commentReference`);
  });

  it("is idempotent — 锚点不含文字，两次写出一致", () => {
    const once = pmToOoxml(ooxmlToPm(INLINE));
    expect(pmToOoxml(ooxmlToPm(once))).toBe(once);
  });

  it("writes an empty range as-is — start 紧跟 end 也是合法区间", () => {
    const pm: PMDoc = {
      type: "doc",
      content: [
        {
          type: "paragraph",
          content: [
            { type: "text", text: "整段被批注" },
            { type: "commentRangeStart", attrs: { id: "1" } },
            { type: "commentRangeEnd", attrs: { id: "1" } },
            { type: "commentReference", attrs: { id: "1" } },
          ],
        },
      ],
    };
    expect(pmToOoxml(pm)).toContain(
      `<w:commentRangeStart w:id="1"/><w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r>`,
    );
  });

  it("ignores anchors whose id is not a number — 读不懂就不建模", () => {
    const pm = ooxmlToPm(
      bodyXml(
        `<w:p><w:commentRangeStart w:id="0x"/><w:r><w:t>x</w:t></w:r><w:commentRangeEnd w:id="0x"/>` +
          `<w:r><w:commentReference w:id="0x"/></w:r></w:p>`,
      ),
    );
    expect(pm.content[0]!.content?.map((n) => n.type)).toEqual(["text"]);
  });

  it("shares the numbering space with bookmarks — 同 id 的书签与批注互不干扰", () => {
    const pm = ooxmlToPm(
      bodyXml(
        `<w:p><w:bookmarkStart w:id="0" w:name="_Toc1"/><w:commentRangeStart w:id="0"/>` +
          `<w:r><w:t>x</w:t></w:r><w:commentRangeEnd w:id="0"/><w:bookmarkEnd w:id="0"/></w:p>`,
      ),
    );
    expect(pm.content[0]!.content?.map((n) => n.type)).toEqual([
      "bookmarkStart",
      "commentRangeStart",
      "text",
      "commentRangeEnd",
      "bookmarkEnd",
    ]);
  });
});

describe("修订跟踪（w:ins / w:del / w:delText）", () => {
  /** Word 写修订的形状：`w:ins`/`w:del` 与 `w:r` 平级，删除的文字换成 `w:delText`。 */
  const TRACKED = bodyXml(
    `<w:p><w:r><w:t>正文</w:t></w:r>` +
      `<w:ins w:id="1" w:author="张三" w:date="2026-01-02T03:04:05Z"><w:r><w:t>插入的字</w:t></w:r></w:ins>` +
      `<w:del w:id="2" w:author="李四"><w:r><w:delText>删掉的字</w:delText></w:r></w:del></w:p>`,
  );

  it("reads ins/del into a revision mark instead of dropping the deleted text", () => {
    const inlines = ooxmlToPm(TRACKED).content[0]!.content!;
    expect(inlines.map((n) => n.text)).toEqual(["正文", "插入的字", "删掉的字"]);
    expect(inlines[0]!.marks ?? []).toEqual([]);
    expect(inlines[1]!.marks).toEqual([
      { type: "revision", attrs: { kind: "ins", id: "1", author: "张三", date: "2026-01-02T03:04:05Z" } },
    ]);
    expect(inlines[2]!.marks).toEqual([
      { type: "revision", attrs: { kind: "del", id: "2", author: "李四" } },
    ]);
  });

  it("writes the region back with the same element, attributes and delText", () => {
    const once = pmToOoxml(ooxmlToPm(TRACKED));
    expect(once).toContain(
      `<w:ins w:id="1" w:author="张三" w:date="2026-01-02T03:04:05Z"><w:r><w:t xml:space="preserve">插入的字</w:t></w:r></w:ins>`,
    );
    expect(once).toContain(
      `<w:del w:id="2" w:author="李四"><w:r><w:delText xml:space="preserve">删掉的字</w:delText></w:r></w:del>`,
    );
    expect(pmToOoxml(ooxmlToPm(once))).toBe(once);
  });

  it("merges the split runs of one revision into a single element", () => {
    // Word 的一次插入是若干个 run（各有自己的字符样式）；按 marks 分组必然切开，
    // 逐组各包一次会让同一个 `w:id` 出现两条，而"接受此修订"是按 id 找对象的。
    const xml = pmToOoxml(
      ooxmlToPm(
        bodyXml(
          `<w:p><w:ins w:id="5" w:author="张三">` +
            `<w:r><w:t>普通</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>加粗</w:t></w:r></w:ins></w:p>`,
        ),
      ),
    );
    expect(xml.match(/<w:ins /g)).toHaveLength(1);
    expect(xml.match(/<\/w:ins>/g)).toHaveLength(1);
    const region = xml.match(/<w:ins [^>]*>([\s\S]*?)<\/w:ins>/)![1]!;
    expect(region.match(/<w:r>/g)).toHaveLength(2);
    expect(region).toContain("普通");
    expect(region).toContain("加粗");
  });

  it("keeps the outer operation when ins nests in del", () => {
    // 先插入、后删除：Word 把后来的那次写在外面，显示状态由外层决定，所以标记取外层。
    const pm = ooxmlToPm(
      bodyXml(
        `<w:p><w:del w:id="8" w:author="王五"><w:ins w:id="7" w:author="张三">` +
          `<w:r><w:delText>先插后删</w:delText></w:r></w:ins></w:del></w:p>`,
      ),
    );
    expect(pm.content[0]!.content![0]!.marks).toEqual([
      { type: "revision", attrs: { kind: "del", id: "8", author: "王五" } },
    ]);
    const xml = pmToOoxml(pm);
    expect(xml).not.toContain("<w:ins");
    expect(xml).toContain(
      `<w:del w:id="8" w:author="王五"><w:r><w:delText xml:space="preserve">先插后删</w:delText></w:r></w:del>`,
    );
  });

  it("puts the revision inside the hyperlink — 只有这个方向有规范依据", () => {
    const xml = pmToOoxml(
      ooxmlToPm(
        bodyXml(
          `<w:p><w:hyperlink r:id="rId5"><w:ins w:id="9" w:author="张三">` +
            `<w:r><w:t>链接里的插入</w:t></w:r></w:ins></w:hyperlink></w:p>`,
        ),
      ),
    );
    expect(xml).toContain(`<w:hyperlink r:id="rId5"><w:ins w:id="9" w:author="张三"><w:r>`);
    expect(xml).not.toMatch(/<w:ins[^>]*><w:hyperlink/);
  });

  it("carries a paragraph-mark revision through pPr/rPr untouched", () => {
    // 段落标记自己的增删在 `w:pPr/w:rPr` 里，没有运行可走；这份靠 rawPpr 原样透传。
    const xml = pmToOoxml(
      ooxmlToPm(
        bodyXml(
          `<w:p><w:pPr><w:rPr><w:ins w:id="3" w:author="张三" w:date="2026-01-02T03:04:05Z"/></w:rPr></w:pPr>` +
            `<w:r><w:t>本段被插入</w:t></w:r></w:p>`,
        ),
      ),
    );
    expect(xml).toContain(`<w:rPr><w:ins w:id="3" w:author="张三" w:date="2026-01-02T03:04:05Z"/></w:rPr>`);
    expect(xml).not.toContain(`w:author="张三"><w:r>`);
  });

  it("invents no author when the document did not name one", () => {
    const xml = pmToOoxml(
      ooxmlToPm(bodyXml(`<w:p><w:ins w:id="4"><w:r><w:t>x</w:t></w:r></w:ins></w:p>`)),
    );
    expect(xml).toContain(`<w:ins w:id="4"><w:r>`);
    expect(xml).not.toContain("w:author");
  });

  it("keeps a deleted tab in the same deletion — 制表符也是被删掉的内容", () => {
    const xml = pmToOoxml(
      ooxmlToPm(bodyXml(`<w:p><w:del w:id="6"><w:r><w:delText>a</w:delText></w:r><w:r><w:tab/></w:r></w:del></w:p>`)),
    );
    expect(xml).toContain(
      `<w:del w:id="6"><w:r><w:delText xml:space="preserve">a</w:delText><w:tab/></w:r></w:del>`,
    );
  });
});

describe("本 app 生成的目录写进 docx", () => {
  /**
   * `apps/desktop/src/editor/toc.ts` 的 `tocEntryJson` 产出的那份 PM 文档，原样搬来：
   * 制表位位置 595.5 磅是 794 CSS px 正文区宽度经 `cssPxToPt` 换来的（= 11910 twips），
   * 二级条目缩进 20 磅（= 400 twips）。
   * 条目不写 `w:pStyle`——本文档的样式表里没有 TOC1/TOC2，写了就是悬空引用。
   */
  const rightDotStop: TabStop = { positionPt: 595.5, alignment: "right", leader: "dot" };
  const GENERATED: PMDoc = {
    type: "doc",
    content: [
      {
        type: "paragraph",
        attrs: { tabStops: [rightDotStop] },
        content: [
          {
            type: "fieldBegin",
            attrs: { kind: "TOC", instr: 'TOC \\o "1-2" \\h \\z \\u', isSeparated: true },
          },
          { type: "text", text: "第一章\t" },
          { type: "field", attrs: { kind: "PAGEREF", instr: "PAGEREF _Toc1 \\h", display: "1" } },
        ],
      },
      {
        type: "paragraph",
        attrs: { indent: 20, tabStops: [rightDotStop] },
        content: [
          { type: "text", text: "小节\t" },
          { type: "field", attrs: { kind: "PAGEREF", instr: "PAGEREF _Toc2 \\h", display: "2" } },
          { type: "fieldEnd" },
        ],
      },
      {
        type: "heading",
        attrs: { level: 1 },
        content: [
          { type: "bookmarkStart", attrs: { id: "1", name: "_Toc1" } },
          { type: "text", text: "第一章" },
          { type: "bookmarkEnd", attrs: { id: "1" } },
        ],
      },
    ],
  };

  it("写出 Word 认的形状：条目各自一条 w:tabs，页码前的制表符是 w:tab", () => {
    const xml = pmToOoxml(GENERATED);
    expect(xml.match(/<w:tabs>/g)).toHaveLength(2);
    expect(xml.match(/w:pos="11910"/g)).toHaveLength(2);
    // CT_PPr 的顺序：w:tabs 在 w:ind 之前。
    expect(xml).toContain(
      '<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="11910"/></w:tabs><w:ind w:left="400"/>',
    );
    expect(xml).toContain('<w:r><w:t xml:space="preserve">第一章</w:t><w:tab/></w:r>');
    // 外层只有一条 TOC 指令；两条 PAGEREF 子域各自成对，结束记号共 3 个（两条页码 + 目录本身）。
    expect(xml.match(/TOC \\o &quot;1-2&quot; \\h \\z \\u/g)).toHaveLength(1);
    expect(xml.match(/PAGEREF _Toc/g)).toHaveLength(2);
    expect(xml.match(/fldCharType="begin"/g)).toHaveLength(3);
    expect(xml.match(/fldCharType="end"/g)).toHaveLength(3);
    // 书签仍是 w:p 的直接子元素，名字与 PAGEREF 指向的一致。
    expect(xml).toContain(
      '<w:bookmarkStart w:id="1" w:name="_Toc1"/><w:r><w:t xml:space="preserve">第一章</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>',
    );
  });

  it("读回来还是同一份结构，二次保存字节一致", () => {
    const once = pmToOoxml(GENERATED);
    const back = ooxmlToPm(once);
    expect(back.content?.[0]?.content).toEqual([
      {
        type: "fieldBegin",
        attrs: { kind: "TOC", instr: 'TOC \\o "1-2" \\h \\z \\u', isSeparated: true },
      },
      { type: "text", text: "第一章\t" },
      { type: "field", attrs: { kind: "PAGEREF", instr: "PAGEREF _Toc1 \\h", display: "1" } },
    ]);
    expect(back.content?.[1]?.content?.at(-1)).toEqual({ type: "fieldEnd" });
    expect(back.content?.[1]?.attrs).toMatchObject({
      indent: 20,
      tabStops: [{ positionPt: 595.5, alignment: "right", leader: "dot" }],
    });
    expect(back.content?.[2]?.content).toEqual([
      { type: "bookmarkStart", attrs: { id: "1", name: "_Toc1" } },
      { type: "text", text: "第一章" },
      { type: "bookmarkEnd", attrs: { id: "1" } },
    ]);
    expect(pmToOoxml(back)).toBe(once);
  });

  it("版心宽度从 CSS px 换成磅只在这一处发生，且落在整数 twips 上", () => {
    // 编辑器/SectionGeometry 用 96DPI 的 px，OOXML 用 twips：595.5 磅 = 11910 twips。
    expect(cssPxToPt(794)).toBe(595.5);
    // Word 真实目录条目的制表位位置 8306 twips 对应 A4 减边距后的正文宽度。
    expect(cssPxToPt(553.7)).toBe(415.3);
  });

  it("写进真实 .docx 再重开，读回的还是同一份目录", async () => {
    // 走一遍包（[Content_Types].xml + rels + document.xml），证明这份 XML 不只是
    // 字符串能对上，而是 DocxFile 认得的部件；书签与域都不需要额外关系。
    const docx = await DocxFile.fromBytes(await minimalDocx(pmToOoxml(GENERATED)));
    const { doc } = await docx.toProseMirror();
    expect(doc.content?.[0]?.content?.map((n) => n.type)).toEqual(["fieldBegin", "text", "field"]);
    expect(doc.content?.[0]?.attrs?.tabStops).toEqual([rightDotStop]);
    await docx.replaceProseMirror(doc);
    const reopened = await DocxFile.fromBytes(await docx.toBytes());
    expect((await reopened.toProseMirror()).doc).toEqual(doc);
  });
});

describe("DocxFile toProseMirror / replaceProseMirror", () => {
  it("round-trips a minimal docx via zip", async () => {
    const bytes = await minimalDocx(
      bodyXml(`<w:p><w:r><w:t>alpha</w:t></w:r></w:p>`, `<w:p><w:r><w:t>beta</w:t></w:r></w:p>`),
    );
    const docx = await DocxFile.fromBytes(bytes);
    const { doc } = await docx.toProseMirror();
    expect(doc.content.length).toBeGreaterThanOrEqual(2);
    await docx.replaceProseMirror(doc);
    const out = await docx.toBytes();
    const reopen = await DocxFile.fromBytes(out);
    const paras = reopen.listParagraphs().map((p) => p.text);
    expect(paras.join("|")).toContain("alpha");
    expect(paras.join("|")).toContain("beta");
  });

  it("keeps tracked insertions and deletions across open → save", async () => {
    // 桌面 app 每次保存都是"读原文件 → 换正文 → 写回"，这条就是那条链路：
    // 一次存盘不能把修订悄悄接受掉，也不能把被删的文字弄没。
    const bytes = await minimalDocx(
      bodyXml(
        `<w:p><w:r><w:t>正文</w:t></w:r>` +
          `<w:del w:id="1" w:author="张三"><w:r><w:delText>删掉</w:delText></w:r></w:del></w:p>`,
      ),
    );
    const docx = await DocxFile.fromBytes(bytes);
    const { doc } = await docx.toProseMirror();
    await docx.replaceProseMirror(doc);
    const reopen = await DocxFile.fromBytes(await docx.toBytes());
    const again = (await reopen.toProseMirror()).doc;
    const inlines = again.content[0]!.content!;
    expect(inlines.map((n) => n.text).join("")).toBe("正文删掉");
    expect(inlines.at(-1)!.marks).toEqual([
      { type: "revision", attrs: { kind: "del", id: "1", author: "张三" } },
    ]);
  });
});
