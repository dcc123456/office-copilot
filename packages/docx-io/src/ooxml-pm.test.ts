import { describe, it, expect } from "vitest";
import JSZip from "jszip";
import {
  ooxmlToPm,
  pmToOoxml,
  ptToCssFontSize,
  cssFontSizeToPt,
  ptToCssLetterSpacing,
  cssColorToOoxmlHex,
  canonColor,
} from "./ooxml-pm.js";
import { StyleMap } from "./styles.js";
import type { PPrAttrs, PMDoc } from "./pm-schema.js";
import { DocxFile } from "./docx.js";

function bodyXml(...inner: string[]): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${inner.join("")}<w:sectPr/></w:body></w:document>`;
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
});

describe("DocxFile toProseMirror / replaceProseMirror", () => {
  it("round-trips a minimal docx via zip", async () => {
    const zip = new JSZip();
    zip.file(
      "[Content_Types].xml",
      `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
    );
    zip.file(
      "_rels/.rels",
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="word/document.xml"/></Relationships>`,
    );
    zip.file("word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`);
    zip.file(
      "word/document.xml",
      bodyXml(
        `<w:p><w:r><w:t>alpha</w:t></w:r></w:p>`,
        `<w:p><w:r><w:t>beta</w:t></w:r></w:p>`,
      ),
    );
    const bytes = await zip.generateAsync({ type: "uint8array" });
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
});
