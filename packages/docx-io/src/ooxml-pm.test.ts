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
import type { PMDoc } from "./pm-schema.js";
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
