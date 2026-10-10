import { describe, expect, it } from "vitest";
import { StyleMap } from "./styles.js";
import { resolvedStyles } from "./style-format.js";

/**
 * 一份浓缩的真实样式表（取自 WPS 写出的简历文档，只留下与解析相关的部分）：
 * Normal 是对齐与字号的来源，heading 1 只写了自己的那几项，其余靠 basedOn 继承。
 */
const WPS_STYLES =
  `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
  `<w:docDefaults><w:rPrDefault><w:rPr>` +
  `<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="宋体"/>` +
  `</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>` +
  `<w:style w:type="paragraph" w:default="1" w:styleId="1"><w:name w:val="Normal"/>` +
  `<w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr>` +
  `<w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia"/><w:sz w:val="21"/></w:rPr>` +
  `</w:style>` +
  `<w:style w:type="paragraph" w:styleId="2"><w:name w:val="heading 1"/><w:basedOn w:val="1"/>` +
  `<w:pPr><w:keepNext/><w:spacing w:before="340" w:after="330" w:line="576" w:lineRule="auto"/>` +
  `<w:outlineLvl w:val="0"/></w:pPr>` +
  `<w:rPr><w:b/><w:sz w:val="44"/></w:rPr>` +
  `</w:style>` +
  `<w:style w:type="paragraph" w:styleId="3"><w:name w:val="标题 1 反粗"/>` +
  `<w:basedOn w:val="2"/><w:rPr><w:b w:val="0"/></w:rPr></w:style>` +
  `<w:style w:type="character" w:styleId="9"><w:name w:val="强调整体"/><w:rPr><w:i/></w:rPr></w:style>` +
  `</w:styles>`;

function formatOf(styleId: string | null, xml = WPS_STYLES) {
  const hit = resolvedStyles(StyleMap.fromXml(xml)).find((s) => s.styleId === styleId);
  return hit?.format;
}

describe("样式链解析（排名 2 ⑤）", () => {
  it("默认段落样式那条排在最前，并且是 docDefaults + 文档的 default 样式", () => {
    const resolved = resolvedStyles(StyleMap.fromXml(WPS_STYLES));
    expect(resolved[0]?.styleId).toBeNull();
    // docDefaults 的字族 + Normal 自己的 10.5 磅（→ CSS px）与两端对齐。
    expect(resolved[0]?.format).toEqual({
      textAlign: "justify",
      fontFamily: "宋体",
      fontFamilyWest: "Times New Roman",
      fontSize: "14px",
    });
  });

  it("basedOn 链：标题 1 拿到 Normal 的对齐与字号之外的自己那套，段前段后是 pt", () => {
    expect(formatOf("2")).toEqual({
      textAlign: "justify",
      spacingBefore: 17,
      spacingAfter: 16.5,
      lineHeight: "2.4",
      fontFamily: "宋体",
      fontFamilyWest: "Times New Roman",
      fontSize: "29.3px",
      bold: true,
    });
  });

  it("`w:b w:val=\"0\"` 解析成显式的 false，不是\"没提过\"", () => {
    // 三态在这里必须有区别：孙样式要能把从父样式继承来的加粗关掉。
    expect(formatOf("2")?.bold).toBe(true);
    expect(formatOf("3")?.bold).toBe(false);
    expect(formatOf("3")?.fontSize).toBe("29.3px");
    const withoutBold = formatOf("2", `<w:styles xmlns:w="x"><w:style w:type="paragraph" w:styleId="2"/>`);
    expect(withoutBold?.bold).toBeUndefined();
  });

  it("只写主题字族的 docDefaults 不产出字族：不替文档编造一个 Word 用的字体名", () => {
    const themeOnly =
      `<w:styles xmlns:w="x"><w:docDefaults><w:rPrDefault><w:rPr>` +
      `<w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia"/>` +
      `</w:rPr></w:rPrDefault></w:docDefaults></w:styles>`;
    expect(resolvedStyles(StyleMap.fromXml(themeOnly))).toEqual([]);
  });

  it("没有 styles.xml（新建的空白文档）就一条都不产出", () => {
    expect(resolvedStyles(StyleMap.fromXml(undefined))).toEqual([]);
  });

  it("字符样式不参与：w:pStyle 只会指向段落样式", () => {
    const ids = resolvedStyles(StyleMap.fromXml(WPS_STYLES)).map((s) => s.styleId);
    expect(ids).not.toContain("9");
  });

  it("`w:hidden` 只影响窗格列不列它，引用它的正文仍按它排版", () => {
    const hidden =
      `<w:styles xmlns:w="x"><w:style w:type="paragraph" w:styleId="h"><w:name w:val="藏起来的"/>` +
      `<w:hidden/><w:rPr><w:sz w:val="28"/></w:rPr></w:style></w:styles>`;
    expect(resolvedStyles(StyleMap.fromXml(hidden))).toEqual([
      { styleId: "h", format: { fontSize: "18.7px" } },
    ]);
  });

  it("w:basedOn 成环时在第一次重复处截断，不会无限循环", () => {
    const cyclic =
      `<w:styles xmlns:w="x">` +
      `<w:style w:type="paragraph" w:styleId="p1"><w:basedOn w:val="p2"/><w:rPr><w:i/></w:rPr></w:style>` +
      `<w:style w:type="paragraph" w:styleId="p2"><w:basedOn w:val="p1"/></w:style>` +
      `</w:styles>`;
    expect(resolvedStyles(StyleMap.fromXml(cyclic))).toEqual([
      { styleId: "p1", format: { italic: true } },
      { styleId: "p2", format: { italic: true } },
    ]);
  });

  it("段落标记的 w:pPr/w:rPr 不算进正文格式", () => {
    // Word 把段落标记的字符属性只给那个段落标记用；读成正文格式会让整段凭空变样。
    const markOnly =
      `<w:styles xmlns:w="x"><w:style w:type="paragraph" w:styleId="m">` +
      `<w:pPr><w:rPr><w:sz w:val="40"/></w:rPr></w:pPr></w:style></w:styles>`;
    expect(resolvedStyles(StyleMap.fromXml(markOnly))).toEqual([]);
  });
});
