import { describe, it, expect } from "vitest";
import { StyleMap, fitStyleId, headingLookupFromPaneStyles } from "./styles.js";

/** Word 自己写出的 styles.xml 是紧凑一行的，夹具照这个形态来。 */
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

/** 只有 w:style 子元素的一份 styles.xml，用于整文件重建比对。 */
const ONLY_STYLES =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<w:styles ${W}>` +
  `<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:after="120" w:line="360" w:lineRule="auto"/></w:pPr></w:style>` +
  `<w:style w:type="character" w:default="1" w:styleId="a0"><w:name w:val="Default Paragraph Font"/><w:semiHidden/></w:style>` +
  `<w:style w:type="paragraph" w:styleId="1"><w:name w:val="heading 1"/><w:basedOn w:val="a"/><w:next w:val="a"/><w:link w:val="10"/><w:uiPriority w:val="9"/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="240"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="44"/></w:rPr></w:style>` +
  `<w:style w:type="character" w:styleId="10"><w:name w:val="标题 1 字符"/><w:basedOn w:val="a0"/><w:uiPriority w:val="9"/><w:semiHidden/><w:unhideWhenUsed/><w:rPr><w:b/><w:sz w:val="44"/></w:rPr></w:style>` +
  `<w:style w:type="table" w:default="1" w:styleId="a1"><w:name w:val="Normal Table"/><w:tblPr><w:tblCellMar><w:left w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>` +
  `</w:styles>`;

/** 一份带 docDefaults / latentStyles 的 WPS 风格文档，未建模的部分必须原样留着。 */
const WITH_NOISE =
  `<w:styles ${W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman"/><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults>` +
  `<w:latentStyles w:defLockedState="0" w:uiPriority="99"/><w:associatedStyles w:styleId="af"/><w:style w:type="paragraph" w:styleId="af2"><w:name w:val="列表段落"/><w:pPr><w:ind w:left="420" w:leftChars="200"/></w:pPr></w:style></w:styles>`;

describe("StyleMap：具名样式清单（排名 2 的读侧）", () => {
  const map = StyleMap.fromXml(ONLY_STYLES);

  it("按文档顺序列出每一条样式，带着名称与类型", () => {
    expect(map.definitions().map((d) => [d.styleId, d.name, d.type])).toEqual([
      ["a", "Normal", "paragraph"],
      ["a0", "Default Paragraph Font", "character"],
      ["1", "heading 1", "paragraph"],
      ["10", "标题 1 字符", "character"],
      ["a1", "Normal Table", "table"],
    ]);
  });

  it("读样式链字段：basedOn / next / link / uiPriority / default / semiHidden", () => {
    const h1 = map.byId("1")!;
    expect(h1).toMatchObject({
      basedOn: "a",
      next: "a",
      link: "10",
      uiPriority: 9,
      isDefault: false,
      semiHidden: false,
      hidden: false,
    });
    // 自闭合的 w:semiHidden 是 CT_OnOff 的"裸元素即为真"写法。
    expect(map.byId("a0")!.semiHidden).toBe(true);
    expect(map.byId("a")!.isDefault).toBe(true);
  });

  it("保留样式自己的 pPr/rPr 原文，一个字符单位都不换算", () => {
    expect(map.byId("1")!.pPrInner).toBe(
      `<w:keepNext/><w:keepLines/><w:spacing w:before="240"/><w:outlineLvl w:val="0"/>`,
    );
    expect(map.byId("1")!.rPrInner).toBe(`<w:b/><w:sz w:val="44"/>`);
    // 没写 pPr 的样式就是 undefined，不能凭空造一个空的。
    expect(map.byId("a0")!.pPrInner).toBeUndefined();
  });

  it("没改动过的样式按原样写回：重建出的文件与读进来的一字不差", () => {
    const rebuilt =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<w:styles ${W}>` +
      map.definitions().map((d) => d.xml).join("") +
      `</w:styles>`;
    expect(rebuilt).toBe(ONLY_STYLES);
  });

  it("沿 basedOn 上溯到父样式，父样式缺失就停在链尾", () => {
    expect(map.chain("1").map((d) => d.styleId)).toEqual(["1", "a"]);
    // Normal 没有 basedOn，链到自己为止。
    expect(map.chain("a").map((d) => d.styleId)).toEqual(["a"]);
    expect(map.chain("未声明")).toEqual([]);
    expect(map.chain(undefined)).toEqual([]);
  });

  it("认得内建标题名与默认段落的段后距，这两件事不能因为结构化而退化", () => {
    expect(map.headingLevel("1")).toBe(1);
    // 半磅/磅都不在这里换算：spacing 的 w:after=120 twips → 8px。
    expect(map.defaultSpacingAfterPx).toBe(8);
    expect(map.isDeclared("a1")).toBe(true);
  });
});

describe("StyleMap：不认识的部件不抢写", () => {
  const map = StyleMap.fromXml(WITH_NOISE);

  it("docDefaults 与 latentStyles 不进结构化视图，原文照旧", () => {
    expect(map.definitions().map((d) => d.styleId)).toEqual(["af2"]);
    expect(map.byId("af2")!.pPrInner).toBe(`<w:ind w:left="420" w:leftChars="200"/>`);
    // 写侧要靠"未改动的样式原样 splice 回去"，所以每条定义的 xml 必须能在原文里找到。
    expect(WITH_NOISE).toContain(map.byId("af2")!.xml);
  });

  it("w:basedOn 成环时在第一次重复处截断，不会转不完", () => {
    const cyclic = StyleMap.fromXml(
      `<w:styles ${W}>` +
        `<w:style w:type="paragraph" w:styleId="p1"><w:name w:val="一"/><w:basedOn w:val="p2"/></w:style>` +
        `<w:style w:type="paragraph" w:styleId="p2"><w:name w:val="二"/><w:basedOn w:val="p1"/></w:style>` +
        `</w:styles>`,
    );
    expect(cyclic.chain("p1").map((d) => d.styleId)).toEqual(["p1", "p2"]);
  });

  it("读不懂的 uiPriority 与缺 @w:type 都不算失败", () => {
    const odd = StyleMap.fromXml(
      `<w:styles ${W}>` +
        `<w:style w:styleId="x"><w:name w:val="没有类型"/><w:uiPriority w:val="abc"/></w:style>` +
        `<w:style w:type="paragraph" w:default="0" w:styleId="y"><w:name w:val="显式不为默认"/></w:style>` +
        `</w:styles>`,
    );
    expect(odd.byId("x")!.type).toBe("paragraph");
    expect(odd.byId("x")!.uiPriority).toBeUndefined();
    expect(odd.byId("y")!.isDefault).toBe(false);
  });

  it("没有 styles.xml 时是一张空表，不抛错", () => {
    expect(StyleMap.fromXml(undefined).definitions()).toEqual([]);
    expect(StyleMap.fromXml("<w:styles />").definitions()).toEqual([]);
  });
});

describe("StyleMap.paneStyles：样式窗格的数据源（排名 2 的读侧终点）", () => {
  const map = StyleMap.fromXml(ONLY_STYLES);

  it("只列段落样式：写了 uiPriority 的在前，没写的按声明次序垫底", () => {
    expect(map.paneStyles()).toEqual([
      { styleId: "1", name: "heading 1", headingLevel: 1, isDefault: false },
      // Normal 没有 uiPriority 所以垫底；字符样式 a0、表格样式 a1 根本不进窗格。
      { styleId: "a", name: "Normal", isDefault: true },
    ]);
  });

  it("w:hidden / w:semiHidden 的段落样式不列出，但样式表照旧认得它们", () => {
    const hidden = StyleMap.fromXml(
      `<w:styles ${W}>` +
        `<w:style w:type="paragraph" w:styleId="p1"><w:name w:val="普通段落"/></w:style>` +
        `<w:style w:type="paragraph" w:styleId="p2"><w:name w:val="隐藏段落"/><w:hidden/></w:style>` +
        `<w:style w:type="paragraph" w:styleId="p3"><w:name w:val="半隐藏段落"/><w:semiHidden/></w:style>` +
        `</w:styles>`,
    );
    expect(hidden.paneStyles().map((s) => s.styleId)).toEqual(["p1"]);
    // 窗格看不见不等于样式不存在：段落里已有的引用写回时仍要判定为已声明。
    expect(hidden.isDeclared("p2")).toBe(true);
  });

  it("没有 styles.xml 时窗格是空列表，不抛错", () => {
    expect(StyleMap.fromXml(undefined).paneStyles()).toEqual([]);
  });

  it("清单能拼出 fitStyleId 的查询面，窗格高亮因此与写侧同一条规则", () => {
    const lookup = headingLookupFromPaneStyles(map.paneStyles());
    expect(lookup.headingLevel("1")).toBe(1);
    expect(lookup.styleIdForHeadingLevel(1)).toBe("1");
    // 清单里没有的级别查不到 id，交给调用方按节点类型推断。
    expect(lookup.styleIdForHeadingLevel(3)).toBeUndefined();
    // 段落套着一级标题、级别已经改成 2：窗格显示与导出都该指向二级那条。
    expect(fitStyleId("1", 2, lookup)).toBeUndefined();
    expect(fitStyleId("1", 1, lookup)).toBe("1");
  });
});
