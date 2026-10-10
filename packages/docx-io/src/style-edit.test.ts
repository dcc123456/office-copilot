import { describe, it, expect } from "vitest";
import { mergeStylesXml, nextStyleId, type StyleEdit } from "./style-edit.js";
import { stylePropertiesXml } from "./ooxml-pm.js";
import { resolvedStyles } from "./style-format.js";
import type { PMMark } from "./pm-schema.js";
import { StyleMap } from "./styles.js";

/** 一份带 docDefaults、latentStyles 与两条已声明样式的 styles.xml（WPS 的形状：数字 id）。 */
const ORIGINAL = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><w:docDefaults><w:pPrDefault><w:pPr><w:spacing w:after="120"/></w:pPr></w:pPrDefault></w:docDefaults><w:latentStyles w:defLockedState="0" w:latentStyleCount="376"/><w:style w:type="paragraph" w:default="1" w:styleId="1"><w:name w:val="Normal"/><w:qFormat/></w:style><w:style w:type="paragraph" w:styleId="2"><w:name w:val="heading 1"/><w:basedOn w:val="1"/><w:uiPriority w:val="1"/><w:pPr><w:spacing w:before="240" w:beforeLines="50"/><w:ind w:left="0" w:leftChars="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="48"/><w:szCs w:val="48"/></w:rPr></w:style><w:style w:type="character" w:default="1" w:styleId="8"><w:name w:val="Default Paragraph Font"/></w:style></w:styles>`;

const BOLD_RED: PMMark[] = [
  { type: "bold" },
  { type: "textStyle", attrs: { fontSize: "24px", color: "#c00000" } },
];

/** 一条"从选区抓到的格式"：左缩进 36pt、段前 12pt、居中、加粗红字 18pt。 */
function edit(patch: Partial<StyleEdit> = {}): StyleEdit {
  return {
    name: "重点段落",
    paragraph: { indent: 36, textAlign: "center", spacingBefore: 12 },
    marks: BOLD_RED,
    ...patch,
  };
}

/** 从合并结果里抠出某条样式的 XML；找不到返回空串。 */
function styleOf(xml: string | undefined, styleId: string): string {
  const match = new RegExp(
    `<w:style\\b[^>]*w:styleId="${styleId}"[^>]*>[\\s\\S]*?</w:style>`,
  ).exec(xml ?? "");
  return match?.[0] ?? "";
}

describe("nextStyleId", () => {
  it("把样式名压成 Word 那种 ASCII id，中文名单退回前缀", () => {
    expect(nextStyleId(new Set(), "Overdue Amount Para")).toBe("OverdueAmountPara");
    expect(nextStyleId(new Set(), "重点段落")).toBe("Style");
    expect(nextStyleId(new Set(["Style"]), "重点段落")).toBe("Style1");
    expect(nextStyleId(new Set(["Existing"]), "Existing")).toBe("Existing1");
  });
});

describe("stylePropertiesXml", () => {
  it("段落 attrs 按 CT_PPr 顺序换算成 twips，marks 换算成半磅与 hex", () => {
    const { pPrInner, rPrInner } = stylePropertiesXml({
      paragraph: { indent: 36, textAlign: "center", spacingBefore: 12 },
      marks: BOLD_RED,
    });
    expect(pPrInner).toBe('<w:spacing w:before="240"/><w:ind w:left="720"/><w:jc w:val="center"/>');
    expect(rPrInner).toBe('<w:b/><w:color w:val="C00000"/><w:sz w:val="36"/><w:szCs w:val="36"/>');
  });

  it("与原定义合并：建模过的顶替，没建模过的留着", () => {
    const { pPrInner, rPrInner } = stylePropertiesXml(
      { paragraph: { spacingBefore: 6 }, marks: [{ type: "textStyle", attrs: { fontSize: "12px" } }] },
      {
        pPrInner:
          '<w:spacing w:before="240" w:beforeLines="50"/><w:tabs><w:tab w:val="left" w:pos="720"/></w:tabs>',
        rPrInner: '<w:b/><w:sz w:val="48"/><w:szCs w:val="48"/><w:lang w:val="zh-CN"/>',
      },
    );
    // before 被这次抓取接管（120 ≠ 原来的 240），所以会压过它的 w:beforeLines 必须清掉；
    // w:tabs 没建模过，原样带走，且按 CT_PPr 的顺序排在 spacing 之前。
    expect(pPrInner).toBe(
      '<w:tabs><w:tab w:val="left" w:pos="720"/></w:tabs><w:spacing w:before="120"/>',
    );
    // 旧字号连着只镜像西文的 szCs 一起换掉，没建模的 w:lang 留在后面。
    expect(rPrInner).toBe('<w:b/><w:sz w:val="18"/><w:szCs w:val="18"/><w:lang w:val="zh-CN"/>');
  });

  it("抓来的制表位顶替样式原有的那条，顺序仍按 CT_PPr", () => {
    // 样式定义与段落共用同一批生成器，所以 w:tabs 的换算这里只有一份（§2.1）。
    const { pPrInner } = stylePropertiesXml(
      {
        paragraph: {
          tabStops: [{ positionPt: 415.3, alignment: "right", leader: "dot" }],
          spacingBefore: 6,
        },
        marks: [],
      },
      { pPrInner: '<w:tabs><w:tab w:val="left" w:pos="720"/></w:tabs>' },
    );
    expect(pPrInner).toBe(
      '<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="8306"/></w:tabs><w:spacing w:before="120"/>',
    );
  });

  it("选区是「全部清除」时样式里那条 w:tabs 一起删掉，也不留空的 w:tabs", () => {
    const { pPrInner } = stylePropertiesXml(
      { paragraph: { tabStops: [] }, marks: [] },
      { pPrInner: '<w:tabs><w:tab w:val="left" w:pos="720"/></w:tabs><w:jc w:val="center"/>' },
    );
    expect(pPrInner).toBe('<w:jc w:val="center"/>');
  });

  it("attrs 里的 styleId / rawPpr 不会被写进样式定义", () => {
    const { pPrInner } = stylePropertiesXml({
      paragraph: { styleId: "2", rawPpr: '<w:numPr><w:numId w:val="5"/></w:numPr>' },
      marks: [],
    });
    expect(pPrInner).toBe("");
  });

  /** 编号挂在样式上时也是用户的格式：抓别的属性不能把它顺手删了（§2.3 同一红线）。 */
  it("被改样式原本的 numPr 原样带走", () => {
    const { pPrInner } = stylePropertiesXml(
      { paragraph: { indent: 36 }, marks: [] },
      { pPrInner: '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr>' },
    );
    expect(pPrInner).toBe(
      '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr><w:ind w:left="720"/>',
    );
  });
});

describe("mergeStylesXml", () => {
  it("新建：追加在末尾，带 customStyle，name 在前属性块在后", () => {
    const xml = mergeStylesXml(ORIGINAL, [edit({ styleId: undefined, name: "Overdue Para" })]);
    const created = styleOf(xml, "OverduePara");
    expect(created).toContain('<w:style w:type="paragraph" w:customStyle="1" w:styleId="OverduePara">');
    expect(created.indexOf('<w:name w:val="Overdue Para"/>')).toBeLessThan(
      created.indexOf("<w:pPr>"),
    );
    expect(xml?.lastIndexOf("OverduePara")).toBeGreaterThan(xml?.indexOf('w:styleId="8"') ?? -1);
  });

  it("splice-and-keep：docDefaults、latentStyles 与没点过的样式一条不少", () => {
    const xml = mergeStylesXml(ORIGINAL, [edit({ styleId: "2", name: "标题 1" })]);
    expect(xml).toContain("<w:docDefaults>");
    expect(xml).toContain('<w:latentStyles w:defLockedState="0"');
    expect(styleOf(xml, "1")).toBe(
      '<w:style w:type="paragraph" w:default="1" w:styleId="1"><w:name w:val="Normal"/><w:qFormat/></w:style>',
    );
    expect(styleOf(xml, "8")).toContain('<w:name w:val="Default Paragraph Font"/>');
  });

  it("改一条：只换 name 与属性块，basedOn / uiPriority 这类不动", () => {
    // 段前取 6pt（=120 缇）与原来的 240 不同，这次抓取算接管了这个属性。
    const xml = mergeStylesXml(ORIGINAL, [
      edit({ styleId: "2", name: "标题 1", paragraph: { indent: 36, textAlign: "center", spacingBefore: 6 } }),
    ]);
    const heading = styleOf(xml, "2");
    expect(heading).toContain('<w:name w:val="标题 1"/>');
    expect(heading).toContain('<w:basedOn w:val="1"/>');
    expect(heading).toContain('<w:uiPriority w:val="1"/>');
    expect(heading).toContain('<w:spacing w:before="120"/>');
    expect(heading).toContain('<w:ind w:left="720"/>');
    expect(heading).toContain('<w:jc w:val="center"/>');
    expect(heading).not.toContain('<w:sz w:val="48"/>');
    // 这次抓取接管了 before 与 left，所以它们的同义写法必须一起清掉，
    // 否则 Word 仍按字符/行单位排，用户的修改等于没生效。
    expect(heading).not.toContain("w:beforeLines");
    expect(heading).not.toContain("w:leftChars");
  });

  /** 取值没变的属性不算被接管，同义写法原样留着——重复保存不侵蚀用户的"按行排版"意图。 */
  it("段前取回同一个值时不清同义写法", () => {
    const xml = mergeStylesXml(ORIGINAL, [edit({ styleId: "2", name: "标题 1" })]);
    expect(styleOf(xml, "2")).toContain('w:beforeLines="50"');
  });

  it("同一轮改两次仍是同一条样式，第二次的合并不丢第一次写进去的加粗", () => {
    const xml = mergeStylesXml(ORIGINAL, [
      edit({ styleId: "2", name: "标题 1" }),
      edit({ styleId: "2", name: "标题 1（改）", paragraph: { indent: 18 } }),
    ]);
    expect(xml?.match(/w:styleId="2"/g)).toHaveLength(1);
    const heading = styleOf(xml, "2");
    expect(heading).toContain('<w:ind w:left="360"/>');
    expect(heading).toContain("<w:b/>");
    expect(heading).toContain('<w:color w:val="C00000"/>');
  });

  it("名字为空的编辑整条忽略；没有有效编辑时原文件一个字节都不动", () => {
    expect(mergeStylesXml(ORIGINAL, [edit({ name: "   " })])).toBe(ORIGINAL);
    expect(mergeStylesXml(ORIGINAL, [])).toBe(ORIGINAL);
    expect(mergeStylesXml(undefined, [])).toBeUndefined();
  });

  it("这份文档还没有样式表时自建根元素，内容照常写出", () => {
    const xml = mergeStylesXml(undefined, [edit({ styleId: undefined, name: "My Note" })]);
    expect(xml).toContain('<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">');
    expect(styleOf(xml, "MyNote")).toContain('<w:spacing w:before="240"/>');
  });

  it("原文件读不懂时按没有样式表处理，不把编辑静默丢掉", () => {
    const xml = mergeStylesXml("<w:styles><oops", [edit({ styleId: undefined, name: "My Note" })]);
    expect(styleOf(xml, "MyNote")).not.toBe("");
  });

  it("样式名里的引号与 & 转义，不会提前结束属性值", () => {
    const xml = mergeStylesXml(ORIGINAL, [edit({ styleId: undefined, name: `A"&<B` })]);
    expect(xml).toContain('<w:name w:val="A&quot;&amp;&lt;B"/>');
    expect(/<w:name w:val="A"&/.test(xml ?? "")).toBe(false);
  });

  it("重复跑同一次合并不改字节（保存两次不会把格式一步步磨掉）", () => {
    const edits: StyleEdit[] = [edit({ styleId: "2", name: "标题 1" })];
    const once = mergeStylesXml(ORIGINAL, edits);
    expect(mergeStylesXml(once, edits)).toBe(once);
  });

  it("合并完的 styles.xml 读得回来：新样式进窗格，格式也解析得出", () => {
    const xml = mergeStylesXml(ORIGINAL, [
      edit({ styleId: undefined, name: "重点段落", basedOn: "1" }),
    ]);
    const map = StyleMap.fromXml(xml);
    const pane = map.paneStyles();
    expect(pane.map((s) => s.name)).toContain("重点段落");
    expect(pane.map((s) => s.name)).toContain("heading 1");
    const formats = resolvedStyles(map);
    const created = formats.find((entry) => entry.styleId === "Style");
    // 单位是编辑器那套词汇：缩进 pt，字号 CSS px，颜色 hex（AGENTS §2.1）。
    expect(created?.format).toMatchObject({
      indent: 36,
      spacingBefore: 12,
      textAlign: "center",
      bold: true,
      color: "#c00000",
      fontSize: "24px",
      // basedOn 的那条 + docDefaults 的段后距一起继承下来。
      spacingAfter: 6,
    });
  });

  it("引用了没声明的父样式时不写 basedOn（Word 对指向空气的引用按默认样式排）", () => {
    const xml = mergeStylesXml(ORIGINAL, [
      edit({ styleId: undefined, name: "Loose", basedOn: "Ghost" }),
    ]);
    expect(xml).not.toContain('<w:basedOn w:val="Ghost"/>');
  });
});
