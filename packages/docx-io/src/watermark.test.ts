import { describe, expect, it } from "vitest";
import {
  applyWatermark,
  readWatermark,
  sameWatermark,
  type SectionWatermark,
} from "./watermark.js";

const HDR_OPEN =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
  `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">`;

/** Word 真的写出来的那一份 `_x0000_t136` 定义块，测试里当"外部产物"用。 */
const SHAPE_TYPE =
  `<v:shapetype id="_x0000_t136" coordsize="21600,21600" o:spt="136" adj="10800" path="m@7,l@8,m@5,21600l@6,21600e">` +
  `<v:formulas><v:f eqn="sum #0 0 10800"/></v:formulas>` +
  `<v:path textpathok="t" o:connecttype="custom" o:connectlocs="@9,0" o:connectangles="270"/>` +
  `<v:textpath on="t" fitshape="t"/>` +
  `<o:lock v:ext="edit" text="t" shapetype="t"/>` +
  `</v:shapetype>`;

/**
 * 造一份带水印的页眉 XML，形状按 Word 的写法来（属性顺序、定点透明度都照抄）。
 *
 * @param inner `w:pict` 里的内容；默认给一个斜排半透明的灰色水印。
 * @returns 完整的 `<w:hdr>` 文档，前面已声明 VML 命名空间。
 */
function headerWith(inner: string): string {
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
    `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ` +
    `xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" ` +
    `xmlns:w10="urn:schemas-microsoft-com:office:word">` +
    `<w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr><w:r><w:rPr><w:noProof/></w:rPr><w:pict>` +
    inner +
    `</w:pict></w:r></w:p></w:hdr>`
  );
}

const WORD_WATERMARK =
  SHAPE_TYPE +
  `<v:shape id="PowerPoint 水印" o:spid="_x0000_s1026" type="#_x0000_t136" ` +
  `style="position:absolute;margin-left:100pt;margin-top:120pt;width:468pt;height:117pt;` +
  `z-index:-251654144;mso-wrap-edited:f" rotation="315" fillcolor="#c0c0c0" stroked="f">` +
  `<v:fill on="t" focalscale="none" opacity="50000f"/><v:stroke on="f"/>` +
  `<v:textpath style="font-family:'SimHei';font-size:1pt" fitshape="t" string="机密★内部资料"/>` +
  `<o:lock v:ext="edit" shapetype="t"/></v:shape>`;

describe("readWatermark", () => {
  it("reads a Word-written diagonal semi-transparent watermark", () => {
    expect(readWatermark(headerWith(WORD_WATERMARK))).toEqual({
      text: "机密★内部资料",
      font: "SimHei",
      sizePt: null,
      color: "#c0c0c0",
      semiTransparent: true,
      layout: "diagonal",
    });
  });

  it("reads a fixed size and a horizontal layout", () => {
    const xml = headerWith(
      `<v:shape type="#_x0000_t136" fillcolor="#FF0000">` +
        `<v:textpath style="font-family:'宋体';font-size:36pt" string="草稿"/>` +
        `</v:shape>`,
    );
    expect(readWatermark(xml)).toEqual({
      text: "草稿",
      font: "宋体",
      sizePt: 36,
      color: "#ff0000",
      semiTransparent: false,
      layout: "horizontal",
    });
  });

  it("decodes entities in the watermark text", () => {
    const xml = headerWith(
      `<v:shape type="#_x0000_t136" fillcolor="#000000">` +
        `<v:textpath style="font-family:'宋体';font-size:20pt" string="A&amp;B&lt;C&quot;"/>` +
        `</v:shape>`,
    );
    expect(readWatermark(xml)?.text).toBe('A&B<C"');
  });

  it("returns undefined when there is no watermark at all", () => {
    expect(readWatermark(undefined)).toBeUndefined();
    expect(readWatermark(`${HDR_OPEN}<w:p><w:r><w:t>页眉文字</w:t></w:r></w:p></w:hdr>`)).toBeUndefined();
  });

  it("leaves decorative WordArt and unreadable colors unmodeled", () => {
    // 没有 @string 的艺术字不是水印；颜色名读不懂时宁可不动，也不把看不懂的值写回去（§2.2）。
    const noString = headerWith(
      SHAPE_TYPE + `<v:shape type="#_x0000_t136" fillcolor="#c0c0c0"><v:textpath on="t"/></v:shape>`,
    );
    expect(readWatermark(noString)).toBeUndefined();
    const namedColor = headerWith(
      `<v:shape type="#_x0000_t136" fillcolor="gray">` +
        `<v:textpath style="font-family:'宋体';font-size:20pt" string="水印"/></v:shape>`,
    );
    expect(readWatermark(namedColor)).toBeUndefined();
  });
});

describe("applyWatermark", () => {
  const wm: SectionWatermark = {
    text: "内部资料",
    font: "黑体",
    sizePt: null,
    color: "#c00000",
    semiTransparent: true,
    layout: "diagonal",
  };

  it("writes the shape into a plain header and keeps its own text", () => {
    const plain = `${HDR_OPEN}<w:p><w:r><w:t>页眉：合同编号</w:t></w:r></w:p></w:hdr>`;
    const out = applyWatermark(plain, wm);
    expect(out).toContain("页眉：合同编号");
    // 页眉根元素原本只声明了 w:/r:，形状用的前缀必须补上，否则是非法 XML。
    expect(out).toContain('xmlns:v="urn:schemas-microsoft-com:vml"');
    expect(out).toContain('string="内部资料"');
    expect(out).toContain('fillcolor="#c00000"');
    expect(out).toContain('rotation="315"');
    expect(out).toContain('opacity="50%"');
    expect(out).toContain(`font-family:'黑体'`);
    expect(out).toContain('fitshape="t"');
  });

  it("round-trips through readWatermark", () => {
    const plain = `${HDR_OPEN}<w:p><w:r><w:t>页眉</w:t></w:r></w:p></w:hdr>`;
    expect(readWatermark(applyWatermark(plain, wm))).toEqual(wm);
    expect(readWatermark(applyWatermark(plain, { ...wm, sizePt: 48, layout: "horizontal", semiTransparent: false }))).toEqual({
      ...wm,
      sizePt: 48,
      layout: "horizontal",
      semiTransparent: false,
    });
  });

  it("replaces an existing watermark instead of stacking a second shape", () => {
    const out = applyWatermark(headerWith(WORD_WATERMARK), wm);
    expect(out.match(/<v:shape\b/g)).toHaveLength(1);
    // 定义块已经在了，再写一遍就是重复 id。
    expect(out.match(/<v:shapetype\b/g)).toHaveLength(1);
    expect(out).toContain('string="内部资料"');
    expect(out).not.toContain("机密★内部资料");
    expect(readWatermark(out)).toEqual(wm);
  });

  it("strips the watermark paragraph but keeps everything else", () => {
    const withText =
      headerWith(WORD_WATERMARK).replace(
        "</w:hdr>",
        `<w:p><w:r><w:t>页脚上方的标题</w:t></w:r></w:p></w:hdr>`,
      );
    const out = applyWatermark(withText, null);
    expect(out).not.toContain("v:shape");
    expect(out).toContain("页脚上方的标题");
    expect(readWatermark(out)).toBeUndefined();
  });

  it("never removes a decorative WordArt living inside a text box", () => {
    // 文本框里的艺术字也带 t136 + string，但它不是水印；删了就是改用户的版式（§2.7）。
    const decorative =
      `${HDR_OPEN}<w:p><w:r><w:pict>` +
      `<v:shape type="#_x0000_t136" fillcolor="#c0c0c0"><v:textbox><w:txbxContent>` +
      `<w:p><w:r><w:t>框里</w:t></w:r></w:p>` +
      `</w:txbxContent></v:textbox>` +
      `<v:textpath style="font-family:'宋体';font-size:20pt" string="艺术字"/></v:shape>` +
      `</w:pict></w:r></w:p></w:hdr>`;
    expect(applyWatermark(decorative, null)).toBe(decorative);
  });

  it("escapes the text so a quote cannot break out of the attribute", () => {
    const out = applyWatermark(`${HDR_OPEN}</w:hdr>`, { ...wm, text: '"><w:t>注入' });
    expect(out).not.toContain('string=""><w:t>注入');
    expect(readWatermark(out)?.text).toBe('"><w:t>注入');
  });
});

describe("sameWatermark", () => {
  const base: SectionWatermark = {
    text: "水印",
    font: "宋体",
    sizePt: null,
    color: "#c0c0c0",
    semiTransparent: true,
    layout: "diagonal",
  };

  it("treats two undefined as equal and any missing side as different", () => {
    expect(sameWatermark(undefined, undefined)).toBe(true);
    expect(sameWatermark(base, undefined)).toBe(false);
    expect(sameWatermark(undefined, base)).toBe(false);
  });

  it("compares every field", () => {
    expect(sameWatermark(base, { ...base })).toBe(true);
    for (const patch of [
      { text: "别的" },
      { font: "黑体" },
      { sizePt: 20 },
      { color: "#000000" },
      { semiTransparent: false },
      { layout: "horizontal" as const },
    ]) {
      expect(sameWatermark(base, { ...base, ...patch })).toBe(false);
    }
  });
});
