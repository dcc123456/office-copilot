import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { loadPackage, parseRelationships, relativeTarget, relsPathFor, resolveTarget, unmodelledParts } from "./opc.js";
import * as xml from "./xml.js";

/**
 * 造一份最小 OPC 包字节，只够覆盖"关系表 + 内容类型 + 部件清单"三件事。
 *
 * @returns zip 的 Uint8Array；部件路径与内容都写死在函数里，改夹具时一眼能看见。
 */
async function minimalPackage(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/></Types>');
  zip.file("_rels/.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>');
  zip.file("ppt/_rels/presentation.xml.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>');
  zip.file("ppt/presentation.xml", "<p:presentation/>");
  zip.file("ppt/slides/slide1.xml", "<p:sld/>");
  zip.file("docProps/thumbnail.jpeg", "");
  return zip.generateAsync({ type: "uint8array" });
}

describe("opc：包与关系（格式无关的那一层）", () => {
  it("relativeTarget 与 resolveTarget 互为反向，跨目录也成立", () => {
    const cases: Array<[string, string]> = [
      ["xl/worksheets/sheet1.xml", "xl/drawings/drawing1.xml"],
      ["ppt/slides/slide1.xml", "ppt/slides/_rels/slide1.xml.rels"],
      ["ppt/presentation.xml", "ppt/slideMasters/slideMaster1.xml"],
      ["xl/workbook.xml", "xl/sharedStrings.xml"],
    ];
    for (const [source, destination] of cases) {
      const relative = relativeTarget(source, destination);
      expect(resolveTarget(source, relative), `${source} → ${destination} 走一圈 ${relative}`).toBe(destination);
    }
  });

  it("绝对写法与带 .. 的写法都收敛到同一个包内路径", () => {
    expect(resolveTarget("xl/workbook.xml", "/docProps/core.xml")).toBe("docProps/core.xml");
    expect(resolveTarget("xl/worksheets/sheet1.xml", "../sharedStrings.xml")).toBe("xl/sharedStrings.xml");
    expect(resolveTarget("xl/workbook.xml", "./styles.xml")).toBe("xl/styles.xml");
  });

  it("关系文件路径永远在源部件自己的 _rels 目录里", () => {
    expect(relsPathFor("xl/worksheets/sheet1.xml")).toBe("xl/worksheets/_rels/sheet1.xml.rels");
    expect(relsPathFor("[Content_Types].xml")).toBe("_rels/[Content_Types].xml.rels");
  });

  it("空的关系文本回空数组而不是抛错", () => {
    expect(parseRelationships("")).toEqual([]);
    expect(parseRelationships("   \n  ")).toEqual([]);
  });

  it("loadPackage 按源部件归好关系表，内容类型分扩展名与 Override 两张表", async () => {
    const pkg = await loadPackage(await minimalPackage());
    expect(pkg.entryNames).toContain("ppt/slides/slide1.xml");
    expect((pkg.relsBySourcePart.get("ppt/presentation.xml") ?? []).map((rel) => rel.target)).toEqual(["slides/slide1.xml"]);
    expect(pkg.contentTypes.get("rels")).toBe("application/vnd.openxmlformats-package.relationships+xml");
    expect(pkg.overrideParts.get("ppt/presentation.xml")).toContain("presentationml.presentation.main+xml");
  });

  it("unmodelledParts 不把关系文件与目录项算成内容部件", async () => {
    const pkg = await loadPackage(await minimalPackage());
    const rest = unmodelledParts(pkg, new Set(["ppt/presentation.xml", "ppt/slides/slide1.xml", "docProps/thumbnail.jpeg"]));
    expect(rest).toEqual(["[Content_Types].xml"]);
  });
});

describe("xml：通用序列化（往返不许丢读不懂的东西）", () => {
  it("读不懂的属性与命名空间前缀原样带回", () => {
    const source = '<p:sld xmlns:a="http://x" xmlns:p="http://y" p:unknown="7"><a:t>甲</a:t><![CDATA[原样]]><!--注--></p:sld>';
    const root = xml.parseDocument(source);
    expect(xml.serialize(root)).toBe(source);
    expect(xml.text(root)).toBe("甲");
    expect(xml.serialize(xml.children(root)[0] ?? "")).toBe("<a:t>甲</a:t>");
  });

  it("innerXml 只取子元素，纯文本与自闭合都回空串", () => {
    const root = xml.parseDocument('<a><b x="1"/><c><d>乙</d></c><e>仅文本</e></a>');
    const [b, c, e] = xml.children(root);
    expect(xml.innerXml(b!)).toBe("");
    expect(xml.innerXml(c!)).toBe("<d>乙</d>");
    expect(xml.innerXml(e!)).toBe("");
  });
});
