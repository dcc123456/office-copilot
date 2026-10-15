import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { loadPackage, parseRelationships, relativeTarget, relsPathFor, resolveTarget, unmodelledParts } from "./opc.js";
import { newlyIntroducedIssues, type OpcAuditIssue } from "./audit.js";
import { contentTypesWith, nextPartIndex, nextRelIdFor, relsWith, serializePackage } from "./write.js";
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

describe("write：§2.10 新建部件的关系与内容类型登记原语", () => {
  it("nextRelIdFor 在已有 rId1 的源部件上续到 rId2，不撞号", async () => {
    const pkg = await loadPackage(await minimalPackage());
    expect(nextRelIdFor(pkg, "ppt/presentation.xml")).toBe("rId2");
    // 没有任何关系的源部件从 rId1 起。
    expect(nextRelIdFor(pkg, "ppt/slideMasters/slideMaster1.xml")).toBe("rId1");
  });

  it("relsWith 保留原关系并把目标写成相对源部件目录", async () => {
    const pkg = await loadPackage(await minimalPackage());
    const rels = await relsWith(pkg, "ppt/presentation.xml", [
      { id: "rId2", type: "http://x/slideLayout", targetPartPath: "ppt/slideLayouts/slideLayout1.xml" },
    ]);
    expect(rels).toContain('Id="rId1"');
    expect(rels).toContain('Id="rId2"');
    expect(rels).toContain('Target="slideLayouts/slideLayout1.xml"');
  });

  it("relsWith 一次交两条时两条都在（新建部件自带的关系表不能后一条盖掉前一条）", async () => {
    const pkg = await loadPackage(await minimalPackage());
    // `ppt/notesSlides/` 这一层在原包里根本不存在，走的是"整篇新建"那一档。
    const rels = await relsWith(pkg, "ppt/notesSlides/notesSlide9.xml", [
      { id: "rId1", type: "http://x/notesMaster", targetPartPath: "ppt/notesMasters/notesMaster1.xml" },
      { id: "rId2", type: "http://x/slide", targetPartPath: "ppt/slides/slide1.xml" },
    ]);
    expect(rels).toContain('Id="rId1" Type="http://x/notesMaster" Target="../notesMasters/notesMaster1.xml"');
    expect(rels).toContain('Id="rId2" Type="http://x/slide" Target="../slides/slide1.xml"');
  });

  it("contentTypesWith 追加 Override 且同名不重复", async () => {
    const pkg = await loadPackage(await minimalPackage());
    const once = await contentTypesWith(pkg, [["/ppt/slideMasters/slideMaster1.xml", "ct-master"]]);
    expect(once).toContain('PartName="/ppt/slideMasters/slideMaster1.xml"');
    const twice = await contentTypesWith(pkg, [["/ppt/presentation.xml", "ct-whatever"]]);
    expect(twice.match(/PartName="\/ppt\/presentation\.xml"/g)).toHaveLength(1);
  });

  it("nextPartIndex 数已有的同类部件再 +1（1 基）", () => {
    expect(nextPartIndex(["xl/charts/chart1.xml", "xl/charts/chart2.xml", "xl/styles.xml"], "xl/charts/chart", ".xml")).toBe(3);
    expect(nextPartIndex(["ppt/presentation.xml"], "ppt/slides/slide", ".xml")).toBe(1);
  });
});

/**
 * 保存前那道闸门的判定：只拦"这次改动新引入的"结构问题。
 *
 * 拦下后调用方整批不写盘（`AGENTS.md` §2.4），渲染层把问题拼成一句中文，
 * 所以这里钉的是"什么该拦"，不是文案。`.xlsx` 与 `.pptx` 两条保存路共用这一份判定。
 */
describe("newlyIntroducedIssues：闸门只对本次事务负责", () => {
  /** 造一条问题记录；码取真实存在的通用码，别编。 */
  function issue(code: OpcAuditIssue["code"], part: string, detail = ""): OpcAuditIssue {
    return { code, part, detail };
  }

  it("底包本来就有的问题不拦（第三方导出器留下的，不该因此让人存不了盘）", () => {
    const base = [issue("part-not-declared", "xl/printerSettings.bin", "bin")];
    expect(newlyIntroducedIssues(base, base)).toEqual([]);
    expect(
      newlyIntroducedIssues(base, [...base, issue("unresolved-relationship-id", "xl/worksheets/sheet1.xml", "rId7")]),
    ).toEqual([issue("unresolved-relationship-id", "xl/worksheets/sheet1.xml", "rId7")]);
  });

  it("同一个码落在不同部件上是两条问题，不许互相抵掉", () => {
    const base = [issue("relationship-target-missing", "ppt/slides/slide1.xml", "rId1 -> ppt/slideLayouts/slideLayout9.xml")];
    const out = [...base, issue("relationship-target-missing", "ppt/slides/slide2.xml", "rId1 -> ppt/slideLayouts/slideLayout9.xml")];
    expect(newlyIntroducedIssues(base, out)).toEqual([
      issue("relationship-target-missing", "ppt/slides/slide2.xml", "rId1 -> ppt/slideLayouts/slideLayout9.xml"),
    ]);
  });

  it("输出干净时什么都不拦", () => {
    expect(newlyIntroducedIssues([issue("declared-part-missing", "ppt/media/image1.png")], [])).toEqual([]);
  });
});

describe("opc：序列化不写目录项（.docx/.xlsx/.pptx 三处 toBytes 共用）", () => {
  /** 把序列化产物重新按目录项/文件项拆开列出来，`unzip -l` 的等价判据。 */
  async function entriesOf(bytes: Uint8Array): Promise<string[]> {
    const reloaded = await JSZip.loadAsync(bytes);
    return Object.keys(reloaded.files).sort();
  }

  it("新增嵌套部件后产物只剩文件条目，不多出 ppt/ 与 ppt/slides/", async () => {
    const zip = await JSZip.loadAsync(await minimalPackage());
    // 这一步会让 JSZip 合成 ppt/ 与 ppt/slides/ 两个目录对象（实测 createFolders:false 挡不住）。
    zip.file("ppt/slides/slide2.xml", "<p:sld>乙</p:sld>");

    const bytes = await serializePackage(zip);
    const entries = await entriesOf(bytes);
    expect(entries.filter((name) => name.endsWith("/"))).toEqual([]);
    // 原部件与新增部件都必须在，一个不许丢（delete dir 对象不能把子树删没）。
    expect(entries).toContain("ppt/presentation.xml");
    expect(entries).toContain("ppt/slides/slide1.xml");
    expect(entries).toContain("ppt/slides/slide2.xml");
    const reloaded = await JSZip.loadAsync(bytes);
    expect(await reloaded.file("ppt/slides/slide2.xml")?.async("string")).toBe("<p:sld>乙</p:sld>");
  });

  it("没新增部件时，序列化不改变文件部件清单（只去目录项）", async () => {
    const original = await minimalPackage();
    const bytes = await serializePackage(await JSZip.loadAsync(original));
    const before = (await entriesOf(original)).filter((name) => !name.endsWith("/"));
    expect(await entriesOf(bytes)).toEqual(before);
  });
});
