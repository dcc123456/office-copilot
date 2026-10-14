/**
 * 「这份 `.pptx` 能不能被办公软件直接打开」的结构审计（`docs/office-parity-spec.md` 阶段 4 的验收面）。
 *
 * 与本机的其它两条审计同构（sheet 侧是 `compat-audit.ts`）：装不了 PowerPoint，Keynote 只能告诉我们
 * "能不能开"、python-pptx 只读值，都不做结构合规判定（`AGENTS.md` §8）。所以把 PowerPoint 打开时会检查的
 * 硬约束写成可机检的判定，产出结构化问题码，文案由调用方组装（§5.7）。
 *
 * 与格式无关的五条（内容类型覆盖、关系可解析、`r:id` 落点、Id 不重复）一律来自 `@office-copilot/opc` 的
 * `auditOpcPackage`，全仓只一份 OPC 判定（§3.2）；本文件只补 `.pptx` 专属的三条：
 * - `<p:sld>` 必须有 `<p:cSld>`，且 `cSld` 里必须有 `<p:spTree>`（Part 1 §19.3.1.16 的 cSld；官方
 *   "Working with presentation slides" 明确 SDK 里 `CommonSlideData` 包 `ShapeTree` 才凑齐"表示一张幻灯片所需的最小元素"）；
 * - 幻灯片部件必须有一条指向版式（slideLayout）的内部关系（同页原话："Each instance of the Slide class must also be
 *   associated with instances of the SlideLayout and SlideMaster classes"；§19.3.1 里没有这条 MUST，它属 PowerPoint 侧约束）。
 *   关系指向的版式部件**是否存在**已经由通用审计那条覆盖，这里只判"有没有这条关系"；母版（slideMaster）由版式自己的关系链到，未单列判定。
 * - 几何元素 `<a:off>`/`<a:ext>`（组合形状那对是 `<a:chOff>`/`<a:chExt>`）的必填属性不许只写一半：Part 1 §20.1.1.15 `CT_Point2D`
 *   与 §20.1.1.14 `CT_PositiveSize2D` 把 x、y 与 cx、cy 都定为 required attribute，python-pptx 读到半个 `<a:off y>` 直接抛 `InvalidXmlError`（本机实测）。
 *
 * 只判"会不会被判损坏"，不判版式好不好看：继承几何、主题解析由 `read.ts` 负责。
 */
import { auditOpcPackage, loadPackage, readText, xml, type OpcAuditCode, type OpcRelationship } from "@office-copilot/opc";

/** 一条 `.pptx` 结构问题：通用码 + 幻灯片专属的三码。 */
export interface PptxAuditIssue {
  /** 机器可判的问题码；文案由调用方组装（`AGENTS.md` §5.7）。 */
  code: OpcAuditCode | "slide-missing-shape-tree" | "slide-missing-layout-relationship" | "slide-incomplete-transform";
  /** 出问题的部件路径（包内路径，如 `ppt/slides/slide1.xml`）。 */
  part: string;
  /** 结构化细节（元素名、关系 Id、目标路径等）。 */
  detail: string;
}

/** 幻灯片部件的路径形状；母版/版式/备注的专属约束暂未纳入本机验收面。 */
const SLIDE_PART = /^ppt\/slides\/slide\d+\.xml$/;

/**
 * 审计一份 `.pptx` 的包完整性与幻灯片必选元素。
 *
 * @param bytes `.pptx` 的原始字节（磁盘上的、或 `PptxFile.toBytes()` 刚写出的）。
 * @returns 问题数组，**空数组代表没发现结构缺陷**；不是 zip 时由 `loadPackage` 抛错，调用方按读文件失败处理。
 */
export async function auditPptxPackage(bytes: Uint8Array): Promise<PptxAuditIssue[]> {
  const pkg = await loadPackage(bytes);
  const issues: PptxAuditIssue[] = [...(await auditOpcPackage(pkg))];

  for (const part of pkg.entryNames.filter((name) => SLIDE_PART.test(name))) {
    const text = await readText(pkg.zip, part);
    if (!text) continue;
    const root = xml.parseDocument(text);
    if (!hasShapeTree(root)) issues.push({ code: "slide-missing-shape-tree", part, detail: "" });
    if (!hasLayoutRelationship(pkg.relsBySourcePart.get(part)))
      issues.push({ code: "slide-missing-layout-relationship", part, detail: "" });
    const incomplete = incompleteTransforms(root);
    if (incomplete.length > 0) issues.push({ code: "slide-incomplete-transform", part, detail: incomplete.join("；") });
  }
  return issues;
}

/** 几何元素各自成对的必填属性（`CT_Point2D` 的 x/y、`CT_PositiveSize2D` 的 cx/cy；组合的那对同形）。 */
const REQUIRED_TRANSFORM_PAIRS: Record<string, [string, string]> = {
  off: ["x", "y"],
  chOff: ["x", "y"],
  ext: ["cx", "cy"],
  chExt: ["cx", "cy"],
};

/**
 * 找出一张幻灯片里"只写了一半"的几何元素。
 *
 * 与写侧那条守卫（`edit.ts` 的 `missingPairValues`）配成一对：那边是不许写出来，这边是万一从别处
 * 进了这种 XML（旧产物、别的工具），落盘前的审计要能看见（§2 红线；python-pptx 读到半个 `<a:off y>` 会抛 `InvalidXmlError`）。
 *
 * @param root 幻灯片部件的根元素。
 * @returns 每处一句"元素名 缺 x、y"；没有时是空数组。
 */
function incompleteTransforms(root: xml.Element): string[] {
  const found: string[] = [];
  for (const [tagName, pair] of Object.entries(REQUIRED_TRANSFORM_PAIRS)) {
    for (const el of xml.collectDescendantsByLocalName(root, tagName)) {
      const missing = pair.filter((attribute) => el.attributes[attribute] === undefined);
      if (missing.length > 0) found.push(`${el.name} 缺 ${missing.join("、")}`);
    }
  }
  return found;
}

/**
 * 判断幻灯片根元素下是否真的存在 `cSld > spTree`（按本地名比对，不依赖前缀写法）。
 *
 * @param root 幻灯片部件解析出的根元素。
 * @returns 两层必选元素都在时 true。
 */
function hasShapeTree(root: xml.Element): boolean {
  const slide = xml.localName(root.name) === "sld" ? root : xml.findChildren(root, "sld")[0];
  if (!slide) return false;
  const commonData = xml.children(slide).find((el) => xml.localName(el.name) === "cSld");
  if (!commonData) return false;
  return xml.children(commonData).some((el) => xml.localName(el.name) === "spTree");
}

/**
 * 找一张幻灯片是否有指向版式部件的内部关系（官方 "Working with presentation slides" 把与 SlideLayout 的关联列为必须）。
 *
 * @param rels 该幻灯片自己的关系表（没有关系表时 undefined）。
 * @returns 有这条内部关系时 true；只有 External 关系或类型对不上时 false。
 */
function hasLayoutRelationship(rels: OpcRelationship[] | undefined): boolean {
  return (rels ?? []).some((rel) => rel.targetMode !== "External" && rel.type.endsWith("/relationships/slideLayout"));
}
