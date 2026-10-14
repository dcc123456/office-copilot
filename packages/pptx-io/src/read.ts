/**
 * `.pptx` 读模型：把 `ppt/presentation.xml` 的幻灯片清单与各张表的形状树读成结构化数据。
 *
 * 单位（`AGENTS.md` §2.1）：页面尺寸、形状偏移/尺寸都是 **EMU**（914400/inch），这里原样读、不换算；
 * 换算只在渲染层做。
 *
 * 溯源红线（spec D3）：占位符（title / body 等）在幻灯片里 `<p:spPr/>` 常是**空的**，它的位置、字体、
 * 颜色全部从版式/母版继承。我们只记录"这格自己有没有写几何（`hasOwnGeometry`）"，绝不把继承来的坐标
 * 烘焙进形状——一烘焙就切断了 PowerPoint 自己的继承链，改母版时不再联动，也违反 §2.3"不许拍平"。
 */
import { children, findChildren, firstChild, loadPackage, parseDocument, readText, resolveTarget, text, type Element, type OpcPackage } from "@office-copilot/opc";
import { PRESENTATION_PART } from "./parts.js";

/** 一个形状（`<p:sp>` / `<p:pic>` / `<p:graphicFrame>` / `<p:grpSp>`）的只读快照。 */
export interface SlideShape {
  /** `<p:cNvPr @id>`，版式内寻址用的形状号。 */
  shapeId: number;
  /** `<p:cNvPr @name>`，PowerPoint 侧边栏显示的名字。 */
  name: string;
  /** 形状大类：形状 / 图片 / 图形框（表格、图表走这个） / 组合。 */
  kind: "shape" | "picture" | "graphicFrame" | "group";
  /** 是不是占位符（`<p:nvPr>` 里有 `<p:ph>`）。 */
  isPlaceholder: boolean;
  /** 占位符类型（`<p:ph @type>`）或索引（`@idx`）；非占位符时 undefined。 */
  placeholderRef?: string;
  /** 形状自己有没有写 `<a:xfrm>`。false 表示几何来自继承，禁止烘焙（见文件头 D3）。 */
  hasOwnGeometry: boolean;
  /** 每段文字（`<a:p>` 内所有 `<a:t>` 拼接）；图片/图形框无 txBody 时为空数组。 */
  paragraphTexts: string[];
  /**
   * 每段自己的 `<a:r>` 条数，与 `paragraphTexts` 一一对应。
   *
   * 视图靠它决定"这段能不能整段改"：一段里有多条 run 说明各截文字带着不同的格式，
   * 整段覆盖会把局部格式拍平（`AGENTS.md` §2.3），那种段落只许按 run 逐条改。
   */
  paragraphRunCounts: number[];
}

/** 一张幻灯片。 */
export interface Slide {
  /** 包内路径，如 `ppt/slides/slide1.xml`。 */
  partPath: string;
  /** 在 `<p:sldIdLst>` 里的序号（0 基），等于放映顺序，不等于文件名编号。 */
  order: number;
  /** `<p:sldId @id>`，撤销/寻址用的稳定 id。 */
  slideId: number;
  /** 形状树顶层子元素（组合不展开）。 */
  shapes: SlideShape[];
}

/** 整份演示文稿的只读模型。 */
export interface PresentationModel {
  /** 页面尺寸，单位 EMU。 */
  slideSizeEmu: { cx: number; cy: number };
  /** 按放映顺序排列的幻灯片。 */
  slides: Slide[];
}

const SHAPE_TAGS: Record<string, SlideShape["kind"]> = {
  "p:sp": "shape",
  "p:pic": "picture",
  "p:graphicFrame": "graphicFrame",
  "p:grpSp": "group",
};

/**
 * 从字节读整份 `.pptx` 的结构模型。
 *
 * @param bytes `.pptx` 的原始字节。
 * @returns 已加载的包 + 幻灯片清单与形状树；不是合法 zip 时由 `loadPackage` 抛错，调用方转成用户可读提示。
 */
export async function readPresentation(bytes: Uint8Array): Promise<{ pkg: OpcPackage; model: PresentationModel }> {
  const pkg = await loadPackage(bytes);
  return { pkg, model: await readModelFromPackage(pkg) };
}

/**
 * 在已加载的包上读结构模型（供 `PptxFile` 复用，避免重复解压）。
 *
 * @param pkg 已 `loadPackage` 的包。
 * @returns `readPresentation` 的 model 部分。
 */
export async function readModelFromPackage(pkg: OpcPackage): Promise<PresentationModel> {
  const presentationXml = await readText(pkg.zip, PRESENTATION_PART);
  if (presentationXml === undefined) throw new Error("演示文稿缺少主部件 ppt/presentation.xml。");
  const root = parseDocument(presentationXml);
  const slideSize = firstChild(root, "p:sldSz");
  const slideSizeEmu = {
    cx: Number(slideSize?.attributes["cx"] ?? 0),
    cy: Number(slideSize?.attributes["cy"] ?? 0),
  };
  const rels = pkg.relsBySourcePart.get(PRESENTATION_PART) ?? [];
  const idList = firstChild(root, "p:sldIdLst");
  const sldIds = idList ? findChildren(idList, "p:sldId") : [];
  const slides: Slide[] = [];
  for (const [order, sldId] of sldIds.entries()) {
    const relId = sldId.attributes["r:id"];
    const rel = rels.find((entry) => entry.id === relId);
    if (!rel) continue; // 关系解析不到：跳过这张，宁可少读也不凭空造一张空幻灯片
    const partPath = resolveTarget(PRESENTATION_PART, rel.target);
    slides.push({
      partPath,
      order,
      slideId: Number(sldId.attributes["id"] ?? 0),
      shapes: await readSlideShapes(pkg, partPath),
    });
  }
  return { slideSizeEmu, slides };
}

/**
 * 读一张幻灯片的形状树顶层。
 *
 * @param pkg 已加载的包。
 * @param partPath 幻灯片部件路径。
 * @returns 顶层形状数组；部件缺失或解析失败返回空数组（读不动不等于崩，交给上层提示）。
 */
async function readSlideShapes(pkg: OpcPackage, partPath: string): Promise<SlideShape[]> {
  const xml = await readText(pkg.zip, partPath);
  if (xml === undefined) return [];
  let slideRoot: Element;
  try {
    slideRoot = parseDocument(xml);
  } catch {
    return [];
  }
  const cSld = firstChild(slideRoot, "p:cSld");
  const spTree = cSld ? firstChild(cSld, "p:spTree") : undefined;
  if (!spTree) return [];
  const shapes: SlideShape[] = [];
  for (const el of children(spTree)) {
    const kind = SHAPE_TAGS[el.name];
    if (!kind) continue; // nvGrpSpPr / grpSpPr 等非形状子元素跳过
    shapes.push(readShape(el, kind));
  }
  return shapes;
}

/**
 * 读单个形状元素。
 *
 * @param el 形状根元素（`<p:sp>` 等）。
 * @param kind 已在 `readSlideShapes` 里判定的大类。
 * @returns 只读快照。
 */
function readShape(el: Element, kind: SlideShape["kind"]): SlideShape {
  const nvPr = firstChild(el, "p:nvSpPr") ?? firstChild(el, "p:nvPicPr") ?? firstChild(el, "p:nvGraphicFramePr") ?? firstChild(el, "p:nvGrpSpPr");
  const cNvPr = nvPr ? firstChild(nvPr, "p:cNvPr") : undefined;
  const phHost = nvPr ? firstChild(nvPr, "p:nvPr") : undefined;
  const ph = phHost ? firstChild(phHost, "p:ph") : undefined;
  const spPr = firstChild(el, "p:spPr");
  const txBody = firstChild(el, "p:txBody");
  const paragraphs = txBody ? findChildren(txBody, "a:p") : [];
  return {
    shapeId: Number(cNvPr?.attributes["id"] ?? 0),
    name: String(cNvPr?.attributes["name"] ?? ""),
    kind,
    isPlaceholder: !!ph,
    placeholderRef: ph ? String(ph.attributes["type"] ?? ph.attributes["idx"] ?? "") : undefined,
    hasOwnGeometry: !!spPr && !!firstChild(spPr, "a:xfrm"),
    paragraphTexts: paragraphs.map((p) => text(p)),
    paragraphRunCounts: paragraphs.map((p) => findChildren(p, "a:r").length),
  };
}
