/**
 * `.pptx` 读模型：把 `ppt/presentation.xml` 的幻灯片清单与各张表的形状树读成结构化数据。
 *
 * 单位（`AGENTS.md` §2.1 / spec D4）：换算发生在**这里**（IO 边），交给 UI 的一律是 CSS px 与 CSS pt；
 * 只有页面尺寸 `slideSizeEmu` 保留 EMU（名字里就写着单位）。形状自己的 `<a:xfrm>` 走 `units.ts` 换成 px，
 * run 的 `sz` 换成 pt——渲染层再乘一次 100 或 9525 就是第二处换算，属缺陷。
 *
 * 溯源红线（spec D3）：占位符（title / body 等）在幻灯片里 `<p:spPr/>` 常是**空的**，它的位置、字体、
 * 颜色全部从版式/母版继承。我们只记录"这格自己有没有写几何（`hasOwnGeometry`）"，绝不把继承来的坐标
 * 烘焙进形状——一烘焙就切断了 PowerPoint 自己的继承链，改母版时不再联动，也违反 §2.3"不许拍平"。
 * 继承值本身由 `inherit.ts` 现算成 `effectiveBoxPx` + `boxPxSources`（有效值 + 来源层两个都给，视图不许猜），
 * 它们是只读事实，不进任何写侧入参。
 */
import { children, findChildren, firstChild, loadPackage, parseDocument, readText, resolveTarget, text, type Element, type OpcPackage } from "@office-copilot/opc";
import { createGeometryInheritance, placeholderKeyOf, readExplicitBoxAxes, type BoxAxisSource, type GeometryInheritance } from "./inherit.js";
import { PRESENTATION_PART, isNotesSlideRelationship } from "./parts.js";
import { createColorInheritance, type ColorInheritance } from "./theme.js";
import { createTextStyleInheritance, readShapeTextStyles, type InheritedSizePt, type SizePtSource, type TextStyleInheritance } from "./text-styles.js";
import { halfPointToPt, type ShapeBoxPx, type ShapeGeometryPatch } from "./units.js";

/** 一个形状（`<p:sp>` / `<p:pic>` / `<p:graphicFrame>` / `<p:grpSp>` / `<p:cxnSp>`）的只读快照。 */
export interface SlideShape {
  /** `<p:cNvPr @id>`，版式内寻址用的形状号。 */
  shapeId: number;
  /** `<p:cNvPr @name>`，PowerPoint 侧边栏显示的名字。 */
  name: string;
  /** 形状大类：形状 / 图片 / 图形框（表格、图表走这个） / 组合 / 连接线（`<p:cxnSp>`）。 */
  kind: "shape" | "picture" | "graphicFrame" | "group" | "connector";
  /**
   * 这个形状在哪个组合里（父组 `<p:cNvPr @id>`）；顶层形状为 undefined。
   *
   * 为什么要单独报出来：组里的形状，几何写在**组的子坐标系**里（`chOff`/`chExt` 那一套映射），
   * 这一档不做坐标映射，所以它的 `boxPx`/`effectiveBoxPx` 一律为空、几何那一条 op 整条拒绝——
   * 宁可不给数也不给错数（与 D3 同一条线），而文字、字体颜色、删形状与组内换叠放次序都不受影响，照样能按 `shapeId` 寻址。
   */
  groupShapeId?: number;
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
  /** 每段里逐 run 的现文字，与 `paragraphRunFormats` 一一对应；AI 提 run 级改动要靠它对准 run 序号。 */
  paragraphRunTexts: string[][];
  /** 每段的 run 格式，与 `paragraphTexts` 一一对应（该段第 i 条 run 是第 i 项）；只含该 run 显式写了的属性。 */
  paragraphRunFormats: RunFormat[][];
  /** 形状自己写的 `<a:xfrm>`（CSS px @96DPI）；没有写（几何继承自版式）时 undefined，禁止烘焙（D3）。 */
  boxPx?: ShapeBoxPx;
  /**
   * 算完继承链的**有效**几何（CSS px @96DPI，逐轴，允许只有其中几轴）。
   *
   * 与 `boxPx` 的分工：那份是"这格自己写了什么"（写侧只许照它判断要不要覆盖，D3 不许烘焙），
   * 这份是" PowerPoint 实际会画在哪"（视图与 AI 要显示/回答位置时用这份，别再自己猜）。
   * 非占位符不继承，所以就是自己写的那几轴；三层都没写的轴不在表里。
   */
  effectiveBoxPx: ShapeGeometryPatch;
  /** `effectiveBoxPx` 每一轴的来源层，键与它一一对应（spec D3 要求的 `src`）。 */
  boxPxSources: Partial<Record<keyof ShapeBoxPx, BoxAxisSource>>;
}

/** 一条 run 自己的字符格式（读侧已换算成 pt 与 `#rrggbb`）。 */
export interface RunFormat {
  /** 字号 pt；**只表示这条 run 自己写了 `sz`**（从版式/母版继承时 undefined，写侧只认这一格，D3 不许烘焙）。 */
  sizePt?: number;
  /**
   * PowerPoint 实际画多大的**有效**字号（pt）：自己写了 `sz` 就用它，没写则沿
   * 段落 `a:pPr/a:defRPr` → 形状 `a:lstStyle` → 版式那一格 → 母版那一格 → 母版 `p:txStyles` 逐层算出来。
   *
   * 与几何的 `boxPx` / `effectiveBoxPx` 同一分工：这份只读，**绝不烤回 `<a:rPr sz>`**（一烤就永久切断母版联动，`AGENTS.md` §2.12）。
   * 整条链都没写时 undefined，不许拿 18pt 这类"常见默认"凑。
   */
  effectiveSizePt?: number;
  /** `effectiveSizePt` 来自哪一层；与它成对出现，视图那句"这个字号继承自…"靠它，不许自己猜。 */
  effectiveSizeSource?: SizePtSource;
  /** 粗体；只有显式写了 `b` 才给值，没写给 undefined（同样是继承）。 */
  bold?: boolean;
  /** 文字颜色 `#rrggbb`（§2.2）：`srgbClr` 直接要，`schemeClr` 沿母版 `p:clrMap` 与主题 `a:clrScheme` 解出来；带变换的、`phClr` 与继承来的都给 undefined。 */
  hex?: string;
}

/**
 * 一张幻灯片读不懂的那几种原因（结构化错误码，中文文案在装配层，`AGENTS.md` §5.7）。
 *
 * 码与句子都住在 `@office-copilot/shared` 那一份：桌面面板、`oc ppt` 与上送给模型的理由念同一套话（§3.2）。
 */
import { type SlideReadFailure } from "@office-copilot/shared";

export type { SlideReadFailure };

/** 一张幻灯片的演讲者备注（`ppt/notesSlides/notesSlideN.xml` 里那个 `type="body"` 占位符）。 */
export interface SlideNotes {
  /** 备注部件的包内路径，便于排查（不是关系 id，关系解析在我们这边已经完成）。 */
  partPath: string;
  /** 备注正文逐段文字（只取 `p:ph type="body"` 那一个形状；页码、幻灯片缩略图那些占位符不算备注）。 */
  paragraphTexts: string[];
  /**
   * 备注逐段的 run 截数，与 `paragraphTexts` 同长同序（一段一截就是 `1`）。
   *
   * 这是"整篇覆盖会不会拍平"的那半个事实：住在读模型里，装配层（`slideAgent`）才敢在提案阶段就判一刀，
   * 而不是等用户点确认、主进程整批不落地。
   */
  paragraphRunCounts: number[];
  /** 备注里带字符格式（`<a:rPr>`：颜色、加粗、字号）的 run 总截数；0 表示这一页的备注全是同一种格式。 */
  styledRunCount: number;
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
  /**
   * 这一页**读不懂**的原因；读得懂（哪怕真的一个形状都没有）时 undefined。
   *
   * 为什么必须有这一格：`shapes: []` 有两种完全不同的来历——"这页确实没形状"与"这页的部件我们读不动"。
   * 折成一句「这张幻灯片没有可读出的形状」会让面板与模型把后者说成前者，用户于是被告知"你的第 3 页是空的"
   * （实测撞到的，spec line 439）。
   */
  readFailure?: SlideReadFailure;
  /**
   * 这一页的演讲者备注（`<p:ph type="body">` 那个占位符，逐段文字）。
   *
   * 没有指向备注部件的关系时整格 undefined；有备注部件但没写字时仍是这一格，只是段落文字是空串
   * （WPS 自带模板那份实测：`<a:p><a:endParaRPr/></a:p>` 读回来就是 `[""]`）。
   * 这两件事必须分得开：AI 问"这页有备注吗"时，缺部件回"没有备注页"、空备注回"有备注页但没写字"。
   */
  notes?: SlideNotes;
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
  "p:cxnSp": "connector",
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
  const inheritance = createGeometryInheritance(pkg);
  const colors = createColorInheritance(pkg);
  const texts = createTextStyleInheritance(pkg);
  const slides: Slide[] = [];
  for (const [order, sldId] of sldIds.entries()) {
    const relId = sldId.attributes["r:id"];
    const rel = rels.find((entry) => entry.id === relId);
    if (!rel) continue; // 关系解析不到：跳过这张，宁可少读也不凭空造一张空幻灯片
    const partPath = resolveTarget(PRESENTATION_PART, rel.target);
    const { shapes, readFailure } = await readSlideShapes(pkg, partPath, inheritance, colors, texts);
    slides.push({
      partPath,
      order,
      slideId: Number(sldId.attributes["id"] ?? 0),
      shapes,
      ...(readFailure ? { readFailure } : {}),
      notes: await readSlideNotes(pkg, partPath),
    });
  }
  return { slideSizeEmu, slides };
}

/**
 * 读一张幻灯片的形状树顶层，并如实报告**读不动**是哪一种。
 *
 * @param pkg 已加载的包。
 * @param partPath 幻灯片部件路径。
 * @param inheritance 整份 deck 共用一个的几何继承解析器（版式/母版只解析一次）。
 * @param colors 整份 deck 共用一个的主题色解析器（母版 `clrMap` 与 theme 的 `clrScheme` 只解析一次）。
 * @param texts 整份 deck 共用一个的字号继承解析器（版式/母版的 `lstStyle` 与母版 `p:txStyles` 只解析一次）。
 * @returns 形状数组 + 失败码：部件缺失 `missing-part`、XML 解析失败 `invalid-xml`、没有 `cSld > spTree` `no-sp-tree`；
 *          读得动时 `readFailure` 为 undefined，**这时空数组就是"这一页真的没有形状"**。
 */
async function readSlideShapes(
  pkg: OpcPackage,
  partPath: string,
  inheritance: GeometryInheritance,
  colors: ColorInheritance,
  texts: TextStyleInheritance,
): Promise<{ shapes: SlideShape[]; readFailure?: SlideReadFailure }> {
  const xml = await readText(pkg.zip, partPath);
  if (xml === undefined) return { shapes: [], readFailure: "missing-part" };
  let slideRoot: Element;
  try {
    slideRoot = parseDocument(xml);
  } catch {
    return { shapes: [], readFailure: "invalid-xml" };
  }
  const cSld = firstChild(slideRoot, "p:cSld");
  const spTree = cSld ? firstChild(cSld, "p:spTree") : undefined;
  if (!spTree) return { shapes: [], readFailure: "no-sp-tree" };
  const shapes: SlideShape[] = [];
  /**
   * 逐层收形状：`<p:grpSp>` 自己进清单，随后把它内部的形状也收进来（标上 `groupShapeId`）。
   *
   * 为什么要往里走：一份真文件把两个文本框组合起来后，只扫顶层就等于**那两格文字在页面上不存在**——
   * 面板列不出、AI 画像里看不见、按 shapeId 也提不出改动（实测见 spec 阶段 2 那条「成组与连接线」）。
   */
  const walk = async (container: Element, withinGroup?: number): Promise<void> => {
    for (const el of children(container)) {
      const kind = SHAPE_TAGS[el.name];
      if (!kind) continue; // nvGrpSpPr / grpSpPr 等非形状子元素跳过
      shapes.push(await readShape(el, kind, partPath, inheritance, colors, texts, withinGroup));
      if (kind === "group") await walk(el, shapeIdOf(el));
    }
  };
  await walk(spTree);
  return { shapes };
}

/**
 * 读一张幻灯片的演讲者备注（经它自己的关系表找 `<p:notes>` 部件）。
 *
 * @param pkg 已加载的包。
 * @param slidePartPath 幻灯片部件路径，关系表在 `ppt/slides/_rels/` 下按它查。
 * @returns 备注正文（逐段）；这页没有指向备注部件的关系时 undefined（"没有备注页"），
 *   有部件但 `type="body"` 那个占位符没写字时返回空数组（"有备注页但没写字"）。部件缺失或解析失败一律 undefined，
 *   与 `readSlideShapes` 同一条宽容规则：备注读不动不许把整页判成读不懂。
 */
async function readSlideNotes(pkg: OpcPackage, slidePartPath: string): Promise<SlideNotes | undefined> {
  const rels = pkg.relsBySourcePart.get(slidePartPath) ?? [];
  const rel = rels.find((entry) => isNotesSlideRelationship(entry.type));
  if (!rel) return undefined;
  const partPath = resolveTarget(slidePartPath, rel.target);
  const xml = await readText(pkg.zip, partPath);
  if (xml === undefined) return undefined;
  let notesRoot: Element;
  try {
    notesRoot = parseDocument(xml);
  } catch {
    return undefined;
  }
  const cSld = firstChild(notesRoot, "p:cSld");
  if (!cSld || !firstChild(cSld, "p:spTree")) return undefined;
  const bodyShape = findNotesBodyShape(notesRoot);
  const txBody = bodyShape ? firstChild(bodyShape, "p:txBody") : undefined;
  const paragraphs = txBody ? findChildren(txBody, "a:p") : [];
  const runCounts = paragraphs.map((p) => findChildren(p, "a:r").length);
  const styledRunCount = paragraphs.reduce(
    (total, p) => total + findChildren(p, "a:r").filter((run) => firstChild(run, "a:rPr")).length,
    0,
  );
  return {
    partPath,
    paragraphTexts: paragraphs.map((p) => text(p)),
    paragraphRunCounts: runCounts,
    styledRunCount,
  };
}

/**
 * 在备注部件里找出正文那一格占位符（`<p:sp>` 且 `p:nvPr/p:ph/@type === "body"`）。
 *
 * 读侧与写侧共用这一份判定（§3.2）：备注页上还有 `sldImg`（幻灯片缩略图）与 `sldNum`/`hdr`/`ftr`（正文是 `<a:fld>` 算出来的），
 * 把它们念成备注是编内容，把它们写成备注是改坏了不该动的东西。
 *
 * @param notesRoot 备注部件（`<p:notes>`）的根元素。
 * @returns 正文那一格 `<p:sp>`；没有形状树或没有 body 占位符时 undefined。
 */
export function findNotesBodyShape(notesRoot: Element): Element | undefined {
  const cSld = firstChild(notesRoot, "p:cSld");
  const spTree = cSld ? firstChild(cSld, "p:spTree") : undefined;
  if (!spTree) return undefined;
  for (const el of children(spTree)) {
    if (el.name !== "p:sp") continue;
    const nvPr = firstChild(el, "p:nvSpPr");
    const phHost = nvPr ? firstChild(nvPr, "p:nvPr") : undefined;
    const ph = phHost ? firstChild(phHost, "p:ph") : undefined;
    if (ph?.attributes["type"] === "body") return el;
  }
  return undefined;
}

/**
 * 读单个形状元素。
 *
 * @param el 形状根元素（`<p:sp>` 等）。
 * @param kind 已在 `readSlideShapes` 里判定的大类。
 * @param slidePartPath 所属幻灯片的部件路径，用来找它自己的版式（只有占位符会用到）。
 * @param inheritance 几何继承解析器。
 * @param colors 主题色解析器（run 上只写了 `a:schemeClr` 时用它解成 hex）。
 * @param texts 字号继承解析器（run 没写 `sz` 时沿段落→形状→版式→母版→`p:txStyles` 求有效字号）。
 * @returns 只读快照：自己写的事实（`boxPx`/`hasOwnGeometry`/`sizePt`）与算完继承的事实（`effectiveBoxPx`/`boxPxSources`/`effectiveSizePt`）。
 */
async function readShape(
  el: Element,
  kind: SlideShape["kind"],
  slidePartPath: string,
  inheritance: GeometryInheritance,
  colors: ColorInheritance,
  texts: TextStyleInheritance,
  groupShapeId?: number,
): Promise<SlideShape> {
  const nvPr = shapeNvPrOf(el);
  const cNvPr = nvPr ? firstChild(nvPr, "p:cNvPr") : undefined;
  const phHost = nvPr ? firstChild(nvPr, "p:nvPr") : undefined;
  const ph = phHost ? firstChild(phHost, "p:ph") : undefined;
  const spPr = firstChild(el, "p:spPr");
  const txBody = firstChild(el, "p:txBody");
  const paragraphs = txBody ? findChildren(txBody, "a:p") : [];
  // 一次扫描拿到每段的 `<a:r>`，条数 / 文字 / 格式三份都从它派生（扫三遍是同一件事重复做，§3.2）。
  const runsByParagraph = paragraphs.map((p) => findChildren(p, "a:r"));
  const ownAxes = readExplicitBoxAxes(spPr);
  // 非占位符没有继承这一说（PowerPoint 不给它查版式），key 传 undefined 让解析器只回自己写的那几轴。
  const key = placeholderKeyOf(el);
  const effective = await inheritance.resolve(slidePartPath, key, ownAxes);
  // 每段一档继承字号（链上要读版式/母版部件，形状那一档读一次就够）。
  const shapeTextStyles = readShapeTextStyles(el);
  const inheritedSizeByParagraph = await Promise.all(paragraphs.map((p) => texts.resolve(slidePartPath, shapeTextStyles, paragraphSizeInput(p))));
  return {
    shapeId: shapeIdOf(el) ?? 0,
    name: String(cNvPr?.attributes["name"] ?? ""),
    kind,
    groupShapeId,
    isPlaceholder: !!ph,
    placeholderRef: ph ? String(ph.attributes["type"] ?? ph.attributes["idx"] ?? "") : undefined,
    hasOwnGeometry: !!spPr && !!firstChild(spPr, "a:xfrm"),
    paragraphTexts: paragraphs.map((p) => text(p)),
    paragraphRunCounts: runsByParagraph.map((runs) => runs.length),
    paragraphRunTexts: runsByParagraph.map((runs) => runs.map((run) => text(run))),
    // run 格式可能要吃两次继承链（解色要读母版与主题，字号要读版式与 `p:txStyles`），所以逐 run 异步算、并发读，顺序不变。
    paragraphRunFormats: await Promise.all(
      runsByParagraph.map(async (runs, paragraphIndex) => {
        const inherited = inheritedSizeByParagraph[paragraphIndex];
        return Promise.all(runs.map(async (run) => withEffectiveSize(await readRunFormat(run, colors, slidePartPath), inherited)));
      }),
    ),
    // 组里的形状：它的 `<a:xfrm>` 是组的子坐标系里的数（要经 chOff/chExt 与 ext 的比例映射回页面），
    // 这一档不做那个映射，所以三个几何字段一律留空——给一个错的位置比不给更坏（与 D3 同一条线）。
    boxPx: groupShapeId === undefined ? completeBox(ownAxes) : undefined,
    effectiveBoxPx: groupShapeId === undefined ? effective.box : {},
    boxPxSources: groupShapeId === undefined ? effective.sources : {},
  };
}

/**
 * 形状自己那节非可视属性（`<p:nvSpPr>` 一类）——五种形状宿主的名字不一样，只有这一处把它们并成一条链。
 *
 * @param el 形状根元素（`p:sp`/`p:pic`/`p:graphicFrame`/`p:grpSp`/`p:cxnSp`）。
 * @returns 该宿主的 `p:nv*Pr` 元素；一个都没有时 undefined（读成"没有 id/名字"，不抛错）。
 */
function shapeNvPrOf(el: Element): Element | undefined {
  return (
    firstChild(el, "p:nvSpPr") ??
    firstChild(el, "p:nvPicPr") ??
    firstChild(el, "p:nvGraphicFramePr") ??
    firstChild(el, "p:nvCxnSpPr") ??
    firstChild(el, "p:nvGrpSpPr")
  );
}

/**
 * 形状元素自己的 `<p:cNvPr @id>`。
 *
 * @param el 形状根元素。
 * @returns 数字形状号；没有 `cNvPr`、或 id 不是数字时 undefined。
 */
export function shapeIdOf(el: Element): number | undefined {
  const nvPr = shapeNvPrOf(el);
  const raw = nvPr ? firstChild(nvPr, "p:cNvPr")?.attributes["id"] : undefined;
  return raw !== undefined && /^\d+$/.test(raw) ? Number(raw) : undefined;
}

/**
 * 读一个段落的字号链输入：层级与这一段自己写的 `defRPr @sz`。
 *
 * @param paragraph 段落元素 `<a:p>`。
 * @returns `level` 是 `<a:pPr @lvl>`（0 基，没写按 0）；`sizePt` 只在 `<a:pPr><a:defRPr @sz>` 真写了才给（半磅已换成 pt，§2.1）。
 */
function paragraphSizeInput(paragraph: Element): { level: number; sizePt?: number } {
  const pPr = firstChild(paragraph, "a:pPr");
  const rawLvl = Number(pPr?.attributes["lvl"]);
  const rawSz = pPr ? firstChild(pPr, "a:defRPr")?.attributes["sz"] : undefined;
  const halfPoints = Number(rawSz);
  return {
    level: Number.isFinite(rawLvl) ? rawLvl : 0,
    ...(rawSz !== undefined && Number.isFinite(halfPoints) ? { sizePt: halfPointToPt(halfPoints) } : {}),
  };
}

/**
 * 给一条 run 的格式补上"实际画多大"那一档，**一个字都不碰 `sizePt`**。
 *
 * @param format `readRunFormat` 读出的显式格式（写侧只看这份）。
 * @param inherited 这一段在继承链上的字号；整条链都没写时 undefined。
 * @returns `sizePt` 有值时有效字号就是它自己（来源记 `run`）；否则用链上那一档；两边都没有时原样返回。
 */
function withEffectiveSize(format: RunFormat, inherited: InheritedSizePt | undefined): RunFormat {
  if (format.sizePt !== undefined) return { ...format, effectiveSizePt: format.sizePt, effectiveSizeSource: "run" };
  if (!inherited) return format;
  return { ...format, effectiveSizePt: inherited.sizePt, effectiveSizeSource: inherited.source };
}

/**
 * 读一条 run 自己的 `<a:rPr>`：字号（半磅→pt）、粗体、颜色（`srgbClr` 直接要，`schemeClr` 沿母版与主题解成 hex）。
 *
 * @param run 段落里的 `<a:r>` 元素。
 * @param colors 主题色解析器（只在写了 `a:schemeClr` 时才会用到）。
 * @param slidePartPath 所属幻灯片部件路径，解色时从这里找版式→母版→theme。
 * @returns 只填了该 run **显式写了**的属性；没写的一律 undefined，表示"这个值是从版式继承的"（D3 不许烘焙）。
 *   带 `a:lumMod`/`a:tint` 这类变换的主题色**不给 hex**（本机没有能核对 PowerPoint 画面的通道，算错的颜色比没有颜色更坏）。
 */
async function readRunFormat(run: Element, colors: ColorInheritance, slidePartPath: string): Promise<RunFormat> {
  const rPr = firstChild(run, "a:rPr");
  if (!rPr) return {};
  const format: RunFormat = {};
  const rawSz = rPr.attributes["sz"];
  const sizePt = Number(rawSz);
  if (rawSz !== undefined && Number.isFinite(sizePt)) format.sizePt = halfPointToPt(sizePt);
  const rawBold = rPr.attributes["b"];
  if (rawBold !== undefined) format.bold = rawBold === "1" || rawBold === "true" || rawBold === "on";
  const solidFill = firstChild(rPr, "a:solidFill");
  const srgbClr = solidFill ? firstChild(solidFill, "a:srgbClr") : undefined;
  const val = srgbClr?.attributes["val"];
  if (val && /^[0-9a-fA-F]{6}$/.test(val)) {
    format.hex = `#${val.toLowerCase()}`;
    return format;
  }
  const schemeClr = solidFill ? firstChild(solidFill, "a:schemeClr") : undefined;
  const role = schemeClr?.attributes["val"];
  if (schemeClr && role && children(schemeClr).length === 0) {
    const hex = await colors.resolve(slidePartPath, role);
    if (hex) format.hex = hex;
  }
  return format;
}

/**
 * 把逐轴的几何轴收成"四个都齐"的 `ShapeBoxPx`。
 *
 * @param axes `readExplicitBoxAxes` 读出的轴（单位 px）；可以只有一两轴。
 * @returns 四个轴都在时给整份盒子；缺任何一个时 undefined——视图与写侧都按"这格没写几何"处理（D3 不许拿继承值来凑齐）。
 */
function completeBox(axes: ShapeGeometryPatch): ShapeBoxPx | undefined {
  const { leftPx, topPx, widthPx, heightPx } = axes;
  if (leftPx === undefined || topPx === undefined || widthPx === undefined || heightPx === undefined) return undefined;
  return { leftPx, topPx, widthPx, heightPx };
}
