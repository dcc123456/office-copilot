/**
 * 占位符几何的继承链解析（`docs/office-parity-spec.md` D3：**只读解析，绝不写回拍平**）。
 *
 * 一手依据是本机读到的 python-pptx 1.0.2 源码（`tmp/pylibs/pptx/shapes/placeholder.py` 的 `_InheritsDimensions`），
 * 它是这条规则当前最主流的落地，也是本仓 §8 那个"独立读取器"：
 * - 只有带 `<p:ph>` 的占位符继承几何，且**四个轴各自独立回退**（`_effective_value`：自己有就用自己的，没有才问上一层，
 *   所以"自己写了 `a:off`、没写 `a:ext`"这种半份几何是合法的，宽/高从版式补）；
 * - 幻灯片 → 版式按 `<p:ph @idx>` 匹配（`CT_Placeholder.idx` 的 default=0，见 `pptx/oxml/shapes/shared.py`），取第一个 idx 相同的格子；
 * - 版式 → 母版按**版式那一格自己的类型**映射到母版的类型再匹配（`ctrTitle`→`title`，
 *   `body`/`subTitle`/`obj`/`chart`/`tbl`/`pic`/`clipArt`/`dgm`/`media`→`body`，`title`/`dt`/`ftr`/`sldNum` 原样；
 *   表里没有的类型原样传下去——python-pptx 那种情况直接 KeyError，我们只当查不到）；
 * - 母版那一格也没写就是没有：不往下用 `defaultTextStyle`/`p:sldSz` 猜，也不烤进形状（一烤就永久切断版式与母版的联动，§2.3）。
 */
import { firstChild, localName, readText, resolveTarget, xml, type OpcPackage } from "@office-copilot/opc";
import { emuToPx, type ShapeGeometryPatch } from "./units.js";

/** 一个几何轴的来源层：形状自己写的、版式那一格写的、母版那一格写的。 */
export type BoxAxisSource = "shape" | "layout" | "master";

/** 占位符的有效几何：每个轴的值 + 这个值来自哪一层（D3 要求"有效值 + 来源"两个都给，视图不许自己猜）。 */
export interface EffectiveBox {
  /** 四个轴里"三层中有一层写了"的那些；都没写的轴不在表里。单位 CSS px @96DPI。 */
  box: ShapeGeometryPatch;
  /** `box` 里每一轴来自哪一层，键与 `box` 一一对应。 */
  sources: Partial<Record<keyof ShapeGeometryPatch, BoxAxisSource>>;
}

/** `<p:ph>` 的匹配键。`idx` 是 `@idx`（缺省 0），`type` 是 `@type`（缺省 `obj`，与 `CT_Placeholder` 的 default 一致）。 */
export interface PlaceholderKey {
  idx: number;
  type: string;
}

/** 某个部件（版式或母版）里一个占位符格子：匹配键 + 它自己显式写的几何轴。 */
interface PlaceholderCell {
  key: PlaceholderKey;
  axes: ShapeGeometryPatch;
}

/** 版式那一格的类型 → 母版那一格的类型（`LayoutPlaceholder._base_placeholder` 的那张表，键是 XML 里的类型串）。 */
const MASTER_TYPE_OF_LAYOUT_TYPE: Record<string, string> = {
  title: "title",
  ctrTitle: "title",
  body: "body",
  subTitle: "body",
  obj: "body",
  chart: "body",
  tbl: "body",
  pic: "body",
  clipArt: "body",
  dgm: "body",
  media: "body",
  dt: "dt",
  ftr: "ftr",
  sldNum: "sldNum",
};

/** 四个轴的读取顺序：`a:off @x`、`a:off @y`、`a:ext @cx`、`a:ext @cy`。 */
const BOX_AXES: Array<{ axis: keyof ShapeGeometryPatch; element: "a:off" | "a:ext"; attribute: "x" | "y" | "cx" | "cy" }> = [
  { axis: "leftPx", element: "a:off", attribute: "x" },
  { axis: "topPx", element: "a:off", attribute: "y" },
  { axis: "widthPx", element: "a:ext", attribute: "cx" },
  { axis: "heightPx", element: "a:ext", attribute: "cy" },
];

/**
 * 读一个 `<p:spPr>` 里**显式写了**的几何轴（EMU→px 走 `units.ts`，§2.1）。
 *
 * 与读模型那个"四个都齐才给"的 `boxPx` 不同，这里是逐轴的：占位符可以只写 `a:off`、宽/高从版式来（D3 的逐轴回退）。
 *
 * @param spPr 形状的 `<p:spPr>`，可以不存在。
 * @returns 只含写了且读得出数字的轴（单位 px）；一个都没写时是空对象。
 */
export function readExplicitBoxAxes(spPr: xml.Element | undefined): ShapeGeometryPatch {
  const xfrm = spPr ? firstChild(spPr, "a:xfrm") : undefined;
  if (!xfrm) return {};
  const axes: ShapeGeometryPatch = {};
  for (const { axis, element, attribute } of BOX_AXES) {
    const host = firstChild(xfrm, element);
    const raw = host?.attributes[attribute];
    const emu = Number(raw);
    if (raw !== undefined && Number.isFinite(emu)) axes[axis] = emuToPx(emu);
  }
  return axes;
}

/**
 * 建一个继承链解析器：一份包一个，版式/母版部件只解析一次（多张幻灯片共用同一版式）。
 *
 * @param pkg 已 `loadPackage` 的包。
 * @returns 用形状自己的几何轴去补版式与母版那两层的解析入口（{@link GeometryInheritance}）。
 */
export function createGeometryInheritance(pkg: OpcPackage) {
  const cellsByPart = new Map<string, Promise<PlaceholderCell[]>>();

  /** 读一个部件里所有顶层占位符格子（版式与母版的形状树与幻灯片同形）。 */
  async function cellsOf(partPath: string): Promise<PlaceholderCell[]> {
    let cached = cellsByPart.get(partPath);
    if (!cached) {
      cached = readPlaceholderCells(pkg, partPath);
      cellsByPart.set(partPath, cached);
    }
    return cached;
  }

  return {
    /**
     * 求一个形状的有效几何：自己写的轴优先，没写的按 版式 → 母版 逐轴回退。
     *
     * @param slidePartPath 幻灯片部件路径（从这里找它自己的 slideLayout 关系）。
     * @param key 形状 `<p:ph>` 的匹配键（idx/type）；**不是占位符时传 undefined**——PowerPoint 不让普通形状继承几何，
     *   传了会照着版式第一格凭空补数。
     * @param own 形状自己显式写的几何轴。
     * @returns 逐轴的值与来源层；版式/母版关系查不到时只有 `own` 那些轴（不猜）。
     */
    async resolve(slidePartPath: string, key: PlaceholderKey | undefined, own: ShapeGeometryPatch): Promise<EffectiveBox> {
      const layers: Array<[BoxAxisSource, ShapeGeometryPatch]> = [["shape", own]];
      const layoutPart = key ? relatedPartPath(pkg, slidePartPath, "/slideLayout") : undefined;
      if (key && layoutPart) {
        const cell = (await cellsOf(layoutPart)).find((entry) => entry.key.idx === key.idx);
        if (cell) {
          layers.push(["layout", cell.axes]);
          const masterPart = relatedPartPath(pkg, layoutPart, "/slideMaster");
          if (masterPart) {
            const masterType = MASTER_TYPE_OF_LAYOUT_TYPE[cell.key.type] ?? cell.key.type;
            const masterCell = (await cellsOf(masterPart)).find((entry) => entry.key.type === masterType);
            if (masterCell) layers.push(["master", masterCell.axes]);
          }
        }
      }
      return mergeAxisLayers(layers);
    },
  };
}

/** {@link createGeometryInheritance} 的返回值：一个包一份，供 `read.ts` 逐形状求有效几何。 */
export type GeometryInheritance = ReturnType<typeof createGeometryInheritance>;

/**
 * 按"逐轴取第一个写了的层"合并三层几何（`_effective_value` 的那句"自己有就用，否则问上一层"）。
 *
 * @param layers 从近到远排列的层（shape → layout → master）。
 * @returns 有效几何与每一轴的来源；某轴所有层都没写时该轴既不在 `box` 也不在 `sources`。
 */
function mergeAxisLayers(layers: Array<[BoxAxisSource, ShapeGeometryPatch]>): EffectiveBox {
  const box: ShapeGeometryPatch = {};
  const sources: Partial<Record<keyof ShapeGeometryPatch, BoxAxisSource>> = {};
  for (const [source, axes] of layers) {
    for (const axis of Object.keys(axes) as Array<keyof ShapeGeometryPatch>) {
      if (box[axis] === undefined && axes[axis] !== undefined) {
        box[axis] = axes[axis];
        sources[axis] = source;
      }
    }
  }
  return { box, sources };
}

/**
 * 读一个部件（幻灯片/版式/母版）`spTree` 顶层的占位符格子。
 *
 * @param pkg 已加载的包。
 * @param partPath 部件路径。
 * @returns 格子数组；部件读不到、解析不出或没有形状树时是空数组（继承链断在这里，不猜）。
 */
async function readPlaceholderCells(pkg: OpcPackage, partPath: string): Promise<PlaceholderCell[]> {
  const text = await readText(pkg.zip, partPath);
  if (text === undefined) return [];
  let root: xml.Element;
  try {
    root = xml.parseDocument(text);
  } catch {
    return [];
  }
  const cSld = firstChild(root, "p:cSld");
  const spTree = cSld ? firstChild(cSld, "p:spTree") : undefined;
  if (!spTree) return [];
  const cells: PlaceholderCell[] = [];
  for (const el of xml.children(spTree)) {
    if (localName(el.name) !== "sp") continue; // 只认顶层形状：组合里的子形状不参与这套匹配
    const ph = placeholderOf(el);
    if (!ph) continue;
    cells.push({ key: ph, axes: readExplicitBoxAxes(firstChild(el, "p:spPr")) });
  }
  return cells;
}

/**
 * 取一个形状元素的 `<p:ph>` 匹配键。
 *
 * @param shapeEl `<p:sp>` 元素。
 * @returns 占位符键（idx 缺省 0、type 缺省 `obj`）；不是占位符时 undefined。
 */
function placeholderOf(shapeEl: xml.Element): PlaceholderKey | undefined {
  const nvPr = firstChild(shapeEl, "p:nvSpPr") ?? firstChild(shapeEl, "p:nvGrpSpPr");
  const inner = nvPr ? firstChild(nvPr, "p:nvPr") : undefined;
  const ph = inner ? firstChild(inner, "p:ph") : undefined;
  if (!ph) return undefined;
  const idx = Number(ph.attributes["idx"]);
  return { idx: Number.isFinite(idx) ? idx : 0, type: String(ph.attributes["type"] ?? "obj") };
}

/**
 * 顺着一条关系找到目标部件的包内路径。
 *
 * 几何链（这里）与主题色链（`theme.ts` 的 版式 → 母版 → theme）走的是同一件事：按关系类型 URI 结尾找目标，
 * 所以只有这一份实现（§3.2）。
 *
 * @param pkg 已加载的包。
 * @param sourcePart 起点部件路径。
 * @param typeSuffix 关系类型 URI 的结尾（如 `/slideLayout`）。
 * @returns 目标部件路径；没有这条关系、是外部目标或路径算不出来时 undefined。
 */
export function relatedPartPath(pkg: OpcPackage, sourcePart: string, typeSuffix: string): string | undefined {
  const rel = (pkg.relsBySourcePart.get(sourcePart) ?? []).find((entry) => entry.targetMode !== "External" && entry.type.endsWith(typeSuffix));
  return rel ? resolveTarget(sourcePart, rel.target) : undefined;
}
