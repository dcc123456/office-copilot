/**
 * 有效字号（`sz`）的继承链解析：run 自己没写字号时，PowerPoint 实际拿哪一档来画。
 *
 * 一手依据（`AGENTS.md` §6.1/§6.2）：
 * - ECMA-376 Part 1 §19.2.5（`p:sldMaster` 那一族，含母版 `p:txStyles` 的三档 `titleStyle`/`bodyStyle`/`otherStyle`，
 *   每档下面是 `lvl1pPr…lvl9pPr`，字号写在其中的 `a:defRPr @sz`）。**小节号只核到 §19.2.5 这一层**，更细的号没逐条比对 ISO 原文，
 *   所以代码一律按**元素名**匹配，不靠号。
 * - 分层顺序与第三方口径一致：run → 段落 `pPr/defRPr` → 形状自己的 `txBody/lstStyle` → 版式那一格 → 母版那一格 → 母版 `p:txStyles`。
 *   本机 python-pptx 1.0.2 **不计算**最终继承字号（它只记录各层显式写了的值，`font.size` 对继承的文字回 `None`），
 *   所以这一档没有现成算法可抄，只能照规范走链——与 `theme.ts` 解 `schemeClr` 同一处境。
 *
 * 红线（`docs/office-parity-spec.md` D3 / `AGENTS.md` §2.12）：继承来的字号**只交视图与模型看**，绝不烤进 `<a:rPr sz>`——
 * 一烤就永久切断版式与母版的联动。所以读模型里 `RunFormat.sizePt` 仍只表示"这条 run 自己写了 sz"（写侧照旧只看它），
 * 继承值另放 `effectiveSizePt` + `effectiveSizeSource`，与几何的 `boxPx` / `effectiveBoxPx` 同一分工。
 */
import { children, firstChild, localName, parseDocument, readText, type Element, type OpcPackage } from "@office-copilot/opc";
import { MASTER_TYPE_OF_LAYOUT_TYPE, placeholderKeyOf, relatedPartPath, type PlaceholderKey } from "./inherit.js";
import { halfPointToPt } from "./units.js";

/**
 * 有效字号来自哪一层。`run` 表示这条 run 自己写了 `<a:rPr @sz>`：那一档由调用方（`read.ts`）直接给，
 * {@link TextStyleInheritance.resolve} 从不返回它——写侧只认 `RunFormat.sizePt`，所以链上不需要 run 那一层。
 */
export type SizePtSource = "run" | "paragraph" | "shape-list" | "layout" | "master-list" | "master-text-styles";

/** 求出来的一档继承字号。 */
export interface InheritedSizePt {
  /** 单位 pt（半磅在链条入口就换算完，§2.1 换算单点）。 */
  sizePt: number;
  /** 这一档来自哪一层。 */
  source: SizePtSource;
}

/** {@link createTextStyleInheritance} 的返回值。 */
export type TextStyleInheritance = ReturnType<typeof createTextStyleInheritance>;

/** 一个形状送进继承链所需的全部信息（由 `readShapeTextStyles` 从手上的元素读出来）。 */
export interface ShapeTextStyles {
  /** 占位符匹配键；非占位符是 undefined。 */
  key: PlaceholderKey | undefined;
  /** 该形状自己 `<p:txBody><a:lstStyle>` 的逐档字号（pt，下标即层级）。 */
  ownLvlSizes: (number | undefined)[];
}

/** 一个版式/母版部件里的占位符格子：匹配键 + 它自己 `txBody/lstStyle` 的逐档字号。 */
interface StyleCell {
  key: PlaceholderKey;
  lvlSizes: (number | undefined)[];
}

/**
 * 读一个形状元素送进字号链的那两样。
 *
 * @param el 形状根元素（`<p:sp>`）。
 * @returns 占位符键（非占位符为 undefined）与该形状自己写的 `lstStyle` 逐档字号。
 */
export function readShapeTextStyles(el: Element): ShapeTextStyles {
  return { key: placeholderKeyOf(el), ownLvlSizes: readLvlSizePts(firstChild(el, "p:txBody")) };
}

/**
 * 读一个 `<p:txBody>` 里**这个形状自己**写的 `a:lstStyle` 逐档字号。
 *
 * @param txBody 形状或格子元素的 `<p:txBody>`，可以不存在。
 * @returns 下标即层级（0 基）的字号表（pt）；没写的档是 `undefined`。
 */
export function readLvlSizePts(txBody: Element | undefined): (number | undefined)[] {
  const lstStyle = txBody ? firstChild(txBody, "a:lstStyle") : undefined;
  return lstStyle ? lvlSizesOf(lstStyle) : [];
}

/**
 * 建一条字号继承链：一份包一个，版式与母版的格子与 `p:txStyles` 都只解析一次（几十页共用同一版式）。
 *
 * @param pkg 已 `loadPackage` 的包。
 * @returns {@link TextStyleInheritance}：给它幻灯片路径、形状的字号输入与段落层级，它回继承字号与来源层。
 */
export function createTextStyleInheritance(pkg: OpcPackage) {
  const cellsByPart = new Map<string, Promise<StyleCell[]>>();
  const txStylesByMaster = new Map<string, Promise<MasterTextStyles>>();

  async function cellsOf(partPath: string): Promise<StyleCell[]> {
    let cached = cellsByPart.get(partPath);
    if (!cached) {
      cached = readPlaceholderStyleCells(pkg, partPath);
      cellsByPart.set(partPath, cached);
    }
    return cached;
  }

  async function txStylesOf(masterPart: string): Promise<MasterTextStyles> {
    let cached = txStylesByMaster.get(masterPart);
    if (!cached) {
      cached = readMasterTxStyles(pkg, masterPart);
      txStylesByMaster.set(masterPart, cached);
    }
    return cached;
  }

  return {
    /**
     * 求一个段落在继承链上的字号（run 自己写了 `sz` 时调用方不该来这里）。
     *
     * @param slidePartPath 幻灯片部件路径，链从它的 slideLayout 关系往上走。
     * @param shape 该形状的字号输入（{@link readShapeTextStyles} 读出来的那份）。
     * @param paragraph 该段的层级（`<a:pPr @lvl>`，0 基）与它自己写了的 `defRPr @sz`（pt）。
     * @returns 第一档写了 `sz` 的那层（paragraph → shape-list → layout → master-list → 母版 `p:txStyles` 依次问）；
     *   整条链都没写、或版式/母版关系断掉时 undefined——**不许拿 18pt 这种"常见默认"凑**。
     */
    async resolve(slidePartPath: string, shape: ShapeTextStyles, paragraph: { level: number; sizePt?: number }): Promise<InheritedSizePt | undefined> {
      if (paragraph.sizePt !== undefined) return { sizePt: paragraph.sizePt, source: "paragraph" };
      const level = paragraph.level;
      const own = shape.ownLvlSizes[level];
      if (own !== undefined) return { sizePt: own, source: "shape-list" };
      const layoutPart = relatedPartPath(pkg, slidePartPath, "/slideLayout");
      if (!layoutPart) return undefined;
      // 与几何那条链同一套匹配：幻灯片 → 版式按 `@idx`，版式 → 母版按**版式那一格的类型**映射（`ctrTitle` → `title`）。
      let masterType: string | undefined;
      if (shape.key) {
        masterType = MASTER_TYPE_OF_LAYOUT_TYPE[shape.key.type] ?? shape.key.type;
        const layoutCell = (await cellsOf(layoutPart)).find((cell) => cell.key.idx === shape.key!.idx);
        if (layoutCell) {
          const fromLayout = layoutCell.lvlSizes[level];
          if (fromLayout !== undefined) return { sizePt: fromLayout, source: "layout" };
          masterType = MASTER_TYPE_OF_LAYOUT_TYPE[layoutCell.key.type] ?? layoutCell.key.type;
        }
      }
      const masterPart = relatedPartPath(pkg, layoutPart, "/slideMaster");
      if (!masterPart) return undefined;
      if (masterType) {
        const masterCell = (await cellsOf(masterPart)).find((cell) => cell.key.type === masterType);
        const fromMasterCell = masterCell?.lvlSizes[level];
        if (fromMasterCell !== undefined) return { sizePt: fromMasterCell, source: "master-list" };
      }
      const fromTxStyles = (await txStylesOf(masterPart))[roleOfMasterType(masterType)][level];
      return fromTxStyles === undefined ? undefined : { sizePt: fromTxStyles, source: "master-text-styles" };
    },
  };
}

/** 母版 `p:txStyles` 的三档名（`p:titleStyle` / `p:bodyStyle` / `p:otherStyle`）。 */
type TextStyleRole = "titleStyle" | "bodyStyle" | "otherStyle";

/** 母版 `p:txStyles` 读出来的那份表：档名 → 逐档字号（pt，下标即层级，没写的档是 undefined）。 */
type MasterTextStyles = Record<TextStyleRole, (number | undefined)[]>;

/**
 * 母版占位符类型 → `p:txStyles` 该看哪一档。
 *
 * 真文件核过：本机 python-pptx 默认母版与 WPS 模板母版的这三档**都不带 `@phType` / `@styleId` 属性**，
 * 所以只能按档名与类型对应（`title`→`titleStyle`、`body`→`bodyStyle`，其余（页脚、日期、页码、非占位符的文本框）
 * 一律落 `otherStyle`——这也是普通文本框不写字号时画出来是 18pt 的那条路，两份真母版的 `otherStyle/lvl1pPr` 都写着 `sz="1800"`）。
 *
 * @param masterType 版式那一格映射后的母版类型；非占位符时 undefined。
 * @returns 三档名之一。
 */
function roleOfMasterType(masterType: string | undefined): TextStyleRole {
  if (masterType === "title") return "titleStyle";
  if (masterType === "body") return "bodyStyle";
  return "otherStyle";
}

/**
 * 读母版的 `p:txStyles`：三档各自的逐档字号。
 *
 * @param pkg 已加载的包。
 * @param masterPart 母版部件路径。
 * @returns 三档名 → 字号表（pt）；缺哪一档给空表，母版读不到时三档全空（调用方拿 undefined，不猜）。
 */
async function readMasterTxStyles(pkg: OpcPackage, masterPart: string): Promise<MasterTextStyles> {
  const empty: MasterTextStyles = { titleStyle: [], bodyStyle: [], otherStyle: [] };
  const text = await readText(pkg.zip, masterPart);
  if (text === undefined) return empty;
  let root: Element;
  try {
    root = parseDocument(text);
  } catch {
    return empty;
  }
  const txStyles = firstChild(root, "p:txStyles");
  if (!txStyles) return empty;
  const roles: Record<TextStyleRole, (number | undefined)[]> = { titleStyle: [], bodyStyle: [], otherStyle: [] };
  for (const style of children(txStyles)) {
    const name = localName(style.name);
    if (name !== "titleStyle" && name !== "bodyStyle" && name !== "otherStyle") continue;
    roles[name] = lvlSizesOf(style);
  }
  return roles;
}

/**
 * 取一个带 `lvlNpPr` 的元素（`a:lstStyle` 或 `p:titleStyle` 这类档）的逐档字号。
 *
 * @param host 带 `lvl1pPr…lvl9pPr` 子元素的元素。
 * @returns 下标即层级（0 基）的字号（pt）；该档没 `defRPr @sz` 时 `undefined`。
 */
function lvlSizesOf(host: Element): (number | undefined)[] {
  const sizes: (number | undefined)[] = [];
  for (const lvl of children(host)) {
    const match = /^lvl(\d)pPr$/.exec(localName(lvl.name));
    if (!match) continue;
    const rawSz = firstChild(lvl, "a:defRPr")?.attributes["sz"];
    const halfPoints = Number(rawSz);
    sizes[Number(match[1]) - 1] = rawSz !== undefined && Number.isFinite(halfPoints) ? halfPointToPt(halfPoints) : undefined;
  }
  return sizes;
}

/**
 * 读一个部件（版式或母版）里顶层占位符格子的 `txBody/lstStyle` 字号表。
 *
 * @param pkg 已加载的包。
 * @param partPath 部件路径。
 * @returns 格子数组（匹配键 + 逐档字号）；部件缺失或解析失败时空数组——链断在这里，调用方拿 undefined。
 */
async function readPlaceholderStyleCells(pkg: OpcPackage, partPath: string): Promise<StyleCell[]> {
  const text = await readText(pkg.zip, partPath);
  if (text === undefined) return [];
  let root: Element;
  try {
    root = parseDocument(text);
  } catch {
    return [];
  }
  const cSld = firstChild(root, "p:cSld");
  const spTree = cSld ? firstChild(cSld, "p:spTree") : undefined;
  if (!spTree) return [];
  const cells: StyleCell[] = [];
  for (const el of children(spTree)) {
    if (localName(el.name) !== "sp") continue; // 只认顶层形状：组合里的子形状不参与这套匹配
    const key = placeholderKeyOf(el);
    if (!key) continue;
    cells.push({ key, lvlSizes: readLvlSizePts(firstChild(el, "p:txBody")) });
  }
  return cells;
}
