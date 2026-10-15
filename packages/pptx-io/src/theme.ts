/**
 * 主题色（`a:schemeClr`）的解析：把"应用角色"这一层色名解成真正画出来的 `#rrggbb`。
 *
 * 一手依据是 ECMA-376 Part 1 §19.2.1.9（`p:clrMap`：幻灯片配色方案里的角色 → 主题色槽位）与 §7.3.2.28
 * （`a:schemeClr` 的 `@val` 就是这些角色名），加上本机 python-pptx 1.0.2 的落地（`pptx/dml/color.py` 的
 * `theme_color` 只回槽位名，**它不做 clrMap 与 theme 的解色**——所以这档必须我们自己算，没有现成算法可抄）。
 *
 * 红线（`docs/office-parity-spec.md` D3 / `AGENTS.md` §2.2）：
 * - 只解**没有变换子元素**的 `schemeClr`。带 `a:lumMod`/`a:lumOff`/`a:tint`/`a:shade`/`a:alpha` 时回 undefined——
 *   变换的 HSL 语义本机能算但拿不到 PowerPoint 的画面核对，**算错的颜色比没有颜色更坏**（会被当成事实烤回文件）。
 * - `val="phClr"` 是"由继承上下文决定"的占位色，单独一层解不出来，一律 undefined。
 * - 解出来的 hex 只作**只读事实**给视图与模型，绝不写回部件（§2.2 那条"不烤回 slide"）。
 */
import { children, firstChild, localName, parseDocument, readText, type Element, type OpcPackage } from "@office-copilot/opc";
import { relatedPartPath } from "./inherit.js";

/** `a:clrScheme` 里可能出现（也按这个顺序读）的槽位名，`dk1`/`lt1` 是主题自己的叫法。 */
const THEME_SLOTS = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];

/** 一张幻灯片可用的色板：母版的 `p:clrMap` 加上它背后那份主题的 `a:clrScheme`。 */
interface SlideColorMap {
  /** 应用角色（`schemeClr @val`）→ 主题槽位名；`p:clrMap` 没写这条属性时按同名回退（`accent1` → `accent1`）。 */
  slotOfRole: Record<string, string>;
  /** 主题槽位名 → `#rrggbb`；解不出（缺部件、既没有 `srgbClr` 也没有 `lastClr`）的槽位不在表里。 */
  hexOfSlot: Record<string, string>;
}

/** {@link createColorInheritance} 的返回值：一个包一份，按幻灯片缓存色板。 */
export type ColorInheritance = ReturnType<typeof createColorInheritance>;

/**
 * 建一个主题色解析器：一张幻灯片的色板只算一次（一份 deck 几十页共用同一母版与主题）。
 *
 * @param pkg 已 `loadPackage` 的包。
 * @returns {@link ColorInheritance}：给它幻灯片部件路径与 `schemeClr @val`，它回 hex 或 undefined。
 */
export function createColorInheritance(pkg: OpcPackage) {
  const cache = new Map<string, Promise<SlideColorMap>>();
  return {
    /**
     * 求某个 `schemeClr @val` 在这张幻灯片上实际是哪一色。
     *
     * @param slidePartPath 幻灯片部件路径（从这里沿 `/slideLayout` → `/slideMaster` → `/theme` 找色板）。
     * @param role `a:schemeClr @val` 的原样字符串（如 `tx1` / `accent1` / `phClr`）。
     * @returns `#rrggbb`（小写）；链路任一层缺失、或这个 `val` 根本不该由我们解时 undefined（**不许猜**）。
     */
    async resolve(slidePartPath: string, role: string): Promise<string | undefined> {
      if (role.toLowerCase() === "phclr") return undefined;
      let pending = cache.get(slidePartPath);
      if (!pending) {
        pending = readColorMap(pkg, slidePartPath);
        cache.set(slidePartPath, pending);
      }
      const colorMap = await pending;
      const slot = colorMap.slotOfRole[role] ?? role;
      return colorMap.hexOfSlot[slot];
    },
  };
}

/**
 * 沿 幻灯片 → 版式 → 母版 → 主题 读出色板。
 *
 * @param pkg 已加载的包。
 * @param slidePartPath 幻灯片部件路径。
 * @returns 角色→槽位与槽位→hex 两张表；查不到母版或主题时两张都是空的（调用方拿到 undefined，即"这色我们没解出来"）。
 */
async function readColorMap(pkg: OpcPackage, slidePartPath: string): Promise<SlideColorMap> {
  const layoutPart = relatedPartPath(pkg, slidePartPath, "/slideLayout");
  const masterPart = layoutPart ? relatedPartPath(pkg, layoutPart, "/slideMaster") : undefined;
  if (!masterPart) return { slotOfRole: {}, hexOfSlot: {} };
  const themePart = relatedPartPath(pkg, masterPart, "/theme");
  const clrMap = await readClrMap(pkg, masterPart);
  const hexOfSlot = themePart ? await readThemeColors(pkg, themePart) : {};
  return { slotOfRole: clrMap, hexOfSlot };
}

/**
 * 读母版的 `p:clrMap`（属性名是应用角色，值是主题槽位）。
 *
 * @param pkg 已加载的包。
 * @param masterPart 母版部件路径。
 * @returns 角色→槽位；没有这张表或读不动时空对象（后面按同名回退）。
 */
async function readClrMap(pkg: OpcPackage, masterPart: string): Promise<Record<string, string>> {
  const xmlText = await readText(pkg.zip, masterPart);
  if (xmlText === undefined) return {};
  let root: Element;
  try {
    root = parseDocument(xmlText);
  } catch {
    return {};
  }
  const clrMap = firstChild(root, "p:clrMap");
  if (!clrMap) return {};
  const slotOfRole: Record<string, string> = {};
  for (const [attribute, value] of Object.entries(clrMap.attributes)) {
    slotOfRole[attribute] = value;
  }
  return slotOfRole;
}

/**
 * 读主题的 `a:clrScheme`：`srgbClr @val` 直接要，`sysClr` 只认 `@lastClr`（办公软件上一次画出来的那色）。
 *
 * @param pkg 已加载的包。
 * @param themePart 主题部件路径。
 * @returns 槽位→`#rrggbb`；只有 `THEME_SLOTS` 里那十二个槽位会进来，缺一个就不在表里。
 */
async function readThemeColors(pkg: OpcPackage, themePart: string): Promise<Record<string, string>> {
  const xmlText = await readText(pkg.zip, themePart);
  if (xmlText === undefined) return {};
  let root: Element;
  try {
    root = parseDocument(xmlText);
  } catch {
    return {};
  }
  const elements = firstChild(root, "a:themeElements");
  const clrScheme = elements ? firstChild(elements, "a:clrScheme") : undefined;
  if (!clrScheme) return {};
  const hexOfSlot: Record<string, string> = {};
  for (const entry of children(clrScheme)) {
    const slot = localName(entry.name);
    if (!THEME_SLOTS.includes(slot)) continue;
    const inside = children(entry);
    const srgb = inside.find((child) => localName(child.name) === "srgbClr");
    const sys = inside.find((child) => localName(child.name) === "sysClr");
    const raw = srgb?.attributes["val"] ?? sys?.attributes["lastClr"];
    if (raw && /^[0-9a-fA-F]{6}$/.test(raw)) hexOfSlot[slot] = `#${raw.toLowerCase()}`;
  }
  return hexOfSlot;
}
