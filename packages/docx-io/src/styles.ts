import {
  children,
  firstChild,
  onOffValue,
  parseDocument,
  serialize,
  val,
} from "./xml.js";
import type { Element } from "./xml.js";

/** Twentieths of a point in one CSS px, the inverse of docx.ts's TWIP_TO_PX. */
const TWIPS_PER_PX = 15;

/**
 * `w:style` 的 @w:type 取值（ECMA-376 §17.5.2.27 列出六类：段落、字符、
 * 通过 w:link 联动的"段落+字符"、表格、编号、以及默认的段落/字符属性）。
 */
export type StyleType = "paragraph" | "character" | "numbering" | "table";

/**
 * 一条具名样式的结构化视图（CT_Style）。
 *
 * `pPrInner` / `rPrInner` 保留元素内部原文，**不做单位换算**：样式定义要么原样
 * 透传，要么由写侧自己按 twips 生成，换算点只能有一个（AGENTS §2.1）。
 * `xml` 是整条 `w:style` 重新序列化的结果，没被改动过的样式按它写回，
 * 于是"改一条样式"不会牵连同文件里其他样式的属性顺序或自闭合写法。
 */
export interface StyleDefinition {
  /** @w:styleId，正文里 `w:pStyle` / `w:rStyle` 引用的就是它。 */
  styleId: string;
  /** `w:name`：Word 的内建名（"Normal"、"heading 1"）或用户起的名字。 */
  name?: string;
  type: StyleType;
  /** 样式链的父样式 id；父样式自身的格式在这条之前生效。 */
  basedOn?: string;
  /** 套用这条样式后回车得到的下一条样式（标题的 next 通常是正文）。 */
  next?: string;
  /** 联动的字符样式 id（标题的加粗与它的字符样式是一对）。 */
  link?: string;
  /** @w:default：该类型的默认样式，Word 把它当作"清除格式"的落点。 */
  isDefault: boolean;
  /** `w:hidden`：样式不出现在窗格里，但引用它的正文照样生效。 */
  hidden: boolean;
  /** `w:semiHidden`：Word 只在样式窗格勾掉"隐藏样式"时才列它。 */
  semiHidden: boolean;
  /** `w:uiPriority`：窗格排序依据，数字小的在前；文档没写时为 undefined。 */
  uiPriority?: number;
  /** 这条样式自己的 `<w:pPr>` 内部 XML；没有段落属性时 undefined。 */
  pPrInner?: string;
  /** 这条样式自己的 `<w:rPr>` 内部 XML；没有字符属性时 undefined。 */
  rPrInner?: string;
  /** 整条 `w:style` 的原样序列化，未改动时按它写回。 */
  xml: string;
}

/**
 * 把一条 `w:style` 读成 StyleDefinition。
 *
 * @param el 已解析的 `w:style` 元素。
 * @param styleId 已确认非空的 @w:styleId。
 * @returns 结构化定义。`w:name` / `w:basedOn` 等缺失时对应字段是 undefined，
 *   不替 Word 补默认值。
 */
function readStyle(el: Element, styleId: string): StyleDefinition {
  const pPr = firstChild(el, "w:pPr");
  const rPr = firstChild(el, "w:rPr");
  const priority = Number(val(firstChild(el, "w:uiPriority")));
  return {
    styleId,
    name: val(firstChild(el, "w:name")),
    type: styleType(el),
    basedOn: val(firstChild(el, "w:basedOn")),
    next: val(firstChild(el, "w:next")),
    link: val(firstChild(el, "w:link")),
    isDefault: attrOnOff(el, "w:default"),
    hidden: onOffValue(firstChild(el, "w:hidden")) === true,
    semiHidden: onOffValue(firstChild(el, "w:semiHidden")) === true,
    uiPriority: Number.isFinite(priority) ? priority : undefined,
    pPrInner: pPr ? innerXml(pPr) : undefined,
    rPrInner: rPr ? innerXml(rPr) : undefined,
    xml: serialize(el),
  };
}

/**
 * 读 @w:type。
 *
 * @param el `w:style` 元素。
 * @returns 四类之一。schema 里 @w:type 是必填的，所以缺失或本仓不认识的写法统一
 *   按段落样式返回——按未知类型处理的话整条样式读不了，反而更难用。
 */
function styleType(el: Element): StyleType {
  switch (String(el.attributes?.["w:type"] ?? "")) {
    case "character":
      return "character";
    case "numbering":
      return "numbering";
    case "table":
      return "table";
    default:
      return "paragraph";
  }
}

/**
 * 读 CT_OnOff 型的**属性**（@w:default 是属性，不是子元素，所以不能走 onOffValue）。
 *
 * @param el 承载属性的元素。
 * @param name 属性名，如 "w:default"。
 * @returns 没写这个属性即 false；显式的 0/false/off 也是 false，其余写法为 true。
 */
function attrOnOff(el: Element, name: string): boolean {
  const raw = el.attributes?.[name];
  if (raw === undefined) return false;
  const v = String(raw).toLowerCase();
  return !(v === "0" || v === "false" || v === "off");
}

/** 元素内部 XML（不含外层标签），用于把一段 pPr/rPr 原样带走。 */
function innerXml(el: Element): string {
  return children(el)
    .map(serialize)
    .join("");
}

/**
 * 样式窗格里的一条样式：只带"列出来 + 点一下"所需的字段，不含任何 XML，
 * 所以可以原样通过 Electron 的 IPC 送到渲染层（渲染层拿不到 styles.xml 全文）。
 */
export interface PaneStyle {
  styleId: string;
  /** `w:name` 原值。内建样式是英文（"heading 1"），界面要显示的中文名由渲染层映射。 */
  name: string;
  /** 这一级的标题样式（1–6）；不是标题样式时 undefined。套用它时段落要同时变成这一级的标题。 */
  headingLevel?: number;
  /** @w:default：该类型的默认样式。Word 的"清除样式"就是把段落退回这条。 */
  isDefault: boolean;
}

/**
 * `fitStyleId` 需要的最小查询面。`StyleMap` 天生满足它；渲染层只拿到窗格清单
 * （`PaneStyle[]`），用它拼出同一个接口，两处走的就是同一条判定规则。
 */
export interface HeadingStyleLookup {
  /** 该样式在 styles.xml 里声明的标题级别；不是标题样式时 undefined。 */
  headingLevel(styleId: string): number | undefined;
  /** 本文档为某一级标题声明的样式 id；没声明时 undefined（不许凭空造 id）。 */
  styleIdForHeadingLevel(level: number): string | undefined;
}

/**
 * 判定一段现在该对外承认哪个具名样式 id。
 *
 * 编辑器里改标题级别、把标题转成正文时只动节点的 type / level，styleId 留在 attrs
 * 里没人负责跟着改——照原样写回的话 Word 仍按旧的那一级排版，用户的操作等于没生效。
 *
 * @param styleId attrs 里的具名样式 id：字符串=套用，null=用户清除过，undefined=没建模过。
 * @param headingLevel 段落当前的标题级别；正文传 undefined。
 * @param styles 本文档的样式表视图；没有它就没法判断，只能放行。
 * @returns 写侧与窗格都该用的 id。标题换了级别时返回这一级真正声明的 id，样式表没有
 *   这一级就返回 undefined（由调用方按节点类型推断）；标题被改成正文返回 null（引用
 *   必须删掉）；不是标题样式、或用户显式清除过样式时原样返回。
 */
export function fitStyleId(
  styleId: string | null | undefined,
  headingLevel: number | undefined,
  styles: HeadingStyleLookup | undefined,
): string | null | undefined {
  if (!styles) return styleId;
  if (typeof styleId === "string") {
    const declaredLevel = styles.headingLevel(styleId);
    if (declaredLevel === undefined || headingLevel === declaredLevel) return styleId;
    return headingLevel === undefined ? null : styles.styleIdForHeadingLevel(headingLevel);
  }
  // null = 用户显式清除过样式，不能替他改主意。
  if (styleId === null) return null;
  // 没建模过：正文保持源文档写法；标题在编辑器里新建时天生没有引用，
  // 只能从样式表取这一级声明的那条，否则会写出本文档不存在的样式 id。
  return headingLevel === undefined ? undefined : styles.styleIdForHeadingLevel(headingLevel);
}

/**
 * 把窗格清单变成 `fitStyleId` 需要的查询面。渲染层只拿到 `PaneStyle[]`（IPC 里
 * 没有 styles.xml），有了它就不必自己再实现一遍"级别变了样式 id 该怎么改"。
 *
 * @param styles `StyleMap.paneStyles()` 的结果。
 * @returns 只认清单里那几条样式的双向查找；同一级出现多条时取排在最前的。
 */
export function headingLookupFromPaneStyles(styles: PaneStyle[]): HeadingStyleLookup {
  const levelById = new Map<string, number>();
  const idByLevel = new Map<number, string>();
  for (const style of styles) {
    if (style.headingLevel === undefined) continue;
    levelById.set(style.styleId, style.headingLevel);
    if (!idByLevel.has(style.headingLevel)) idByLevel.set(style.headingLevel, style.styleId);
  }
  return {
    headingLevel: (styleId) => levelById.get(styleId),
    styleIdForHeadingLevel: (level) => idByLevel.get(level),
  };
}

/**
 * Parsed lookups from word/styles.xml. Real-world .docx files use arbitrary
 * numeric styleIds (WPS writes heading 1 with styleId="2"), so nothing here may
 * assume Word's own ids; and the style pane needs every declared style, not just
 * the headings, which is why each `w:style` is also kept as a StyleDefinition.
 */
export class StyleMap {
  /** styleId → heading level (1..6). Missing key means not a heading. */
  private headingById = new Map<string, number>();
  /** Every styleId styles.xml actually declares, so unknown ids can be guessed. */
  private declared = new Set<string>();
  /** styleId → true if this is a list-related style (ListBullet/ListNumber). */
  listById = new Set<string>();
  /** 按 styles.xml 里的出现顺序保存全部样式定义，窗格与写侧都用它。 */
  private defs: StyleDefinition[] = [];
  /**
   * Space below a paragraph the document leaves unspecified, in CSS px. Word
   * takes it from the default paragraph style, and a template whose Normal
   * carries no `w:spacing` puts body text edge to edge — the flow height that
   * floating boxes anchored to those paragraphs are measured against.
   */
  defaultSpacingAfterPx = 0;
  /**
   * `w:docDefaults/w:pPrDefault/w:pPr` 的内部 XML。ISO/IEC 29500-1 §17.7.2.5 说这些
   * 属性"applied first in the style hierarchy"，所以样式链解析要把它垫在最底下。
   */
  docDefaultsPPrInner?: string;
  /** `w:docDefaults/w:rPrDefault/w:rPr` 的内部 XML，同上。 */
  docDefaultsRPrInner?: string;
  /**
   * @w:default 的段落样式 id。Word 把它当作"没有 `w:pStyle` 的段落"与"清除样式"的落点，
   * 真实文档里通常是 Normal——但 id 由文档自己决定，这里不假设它叫 Normal。
   */
  defaultParagraphStyleId?: string;

  static fromXml(xml: string | undefined | null): StyleMap {
    const map = new StyleMap();
    if (!xml) return map;
    try {
      const root = parseDocument(xml);
      for (const s of children(root).filter((c) => c.name === "w:style")) {
        const id = s.attributes?.["w:styleId"] as string | undefined;
        if (!id) continue;
        const def = readStyle(s, id);
        map.declared.add(id);
        map.defs.push(def);
        const name = def.name?.toLowerCase() ?? "";
        const m = name.match(/^heading\s+([1-6])/);
        if (m) {
          map.headingById.set(id, Number(m[1]));
          continue;
        }
        if (name.includes("list bullet") || name.includes("list number")) {
          map.listById.add(id);
        }
        if (def.type === "paragraph" && def.isDefault) {
          map.defaultParagraphStyleId = id;
          const pPr = firstChild(s, "w:pPr");
          const spacing = pPr && firstChild(pPr, "w:spacing");
          // `w:after` counts twentieths of a point; 15 of them make a CSS px.
          map.defaultSpacingAfterPx =
            Number(spacing?.attributes?.["w:after"] ?? 0) / TWIPS_PER_PX;
        }
      }
      const defaults = firstChild(root, "w:docDefaults");
      const pPrDefault = defaults && firstChild(defaults, "w:pPrDefault");
      const rPrDefault = defaults && firstChild(defaults, "w:rPrDefault");
      const defaultsPPr = pPrDefault && firstChild(pPrDefault, "w:pPr");
      const defaultsRPr = rPrDefault && firstChild(rPrDefault, "w:rPr");
      if (defaultsPPr) map.docDefaultsPPrInner = innerXml(defaultsPPr);
      if (defaultsRPr) map.docDefaultsRPrInner = innerXml(defaultsRPr);
    } catch {
      // Ignore malformed styles.
    }
    return map;
  }

  headingLevel(styleId: string | undefined): number | undefined {
    if (!styleId) return undefined;
    return this.headingById.get(styleId);
  }

  /**
   * 某一级的标题样式 id，`headingLevel` 的反向查找。
   *
   * @param level 标题级别 1–6。
   * @returns 文档里这一级的标题样式 id；同一级声明了多条时取 styles.xml 里最先出现
   *   的那条。文档没有这一级的标题样式时返回 undefined——不能凭空造一个没声明的 id，
   *   Word 对未定义的引用是按默认样式排的，段落会直接丢掉标题格式。
   */
  styleIdForHeadingLevel(level: number): string | undefined {
    for (const def of this.defs) {
      if (this.headingById.get(def.styleId) === level) return def.styleId;
    }
    return undefined;
  }

  /** True when styles.xml names this style, so its role is known, not guessed. */
  isDeclared(styleId: string | undefined): boolean {
    return styleId !== undefined && this.declared.has(styleId);
  }

  /**
   * 文档声明的全部样式，按 styles.xml 里的顺序。
   *
   * @returns 只读数组；空数组表示这份文档没有 styles.xml，或一条样式都没声明。
   */
  definitions(): readonly StyleDefinition[] {
    return this.defs;
  }

  /**
   * 按 @w:styleId 取一条样式定义。
   *
   * @param styleId 正文引用的样式 id。
   * @returns 定义本身；文档没声明这个 id 时返回 undefined。调用方不得把它当作
   *   "无格式"——Word 对未定义 id 是按默认样式排的，不是清零。
   */
  byId(styleId: string | undefined): StyleDefinition | undefined {
    if (!styleId) return undefined;
    return this.defs.find((d) => d.styleId === styleId);
  }

  /**
   * 样式链：自身在前，然后沿 `w:basedOn` 往上。
   *
   * @param styleId 起点样式 id。
   * @returns 链上的定义，按"自身 → 父 → 祖父"排列；id 未声明时返回空数组。
   *   `w:basedOn` 成环（WPS 文档里真出现过两条互指）时在第一次重复处截断。
   */
  chain(styleId: string | undefined): StyleDefinition[] {
    const out: StyleDefinition[] = [];
    const seen = new Set<string>();
    let current = this.byId(styleId);
    while (current && !seen.has(current.styleId)) {
      seen.add(current.styleId);
      out.push(current);
      current = this.byId(current.basedOn);
    }
    return out;
  }

  /**
   * 样式窗格要列出的段落样式。
   *
   * @returns 按 `w:uiPriority` 升序（没写优先级的排到最后，同级保持 styles.xml 里的
   *   声明次序）、已过滤 `w:hidden` / `w:semiHidden` 与字符/表格/编号样式的列表。
   *   空数组表示这份文档没有可显示的段落样式（例如根本没有 styles.xml）。
   *   被过滤掉的样式只是不出现在窗格里，引用它们的正文不受影响——那是 `w:hidden`
   *   的语义，不等于"无格式"。
   */
  paneStyles(): PaneStyle[] {
    const priority = (d: StyleDefinition): number => d.uiPriority ?? Number.MAX_SAFE_INTEGER;
    return this.defs
      .filter((d) => d.type === "paragraph" && !d.hidden && !d.semiHidden)
      // Array.prototype.sort 是稳定排序：同优先级（含都没写优先级）时保持声明次序。
      .sort((a, b) => priority(a) - priority(b))
      .map((d) => {
        const level = this.headingLevel(d.styleId);
        return {
          styleId: d.styleId,
          name: d.name ?? d.styleId,
          ...(level !== undefined ? { headingLevel: level } : {}),
          isDefault: d.isDefault,
        };
      });
  }
}
