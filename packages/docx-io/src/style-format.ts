/**
 * 样式链解析成可见格式（`docs/word-parity-spec.md` 排名 2 ⑤）。
 *
 * Word 的排版优先级有明文规定：ISO/IEC 29500-1 §17.7.2.5 说 `w:docDefaults` 的属性
 * "applied first in the style hierarchy; therefore they are superseded by any further
 * conflicting formatting"。于是顺序是 docDefaults → `w:basedOn` 链（根在前）→ 这条样式
 * 自己 → 段落与文字的直接格式。本模块只做**前三层**的合并，产出的还是编辑器那套词汇
 * （段间距/缩进是 pt，行距是倍数或 `pt:N`/`atLeast:N`，字号/字距是 CSS px，颜色 #rrggbb），
 * 由渲染层拼成文档作用域的样式表。
 *
 * 刻意不在导入时把解析结果写进段落 attrs：那等于把样式定义盖成用户从没做过的直接格式，
 * 导出后就再也分不清哪部分是用户的（AGENTS §2.3）。直接格式留在 attrs 里，靠 CSS 优先级
 * （行内 style 胜过样式表规则）继续在 Word 语义的位置上赢。
 */
import {
  jcToTextAlign,
  readParagraphFormat,
  readRunProperties,
  type RunStyleProps,
} from "./ooxml-pm.js";
import type { PPrAttrs } from "./pm-schema.js";
import { StyleMap } from "./styles.js";
import { firstChild, parseFragment, val } from "./xml.js";
import type { Element } from "./xml.js";

/** 本模块建模的段落属性；单位与 `PPrAttrs` 同名项一致。 */
type ResolvedParagraphFormat = Partial<
  Pick<
    PPrAttrs,
    "textAlign" | "spacingBefore" | "spacingAfter" | "indent" | "indentRight" | "indentFirst" | "lineHeight"
  >
>;

/** 一条样式解析后的可见格式：段落部分 + 字符部分（`w:rPr` 直接子元素，不是段落标记的那份）。 */
export type ResolvedStyleFormat = ResolvedParagraphFormat & RunStyleProps;

/**
 * 样式表里一条段落样式的可见格式。
 *
 * `styleId` 为 null 的那条是**文档默认段落样式**：Word 把 docDefaults + `@w:default`
 * 的段落样式给所有没写 `w:pStyle` 的段落，也就是"清除样式"之后的落点。
 */
export interface ResolvedStyle {
  styleId: string | null;
  /** 这一条的格式；只含样式链真写过的键。 */
  format: ResolvedStyleFormat;
}

/**
 * 把 `<w:pPr>` / `<w:rPr>` 的内部原文还原成元素。
 *
 * @param tag 元素名（含 `w:` 前缀）。
 * @param inner 元素内部 XML（`StyleDefinition.pPrInner` 那种原文）；空白表示这一层没写这组属性。
 * @returns 可直接交给读侧函数的元素；这一层为空时返回 undefined。
 *   本仓的 XML 解析按字面前缀名匹配，`w14:` / `mc:` 这类外部前缀的子元素照常带着，
 *   不会因为没声明命名空间而报错（实测：`parseFragment('<w:pPr><w14:textFill/></w:pPr>')` 可用）。
 */
function wrapProps(tag: string, inner: string | undefined): Element | undefined {
  if (!inner || !inner.trim()) return undefined;
  return parseFragment(`<${tag}>${inner}</${tag}>`)[0];
}

/**
 * 读一层段落属性。
 *
 * @param inner `<w:pPr>` 的内部 XML。
 * @returns 只带这一层真写过的键——合并用的是对象展开，留着值为 undefined 的键会把
 *   下一层（优先级更高的那层）读到的值擦掉。段落标记的 `w:pPr/w:rPr` 不在这里读：
 *   那是 Word 给"段落标记本身"的格式，不影响正文文字。
 */
function paragraphPart(inner: string | undefined): ResolvedParagraphFormat {
  const pPr = wrapProps("w:pPr", inner);
  if (!pPr) return {};
  const read = readParagraphFormat(pPr);
  // w:jc 走正文读侧的同一个映射函数，否则同一份对齐在样式里和正文里会读出不同结果。
  const textAlign = jcToTextAlign(val(firstChild(pPr, "w:jc")));
  const part: ResolvedParagraphFormat = {};
  if (textAlign) part.textAlign = textAlign;
  if (read.spacingBefore !== undefined) part.spacingBefore = read.spacingBefore;
  if (read.spacingAfter !== undefined) part.spacingAfter = read.spacingAfter;
  if (read.indent !== undefined) part.indent = read.indent;
  if (read.indentRight !== undefined) part.indentRight = read.indentRight;
  if (read.indentFirst !== undefined) part.indentFirst = read.indentFirst;
  if (read.lineHeight !== undefined) part.lineHeight = read.lineHeight;
  return part;
}

/**
 * 读一层字符属性（`w:style` 的直接子 `w:rPr`）。
 *
 * @param inner `<w:rPr>` 的内部 XML。
 * @returns 见 `readRunProperties`：只带写过的键，开关保留三态，
 *   所以 `<w:b w:val="0"/>` 这种"显式取消继承来的加粗"不会被当成没提过。
 */
function runPart(inner: string | undefined): RunStyleProps {
  return readRunProperties(wrapProps("w:rPr", inner));
}

/**
 * 把样式链上的一层（段落属性 + 字符属性）读成一个格式片段。
 *
 * @param def 一条样式定义的 pPr/rPr 原文；docDefaults 没有 `w:style` 外壳，
 *   调用方把它的两份内部原文按同形对象递进来。
 * @returns 这一层写过的属性。空定义（`{}`）读出空对象，展开时什么都不覆盖。
 */
function layerOf(def: { pPrInner?: string; rPrInner?: string }): ResolvedStyleFormat {
  return { ...paragraphPart(def.pPrInner), ...runPart(def.rPrInner) };
}

/**
 * 按 Word 的优先级合并一条样式链。
 *
 * @param styles 这份文档的样式表视图；docDefaults 也从它取。
 * @param chainFor 要算哪条样式；undefined 表示"没有具名引用"，只算 docDefaults + 默认段落样式。
 * @returns 合并后的格式，低优先级在前：docDefaults →（无引用或引用未声明时的）默认段落样式 →
 *   链上由根到自身。引用了本文档没声明的 id 时按 Word 的行为退回默认样式，而不是当"无格式"
 *   （与 `StyleMap.byId` 的告诫一致）。
 */
function mergeChain(styles: StyleMap, chainFor: string | undefined): ResolvedStyleFormat {
  const layers: ResolvedStyleFormat[] = [
    layerOf({ pPrInner: styles.docDefaultsPPrInner, rPrInner: styles.docDefaultsRPrInner }),
  ];
  const declared = chainFor !== undefined && styles.isDeclared(chainFor);
  if (!declared) layers.push(layerOf(styles.byId(styles.defaultParagraphStyleId) ?? {}));
  if (declared) {
    // chain() 是"自身 → 父 → 祖父"，合并得从最外层开始，所以先倒过来。
    for (const def of [...styles.chain(chainFor)].reverse()) layers.push(layerOf(def));
  }
  const out: ResolvedStyleFormat = {};
  for (const layer of layers) Object.assign(out, layer);
  return out;
}

/**
 * 把整份样式表解析成"每条段落样式一条可见格式"。
 *
 * @param styles `StyleMap.fromXml` 的结果；没有 styles.xml 的文档（新建的空白文档）
 *   得到空数组，渲染层因此什么都不注入。
 * @returns 顺序是：文档默认段落样式（`styleId: null`）在最前，然后是 styles.xml 里
 *   声明顺序的每一条段落样式。`w:hidden` 的样式**照旧解析**——那只影响窗格列不列它，
 *   引用它的正文仍然按它排版。链上一个属性都没写过的样式不出现（免得产出空规则）。
 */
export function resolvedStyles(styles: StyleMap): ResolvedStyle[] {
  const out: ResolvedStyle[] = [];
  const defaultFormat = mergeChain(styles, undefined);
  if (Object.keys(defaultFormat).length) out.push({ styleId: null, format: defaultFormat });
  for (const def of styles.definitions()) {
    // 字符/表格/编号样式不是段落引用能指向的，正文里的 `w:pStyle` 只认段落样式。
    if (def.type !== "paragraph") continue;
    const format = mergeChain(styles, def.styleId);
    if (Object.keys(format).length) out.push({ styleId: def.styleId, format });
  }
  return out;
}
