import { stylePropertiesXml } from "./ooxml-pm.js";
import type { PMMark, PPrAttrs } from "./pm-schema.js";
import {
  children,
  escapeAttr,
  firstChild,
  innerXml,
  parseDocument,
  parseFragment,
  serialize,
} from "./xml.js";
import type { Element } from "./xml.js";

/**
 * 一次"新建样式 / 修改样式"的编辑：装的是**编辑器里抓到的格式**，还没有换算成 OOXML。
 *
 * 换算发生在 `stylePropertiesXml`（也就是 `docx-io` 内部），所以渲染层与主进程都只管
 * 搬运这份结构化数据，单位仍然只在一个地方转换（AGENTS §2.1）。
 */
export interface StyleEdit {
  /** undefined = 新建（id 由 `nextStyleId` 分配）；字符串 = 改这条已声明的样式。 */
  styleId?: string;
  /** `w:name`：窗格与 Word 样式列表里显示的名字，可以是中文。 */
  name: string;
  /** 段落 attrs（段前/缩进是 pt，lineHeight 是倍数或 pt:N / atLeast:N）。 */
  paragraph: PPrAttrs;
  /** 选区那段文字的行内标记（字号是 CSS px、颜色是 #rrggbb）。 */
  marks: PMMark[];
  /** 只在新建时用：父样式 id（Word 对话框里的"基准样式"）。 */
  basedOn?: string;
}

/** 能安全写进 @w:styleId、也安全进 CSS 属性选择器的 id 形状。 */
const SAFE_STYLE_ID = /^[A-Za-z0-9][A-Za-z0-9_-]{0,63}$/;

/** 新建的样式都带上它：ECMA-376 用它区分"这份文档自己定义的用户样式"与内建样式。 */
const CUSTOM_STYLE_ATTR = ' w:customStyle="1"';

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/**
 * 给新建的样式分配一个不与本文档冲突的 id。
 *
 * Word 的做法是把样式名里的空格与标点去掉（"Overdue Amount Para" → `OverdueAmountPara`），
 * 本函数沿用这一规则；中文名去掉非 ASCII 后什么都不剩，退回 `Style` 前缀——id 只要稳定
 * 唯一即可，界面显示用的是 `w:name`。
 *
 * @param taken 已经被占用的 id（本文档声明过的 + 这一轮已经分配过的）。
 * @param name 用户起的样式名。
 * @returns 合 `SAFE_STYLE_ID` 且不在 taken 里的 id。
 */
export function nextStyleId(taken: ReadonlySet<string>, name: string): string {
  const slugged = name.replace(/[^A-Za-z0-9_-]/g, "").slice(0, 56);
  const base = /^[A-Za-z]/.test(slugged) ? slugged : "Style";
  if (!taken.has(base)) return base;
  let index = 1;
  while (taken.has(`${base}${index}`)) index += 1;
  return `${base}${index}`;
}

/** 元素开标签，属性按原顺序写出并转义。 */
function openTag(el: Element): string {
  const attrs = Object.entries(el.attributes ?? {})
    .map(([name, value]) => ` ${name}="${escapeAttr(String(value))}"`)
    .join("");
  return `<${el.name}${attrs}>`;
}

/** 取 `w:style` 里某个属性块的内部 XML；没有这个子元素时返回 undefined。 */
function innerOf(el: Element, name: string): string | undefined {
  const child = firstChild(el, name);
  return child ? innerXml(child) : undefined;
}

/**
 * 把样式属性块拼进 `w:style` 的尾部（CT_Style 规定 pPr / rPr 在所有其他子元素之后）。
 *
 * @param parts 已有的子元素串。
 * @param props 换算好的 pPr / rPr 内部 XML；空串表示这条样式不带该属性块。
 * @returns 追加完的子元素串。
 */
function withPropertyBlocks(
  parts: string[],
  props: { pPrInner: string; rPrInner: string },
): string[] {
  if (props.pPrInner) parts.push(`<w:pPr>${props.pPrInner}</w:pPr>`);
  if (props.rPrInner) parts.push(`<w:rPr>${props.rPrInner}</w:rPr>`);
  return parts;
}

/**
 * 把一条**新**样式拼成 `w:style`。
 *
 * @param edit 这次编辑，`styleId` 与 `name` 已经定稿。
 * @param props 换算好的 pPr / rPr 内部 XML。
 * @param declared 样式表里已声明的 id；`basedOn` 指向它们之外的一律不写——Word 对
 *   没声明的引用是按默认样式排的，与其留一条指向空气的引用，不如干脆不写。
 * @returns 完整的 `w:style` 串。子元素顺序按 CT_Style（§17.5.2.27）：
 *   `w:name` → `w:basedOn` → `w:pPr` → `w:rPr`；Word 校验的是先后而不是元素齐不齐，
 *   顺序错会弹"文档内容有问题"。
 */
function newStyleXml(
  edit: RequiredPick<StyleEdit, "styleId">,
  props: StyleProps,
  declared: ReadonlySet<string>,
): string {
  const parts = [`<w:name w:val="${escapeAttr(edit.name)}"/>`];
  if (edit.basedOn && SAFE_STYLE_ID.test(edit.basedOn) && declared.has(edit.basedOn)) {
    parts.push(`<w:basedOn w:val="${escapeAttr(edit.basedOn)}"/>`);
  }
  return `<w:style w:type="paragraph"${CUSTOM_STYLE_ATTR} w:styleId="${escapeAttr(edit.styleId)}">${withPropertyBlocks(parts, props).join("")}</w:style>`;
}

/**
 * 改一条**已存在**的样式：只替换 `w:name`、`w:pPr`、`w:rPr`，其余子元素原样留着。
 *
 * @param el 原 `w:style` 元素（或本轮上一次编辑刚产出的那一条）。
 * @param edit 这次编辑。
 * @param props 与原定义合并后的 pPr / rPr 内部 XML。
 * @returns 重新序列化出来的 `w:style` 串。@w:type、@w:default、`w:basedOn`、`w:qFormat`、
 *   `w:uiPriority` 这些"从选区抓格式"覆盖不到的东西一律不动。
 */
function restyledXml(el: Element, edit: StyleEdit, props: StyleProps): string {
  // 名字挪到最前（CT_Style 要求），其余保留原有相对次序，属性块永远排在最后。
  const kept = children(el)
    .filter((c) => c.name !== "w:name" && c.name !== "w:pPr" && c.name !== "w:rPr")
    .map(serialize);
  const parts = [`<w:name w:val="${escapeAttr(edit.name)}"/>`, ...kept];
  return `${openTag(el)}${withPropertyBlocks(parts, props).join("")}</w:style>`;
}

/** 只把 styleId 收紧成必填，其余字段沿用 StyleEdit。 */
type RequiredPick<T, K extends keyof T> = Omit<T, K> & Required<Pick<T, K>>;

/** `stylePropertiesXml` 的返回形状。 */
type StyleProps = { pPrInner: string; rPrInner: string };

/**
 * 把样式编辑并进 `word/styles.xml`，splice-and-keep。
 *
 * 保住的东西：根元素上的全部命名空间声明、`w:docDefaults`、`w:latentStyles`、
 * 以及每一条**没被改动**的 `w:style`——样式链解析（`style-format.ts`）依赖
 * `w:docDefaults`，写侧若整份重建就会把全文档的默认格式抹掉。
 *
 * @param originalXml 原来的 styles.xml；undefined 或读不懂时按"这份文档还没有样式表"
 *   从头建根元素，编辑内容照常写出，不会静默丢。
 * @param edits 本会话累计的新建与修改，按用户的操作顺序；同一 id 改多次时后一次在
 *   前一次的结果上继续合并（不会丢掉头一次写进去的属性）。
 * @returns 新的 styles.xml 全文；一条有效编辑都没有时原样返回 `originalXml`
 *   （可能是 undefined，含义是"这次保存不该动 styles.xml"）。
 */
export function mergeStylesXml(
  originalXml: string | undefined,
  edits: readonly StyleEdit[],
): string | undefined {
  if (!edits.length) return originalXml;
  let root: Element | undefined;
  if (originalXml) {
    try {
      root = parseDocument(originalXml);
    } catch {
      root = undefined;
    }
  }
  const entries = root ? children(root).filter((c) => c.name === "w:style") : [];
  const taken = new Set(entries.map((el) => String(el.attributes?.["w:styleId"] ?? "")));
  const byId = new Map(entries.map((el) => [String(el.attributes?.["w:styleId"] ?? ""), el]));
  const replaced = new Map<string, string>();
  const fresh: string[] = [];
  let applied = false;

  for (const edit of edits) {
    const name = edit.name.trim();
    if (!name) continue;
    const target = (edit.styleId && byId.get(edit.styleId)) ?? undefined;
    const styleId = edit.styleId ?? nextStyleId(taken, name);
    const props = stylePropertiesXml(
      { paragraph: edit.paragraph, marks: edit.marks },
      target
        ? { pPrInner: innerOf(target, "w:pPr"), rPrInner: innerOf(target, "w:rPr") }
        : {},
    );
    const xml = target
      ? restyledXml(target, { ...edit, name, styleId }, props)
      : newStyleXml({ ...edit, name, styleId }, props, taken);
    if (target) {
      replaced.set(styleId, xml);
      // 同一轮里再改这条时，基准是刚刚产出的定义，而不是磁盘上那份旧的。
      byId.set(styleId, parseFragment(xml)[0] as Element);
    } else {
      fresh.push(xml);
    }
    taken.add(styleId);
    applied = true;
  }

  if (!applied) return originalXml;

  const body: string[] = [];
  for (const el of root ? children(root) : []) {
    if (el.name !== "w:style") {
      body.push(serialize(el));
      continue;
    }
    body.push(replaced.get(String(el.attributes?.["w:styleId"] ?? "")) ?? serialize(el));
  }
  body.push(...fresh);
  const open = root ? openTag(root) : `<w:styles xmlns:w="${W_NS}">`;
  return `${XML_DECL}\n${open}${body.join("")}</w:styles>`;
}
