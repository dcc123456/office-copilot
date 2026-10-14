/**
 * OOXML 通用 XML 层（从 `packages/docx-io/src/xml.ts` 提上来的一份，不是第二份）。
 *
 * 为什么在这一层：`.docx`/`.xlsx`/`.pptx` 三种部件用的是同一套 XML 形状，
 * 之前通用部分住在 docx-io 里、由 sheet-io 反向依赖去复用（`AGENTS.md` §3.1），
 * 再给 pptx-io 用就是第三次。命名空间前缀一律原样保留（OOXML 各部件都靠 `w:`/`x:`/`a:` 前缀区分），
 * 所以解析时不做前缀消解，序列化时也不重排属性。
 */
import { parseXml, XmlCdata, XmlComment, type XmlElement, type XmlNode, type XmlText } from "@rgrove/parse-xml";

export type Element = XmlElement;
export type TextNode = XmlText;

/** XML 节点的 union（元素 / 文本 / CDATA / 注释）：自己走子树时要按它接 `isElement` 收窄。 */
export type { XmlNode };

/** OOXML 各部件统一的 XML 声明；实测所有部件都带 `standalone="yes"` 且声明后不接换行。 */
export const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

/**
 * 解析一段 XML。
 *
 * @param xml 部件正文。未定义实体、CDATA 与注释都原样留着，避免往返时把用户写的内容改掉。
 * @returns 根元素；没有根元素时抛错，由调用方转成"这份文件读不了"的用户提示。
 */
export function parseDocument(xml: string): Element {
  const doc = parseXml(xml, {
    ignoreUndefinedEntities: true,
    preserveCdata: true,
    preserveComments: true,
  });
  if (!doc.root) throw new Error("XML 文档没有根元素。");
  return doc.root;
}

/**
 * 子元素（只含元素节点，文本/注释跳过）。
 *
 * @param el 父元素。
 * @returns 文档顺序的子元素数组。
 */
export function children(el: Element): Element[] {
  return (el.children ?? []).filter(isElement);
}

/**
 * 节点是不是元素。
 *
 * @param n 任意节点，可能是 undefined。
 * @returns 是元素时 true，同时把类型窄化成 `Element`。
 */
export function isElement(n: XmlNode | undefined): n is Element {
  return !!n && n.type === "element";
}

/**
 * 节点是不是文本节点。
 *
 * @param n 任意节点，可能是 undefined。
 * @returns 是文本时 true，同时把类型窄化成 `TextNode`。
 */
export function isText(n: XmlNode | undefined): n is TextNode {
  return !!n && n.type === "text";
}

/**
 * 节点是不是 CDATA 段。
 *
 * `XmlCdata` 是 `XmlText` 的子类，它的 `type` getter 返回 `cdata`，所以 `isText` 天然不会命中它；
 * 但 `@rgrove/parse-xml` 把 `type` 声明成宽 `string`（不是字面量判别式），TS 无法靠 `node.type === "cdata"` 窄化，
 * 只能用 `instanceof` 拿到 `.text`。
 */
export function isCdata(n: XmlNode): n is XmlCdata {
  return n instanceof XmlCdata;
}

/**
 * 节点是不是注释。同上：`type` 不是判别式，用 `instanceof` 才能安全读 `.content`。
 */
export function isComment(n: XmlNode): n is XmlComment {
  return n instanceof XmlComment;
}

/**
 * 按标签名取直接子元素。
 *
 * @param el 父元素。
 * @param name 带前缀的标签名原值（如 `w:p`、`c:ser`），不做命名空间解析。
 * @returns 匹配的子元素，没有则空数组。
 */
export function findChildren(el: Element, name: string): Element[] {
  return children(el).filter((c) => c.name === name);
}

/**
 * 去掉命名空间前缀，只留本地名。
 *
 * @param name 元素名原值（可能是 `p:cSld`，也可能是默认命名空间下的 `cSld`）。
 * @returns 本地名；无冒号时原样返回。各格式的审计都要按本地名比对（`findChildren` 是精确匹配带前缀的原名）。
 */
export function localName(name: string): string {
  const colon = name.indexOf(":");
  return colon < 0 ? name : name.slice(colon + 1);
}

/**
 * 按标签名取第一个直接子元素。
 *
 * @param el 父元素。
 * @param name 带前缀的标签名原值。
 * @returns 找到的元素；没有则 undefined。
 */
export function firstChild(el: Element, name: string): Element | undefined {
  return children(el).find((c) => c.name === name);
}

/**
 * 把元素内所有后代文本拼起来。
 *
 * @param el 元素。
 * @returns 纯文本；元素里没有文本时返回空串（不是 undefined）。
 */
export function text(el: Element): string {
  let out = "";
  const walk = (n: XmlNode) => {
    if (isText(n)) {
      out += n.text;
    } else if (isElement(n)) {
      for (const c of n.children ?? []) walk(c);
    }
  };
  for (const c of el.children ?? []) walk(c);
  return out;
}

/**
 * 收集所有后代里指定标签名的元素，文档顺序。
 *
 * @param el 起点元素（自身不计入）。
 * @param name 带前缀的标签名原值。
 * @param acc 累加器，递归用；外部一般不传。
 * @returns 匹配的元素数组。
 */
export function collectDescendants(el: Element, name: string, acc: Element[] = []): Element[] {
  for (const c of children(el)) {
    if (c.name === name) acc.push(c);
    collectDescendants(c, name, acc);
  }
  return acc;
}

/**
 * 把节点树序列化回 XML 串。
 *
 * 不用通用序列化库：多处只要某个子树（例如整棵 `<w:body>`），而且 OOXML 不在乎自闭合写法，
 * 属性顺序按对象键序原样输出——**属性表是整 map 保留的，读不懂的属性也会原样带回**，
 * 这是"未建模部件不被拍平"（§2.3）的前提。
 * 注释与 CDATA 同样原样带回：解析时特意 `preserveComments`/`preserveCdata` 留了，
 * 序列化若丢掉它们，"我们没改过的部件"就不再逐字一致。
 *
 * @param node 任意节点。
 * @returns XML 串；文本/注释/CDATA/元素之外的节点（如文档声明）返回空串。
 */
export function serialize(node: XmlNode): string {
  if (isCdata(node)) return `<![CDATA[${node.text}]]>`;
  if (isComment(node)) return `<!--${node.content}-->`;
  if (isText(node)) return escapeText(node.text);
  if (!isElement(node)) return "";

  const attrs = Object.entries(node.attributes ?? {})
    .map(([k, v]) => ` ${k}="${escapeAttr(String(v))}"`)
    .join("");

  const inner = (node.children ?? []).map(serialize).join("");
  const selfClosing = inner.length === 0;
  return selfClosing
    ? `<${node.name}${attrs}/>`
    : `<${node.name}${attrs}>${inner}</${node.name}>`;
}

/**
 * 转义文本节点内容。
 *
 * @param s 原文。
 * @returns `&`/`<`/`>` 已转义的串。
 */
export function escapeText(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

/**
 * 转义属性值。
 *
 * @param s 原文。
 * @returns 在 `escapeText` 之外再转义 `"` 与 `'`。
 */
export function escapeAttr(s: string): string {
  return escapeText(s).replace(/"/g, "&quot;").replace(/'/g, "&apos;");
}

/**
 * 元素内部 XML（不含外层标签）。
 *
 * 读侧要"这个元素里面原来写了什么"、写侧要"把原定义的内部串取出来再合并"，
 * 两处是同一件事，所以只留这一份实现。
 *
 * @param el 已解析的元素。
 * @returns 其子元素按顺序序列化的结果；没有子元素时返回空串。
 */
export function innerXml(el: Element): string {
  return children(el)
    .map(serialize)
    .join("");
}
