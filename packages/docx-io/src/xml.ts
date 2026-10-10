/**
 * Minimal XML DOM helpers on top of @rgrove/parse-xml. We keep a tiny
 * namespace-free model because .docx OOXML always uses the same `w:` prefix.
 */
import { parseXml, type XmlElement, type XmlNode, type XmlText } from "@rgrove/parse-xml";

export type Element = XmlElement;
export type TextNode = XmlText;

export function parseDocument(xml: string): Element {
  const doc = parseXml(xml, {
    ignoreUndefinedEntities: true,
    preserveCdata: true,
    preserveComments: true,
  });
  if (!doc.root) throw new Error("XML 文档没有根元素。");
  return doc.root;
}

export function children(el: Element): Element[] {
  return (el.children ?? []).filter(isElement);
}

/**
 * 把 `word/_rels/document.xml.rels` 里的 `@Target` 换算成包内路径。
 *
 * rels 的 Target 按约定相对 `word/` 写（`header1.xml`、`./styles.xml`），也有绝对写法
 * （`/word/foo.xml`），OPC 还允许 `..` 往上一层；JSZip 的键却是全包路径且**不做规范化**，
 * 留着 `../` 去查就查不到。几种写法在这一处收敛，各部件读取处不再各算一遍。
 *
 * @param target `Relationship/@Target` 的原始字符串。
 * @returns 相对包根的路径，例如 `word/comments.xml`；`..` 已按段消掉。
 */
export function wordPartPath(target: string): string {
  const relativeToRoot = target.startsWith("/");
  const segments = (relativeToRoot ? target.slice(1) : `word/${target}`).split("/");
  const resolved: string[] = [];
  for (const segment of segments) {
    if (segment === "" || segment === ".") continue;
    if (segment === "..") resolved.pop();
    else resolved.push(segment);
  }
  return resolved.join("/");
}

export function isElement(n: XmlNode | undefined): n is Element {
  return !!n && n.type === "element";
}

export function isText(n: XmlNode | undefined): n is TextNode {
  return !!n && n.type === "text";
}

export function findChildren(el: Element, name: string): Element[] {
  return children(el).filter((c) => c.name === name);
}

export function firstChild(el: Element, name: string): Element | undefined {
  return children(el).find((c) => c.name === name);
}

export function text(el: Element): string {
  // Concatenate text from all w:t descendants.
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

/** Collect every descendant with the given element name, in document order. */
export function collectDescendants(el: Element, name: string, acc: Element[] = []): Element[] {
  for (const c of children(el)) {
    if (c.name === name) acc.push(c);
    collectDescendants(c, name, acc);
  }
  return acc;
}

/**
 * Serialize an Element tree back to an XML string. We do not use a generic
 * serialization library because we need to output only the `<w:body ...>`
 * subtree in many places, and preserving self-closing tags is not important
 * for OOXML.
 */
export function serialize(node: XmlNode): string {
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

export function escapeText(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

export function escapeAttr(s: string): string {
  return escapeText(s).replace(/"/g, "&quot;").replace(/'/g, "&apos;");
}

/**
 * 读 OOXML 的开关型属性（CT_OnOff，如 `w:keepNext` / `w:hidden` / `w:default`）。
 *
 * `styles.ts` 与 `ooxml-pm.ts` 都要读它，所以放在这个最底层的模块里，
 * 避免出现第二份"裸元素算 true"的判断。
 *
 * @param el 承载开关的元素；元素不存在时传 undefined/null。
 * @returns 裸元素（`<w:keepNext/>`）为 true，`w:val` 是 0/false/off 为 false；
 *   元素缺失或值读不懂时返回 undefined，含义是"这里没建模过它"，
 *   调用方不得替 Word 补一个它没写过的值。
 */
export function onOffValue(el: Element | null | undefined): boolean | undefined {
  if (!el) return undefined;
  const raw = el.attributes?.["w:val"];
  if (raw === undefined) return true;
  const v = String(raw).toLowerCase();
  if (v === "0" || v === "false" || v === "off") return false;
  return v === "1" || v === "true" || v === "on" ? true : undefined;
}

/**
 * 读 OOXML 属性表里那个通用的 @w:val（`w:jc`、`w:pStyle`、`w:sz` 都以它装值）。
 *
 * `docx-io` 里每个读属性的模块都要它，所以和 `onOffValue` 一起放在最底层：
 * 之前 styles/numbering/ooxml-pm 各写了一份一模一样的，读法稍有出入就会让
 * "正文读得到、样式读不到"这类分歧无从对照。
 *
 * @param el 承载 @w:val 的元素；不存在时传 undefined/null。
 * @returns 属性原值（不换算单位、不解释语义）；元素缺失或没写 @w:val 时返回 undefined，
 *   含义是"这里没建模过"，调用方不得替 Word 补一个它没写过的值。
 */
export function val(el: Element | null | undefined): string | undefined {
  if (!el) return undefined;
  return (el.attributes?.["w:val"] as string | undefined) ?? undefined;
}

/**
 * 元素内部 XML（不含外层标签）。
 *
 * `styles.ts`（读 `w:style` 的 `w:pPr` / `w:rPr`）与 `style-edit.ts`（写回时取原定义的
 * 内部串再合并）要的是同一件事，所以放在最底层的 XML 模块里，避免两份实现的
 * 子元素序列化方式走偏。
 *
 * @param el 已解析的元素。
 * @returns 其子元素按顺序序列化的结果；没有子元素时返回空串。
 */
export function innerXml(el: Element): string {
  return children(el)
    .map(serialize)
    .join("");
}

/**
 * Parse a standalone XML fragment (multiple top-level elements allowed) and
 * return them as Element nodes. Wraps the fragment in a synthetic root.
 */
export function parseFragment(xml: string): Element[] {
  const wrapped = `<root xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">${xml}</root>`;
  const root = parseDocument(wrapped);
  return children(root);
}
