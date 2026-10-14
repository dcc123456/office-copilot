/**
 * 通用 XML 层现在住在 `@office-copilot/opc`（`AGENTS.md` §3.2：三种格式共用一份，不许第二套）。
 *
 * 这里只保留**docx 专属**的三件：`wordPartPath`（基准目录硬编码成 `word/`）、
 * `onOffValue` 与 `val`（读的都是 `@w:val`，属性名前缀是 `w:`），其余通用函数一律再导出。
 */
import { children, parseDocument, type Element } from "@office-copilot/opc";

export {
  parseDocument,
  children,
  isElement,
  isText,
  findChildren,
  firstChild,
  text,
  collectDescendants,
  serialize,
  escapeText,
  escapeAttr,
  innerXml,
  XML_DECLARATION,
} from "@office-copilot/opc";
export type { Element, TextNode } from "@office-copilot/opc";

/**
 * 把 `word/_rels/document.xml.rels` 里的 `@Target` 换算成包内路径。
 *
 * rels 的 Target 按约定相对 `word/` 写（`header1.xml`、`./styles.xml`），也有绝对写法
 * （`/word/foo.xml`），OPC 还允许 `..` 往上一层；JSZip 的键却是全包路径且**不做规范化**，
 * 留着 `../` 去查就查不到。几种写法在这一处收敛，各部件读取处不再各算一遍。
 *
 * 注意：这是 **docx 专属**的（基准写死 `word/`）。跨格式的通用版本是 `@office-copilot/opc` 的
 * `resolveTarget(sourcePartPath, target)`，它从源部件路径推基准。
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
 * 解析一段独立的 XML 片段（允许多个顶层元素）并返回元素节点。
 *
 * @param xml 片段正文。
 * @returns 顶层元素数组；包了一层带 wordprocessingml/relationships/drawingml 命名空间的合成根，
 *   所以片段里用到的前缀都能解析出来。
 */
export function parseFragment(xml: string): Element[] {
  const wrapped = `<root xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">${xml}</root>`;
  const root = parseDocument(wrapped);
  return children(root);
}
