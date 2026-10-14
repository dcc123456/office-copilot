/**
 * 与 `@office-copilot/opc` 共用的 XML 助手出口。
 *
 * 这些遍历/序列化函数是通用 OOXML 能力，三种格式（docx/xlsx/pptx）共用一份，
 * 这里只做**再导出 + 类型别名 + 表格专属的两个小工具**（`AGENTS.md` §3.1/§3.2）。
 */
import { xml } from "@office-copilot/opc";

export const {
  parseDocument,
  serialize,
  children,
  findChildren,
  firstChild,
  isElement,
  isText,
  text,
  innerXml,
  escapeText,
  escapeAttr,
} = xml;

/** 解析后的元素节点（类型从复用入口推导，避免手写第三份形状）。 */
export type Element = ReturnType<typeof xml.parseDocument>;

/** 序列化接受的任何节点。 */
export type XmlNode = Parameters<typeof xml.serialize>[0];

/** OOXML 统一的 XML 声明；定义在 `@office-copilot/opc`，这里只是别名，不留第二份。 */
export const XML_DECLARATION = xml.XML_DECLARATION;

/** spreadsheetml 主命名空间（实测所有 `xl/**` 部件的默认 xmlns 都是它）。 */
export const SPREADSHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

/**
 * 造一个属于 spreadsheetml 的元素节点。
 *
 * @param markup 单根元素的 XML 片段，如 `<numFmt numFmtId="164" formatCode="yyyy-mm-dd"/>`。
 * @returns 不带显式 `xmlns` 属性的元素——直接 `push` 进已有树时，命名空间由父根的默认声明继承，
 *   若子元素自己再写一份 `xmlns` 虽然合法但会让往返结果与原文件对不上。
 */
export function sheetElement(markup: string): Element {
  const wrapper = parseDocument(`<root xmlns="${SPREADSHEET_NS}">${markup}</root>`);
  const first = children(wrapper)[0];
  if (!first) throw new Error(`无法解析 XML 片段：${markup}`);
  return first;
}

/**
 * 按 OOXML 的 schema 顺序插入元素。
 *
 * `styleSheet` 的子元素顺序是固定的（`numFmts` 必须在 `fonts` 前），插错位置 Word/Excel
 * 会判文档损坏；这里按给定顺序表找到第一个"排在后面的兄弟"再插到它前面。
 *
 * @param parent 根元素。
 * @param newEl 待插元素。
 * @param order 该容器允许的子元素顺序表（元素名）。
 */
export function insertInSchemaOrder(parent: Element, newEl: Element, order: readonly string[]): void {
  const kids = children(parent);
  const wantedIndex = order.indexOf(newEl.name);
  for (const existing of kids) {
    const existingIndex = order.indexOf(existing.name);
    if (existingIndex < 0) continue;
    if (wantedIndex >= 0 && existingIndex > wantedIndex) {
      const at = (parent.children ?? []).indexOf(existing);
      parent.children?.splice(at, 0, newEl);
      return;
    }
  }
  parent.children?.push(newEl);
}
