/**
 * OPC 写侧的"新建部件三处一起登记"原语（ECMA-376 Part 1、OPC §3）。
 *
 * 为什么单列一层：`AGENTS.md` §2.10 规定新建一个部件必须同时补**部件本体 + 指向它的 `<Relationship>` +
 * `[Content_Types].xml` 里的 Override**，少任一件办公软件判文件损坏。这段逻辑 `.xlsx`（图表、画图、凭空新建的
 * `sharedStrings`）与 `.pptx`（幻灯片、母版、嵌入工作簿）完全一样，只落在 `SheetFile` 里的话，
 * 等 `pptx-io` 要用的时候就是第二次复制（§3.2）。所以提到格式无关的 OPC 层，两边都调这一份。
 *
 * 这些函数都只**产出文本 / 算号**，不碰任何暂存 Map——"改写还是新增"是上层写器的记账，不属于 OPC。
 */
import { readText, relsPathFor, relativeTarget, type OpcPackage } from "./opc.js";

/** `[Content_Types].xml` 缺失时的最小骨架（实测原文件多半有，缺时整篇新建）。 */
const TYPES_SKELETON =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>';

/**
 * 源部件下一个可用的关系编号。
 *
 * @param pkg 已加载的包（读它的 `relsBySourcePart`）。
 * @param sourcePartPath 关系所属的源部件，如 `xl/workbook.xml`、`ppt/slides/slide1.xml`。
 * @returns `rIdN` 形式的 id；已有的超链接/图表关系原样让位，不许撞号。
 */
export function nextRelIdFor(pkg: OpcPackage, sourcePartPath: string): string {
  const existing = pkg.relsBySourcePart.get(sourcePartPath) ?? [];
  const highest = existing.reduce((max, rel) => Math.max(max, Number(/^rId(\d+)$/.exec(rel.id)?.[1] ?? 0)), 0);
  return `rId${highest + 1}`;
}

/**
 * 源部件的关系文件：保留原文，追加一条指向新部件的关系。
 *
 * @param pkg 已加载的包（读它的 `zip`）。
 * @param sourcePartPath 关系所属的源部件。
 * @param relId 新关系编号（来自 `nextRelIdFor`）。
 * @param relType 关系类型 URI。
 * @param targetPartPath 目标部件路径；写成**相对源部件目录**的形式（实测绝对写法办公软件拒开）。
 * @returns 可直接写回的 `.rels` 文本；原文件没有这个关系文件时整篇新建。
 */
export async function relsWith(pkg: OpcPackage, sourcePartPath: string, relId: string, relType: string, targetPartPath: string): Promise<string> {
  const relationship = `<Relationship Id="${relId}" Type="${relType}" Target="${relativeTarget(sourcePartPath, targetPartPath)}"/>`;
  const original = await readText(pkg.zip, relsPathFor(sourcePartPath));
  return original
    ? original.replace("</Relationships>", `${relationship}</Relationships>`)
    : `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationship}</Relationships>`;
}

/**
 * 给 `[Content_Types].xml` 追加 Override 声明。
 *
 * @param pkg 已加载的包（读它的 `zip`）。
 * @param entries `[部件路径, 内容类型]` 二元组，路径要带前导斜杠。
 * @returns 新的部件文本；同名 Override 已存在时不重复追加。
 */
export async function contentTypesWith(pkg: OpcPackage, entries: ReadonlyArray<readonly [string, string]>): Promise<string> {
  const original = (await readText(pkg.zip, "[Content_Types].xml")) ?? TYPES_SKELETON;
  return entries.reduce((xml, [partName, contentType]) => {
    if (xml.includes(`PartName="${partName}"`)) return xml;
    return xml.replace("</Types>", `<Override PartName="${partName}" ContentType="${contentType}"/></Types>`);
  }, original);
}

/**
 * 下一个可用的同名部件序号，避开包里已有的与本次已排队的同类部件。
 *
 * @param takenNames 已占用与待新增的部件路径合集（调用方把包内清单与自己排队的键拼一起传进来）。
 * @param prefix 部件路径前缀，如 `xl/charts/chart`、`ppt/slides/slide`。
 * @param suffix 扩展名，如 `.xml`。
 * @returns 命中数量 + 1（1 基）。
 */
export function nextPartIndex(takenNames: Iterable<string>, prefix: string, suffix: string): number {
  let taken = 0;
  for (const name of takenNames) {
    if (name.startsWith(prefix) && name.endsWith(suffix)) taken += 1;
  }
  return taken + 1;
}
