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
import JSZip from "jszip";
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
 * 源部件的关系文件：保留原文，追加指向新部件的关系。
 *
 * 为什么一次收多条而不是循环调一次一个：新建一个部件常常要**同时**给它挂上两条关系（`.pptx` 的备注页要指向
 * notesMaster 与幻灯片，`.xlsx` 的画图部件要指向图表与内嵌工作簿），而这一段在"原文件里没有这个关系文件"时
 * 是**整篇新建**——循环调两次就会各拿一份只有一条关系的骨架，第二次把第一次那条盖掉。
 *
 * @param pkg 已加载的包（读它的 `zip`）。
 * @param sourcePartPath 关系所属的源部件。
 * @param entries 要追加的关系：编号（来自 `nextRelIdFor`）、类型 URI、目标部件路径。
 * @returns 可直接写回的 `.rels` 文本；原文件没有这个关系文件时整篇新建。
 */
export async function relsWith(
  pkg: OpcPackage,
  sourcePartPath: string,
  entries: ReadonlyArray<{ id: string; type: string; targetPartPath: string }>,
): Promise<string> {
  const relationships = entries.map(
    ({ id, type, targetPartPath }) =>
      `<Relationship Id="${id}" Type="${type}" Target="${relativeTarget(sourcePartPath, targetPartPath)}"/>`,
  );
  const original = await readText(pkg.zip, relsPathFor(sourcePartPath));
  return original
    ? original.replace("</Relationships>", `${relationships.join("")}</Relationships>`)
    : `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationships.join("")}</Relationships>`;
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

/**
 * 把装好部件的 JSZip 序列化成办公文档字节，并且**不写目录项**。
 *
 * 为什么单列这一份（`.docx`/`.xlsx`/`.pptx` 三处 `toBytes()` 共用）：写侧新增一个嵌套部件
 * （`zip.file("ppt/slides/slide2.xml", …)`、`word/media/image1.png`）时，JSZip 会顺手把 `ppt/`、
 * `ppt/slides/` 这类 `dir:true` 目录对象塞进 `zip.files`，`generateAsync` 再把它们写进包——于是我们产的
 * 文件比 python-pptx 那种"只有文件条目"的原文件多出目录项。**实测**：`generateAsync({createFolders:false})`
 * 挡不住（那些目录对象是 `file()` 当场造的，不是生成时补的），必须在生成前把 `dir` 条目删掉；删 `zip.files`
 * 里的对象而不是 `zip.remove(name)`，因为 `remove` 按前缀递归、会把 `ppt/` 底下的真部件一起删没。
 * 目录项对 OPC 是可选的（办公软件忽略，`unmodelledParts` 也按 `/` 结尾过滤），所以三份统一都不写最省心。
 *
 * @param zip 已装好全部部件的包（没改过的部件原样留着，含二进制部件）。
 * @returns DEFLATE 压缩的 zip 字节；产物不含任何目录项，文件部件一个不少。
 */
export async function serializePackage(zip: JSZip): Promise<Uint8Array> {
  for (const name of Object.keys(zip.files)) if (zip.files[name]?.dir) delete zip.files[name];
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}
