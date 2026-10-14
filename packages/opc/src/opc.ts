/**
 * OPC（zip 包）读取：部件清单、`[Content_Types].xml`、关系表。
 *
 * 这一层是**格式无关**的——`.docx`/`.xlsx`/`.pptx` 的包结构规则同属 ECMA-376 Part 2 (OPC)，
 * 之前它住在 `sheet-io` 里且没有对外导出，pptx 线要用就只能再写第三套（`AGENTS.md` §3.2）。
 * 各格式特有的关系类型判断（例如"这是不是一张 worksheet"）留在各自的包里，不塞进这里。
 */
import JSZip from "jszip";
import * as xml from "./xml.js";

/** 已加载的包。`entryNames` 含目录项，读内容前要先过 `readText`。 */
export interface OpcPackage {
  zip: JSZip;
  entryNames: string[];
  /** 关系表：键是承载关系的部件路径（如 `xl/workbook.xml`），值是其关系。 */
  relsBySourcePart: Map<string, OpcRelationship[]>;
  contentTypes: Map<string, string>;
  overrideParts: Map<string, string>;
}

/** 一条 OPC 关系。 */
export interface OpcRelationship {
  id: string;
  type: string;
  target: string;
  targetMode: string;
}

/**
 * 读取包并解析全部关系表。
 *
 * @param bytes 办公文档的原始字节（`.xlsx`/`.docx`/`.pptx` 都是 zip）。
 * @returns 可直接查部件文本的包对象；不是 zip 时由 JSZip 抛错，调用方负责转成用户可读提示。
 */
export async function loadPackage(bytes: Uint8Array): Promise<OpcPackage> {
  const zip = await JSZip.loadAsync(bytes);
  const entryNames = Object.keys(zip.files).sort();

  const relsBySourcePart = new Map<string, OpcRelationship[]>();
  for (const name of entryNames) {
    const match = /^(.*)\/_rels\/([^/]+)\.rels$/.exec(name);
    if (!match) continue;
    const folder = match[1] ?? "";
    const sourceFile = match[2] ?? "";
    const sourcePart = folder ? `${folder}/${sourceFile}` : sourceFile;
    relsBySourcePart.set(sourcePart, parseRelationships((await readText(zip, name)) ?? ""));
  }

  const contentTypes = new Map<string, string>();
  const overrideParts = new Map<string, string>();
  const typesXml = await readText(zip, "[Content_Types].xml");
  if (typesXml) {
    const root = xml.parseDocument(typesXml);
    for (const el of xml.children(root)) {
      if (el.name === "Default") {
        const ext = el.attributes["Extension"];
        const contentType = el.attributes["ContentType"];
        if (ext && contentType) contentTypes.set(String(ext).toLowerCase(), String(contentType));
      } else if (el.name === "Override") {
        const part = el.attributes["PartName"];
        const contentType = el.attributes["ContentType"];
        if (part && contentType) overrideParts.set(String(part).replace(/^\//, ""), String(contentType));
      }
    }
  }

  return { zip, entryNames, relsBySourcePart, contentTypes, overrideParts };
}

/**
 * 解析一个 `.rels` 文件。
 *
 * @param relsXml 关系表 XML 文本。
 * @returns 关系数组；文本为空时返回空数组。
 */
export function parseRelationships(relsXml: string): OpcRelationship[] {
  if (!relsXml.trim()) return [];
  const root = xml.parseDocument(relsXml);
  return xml.findChildren(root, "Relationship").map((el) => ({
    id: String(el.attributes["Id"] ?? ""),
    type: String(el.attributes["Type"] ?? ""),
    target: String(el.attributes["Target"] ?? ""),
    targetMode: String(el.attributes["TargetMode"] ?? "Internal"),
  }));
}

/**
 * 把 `@Target` 换算成包内路径。
 *
 * OPC 允许三种写法：相对关系文件所在目录（`worksheets/sheet1.xml`）、相对源部件目录、
 * 绝对 `/xl/...`，还允许 `..`。`docx-io` 里的 `wordPartPath` 把基准硬编码成了 `word/`，
 * 这里的基准是从关系文件路径推出来的，所以对任何格式的目录都成立。
 *
 * @param sourcePartPath 源部件路径，如 `xl/workbook.xml`（不是 `.rels` 本身）。
 * @param target `Relationship/@Target` 原值。
 * @returns 规范化的包内路径，如 `xl/worksheets/sheet1.xml`。
 */
export function resolveTarget(sourcePartPath: string, target: string): string {
  if (target.startsWith("/")) return normalizeSegments(target.slice(1));
  const folder = sourcePartPath.includes("/") ? sourcePartPath.slice(0, sourcePartPath.lastIndexOf("/")) : "";
  return normalizeSegments(folder ? `${folder}/${target}` : target);
}

/**
 * `resolveTarget` 的反向：把包内绝对路径写成关系文件里该出现的相对 `@Target`。
 *
 * 只在**新建**关系时用（图表、画图部件）。写错一层目录办公软件会直接判文档损坏，
 * 所以这里给出唯一的拼法，并且保证 `resolveTarget(src, relativeTarget(src, dst)) === dst`。
 *
 * @param sourcePartPath 源部件路径，如 `xl/worksheets/sheet1.xml`。
 * @param targetPartPath 目标部件的包内路径，如 `xl/drawings/drawing1.xml`。
 * @returns 相对源部件目录的写法，如 `../drawings/drawing1.xml`。
 */
export function relativeTarget(sourcePartPath: string, targetPartPath: string): string {
  const base = sourcePartPath.includes("/") ? sourcePartPath.slice(0, sourcePartPath.lastIndexOf("/")).split("/") : [];
  const target = normalizeSegments(targetPartPath).split("/");
  let shared = 0;
  while (shared < base.length && shared < target.length - 1 && base[shared] === target[shared]) shared += 1;
  return `${"../".repeat(base.length - shared)}${target.slice(shared).join("/")}`;
}

function normalizeSegments(path: string): string {
  const resolved: string[] = [];
  for (const segment of path.split("/")) {
    if (segment === "" || segment === ".") continue;
    if (segment === "..") resolved.pop();
    else resolved.push(segment);
  }
  return resolved.join("/");
}

/**
 * 读部件文本。
 *
 * @param zip 包。
 * @param partPath 包内路径。
 * @returns UTF-8 文本；部件不存在（含目录项）返回 undefined。
 */
export async function readText(zip: JSZip, partPath: string): Promise<string | undefined> {
  const entry = zip.file(partPath);
  if (!entry) return undefined;
  return entry.async("string");
}

/**
 * 关系文件的位置：`xl/worksheets/sheet1.xml` → `xl/worksheets/_rels/sheet1.xml.rels`；
 * 根级部件（不含 `/`）落在包根的 `_rels/` 下。
 *
 * @param partPath 源部件路径。
 * @returns 对应的 `.rels` 路径。
 */
export function relsPathFor(partPath: string): string {
  const slash = partPath.lastIndexOf("/");
  const folder = slash < 0 ? "" : partPath.slice(0, slash);
  const file = slash < 0 ? partPath : partPath.slice(slash + 1);
  return `${folder ? `${folder}/` : ""}_rels/${file}.rels`;
}

/**
 * 列出包里"存在但没有被建模"的部件（保真验收用）。
 *
 * @param pkg 已加载的包。
 * @param modelledPaths 我们已知会读写的路径集合。
 * @returns 其余文件部件的路径，顺序稳定（按路径排序）；关系文件不计入（它属于登记结构，不是内容部件）。
 */
export function unmodelledParts(pkg: OpcPackage, modelledPaths: ReadonlySet<string>): string[] {
  return pkg.entryNames.filter(
    (name) => !name.endsWith("/") && !modelledPaths.has(name) && !name.endsWith(".rels") && !name.endsWith("/"),
  );
}
