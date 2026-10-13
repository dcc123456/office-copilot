/**
 * 「这份 `.xlsx` 能不能被办公软件直接打开」的结构审计（spec 阶段 6 之后的兼容性切片）。
 *
 * 为什么单独立一层：本机的验收通道只有 openpyxl（只读值，不做结构合规判定）与 Numbers（能开就是能开，
 * 不能开只给一句"文件已损坏"），装不了桌面 Excel（`AGENTS.md` §8）。所以把 Excel 打开时会检查的
 * **OPC 完整性与部件内元素顺序**写成可机检的判定，产出结构化问题码，让 CLI/UI 各端自己拼文案（§5.7）。
 *
 * 判定依据一律取 ECMA-376 / ISO-IEC 29500-1 的显式约束（元素序列与 `minOccurs`），不取"Excel 大概会喜欢"：
 * - `[Content_Types].xml` 必须覆盖每个部件（Part 1 §10.1.2）；
 * - 每条非 External 关系的 Target 必须解析到真实存在的部件（§9.2）；
 * - 部件里引用的 `r:id` 必须在自己的关系表里存在（§13.2.1）；
 * - `<worksheet>` 的子元素顺序是 XSD 的 `xsd:sequence`，顺序错 = 不合规（§18.3.1.100，MS-OFFMACRO2 附的 schema 片段可逐条核对）；
 * - `<sheetData>` 是 `worksheet` 唯一 `minOccurs="1"` 的孩子（同上）。
 *
 * 只判"会不会被判损坏"，不判"图好不好看"：数值口径、图表语义由 `analysis/` 与 `chart-spec.ts` 负责。
 */
import { loadPackage, readText, resolveTarget } from "./opc.js";
import { xml } from "@office-copilot/docx-io";

/**
 * `<worksheet>` 允许的子元素顺序（ECMA-376 Part 1 §18.3.1.100 的 `xsd:sequence`，与 MS-OFFMACRO2 公布的
 * schema 片段逐条对齐；`legacyDrawing`/`legacyDrawingHF` 在 ISO 版本里排在 `drawing` 之后）。
 */
const WORKSHEET_SEQUENCE = [
  "sheetPr",
  "dimension",
  "sheetViews",
  "sheetFormatPr",
  "cols",
  "sheetData",
  "sheetCalcPr",
  "sheetProtection",
  "protectedRanges",
  "scenarios",
  "autoFilter",
  "sortState",
  "dataConsolidate",
  "customSheetViews",
  "mergeCells",
  "phoneticPr",
  "conditionalFormatting",
  "dataValidations",
  "hyperlinks",
  "printOptions",
  "pageMargins",
  "pageSetup",
  "headerFooter",
  "rowBreaks",
  "colBreaks",
  "customProperties",
  "cellWatches",
  "ignoredErrors",
  "smartTags",
  "drawing",
  "legacyDrawing",
  "legacyDrawingHF",
  "picture",
  "oleObjects",
  "controls",
  "webPublishItems",
  "tableParts",
  "extLst",
] as const;

/** 一条结构问题。 */
export interface CompatIssue {
  /** 机器可判的问题码；文案由调用方组装（`AGENTS.md` §5.7）。 */
  code:
    | "part-not-declared"
    | "declared-part-missing"
    | "relationship-target-missing"
    | "relationship-id-duplicate"
    | "unresolved-relationship-id"
    | "worksheet-child-order"
    | "worksheet-missing-sheet-data"
    | "chart-missing-chart";
  /** 出问题的部件路径（包内路径，如 `xl/worksheets/sheet1.xml`）。 */
  part: string;
  /** 结构化细节（元素名、关系 Id、目标路径等），便于渲染层拼中文句子。 */
  detail: string;
}

/**
 * 去掉命名空间前缀，只留本地名。
 *
 * @param name 元素名（可能是 `c:chartSpace`，也可能是默认命名空间下的 `chartSpace`）。
 * @returns 本地名；无冒号时原样返回。
 */
function localName(name: string): string {
  const colon = name.indexOf(":");
  return colon < 0 ? name : name.slice(colon + 1);
}

/**
 * 审计一份 `.xlsx` 的包结构与部件顺序。
 *
 * @param bytes `.xlsx` 的原始字节（磁盘上的、或本包 `toBytes()` 刚写出的）。
 * @returns 问题数组，**空数组代表没发现结构缺陷**；不是 zip 时由 `loadPackage` 抛 `SheetOpenError` 同源错误，调用方按读文件失败处理。
 *   只报"会被办公软件判为损坏"级别的确定问题，不推测渲染结果。
 */
export async function auditXlsxPackage(bytes: Uint8Array): Promise<CompatIssue[]> {
  const pkg = await loadPackage(bytes);
  const issues: CompatIssue[] = [];
  const entrySet = new Set(pkg.entryNames);
  const partNames = pkg.entryNames.filter((name) => !name.endsWith("/"));

  // §10.1.2：每个部件都要被 Default（按扩展名）或 Override（按部件名）覆盖，漏了 Excel 直接不认这个部件。
  for (const part of partNames) {
    if (part === "[Content_Types].xml") continue;
    const extension = part.slice(part.lastIndexOf(".") + 1).toLowerCase();
    if (pkg.overrideParts.has(part) || pkg.contentTypes.has(extension)) continue;
    issues.push({ code: "part-not-declared", part, detail: extension });
  }
  // Override 指向不存在的部件同样是"包内容类型表说谎"，Excel 会在打开时报错。
  for (const part of pkg.overrideParts.keys()) {
    if (!entrySet.has(part)) issues.push({ code: "declared-part-missing", part, detail: "" });
  }

  // §9.2：每条非 External 关系的 Target 必须解析到真实存在的部件，同一张关系表里 Id 不许重复。
  const idsBySourcePart = new Map<string, Set<string>>();
  for (const [sourcePart, rels] of pkg.relsBySourcePart) {
    const seen = new Set<string>();
    for (const rel of rels) {
      if (seen.has(rel.id)) issues.push({ code: "relationship-id-duplicate", part: sourcePart, detail: rel.id });
      seen.add(rel.id);
      if (rel.targetMode === "External") continue;
      const target = resolveTarget(sourcePart, rel.target);
      if (!entrySet.has(target)) {
        issues.push({ code: "relationship-target-missing", part: sourcePart, detail: `${rel.id} -> ${target}` });
      }
    }
    idsBySourcePart.set(sourcePart, seen);
  }

  for (const part of partNames) {
    if (part === "[Content_Types].xml" || part.endsWith(".rels")) continue;
    const text = await readText(pkg.zip, part);
    if (!text) continue;
    // §13.2.1：部件里引用的每个 `r:id` 都要能在**它自己**的关系表里查到（没有关系表就是全都查不到）。
    const ids = idsBySourcePart.get(part) ?? new Set<string>();
    for (const id of referencedRelIds(text)) {
      if (!ids.has(id)) issues.push({ code: "unresolved-relationship-id", part, detail: id });
    }
    if (/^xl\/worksheets\/[^/]+\.xml$/.test(part)) issues.push(...auditWorksheet(part, text));
    if (/^xl\/charts\/[^/]+\.xml$/.test(part)) {
      const root = xml.parseDocument(text);
      // 图表部件一律带 `c:` 前缀，而 `findChildren` 按名字精确匹配，所以这里自己比本地名。
      if (root && !xml.children(root).some((el) => localName(el.name) === "chart")) {
        issues.push({ code: "chart-missing-chart", part, detail: "" });
      }
    }
  }
  return issues;
}

/**
 * 扫出部件文本里引用到的关系 Id（`r:id="..."` 这一族属性）。
 *
 * @param text 部件 XML 原文。
 * @returns 去重后的引用 Id 列表；写错的引用（前缀不是 `r` 的 `id=` 属性如 `c:idx val`）不会被误收。
 */
function referencedRelIds(text: string): string[] {
  const ids = new Set<string>();
  for (const match of text.matchAll(/\br:id="([^"]+)"/g)) ids.add(match[1] ?? "");
  return [...ids].filter(Boolean);
}

/**
 * 审 `<worksheet>` 的孩子顺序与必选元素（§18.3.1.100）。
 *
 * @param part 部件路径（只用于问题定位）。
 * @param text 表部件原文。
 * @returns 发现的问题；顺序全对且有 `<sheetData>` 时为空数组。
 *   序列之外的陌生元素只跳过、不报——扩展元素（`extLst` 内部、`mc:AlternateContent`）合法存在。
 */
function auditWorksheet(part: string, text: string): CompatIssue[] {
  const root = xml.parseDocument(text);
  if (!root) return [];
  const worksheet = localName(root.name) === "worksheet" ? root : xml.findChildren(root, "worksheet")[0] ?? root;
  const issues: CompatIssue[] = [];
  const children = xml.children(worksheet).map((el) => localName(el.name));
  if (!children.includes("sheetData")) issues.push({ code: "worksheet-missing-sheet-data", part, detail: "" });

  let previousIndex = -1;
  let previousName = "";
  for (const name of children) {
    const index = WORKSHEET_SEQUENCE.indexOf(name as (typeof WORKSHEET_SEQUENCE)[number]);
    if (index < 0) continue;
    // 逐个与**前一个**已知元素比：只报"相邻两个颠倒了"，不拿全程最大值去卡，
    // 否则一处颠倒会把后面本来合规的元素全牵连成误报。
    if (index < previousIndex) issues.push({ code: "worksheet-child-order", part, detail: `${previousName} -> ${name}` });
    previousIndex = index;
    previousName = name;
  }
  return issues;
}
