/**
 * 「这份 `.xlsx` 能不能被办公软件直接打开」的结构审计（spec 阶段 6 之后的兼容性切片）。
 *
 * 为什么单独立一层：本机的验收通道只有 openpyxl（只读值，不做结构合规判定）与 Numbers（能开就是能开，
 * 不能开只给一句"文件已损坏"），装不了桌面 Excel（`AGENTS.md` §8）。所以把 Excel 打开时会检查的
 * **OPC 完整性与部件内元素顺序**写成可机检的判定，产出结构化问题码，让 CLI/UI 各端自己拼文案（§5.7）。
 *
 * 其中与格式无关的五条（内容类型覆盖、关系可解析、`r:id` 落点、Id 不重复）来自 `@office-copilot/opc` 的
 * 通用审计 `auditOpcPackage`（`AGENTS.md` §3，全仓只一份 OPC 判定），本文件只补 `.xlsx` 专属的两条部件顺序判定：
 * - `<worksheet>` 的子元素顺序是 XSD 的 `xsd:sequence`，顺序错 = 不合规（§18.3.1.100，MS-OFFMACRO2 附的 schema 片段可逐条核对）；
 * - `<sheetData>` 是 `worksheet` 唯一 `minOccurs="1"` 的孩子（同上）。
 *
 * 只判"会不会被判损坏"，不判"图好不好看"：数值口径、图表语义由 `analysis/` 与 `chart-spec.ts` 负责。
 */
import { auditOpcPackage, loadPackage, readText, xml, type OpcAuditCode } from "@office-copilot/opc";

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

/** 一条结构问题：通用码来自 `opc`，后面两项是 `.xlsx` 专属的部件顺序判定。 */
export interface CompatIssue {
  /** 机器可判的问题码；文案由调用方组装（`AGENTS.md` §5.7）。 */
  code: OpcAuditCode | "worksheet-child-order" | "worksheet-missing-sheet-data" | "chart-missing-chart";
  /** 出问题的部件路径（包内路径，如 `xl/worksheets/sheet1.xml`）。 */
  part: string;
  /** 结构化细节（元素名、关系 Id、目标路径等），便于渲染层拼中文句子。 */
  detail: string;
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
  const issues: CompatIssue[] = [...(await auditOpcPackage(pkg))];

  for (const part of pkg.entryNames.filter((name) => /^xl\/(worksheets|charts)\/[^/]+\.xml$/.test(name))) {
    const text = await readText(pkg.zip, part);
    if (!text) continue;
    if (part.startsWith("xl/worksheets/")) issues.push(...auditWorksheet(part, text));
    if (part.startsWith("xl/charts/")) {
      const root = xml.parseDocument(text);
      // 图表部件一律带 `c:` 前缀，而 `findChildren` 按名字精确匹配，所以这里自己比本地名。
      if (root && !xml.children(root).some((el) => xml.localName(el.name) === "chart")) {
        issues.push({ code: "chart-missing-chart", part, detail: "" });
      }
    }
  }
  return issues;
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
  const worksheet = xml.localName(root.name) === "worksheet" ? root : xml.findChildren(root, "worksheet")[0] ?? root;
  const issues: CompatIssue[] = [];
  const children = xml.children(worksheet).map((el) => xml.localName(el.name));
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
