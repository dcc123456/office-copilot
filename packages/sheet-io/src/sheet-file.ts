/**
 * `.xlsx` 文件的出入口（spec D1 的 `packages/sheet-io` 门面）。
 *
 * 三条纪律：
 * 1. 没被我们改写的部件在 `toBytes()` 里逐字节原样回写（spec §零.6 / 验收 A1）。
 * 2. 单位/颜色/序列号的解释只发生在本包（`AGENTS.md` §2.1 的同一条纪律，换成表格域）。
 * 3. 读不懂的东西记进 `warnings`/`unsupported` 一路带到 UI，不许静默丢弃。
 */
import { isWorksheetRelationship, loadPackage, readText, relativeTarget, resolveTarget, type OpcPackage } from "./opc.js";
import { children, isElement, parseDocument, type Element } from "./xml.js";
import type { SheetCell, SheetModel, UnsupportedFeature, WorkbookModel } from "./model.js";
import type { ResolvedChart } from "./chart-spec.js";
import { buildChartParts, CHART_CONTENT_TYPE, DRAWING_CONTENT_TYPE, DRAWING_REL_TYPE } from "./xlsx-chart.js";
import { SHARED_STRINGS_CONTENT_TYPE, SHARED_STRINGS_PART, SHARED_STRINGS_REL_TYPE, SharedStringTable } from "./xlsx-strings.js";
import { StyleTable } from "./xlsx-styles.js";
import { parseSheet, recalculateSheetExtent, serializeSheetXml } from "./xlsx-sheet.js";

/** 打不开时抛出的错误；`reason` 是给渲染层组装文案用的错误码，不是给人看的句子（`AGENTS.md` §5.7）。 */
export class SheetOpenError extends Error {
  constructor(readonly reason: "macro-enabled" | "not-a-zip" | "no-workbook", message: string) {
    super(message);
    this.name = "SheetOpenError";
  }
}

/** 打开参数。 */
export interface SheetOpenOptions {
  /** 文件名，用于识别 `.xlsm` 与在 modlog 里定位文档。 */
  fileName: string;
}

/** 工作簿里一张表的定位信息（未解析时也有）。 */
export interface SheetRef {
  name: string;
  sheetId: number;
  partPath: string;
  hidden: boolean;
}

/** 快照：本包可能改写的所有部件的原文，`restoreSnapshot` 用它整批回退（`AGENTS.md` §2.4）。 */
export interface SheetSnapshot {
  parts: Record<string, string>;
}

/** 一次图表写入的结果。失败只给错误码，文案由渲染层组装（`AGENTS.md` §5.7）。 */
export type ChartAddResult =
  | { ok: true; chartPartPath: string; drawingPartPath: string; sheetPartPath: string }
  | { ok: false; code: "sheet-missing" | "drawing-exists" | "no-worksheet-root" };

/** 关系属性的命名空间，`<drawing r:id>` 要用它前缀。 */
const DRAWING_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/** 工作簿本体的部件路径：新建的关系（字符串表）都挂在它下面。 */
const WORKBOOK_PART = "xl/workbook.xml";

/** 排在 `<drawing>` **之后**的 `CT_Worksheet` 元素（ECMA-376 Part 1 §18.3.1.100 的序列），插入点取第一个命中的。 */
const AFTER_DRAWING_PATTERN =
  /<(?:legacyDrawing|legacyDrawingHF|picture|oleObjects|controls|webPublishItems|tableParts|extLst)\b/;

/** 表里已经有画图部件时的兜底 content types 根（原文件连 `[Content_Types].xml` 都缺才用得到）。 */
const TYPES_SKELETON =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>';

/**
 * 在表部件里放一个 `<drawing r:id>`，指向本次新建的画图部件。
 *
 * 一张表按规范只能有一个 `<drawing>`，所以**已有画图部件（图片、其它图表）的表这里直接拒绝**：
 * 往里追加要重排既有 drawing 的关系，做错了会牵连用户的图片（`AGENTS.md` §2.4 的整批放弃语义）。
 *
 * @param xml 表部件原文。
 * @param relId 本次新建关系用的编号。
 * @returns 新文本；连 `</worksheet>` 都找不到时返回 null，调用方据此不落任何部件。
 */
function insertDrawingElement(xml: string, relId: string): string | null {
  const root = /<worksheet\b[^>]*>/.exec(xml);
  if (!root) return null;
  // 关系属性要 `r` 前缀；原文件没声明过就补一次（补在根标签上，别处不动）。
  const withNs = root[0].includes("xmlns:r=")
    ? xml
    : xml.replace(root[0], `${root[0].slice(0, -1)} xmlns:r="${DRAWING_REL_NS}">`);
  const element = `<drawing r:id="${relId}"/>`;
  const anchor = AFTER_DRAWING_PATTERN.exec(withNs);
  if (anchor?.index !== undefined) return withNs.slice(0, anchor.index) + element + withNs.slice(anchor.index);
  const close = withNs.lastIndexOf("</worksheet>");
  return close < 0 ? null : `${withNs.slice(0, close)}${element}${withNs.slice(close)}`;
}

/** 会破坏"打开即回写也不坏"承诺的部件前缀；命中只记进 `unsupported`，内容仍原样保留。 */
const UNSUPPORTED_PREFIXES: ReadonlyArray<readonly [string, string]> = [
  ["xl/vbaProject.bin", "macro"],
  ["xl/externalLinks/", "externalLinks"],
  ["xl/pivotTables/", "pivotTable"],
  ["xl/pivotCache/", "pivotCache"],
  ["xl/charts/", "chart"],
  ["xl/drawings/", "drawing"],
  ["xl/tables/", "table"],
  ["xl/queryTables/", "queryTable"],
  ["xl/timelines/", "timeline"],
];

/** 已打开的工作簿。 */
export class SheetFile {
  private readonly loadedSheets = new Map<string, SheetModel>();
  private readonly rewrittenParts = new Map<string, string>();
  /** 原文件里没有、本次保存要新增的部件（图表本体、画图部件、它们的关系与 content types）。 */
  private readonly addedParts = new Map<string, string>();
  private readonly originalParts = new Map<string, string>();
  private readonly refs: SheetRef[];
  private readonly sharedStrings: SharedStringTable;
  private readonly styles: StyleTable;

  private constructor(
    private readonly pkg: OpcPackage,
    refs: SheetRef[],
    sharedStrings: SharedStringTable,
    styles: StyleTable,
    readonly fileName: string,
    readonly warnings: string[],
    readonly unsupported: UnsupportedFeature[],
  ) {
    this.refs = refs;
    this.sharedStrings = sharedStrings;
    this.styles = styles;
  }

  /**
   * 从字节打开。
   *
   * @param bytes `.xlsx` 原始字节。
   * @param options 文件名等。
   * @returns 可读写的工作簿；此时表内容还没解析（惰性）。
   * @throws SheetOpenError 宏工作簿（拒绝，不执行任何宏）、不是 zip、或找不到 `xl/workbook.xml`。
   */
  static async fromBytes(bytes: Uint8Array, options: SheetOpenOptions): Promise<SheetFile> {
    let pkg: OpcPackage;
    try {
      pkg = await loadPackage(bytes);
    } catch {
      throw new SheetOpenError("not-a-zip", `${options.fileName} 不是有效的 OOXML 包。`);
    }

    const workbookXml = await readText(pkg.zip, WORKBOOK_PART);
    if (!workbookXml) throw new SheetOpenError("no-workbook", `${options.fileName} 缺少 xl/workbook.xml。`);

    const warnings: string[] = [];
    const unsupported: UnsupportedFeature[] = [];
    for (const name of pkg.entryNames) {
      for (const [prefix, reason] of UNSUPPORTED_PREFIXES) {
        if (name.startsWith(prefix)) {
          if (reason === "macro") {
            throw new SheetOpenError("macro-enabled", `${options.fileName} 含 VBA 宏（${name}），请先在 Excel/WPS 里另存为 .xlsx。`);
          }
          unsupported.push({ partPath: name, reason });
        }
      }
    }
    for (const [part, contentType] of pkg.overrideParts) {
      if (contentType.includes("macroEnabled")) {
        throw new SheetOpenError("macro-enabled", `${options.fileName} 是宏工作簿，本工具不打开它（见 spec §一 非目标 4）。`);
      }
      if (contentType.includes("pivotCaches") || contentType.includes("slicerCaches")) unsupported.push({ partPath: part, reason: "pivotCache" });
    }

    const relationships = pkg.relsBySourcePart.get(WORKBOOK_PART) ?? [];
    const relById = new Map(relationships.map((rel) => [rel.id, rel]));
    const refs: SheetRef[] = [];
    // `<sheet>` 在 `<sheets>` 容器里，不是 `<workbook>` 的直接子元素。
    const sheetElements = childElements(workbookXml, "sheets", "sheet");
    for (const el of sheetElements) {
      const relId = String(el.attributes["r:id"] ?? el.attributes["id"] ?? "");
      const rel = relById.get(relId);
      if (!rel || !isWorksheetRelationship(rel.type)) {
        warnings.push(`sheet-rel-missing:${String(el.attributes["name"] ?? "")}`);
        continue;
      }
      refs.push({
        name: String(el.attributes["name"] ?? ""),
        sheetId: Number(el.attributes["sheetId"] ?? 0),
        partPath: resolveTarget(WORKBOOK_PART, rel.target),
        hidden: String(el.attributes["state"] ?? "visible") === "hidden",
      });
    }

    const sharedStrings = SharedStringTable.parse(await readText(pkg.zip, SHARED_STRINGS_PART));
    const styles = StyleTable.parse(await readText(pkg.zip, "xl/styles.xml"));

    const file = new SheetFile(pkg, refs, sharedStrings, styles, options.fileName, warnings, unsupported);
    for (const ref of refs) {
      const original = await readText(pkg.zip, ref.partPath);
      if (original === undefined) {
        file.warnings.push(`missing-sheet-part:${ref.partPath}`);
        continue;
      }
      file.originalParts.set(ref.partPath, original);
    }
    const sstText = await readText(pkg.zip, SHARED_STRINGS_PART);
    if (sstText) file.originalParts.set(SHARED_STRINGS_PART, sstText);
    const stylesText = await readText(pkg.zip, "xl/styles.xml");
    if (stylesText) file.originalParts.set("xl/styles.xml", stylesText);
    // `xl/workbook.xml` 也存原文：保存时要按它现有的样子补 `calcPr/@fullCalcOnLoad`。
    file.originalParts.set(WORKBOOK_PART, workbookXml);
    return file;
  }

  /** 表数量（含隐藏表）。 */
  get sheetCount(): number {
    return this.refs.length;
  }

  /**
   * 表的定位信息。
   *
   * @param index 表序号（0 基）。
   * @returns 引用信息；越界返回 undefined。
   */
  sheetRefAt(index: number): SheetRef | undefined {
    return this.refs[index];
  }

  /**
   * 解析（或取回已解析的）某张表。
   *
   * @param index 表序号。
   * @returns 表模型；序号越界时返回 undefined，调用方按"没有这张表"处理。
   */
  async sheetAt(index: number): Promise<SheetModel | undefined> {
    const ref = this.refs[index];
    if (!ref) return undefined;
    const cached = this.loadedSheets.get(ref.partPath);
    if (cached) return cached;
    const xmlText = this.originalParts.get(ref.partPath) ?? (await readText(this.pkg.zip, ref.partPath));
    if (xmlText === undefined) return undefined;
    const sheet = parseSheet({ xmlText, partPath: ref.partPath, name: ref.name, sheetId: ref.sheetId, sharedStrings: this.sharedStrings, styles: this.styles });
    this.loadedSheets.set(ref.partPath, sheet);
    return sheet;
  }

  /**
   * 解析全部表，供 CLI 与 AI 工具层使用。
   *
   * @returns 工作簿模型（含未建模部件清单与警告）。
   */
  async toWorkbook(): Promise<WorkbookModel> {
    const sheets: SheetModel[] = [];
    for (let i = 0; i < this.refs.length; i += 1) {
      const sheet = await this.sheetAt(i);
      if (sheet) sheets.push(sheet);
    }
    return {
      fileName: this.fileName,
      sheets,
      sharedStrings: this.sharedStrings.entriesForRead(),
      warnings: [...this.warnings],
      unsupported: [...this.unsupported],
    };
  }

  /** 字符串表（新字符串要经它追加索引，不许在外面自己数）。 */
  get strings(): SharedStringTable {
    return this.sharedStrings;
  }

  /** 样式表（追加唯一入口见 `StyleTable`）。 */
  get styleTable(): StyleTable {
    return this.styles;
  }

  /**
   * 把一张表的当前内容登记为待写入。
   *
   * @param sheet 已被上层（手动编辑或 op 落地）改过的表；调用前不需自己算维度。
   */
  writeSheet(sheet: SheetModel): void {
    this.resolveNumberFormatStyles(sheet);
    recalculateSheetExtent(sheet);
    this.rewrittenParts.set(sheet.partPath, serializeSheetXml(sheet, this.sharedStrings));
    this.loadedSheets.set(sheet.partPath, sheet);
  }

  /**
   * 把 `sheet-io` 解析好的图表作为**原生图表部件**加进某张表（spec 决策 D5）。
   *
   * 一次登记五件事，少一件 Excel 就判文件损坏：图表部件本体、画图部件、画图部件→图表的关系、
   * 表→画图部件的关系、`[Content_Types].xml` 的两条 Override。表部件里只插一个 `<drawing r:id>`，
   * 其余字节照旧由 `serializeSheetXml` 的定点替换带着走，所以已有的图片/图形/超链接不受牵连。
   *
   * @param sheetPartPath 目标表的包内部件路径（来自 `SheetPayload.partPath`）。
   * @param chart `resolveChartSpec` 通过的结果；本方法不复核数据，数据校验只在那一层。
   * @returns 新部件的路径；表不存在或表部件里找不到插入点时返回错误码，本次不落任何部件。
   */
  async addChart(sheetPartPath: string, chart: ResolvedChart): Promise<ChartAddResult> {
    const index = this.refs.findIndex((ref) => ref.partPath === sheetPartPath);
    if (index < 0) return { ok: false, code: "sheet-missing" };
    const sheet = await this.sheetAt(index);
    if (!sheet) return { ok: false, code: "sheet-missing" };
    const relId = this.nextRelIdFor(sheetPartPath);
    if (/<drawing\b/.test(sheet.sourceXml)) return { ok: false, code: "drawing-exists" };
    const withDrawing = insertDrawingElement(sheet.sourceXml, relId);
    if (!withDrawing) return { ok: false, code: "no-worksheet-root" };

    const parts = buildChartParts(chart, {
      chart: this.nextPartIndex("xl/charts/chart", ".xml"),
      drawing: this.nextPartIndex("xl/drawings/drawing", ".xml"),
    });
    sheet.sourceXml = withDrawing;
    this.writeSheet(sheet);
    this.stagePart(parts.chartPartPath, parts.chartXml);
    this.stagePart(parts.drawingPartPath, parts.drawingXml);
    this.stagePart(parts.drawingRelsPath, parts.drawingRelsXml);
    this.stagePart(
      this.relsPathFor(sheetPartPath),
      await this.relsWith(sheetPartPath, relId, DRAWING_REL_TYPE, parts.drawingPartPath),
    );
    this.stagePart(
      "[Content_Types].xml",
      await this.contentTypesWith([
        [`/${parts.chartPartPath}`, CHART_CONTENT_TYPE],
        [`/${parts.drawingPartPath}`, DRAWING_CONTENT_TYPE],
      ]),
    );
    return { ok: true, chartPartPath: parts.chartPartPath, drawingPartPath: parts.drawingPartPath, sheetPartPath };
  }

  /**
   * 关系文件的位置：`xl/worksheets/sheet1.xml` → `xl/worksheets/_rels/sheet1.xml.rels`。
   *
   * @param partPath 源部件路径。
   * @returns 对应的 `.rels` 路径。
   */
  private relsPathFor(partPath: string): string {
    const slash = partPath.lastIndexOf("/");
    return `${partPath.slice(0, slash)}/_rels/${partPath.slice(slash + 1)}.rels`;
  }

  /**
   * 源部件下一个可用的关系编号。
   *
   * @param sourcePartPath 关系所属的源部件，如 `xl/worksheets/sheet1.xml`、`xl/workbook.xml`。
   * @returns `rIdN` 形式的 id；已有的超链接/图表关系原样让位，不许撞号。
   */
  private nextRelIdFor(sourcePartPath: string): string {
    const existing = this.pkg.relsBySourcePart.get(sourcePartPath) ?? [];
    const highest = existing.reduce((max, rel) => Math.max(max, Number(/^rId(\d+)$/.exec(rel.id)?.[1] ?? 0)), 0);
    return `rId${highest + 1}`;
  }

  /**
   * 源部件的关系文件：保留原文，追加一条指向新部件的关系。
   *
   * @param sourcePartPath 关系所属的源部件。
   * @param relId 新关系编号（来自 `nextRelIdFor`）。
   * @param relType 关系类型 URI。
   * @param targetPartPath 目标部件路径；写成**相对源部件目录**的形式（实测绝对写法 Excel 拒开）。
   * @returns 可直接写回的 `.rels` 文本；原文件没有这个关系文件时整篇新建。
   */
  private async relsWith(sourcePartPath: string, relId: string, relType: string, targetPartPath: string): Promise<string> {
    const relationship = `<Relationship Id="${relId}" Type="${relType}" Target="${relativeTarget(sourcePartPath, targetPartPath)}"/>`;
    const original = await readText(this.pkg.zip, this.relsPathFor(sourcePartPath));
    return original
      ? original.replace("</Relationships>", `${relationship}</Relationships>`)
      : `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationship}</Relationships>`;
  }

  /**
   * 给 `[Content_Types].xml` 追加 Override 声明。
   *
   * @param entries `[部件路径, 内容类型]` 二元组，路径要带前导斜杠。
   * @returns 新的部件文本；同名 Override 已存在时不重复追加。
   */
  private async contentTypesWith(entries: ReadonlyArray<readonly [string, string]>): Promise<string> {
    const original = (await readText(this.pkg.zip, "[Content_Types].xml")) ?? TYPES_SKELETON;
    return entries.reduce((xml, [partName, contentType]) => {
      if (xml.includes(`PartName="${partName}"`)) return xml;
      return xml.replace("</Types>", `<Override PartName="${partName}" ContentType="${contentType}"/></Types>`);
    }, original);
  }

  /**
   * 下一个可用的同名部件序号，避开文件里已有的图表/画图部件。
   *
   * @param prefix 部件路径前缀，如 `xl/charts/chart`。
   * @param suffix 扩展名，如 `.xml`。
   * @returns 已有数量 + 1（1 基）。
   */
  private nextPartIndex(prefix: string, suffix: string): number {
    const taken = (names: string[]) => names.filter((name) => name.startsWith(prefix) && name.endsWith(suffix)).length;
    return taken(this.pkg.entryNames) + taken([...this.addedParts.keys()]) + 1;
  }

  /**
   * 登记一个部件的内容：原文件里有的走改写，没有的走新增。
   *
   * @param path 包内部件路径。
   * @param content 部件正文。
   */
  private stagePart(path: string, content: string): void {
    if (this.originalParts.has(path)) this.rewrittenParts.set(path, content);
    else this.addedParts.set(path, content);
  }

  /**
   * 当前包里已有的部件名（测试与验收用它核对"新部件真的进去了"）。
   *
   * @returns 已存在的 + 本次新增的路径集合，已排序。
   */
  partNames(): string[] {
    return [...new Set([...this.pkg.entryNames, ...this.addedParts.keys()])].sort();
  }

  /**
   * 把上层表达的"显示意图"收敛成 `cellXfs` 位置索引（只追加或复用，**绝不重排**，实测 §二 P3）。
   *
   * 渲染层只带 `numberFormat`（它在 IO 层读到的语义），不自己数样式索引；
   * 已经是目标格式的格原样带过（`s` 不动），所以在表里改一个数字格式不会牵连其它格。
   * 换格式时该格原有的字体/底色/边框/对齐一律随格带走（§2.3 不许拍平局部格式）。
   * 读不懂的格（`rawXml` 原样回写的）不碰。
   *
   * @param sheet 待写入的表；需要换格式的格会换成一份补好 `styleIndex` 的副本，上层载荷原样不动。
   */
  private resolveNumberFormatStyles(sheet: SheetModel): void {
    const resolved: Record<string, SheetCell> = {};
    for (const [address, cell] of Object.entries(sheet.cells)) {
      if (!cell.numberFormat || cell.rawXml) continue;
      const current = this.styles.numberFormatFor(cell.styleIndex);
      if (current.numFmtId === cell.numberFormat.numFmtId && current.formatCode === cell.numberFormat.formatCode) continue;
      const styleIndex = this.styles.withNumberFormat(cell.styleIndex, cell.numberFormat);
      resolved[address] = { ...cell, styleIndex, numberFormat: this.styles.numberFormatFor(styleIndex) };
    }
    // 必须整体换表而不是就地改格子：`applySheetPayload` 让表和上层载荷共用同一批 cell 对象，
    // 就地写 `styleIndex` 会把上层 op 记录的前像一起改掉，存过一次盘后撤销就认不出自己（§2.5）。
    if (Object.keys(resolved).length > 0) sheet.cells = { ...sheet.cells, ...resolved };
  }

  /**
   * 当前可改写部件的快照。
   *
   * @returns 快照对象；`restoreSnapshot` 之前多次调用都安全（只记原文）。
   */
  getSnapshot(): SheetSnapshot {
    const parts: Record<string, string> = {};
    for (const [path, text] of this.originalParts) parts[path] = text;
    return { parts };
  }

  /**
   * 整批回退到快照（`AGENTS.md` §2.4 的事务语义：落地不一致就整批放弃）。
   *
   * @param snapshot 之前 `getSnapshot()` 的结果；传 undefined 时什么都不做。
   */
  restoreSnapshot(snapshot: SheetSnapshot | undefined): void {
    if (!snapshot) return;
    this.rewrittenParts.clear();
    this.addedParts.clear();
    for (const [path, text] of Object.entries(snapshot.parts)) this.originalParts.set(path, text);
    this.loadedSheets.clear();
  }

  /**
   * 产出文件字节。
   *
   * @returns 新的 `.xlsx` 字节：改写过的部件用新内容，其余部件原样带过。
   *   字符串表/样式表在没有追加过条目时**完全不重写**，所以"空改动往返"能做到逐字节一致。
   */
  async toBytes(): Promise<Uint8Array> {
    // 原文件里没有的部件不凭空造（例如缺 sheet 部件时只记 warning，见 fromBytes）。
    for (const [path, content] of this.rewrittenParts) {
      if (this.originalParts.has(path)) this.pkg.zip.file(path, content);
    }
    // 本次新建的部件（图表、画图部件、它们的关系、补过的 content types）原文件里没有，要凭空造出来。
    for (const [path, content] of this.addedParts) this.pkg.zip.file(path, content);
    // 改写过任何一张表就要让 Excel 开文件时全量重算：我们只写缓存值，不维护 `calcChain`。
    const workbookXml = this.originalParts.get(WORKBOOK_PART);
    if (this.rewrittenParts.size > 0 && workbookXml) {
      this.pkg.zip.file(WORKBOOK_PART, withFullCalcOnLoad(workbookXml));
    }
    if (this.sharedStrings.addedCount > 0) {
      const allLoaded = this.loadedSheets.size === this.refs.length;
      if (allLoaded) this.sharedStrings.setUsageCount(this.countStringReferences());
      const sstXml = this.sharedStrings.toXml();
      if (sstXml) await this.writeSharedStringsPart(sstXml);
    }
    if (this.styles.isDirty) this.pkg.zip.file("xl/styles.xml", this.styles.toXml());
    return this.pkg.zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
  }

  /**
   * 写字符串表部件；**原文件里没有这张表时**把关系与内容类型一起补上。
   *
   * 少补任一条，产出的文件就没人读得回：openpyxl 实测 `IndexError: list index out of range`
   * （它按 `xl/workbook.xml` 的关系找表），Excel 判文件损坏——只写部件是不够的（ECMA-376 Part 1 §18.2.23、OPC §3）。
   * 触发场景实测过：只有 `<c t="inlineStr">` 的文件（openpyxl / 部分导出器产出）本来就没有 `xl/sharedStrings.xml`，
   * 而我们一律把文本写成 `t="s"`，就必须凭空新建这张表并登记它。
   *
   * @param sstXml 序列化好的 `<sst>` 文本。
   */
  private async writeSharedStringsPart(sstXml: string): Promise<void> {
    this.pkg.zip.file(SHARED_STRINGS_PART, sstXml);
    // 表原本就有：关系与内容类型都在原文里，追加条目不改变包结构。
    if (this.originalParts.has(SHARED_STRINGS_PART)) return;
    const relId = this.nextRelIdFor(WORKBOOK_PART);
    this.pkg.zip.file(
      this.relsPathFor(WORKBOOK_PART),
      await this.relsWith(WORKBOOK_PART, relId, SHARED_STRINGS_REL_TYPE, SHARED_STRINGS_PART),
    );
    this.pkg.zip.file(
      "[Content_Types].xml",
      await this.contentTypesWith([[`/${SHARED_STRINGS_PART}`, SHARED_STRINGS_CONTENT_TYPE]]),
    );
  }

  private countStringReferences(): number {
    let total = 0;
    for (const sheet of this.loadedSheets.values()) {
      for (const cell of Object.values(sheet.cells)) {
        if (cell.kind === "string" && !cell.formula && !cell.rawXml) total += 1;
      }
    }
    return total;
  }
}

/** `<calcPr>` 之后还允许出现的元素（ECMA-376 Part 1 §18.2.7 的 CT_Workbook 序列），插入点取第一个命中的。 */
const AFTER_CALC_PR = ["oleSize", "customWorkbookViews", "pivotCaches", "smartTagPr", "smartTagTypes", "webPublishing", "fileRecoveryPr", "webPublishObjects", "extLst"];

/**
 * 给 `xl/workbook.xml` 打上"打开即全量重算"的标记。
 *
 * 为什么必须打：我们改一格之后，依赖它的公式里 `<v>` 缓存值就旧了，而 `calcChain.xml` 我们既不读也不维护；
 * 没有这个标记时 Excel 会信任旧缓存值显示出来（用户看到的是改之前的合计）。
 * `calcPr/@fullCalcOnLoad="1"` 是让消费方打开文件时重算全部公式的开关（ECMA-376 Part 1 §18.2.7）。
 *
 * @param workbookXml 原 `xl/workbook.xml` 文本。
 * @returns 带标记的文本：已有 `<calcPr>` 时只补这一个属性（其余属性原样留），没有时按 schema 顺序插在
 *   `<calcPr>` 该在的位置；属性已经是 1 时原样返回（重复保存不许叠加）。
 */
function withFullCalcOnLoad(workbookXml: string): string {
  if (/<calcPr\b[^>]*fullCalcOnLoad="1"/.test(workbookXml)) return workbookXml;
  const existing = /<calcPr\b([^>]*)\/>/.exec(workbookXml);
  if (existing) return workbookXml.replace(existing[0], `<calcPr${existing[1]} fullCalcOnLoad="1"/>`);
  const anchor = AFTER_CALC_PR.map((name) => workbookXml.indexOf(`<${name}`)).find((index) => index >= 0);
  const insertion = anchor ?? workbookXml.indexOf("</workbook>");
  return `${workbookXml.slice(0, insertion)}<calcPr fullCalcOnLoad="1"/>${workbookXml.slice(insertion)}`;
}

/**
 * 从部件文本里取元素：先按容器名下一层，再取其直接子元素。
 *
 * @param xmlText 部件文本。
 * @param containerName 容器元素名，如 `sheets`；传 undefined 表示直接在根下取。
 * @param name 目标子元素名，如 `sheet`。
 * @returns 元素列表；容器不存在时返回空列表。
 */
function childElements(xmlText: string, containerName: string | undefined, name: string): Element[] {
  const root = parseDocument(xmlText);
  const container = containerName ? children(root).find((node) => isElement(node) && node.name === containerName) : root;
  if (!container) return [];
  return children(container).filter((node) => isElement(node) && node.name === name);
}
