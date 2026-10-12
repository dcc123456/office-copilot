/**
 * `.xlsx` 文件的出入口（spec D1 的 `packages/sheet-io` 门面）。
 *
 * 三条纪律：
 * 1. 没被我们改写的部件在 `toBytes()` 里逐字节原样回写（spec §零.6 / 验收 A1）。
 * 2. 单位/颜色/序列号的解释只发生在本包（`AGENTS.md` §2.1 的同一条纪律，换成表格域）。
 * 3. 读不懂的东西记进 `warnings`/`unsupported` 一路带到 UI，不许静默丢弃。
 */
import { isWorksheetRelationship, loadPackage, readText, resolveTarget, type OpcPackage } from "./opc.js";
import { children, isElement, parseDocument, type Element } from "./xml.js";
import type { SheetModel, UnsupportedFeature, WorkbookModel } from "./model.js";
import { SharedStringTable } from "./xlsx-strings.js";
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

    const workbookXml = await readText(pkg.zip, "xl/workbook.xml");
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

    const relationships = pkg.relsBySourcePart.get("xl/workbook.xml") ?? [];
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
        partPath: resolveTarget("xl/workbook.xml", rel.target),
        hidden: String(el.attributes["state"] ?? "visible") === "hidden",
      });
    }

    const sharedStrings = SharedStringTable.parse(await readText(pkg.zip, "xl/sharedStrings.xml"));
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
    const sstText = await readText(pkg.zip, "xl/sharedStrings.xml");
    if (sstText) file.originalParts.set("xl/sharedStrings.xml", sstText);
    const stylesText = await readText(pkg.zip, "xl/styles.xml");
    if (stylesText) file.originalParts.set("xl/styles.xml", stylesText);
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
   * 把上层表达的"显示意图"收敛成 `cellXfs` 位置索引（只追加或复用，**绝不重排**，实测 §二 P3）。
   *
   * 渲染层只带 `numberFormat`（它在 IO 层读到的语义），不自己数样式索引；
   * 已经是目标格式的格原样带过（`s` 不动），所以在表里改一个数字格式不会牵连其它格。
   * 读不懂的格（`rawXml` 原样回写的）不碰。
   *
   * @param sheet 待写入的表；就地补 `cell.styleIndex`。
   */
  private resolveNumberFormatStyles(sheet: SheetModel): void {
    for (const cell of Object.values(sheet.cells)) {
      if (!cell.numberFormat || cell.rawXml) continue;
      const current = this.styles.numberFormatFor(cell.styleIndex);
      if (current.numFmtId === cell.numberFormat.numFmtId && current.formatCode === cell.numberFormat.formatCode) continue;
      cell.styleIndex = this.styles.appendCellStyle(
        cell.numberFormat.formatCode !== undefined
          ? { formatCode: cell.numberFormat.formatCode }
          : { numFmtId: cell.numberFormat.numFmtId },
      );
      cell.numberFormat = this.styles.numberFormatFor(cell.styleIndex);
    }
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
    if (this.sharedStrings.addedCount > 0) {
      const allLoaded = this.loadedSheets.size === this.refs.length;
      if (allLoaded) this.sharedStrings.setUsageCount(this.countStringReferences());
      const sstXml = this.sharedStrings.toXml();
      if (sstXml) this.pkg.zip.file("xl/sharedStrings.xml", sstXml);
    }
    if (this.styles.isDirty) this.pkg.zip.file("xl/styles.xml", this.styles.toXml());
    return this.pkg.zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
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
