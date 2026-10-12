/**
 * `xl/styles.xml` 的读取与**只追加**改写。
 *
 * 实测（spec §二 P3）：`<c/@s>` 是 `cellXfs` 的**位置索引**，`s="2"` 与 `numFmtId="164"` 之间
 * 只靠顺序相连。所以任何"重排/去重/压缩样式表"的实现都会让用户已有格式整体错位——
 * 本模块因此只允许追加，绝不动既有索引。
 */
import type { FontSpec } from "./color.js";
import { argbToSheetColor, sheetColorToArgb } from "./color.js";
import { BUILTIN_DATE_TIME_NUMFMT_IDS, isDateFormatCode } from "./datetime.js";
import type { ResolvedNumberFormat } from "./model.js";
import {
  XML_DECLARATION,
  type Element,
  escapeAttr,
  firstChild,
  findChildren,
  insertInSchemaOrder,
  parseDocument,
  serialize,
  sheetElement,
} from "./xml.js";

/** `styleSheet` 子元素的固定顺序（ECMA-376 18.8.11 CT_Stylesheet 的序列）。插错位置 Excel 判文件损坏。 */
const STYLESHEET_CHILD_ORDER = [
  "numFmts",
  "fonts",
  "fills",
  "borders",
  "cellStyleXfs",
  "cellXfs",
  "cellStyles",
  "dxfs",
  "tableStyles",
  "colors",
  "extLst",
] as const;

/** 自定义格式 id 起点：小于它的是内建 id，不允许占用（ECMA-376 18.8.30）。 */
const FIRST_CUSTOM_NUMFMT_ID = 164;

/** 一条 `cellXfs/xf` 的组成引用。 */
export interface CellStyleParts {
  numFmtId?: number;
  /**
   * 自定义格式码（`numFmtId >= 164` 时用）。传了就先把格式码追加/复用到 `<numFmts>`，
   * 再把得到的 id 写进 `xf`；不传则只用 `numFmtId`（内建 id 不需要格式码）。
   */
  formatCode?: string;
  fontId?: number;
  fillId?: number;
  borderId?: number;
  font?: FontSpec;
}

/** 样式表：持有解析后的 DOM，改写时只在原 DOM 上追加，其余部分逐字保留。 */
export class StyleTable {
  private constructor(
    private readonly root: Element,
    private readonly numFmtCodes: Map<number, string>,
    private readonly numFmtIdByCode: Map<string, number>,
    private readonly cellXfs: Element[],
    private readonly fonts: Element[],
    readonly sourceXml: string,
    private dirty = false,
  ) {}

  /**
   * 解析 `styles.xml`。
   *
   * @param xmlText 部件文本；缺失时（极端最小包）传 undefined，得到一个带默认 0 号样式的空表。
   * @returns 样式表实例。
   */
  static parse(xmlText: string | undefined): StyleTable {
    const source =
      xmlText ??
      `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
        `<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>` +
        `<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>` +
        `<borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs>` +
        `<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>`;
    const root = parseDocument(source);
    const numFmtCodes = new Map<number, string>();
    const numFmtIdByCode = new Map<string, number>();
    const numFmts = firstChild(root, "numFmts");
    if (numFmts) {
      for (const el of findChildren(numFmts, "numFmt")) {
        const id = Number(el.attributes["numFmtId"]);
        const code = String(el.attributes["formatCode"] ?? "");
        if (Number.isInteger(id) && code) {
          numFmtCodes.set(id, code);
          numFmtIdByCode.set(code, id);
        }
      }
    }
    return new StyleTable(
      root,
      numFmtCodes,
      numFmtIdByCode,
      findChildren(firstChild(root, "cellXfs") ?? root, "xf"),
      findChildren(firstChild(root, "fonts") ?? root, "font"),
      source,
    );
  }

  /**
   * 取某个 `@s` 的数字格式。
   *
   * @param styleIndex `cellXfs` 位置索引（`<c/@s>` 原值）。
   * @returns 解析结果；索引越界返回 `{ numFmtId: 0, isDateLike: false }`，
   *   含义是"按 General 显示"，不替 Excel 猜格式。
   */
  numberFormatFor(styleIndex: number | undefined): ResolvedNumberFormat {
    const xf = this.cellXfs[styleIndex ?? -1];
    const numFmtId = xf ? Number(xf.attributes["numFmtId"] ?? 0) : 0;
    const formatCode = this.numFmtCodes.get(numFmtId);
    const isDateLike =
      BUILTIN_DATE_TIME_NUMFMT_IDS.has(numFmtId) || (formatCode !== undefined && isDateFormatCode(formatCode));
    return { numFmtId, formatCode, isDateLike };
  }

  /**
   * 取某个 `@s` 的字体标记。
   *
   * @param styleIndex `cellXfs` 位置索引。
   * @returns 读得懂的字体属性；没有对应字体时返回空对象。
   */
  fontFor(styleIndex: number | undefined): FontSpec {
    const xf = this.cellXfs[styleIndex ?? -1];
    const fontId = xf ? Number(xf.attributes["fontId"] ?? 0) : 0;
    return StyleTable.readFont(this.fonts[fontId]);
  }

  /** 已定义的 `cellXfs` 条数；单测用它断言"只追加不重排"。 */
  get cellXfCount(): number {
    return this.cellXfs.length;
  }

  /** 是否追加过条目；没追加过时 `toBytes()` 完全不重写 `styles.xml`，保证空改动逐字节一致。 */
  get isDirty(): boolean {
    return this.dirty;
  }

  /**
   * 追加一个自定义数字格式码。
   *
   * @param formatCode 如 `0.00%`；已有相同码时复用，不产生重复条目。
   * @returns `numFmtId`（已有值或 ≥164 的新值）。
   */
  appendNumberFormat(formatCode: string): number {
    const existing = this.numFmtIdByCode.get(formatCode);
    if (existing !== undefined) return existing;
    let nextId = FIRST_CUSTOM_NUMFMT_ID;
    for (const id of this.numFmtCodes.keys()) nextId = Math.max(nextId, id + 1);
    this.numFmtCodes.set(nextId, formatCode);
    this.numFmtIdByCode.set(formatCode, nextId);

    this.dirty = true;
    let container = firstChild(this.root, "numFmts");
    if (!container) {
      container = sheetElement("<numFmts/>");
      insertInSchemaOrder(this.root, container, STYLESHEET_CHILD_ORDER);
    }
    container.children?.push(
      sheetElement(`<numFmt numFmtId="${nextId}" formatCode="${escapeAttr(formatCode)}"/>`),
    );
    container.attributes["count"] = String(findChildren(container, "numFmt").length);
    return nextId;
  }

  /**
   * 追加一条单元格样式（`cellXfs/xf`）。
   *
   * @param parts 各引用；未给的按 0。传 `font` 时按需往 `fonts` 追加条目。
   * @returns 新的 `cellXfs` 位置索引（可直接写进 `<c/@s>`）；完全相同的样式复用已有索引，不产生重复。
   */
  appendCellStyle(parts: CellStyleParts): number {
    const fontId = parts.font ? this.ensureFont(parts.font) : (parts.fontId ?? 0);
    // 给了格式码就先把码落到 `<numFmts>`（已有同码则复用 id），否则 Excel 读不到自定义格式。
    const numFmtId = parts.formatCode !== undefined ? this.appendNumberFormat(parts.formatCode) : (parts.numFmtId ?? 0);
    const wanted = {
      numFmtId,
      fontId,
      fillId: parts.fillId ?? 0,
      borderId: parts.borderId ?? 0,
    };
    const signature = StyleTable.xfSignature(wanted);
    const existingIndex = this.cellXfs.findIndex((xf) => StyleTable.xfSignature(xf) === signature);
    if (existingIndex >= 0) return existingIndex;
    this.dirty = true;

    const attrs = [
      `numFmtId="${wanted.numFmtId}"`,
      `fontId="${wanted.fontId}"`,
      `fillId="${wanted.fillId}"`,
      `borderId="${wanted.borderId}"`,
      `xfId="0"`,
    ];
    if (parts.numFmtId !== undefined || parts.formatCode !== undefined) attrs.push('applyNumberFormat="1"');
    if (parts.font !== undefined) attrs.push('applyFont="1"');
    if (parts.fillId !== undefined) attrs.push('applyFill="1"');
    if (parts.borderId !== undefined) attrs.push('applyBorder="1"');
    const el = sheetElement(`<xf ${attrs.join(" ")}/>`);
    const container = firstChild(this.root, "cellXfs") ?? this.root;
    container.children?.push(el);
    this.cellXfs.push(el);
    container.attributes["count"] = String(this.cellXfs.length);
    return this.cellXfs.length - 1;
  }

  /**
   * 序列化回 `styles.xml`。
   *
   * @returns 带 XML 声明的完整部件文本。
   */
  toXml(): string {
    return `${XML_DECLARATION}${serialize(this.root)}`;
  }

  /**
   * 字体表里找同款，没有就追加。
   *
   * @param spec 字体属性。
   * @returns `fonts` 的位置索引。
   */
  private ensureFont(spec: FontSpec): number {
    const signature = StyleTable.fontSignature(spec);
    for (let i = 0; i < this.fonts.length; i += 1) {
      if (StyleTable.fontSignature(StyleTable.readFont(this.fonts[i])) === signature) return i;
    }
    const parts: string[] = [];
    if (spec.bold) parts.push("<b/>");
    if (spec.italic) parts.push("<i/>");
    if (spec.underline) parts.push("<u/>");
    if (spec.strike) parts.push("<strike/>");
    if (spec.sizePt !== undefined) parts.push(`<sz val="${spec.sizePt}"/>`);
    if (spec.name) parts.push(`<name val="${escapeAttr(spec.name)}"/>`);
    const argb = spec.color ? sheetColorToArgb(spec.color) : undefined;
    if (argb) parts.push(`<color rgb="${argb}"/>`);
    else if (spec.color?.themeRef) {
      const tint = spec.color.themeRef.tint;
      parts.push(`<color theme="${spec.color.themeRef.theme}"${tint === undefined ? "" : ` tint="${tint}"`}/>`);
    }
    this.dirty = true;
    const el = sheetElement(`<font>${parts.join("")}</font>`);
    const container = firstChild(this.root, "fonts") ?? this.root;
    container.children?.push(el);
    this.fonts.push(el);
    container.attributes["count"] = String(this.fonts.length);
    return this.fonts.length - 1;
  }

  private static readFont(el: Element | undefined): FontSpec {
    if (!el) return {};
    const spec: FontSpec = {};
    if (firstChild(el, "b")) spec.bold = true;
    if (firstChild(el, "i")) spec.italic = true;
    if (firstChild(el, "u")) spec.underline = true;
    if (firstChild(el, "strike")) spec.strike = true;
    const sz = firstChild(el, "sz")?.attributes["val"];
    if (sz !== undefined) spec.sizePt = Number(sz);
    const name = firstChild(el, "name")?.attributes["val"];
    if (name !== undefined) spec.name = String(name);
    const colorEl = firstChild(el, "color");
    const rgb = colorEl?.attributes["rgb"];
    const theme = colorEl?.attributes["theme"];
    if (rgb) spec.color = argbToSheetColor(String(rgb));
    else if (theme !== undefined) {
      const tint = colorEl?.attributes["tint"];
      spec.color = {
        hex: "",
        themeRef: { theme: Number(theme), tint: tint === undefined ? undefined : Number(tint) },
      };
    }
    return spec;
  }

  private static fontSignature(spec: FontSpec): string {
    const argb = spec.color ? (sheetColorToArgb(spec.color) ?? "") : "";
    const themeRef = spec.color?.themeRef ? `${spec.color.themeRef.theme}/${spec.color.themeRef.tint ?? ""}` : "";
    return `${spec.bold ? "b" : ""}${spec.italic ? "i" : ""}${spec.underline ? "u" : ""}${spec.strike ? "s" : ""}|${spec.sizePt ?? ""}|${spec.name ?? ""}|${argb}|${themeRef}`;
  }

  private static xfSignature(input: Element | { numFmtId: number; fontId: number; fillId: number; borderId: number }): string {
    const isElement = "attributes" in input;
    const numFmtId = isElement ? Number(input.attributes["numFmtId"] ?? 0) : input.numFmtId;
    const fontId = isElement ? Number(input.attributes["fontId"] ?? 0) : input.fontId;
    const fillId = isElement ? Number(input.attributes["fillId"] ?? 0) : input.fillId;
    const borderId = isElement ? Number(input.attributes["borderId"] ?? 0) : input.borderId;
    return `${numFmtId}|${fontId}|${fillId}|${borderId}`;
  }
}
