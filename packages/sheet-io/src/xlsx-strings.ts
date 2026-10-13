/**
 * `xl/sharedStrings.xml`（字符串表）的读取与**只追加**改写。
 *
 * 实测（spec §二 P2）：`<sst>` 上 `count="12" uniqueCount="11"` 是两个不同的数——
 * 前者是**引用次数**、后者才是 `<si>` 条目数；我们追加条目时只能让两者各自增加。
 * 未解释的子元素（`<rPh>` 注音等）靠 `rawSi` 原文回写，不许丢。
 */
import type { CellRun, SharedStringEntry } from "./model.js";
import { argbToSheetColor, sheetColorToArgb } from "./color.js";
import {
  SPREADSHEET_NS,
  XML_DECLARATION,
  type Element,
  escapeAttr,
  findChildren,
  isElement,
  parseDocument,
  serialize,
  text,
} from "./xml.js";

/** 字符串表。`entries` 的下标就是 `<c/@t="s">` 里 `<v>` 的取值。 */
export class SharedStringTable {
  private readonly entries: SharedStringEntry[] = [];
  /** 原 `@count`，追加时按新增条目数同步增加（引用数只能由调用方补充增量）。 */
  private usageCount: number;
  private readonly textIndex = new Map<string, number>();
  /** 富文本条目按 run 指纹复用：渲染层过来的载荷没有 `sharedStringIndex`，只靠它就每保存一次追加一条。 */
  private readonly runIndex = new Map<string, number>();
  private addedDelta = 0;

  private constructor() {
    this.usageCount = 0;
  }

  /**
   * 解析字符串表。
   *
   * @param xmlText 部件文本；部件不存在（表里没有字符串）时传 undefined，得到空表。
   * @returns 字符串表。
   */
  static parse(xmlText: string | undefined): SharedStringTable {
    const table = new SharedStringTable();
    if (!xmlText) return table;
    const root = parseDocument(xmlText);
    table.usageCount = Number(root.attributes["count"] ?? 0);
    for (const si of findChildren(root, "si")) {
      const entry = SharedStringTable.readEntry(si);
      const index = table.entries.length;
      table.entries.push(entry);
      if (!entry.runs) {
        if (!table.textIndex.has(entry.text)) table.textIndex.set(entry.text, index);
        continue;
      }
      const key = runSignature(entry.runs);
      if (!table.runIndex.has(key)) table.runIndex.set(key, index);
    }
    return table;
  }

  /** 条目总数（写回时对应 `@uniqueCount`）。 */
  get size(): number {
    return this.entries.length;
  }

  /** 本次加载以来新增的条目数。 */
  get addedCount(): number {
    return this.addedDelta;
  }

  /**
   * 按下标取条目。
   *
   * @param index `<v>` 里的索引。
   * @returns 条目；越界返回 undefined，调用方按"读不懂的引用"出 warning。
   */
  at(index: number): SharedStringEntry | undefined {
    return this.entries[index];
  }

  /**
   * 取或追加一个字符串索引。
   *
   * @param value 要写进表的文本。
   * @param runs 富文本 run；给了就按 run 生成 `<si><r>…`，不走纯文本复用。
   * @returns 可写进 `<v>` 的索引。纯文本相同（且原条目无 run）、或 run 逐字段相同，都复用原条目不追加。
   */
  intern(value: string, runs?: CellRun[]): number {
    if (!runs?.length) {
      const existing = this.textIndex.get(value);
      if (existing !== undefined) return existing;
    } else {
      const rich = this.runIndex.get(runSignature(runs));
      if (rich !== undefined) return rich;
    }
    const index = this.entries.length;
    const entry: SharedStringEntry = runs?.length ? { text: runs.map((run) => run.text).join(""), runs } : { text: value };
    this.entries.push(entry);
    if (entry.runs) this.runIndex.set(runSignature(entry.runs), index);
    else this.textIndex.set(value, index);
    this.addedDelta += 1;
    return index;
  }

  /**
   * 该占哪个索引：读入时那个条目内容没变就沿用它，否则追加。
   *
   * 为什么单独有这一步：`<si>` 原文（`rawSi`）是回写时保格式的唯一凭据，沿用了原索引就沿用了原条目，
   * 富文本的颜色/字体不必由我们重新序列化；换成 intern 则会追加一条新 `<si>`，
   * 每保存一次多一条重复条目，且格式只剩我们能解释的那部分。
   *
   * @param value 要写进表的文本（富文本为各 run 拼接后的串）。
   * @param runs 格子里的富文本 run；与条目逐字段相同才算同一个字符串。
   * @param preferred 读入时的原索引，undefined 表示这格不是从字符串表读来的。
   * @returns 可写进 `<v>` 的索引。
   */
  reuseOrAppend(value: string, runs?: CellRun[], preferred?: number): number {
    if (preferred !== undefined) {
      const entry = this.entries[preferred];
      if (entry && entry.text === value && runSignature(entry.runs) === runSignature(runs)) return preferred;
    }
    return this.intern(value, runs);
  }

  /**
   * 设定写回时的 `@count`（引用总数）。
   *
   * @param count 由调用方在全部表都已解析后数出来的 `t="s"` 格数；
   *   没有全部解析时传原值加新增量，保证 count 只会随真实引用走。
   */
  setUsageCount(count: number): void {
    this.usageCount = count;
  }

  /**
   * 给上层读取的条目快照。
   *
   * @returns 条目的浅拷贝数组（`rawSi` 不外泄给渲染层，避免 UI 直接改 XML）。
   */
  entriesForRead(): SharedStringEntry[] {
    return this.entries.map((entry) => ({ text: entry.text, runs: entry.runs }));
  }

  /**
   * 序列化回 `sharedStrings.xml`。
   *
   * @returns 完整部件文本；空表时返回 `undefined`（该部件本就可以不存在）。
   */
  toXml(): string | undefined {
    if (this.entries.length === 0) return undefined;
    const body = this.entries.map((entry) => entry.rawSi ?? SharedStringTable.writeEntry(entry)).join("");
    return `${XML_DECLARATION}<sst xmlns="${SPREADSHEET_NS}" count="${this.usageCount}" uniqueCount="${this.entries.length}">${body}</sst>`;
  }

  private static readEntry(si: Element): SharedStringEntry {
    const runs = findChildren(si, "r");
    const entry: SharedStringEntry = { text: text(si), rawSi: serialize(si) };
    if (runs.length > 1) entry.runs = runs.map(SharedStringTable.readRun);
    return entry;
  }

  private static readRun(r: Element): CellRun {
    const runText = findChildren(r, "t").map((tEl) => text(tEl)).join("");
    const run: CellRun = { text: runText };
    const rPr = r.children?.find((n) => isElement(n) && n.name === "rPr") as Element | undefined;
    if (!rPr) return run;
    if (rPr.children?.some((n) => isElement(n) && n.name === "b")) run.bold = true;
    if (rPr.children?.some((n) => isElement(n) && n.name === "i")) run.italic = true;
    if (rPr.children?.some((n) => isElement(n) && n.name === "u")) run.underline = true;
    if (rPr.children?.some((n) => isElement(n) && n.name === "strike")) run.strike = true;
    const sz = rPr.children?.find((n) => isElement(n) && n.name === "sz") as Element | undefined;
    if (sz?.attributes["val"] !== undefined) run.sizePt = Number(sz.attributes["val"]);
    const name = rPr.children?.find((n) => isElement(n) && n.name === "rFont") as Element | undefined;
    if (name?.attributes["val"] !== undefined) run.name = String(name.attributes["val"]);
    const color = rPr.children?.find((n) => isElement(n) && n.name === "color") as Element | undefined;
    const rgb = color?.attributes["rgb"];
    if (rgb) run.color = argbToSheetColor(String(rgb));
    return run;
  }

  private static writeEntry(entry: SharedStringEntry): string {
    if (!entry.runs?.length) {
      const preserve = entry.text !== entry.text.trim() ? ' xml:space="preserve"' : "";
      return `<si><t${preserve}>${escapeText(entry.text)}</t></si>`;
    }
    const body = entry.runs
      .map((run) => {
        const marks: string[] = [];
        if (run.bold) marks.push("<b/>");
        if (run.italic) marks.push("<i/>");
        if (run.underline) marks.push("<u/>");
        if (run.strike) marks.push("<strike/>");
        if (run.sizePt !== undefined) marks.push(`<sz val="${run.sizePt}"/>`);
        if (run.name) marks.push(`<rFont val="${escapeAttr(run.name)}"/>`);
        // 颜色必须跟着走：把红字写成黑字就是拍平局部格式，`AGENTS.md` §2.3 不许。
        const argb = run.color ? sheetColorToArgb(run.color) : undefined;
        if (argb) marks.push(`<color rgb="${argb}"/>`);
        const rPr = marks.length ? `<rPr>${marks.join("")}</rPr>` : "";
        const preserve = run.text !== run.text.trim() ? ' xml:space="preserve"' : "";
        return `<r>${rPr}<t${preserve}>${escapeText(run.text)}</t></r>`;
      })
      .join("");
    return `<si>${body}</si>`;
  }
}

/**
 * run 数组的指纹：逐字段列出 `writeEntry` 会写进 `<rPr>` 与 `<t>` 的全部内容。
 *
 * @param runs 格子里的富文本 run；纯文本条目传 undefined（签名为空串）。
 * @returns 可直接比较的字符串。只有"序列化出来一模一样"的两段才算同一个字符串。
 */
function runSignature(runs?: CellRun[]): string {
  return (runs ?? [])
    .map((run) =>
      JSON.stringify([
        run.text,
        run.bold ?? false,
        run.italic ?? false,
        run.underline ?? false,
        run.strike ?? false,
        run.sizePt ?? null,
        run.name ?? null,
        run.color ? (sheetColorToArgb(run.color) ?? null) : null,
      ]),
    )
    .join(",");
}

function escapeText(value: string): string {
  return value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}
