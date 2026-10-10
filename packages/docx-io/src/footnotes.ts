import type JSZip from "jszip";
import { children, parseDocument, serialize, wordPartPath } from "./xml.js";
import { FOOTNOTE_ID, ooxmlFragmentToPm, pmToOoxmlFragment, samePmFragment } from "./ooxml-pm.js";
import { emptyPmDoc } from "./pm-schema.js";
import type { PMDoc, PMNode } from "./pm-schema.js";
import type { NumberingMap } from "./numbering.js";
import type { StyleMap } from "./styles.js";

/**
 * 脚注与尾注部件（`word/footnotes.xml` / `word/endnotes.xml`）。
 *
 * ECMA-376 §17.11：正文里只有引用点 `<w:footnoteReference w:id="N"/>`，内容住在
 * 这个独立部件里，靠 id 对应。部件里除了正常条目还有两条特殊条目（`w:type` 为
 * `separator` 与 `continuationSeparator`），它们不是用户的脚注，而是脚注区那条横线
 * 和跨页时的续行记号——少了它们 Word 会自己补，但我们已经有的必须原样带走。
 *
 * 本模块只管"把部件读进来、把部件写回去、保持 id 对得上"。脚注正文的编辑入口在
 * 编辑器层，写回时未改动的条目按原始 XML 字节透传（AGENTS §2.7：判不了角色的块不产生改动）。
 */

export type FootnoteKind = "footnote" | "endnote";

/** 一条脚注/尾注条目。 */
export interface FootnoteEntry {
  /** `w:id`，ST_DecimalNumber 的字符串原样（分隔符条目可以是负数）。 */
  id: string;
  /** `w:type`：`separator` / `continuationSeparator`；正常脚注没有这个属性。 */
  role?: string;
  /** 条目正文，段落与正文同构；未改动时只用来读文字，写回走 `rawXml`。 */
  pm: PMDoc;
  /** 整个 `w:footnote` 元素的原始 XML，脏之前按字节写回。 */
  rawXml: string;
  /** true = 用户改过正文，写回时重新序列化。 */
  dirty?: boolean;
}

/** 两种部件的条目，按 id 对应正文里的引用点。 */
export interface FootnoteStore {
  footnote: FootnoteEntry[];
  endnote: FootnoteEntry[];
  /**
   * 有条目没能建模（id 不合 ST_DecimalNumber）的部件。这些部件整个原样保留、
   * 永不重写——按条目重建会把读不进来的那几条丢掉。
   */
  unmodeled: FootnoteKind[];
}

/** 一条脚注/尾注的正文，跨界用的形状：只有 id 与 PM 正文，不带部件里的原始 XML。 */
export interface FootnoteBody {
  /** `w:id`，与正文里的引用点同一个字符串。 */
  id: string;
  /** 条目正文。 */
  doc: PMDoc;
}

/** 两种部件的正文，按部件分组；没有正文的那种部件可以整个缺省。 */
export type FootnoteBodies = Partial<Record<FootnoteKind, FootnoteBody[]>>;

/** 引用点的一个实例，按正文出现顺序排列。 */
interface FootnoteRefOccurrence {
  kind: FootnoteKind;
  id: string;
}

export const FOOTNOTE_PART: Record<FootnoteKind, string> = {
  footnote: "word/footnotes.xml",
  endnote: "word/endnotes.xml",
};

export const FOOTNOTE_REL_TYPE: Record<FootnoteKind, string> = {
  footnote: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes",
  endnote: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes",
};

export const FOOTNOTE_CONTENT_TYPE: Record<FootnoteKind, string> = {
  footnote:
    "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
  endnote: "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml",
};

const ROOT_TAG: Record<FootnoteKind, string> = {
  footnote: "w:footnotes",
  endnote: "w:endnotes",
};
const ENTRY_TAG: Record<FootnoteKind, string> = {
  footnote: "w:footnote",
  endnote: "w:endnote",
};
/** 条目正文开头那个"编号记号"，Word 靠它在版面上排编号；引用它的是 `w:footnoteReference`。 */
const REF_MARK: Record<FootnoteKind, string> = {
  footnote: "<w:r><w:rPr><w:vertAlign w:val=\"superscript\"/></w:rPr><w:footnoteRef/></w:r>",
  endnote: "<w:r><w:rPr><w:vertAlign w:val=\"superscript\"/></w:rPr><w:endnoteRef/></w:r>",
};

/**
 * 读脚注与尾注部件。
 *
 * 部件路径从 `document.xml.rels` 里按关系类型找（Word 不保证叫 footnotes.xml，
 * 尽管它几乎总是）；rels 里没有、或部件文件不存在，都算"这份文档没有脚注"。
 *
 * @param zip 已加载的包。
 * @param relsXml `word/_rels/document.xml.rels` 的文本，缺省空串。
 * @param numbering 正文那份编号表，脚注里的列表要用。
 * @param styles 正文那份样式表，脚注段落常带 `FootnoteText` 样式。
 * @returns 两种部件的条目；文件损坏或 id 缺失的条目直接跳过，不猜。
 */
export async function loadFootnotes(
  zip: JSZip,
  relsXml: string,
  numbering: NumberingMap,
  styles: StyleMap,
): Promise<FootnoteStore> {
  const store: FootnoteStore = { footnote: [], endnote: [], unmodeled: [] };
  if (!relsXml || !/<Relationships\b/.test(relsXml)) return store;
  const relsRoot = parseDocument(relsXml);
  const byType = new Map<string, string>();
  for (const rel of children(relsRoot).filter((c) => c.name === "Relationship")) {
    const type = rel.attributes?.Type as string | undefined;
    const target = rel.attributes?.Target as string | undefined;
    if (type && target) byType.set(type, target);
  }

  for (const kind of ["footnote", "endnote"] as FootnoteKind[]) {
    const target = byType.get(FOOTNOTE_REL_TYPE[kind]);
    if (!target) continue;
    // rels 的 Target 相对 word/ 解析。
    const xml = await zip.file(wordPartPath(target))?.async("string");
    if (!xml) continue;
    const root = parseDocument(xml);
    let dropped = false;
    for (const el of children(root)) {
      if (el.name !== ENTRY_TAG[kind]) continue;
      const id = el.attributes?.["w:id"] as string | undefined;
      // id 不是 ST_DecimalNumber 就记下来：正文里的引用点本来也配不上它，
      // 但这个部件从此不能按条目重建，否则这一条就消失了。
      if (!id || !FOOTNOTE_ID.test(id)) {
        dropped = true;
        continue;
      }
      store[kind].push({
        id,
        role: el.attributes?.["w:type"] as string | undefined,
        pm: ooxmlFragmentToPm(el, numbering, styles),
        rawXml: serialize(el),
      });
    }
    if (dropped) store.unmodeled.push(kind);
  }
  return store;
}

/**
 * 收集一篇正文引用了哪些脚注与尾注。
 *
 * 引用点是唯一的真相来源：正文里没有的 id 就不该在部件里留条目，反之缺的条目要补。
 *
 * @param doc 正文 PM doc（不含脚注部件自己的内容）。
 * @returns 两种部件各自的引用 id 集合；正文一个引用都没有时为空集合。
 */
export function collectFootnoteRefs(doc: PMDoc): {
  footnote: Set<string>;
  endnote: Set<string>;
} {
  const out = { footnote: new Set<string>(), endnote: new Set<string>() };
  for (const ref of footnoteRefOccurrences(doc)) out[ref.kind].add(ref.id);
  return out;
}

/**
 * 按正文顺序列出所有引用点。
 *
 * 对齐条目要的是这份顺序，遍历因此只写这一处。id 不合 `ST_DecimalNumber` 的引用点被跳过——那种引用在部件里
 * 没有条目可读，读侧已把整个部件标成 `unmodeled`。
 *
 * @param doc 正文 PM doc。
 * @returns 引用点列表；同一 id 被引用两次就有两条。
 */
function footnoteRefOccurrences(doc: PMDoc): FootnoteRefOccurrence[] {
  const out: FootnoteRefOccurrence[] = [];
  const walk = (nodes: readonly PMNode[]): void => {
    for (const n of nodes) {
      if (n.type === "footnoteReference" || n.type === "endnoteReference") {
        const id = String(n.attrs.id ?? "");
        if (FOOTNOTE_ID.test(id))
          out.push({ kind: n.type === "footnoteReference" ? "footnote" : "endnote", id });
      }
      if ("content" in n) walk((n.content ?? []) as PMNode[]);
    }
  };
  walk(doc.content ?? []);
  return out;
}

/**
 * 给新插入的脚注/尾注分一个 id。
 *
 * 取现有正常条目里最大数字 id 加一。`separator`（负数）与 `continuationSeparator`（0）
 * 是版式件，不参与取最大值；一个正常条目都没有时从 1 起，与 Word 新建文档的结果一致。
 *
 * @param entries 该部件的现有条目（只要 `id` 与 `role`，正文不必参与）。
 * @returns 可以用的 id（字符串）。
 */
export function nextFootnoteId(
  entries: readonly { id: string; role?: string }[],
): string {
  let max = 0;
  for (const entry of entries) {
    if (entry.role) continue;
    const value = Number(entry.id);
    if (Number.isFinite(value) && value > max) max = value;
  }
  return String(max + 1);
}

/**
 * 把编辑器改过的正文并回 store。
 *
 * 认不出的 id 也照样收下（作为新条目）：正文里的引用点是刚插的，部件里还没有条目，
 * 下一次保存 `syncFootnoteIds` 会按引用对齐。分隔符条目不接受编辑。
 *
 * @param store 当前条目（原地不改）。
 * @param kind 脚注还是尾注。
 * @param bodies 编辑器交回来的正文，按 id 对应。
 * @returns 并好正文的新 store；被标为 `unmodeled` 的部件原样返回，一个字节都不动。
 */
export function applyFootnoteBodies(
  store: FootnoteStore,
  kind: FootnoteKind,
  bodies: readonly FootnoteBody[],
): FootnoteStore {
  if (store.unmodeled.includes(kind)) return store;
  const byId = new Map(bodies.map((body) => [body.id, body]));
  const known = new Set(store[kind].map((entry) => entry.id));
  const merged: FootnoteEntry[] = store[kind].map((entry) => {
    const body = entry.role ? undefined : byId.get(entry.id);
    if (!body) return entry;
    // 文字没变就不标 dirty：下次保存继续按原始 XML 字节透传，保住未建模的属性。
    if (samePmFragment(entry.pm, body.doc)) return entry;
    return { ...entry, pm: body.doc, dirty: true };
  });
  for (const body of bodies) {
    if (known.has(body.id)) continue;
    merged.push({ id: body.id, pm: body.doc, rawXml: "", dirty: true });
  }
  const out: FootnoteStore = {
    footnote: store.footnote,
    endnote: store.endnote,
    unmodeled: store.unmodeled,
  };
  out[kind] = merged;
  return out;
}

/**
 * 把正文里还存在的引用 id 与部件条目对齐：删掉没人引用的正常脚注，
 * 给新引用补一条空正文的条目。
 *
 * 只对齐正常条目：`separator` / `continuationSeparator` 是脚注区的版式件，
 * 正文永远不会引用它们，删了会让 Word 重排整个脚注区。
 *
 * @param store 当前条目（原地不改）。
 * @param referenced 正文里出现过的引用 id 集合。
 * @returns 对齐后的新 store。
 */
export function syncFootnoteIds(
  store: FootnoteStore,
  referenced: { footnote: Set<string>; endnote: Set<string> },
): FootnoteStore {
  const out: FootnoteStore = { footnote: [], endnote: [], unmodeled: store.unmodeled };
  for (const kind of ["footnote", "endnote"] as FootnoteKind[]) {
    if (store.unmodeled.includes(kind)) {
      // 这个部件不会被重写，所以也不动它的条目。
      out[kind] = [...store[kind]];
      continue;
    }
    const keep = referenced[kind];
    const entries: FootnoteEntry[] = [];
    for (const entry of store[kind]) {
      if (!entry.role && !keep.has(entry.id)) continue;
      entries.push(entry);
    }
    for (const id of keep) {
      if (entries.some((e) => e.id === id)) continue;
      entries.push({ id, pm: emptyPmDoc(), rawXml: "", dirty: true });
    }
    out[kind] = entries;
  }
  return out;
}

/**
 * 序列化整个脚注（或尾注）部件。
 *
 * 未改动的条目按原始 XML 透传，改过的重新走一遍正文那套转换器，并在第一段开头
 * 补回编号记号：那个记号不属于正文模型（读的时候就被忽略），不补的话 Word 会渲染出
 * 一条没有编号的脚注。
 *
 * @param kind 脚注还是尾注。
 * @param entries 该部件的全部条目，顺序即 Word 写出的顺序（分隔符在前）。
 * @param numbering 正文那份编号表。
 * @returns 完整的部件 XML；一条条目都没有时返回空串，调用方据此不写部件。
 */
export function serializeFootnotePart(
  kind: FootnoteKind,
  entries: readonly FootnoteEntry[],
  numbering: NumberingMap,
): string {
  if (!entries.length) return "";
  const body = entries
    .map((entry) => {
      // 没改过的整条原样带走：连 `w:footnote` 自己身上那些没建模的属性一起。
      if (!entry.dirty) return entry.rawXml;
      const inner = withRefMark(kind, pmToOoxmlFragment(entry.pm, { numbering }));
      const role = entry.role ? ` w:type="${entry.role}"` : "";
      return `<${ENTRY_TAG[kind]} w:id="${entry.id}"${role}>${inner}</${ENTRY_TAG[kind]}>`;
    })
    .join("");
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<${ROOT_TAG[kind]} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
    `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
    body +
    `</${ROOT_TAG[kind]}>`
  );
}

/**
 * 在条目正文的第一段开头补回编号记号。
 *
 * 记号要排在段落自己的 `w:pPr` 之后（`EG_PContent` 在 `EG_PPr` 之后才出现），
 * 第一段是自闭合 `<w:p/>`（新建的脚注还没有正文）时把它展开。
 *
 * @param kind 脚注还是尾注，决定记号用 `w:footnoteRef` 还是 `w:endnoteRef`。
 * @param inner 序列化好的正文片段，可能为空。
 * @returns 带记号的正文片段。
 */
function withRefMark(kind: FootnoteKind, inner: string): string {
  const mark = REF_MARK[kind];
  const open = inner.match(/^<w:p(?: [^>]*)?(\/?)>/);
  if (!open || open[1]) return `<w:p>${mark}</w:p>` + inner.slice(open ? open[0].length : 0);
  const after = inner.slice(open[0].length);
  const props = after.match(/^<w:pPr>[\s\S]*?<\/w:pPr>/);
  const at = open[0].length + (props ? props[0].length : 0);
  return inner.slice(0, at) + mark + inner.slice(at);
}
