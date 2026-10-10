import type JSZip from "jszip";
import { children, parseDocument, serialize, wordPartPath } from "./xml.js";
import {
  ANNOTATION_ID,
  escapeXml,
  ooxmlFragmentToPm,
  pmToOoxmlFragment,
  samePmFragment,
} from "./ooxml-pm.js";
import { emptyPmDoc } from "./pm-schema.js";
import type { PMDoc, PMNode } from "./pm-schema.js";
import type { NumberingMap } from "./numbering.js";
import type { StyleMap } from "./styles.js";

/**
 * 批注部件（`word/comments.xml`）。
 *
 * 与脚注同一套分工：正文里只有三个锚点（`commentRangeStart` / `commentRangeEnd` /
 * `commentReference`，都只带 `w:id`），批注的**内容**住在这个独立部件里，
 * `CT_Comment` 要求 `w:author` 必填，`w:date` 与 `w:initials` 可选，正文可以是整段。
 *
 * `w15:commentEx`（已解决状态）、`w15:people`、`w16cid:commentId` 都是微软扩展，
 * 住在别的部件里：我们从不重写那些部件，JSZip 会把它们原样带走。第三方生产者省略它们会
 * 怎样没有官方说明（spec 排名 4），所以我们也不主动生成。
 *
 * 本模块只管"把部件读进来、把部件写回去、保持 id 对得上"。未改动的条目按原始 XML 字节
 * 透传（AGENTS §2.7：判不了角色的块不产生改动）。
 */

/** 一条批注条目。 */
export interface CommentEntry {
  /** `w:id`，ST_DecimalNumber 的字符串原样（与书签同一个编号空间）。 */
  id: string;
  /** `w:author`，规范里必填；读到的条目没有就留占位名。 */
  author: string;
  /** `w:initials`，可选。 */
  initials?: string;
  /** `w:date`，可选，Word 写的是 ISO 8601（`2026-01-02T10:30:00Z`），原样存。 */
  date?: string;
  /** 批注正文，段落与正文同构；未改动时只用来读文字，写回走 `rawXml`。 */
  pm: PMDoc;
  /** 整个 `w:comment` 元素的原始 XML，脏之前按字节写回。 */
  rawXml: string;
  /** true = 用户改过正文或元信息，写回时重新序列化。 */
  dirty?: boolean;
}

/** 批注部件的全部条目，按 id 对应正文里的锚点。 */
export interface CommentStore {
  entries: CommentEntry[];
  /**
   * 有条目没能建模（`w:id` 不合判据）时为 true：整个部件从此原样保留、永不重写。
   * 按条目重建会把读不进来的那几条丢掉。
   */
  isUnmodeled: boolean;
  /**
   * 部件在包里的实际路径（从 rels 的 Target 解析出来，可能是 `word/comments.xml` 之外的名字）。
   * `undefined` = 这份文档没有批注部件，新建时用默认路径。
   */
  partPath?: string;
}

/** 编辑器交回来的批注正文；作者等元信息可以一并改。 */
export interface CommentBody {
  /** `w:id`，与正文里的锚点同一个字符串。 */
  id: string;
  /** 批注正文。 */
  doc: PMDoc;
  /** `w:author`，缺省保留原条目里的名字。 */
  author?: string;
  /** `w:initials`，缺省保留原值。 */
  initials?: string;
  /** `w:date`，缺省保留原值。 */
  date?: string;
}

export const COMMENT_PART = "word/comments.xml";

export const COMMENT_REL_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";

export const COMMENT_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";

const ROOT_TAG = "w:comments";
const ENTRY_TAG = "w:comment";

/**
 * `w:author` 是必填属性，而新建一条批注时编辑器可能还没给出作者名（作者名要到
 * 排名 4 ⑤ 才接设置里的用户）。留一个稳定的占位名，比写空串更接近 Word 自己的产物。
 */
export const FALLBACK_COMMENT_AUTHOR = "Office Copilot";

/**
 * 读批注部件。
 *
 * 部件路径从 `document.xml.rels` 里按关系类型找（Word 不保证叫 comments.xml，尽管它几乎总是）；
 * rels 里没有、或部件文件不存在，都算"这份文档没有批注"。
 *
 * @param zip 已加载的包。
 * @param relsXml `word/_rels/document.xml.rels` 的文本，缺省空串。
 * @param numbering 正文那份编号表，批注里的列表要用。
 * @param styles 正文那份样式表，批注段落常带 `CommentText` 样式。
 * @returns 部件的条目；`w:id` 缺失或不合法的条目被跳过，同时把整个部件标为不可重写。
 */
export async function loadComments(
  zip: JSZip,
  relsXml: string,
  numbering: NumberingMap,
  styles: StyleMap,
): Promise<CommentStore> {
  const store: CommentStore = { entries: [], isUnmodeled: false };
  if (!relsXml || !/<Relationships\b/.test(relsXml)) return store;
  const relsRoot = parseDocument(relsXml);
  let target: string | undefined;
  for (const rel of children(relsRoot).filter((c) => c.name === "Relationship")) {
    if (rel.attributes?.Type === COMMENT_REL_TYPE)
      target = rel.attributes?.Target as string | undefined;
  }
  if (!target) return store;
  const partPath = wordPartPath(target);
  const xml = await zip.file(partPath)?.async("string");
  if (!xml) return store;
  store.partPath = partPath;
  const root = parseDocument(xml);
  for (const el of children(root)) {
    if (el.name !== ENTRY_TAG) continue;
    const id = el.attributes?.["w:id"] as string | undefined;
    // id 不合判据就记下来：正文里的锚点本来也配不上它，但这个部件从此不能按条目重建，
    // 否则这一条就消失了。
    if (!id || !ANNOTATION_ID.test(id)) {
      store.isUnmodeled = true;
      continue;
    }
    store.entries.push({
      id,
      author: (el.attributes?.["w:author"] as string | undefined) ?? FALLBACK_COMMENT_AUTHOR,
      initials: el.attributes?.["w:initials"] as string | undefined,
      date: el.attributes?.["w:date"] as string | undefined,
      pm: ooxmlFragmentToPm(el, numbering, styles),
      rawXml: serialize(el),
    });
  }
  return store;
}

/**
 * 收集一篇正文引用了哪些批注 id。
 *
 * 三种锚点都算引用：只写了区间两端、没有引用记号的文档（别的生成器会这样）里的批注
 * 必须留下，否则按引用对齐会把条目删掉。
 *
 * @param doc 正文 PM doc。
 * @returns 出现过的批注 id 集合；正文一个锚点都没有时为空集合。
 */
export function collectCommentRefs(doc: PMDoc): Set<string> {
  const out = new Set<string>();
  const walk = (nodes: readonly PMNode[]): void => {
    for (const n of nodes) {
      if (
        n.type === "commentRangeStart" ||
        n.type === "commentRangeEnd" ||
        n.type === "commentReference"
      ) {
        const id = String(n.attrs.id ?? "");
        if (ANNOTATION_ID.test(id)) out.add(id);
      }
      if ("content" in n) walk((n.content ?? []) as PMNode[]);
    }
  };
  walk(doc.content ?? []);
  return out;
}

/**
 * 给新插入的批注分一个 id。
 *
 * 取现有条目里最大数字 id 加一；一条批注都没有时返回 `0`，与规范样例里第一条批注的
 * `w:id="0"` 一致（脚注那边要从 1 起是因为部件里占了 -1 与 0，这里没有那种版式件）。
 * 只在批注内部取最大值：书签也用 `w:id`，但两者的编号是否必须互不重复没有官方说明，
 * 与 Word 实际写出的文件一样按各自部件各算各的。
 *
 * @param entries 现有条目（只要 `id`，正文不必参与）。
 * @returns 可以用的 id（字符串）。
 */
export function nextCommentId(entries: readonly { id: string }[]): string {
  let max = -1;
  for (const entry of entries) {
    const value = Number(entry.id);
    if (Number.isFinite(value) && value > max) max = value;
  }
  return String(max + 1);
}

/**
 * 把编辑器改过的批注并回 store。
 *
 * 认不出的 id 也照样收下（作为新条目）：正文里的锚点是刚插的，部件里还没有条目。
 *
 * @param store 当前条目（原地不改）。
 * @param bodies 编辑器交回的正文与元信息，按 id 对应。
 * @returns 并好正文的新 store；被标为 `isUnmodeled` 的部件原样返回，一个字节都不动。
 */
export function applyCommentBodies(
  store: CommentStore,
  bodies: readonly CommentBody[],
): CommentStore {
  if (store.isUnmodeled) return store;
  const byId = new Map(bodies.map((body) => [body.id, body]));
  const known = new Set(store.entries.map((entry) => entry.id));
  const merged: CommentEntry[] = store.entries.map((entry) => {
    const body = byId.get(entry.id);
    if (!body) return entry;
    const next: CommentEntry = {
      ...entry,
      pm: body.doc,
      author: body.author ?? entry.author,
      initials: body.initials ?? entry.initials,
      date: body.date ?? entry.date,
    };
    // 正文与元信息都没变就不标 dirty：下次保存继续按原始 XML 字节透传，保住未建模的属性。
    if (
      samePmFragment(entry.pm, body.doc) &&
      next.author === entry.author &&
      next.initials === entry.initials &&
      next.date === entry.date
    )
      return entry;
    return { ...next, dirty: true };
  });
  for (const body of bodies) {
    if (known.has(body.id)) continue;
    merged.push({
      id: body.id,
      author: body.author ?? FALLBACK_COMMENT_AUTHOR,
      initials: body.initials,
      date: body.date,
      pm: body.doc,
      rawXml: "",
      dirty: true,
    });
  }
  return { entries: merged, isUnmodeled: false, partPath: store.partPath };
}

/**
 * 把正文里还存在的锚点 id 与部件条目对齐：删掉没人引用的批注，给新锚点补一条空正文的条目。
 *
 * @param store 当前条目（原地不改）。
 * @param referenced 正文里出现过的批注 id 集合。
 * @returns 对齐后的新 store。
 */
export function syncCommentIds(store: CommentStore, referenced: Set<string>): CommentStore {
  if (store.isUnmodeled) return store;
  const entries = store.entries.filter((entry) => referenced.has(entry.id));
  for (const id of referenced) {
    if (entries.some((entry) => entry.id === id)) continue;
    entries.push({
      id,
      author: FALLBACK_COMMENT_AUTHOR,
      pm: emptyPmDoc(),
      rawXml: "",
      dirty: true,
    });
  }
  return { entries, isUnmodeled: false, partPath: store.partPath };
}

/**
 * 序列化整个批注部件。
 *
 * 未改动的条目按原始 XML 透传（连 `w:comment` 自己身上那些没建模的属性一起），
 * 改过的重新走一遍正文那套转换器。批注正文不需要像脚注那样补编号记号——
 * Word 给批注框排版的是外层部件，条目里没有 `w:commentRef` 这种东西。
 *
 * @param entries 全部条目，顺序即 Word 写出的顺序。
 * @param numbering 正文那份编号表。
 * @returns 完整的部件 XML；一条条目都没有时返回空串，调用方据此不写部件。
 */
export function serializeCommentsPart(
  entries: readonly CommentEntry[],
  numbering: NumberingMap,
): string {
  if (!entries.length) return "";
  const body = entries
    .map((entry) => {
      if (!entry.dirty) return entry.rawXml;
      const attrs =
        `w:id="${entry.id}" w:author="${escapeXml(entry.author)}"` +
        (entry.date ? ` w:date="${escapeXml(entry.date)}"` : "") +
        (entry.initials ? ` w:initials="${escapeXml(entry.initials)}"` : "");
      return `<${ENTRY_TAG} ${attrs}>${pmToOoxmlFragment(entry.pm, { numbering })}</${ENTRY_TAG}>`;
    })
    .join("");
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<${ROOT_TAG} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
    `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
    body +
    `</${ROOT_TAG}>`
  );
}
