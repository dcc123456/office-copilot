import JSZip from "jszip";
import {
  children,
  collectDescendants,
  escapeAttr,
  escapeText,
  firstChild,
  parseDocument,
  parseFragment,
  serialize,
  val,
  type Element,
} from "./xml.js";
import { ooxmlToPm, pmToOoxml } from "./ooxml-pm.js";
import type { LinkStyle } from "./ooxml-pm.js";
import type { PMDoc, PMNode } from "./pm-schema.js";
import { NumberingMap } from "./numbering.js";
import { StyleMap, type PaneStyle } from "./styles.js";
import { resolvedStyles, type ResolvedStyle } from "./style-format.js";
import { mergeStylesXml, type StyleEdit } from "./style-edit.js";
import {
  ensureDefaultHeaderFooter,
  loadHeaderFooters,
  serializeHeaderFooters,
  type HeaderFooterBundle,
} from "./headerFooter.js";
import {
  FOOTNOTE_CONTENT_TYPE,
  FOOTNOTE_PART,
  FOOTNOTE_REL_TYPE,
  applyFootnoteBodies,
  collectFootnoteRefs,
  loadFootnotes,
  serializeFootnotePart,
  syncFootnoteIds,
  type FootnoteBody,
  type FootnoteKind,
  type FootnoteStore,
} from "./footnotes.js";
import {
  COMMENT_CONTENT_TYPE,
  COMMENT_PART,
  COMMENT_REL_TYPE,
  applyCommentBodies,
  collectCommentRefs,
  loadComments,
  serializeCommentsPart,
  syncCommentIds,
  type CommentBody,
  type CommentStore,
} from "./comments.js";

export interface DocxParagraph {
  index: number;
  text: string;
  style?: string;
}

/**
 * Like `collectDescendants`, but it does not descend into `mc:Fallback`.
 * Word writes every floating drawing twice — once in the modern `mc:Choice`
 * branch and once as legacy VML in `mc:Fallback` — so a plain descendant scan
 * would report each text box paragraph a second time.
 */
function collectOutsideFallback(
  el: Element,
  name: string,
  acc: Element[] = [],
): Element[] {
  for (const c of children(el)) {
    if (c.name === name) acc.push(c);
    if (c.name === "mc:Fallback") continue;
    collectOutsideFallback(c, name, acc);
  }
  return acc;
}

export interface ReplaceResult {
  count: number;
  previews: Array<[string, string]>;
}

export interface Snapshot {
  xml: string;
  rels: string;
  [k: string]: unknown;
}

export interface PmLoadResult {
  doc: PMDoc;
  /** Original document.xml, retained so save can diff. */
  sourceXml: string;
  warnings: string[];
  /** Parsed headers/footers keyed by "header"/"footer". */
  headers?: HeaderFooterBundle;
  /** Page geometry read from <w:sectPr>; defaults to A4 if absent. */
  section: SectionGeometry;
  /**
   * 样式窗格要列出的段落样式，按 uiPriority 排好、隐藏样式已过滤。
   * 空数组 = 这份文档没有 styles.xml（或里面没有可显示的段落样式）。
   */
  styles: PaneStyle[];
  /**
   * 样式链解析出的可见格式，供渲染层生成文档作用域的样式表（单位与编辑器 attrs 一致：
   * pt / 倍数 / CSS px / #rrggbb）。`styleId: null` 那条是文档默认段落样式，给没有
   * `w:pStyle` 的段落。空数组 = 这份文档什么都排不出来，渲染层因此不注入任何规则。
   */
  styleFormats: ResolvedStyle[];
  /**
   * 脚注与尾注部件的条目，按 `w:id` 对应正文里的引用点。
   * 两个数组都空 = 这份文档没有脚注/尾注；正文里插入新引用时保存会自动补条目。
   */
  footnotes: FootnoteStore;
  /**
   * 批注部件的条目，按 `w:id` 对应正文里的锚点（区间两端 + 引用记号）。
   * 空条目 = 这份文档没有批注；正文里出现新锚点时保存会自动补条目。
   */
  comments: CommentStore;
}

/**
 * Page geometry derived from <w:sectPr>. All values are CSS pixels at
 * 96 DPI. OOXML stores dimensions in twips (1/1440 inch) and EMUs.
 *
 * Defaults match a standard A4 page with 1-inch margins.
 */
export interface SectionGeometry {
  /** Full page width in CSS px. */
  pageWidth: number;
  /** Full page height in CSS px. */
  pageHeight: number;
  /** Distance from page edge to the top of the body text. */
  marginTop: number;
  marginRight: number;
  marginBottom: number;
  marginLeft: number;
  /** Distance from page edge to the top of the header band. */
  header: number;
  /** Distance from page edge to the bottom of the footer band. */
  footer: number;
  /** The width of the printable text area: pageWidth - margins. */
  contentWidth: number;
  /** The height of the printable text area: pageHeight - margins. */
  contentHeight: number;
  /**
   * Height of one snapped line, in px, when the section carries
   * `<w:docGrid w:type="lines">`. Chinese templates set this and Word lays
   * every body line on it, so it — not the font — fixes how tall the anchor
   * paragraphs a page of floating boxes are measured from.
   */
  linePitch?: number;
  /**
   * Space below a paragraph the document gives no spacing of its own, in px,
   * read from the default paragraph style. Zero is Word's own default.
   */
  paraSpacingAfter?: number;
}

/** Document-wide decorations a template owns, applied at save time. */
export interface DocStyle {
  /** null clears the setting, undefined leaves it alone. */
  link?: LinkStyle | null;
}

const DOC_XML = "word/document.xml";
const RELS_XML = "word/_rels/document.xml.rels";
const NUMBERING_XML = "word/numbering.xml";
const STYLES_XML = "word/styles.xml";
const CONTENT_TYPES = "[Content_Types].xml";

/**
 * In-memory .docx editor. A .docx is a zip of XML parts; we only need
 * `word/document.xml` for read/edit and `word/_rels/document.xml.rels`
 * (plus the image itself) to insert an image.
 *
 * Typical flow:
 *   const d = await DocxFile.fromBytes(data);
 *   d.replaceText({find:"甲方", replace:"客户"});
 *   d.insertParagraph({text:"附加条款", position:"end"});
 *   const out = await d.toBytes();
 */
export class DocxFile {
  private zip: JSZip;
  private root: Element;
  private body: Element;
  private paragraphs: Element[];
  private numbering: NumberingMap;
  private styles: StyleMap;
  /**
   * `word/styles.xml` 的原文。合并式写回要在它上面动刀（splice-and-keep），
   * 只留 StyleMap 的话 `docDefaults`、`latentStyles` 这些没建模的块就找不回来了。
   */
  private stylesXml?: string;
  private headers: HeaderFooterBundle = { entries: {} };
  /** 脚注/尾注部件的条目；正文只放引用点，内容住在这里。 */
  private footnotes: FootnoteStore = { footnote: [], endnote: [], unmodeled: [] };
  /** 批注部件的条目；正文只放锚点，内容与作者信息住在这里。 */
  private comments: CommentStore = { entries: [], isUnmodeled: false };
  private pendingSection?: SectionGeometry;
  private pendingStyle?: DocStyle;

  private constructor(
    zip: JSZip,
    root: Element,
    numbering: NumberingMap,
    styles: StyleMap,
    headers?: HeaderFooterBundle,
  ) {
    this.zip = zip;
    this.root = root;
    this.numbering = numbering;
    this.styles = styles;
    if (headers) this.headers = headers;
    const body = firstChild(root, "w:body");
    if (!body) throw new Error("无效的 .docx：未找到 w:body。");
    this.body = body;
    this.paragraphs = children(body).filter((c) => c.name === "w:p");
  }

  static async fromBytes(bytes: Uint8Array | ArrayBuffer | Buffer): Promise<DocxFile> {
    const zip = await JSZip.loadAsync(bytes);
    const xml = await zip.file(DOC_XML)?.async("string");
    if (!xml) throw new Error("无效的 .docx：缺少 word/document.xml。");
    const root = parseDocument(xml);
    const numberingXml = await zip.file(NUMBERING_XML)?.async("string");
    const numbering = NumberingMap.fromXml(numberingXml);
    const stylesXml = await zip.file(STYLES_XML)?.async("string");
    const styles = StyleMap.fromXml(stylesXml);
    const relsXml = (await zip.file(RELS_XML)?.async("string")) ?? "";
    const headers = await loadHeaderFooters(zip, xml, relsXml, numbering, styles);
    const file = new DocxFile(zip, root, numbering, styles, headers);
    file.stylesXml = stylesXml;
    file.footnotes = await loadFootnotes(zip, relsXml, numbering, styles);
    file.comments = await loadComments(zip, relsXml, numbering, styles);
    return file;
  }

  // -------- Read --------

  listParagraphs(): DocxParagraph[] {
    const all = collectOutsideFallback(this.body, "w:p");
    return all.map((p, index) => {
      const pPr = firstChild(p, "w:pPr");
      const pStyle = pPr ? firstChild(pPr, "w:pStyle") : undefined;
      const style = val(pStyle);
      return { index, text: paragraphText(p), style };
    });
  }

  getSnapshot(): Snapshot {
    return { xml: serialize(this.root), rels: "" };
  }

  restoreSnapshot(snapshot: Snapshot): void {
    const root = parseDocument(snapshot.xml);
    const body = firstChild(root, "w:body");
    if (!body) throw new Error("快照损坏。");
    this.root = root;
    this.body = body;
    this.paragraphs = children(body).filter((c) => c.name === "w:p");
    this.zip.file(DOC_XML, serialize(root));
  }

  // -------- ProseMirror bridge (M3) --------

  /**
   * Convert the loaded .docx to a ProseMirror JSON document. Image
   * relationships are resolved to in-memory data URLs so the renderer can
   * display them without the zip.
   */
  async toProseMirror(): Promise<PmLoadResult> {
    const sourceXml = serialize(this.root);
    const pm = ooxmlToPm(sourceXml, this.numbering, this.styles);
    const warnings: string[] = [];
    await this.resolveImages(pm, warnings);
    // Resolve images in headers/footers too so the UI can show them.
    for (const entry of Object.values(this.headers.entries)) {
      await this.resolveImages(entry.pm, warnings);
    }
    const section = readSectionGeometry(this.body);
    section.paraSpacingAfter = this.styles.defaultSpacingAfterPx;
    return {
      doc: pm,
      sourceXml,
      warnings,
      headers: this.headers,
      section,
      styles: this.styles.paneStyles(),
      styleFormats: resolvedStyles(this.styles),
      footnotes: this.footnotes,
      comments: this.comments,
    };
  }

  /**
   * Update the header/footer ProseMirror content. Pass null/undefined to
   * leave an entry untouched. Pass an empty doc to clear it.
   */
  setHeaderFooter(type: "header" | "footer", pm: PMDoc): void {
    let entry = this.headers.entries[type];
    if (!entry) {
      const { headerRId, footerRId, headerPath, footerPath } =
        ensureDefaultHeaderFooter(this.headers, this.numbering);
      if (type === "header") {
        entry = this.headers.entries.header!;
        void headerRId;
        void headerPath;
      } else {
        entry = this.headers.entries.footer!;
        void footerRId;
        void footerPath;
      }
    }
    entry.pm = pm;
    entry.dirty = true;
  }

  getHeaderFooter(type: "header" | "footer"): PMDoc | null {
    return this.headers.entries[type]?.pm ?? null;
  }

  /**
   * 收下脚注/尾注窗格改过的正文，下一次保存写进对应部件。
   *
   * 只并正文：没被列出的条目（包括 `separator` 那两条版式件）保持原样透传，
   * 被标成 `unmodeled` 的部件一个字节都不动。正文删掉的引用点不在此处处理，
   * 由保存时的 `syncFootnoteIds` 按正文对齐。
   *
   * @param kind 脚注还是尾注。
   * @param bodies 编辑器交回的正文，`id` 与正文里的引用点一致。
   */
  setFootnoteBodies(kind: FootnoteKind, bodies: readonly FootnoteBody[]): void {
    this.footnotes = applyFootnoteBodies(this.footnotes, kind, bodies);
  }

  /**
   * 收下批注窗格改过的正文与作者信息，下一次保存写进 `word/comments.xml`。
   *
   * 只并列出来的条目：没被列出的条目保持原样透传，被标为不可重写的部件一个字节都不动。
   * 正文里删掉的锚点不在此处处理，由保存时的 `syncCommentIds` 按锚点对齐。
   *
   * @param bodies 编辑器交回的批注，`id` 与正文里的锚点一致。
   */
  setCommentBodies(bodies: readonly CommentBody[]): void {
    this.comments = applyCommentBodies(this.comments, bodies);
  }

  /**
   * Override the section geometry (page size, orientation, margins).
   * Called when the user edits page setup in the UI; the new values
   * are written into <w:sectPr> on the next save.
   */
  setSectionGeometry(section: SectionGeometry): void {
    this.pendingSection = section;
  }

  /**
   * Set the document-wide decorations (currently only hyperlinks). They live
   * in the run properties we write, not in styles.xml, because Word lets a run's
   * own rPr override the `Hyperlink` character style — and so does WPS.
   */
  setDocStyle(style: DocStyle): void {
    this.pendingStyle = style;
  }

  async replaceProseMirror(doc: PMDoc): Promise<string[]> {
    const warnings: string[] = [];
    await this.embedNewImages(doc, warnings);
    const xml = pmToOoxml(doc, {
      numbering: this.numbering,
      link: this.pendingStyle?.link ?? null,
      // 段落 attrs 里的样式 id 要对着本文档的样式表复核，见 fitStyleId。
      styles: this.styles,
    });
    const newRoot = parseDocument(xml);
    // Keep the source document's root attributes. A round-tripped floating
    // drawing can reference mc:/wps:/wpg:/v:/wp14:, and those prefixes are
    // declared on the root we are replacing — dropping them produces a file
    // Word reports as corrupt.
    for (const [name, value] of Object.entries(this.root.attributes ?? {})) {
      if (!(name in newRoot.attributes)) {
        newRoot.attributes[name] = value as string;
      }
    }
    // Preserve the original sectPr (page size, margins, headers/footers) if
    // the source document had one. Our pmToOoxml always appends a default
    // sectPr; replace it with the original when available.
    const oldBody = firstChild(this.root, "w:body");
    const oldSectPr = oldBody && firstChild(oldBody, "w:sectPr");
    const newBody = firstChild(newRoot, "w:body");
    if (newBody) {
      let sectPr: Element | null = oldSectPr ?? null;
      if (this.pendingSection) {
        sectPr = buildSectPr(sectPr, this.pendingSection);
      }
      const newSectPr = firstChild(newBody, "w:sectPr");
      if (sectPr) {
        if (newSectPr) {
          // Replace new default sectPr with the preserved/updated one.
          const arr = newBody.children as unknown as Element[];
          const idx = arr.indexOf(newSectPr);
          if (idx >= 0) arr[idx] = sectPr;
        } else {
          (newBody.children as unknown as Element[]).push(sectPr);
        }
      }
    }
    this.root = newRoot;
    this.body = firstChild(newRoot, "w:body")!;
    this.paragraphs = children(this.body).filter((c) => c.name === "w:p");
    this.zip.file(DOC_XML, serialize(newRoot));

    // If any list used a numbering definition, (re)write word/numbering.xml
    // and ensure it is referenced from [Content_Types] + document rels.
    await this.writeNumberingIfNeeded();

    // Re-serialize any edited headers/footers and ensure sectPr references
    // point to them. If the original doc had no header/footer at all, we
    // create the parts + relationships now.
    await this.writeHeaderFooters();

    // 脚注/尾注部件：正文里的引用点决定条目的增删，没被编辑过的条目按原始 XML 透传。
    await this.writeFootnoteParts(doc);

    // 批注部件：正文里的锚点决定条目的增删，没被编辑过的条目同样按原始 XML 透传。
    await this.writeCommentPart(doc);

    return warnings;
  }

  private async writeHeaderFooters(): Promise<void> {
    const bundle = this.headers;
    const hasAny = Object.keys(bundle.entries).length > 0;
    if (!hasAny) return;

    // Embed any data-URL images that the user added inside headers/footers.
    for (const entry of Object.values(bundle.entries)) {
      if (entry.dirty) {
        const warnings: string[] = [];
        await this.embedNewImages(entry.pm, warnings);
      }
    }

    const parts = serializeHeaderFooters(bundle, this.numbering);
    for (const [path, xml] of parts) {
      this.zip.file(path, xml);
    }

    // Make sure [Content_Types].xml declares header+footer parts.
    const ctFile = this.zip.file(CONTENT_TYPES);
    if (ctFile) {
      let ct = await ctFile.async("string");
      const add = (ext: string, ctype: string) => {
        if (!ct.includes(`Extension="${ext}"`)) {
          ct = ct.replace(
            "</Types>",
            `<Default Extension="${ext}" ContentType="${ctype}"/></Types>`,
          );
        }
      };
      add("xml", "application/xml");
      if (!ct.includes("header+xml")) {
        ct = ct.replace(
          "</Types>",
          `<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>`,
        );
      }
      if (!ct.includes("footer+xml")) {
        ct = ct.replace(
          "</Types>",
          `<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>`,
        );
      }
      this.zip.file(CONTENT_TYPES, ct);
    }

    // Ensure document.xml.rels contains the relationships.
    const relsFile = this.zip.file(RELS_XML);
    let rels = relsFile ? await relsFile.async("string") : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;
    const ensureRel = (rId: string, target: string, type: string) => {
      if (!rels.includes(`Id="${rId}"`)) {
        rels = rels.replace(
          "</Relationships>",
          `<Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}" Target="${target}"/></Relationships>`,
        );
      }
    };
    if (bundle.entries.header) ensureRel(bundle.entries.header.rId, "header1.xml", "header");
    if (bundle.entries.footer) ensureRel(bundle.entries.footer.rId, "footer1.xml", "footer");
    this.zip.file(RELS_XML, rels);

    // Inject headerReference/footerReference into sectPr if missing.
    // OOXML schema order requires headerReference/footerReference to
    // appear before <w:pgSz>. We also force header before footer.
    const body = firstChild(this.root, "w:body");
    const sectPr = body && firstChild(body, "w:sectPr");
    if (sectPr) {
      const mkRef = (tag: string, type: string, rId: string): Element =>
        ({
          type: "element",
          name: tag,
          attributes: { "r:id": rId, "w:type": type },
          children: [],
        }) as unknown as Element;

      const removeExisting = (tag: string, type: string) => {
        const arr = sectPr.children as unknown as Element[];
        for (let i = arr.length - 1; i >= 0; i--) {
          const c = arr[i]!;
          if (
            c.name === tag &&
            (c.attributes?.["w:type"] as string) === type
          ) {
            arr.splice(i, 1);
          }
        }
      };
      removeExisting("w:headerReference", "default");
      removeExisting("w:footerReference", "default");

      const newRefs: Element[] = [];
      if (bundle.entries.header)
        newRefs.push(mkRef("w:headerReference", "default", bundle.entries.header.rId));
      if (bundle.entries.footer)
        newRefs.push(mkRef("w:footerReference", "default", bundle.entries.footer.rId));

      const arr = sectPr.children as unknown as Element[];
      const pgSz = firstChild(sectPr, "w:pgSz");
      if (pgSz) {
        const idx = arr.indexOf(pgSz);
        arr.splice(idx, 0, ...newRefs);
      } else {
        arr.unshift(...newRefs);
      }
      this.zip.file(DOC_XML, serialize(this.root));
    }
  }

  /**
   * 按正文里的引用点重写脚注与尾注部件，并保证它们在包里登记过。
   *
   * 引用点是唯一真相：删掉引用就删掉条目（Word 也是这个语义），新引用补一条空正文。
   * 某个部件里有任何条目没能建模（id 不是十进制）时整个部件原样保留、永不重写，
   * 否则按条目重建会把读不进来的那几条丢掉。
   *
   * @param doc 即将写进正文的 PM doc，引用点从它身上收集（单位无关，只读 attrs.id）。
   * @returns 无返回值；某一边一条条目都没有时既不写文件也不登记关系。
   */
  private async writeFootnoteParts(doc: PMDoc): Promise<void> {
    this.footnotes = syncFootnoteIds(this.footnotes, collectFootnoteRefs(doc));
    for (const kind of ["footnote", "endnote"] as FootnoteKind[]) {
      if (this.footnotes.unmodeled.includes(kind)) continue;
      const xml = serializeFootnotePart(kind, this.footnotes[kind], this.numbering);
      if (!xml) continue;
      this.zip.file(FOOTNOTE_PART[kind], xml);
      await this.ensurePartReference(
        FOOTNOTE_PART[kind],
        FOOTNOTE_REL_TYPE[kind],
        FOOTNOTE_CONTENT_TYPE[kind],
      );
    }
  }

  /**
   * 写 `word/comments.xml`：正文里的锚点决定条目增删，没编辑过的按原始 XML 透传。
   *
   * 一条批注都没有时既不写文件也不登记关系——空的 `w:comments` 会让 Word 觉得文档
   * 有批注区，比没有这个部件更糟。
   *
   * @param doc 即将写进正文的 PM doc，锚点从它身上收集（只读 attrs.id）。
   * @returns 无返回值。
   */
  private async writeCommentPart(doc: PMDoc): Promise<void> {
    this.comments = syncCommentIds(this.comments, collectCommentRefs(doc));
    if (this.comments.isUnmodeled) return;
    const xml = serializeCommentsPart(this.comments.entries, this.numbering);
    if (!xml) return;
    const path = this.comments.partPath ?? COMMENT_PART;
    this.zip.file(path, xml);
    await this.ensurePartReference(path, COMMENT_REL_TYPE, COMMENT_CONTENT_TYPE);
  }

  // -------- Legacy edit API (CLI) --------

  replaceText(opts: {
    find: string;
    replace: string;
    regex?: boolean;
  }): ReplaceResult {
    const { find, replace, regex = false } = opts;
    if (!find) throw new Error("find 不能为空。");

    const previews: Array<[string, string]> = [];
    let count = 0;

    for (const runText of collectDescendants(this.body, "w:t")) {
      const original = textNodeText(runText);
      if (!original) continue;
      const next = regex
        ? original.replace(new RegExp(find, "g"), replace)
        : original.split(find).join(replace);
      if (next !== original) {
        count += occurrences(original, find, regex);
        previews.push([original, next]);
        setTextNodeText(runText, next);
      }
    }
    return { count, previews };
  }

  insertParagraph(opts: {
    text: string;
    position?: "before" | "after" | "start" | "end";
    style?: string;
    anchor?: number;
  }): number {
    const { text, position = "end", style } = opts;
    const p = makeParagraph(text, style);

    if (position === "end") {
      this.body.children.splice(this.insertIndexForEnd(), 0, p);
    } else if (position === "start") {
      (this.body.children as unknown as Element[]).unshift(p);
    } else if (opts.anchor == null) {
      throw new Error("before/after 插入需要指定 anchor 段落下标。");
    } else {
      const target = this.paragraphs[opts.anchor];
      if (!target) throw new Error(`未找到第 ${opts.anchor} 段。`);
      const arr = this.body.children as unknown as Element[];
      const idx = arr.indexOf(target);
      const insertAt = position === "before" ? idx : idx + 1;
      arr.splice(insertAt, 0, p);
    }

    this.paragraphs = children(this.body).filter((c) => c.name === "w:p");
    return this.paragraphs.indexOf(p);
  }

  insertTable(opts: {
    rows: string[][];
    headerRow?: boolean;
    position?: "start" | "end";
  }): void {
    const { rows, headerRow = true, position = "end" } = opts;
    if (rows.length === 0) throw new Error("表格至少需要一行。");
    const cols = rows[0]!.length;
    if (cols === 0) throw new Error("表格至少需要一列。");
    for (const r of rows) {
      if (r.length !== cols) throw new Error("表格每行列数必须一致。");
    }
    const tbl = makeTable(rows, headerRow);
    if (position === "start") {
      (this.body.children as unknown as Element[]).unshift(tbl);
    } else {
      (this.body.children as unknown as Element[]).splice(
        this.insertIndexForEnd(),
        0,
        tbl,
      );
    }
    this.paragraphs = children(this.body).filter((c) => c.name === "w:p");
  }

  async insertImage(opts: {
    bytes: Uint8Array;
    contentType: "image/png" | "image/jpeg";
    filename: string;
    widthPx?: number;
    heightPx?: number;
  }): Promise<void> {
    const { bytes, contentType, filename, widthPx, heightPx } = opts;
    const maxId = await this.nextMediaId();
    const target = `media/image${maxId}.${contentType === "image/png" ? "png" : "jpeg"}`;
    this.zip.file(`word/${target}`, bytes);

    const relsXml = await this.zip.file(RELS_XML)?.async("string");
    const rels = relsXml
      ? parseDocument(relsXml)
      : newRelsDocument();
    const existingIds = children(rels).map((c) => c.attributes?.["Id"] as string);
    const rId = nextRid(existingIds);
    (rels.children as unknown as Element[]).push({
      type: "element",
      name: "Relationship",
      attributes: {
        Id: rId,
        Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
        Target: target,
      },
      children: [],
    } as unknown as Element);
    this.zip.file(RELS_XML, serialize(rels));

    await this.ensureImageContentTypes(contentType);

    const cx = (widthPx ?? 200) * 9525;
    const cy = (heightPx ?? 200) * 9525;
    const para = makeImageParagraph({ rId, cx, cy, filename });
    (this.body.children as unknown as Element[]).splice(
      this.insertIndexForEnd(),
      0,
      para,
    );
    this.paragraphs = children(this.body).filter((c) => c.name === "w:p");
  }

  async toBytes(): Promise<Uint8Array> {
    this.zip.file(DOC_XML, serialize(this.root));
    return this.zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
  }

  // -------- Internals --------

  private insertIndexForEnd(): number {
    const arr = this.body.children as unknown as Element[];
    const last = arr[arr.length - 1];
    if (last && last.name === "w:sectPr") return arr.length - 1;
    return arr.length;
  }

  private async nextMediaId(): Promise<number> {
    let max = 0;
    this.zip.forEach((p) => {
      const m = p.match(/^word\/media\/image(\d+)\./);
      if (m) max = Math.max(max, Number(m[1]));
    });
    return max + 1;
  }

  private async ensureImageContentTypes(contentType: string): Promise<void> {
    const ext = contentType === "image/png" ? "png" : "jpeg";
    const file = this.zip.file(CONTENT_TYPES);
    if (!file) return;
    const xml = await file.async("string");
    const ct = parseDocument(xml);
    const has = children(ct).some(
      (c) =>
        c.name === "Default" &&
        (c.attributes?.["Extension"] === ext || c.attributes?.["Extension"] === "jpg"),
    );
    if (!has) {
      (ct.children as unknown as Element[]).unshift({
        type: "element",
        name: "Default",
        attributes: { Extension: ext, ContentType: contentType },
        children: [],
      } as unknown as Element);
      this.zip.file(CONTENT_TYPES, serialize(ct));
    }
  }

  /**
   * Walk a PMDoc and replace `relationship:rIdX` image srcs with data URLs
   * read from the zip. Missing relationships produce a warning and a 1x1
   * transparent PNG placeholder so the editor does not break.
   */
  private async resolveImages(doc: PMDoc, warnings: string[]): Promise<void> {
    const relsXml = await this.zip.file(RELS_XML)?.async("string");
    const rels = relsXml ? parseDocument(relsXml) : undefined;
    const ridToTarget = new Map<string, string>();
    if (rels) {
      for (const r of children(rels)) {
        const id = r.attributes?.["Id"] as string | undefined;
        const target = r.attributes?.["Target"] as string | undefined;
        if (id && target) ridToTarget.set(id, target);
      }
    }
    const walk = async (nodes: PMNode[] | undefined): Promise<void> => {
      if (!nodes) return;
      for (const n of nodes) {
        if (n.type === "image" && n.attrs.src.startsWith("relationship:")) {
          const rId = n.attrs.src.slice("relationship:".length);
          const target = ridToTarget.get(rId);
          if (!target) {
            warnings.push(`图片关系 ${rId} 未找到，已用占位图替代。`);
            n.attrs.src = PLACEHOLDER_PNG;
            continue;
          }
          const path = target.startsWith("media/") ? `word/${target}` : `word/${target.replace(/^\.?\//, "")}`;
          const file = this.zip.file(path);
          if (!file) {
            warnings.push(`图片文件 ${path} 未在压缩包里找到，已用占位图替代。`);
            n.attrs.src = PLACEHOLDER_PNG;
            continue;
          }
          const bytes = await file.async("base64");
          const ext = (target.split(".").pop() ?? "png").toLowerCase();
          const mime =
            ext === "jpg" || ext === "jpeg"
              ? "image/jpeg"
              : ext === "gif"
                ? "image/gif"
                : ext === "webp"
                  ? "image/webp"
                  : "image/png";
          n.attrs.src = `data:${mime};base64,${bytes}`;
        }
        if ("content" in n) await walk(n.content as PMNode[] | undefined);
      }
    };
    await walk(doc.content);
  }

  /**
   * For every image whose src is a data URL, write the binary into the zip
   * and rewrite the attribute to `relationship:rIdX`. Existing relationship
   * refs are kept as-is.
   */
  private async embedNewImages(doc: PMDoc, warnings: string[]): Promise<void> {
    const relsXml = await this.zip.file(RELS_XML)?.async("string");
    const rels = relsXml
      ? parseDocument(relsXml)
      : newRelsDocument();
    const existingIds = children(rels).map((c) => c.attributes?.["Id"] as string);

    let nextMediaId = await this.nextMediaId();
    let nextRidNum = existingIds
      .map((id) => {
        const m = id?.match(/^rId(\d+)$/);
        return m ? Number(m[1]) : 0;
      })
      .reduce((a, b) => Math.max(a, b), 0);

    const walk = async (nodes: PMNode[] | undefined): Promise<void> => {
      if (!nodes) return;
      for (const n of nodes) {
        if (n.type === "image" && n.attrs.src.startsWith("data:")) {
          const match = /^data:([^;]+);base64,(.*)$/s.exec(n.attrs.src);
          if (!match) {
            warnings.push("无法识别的图片 data URL，已忽略。");
            continue;
          }
          const mime = match[1]!;
          const base64 = match[2]!;
          const ext =
            mime === "image/jpeg"
              ? "jpeg"
              : mime === "image/gif"
                ? "gif"
                : mime === "image/webp"
                  ? "webp"
                  : "png";
          const bytes = Buffer.from(base64, "base64");
          const target = `media/image${nextMediaId}.${ext}`;
          nextMediaId += 1;
          this.zip.file(`word/${target}`, bytes);
          const rId = `rId${nextRidNum + 1}`;
          nextRidNum += 1;
          (rels.children as unknown as Element[]).push({
            type: "element",
            name: "Relationship",
            attributes: {
              Id: rId,
              Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
              Target: target,
            },
            children: [],
          } as unknown as Element);
          n.attrs.src = `relationship:${rId}`;
          await this.ensureImageContentTypes(mime);
        }
        if ("content" in n) await walk(n.content as PMNode[] | undefined);
      }
    };
    await walk(doc.content);
    this.zip.file(RELS_XML, serialize(rels));
  }

  /**
   * If the document contains any list and the numbering part doesn't yet
   * have a usable definition, write word/numbering.xml with our defaults
   * and ensure [Content_Types].xml + document.xml.rels reference it.
   */
  private async writeNumberingIfNeeded(): Promise<void> {
    if (!this.numbering.hasAny()) return;
    const existingXml = await this.zip.file(NUMBERING_XML)?.async("string");
    this.zip.file(NUMBERING_XML, this.numbering.toXml(existingXml));
    await this.ensurePartReference(
      NUMBERING_XML,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
    );
  }

  /**
   * 把"新建样式 / 修改样式"的编辑并进 `word/styles.xml`（splice-and-keep）。
   *
   * 与写列表那套同一条路：合并只动被点名的 `w:style`，`docDefaults` 与其余样式原样带走，
   * 部件不存在时补登记 `[Content_Types].xml` 与 rels。合并完立刻刷新内存里的样式表，
   * 所以同一次保存里 `replaceProseMirror` 判断 styleId 是否已声明时看到的就是新状态。
   *
   * @param edits 本会话累计的编辑（渲染层带全量，顺序 = 用户操作顺序）；空数组什么都不做。
   * @returns 样式窗格清单与解析好的样式格式，供界面立刻刷新预览；edits 全是无效名字时
   *   zip 不动，返回的仍是原样式表——不会静默造出一份空 styles.xml。
   */
  async applyStyleEdits(
    edits: readonly StyleEdit[],
  ): Promise<{ styles: PaneStyle[]; styleFormats: ResolvedStyle[] }> {
    const projection = (): { styles: PaneStyle[]; styleFormats: ResolvedStyle[] } => ({
      styles: this.styles.paneStyles(),
      styleFormats: resolvedStyles(this.styles),
    });
    if (!edits.length) return projection();
    const merged = mergeStylesXml(this.stylesXml, edits);
    if (merged === undefined || merged === this.stylesXml) return projection();
    this.stylesXml = merged;
    this.zip.file(STYLES_XML, merged);
    await this.ensurePartReference(
      STYLES_XML,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
    );
    this.styles = StyleMap.fromXml(merged);
    return projection();
  }

  /**
   * 确保一个部件在 `[Content_Types].xml` 与 `document.xml.rels` 里登记过。
   *
   * 新建的文档可能压根没有这个部件（例如它原本没有样式表），只写 zip 里的 XML
   * 是不够的：Word 认的是包里的登记关系，缺了就直接判"文档损坏"。
   *
   * @param partName zip 内部路径，如 "word/styles.xml"。
   * @param relationshipTypeUrl rels 里的 @Type，如 "…/relationships/styles"。
   * @param contentType Override 的 @ContentType（不含 `wordprocessingml.` 前缀差异由调用方给全）。
   * @returns 无返回值；两处都按需要写入。包里没有 `[Content_Types].xml` 时跳过登记——
   *   那不是我们要伪造的包；没有 rels 部件时新建一份，Word 只认关系里的引用。
   */
  private async ensurePartReference(
    partName: string,
    relationshipTypeUrl: string,
    contentType: string,
  ): Promise<void> {
    const ctFile = this.zip.file(CONTENT_TYPES);
    if (ctFile) {
      const ct = await ctFile.async("string");
      if (!ct.includes(`/${partName}`)) {
        const override = `<Override PartName="/${partName}" ContentType="${contentType}"/>`;
        this.zip.file(CONTENT_TYPES, ct.replace(/<Types\b[^>]*>/, (m) => `${m}${override}`));
      }
    }

    const relsXml = await this.zip.file(RELS_XML)?.async("string");
    if (relsXml?.includes(relationshipTypeUrl)) return;
    const rels = relsXml ? parseDocument(relsXml) : newRelsDocument();
    const existingIds = children(rels).map((c) => c.attributes?.["Id"] as string);
    (rels.children as unknown as Element[]).push({
      type: "element",
      name: "Relationship",
      attributes: {
        Id: nextRid(existingIds),
        Type: relationshipTypeUrl,
        Target: partName.split("/").pop() ?? partName,
      },
      children: [],
    } as unknown as Element);
    this.zip.file(RELS_XML, serialize(rels));
  }
}

// ---------- Section / page geometry ----------

/** 96 DPI: 1 inch = 96 px; OOXML twips are 1/1440 inch. */
const TWIP_TO_PX = 96 / 1440;
/** A4 in twips: 11906 × 16838. */
const DEFAULT_A4 = { w: 11906, h: 16838 };
/** Default 1-inch (1440 twip) margins; header at 0.5in, footer at 0.5in. */
const DEFAULT_MARGINS = {
  top: 1440,
  right: 1440,
  bottom: 1440,
  left: 1440,
  header: 720,
  footer: 720,
};

function readSectionGeometry(body: Element): SectionGeometry {
  const sectPr = firstChild(body, "w:sectPr");
  const pgSz = sectPr && firstChild(sectPr, "w:pgSz");
  const pgMar = sectPr && firstChild(sectPr, "w:pgMar");
  const num = (v: string | undefined, fallback: number): number => {
    if (v === undefined) return fallback;
    const n = Number(v);
    return Number.isFinite(n) ? n : fallback;
  };
  const w = num(pgSz?.attributes?.["w:w"] as string | undefined, DEFAULT_A4.w);
  const h = num(pgSz?.attributes?.["w:h"] as string | undefined, DEFAULT_A4.h);
  const marginTop = num(
    pgMar?.attributes?.["w:top"] as string | undefined,
    DEFAULT_MARGINS.top,
  );
  const marginRight = num(
    pgMar?.attributes?.["w:right"] as string | undefined,
    DEFAULT_MARGINS.right,
  );
  const marginBottom = num(
    pgMar?.attributes?.["w:bottom"] as string | undefined,
    DEFAULT_MARGINS.bottom,
  );
  const marginLeft = num(
    pgMar?.attributes?.["w:left"] as string | undefined,
    DEFAULT_MARGINS.left,
  );
  const header = num(
    pgMar?.attributes?.["w:header"] as string | undefined,
    DEFAULT_MARGINS.header,
  );
  const footer = num(
    pgMar?.attributes?.["w:footer"] as string | undefined,
    DEFAULT_MARGINS.footer,
  );
  // Orientation swaps width/height for landscape.
  const orient = pgSz?.attributes?.["w:orient"] as string | undefined;
  const pageWidth = (orient === "landscape" ? h : w) * TWIP_TO_PX;
  const pageHeight = (orient === "landscape" ? w : h) * TWIP_TO_PX;
  const ml = marginLeft * TWIP_TO_PX;
  const mr = marginRight * TWIP_TO_PX;
  const mt = marginTop * TWIP_TO_PX;
  const mb = marginBottom * TWIP_TO_PX;
  // `w:linePitch` is the grid step; only a grid of type lines/both snaps text.
  const grid = sectPr && firstChild(sectPr, "w:docGrid");
  const gridType = grid?.attributes?.["w:type"] as string | undefined;
  const pitch = num(grid?.attributes?.["w:linePitch"] as string | undefined, 0);
  const linePitch =
    gridType === "lines" || gridType === "both" ? pitch * TWIP_TO_PX : 0;
  return {
    pageWidth,
    pageHeight,
    marginTop: mt,
    marginRight: mr,
    marginBottom: mb,
    marginLeft: ml,
    header: header * TWIP_TO_PX,
    footer: footer * TWIP_TO_PX,
    contentWidth: pageWidth - ml - mr,
    contentHeight: pageHeight - mt - mb,
    linePitch: linePitch || undefined,
  };
}

// ---------- XML element factories ----------

const TWIP_PER_PX = 1440 / 96; // 15
const PX_TO_TWIP = (px: number) => Math.round(px * TWIP_PER_PX);

/**
 * Build or update a <w:sectPr> with the given geometry. We preserve
 * any other children the source had (header/footer references,
 * docGrid, cols, titlePg) and only replace pgSz/pgMar.
 */
function buildSectPr(
  existing: Element | null,
  g: SectionGeometry,
): Element {
  const orient: "portrait" | "landscape" =
    g.pageWidth > g.pageHeight ? "landscape" : "portrait";
  const wTw = PX_TO_TWIP(
    orient === "landscape" ? g.pageHeight : g.pageWidth,
  );
  const hTw = PX_TO_TWIP(
    orient === "landscape" ? g.pageWidth : g.pageHeight,
  );
  const pgSzAttrs: Record<string, string> = {
    "w:w": String(wTw),
    "w:h": String(hTw),
  };
  if (orient === "landscape") pgSzAttrs["w:orient"] = "landscape";
  const pgMarAttrs: Record<string, string> = {
    "w:top": String(PX_TO_TWIP(g.marginTop)),
    "w:right": String(PX_TO_TWIP(g.marginRight)),
    "w:bottom": String(PX_TO_TWIP(g.marginBottom)),
    "w:left": String(PX_TO_TWIP(g.marginLeft)),
    "w:header": String(PX_TO_TWIP(g.header)),
    "w:footer": String(PX_TO_TWIP(g.footer)),
    "w:gutter": "0",
  };

  const newPgSz: Element = {
    type: "element",
    name: "w:pgSz",
    attributes: pgSzAttrs,
    children: [],
  } as unknown as Element;
  const newPgMar: Element = {
    type: "element",
    name: "w:pgMar",
    attributes: pgMarAttrs,
    children: [],
  } as unknown as Element;

  if (!existing) {
    return {
      type: "element",
      name: "w:sectPr",
      attributes: {},
      children: [newPgSz, newPgMar],
    } as unknown as Element;
  }

  const childrenArr = [...(existing.children as unknown as Element[])];
  const indexOf = (name: string): number =>
    childrenArr.findIndex((c) => c.name === name);
  const pgSzAt = indexOf("w:pgSz");
  if (pgSzAt >= 0) childrenArr[pgSzAt] = newPgSz;
  else childrenArr.unshift(newPgSz);
  // pgMar 必须紧跟 pgSz（ECMA-376 的 CT_SectPr 顺序），往后追加会排到
  // w:cols / w:pgBorders / w:docGrid 之后，Word 判定顺序不合法。
  const pgMarAt = indexOf("w:pgMar");
  if (pgMarAt >= 0) childrenArr[pgMarAt] = newPgMar;
  else childrenArr.splice(indexOf("w:pgSz") + 1, 0, newPgMar);
  // 必须显式写 `type`/`name`：`parseDocument` 出来的节点只有原型上带 `type`，
  // 展开运算符拿不到，`serialize` 认它不是元素就返回空串——整块 sectPr 静默消失。
  return {
    type: "element",
    name: "w:sectPr",
    attributes: existing.attributes ?? {},
    children: childrenArr,
  } as unknown as Element;
}

const PLACEHOLDER_PNG =
  "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=";

function paragraphText(p: Element): string {
  // Recursion stops at a nested `<w:p>`, which is a floating text box's own
  // paragraph (it gets listed separately), and at `mc:Fallback`, Word's
  // duplicate VML branch of the same drawing.
  let out = "";
  const walk = (el: Element): void => {
    for (const c of children(el)) {
      if (c.name === "w:t") out += textNodeText(c);
      else if (c.name !== "w:p" && c.name !== "mc:Fallback") walk(c);
    }
  };
  walk(p);
  return out;
}

function textNodeText(t: Element): string {
  let out = "";
  for (const c of t.children ?? []) {
    if (c.type === "text") out += (c as { text: string }).text;
  }
  return out;
}

function setTextNodeText(t: Element, value: string): void {
  t.children = [
    {
      type: "text",
      text: value,
    } as never,
  ];
  if (value.startsWith(" ") || value.endsWith(" ")) {
    t.attributes = { ...(t.attributes ?? {}), "xml:space": "preserve" };
  }
}

function occurrences(haystack: string, needle: string, regex: boolean): number {
  if (regex) {
    const m = haystack.match(new RegExp(needle, "g"));
    return m ? m.length : 0;
  }
  return haystack.split(needle).length - 1;
}

function makeParagraph(text: string, style?: string): Element {
  const pPr = style
    ? `<w:pPr><w:pStyle w:val="${escapeAttr(style)}"/></w:pPr>`
    : "";
  const safe = text
    .split("\n")
    .map((line) => `<w:r><w:t xml:space="preserve">${escapeText(line)}</w:t></w:r>`)
    .join('<w:r><w:br/></w:r>');
  const [p] = parseFragment(`<w:p>${pPr}${safe}</w:p>`);
  if (!p) throw new Error("构造段落失败。");
  return p;
}

function makeTable(rows: string[][], headerRow: boolean): Element {
  const trs = rows
    .map((r, i) => {
      const bold = headerRow && i === 0;
      const tcs = r
        .map(
          (cell) =>
            `<w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr><w:p>${runXml(
              cell,
              bold,
            )}</w:p></w:tc>`,
        )
        .join("");
      return `<w:tr>${tcs}</w:tr>`;
    })
    .join("");
  const tblPr =
    `<w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/>` +
    `<w:tblBorders>` +
    `<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>` +
    `<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>` +
    `<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>` +
    `<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>` +
    `<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>` +
    `<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>` +
    `</w:tblBorders></w:tblPr>`;
  const [tbl] = parseFragment(`<w:tbl>${tblPr}${trs}</w:tbl>`);
  if (!tbl) throw new Error("构造表格失败。");
  return tbl;
}

function runXml(text: string, bold: boolean): string {
  const rPr = bold ? "<w:rPr><w:b/></w:rPr>" : "";
  return `<w:r>${rPr}<w:t xml:space="preserve">${escapeText(text)}</w:t></w:r>`;
}

function makeImageParagraph(opts: { rId: string; cx: number; cy: number; filename: string }): Element {
  const xml =
    `<w:p><w:r><w:drawing>` +
    `<wp:inline distT="0" distB="0" distL="0" distR="0">` +
    `<wp:extent cx="${opts.cx}" cy="${opts.cy}"/>` +
    `<wp:effectExtent l="0" t="0" r="0" b="0"/>` +
    `<wp:docPr id="1" name="${escapeAttr(opts.filename)}"/>` +
    `<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>` +
    `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
    `<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
    `<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
    `<pic:nvPicPr><pic:cNvPr id="1" name="${escapeAttr(opts.filename)}"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="${opts.rId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${opts.cx}" cy="${opts.cy}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
    `</pic:pic>` +
    `</a:graphicData></a:graphic>` +
    `</wp:inline></w:drawing></w:r></w:p>`;
  const [p] = parseFragment(xml);
  if (!p) throw new Error("构造图片段落失败。");
  return p;
}

function nextRid(existing: string[]): string {
  let max = 0;
  for (const id of existing) {
    const m = id?.match(/^rId(\d+)$/);
    if (m) max = Math.max(max, Number(m[1]));
  }
  return `rId${max + 1}`;
}

function newRelsDocument(): Element {
  const [root] = parseFragment(
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`,
  );
  if (!root) throw new Error("构造 rels 失败。");
  return root;
}
