import JSZip from "jszip";
import {
  children,
  collectDescendants,
  escapeAttr,
  escapeText,
  firstChild,
  isElement,
  parseDocument,
  parseFragment,
  serialize,
  val,
  type Element,
} from "./xml.js";
import {
  colorVal,
  cssColorToOoxmlHex,
  ooxmlToPm,
  pmToOoxml,
} from "./ooxml-pm.js";
import type { LinkStyle } from "./ooxml-pm.js";
import type { PMDoc, PMNode } from "./pm-schema.js";
import { NumberingMap } from "./numbering.js";
import { StyleMap, type PaneStyle } from "./styles.js";
import { resolvedStyles, type ResolvedStyle } from "./style-format.js";
import { mergeStylesXml, type StyleEdit } from "./style-edit.js";
import {
  ensureHeaderFooterPart,
  loadHeaderFooters,
  serializeHeaderFooters,
  type HeaderFooterBundle,
} from "./headerFooter.js";
import {
  applyWatermark,
  readWatermark,
  sameWatermark,
  type SectionWatermark,
} from "./watermark.js";
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
  /**
   * 分栏（`w:cols`）。缺省就是单栏：Word 自己也常整个省略这个元素，所以不拿
   * `count: 1` 去冒充"有分栏设置"。`gap` 与 `widths` 是 CSS px @96DPI，
   * twips 换算只发生在本包的进出口（AGENTS §2.1）。
   */
  columns?: SectionColumns;
  /**
   * 页面边框（`w:pgBorders`）。`undefined` 是"这次保存不涉及边框"，原元素一动不动；
   * `null` 是"用户关掉了边框"，保存时把这个元素删掉。
   */
  pageBorders?: SectionPageBorders | null;
  /** 行号（`w:lnNumType`），`undefined`/`null` 的语义同 `pageBorders`。 */
  lineNumbers?: SectionLineNumbers | null;
  /**
   * 水印。它不是节属性——Word 把它写成**页眉部件里的一个 VML 艺术字形状**，
   * 所以读的时候从页眉原始 XML 里取，写的时候在页眉序列化之后做字符串级替换。
   * `undefined`/`null` 的语义同 `pageBorders`。
   */
  watermark?: SectionWatermark | null;
}

/**
 * `w:cols` 的建模：栏数、栏间距、是否等宽，以及每栏宽度。
 *
 * `widths` 对应 `w:col` 子元素：等宽时 Word 常常一条都不写，靠 `w:num` + `w:space`
 * 自己算，所以我们用空数组表示"没有显式宽度"，写回时也不写 `w:col`。
 * 读到的 `w:cols` 里若有我们没建模的属性（`w:sep` 栏间分隔线），
 * `buildSectPr` 会因为它跟读进来的值完全一致而原样留着整段，不重建。
 */
export interface SectionColumns {
  /** 栏数，`w:num`。 */
  count: number;
  /** 栏与栏之间的空隙，`w:space`，CSS px。 */
  gap: number;
  /** `w:equalWidth`，Word 缺省按等宽处理，所以缺省值是 true。 */
  equalWidth: boolean;
  /** 每栏宽度，`w:col w:w`，CSS px；等宽且 Word 没写子元素时为空数组。 */
  widths: number[];
}

/**
 * 页面边框的一条边，即 `w:pgBorders` 下的 `w:top`/`w:left`/`w:bottom`/`w:right`。
 *
 * 单位跟着 OOXML 的原生量纲留在模型里：`w:sz` 是 1/8 磅、`w:space` 是磅
 * （ECMA-376 §17.3.4 的 CT_Border），所以这里一律存 pt，换算只发生在写 XML 那一刻
 * （AGENTS §2.1）。
 */
export interface PageBorderSide {
  /** 线型 `w:val`，ST_Border 的字面值（`single`/`double`/`dashed`…）。 */
  style: string;
  /** 线宽，pt（写出去是 `w:sz`，1/8 磅）。 */
  sizePt: number;
  /** 边框到页边/正文的间距，pt（`w:space` 本身就是磅，不换算 twips）。 */
  spacePt: number;
  /** 颜色，CSS 写法 `#rrggbb`；null 是 Word 的 `auto`。 */
  color: string | null;
}

/**
 * `w:pgBorders` 的建模：四条边各自可选，null 表示那条边不画。
 *
 * 四条边都是 null 时不写这个元素——空 `w:pgBorders` 和没有它一样是不画边框。
 * 元素上没建模的属性（`w:display`、`w:zOrder`）与边上没建模的 `w:shadow`/`w:frame`/
 * 主题色，靠 `placeOrClear` 的"值等价就整段不重建"原样留着。
 */
export interface SectionPageBorders {
  /** `w:offsetFrom`：`page` 量到页边、`text` 量到正文（规范正文只出现这两个）。 */
  offsetFrom: string;
  top: PageBorderSide | null;
  left: PageBorderSide | null;
  bottom: PageBorderSide | null;
  right: PageBorderSide | null;
}

/**
 * `w:lnNumType` 的建模（行号）。
 *
 * `distance` 是 CSS px（`w:distance` 是 twips，换算在本包进出口）；`start` 只有不等于
 * Word 缺省的 1 才写出去。`w:restart`（每页/每节/连续）还没建模，值没变时整段原样保留。
 */
export interface SectionLineNumbers {
  /** 每几行标一个号，`w:countBy`；0 或负数表示不标行号。 */
  countBy: number;
  /** 行号与正文的距离，CSS px。 */
  distance: number;
  /** 起始号，`w:start`。 */
  start: number;
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
    // 水印住在页眉部件里，形状（VML 艺术字）PM 表示不了，只能读原始 XML 那一份。
    const watermark = readWatermark(this.headers.entries.header?.rawXml);
    if (watermark) section.watermark = watermark;
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
    const entry = ensureHeaderFooterPart(this.headers, type);
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
    const watermark = this.pendingSection?.watermark;
    // 只设了水印、文档本来没有页眉：先补一个页眉部件，形状才有地方放。
    if (watermark && !bundle.entries.header) ensureHeaderFooterPart(bundle, "header");
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
    const header = bundle.entries.header;
    if (watermark !== undefined && header) {
      const xml = parts.get(header.path);
      // 值等价就不重建：页面设置每次保存都会把读到的水印带回来，
      // 重写会把 Word 那份形状换成我们那份，字节级别来回变。
      if (xml && !sameWatermark(readWatermark(xml), watermark)) {
        parts.set(header.path, applyWatermark(xml, watermark));
      }
    }
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
    // 必须解析后追加：rels 一条关系都没有时 serialize 出来的是自闭合的
    // `<Relationships …/>`，字符串替换 `</Relationships>` 匹配不上，页眉关系就悄悄丢了，
    // 而 sectPr 里的 `r:id` 指着不存在的目标 = Word 判文档损坏。
    const relsFile = this.zip.file(RELS_XML);
    const relsRoot = relsFile
      ? parseDocument(await relsFile.async("string"))
      : newRelsDocument();
    const relBase = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
    let addedRel = false;
    if (bundle.entries.header) {
      addedRel =
        addRelationship(relsRoot, bundle.entries.header.rId, `${relBase}header`, "header1.xml") ||
        addedRel;
    }
    if (bundle.entries.footer) {
      addedRel =
        addRelationship(relsRoot, bundle.entries.footer.rId, `${relBase}footer`, "footer1.xml") ||
        addedRel;
    }
    if (addedRel) this.zip.file(RELS_XML, serialize(relsRoot));

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
   * 一条批注都没有时把部件连同关系与 Override 一起删掉：文档里既没有批注区，
   * 也不会留下一份带孤立条目的旧文件。
   *
   * @param doc 即将写进正文的 PM doc，锚点从它身上收集（只读 attrs.id）。
   * @returns 无返回值。
   */
  private async writeCommentPart(doc: PMDoc): Promise<void> {
    this.comments = syncCommentIds(this.comments, collectCommentRefs(doc));
    if (this.comments.isUnmodeled) return;
    const xml = serializeCommentsPart(this.comments.entries, this.numbering);
    // 批注全删光时不能只是"不写部件"：包里那份旧的 word/comments.xml 还带着孤立条目，
    // Word 会把已删的批注重新显示出来。空的 <w:comments> 又会让 Word 以为文档有批注区，
    // 所以按 Word 自己的做法把部件连同登记一起摘掉。
    if (!xml) {
      await this.dropCommentPart(this.comments.partPath ?? COMMENT_PART);
      return;
    }
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
    addRelationship(
      rels,
      nextRid(existingIds),
      relationshipTypeUrl,
      partName.split("/").pop() ?? partName,
    );
    this.zip.file(RELS_XML, serialize(rels));
  }

  /**
   * 摘掉批注部件：文件本身、`[Content_Types].xml` 里那条 Override、文档关系各清一份。
   *
   * 与 {@link ensurePartReference} 相反，但只在包里真的有这个部件时才动手——
   * 一份从来没有批注的文档不该因为保存一次就被改写 rels/Content_Types 的字节。
   *
   * @param partName 部件路径，如 `word/comments.xml`。
   * @returns 无返回值。
   */
  private async dropCommentPart(partName: string): Promise<void> {
    if (!this.zip.file(partName)) return;
    this.zip.remove(partName);

    const ctFile = this.zip.file(CONTENT_TYPES);
    if (ctFile) {
      const ct = parseDocument(await ctFile.async("string"));
      ct.children = (ct.children ?? []).filter(
        (node) =>
          !(
            isElement(node) &&
            node.name === "Override" &&
            node.attributes?.["PartName"] === `/${partName}`
          ),
      );
      this.zip.file(CONTENT_TYPES, serialize(ct));
    }

    const relsFile = this.zip.file(RELS_XML);
    if (relsFile) {
      const rels = parseDocument(await relsFile.async("string"));
      rels.children = (rels.children ?? []).filter(
        (node) => !(isElement(node) && node.attributes?.["Type"] === COMMENT_REL_TYPE),
      );
      this.zip.file(RELS_XML, serialize(rels));
    }
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

/** Word 行号对话框里"距正文"的缺省值 0.25 英寸。本机没装 Word，这是推断，不是实测。 */
const DEFAULT_LN_DISTANCE_TWIP = 360;

/**
 * 把 OOXML 属性值收成数字，读不懂（缺省、非数字）就用兜底值。
 *
 * @param v 属性原文，可能是 undefined。
 * @param fallback 兜底值，语义是"这个属性缺省时 Word 怎么理解"。
 * @returns 有限数字或兜底值。
 */
function toNum(v: string | undefined, fallback: number): number {
  if (v === undefined) return fallback;
  const n = Number(v);
  return Number.isFinite(n) ? n : fallback;
}

function readSectionGeometry(body: Element): SectionGeometry {
  const sectPr = firstChild(body, "w:sectPr");
  const pgSz = sectPr && firstChild(sectPr, "w:pgSz");
  const pgMar = sectPr && firstChild(sectPr, "w:pgMar");
  const w = toNum(pgSz?.attributes?.["w:w"] as string | undefined, DEFAULT_A4.w);
  const h = toNum(pgSz?.attributes?.["w:h"] as string | undefined, DEFAULT_A4.h);
  const marginTop = toNum(
    pgMar?.attributes?.["w:top"] as string | undefined,
    DEFAULT_MARGINS.top,
  );
  const marginRight = toNum(
    pgMar?.attributes?.["w:right"] as string | undefined,
    DEFAULT_MARGINS.right,
  );
  const marginBottom = toNum(
    pgMar?.attributes?.["w:bottom"] as string | undefined,
    DEFAULT_MARGINS.bottom,
  );
  const marginLeft = toNum(
    pgMar?.attributes?.["w:left"] as string | undefined,
    DEFAULT_MARGINS.left,
  );
  const header = toNum(
    pgMar?.attributes?.["w:header"] as string | undefined,
    DEFAULT_MARGINS.header,
  );
  const footer = toNum(
    pgMar?.attributes?.["w:footer"] as string | undefined,
    DEFAULT_MARGINS.footer,
  );
  // `w:orient` 只是"这页按横向排"的标记，页面尺寸仍以 `w:w`/`w:h` 为准：Word 存横向
  // Letter 时写的就是 w=15840 h=12240（长边在 w）。所以读的时候不做任何交换，
  // 谁大谁就是宽——按 orient 反着换会把 Word 的横向文档读成竖版形状。
  const pageWidth = w * TWIP_TO_PX;
  const pageHeight = h * TWIP_TO_PX;
  const ml = marginLeft * TWIP_TO_PX;
  const mr = marginRight * TWIP_TO_PX;
  const mt = marginTop * TWIP_TO_PX;
  const mb = marginBottom * TWIP_TO_PX;
  // `w:linePitch` is the grid step; only a grid of type lines/both snaps text.
  const grid = sectPr && firstChild(sectPr, "w:docGrid");
  const gridType = grid?.attributes?.["w:type"] as string | undefined;
  const pitch = toNum(grid?.attributes?.["w:linePitch"] as string | undefined, 0);
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
    columns: readColumns(sectPr && firstChild(sectPr, "w:cols")),
    pageBorders: readPageBorders(sectPr && firstChild(sectPr, "w:pgBorders")),
    lineNumbers: readLineNumbers(sectPr && firstChild(sectPr, "w:lnNumType")),
  };
}

/**
 * 读 `w:cols`：栏数、栏间距、等宽标记与每栏宽度。
 *
 * @param el body 末尾（或分节符里）那个 `w:cols`；没有就是单栏，返回 undefined。
 * @returns 建模后的分栏；`w:cols` 缺席时 undefined，绝不拿 `count: 1` 冒充"设过分栏"。
 */
function readColumns(
  el: Element | undefined | null,
): SectionColumns | undefined {
  if (!el) return undefined;
  const attrs = el.attributes ?? {};
  const widths = children(el)
    .filter((c) => c.name === "w:col")
    .map((c) => toNum(c.attributes?.["w:w"] as string | undefined, 0) * TWIP_TO_PX);
  // Word 写 `w:equalWidth="0"` 表示不等宽；缺省（没这个属性）按等宽。
  const equalFlag = attrs["w:equalWidth"] as string | undefined;
  return {
    count: toNum(attrs["w:num"] as string | undefined, widths.length || 1),
    gap: toNum(attrs["w:space"] as string | undefined, 0) * TWIP_TO_PX,
    equalWidth: equalFlag !== "0" && equalFlag !== "false",
    widths,
  };
}

/** 边框四条边的固定次序，写出去也按这个次序（规范里 pgBorders 的例子就是它）。 */
const BORDER_SIDES = ["top", "left", "bottom", "right"] as const;
type BorderSide = (typeof BORDER_SIDES)[number];

/**
 * 读 `w:pgBorders`：测量基准加四条边。
 *
 * @param el 节里的 `w:pgBorders`；缺席就是没设过边框。
 * @returns 建模后的边框；`el` 不在时 undefined，不冒充"设过无边框"。
 *          边上的 `w:val="none"`/`"nil"` 收成 null，即"这条边不画"。
 */
function readPageBorders(
  el: Element | undefined | null,
): SectionPageBorders | undefined {
  if (!el) return undefined;
  const borders = el;
  const side = (name: BorderSide): PageBorderSide | null => {
    const kid = firstChild(borders, `w:${name}`);
    if (!kid) return null;
    const attrs = kid.attributes ?? {};
    const style = (attrs["w:val"] as string | undefined) ?? "none";
    if (style === "none" || style === "nil") return null;
    // Word 的 `auto` 收成 null，写回时再还原成 `auto`，这样"用户没动颜色"两边对得上。
    const color = colorVal(attrs["w:color"] as string | undefined);
    return {
      style,
      // `w:sz` 是 1/8 磅，`w:space` 是磅（ECMA-376 §17.3.4）。
      sizePt: toNum(attrs["w:sz"] as string | undefined, 0) / 8,
      spacePt: toNum(attrs["w:space"] as string | undefined, 0),
      color: color?.toLowerCase() === "auto" ? null : color,
    };
  };
  return {
    offsetFrom:
      (borders.attributes?.["w:offsetFrom"] as string | undefined) ?? "page",
    top: side("top"),
    left: side("left"),
    bottom: side("bottom"),
    right: side("right"),
  };
}

/**
 * 读 `w:lnNumType`。
 *
 * @param el 节里的 `w:lnNumType`；缺席就是没开行号。
 * @returns 建模后的行号，`distance` 是 CSS px；`el` 不在时 undefined。
 */
function readLineNumbers(
  el: Element | undefined | null,
): SectionLineNumbers | undefined {
  if (!el) return undefined;
  const attrs = el.attributes ?? {};
  return {
    countBy: toNum(attrs["w:countBy"] as string | undefined, 0),
    distance:
      toNum(attrs["w:distance"] as string | undefined, DEFAULT_LN_DISTANCE_TWIP) *
      TWIP_TO_PX,
    start: toNum(attrs["w:start"] as string | undefined, 1),
  };
}

/**
 * 数值等价判断。属性经过单位往返（px↔twips、pt↔1/8 磅）会有取整误差，
 * 容差按各自的最小单位给，否则"用户什么都没动"也会被当成改动而重建元素。
 *
 * @param x 从文件读出来的值。
 * @param y 保存载荷里的值。
 * @param tolerance 允许差多少（与 x/y 同单位）。
 */
function near(x: number, y: number, tolerance: number): boolean {
  return Math.abs(x - y) <= tolerance;
}

/**
 * 两份分栏设置是否等价——用来决定 `w:cols` 要不要重建。
 *
 * 数值经过 px↔twips 来回，`Math.round` 的误差用 0.5px 容差吸收，
 * 否则"用户什么都没动"也会被判成改动，把 `w:sep` 这类没建模的属性洗掉。
 *
 * @param a 从文件里读出来的那份。
 * @param b 保存载荷里的那份（多半就是 a 走了一趟 IPC 回来的）。
 * @returns 等价则为 true。
 */
function sameColumns(a: SectionColumns, b: SectionColumns): boolean {
  return (
    a.count === b.count &&
    a.equalWidth === b.equalWidth &&
    a.widths.length === b.widths.length &&
    near(a.gap, b.gap, 0.5) &&
    a.widths.every((w, i) => near(w, b.widths[i]!, 0.5))
  );
}

/**
 * 两份页面边框是否等价。线宽经过 pt↔1/8 磅取整，容差给 1/16 磅；
 * 间距经过整数取整，容差半磅。
 */
function samePageBorders(a: SectionPageBorders, b: SectionPageBorders): boolean {
  const sameSide = (x: PageBorderSide | null, y: PageBorderSide | null): boolean => {
    if (!x || !y) return x === y;
    return (
      x.style === y.style &&
      near(x.sizePt, y.sizePt, 1 / 16) &&
      near(x.spacePt, y.spacePt, 0.5) &&
      // 文件里是 `FF0000`，`<input type=color>` 交回来是 `#ff0000`：大小写不算改动。
      x.color?.toLowerCase() === y.color?.toLowerCase()
    );
  };
  return (
    a.offsetFrom === b.offsetFrom &&
    sameSide(a.top, b.top) &&
    sameSide(a.left, b.left) &&
    sameSide(a.bottom, b.bottom) &&
    sameSide(a.right, b.right)
  );
}

/** 两份行号设置是否等价；`distance` 是 px↔twips 往返，容差 0.5px。 */
function sameLineNumbers(a: SectionLineNumbers, b: SectionLineNumbers): boolean {
  return (
    a.countBy === b.countBy &&
    a.start === b.start &&
    near(a.distance, b.distance, 0.5)
  );
}

/**
 * `CT_SectPr` 允许的子孙顺序（ECMA-376 §17.6.1）。重建节属性时按它插入，
 * 排错了 Word 会判文档结构不合法。
 */
const SECT_PR_ORDER = [
  "w:headerReference",
  "w:footerReference",
  "w:footnotePr",
  "w:endnotePr",
  "w:type",
  "w:pgSz",
  "w:pgMar",
  "w:paperSrc",
  "w:pgBorders",
  "w:lnNumType",
  "w:pgNumType",
  "w:cols",
  "w:formProt",
  "w:vAlign",
  "w:noEndnote",
  "w:titlePg",
  "w:textDirection",
  "w:bidi",
  "w:rtlGutter",
  "w:docGrid",
  "w:printerSettings",
];

/**
 * 按 `CT_SectPr` 的顺序把子元素放进 `arr`：已有同名元素就原位替换，
 * 没有就插到第一个"排在它后面"的子元素之前。认不出名字的子元素不参与排序。
 *
 * @param arr 节属性的子元素数组，原地修改。
 * @param name 要放的元素名，如 `w:cols`。
 * @param el 要放的元素。
 */
function placeInSection(arr: Element[], name: string, el: Element): void {
  const rank = SECT_PR_ORDER.indexOf(name);
  const at = arr.findIndex((c) => c.name === name);
  if (at >= 0) {
    arr[at] = el;
    return;
  }
  const after = arr.findIndex(
    (c) => SECT_PR_ORDER.indexOf(c.name) > rank && SECT_PR_ORDER.indexOf(c.name) >= 0,
  );
  if (after >= 0) arr.splice(after, 0, el);
  else arr.push(el);
}

/**
 * 从节属性里删掉某个子元素（用户关掉了这项设置）。
 *
 * @param arr 节属性的子元素数组，原地修改。
 * @param name 要删的元素名。
 */
function removeInSection(arr: Element[], name: string): void {
  const at = arr.findIndex((c) => c.name === name);
  if (at >= 0) arr.splice(at, 1);
}

/**
 * 只有建模值真的变了才重建节属性里的某个子元素。
 *
 * 读进来的元素上常有我们没建模的属性（`w:cols` 的 `w:sep`、`w:lnNumType` 的
 * `w:restart`、`w:pgBorders` 的 `w:display`），值等价就整段原样留着；值变了才按建模
 * 结果重写。
 *
 * @param arr 节属性的子元素数组，原地修改。
 * @param name 子元素名，如 `w:pgBorders`。
 * @param wanted 保存载荷里的值：`undefined` 不碰，`null` 表示关掉（删元素）。
 * @param parse 把已有的元素读回建模值（元素不在就返回 undefined）。
 * @param same 两份建模值是否等价。
 * @param build 按建模值写出元素；返回 null 表示这项设置等价于"没有这个元素"。
 */
function placeOrClear<T>(
  arr: Element[],
  name: string,
  wanted: T | null | undefined,
  parse: (el: Element | undefined) => T | undefined,
  same: (a: T, b: T) => boolean,
  build: (value: T) => Element | null,
): void {
  if (wanted === undefined) return;
  if (wanted === null) {
    removeInSection(arr, name);
    return;
  }
  const existingEl = arr.find((c) => c.name === name);
  const parsed = parse(existingEl);
  if (existingEl && parsed && same(parsed, wanted)) return;
  const built = build(wanted);
  if (built) placeInSection(arr, name, built);
  else removeInSection(arr, name);
}

// ---------- XML element factories ----------

const TWIP_PER_PX = 1440 / 96; // 15
const PX_TO_TWIP = (px: number) => Math.round(px * TWIP_PER_PX);

/**
 * Build or update a <w:sectPr> with the given geometry. pgSz/pgMar are always
 * rewritten from the model; cols/pgBorders/lnNumType are rebuilt only when the
 * model actually differs from what was read, so unmodelled attributes
 * (titlePg, docGrid, headerReference, w:sep, w:restart, …) survive untouched.
 */
function buildSectPr(
  existing: Element | null,
  g: SectionGeometry,
): Element {
  // 与 Word 一致：`w:w` 装的就是页面上看到的宽（横向时是长边），`w:orient` 只是标记。
  const orient: "portrait" | "landscape" =
    g.pageWidth > g.pageHeight ? "landscape" : "portrait";
  const wTw = PX_TO_TWIP(g.pageWidth);
  const hTw = PX_TO_TWIP(g.pageHeight);
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

  // 从空数组开始，和从已有节属性开始走同一条路径：已有的子元素留在数组里，
  // 我们只按 CT_SectPr 的次序放自己的那几项。
  const childrenArr: Element[] = existing
    ? [...(existing.children as unknown as Element[])]
    : [];
  placeInSection(childrenArr, "w:pgSz", newPgSz);
  placeInSection(childrenArr, "w:pgMar", newPgMar);
  // 分栏/页面边框/行号：只有值真的变了才重建，否则整段原样留着，
  // 免得把 `w:sep`、`w:restart`、`w:display` 这类没建模的属性洗掉（AGENTS §2.7）。
  placeOrClear(childrenArr, "w:cols", g.columns, readColumns, sameColumns, colsXml);
  placeOrClear(
    childrenArr,
    "w:pgBorders",
    g.pageBorders,
    readPageBorders,
    samePageBorders,
    pgBordersXml,
  );
  placeOrClear(
    childrenArr,
    "w:lnNumType",
    g.lineNumbers,
    readLineNumbers,
    sameLineNumbers,
    lnNumTypeXml,
  );
  // 必须显式写 `type`/`name`：`parseDocument` 出来的节点只有原型上带 `type`，
  // 展开运算符拿不到，`serialize` 认它不是元素就返回空串——整块 sectPr 静默消失。
  return {
    type: "element",
    name: "w:sectPr",
    attributes: existing?.attributes ?? {},
    children: childrenArr,
  } as unknown as Element;
}

/**
 * 把建模后的分栏写成 `w:cols`。
 *
 * `w:num` 给栏数、`w:space` 给栏间距（twips，换算只在本包发生）；`widths` 非空才写
 * `w:col` 子元素，等宽时留空让 Word 自己算，免得写出一份和 `w:num` 打架的宽度表。
 *
 * @param c 分栏设置。
 * @returns 可以直接放进 `w:sectPr` 的元素。
 */
function colsXml(c: SectionColumns): Element {
  const attrs: Record<string, string> = {
    "w:space": String(PX_TO_TWIP(c.gap)),
  };
  if (c.count > 0) attrs["w:num"] = String(c.count);
  if (!c.equalWidth) attrs["w:equalWidth"] = "0";
  const kids = c.widths.map(
    (w) =>
      ({
        type: "element",
        name: "w:col",
        attributes: { "w:w": String(PX_TO_TWIP(w)) },
        children: [],
      }) as unknown as Element,
  );
  return {
    type: "element",
    name: "w:cols",
    attributes: attrs,
    children: kids,
  } as unknown as Element;
}

/**
 * 把建模后的页面边框写成 `w:pgBorders`，四条边按 top→left→bottom→right 排。
 *
 * @param b 边框设置。
 * @returns 可放进 `w:sectPr` 的元素；四条边都不画时返回 null，语义是"删掉这个元素"。
 */
function pgBordersXml(b: SectionPageBorders): Element | null {
  const kids = BORDER_SIDES.map((name) => ({ name, side: b[name] }))
    .filter((e): e is { name: BorderSide; side: PageBorderSide } => e.side !== null)
    .map(({ name, side }) => {
      const attrs: Record<string, string> = {
        "w:val": side.style,
        "w:sz": String(Math.round(side.sizePt * 8)),
        "w:space": String(Math.round(side.spacePt)),
        // 读不懂的颜色宁可写成 auto，也不能把 `rgb(...)` 塞进 w:color（AGENTS §2.2）。
        "w:color": cssColorToOoxmlHex(side.color) ?? "auto",
      };
      return {
        type: "element",
        name: `w:${name}`,
        attributes: attrs,
        children: [],
      } as unknown as Element;
    });
  if (!kids.length) return null;
  return {
    type: "element",
    name: "w:pgBorders",
    attributes: { "w:offsetFrom": b.offsetFrom },
    children: kids,
  } as unknown as Element;
}

/**
 * 把建模后的行号写成 `w:lnNumType`。
 *
 * @param v 行号设置；`w:start` 只在不是 Word 缺省的 1 时才写。
 * @returns 可放进 `w:sectPr` 的元素；`countBy` 非正数时返回 null，即"删掉这个元素"，
 *          没有 `w:lnNumType` 就是不标行号。
 */
function lnNumTypeXml(v: SectionLineNumbers): Element | null {
  if (v.countBy <= 0) return null;
  const attrs: Record<string, string> = {
    "w:countBy": String(v.countBy),
    "w:distance": String(PX_TO_TWIP(v.distance)),
  };
  if (v.start !== 1) attrs["w:start"] = String(v.start);
  return {
    type: "element",
    name: "w:lnNumType",
    attributes: attrs,
    children: [],
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

/**
 * 往已解析的 rels 根上追加一条文档关系，同一个 `@Id` 已经在了就不重复加。
 *
 * @param relsRoot `word/_rels/document.xml.rels` 解析出来的根元素，原地改它的 children。
 * @param id 关系的 `@Id`，正文里的 `r:id` 就指着它。
 * @param typeUrl 完整的 `@Type` URL（不是缩写）。
 * @param target 相对 `word/` 的目标路径，如 "header1.xml"。
 * @returns 真的加了这条关系才 true；调用方据此决定要不要重写部件。
 */
function addRelationship(
  relsRoot: Element,
  id: string,
  typeUrl: string,
  target: string,
): boolean {
  if (children(relsRoot).some((c) => c.attributes?.["Id"] === id)) return false;
  (relsRoot.children as unknown as Element[]).push({
    type: "element",
    name: "Relationship",
    attributes: { Id: id, Type: typeUrl, Target: target },
    children: [],
  } as unknown as Element);
  return true;
}
