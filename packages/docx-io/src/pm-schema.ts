/**
 * ProseMirror schema that matches the TipTap extensions used by the desktop
 * app (StarterKit + Underline + TextAlign + Link + Image + Table). We keep
 * this schema intentionally narrow; anything not represented here is wrapped
 * in an `unmanagedBlock` and round-tripped verbatim through XML.
 *
 * We do NOT depend on ProseMirror at this layer: the types are plain JSON so
 * the converters can run in Node, Electron main, or the renderer without a
 * PM instance. The renderer feeds the result straight into TipTap's
 * editor.commands.setContent().
 */

export type PMMark =
  | { type: "bold" }
  | { type: "italic" }
  | { type: "underline" }
  | { type: "strike" }
  | { type: "code" }
  | {
      type: "textStyle";
      attrs: {
        /** Chinese / East-Asian face (`w:eastAsia`). */
        fontFamily?: string | null;
        /**
         * Latin face (`w:ascii`/`w:hAnsi`/`w:cs`). Only produced when the
         * document really does use a different face for Western text — the
         * common 中文宋体 + 数字 Times New Roman pairing.
         */
        fontFamilyWest?: string | null;
        /** Font size in CSS pixels (OOXML half-points are converted on the fly). */
        fontSize?: string | null;
        /** Letter spacing in CSS pixels; OOXML stores twentieths of a point. */
        letterSpacing?: string | null;
        color?: string | null;
        backgroundColor?: string | null;
      };
    }
  | { type: "link"; attrs: { href: string; target?: string | null } };

export interface PMTextNode {
  type: "text";
  text?: string;
  marks?: PMMark[];
}

/**
 * 段落与标题共用的 `<w:pPr>` 建模属性（用 type 而非 interface：写回侧把它当
 * `Record<string, unknown>` 遍历，type 才有隐式索引签名）。
 *
 * 单位约定（AGENTS §2.1，换算只发生在 docx-io 进出口）：段前/段后/缩进是 pt，
 * `lineHeight` 是倍数或 `pt:N` / `atLeast:N`。
 * 每一项 `undefined` 的含义都是"这一段没建模过这个属性"，写回时保持源文档原样，
 * 而不是补一个默认值——这是重复保存不侵蚀用户格式的前提。
 */
export type PPrAttrs = {
  textAlign?: string;
  /**
   * 这一段套用的具名样式 id（`w:pStyle`/@w:val），原样存 Word 写的那个值，
   * 不做名称换算——styles.xml 里 id 与名称的对应关系只有源文档知道。
   * undefined = 这一段没建模过样式（保留源文档里的 pStyle）；
   * null = 用户在编辑器里主动"清除样式"，源文档那条 pStyle 必须删掉。
   */
  styleId?: string | null;
  /** 段前分页（w:pageBreakBefore）。 */
  pageBreakBefore?: boolean;
  /** 与下段同页（w:keepNext），Word 勾选项"段与下段同页"。 */
  keepWithNext?: boolean;
  /** 段中不分页（w:keepLines），同一段的文字不跨页拆开。 */
  keepLines?: boolean;
  /** 孤行控制（w:widowControl）。Word 默认开，显式关闭才写 w:val="0"。 */
  widowControl?: boolean;
  /** 对齐到网格（w:snapToGrid），关闭后本段不受文档网格行距约束。 */
  snapToGrid?: boolean;
  /** 大纲级别，直接存 OOXML 的 0-9：0 是"1 级"，9 是"正文文本"。 */
  outlineLevel?: number;
  /** Spacing before the paragraph, in points. */
  spacingBefore?: number;
  /** Spacing after the paragraph, in points. */
  spacingAfter?: number;
  /** Left indent, in points. */
  indent?: number;
  /** Right indent, in points (Word's 段落对话框 → 右侧). */
  indentRight?: number;
  /** First-line indent, in points. Negative = hanging indent. */
  indentFirst?: number;
  /** Line height as a CSS-like number (1.0, 1.15, 1.5, 2.0). */
  lineHeight?: string;
  /**
   * Raw <w:pPr> inner XML (everything between <w:pPr> and </w:pPr>)
   * captured from the source document. pmToOoxml merges it with the
   * attrs it models: the modeled children are rebuilt from both sides,
   * every other element (borders, tabs, run properties, …) rides through
   * untouched and keeps its place in CT_PPr's sequence. See
   * `mergedElementXml` and `PPR_CHILD_ORDER` in ooxml-pm.ts.
   */
  rawPpr?: string;
};

export type PMNode =
  | PMTextNode
  | {
      type: "paragraph";
      attrs?: PPrAttrs;
      content?: PMNode[];
    }
  | {
      type: "heading";
      attrs: { level: 1 | 2 | 3 | 4 | 5 | 6 } & PPrAttrs;
      content?: PMNode[];
    }
  | { type: "bulletList"; content?: PMNode[]; attrs?: { numId?: string } }
  | { type: "orderedList"; attrs?: { start?: number; numId?: string }; content?: PMNode[] }
  | { type: "listItem"; content?: PMNode[] }
  | { type: "blockquote"; content?: PMNode[] }
  | { type: "codeBlock"; attrs?: { language?: string | null }; content?: PMNode[] }
  | { type: "horizontalRule" }
  | { type: "hardBreak" }
  | {
      type: "image";
      attrs: {
        src: string;
        alt?: string | null;
        title?: string | null;
        width?: number | null;
        height?: number | null;
        /**
         * When present, this drawing was imported from a .docx and we keep
         * its exact <w:drawing>…</w:drawing> XML so that floating anchors
         * (wp:anchor), text wrapping, and effect extents round-trip. The
         * editor renders the resolved data URL at roughly the right size.
         */
        ooxml?: string | null;
        /** "anchor" (floating) or "inline" (flowing with text). */
        layout?: "anchor" | "inline";
        name?: string | null;
        /** Floating-image position in CSS pixels, converted from EMU. */
        posX?: number | null;
        posY?: number | null;
        /** wrapNone | wrapSquare | wrapTight | wrapThrough | wrapTopAndBottom */
        wrap?: "none" | "square" | "tight" | "through" | "topAndBottom";
      };
    }
  | {
      type: "pageBreak";
    }
  | {
      /**
       * A Word field: PAGE (current page number), NUMPAGES (total page
       * count), DATE, AUTHOR, FILENAME. Serialized as a complex-field
       * run sequence (`<w:fldChar w:fldCharType="begin"/> …`) on save.
       * The `display` attr is the cached result shown in the editor and
       * written into `<w:t>` so Word/WPS can display something even
       * before it recalculates fields.
       */
      type: "field";
      attrs: {
        kind: "page" | "numPages" | "date" | "author" | "fileName" | string;
        instr?: string;
        display: string;
      };
    }
  | {
      /**
       * 书签的起点：OOXML 的 `<w:bookmarkStart w:id=".." w:name=".."/>`。
       *
       * 它是 `<w:p>` 的直接子元素（与 `<w:r>` 平级，ECMA-376 §17.5.2.1 的例子就是
       * 夹在两个 run 之间），所以建模成行内节点才能保住它在正文流里的位置；
       * `w:id` 是 ST_DecimalNumber，原样存字符串不做数字转换（§2.1：换算只在进出口）。
       */
      type: "bookmarkStart";
      attrs: { id: string; name?: string };
    }
  | {
      /**
       * 书签的终点：`<w:bookmarkEnd w:id=".."/>`，与起点靠同一个 `id` 配对。
       *
       * 一个书签可以跨段落（起在第一段、止在第三段），因此这一对节点各自落在自己那段里，
       * 不需要任何跨段的位置信息。
       */
      type: "bookmarkEnd";
      attrs: { id: string };
    }
  | {
      type: "table";
      attrs?: Record<string, unknown>;
      content?: PMNode[];
    }
  | { type: "tableRow"; content?: PMNode[] }
  | {
      type: "tableCell" | "tableHeader";
      attrs?: {
        colspan?: number;
        rowspan?: number;
        colwidth?: number[] | null;
        textAlign?: string;
      };
      content?: PMNode[];
    }
  | {
      /**
       * A floating drawing (Text Box / AutoShape / group of them) anchored at
       * absolute page or paragraph coordinates. Word templates — resumes in
       * particular — put essentially all body text inside these, so the editor
       * has to model them as positioned containers rather than dropping them.
       *
       * `ooxml` keeps the source `<mc:AlternateContent>`/`<w:drawing>`/`<w:pict>`
       * element verbatim as a template. On save we splice each child's
       * `<w:txbxContent>` back into it, so the group transform, fills and the
       * legacy VML fallback branch survive untouched.
       */
      type: "floatGroup";
      attrs: {
        ooxml: string;
        /** Offset from `relH`/`relV`, in CSS px. */
        x: number;
        y: number;
        relH: string;
        relV: string;
        /** `wp:anchor/@relativeHeight`: Word's stacking order for this drawing. */
        z: number;
      };
      content?: PMNode[];
    }
  | {
      /** One shape inside a `floatGroup`, positioned relative to the anchor. */
      type: "floatBox";
      attrs: {
        x: number;
        y: number;
        w: number;
        h: number;
        fill: string | null;
        stroke: string | null;
        strokeWidth: number | null;
        vAlign: string | null;
        name: string | null;
        /**
         * The shape's original `<w:txbxContent>` inner XML. While `dirty` is
         * false this is written back verbatim; the editor flips `dirty` once the
         * content really changes and the box is rebuilt from ProseMirror.
         */
        txbx: string | null;
        dirty: boolean;
      };
      content?: PMNode[];
    }
  | {
      /**
       * Block the editor does not model. We keep its raw OOXML so save can
       * write it back unchanged (SmartArt, charts, embedded objects, etc.).
       */
      type: "unmanagedBlock";
      attrs: { kind: string; ooxml: string; label?: string };
    };

export interface PMDoc {
  type: "doc";
  content: PMNode[];
}
