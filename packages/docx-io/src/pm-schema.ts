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

export type PMNode =
  | PMTextNode
  | {
      type: "paragraph";
      attrs?: {
        textAlign?: string;
        pageBreakBefore?: boolean;
        /**
         * Raw <w:pPr> inner XML (everything between <w:pPr> and </w:pPr>)
         * captured from the source document. pmToOoxml merges it with the
         * attrs it models: <w:jc>/<w:spacing>/<w:ind> are rebuilt from both
         * sides, every other element (borders, keepLines, tabs, outlineLvl,
         * …) rides through untouched. See `mergedElementXml` in ooxml-pm.ts.
         */
        rawPpr?: string;
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
      };
      content?: PMNode[];
    }
  | {
      type: "heading";
      attrs: {
        level: 1 | 2 | 3 | 4 | 5 | 6;
        textAlign?: string;
        spacingBefore?: number;
        spacingAfter?: number;
        indent?: number;
        indentRight?: number;
        indentFirst?: number;
        lineHeight?: string;
      };
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
