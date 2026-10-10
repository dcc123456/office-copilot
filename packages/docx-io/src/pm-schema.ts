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

/**
 * `w:vertAlign` 的 `w:val`（ECMA-376 §17.3.2.27 的 ST_VerticalAlignRun，三档：
 * `baseline` 是缺省位置、`superscript` 抬起来、`subscript` 沉下去，后两档同时按
 * "有则缩小"排字）。字号缩小是**排版行为**，文档里不写 `w:sz`，所以换算不涉及这里。
 */
export type VerticalAlign = "baseline" | "superscript" | "subscript";

export type PMMark =
  | { type: "bold" }
  | { type: "italic" }
  | { type: "underline" }
  | { type: "strike" }
  | { type: "code" }
  /** 上标 / 下标。OOXML 只有一个 `w:vertAlign` 元素，所以这里也是一个标记带枚举值，而不是两个标记。 */
  | { type: "vertAlign"; attrs: { val: VerticalAlign } }
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
  | { type: "link"; attrs: { href: string; target?: string | null } }
  /**
   * 修订跟踪：`w:ins`（插入）/ `w:del`（删除）包住的运行（ECMA-376 的 `CT_RunTrackChange`）。
   *
   * 做成 mark 而不是节点：修订是**字符属性**（同一段里可以一半是插入一半不是），而且必须让
   * 被改的文字留在原地——用节点表示就得把插入的文字搬进一个原子节点，光标、选区、撤销全乱。
   * `kind` 存原始操作名（不是"接受/拒绝"后的语义），`id`/`author`/`date` 一律**原样存字符串**：
   * 单位换算与归一化只在 docx-io 进出口（§2.1），而 `w:id` 是文档级编号空间，改动会让 Word
   * 的"接受此修订"找不到对象。
   *
   * 嵌套（`w:del` 里再包 `w:ins`）只保留**外层**那一次操作：外层是后发生的，Word 显示的正是
   * 它的状态（内层 `w:delText` 会被外层的 `w:t` 语义覆盖）。这是刻意的取舍，见规格排名 4 判据。
   */
  | {
      type: "revision";
      attrs: {
        kind: "ins" | "del";
        id?: string | null;
        author?: string | null;
        date?: string | null;
      };
    };

export interface PMTextNode {
  type: "text";
  text?: string;
  marks?: PMMark[];
}

/** `<w:tab>` 的 `w:val`；`clear` 是 Word 自己的写法："这个位置上的制表位在此作废"。 */
export type TabStopAlignment =
  | "left"
  | "center"
  | "right"
  | "decimal"
  | "bar"
  | "list"
  | "clear";

/** `<w:tab>` 的 `w:leader`，即 Word"前导符"下拉框里的七种。 */
export type TabStopLeader =
  | "none"
  | "dot"
  | "hyphen"
  | "dash"
  | "line"
  | "heavy"
  | "middleDot";

/**
 * 一个制表位，即 `<w:pPr><w:tabs><w:tab w:val w:leader w:pos/>`（ECMA-376 §17.3.1.1 的
 * CT_TabStop）。与"这一段没定义制表位"不是一回事的 `clear` 原样保留，不做归一化。
 */
export type TabStop = {
  /** 位置，pt（导出 ×20 成 twips；页面坐标从文字区左边界起算）。 */
  positionPt: number;
  /** `w:val`。缺省按 left，与 Word 一致。 */
  alignment: TabStopAlignment;
  /** `w:leader`。缺省按 none，即空白前导。 */
  leader?: TabStopLeader;
};

/**
 * 首字下沉的位置，直接沿用 `w:dropCap` 的合法值（ECMA-376 §17.3.1.11 ST_DropCap）：
 * `drop` = Word UI 的"下沉"（首字缩进正文区内），`margin` = "首字突出到页边距外"。
 * `none` 不进模型——那是"关掉"，用 attrs 上的 null 表示。
 */
export type DropCapPosition = "drop" | "margin";

/** 首字下沉：位置 + 跨几行（`w:lines`，正整数）。 */
export type DropCap = {
  position: DropCapPosition;
  /**
   * 下沉行数。源文档没写 `w:lines` 时保持 undefined：写侧因此不会补一个我们猜的默认值，
   * 未改动的段落重复保存才能字节一致。
   */
  lines?: number;
};

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
  /**
   * 段落的制表位定义（`<w:tabs>`）。undefined = 这一段没建模过（源文档的 `w:tabs` 原样透传）；
   * 空数组 = 用户在"制表位"里全部清除，写出的 pPr 不再带 `w:tabs`，源文档那条也不留——
   * 这正是 Word 删掉直接制表位后的行为（样式给的制表位仍然生效）。
   */
  tabStops?: TabStop[];
  /** Line height as a CSS-like number (1.0, 1.15, 1.5, 2.0). */
  lineHeight?: string;
  /**
   * 首字下沉（`<w:framePr>` 上的 `w:dropCap` + `w:lines`）。
   * undefined = 这一段没建模过（源文档的 `w:framePr` 原样透传）；
   * null = 用户选了"无"——只带首字下沉属性的框架整个删掉（`w:framePr` 的**存在本身**
   *   就表示"这段是文本框"，留一个空元素会把正文段落变成匿名文本框），
   *   真文本框（还有 `w:w`/`w:h`/`w:hRule` 之类）只把 `w:dropCap` 改写成 none；
   * 有值 = 位置 + 行数。
   */
  dropCap?: DropCap | null;
  /**
   * Raw <w:pPr> inner XML (everything between <w:pPr> and </w:pPr>)
   * captured from the source document. pmToOoxml merges it with the
   * attrs it models: the modeled children are rebuilt from both sides,
   * every other element (borders, run properties, …) rides through
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
       *
       * `display` 是缓存结果的**字面文本**：`\t` / `\n` 分别对应 OOXML 的
       * `<w:tab/>` / `<w:br/>`（目录条目的"标题⇥页码"就靠它对齐），空串表示 Word
       * 还没算过结果——我们不替用户编一个占位文字写回文档。
       *
       * 边界：一个节点只表示**同段落内** begin/end 配对的域。Word 自己生成的多行目录
       * 是 begin 在首段、end 在末段，这种跨段域用 `fieldBegin` / `fieldEnd` 建模。
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
       * 跨段落复合域的**起始边界**：`<w:r><w:fldChar w:fldCharType="begin"/></w:r>`
       * 加上紧随其后的 `<w:instrText>`（指令存进 `instr`，不再作为可见文字），
       * 以及可选的 `<w:fldChar w:fldCharType="separate"/>`（`isSeparated` 记录有没有写）。
       *
       * Word 的目录就是这个形状：首段写 begin+instr+separate，条目段落是域的结果，
       * 末段末尾写 end。用 `field` 表示会把整棵目录吞成一段缓存文本，所以这里只落边界，
       * 边界之后的内容照常按结构读（条目里的 PAGEREF 仍是 `field` 节点）。
       */
      type: "fieldBegin";
      attrs: {
        kind: string;
        instr?: string;
        /** 是否已经写下 separate 记号；false 表示 Word 还没算过结果。 */
        isSeparated?: boolean;
      };
    }
  | {
      /** 跨段落复合域的结束边界，即 `<w:r><w:fldChar w:fldCharType="end"/></w:r>`。 */
      type: "fieldEnd";
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
      /**
       * 脚注引用点：`<w:r><w:footnoteReference w:id="N"/></w:r>`（ECMA-376 §17.11.18）。
       *
       * 正文里只有这一个引用点，脚注的**内容**住在 `word/footnotes.xml` 部件里，
       * 靠同一个 `w:id` 对应；所以这个节点是行内 atom，不带 content。
       * `w:id` 是 ST_DecimalNumber，原样存字符串（可以是负数，部件里的分隔符条目就是 -1）。
       */
      type: "footnoteReference";
      attrs: { id: string };
    }
  | {
      /** 尾注引用点，形状与脚注一致，内容部件是 `word/endnotes.xml`。 */
      type: "endnoteReference";
      attrs: { id: string };
    }
  | {
      /**
       * 批注区间的起点：`<w:commentRangeStart w:id="N"/>`（属 `EG_RangeMarkupElements`，
       * 位置见 spec 排名 4 的内容模型那条）。
       *
       * 它与 `<w:r>` 平级、是 `<w:p>` 的直接子元素（属 `EG_RangeMarkupElements` ⊂ `EG_PContent`），
       * 所以必须建模成行内节点才能保住它在正文流里的位置；批注**内容**住在
       * `word/comments.xml` 部件里，靠同一个 `w:id` 对应，因此节点本身是空 atom。
       * 起点与终点可以落在不同段落里（区间跨段是合法的），两侧各记一次位置即可。
       */
      type: "commentRangeStart";
      attrs: { id: string };
    }
  | {
      /**
       * 批注区间的终点：`<w:commentRangeEnd w:id="N"/>`，与起点靠 `id` 配对。
       *
       * Word 习惯紧跟着写一个带 `w:commentReference` 的运行；那一跑我们独立建模成
       * `commentReference` 节点，不与此合并，因为读到的位置可能不在紧邻处（用户手改过）。
       */
      type: "commentRangeEnd";
      attrs: { id: string };
    }
  | {
      /**
       * 批注引用记号：`<w:r><w:commentReference w:id="N"/></w:r>`（属 `EG_RunInnerContent`，
       * 因此必须在 run 里）。
       *
       * 形状与 `footnoteReference` 一致：正文里只有引用点，编号由 Word 在版面上算，
       * 文件里不存；外面那个运行的 `w:rPr`（上标 + `AnnotationReference` 样式）是呈现，
       * 读侧忽略、写侧固定补，因此读→写→读稳定。
       */
      type: "commentReference";
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

/**
 * 一份只有一个空段落的 doc：透明部件（脚注、批注）里新建的条目还没有正文时用。
 *
 * @returns 每次调用都返回新对象——共享一个常量的话，一处编辑会把别的条目一起改掉。
 */
export function emptyPmDoc(): PMDoc {
  return { type: "doc", content: [{ type: "paragraph", content: [] }] };
}
