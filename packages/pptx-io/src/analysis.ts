/**
 * `.pptx` 的分析读面：把读模型摊成**给 AI 看的结构化画像**（`docs/office-parity-spec.md` 阶段 5、D6）。
 *
 * 与表格侧 `sheet-io` 的 `describeSheet`/`readRange` 同一个分工（`AGENTS.md` §3.3、§2.11）：
 * 页数、形状清单、逐 run 的文字与格式这些**事实**由引擎从现场模型算，工具包（`pptx-tools`）只管入参合不合法，
 * 代理层（`apps/desktop/src/ai/slideAgent.ts`）只搬运——模型不许自己数页、数形状、编 shapeId。
 *
 * 上送内容默认截断（§2.9）：页数、每页形状数、每段 run 数各有上限，单条 run 文字 ≤120 字，
 * 被截掉的部分如实标 `truncated`，让模型接着调 `readSlideText` 往下读，而不是误以为那就是全部。
 */
import { clipText, describeSlideReadFailure } from "@office-copilot/shared";
import { type PresentationModel, type RunFormat, type Slide, type SlideReadFailure, type SlideShape } from "./read.js";
import { emuToPx } from "./units.js";

/** 单条 run 文字上送的上限（字数，含被截断的省略号；§2.9 的表格侧同名单格 120 字）。 */
const RUN_TEXT_LIMIT = 120;

/** 一条 run 摊开给模型看的样子：定位坐标 + 现文字 + 现格式。 */
export interface SlideRunSnapshot {
  /** 段落序号（0 基，与模型提交改动时用的 `paraIndex` 同一个）。 */
  paraIndex: number;
  /** run 序号（0 基）。 */
  runIndex: number;
  /** 这条 run 的现文字；超长时截到 {@link RUN_TEXT_LIMIT} 字并置 `textTruncated`。 */
  text: string;
  /** 文字是否被截断（§2.9：截断要说出来，不许让模型以为看到全部）。 */
  textTruncated: boolean;
  /** 这条 run 自己显式写了的格式；全是 `undefined` 表示这一截的字号/粗体/颜色都继承自版式（D3，不许猜）。 */
  format: RunFormat;
}

/** 表格里一格的快照（行、列就是它在 `rows` 里的位置，与引擎 `TableCellRunTarget` 同一口径）。 */
export interface SlideTableCellSnapshot {
  /** 这一格的文字（各段用 `⏎` 隔开，超长已截）；续格画面上没有字，文本仍照实给。 */
  text: string;
  /** 每段的 `<a:r>` 截数：>1 表示整段覆盖会拍平局部格式（§2.3），只能逐 run 改。 */
  paragraphRunCounts: number[];
  /** 是不是合并区里的**续格**（`hMerge`/`vMerge`）——PowerPoint 不画它的字，对它提改动等于静默无效。 */
  isMergeContinuation: boolean;
}

/** 一个形状摊开给模型看的样子。 */
export interface SlideShapeTextSnapshot {
  /** 形状 `<p:cNvPr @id>`，模型提交改动时用它定位。 */
  shapeId: number;
  /** 形状名。 */
  name: string;
  /** 这个形状能不能改文字（只有 `p:sp` 这类有文本体的才行）。 */
  hasText: boolean;
  /** 逐 run 的快照，按「段—run」顺序排。 */
  runs: SlideRunSnapshot[];
  /** 每段是否只有一条 run（true 才允许整段覆盖式改文字；false 必须逐 run 改，§2.3）。 */
  singleRunParagraphs: boolean[];
  /**
   * 图形框里那张表的逐格快照（`<a:tbl>`）；不是表的形状为 undefined。
   *
   * 为什么单列：表的字不写在形状的 `p:txBody` 里，段/run 序号是**每一格里重新数**的，
   * 上面那套 `runs` 装不下（也不许装——把格压成段就等于编了一套寻址）。
   */
  table?: SlideTableCellSnapshot[][];
}

/** `readSlideText` 的结果。 */
export interface SlideTextReport {
  /** 页号（1 基）。 */
  slideNumber: number;
  /** 该页部件路径，便于排查。 */
  partPath: string;
  /** 页上的形状快照（只有带文字体的形状）。 */
  shapes: SlideShapeTextSnapshot[];
  /** 每段最多列多少条 run，超出的 run 没上送。 */
  runsTruncated: boolean;
}

/** 一页给模型看的概要。 */
export interface SlideOutlineEntry {
  /** 页号（1 基，与 `ppt_read_slide_text` 的入参同一个口径）。 */
  slideNumber: number;
  /** 部件路径。 */
  partPath: string;
  /** 该页形状总数（`maxShapesPerSlide` 截断时用它说明只列了前面几个）。 */
  shapeCount: number;
  /** 列出来的形状。 */
  shapes: Array<{
    /** 形状号，改几何与改文字都靠它。 */
    shapeId: number;
    /** 形状名。 */
    name: string;
    /** 大类：形状 / 图片 / 图形框 / 组合 / 连接线。 */
    kind: SlideShape["kind"];
    /** 占位符角色（`type` 或 `idx`），非占位符为 null。 */
    placeholderRef: string | null;
    /** 有没有自己的 `<a:xfrm>`（false＝几何继承自版式，改几何要把四个值一起给，D3）。 */
    hasOwnGeometry: boolean;
    /**
     * 这个形状是不是装在某个组合（`<p:grpSp>`）里面。
     *
     * 组里的形状只许改文字/字体/颜色与删掉，**不许改几何**：它的坐标写在组内的子坐标系里，换算回页面位置
     * 要连着组的缩放与旋转一起算，这一档不做（引擎那条会整条拒绝并给理由）。叠放次序是"同一层内"的性质，
     * 所以调层级照做，只是只在它所在的那个组里换序。
     */
    inGroup: boolean;
    /** 这一页上这个形状的文字概要（拼接串，超长已截）。 */
    text: string;
    /**
     * 图形框里那张表的逐格概要（行、列就是位置下标，续格标了 `isMergeContinuation`）；
     * 不是表的形状为 null。为什么概要里也要带：`text` 对表恒为空，模型会把一页的表格念成"图形框，没文字"。
     */
    table: SlideTableCellSnapshot[][] | null;
  }>;
  /** 形状清单是否被 `maxShapesPerSlide` 截断。 */
  shapesTruncated: boolean;
  /**
   * 这一页的演讲者备注概要（各段用 `⏎` 隔开，超长已截）；没有备注部件时 null。
   *
   * 与"有备注页但一个字都没写"的差别靠 `notesPartPath` 分：后者是空串。模型回答"这页有没有备注"时
   * 不许把"没有备注页"说成"备注是空的"。
   */
  notes: string | null;
  /** 备注部件路径；没有该部件时 null。 */
  notesPartPath: string | null;
  /**
   * 这一页**读不懂**的码（`missing-part` / `invalid-xml` / `no-sp-tree`），读得懂时 null。
   *
   * 与 `shapeCount: 0` 的区别必须让模型看得见：后者才是"这页真的没形状"，前者要说成"这页读不动"。
   */
  readFailure: SlideReadFailure | null;
}

/** `describeDeck` 的结果。 */
export interface DeckOutline {
  /** 总页数。 */
  slideCount: number;
  /** 页面尺寸（CSS px @96DPI；模型视角一律 px，EMU 换算留在引擎边，§2.1）。 */
  slideSizePx: { widthPx: number; heightPx: number };
  /** 列出来的页（`maxSlides` 截断时后面的页不在这里，但 `slideCount` 仍是全量）。 */
  slides: SlideOutlineEntry[];
  /** 页清单是否被截断。 */
  slidesTruncated: boolean;
}

/** `describeDeck` 的入参上限。 */
export interface DescribeOptions {
  /** 最多描述多少页（1~200）。 */
  maxSlides: number;
  /** 每页最多列多少个形状（1~200）。 */
  maxShapesPerSlide: number;
}

/**
 * 把形状的文字概要拼成一条截断串（描述用，不含逐 run 定位）。
 *
 * @param shape 形状快照。
 * @returns 各段用 `⏎` 隔开的概要文字。
 */
function shapeTextSummary(shape: SlideShape): string {
  // 表里的字不写在 `p:txBody`（在每格自己的文本体里），所以概要保持现状，只额外带出逐格网格。
  return clipText(shape.paragraphTexts.join("⏎"), RUN_TEXT_LIMIT).text;
}

/**
 * 把形状里那张表摊成逐格快照（供概要 `table` 与 `readSlideText` 两处上送）。
 *
 * @param shape 形状快照。
 * @returns 逐行逐格（行列下标即位置，与引擎 `TableCellRunTarget` 同一口径）；不是表时 undefined。
 *   每格文字超长时按 §2.9 逐格截断，续格照实标出来——模型对着续格提改动等于改了个不画出来的地方。
 */
function tableSnapshot(shape: SlideShape): SlideTableCellSnapshot[][] | undefined {
  if (!shape.table) return undefined;
  return shape.table.rows.map((row) =>
    row.map((cell) => ({
      text: clipText(cell.paragraphTexts.join("⏎"), RUN_TEXT_LIMIT).text,
      paragraphRunCounts: cell.paragraphRunCounts,
      isMergeContinuation: cell.isMergeContinuation,
    })),
  );
}

/**
 * 描述整份演示文稿：页 → 形状 → 该页文字概要。
 *
 * @param model 主进程读出来的模型。
 * @param options 页数与每页形状数的上限（工具入参已校验过范围，这里只照做）。
 * @returns 结构化画像；被截断的页与形状都置了 `slidesTruncated` / `shapesTruncated`，不静默丢。
 */
export function describeDeck(model: PresentationModel, options: DescribeOptions): DeckOutline {
  const slides: SlideOutlineEntry[] = [];
  for (const slide of model.slides.slice(0, options.maxSlides)) {
    slides.push(outlineOfSlide(slide, slide.order + 1, options.maxShapesPerSlide));
  }
  return {
    slideCount: model.slides.length,
    slideSizePx: { widthPx: emuToPx(model.slideSizeEmu.cx), heightPx: emuToPx(model.slideSizeEmu.cy) },
    slides,
    slidesTruncated: model.slides.length > slides.length,
  };
}

/**
 * 描述一页的形状清单。
 *
 * @param slide 一页的快照。
 * @param slideNumber 页号（1 基）。
 * @param maxShapesPerSlide 每页最多列多少个形状。
 * @returns 该页概要，形状超出时只列前面那些并标记截断。
 */
function outlineOfSlide(slide: Slide, slideNumber: number, maxShapesPerSlide: number): SlideOutlineEntry {
  const listed = slide.shapes.slice(0, maxShapesPerSlide);
  return {
    slideNumber,
    partPath: slide.partPath,
    shapeCount: slide.shapes.length,
    shapes: listed.map((shape) => ({
      shapeId: shape.shapeId,
      name: shape.name,
      kind: shape.kind,
      placeholderRef: shape.placeholderRef ?? null,
      hasOwnGeometry: shape.hasOwnGeometry,
      inGroup: shape.groupShapeId !== undefined,
      text: shapeTextSummary(shape),
      table: tableSnapshot(shape) ?? null,
    })),
    shapesTruncated: slide.shapes.length > listed.length,
    readFailure: slide.readFailure ?? null,
    notes: slide.notes ? clipText(slide.notes.paragraphTexts.join("⏎"), RUN_TEXT_LIMIT).text : null,
    notesPartPath: slide.notes?.partPath ?? null,
  };
}

/**
 * 把某一页每个带文字的形状摊成「段 → run」清单，供模型对准 `paraIndex` / `runIndex`。
 *
 * @param model 主进程读出来的模型。
 * @param slideNumber 页号（1 基；越界时返回 `ok:false` 与给模型看的中文理由，不抛异常）。
 * @param maxRunsPerParagraph 每段最多列多少条 run（超出置 `runsTruncated`）。
 * @returns 该页的文字报告，或一句中文失败理由（页号越界；或那一页我们**读不动**——这时绝不回一份空形状，
 *          免得模型把"读不懂"当成"这一页是空的"）。
 */
export function readSlideText(
  model: PresentationModel,
  slideNumber: number,
  maxRunsPerParagraph: number,
): { ok: true; report: SlideTextReport } | { ok: false; reason: string } {
  const slide = model.slides.find((entry) => entry.order + 1 === slideNumber);
  if (!slide) return { ok: false, reason: `没有第 ${slideNumber} 页：这份演示文稿一共 ${model.slides.length} 页，先用 ppt_describe 看页号。` };
  if (slide.readFailure) {
    return { ok: false, reason: `第 ${slideNumber} 页读不动：${describeSlideReadFailure(slide.readFailure)}这一页不是"空的"，别对它提改动。` };
  }
  const shapes: SlideShapeTextSnapshot[] = [];
  let runsTruncated = false;
  for (const shape of slide.shapes) {
    if (shape.paragraphTexts.length === 0) {
      // 图片/图形框/连接线没有文本体：列出来只会误导模型去改它。
      // 但**表**是个例外——它的字写在每一格自己的文本体里，`paragraphTexts` 恒空不等于"这格没内容"。
      const table = tableSnapshot(shape);
      if (table) shapes.push({ shapeId: shape.shapeId, name: shape.name, hasText: false, runs: [], singleRunParagraphs: [], table });
      continue;
    }
    const runs: SlideRunSnapshot[] = [];
    shape.paragraphTexts.forEach((_, paraIndex) => {
      const paragraphRuns = shape.paragraphRunTexts[paraIndex] ?? [];
      paragraphRuns.slice(0, maxRunsPerParagraph).forEach((runText, runIndex) => {
        const clipped = clipText(runText, RUN_TEXT_LIMIT);
        runs.push({
          paraIndex,
          runIndex,
          text: clipped.text,
          textTruncated: clipped.truncated,
          format: shape.paragraphRunFormats[paraIndex]?.[runIndex] ?? {},
        });
      });
      if (paragraphRuns.length > maxRunsPerParagraph) runsTruncated = true;
    });
    shapes.push({
      shapeId: shape.shapeId,
      name: shape.name,
      hasText: true,
      runs,
      singleRunParagraphs: shape.paragraphRunCounts.map((runCount) => runCount === 1),
    });
  }
  return { ok: true, report: { slideNumber, partPath: slide.partPath, shapes, runsTruncated } };
}
