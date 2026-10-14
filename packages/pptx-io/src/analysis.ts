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
import { type PresentationModel, type RunFormat, type Slide, type SlideShape } from "./read.js";
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
    /** 大类：形状 / 图片 / 图形框 / 组合。 */
    kind: SlideShape["kind"];
    /** 占位符角色（`type` 或 `idx`），非占位符为 null。 */
    placeholderRef: string | null;
    /** 有没有自己的 `<a:xfrm>`（false＝几何继承自版式，改几何会被拒，D3）。 */
    hasOwnGeometry: boolean;
    /** 这一页上这个形状的文字概要（拼接串，超长已截）。 */
    text: string;
  }>;
  /** 形状清单是否被 `maxShapesPerSlide` 截断。 */
  shapesTruncated: boolean;
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
 * 截断一条上送文字。
 *
 * @param value 现场读出的文字。
 * @returns `{ text, truncated }`：不超上限时原样，超了就切前 {@link RUN_TEXT_LIMIT} 字并置标记。
 */
function clip(value: string): { text: string; truncated: boolean } {
  return value.length <= RUN_TEXT_LIMIT ? { text: value, truncated: false } : { text: `${value.slice(0, RUN_TEXT_LIMIT)}…`, truncated: true };
}

/**
 * 把形状的文字概要拼成一条截断串（描述用，不含逐 run 定位）。
 *
 * @param shape 形状快照。
 * @returns 各段用 `⏎` 隔开的概要文字。
 */
function shapeTextSummary(shape: SlideShape): string {
  return clip(shape.paragraphTexts.join("⏎")).text;
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
      text: shapeTextSummary(shape),
    })),
    shapesTruncated: slide.shapes.length > listed.length,
  };
}

/**
 * 把某一页每个带文字的形状摊成「段 → run」清单，供模型对准 `paraIndex` / `runIndex`。
 *
 * @param model 主进程读出来的模型。
 * @param slideNumber 页号（1 基；越界时返回 `ok:false` 与给模型看的中文理由，不抛异常）。
 * @param maxRunsPerParagraph 每段最多列多少条 run（超出置 `runsTruncated`）。
 * @returns 该页的文字报告，或一句中文失败理由（页号越界）。
 */
export function readSlideText(
  model: PresentationModel,
  slideNumber: number,
  maxRunsPerParagraph: number,
): { ok: true; report: SlideTextReport } | { ok: false; reason: string } {
  const slide = model.slides.find((entry) => entry.order + 1 === slideNumber);
  if (!slide) return { ok: false, reason: `没有第 ${slideNumber} 页：这份演示文稿一共 ${model.slides.length} 页，先用 ppt_describe 看页号。` };
  const shapes: SlideShapeTextSnapshot[] = [];
  let runsTruncated = false;
  for (const shape of slide.shapes) {
    if (shape.paragraphTexts.length === 0) continue; // 图片/图形框没有文本体：列出来只会误导模型去改它
    const runs: SlideRunSnapshot[] = [];
    shape.paragraphTexts.forEach((_, paraIndex) => {
      const paragraphRuns = shape.paragraphRunTexts[paraIndex] ?? [];
      paragraphRuns.slice(0, maxRunsPerParagraph).forEach((runText, runIndex) => {
        const clipped = clip(runText);
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
