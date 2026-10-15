/**
 * 幻灯片视图的草稿模型：把"用户在输入框里敲了什么"翻成主进程认的那批 `PptxEditRequest`。
 *
 * 这里决定**什么会被写进文件**，所以单独成模块、单独测（与 `sheet/sheetGridModel.ts` 同一个分工）：
 * 视图只负责显示，落地范围由这两个函数说一不二。
 */
import { notesFlattenNote, type BoxAxisSource, type PptxEditRequest, type PresentationModel, type RunFontFormat, type RunFormat, type ShapeBoxPx, type ShapeGeometryPatch, type SizePtSource, type Slide, type SlideLayerDirection, type SlideNotes, type SlideShape } from "@office-copilot/pptx-io";

/**
 * 形状几何的四个字段，按视图显示的顺序（左、上、宽、高）。
 *
 * 视图的输入框与 {@link inheritedBoxNote} 都要按同一顺序遍历，写两份就会漂移（§3.2）。单位 CSS px。
 */
export const SHAPE_BOX_FIELDS: (keyof ShapeBoxPx)[] = ["leftPx", "topPx", "widthPx", "heightPx"];

/** 几何继承来源层的中文叫法（`pptx-io` 给的是英文层名，文案在本层组装，§5.7）。 */
const BOX_SOURCE_LABEL: Record<BoxAxisSource, string> = {
  shape: "本形状",
  layout: "版式",
  master: "母版",
};

/**
 * 形状没有自己的 `<a:xfrm>` 时，视图上那句几何说明。
 *
 * 数字来自读模型算完继承链的 `effectiveBoxPx`（spec D3 要求"有效值 + 来源层"两个都给，视图不许自己猜），
 * 但**只读**：这一格的写侧入口仍然要求形状自己写了几何（见 {@link pushGeometryEdits}），
 * 把继承值烤进 slide 就永久切断了版式/母版联动。
 *
 * @param shape 一个形状快照（`boxPx` 应为 undefined，即自己没写几何）。
 * @returns 一句中文：继承来的四数与来源层；三层都没写几何或非占位符时说明"没写"，不给猜测值。
 */
export function inheritedBoxNote(shape: SlideShape): string {
  const fields = SHAPE_BOX_FIELDS.filter((field) => shape.effectiveBoxPx[field] !== undefined);
  if (fields.length === 0) {
    return shape.isPlaceholder
      ? "位置与尺寸：版式与母版都没给这一格写几何，本视图不写死（spec D3）。"
      : "位置与尺寸：这个形状自己没写几何，本视图不写死（spec D3）。";
  }
  const numbers = fields.map((field) => `${SHAPE_BOX_FIELD_LABEL[field]} ${shape.effectiveBoxPx[field]}`).join(" · ");
  const layers = [...new Set(fields.map((field) => shape.boxPxSources[field]).filter((source) => source !== undefined).map((source) => BOX_SOURCE_LABEL[source]))].join("、");
  return `位置与尺寸继承自${layers}：${numbers} px（写死就切断版式联动，本视图只读，spec D3）。`;
}

/**
 * 字号继承来源层的中文叫法（`pptx-io` 给的是英文层名，文案在本层组装，§5.7）。
 * 六档与 `RunFormat.effectiveSizeSource` 一一对应，少写一档 TS 会报错——这正是我们要的兜底。
 */
const SIZE_SOURCE_LABEL: Record<SizePtSource, string> = {
  run: "这条 run 自己写的",
  paragraph: "这一段自己的段落默认字号",
  "shape-list": "这个形状的文本样式",
  layout: "版式那一格",
  "master-list": "母版那一格",
  "master-text-styles": "母版的文本样式",
};

/**
 * 字号输入框上的 hover 说明：显式值与继承值说的是两件事。
 *
 * 输入框的 **value 只放显式值**（`RunFormat.sizePt`）：把继承来的 40pt 填进框里，用户一个字没敲也会产出一条
 * `sz="4000"`，从此改母版字号不再联动——spec D3 那条红线。所以继承值只出现在这句说明里。
 *
 * @param format 这一格第一条 run 的读模型格式。
 * @returns 中文一句：自己写了就说"当前 Xpt"；没写但链上算得出就说"当前 Xpt，继承自…"；整条链都没有才承认"未写"。
 */
export function sizeFieldHint(format: RunFormat): string {
  if (format.sizePt !== undefined) return `当前 ${format.sizePt}pt（${SIZE_SOURCE_LABEL.run}）`;
  const { effectiveSizePt, effectiveSizeSource } = format;
  if (effectiveSizePt !== undefined && effectiveSizeSource !== undefined) {
    return `当前 ${effectiveSizePt}pt，继承自${SIZE_SOURCE_LABEL[effectiveSizeSource]}（本形状没写字号，写死就切断母版联动）`;
  }
  return "这一格没写字号，继承链上也没找到（视图不写死，spec D3）";
}

/**
 * 形状几何四个字段的中文标签（单位是 CSS px，非 pt、非 EMU）。
 *
 * 视图的输入框与 AI 提案卡都要说这四个字段，放这里是因为落地范围也由本模块定（§3.2 不写第二份）。
 */
export const SHAPE_BOX_FIELD_LABEL: Record<keyof ShapeBoxPx, string> = {
  leftPx: "左",
  topPx: "上",
  widthPx: "宽",
  heightPx: "高",
};

/**
 * 叠放方向四个值的中文标签（照 PowerPoint 功能区那四个按钮的叫法）。
 *
 * AI 提案卡的每一行与视图上的层级按钮说的是同一件事，所以只写一份（§3.2、§5.7）。
 */
export const LAYER_DIRECTION_LABEL: Record<SlideLayerDirection, string> = {
  up: "上移一层",
  down: "下移一层",
  top: "置于顶层",
  bottom: "置于底层",
};

/** 一处未保存的草稿：只带用户真改过的那几项（没改的字段留 undefined，不交出、也不覆盖）。 */
export interface SlideDraft {
  /** 段落文字（整段只有一条 run 时才允许改，见 {@link pendingEdits}）。 */
  text?: string;
  /** 字号（CSS pt，非半磅）。 */
  sizePt?: number;
  /** 文字颜色，`#rrggbb`（§2.2）。 */
  hex?: string;
  /** 形状的位置/尺寸（CSS px），只含改过的那几项。 */
  box?: ShapeGeometryPatch;
}

/**
 * 一条 run 级草稿（文字 / 字号 / 颜色）的定位键。
 *
 * @param partPath 幻灯片部件路径，如 `ppt/slides/slide1.xml`。
 * @param shapeId 形状 `<p:cNvPr @id>`。
 * @param paraIndex 段落序号（0 基）。
 * @returns 拼出来的键；run 固定是 0，因为只有"整段只有一条 run"时本视图才允许改。
 */
export function draftKeyOf(partPath: string, shapeId: number, paraIndex: number): string {
  return `${partPath}#${shapeId}#${paraIndex}`;
}

/**
 * 一条形状几何草稿的定位键。
 *
 * @param partPath 幻灯片部件路径。
 * @param shapeId 形状 `<p:cNvPr @id>`。
 * @returns 拼出来的键；带 `@` 前缀，与 run 级那条键不可能撞车。
 */
export function boxDraftKeyOf(partPath: string, shapeId: number): string {
  return `${partPath}@${shapeId}`;
}

/**
 * 删掉一个形状时把它那些还没保存的草稿一并丢掉。
 *
 * 形状号是"这一页已有 `@id` 最大值 +1"（见 `pptx-io/edit.ts` 的 `addTextBox`），删掉最大号那个之后
 * 再新建就会把同一个号发给新形状；旧草稿留着就会把用户给旧框打的字静默写进新框。
 *
 * @param drafts 当前全部草稿，键见 {@link draftKeyOf} 与 {@link boxDraftKeyOf}。
 * @param partPath 被删形状所在页的部件路径。
 * @param shapeId 被删形状的形状号。
 * @returns 去掉该形状那几条（该形状各段 + 它的几何那一条）之后的新对象；没命中时返回原对象。
 */
export function dropShapeDrafts(drafts: Record<string, SlideDraft>, partPath: string, shapeId: number): Record<string, SlideDraft> {
  const runPrefix = `${partPath}#${shapeId}#`;
  const boxKey = boxDraftKeyOf(partPath, shapeId);
  const kept = Object.fromEntries(Object.entries(drafts).filter(([key]) => !key.startsWith(runPrefix) && key !== boxKey));
  return Object.keys(kept).length === Object.keys(drafts).length ? drafts : kept;
}

/**
 * 叠放方向是不是已经到头（AI 提案与视图上的层级按钮共用这一刀，§3.2 不写第二份）。
 *
 * `shapes` 就是 `<p:spTree>` 里形状的文档顺序（读侧按它列），而那个顺序就是叠放次序：
 * 第一个在最底、最后一个在最顶（ECMA-376 §19.3.5 只钉子元素顺序，这层语义取自 python-pptx
 * 的 `shapetree.py`）。所以只要比对首尾，不必再解析一次 XML。
 *
 * @param shapes 该页全部形状（文档顺序）。
 * @param shape 本次要挪的那个。
 * @param direction 要挪的方向。
 * @returns 一句中文理由（该页只有一个形状、或已在要去的那一端）；能挪时返回 null。
 */
export function layerEdgeNote(
  shapes: SlideShape[],
  shape: SlideShape,
  direction: SlideLayerDirection,
): string | null {
  const at = shapes.indexOf(shape);
  if (shapes.length < 2) return "这一页只有这一个形状，没有可叠放的邻居";
  if ((direction === "up" || direction === "top") && at === shapes.length - 1) return "它已经在最顶层";
  if ((direction === "down" || direction === "bottom") && at === 0) return "它已经在最底层";
  return null;
}

/**
 * 把草稿摊成一批待保存的改动。
 *
 * @param deck 主进程读出来的当前模型（草稿的对照基准）。
 * @param drafts 键见 {@link draftKeyOf} 与 {@link boxDraftKeyOf}。
 * @returns 按「页—形状（几何先，再各段）」显示顺序排列的改动。
 *   三条守卫：多条 run 的段落一次都不交出（整段覆盖会拍平局部格式，`AGENTS.md` §2.3）；
 *   与现场值相同的字段不交出（空点"保存"不写盘）；形状没有自己的 `<a:xfrm>` 时几何不交出
 *   （几何来自版式继承，写死它就是切断联动，spec D3）。
 */
export function pendingEdits(deck: PresentationModel, drafts: Record<string, SlideDraft>): PptxEditRequest[] {
  const edits: PptxEditRequest[] = [];
  for (const slide of deck.slides) {
    for (const shape of slide.shapes) {
      pushGeometryEdits(edits, slide.partPath, shape, drafts[boxDraftKeyOf(slide.partPath, shape.shapeId)]);
      shape.paragraphTexts.forEach((currentText, paraIndex) => {
        if (shape.paragraphRunCounts[paraIndex] !== 1) return;
        const runTarget = { partPath: slide.partPath, shapeId: shape.shapeId, paraIndex, runIndex: 0 };
        const draft = drafts[draftKeyOf(slide.partPath, shape.shapeId, paraIndex)];
        if (!draft) return;
        const currentFormat = shape.paragraphRunFormats[paraIndex]?.[0] ?? {};
        if (draft.text !== undefined && draft.text !== currentText) {
          edits.push({ kind: "text", target: runTarget, text: draft.text });
        }
        const format: RunFontFormat = {};
        if (draft.sizePt !== undefined && draft.sizePt !== currentFormat.sizePt) format.sizePt = draft.sizePt;
        if (Object.keys(format).length > 0) edits.push({ kind: "font", target: runTarget, format });
        if (draft.hex !== undefined && draft.hex !== currentFormat.hex) {
          edits.push({ kind: "color", target: runTarget, hex: draft.hex });
        }
      });
    }
  }
  return edits;
}

/**
 * 把一个形状草稿里的几何改动摊成一条 `geometry`（只带真改过的那几项）。
 *
 * @param edits 目标数组，就地追加。
 * @param partPath 幻灯片部件路径。
 * @param shape 该形状的当前快照。
 * @param draft 该形状的草稿；没有草稿或没有自己的 `<a:xfrm>` 时什么都不加。
 */
function pushGeometryEdits(edits: PptxEditRequest[], partPath: string, shape: SlideShape, draft: SlideDraft | undefined): void {
  const box = draft?.box;
  const current = shape.boxPx;
  if (!box || !current) return;
  const patch = changedBoxFields(box, current);
  if (Object.keys(patch).length === 0) return;
  edits.push({ kind: "geometry", target: { partPath, shapeId: shape.shapeId }, box: patch });
}

/**
 * 逐字段比草稿与现场值，只留下改过的那几项。
 *
 * @param box 草稿里的几何（全部可选）。
 * @param current 形状自己写的 `<a:xfrm>`（CSS px）。
 * @returns 只含差异字段的补丁；没有差异时是空对象。
 */
function changedBoxFields(box: ShapeGeometryPatch, current: ShapeBoxPx): ShapeGeometryPatch {
  const patch: ShapeGeometryPatch = {};
  if (box.leftPx !== undefined && box.leftPx !== current.leftPx) patch.leftPx = box.leftPx;
  if (box.topPx !== undefined && box.topPx !== current.topPx) patch.topPx = box.topPx;
  if (box.widthPx !== undefined && box.widthPx !== current.widthPx) patch.widthPx = box.widthPx;
  if (box.heightPx !== undefined && box.heightPx !== current.heightPx) patch.heightPx = box.heightPx;
  return patch;
}

/**
 * 那一页的备注能不能存，不能存时说为什么。
 *
 * 备注住在**另一个部件**里（`ppt/notesSlides/notesSlideN.xml`），落地是整篇换掉那一格，所以这里按三段判：
 * ① 这一页没有备注部件——新建 `notesSlide` 那一档还没做（要 §2.10 三处登记），不许凭空造一份；
 * ② 输入框与现场逐字一样——没有要保存的东西，按钮就该是灰的；
 * ③ 现有备注带多截格式——整篇覆盖会拍平那些格式（§2.3），这条判定与引擎共用 `notesFlattenNote` 那一份。
 *
 * @param slide 该页的读模型快照。
 * @param draftText 输入框里的整段备注，一行一段（`\n` 分段，与 `paragraphTexts` 一一对应）。
 * @returns 一句中文理由；能存时返回 null。
 */
export function notesSaveBlock(slide: Slide, draftText: string): string | null {
  if (!slide.notes) return "这一页没有备注页（新建备注页这一档还没做，只能改已有的备注）。";
  if (draftText === slide.notes.paragraphTexts.join("\n")) return "备注没有改动。";
  return notesFlattenNote(slide.notes.paragraphRunCounts, slide.notes.styledRunCount);
}

/**
 * 把输入框那句备注翻成主进程认的那条改动。
 *
 * 只在 {@link notesSaveBlock} 返回 null 时调用：备注部件路径直接用地读模型给的那份，视图不算路径也不换单位（§2.1）。
 *
 * @param notes 该页的备注快照（`notesSaveBlock` 已经确认这一页有备注部件）。
 * @param draftText 输入框里的整段备注，一行一段；空行就是空串那一段。
 * @returns 一条 `set_notes` 请求，确认后原样送 `ppt:edit` 即可落地并拿到逆操作。
 */
export function notesRequestOf(notes: SlideNotes, draftText: string): PptxEditRequest {
  return { kind: "set_notes", target: { notesPartPath: notes.partPath }, paragraphTexts: draftText.split("\n") };
}
