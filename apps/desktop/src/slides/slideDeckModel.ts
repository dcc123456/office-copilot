/**
 * 幻灯片视图的草稿模型：把"用户在输入框里敲了什么"翻成主进程认的那批 `PptxEditRequest`。
 *
 * 这里决定**什么会被写进文件**，所以单独成模块、单独测（与 `sheet/sheetGridModel.ts` 同一个分工）：
 * 视图只负责显示，落地范围由这两个函数说一不二。
 */
import type { PptxEditRequest, PresentationModel, RunFontFormat, ShapeBoxPx, ShapeGeometryPatch, SlideShape } from "@office-copilot/pptx-io";

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
