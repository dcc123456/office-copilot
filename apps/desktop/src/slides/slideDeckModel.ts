/**
 * 幻灯片视图的草稿模型：把"用户在输入框里敲了什么"翻成主进程认的那批 `PptxEditRequest`。
 *
 * 这里决定**什么会被写进文件**，所以单独成模块、单独测（与 `sheet/sheetGridModel.ts` 同一个分工）：
 * 视图只负责显示，落地范围由这两个函数说一不二。
 */
import type { PptxEditRequest, PresentationModel } from "@office-copilot/pptx-io";

/**
 * 一条草稿的定位键。
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
 * 把草稿摊成一批待保存的改动。
 *
 * @param deck 主进程读出来的当前模型（草稿的对照基准）。
 * @param drafts 键见 {@link draftKeyOf}，值是用户改后的文字。
 * @returns 按"页—形状—段"显示顺序排列的改动；多条 run 的段落不进这批（整段覆盖会拍平局部格式，§2.3），
 *   与原文字相同的草稿也不进，于是一次空点"保存"不会写盘。
 */
export function pendingTextEdits(deck: PresentationModel, drafts: Record<string, string>): PptxEditRequest[] {
  const edits: PptxEditRequest[] = [];
  for (const slide of deck.slides) {
    for (const shape of slide.shapes) {
      shape.paragraphTexts.forEach((currentText, paraIndex) => {
        if (shape.paragraphRunCounts[paraIndex] !== 1) return;
        const draftText = drafts[draftKeyOf(slide.partPath, shape.shapeId, paraIndex)];
        if (draftText === undefined || draftText === currentText) return;
        edits.push({
          kind: "text",
          target: { partPath: slide.partPath, shapeId: shape.shapeId, paraIndex, runIndex: 0 },
          text: draftText,
        });
      });
    }
  }
  return edits;
}
