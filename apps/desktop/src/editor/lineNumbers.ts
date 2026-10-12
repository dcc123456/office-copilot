import type { SectionLineNumbers } from "@office-copilot/docx-io";

/**
 * 一个顶层块在画布上的量测结果。单位一律 CSS px（AGENTS §2.1：换算只在 docx-io 进出口）。
 */
export interface NumberedBlock {
  /** 块顶在 `.page-stage` 坐标系里的纵向偏移（屏幕差值已除掉缩放），px。 */
  topPx: number;
  /** 块高，px（含块内的行，不含块外边距）。跨页的段落这里是整个块的高。 */
  heightPx: number;
  /** 该块的行高，px（computed `line-height`）。 */
  lineHeightPx: number;
}

/** 一页正文带的纵向范围，`.page-stage` 坐标系，px。 */
export interface PageBand {
  topPx: number;
  bottomPx: number;
}

/** 画布上一个行号标签：挂在第几页、那一页正文区的哪个偏移、对齐多高的行。 */
export interface LineNumberLabel {
  /** 显示的数字。Word 的口径是"这一行本身的行号"，不是第几个号。 */
  value: number;
  /** 页序号，从 1 起；调用方用它做 key，位置本身按 `topPx` 直接摆。 */
  page: number;
  /** 号所在行的顶，与 `NumberedBlock.topPx` 同一坐标系（`.page-stage`），px。 */
  topPx: number;
  /** 那一行的行高，用来把号垂直对齐。 */
  lineHeightPx: number;
}

/**
 * 按 Word 的语义算出画布上该画哪些行号。
 *
 * `w:lnNumType` 的规范缺省是"每页重新开始、起始值为 `w:start`"（ISO/IEC 29500-1 §17.6.9
 * 的例子原话：each line, restarting at one at the top of each new page），而 `w:restart`
 * 我们没有建模，所以照缺省来：每页从 `start` 重新数。`w:countBy` 是"每几行显一个号"，
 * 显示的值仍是行本身的号，所以每 countBy 行出现一次。
 *
 * 分页归属按**行自己的位置**算，不按块：段中分页时，后半截的行落在下一页，
 * 它们就该从该页的 `start` 重新数（若按块的首行算，上一段续行会把大号码写进下一页顶部）。
 *
 * @param blocks 文档顺序的顶层块量测结果。
 * @param bands  每一页正文带的范围，按页序；落在所有带之外的行不编号。
 * @param ln     节里的 `w:lnNumType`；`countBy` ≤ 0 等于没开行号。
 * @returns 标签数组；没开行号或还没量到布局时返回空数组。
 */
export function lineNumberLabelsFor(
  blocks: NumberedBlock[],
  bands: PageBand[],
  ln: SectionLineNumbers,
): LineNumberLabel[] {
  const countBy = Math.floor(ln.countBy);
  if (!(countBy >= 1)) return [];
  const start = Math.max(1, Math.floor(ln.start) || 1);
  // 先把块摊成视觉行：行高均匀时按高度取整，至少算一行（空段落 Word 也占一行）。
  const lines: { topPx: number; lineHeightPx: number }[] = [];
  for (const block of blocks) {
    const lineHeightPx = Math.max(block.lineHeightPx, 1);
    const lineCount = Math.max(1, Math.round(block.heightPx / lineHeightPx));
    for (let i = 0; i < lineCount; i++) {
      lines.push({ topPx: block.topPx + i * lineHeightPx, lineHeightPx });
    }
  }
  const labels: LineNumberLabel[] = [];
  bands.forEach((band, index) => {
    let lineInPage = 0;
    for (const line of lines) {
      const center = line.topPx + line.lineHeightPx / 2;
      if (center < band.topPx || center >= band.bottomPx) continue;
      const value = start + lineInPage;
      lineInPage += 1;
      if ((value - start) % countBy === 0) {
        labels.push({
          value,
          page: index + 1,
          topPx: line.topPx,
          lineHeightPx: line.lineHeightPx,
        });
      }
    }
  });
  return labels;
}
