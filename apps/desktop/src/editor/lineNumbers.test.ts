import { describe, expect, it } from "vitest";
import {
  lineNumberLabelsFor,
  type NumberedBlock,
  type PageBand,
} from "./lineNumbers.js";

/**
 * 造一个块的量测结果。
 *
 * @param topPx 块顶（`.page-stage` 坐标系，px）。
 * @param lines 该块占几视觉行，换算成 `heightPx = lines * lineHeightPx`。
 * @param lineHeightPx 行高，px。
 */
const block = (
  topPx: number,
  lines: number,
  lineHeightPx = 20,
): NumberedBlock => ({
  topPx,
  heightPx: lines * lineHeightPx,
  lineHeightPx,
});

/** 造一页正文带。 @param topPx/bottomPx 带的上下边界，px。 */
const band = (topPx: number, bottomPx: number): PageBand => ({ topPx, bottomPx });

describe("画布行号", () => {
  it("每行一个号，号跟着自己的行走", () => {
    const labels = lineNumberLabelsFor(
      [block(100, 3)],
      [band(0, 1000)],
      { countBy: 1, start: 1, distance: 24 },
    );
    expect(labels.map((l) => l.value)).toEqual([1, 2, 3]);
    expect(labels.map((l) => l.topPx)).toEqual([100, 120, 140]);
    expect(labels[2]!.lineHeightPx).toBe(20);
  });

  it("每几行一个号时，显示的还是这一行本身的号", () => {
    // Word 的 `w:countBy="2"` + `w:start="7"` 显示 7、9、11……而不是 2、4、6。
    const labels = lineNumberLabelsFor(
      [block(0, 5)],
      [band(0, 1000)],
      { countBy: 2, start: 7, distance: 24 },
    );
    expect(labels.map((l) => l.value)).toEqual([7, 9, 11]);
  });

  it("每页从起始号重新开始（w:restart 的规范缺省）", () => {
    const labels = lineNumberLabelsFor(
      [block(0, 2), block(50, 2)],
      [band(0, 50), band(50, 100)],
      { countBy: 1, start: 3, distance: 24 },
    );
    expect(labels.map((l) => `${l.page}:${l.value}`)).toEqual([
      "1:3",
      "1:4",
      "2:3",
      "2:4",
    ]);
  });

  it("段中分页时按行归页：续到下一页的行从该页起始号重新数", () => {
    // 一个 4 行的段落横跨两页。按"块所在页"编号会把上一页的大号码（3、4）画到
    // 下一页顶部，而 Word 是每页重新数，所以后两行应该是第二页的 1、2。
    const labels = lineNumberLabelsFor(
      [block(0, 4)],
      [band(0, 40), band(40, 100)],
      { countBy: 1, start: 1, distance: 24 },
    );
    expect(labels.map((l) => `${l.page}:${l.value}@${l.topPx}`)).toEqual([
      "1:1@0",
      "1:2@20",
      "2:1@40",
      "2:2@60",
    ]);
  });

  it("空段落算一行，零头按行高取整", () => {
    const labels = lineNumberLabelsFor(
      [block(0, 0), block(20, 0)],
      [band(0, 100)],
      { countBy: 1, start: 1, distance: 24 },
    );
    expect(labels.map((l) => `${l.topPx}:${l.value}`)).toEqual(["0:1", "20:2"]);
    // 55px / 20px 行高 = 2.75 → 3 行。
    expect(
      lineNumberLabelsFor(
        [{ topPx: 0, heightPx: 55, lineHeightPx: 20 }],
        [band(0, 100)],
        { countBy: 1, start: 1, distance: 24 },
      ),
    ).toHaveLength(3);
  });

  it("countBy 不是正数就不标行号，分页带还没量出来时一页都不画", () => {
    expect(
      lineNumberLabelsFor(
        [block(0, 3)],
        [band(0, 100)],
        { countBy: 0, start: 1, distance: 24 },
      ),
    ).toEqual([]);
    expect(
      lineNumberLabelsFor(
        [block(0, 3)],
        [],
        { countBy: 1, start: 1, distance: 24 },
      ),
    ).toEqual([]);
    // 落在所有带之外的行（还没排版到位）不画号。
    expect(
      lineNumberLabelsFor(
        [block(500, 3)],
        [band(0, 100)],
        { countBy: 1, start: 1, distance: 24 },
      ),
    ).toEqual([]);
  });
});
