/**
 * 内容流扫描器的文字定位用例（spec 阶段 7 的 W2 第一刀）。
 *
 * 这一刀要的是"这条 run 的字形原点在哪"，因为第二刀靠**原点换没换**来拼逻辑文本行与分栏
 * （真实中文产物一个字一条 `Tj`，同一行内的 run 共享原点、换栏必然换 x）。
 * 所以钉的全是 ISO 32000-1 里容易被"想当然"写错的几条：`T*` 走的是**负** leading（§9.4.4 把它定义成
 * `0 -Tl Td`）、`Td`/`TD` 在**行矩阵**上累加而不是替换当前笔位、`TD` 顺手把 leading 设成 `-ty`、
 * `"` 前两个操作数是词/字间距**而不是 leading**、`cm` 要复合进 CTM、`q`/`Q` 把文字状态与图形状态一起存回、
 * `BT` 只清两个矩阵不清字体与 leading、字号要乘上文字矩阵与 CTM 的缩放（WPS 写 `/FT8 439 Tf` 配 `0.05 … Tm`，
 * 看到的是 22pt）。取字号与原点时不许用 0 顶（内容流写歪时宁可保持上一个值）。
 */
import { describe, expect, it } from "vitest";
import { scanContentStream, type ContentScan } from "./content-stream.js";

/** 取每条 run 的原点，便于整条内容流一起比对。 */
const originsOf = (scan: ContentScan): Array<{ x: number; y: number }> => scan.runs.map((run) => run.origin);

describe("scanContentStream 的字形原点", () => {
  it("从没出现过定位算子时按身份矩阵给 (0,0)——这是真值，不是没测到", () => {
    const scan = scanContentStream("BT /F1 12 Tf (Abc) Tj ET");
    expect(scan.runs).toHaveLength(1);
    expect(scan.runs[0]).toMatchObject({ operator: "Tj", fontName: "F1", fontSizePt: 12, origin: { x: 0, y: 0 } });
  });

  it("Tm 是直接指定：原点取那六个数的 e、f", () => {
    const scan = scanContentStream("BT 1 0 0 1 90 700 Tm (Abc) Tj ET");
    expect(originsOf(scan)).toEqual([{ x: 90, y: 700 }]);
  });

  it("Td 在行矩阵上累加，两次是给绝对偏移而不是替换坐标", () => {
    const scan = scanContentStream("BT 10 20 Td (a) Tj 5 5 Td (b) Tj ET");
    expect(originsOf(scan)).toEqual([
      { x: 10, y: 20 },
      { x: 15, y: 25 },
    ]);
  });

  it("同一行内的连续 Tj 共享原点：隐式推进靠字体宽度，本包不读宽度", () => {
    const scan = scanContentStream("BT 72 720 Td (a) Tj (b) Tj (c) Tj ET");
    expect(originsOf(scan)).toEqual([
      { x: 72, y: 720 },
      { x: 72, y: 720 },
      { x: 72, y: 720 },
    ]);
  });

  it("T* 走的是负 leading：leading 存正数 12，下一行的 y 要减 12", () => {
    const scan = scanContentStream("BT 0 700 Td 12 TL (a) Tj T* (b) Tj T* (c) Tj ET");
    expect(originsOf(scan)).toEqual([
      { x: 0, y: 700 },
      { x: 0, y: 688 },
      { x: 0, y: 676 },
    ]);
  });

  it('TD 把 leading 设成 -ty，之后 T* 继续按同一个行距往下走', () => {
    const scan = scanContentStream("BT 72 700 Td (a) Tj 0 -14 TD (b) Tj T* (c) Tj ET");
    expect(originsOf(scan)).toEqual([
      { x: 72, y: 700 },
      { x: 72, y: 686 },
      { x: 72, y: 672 },
    ]);
  });

  it("单引号算子等价于先 T* 再显示；双引号那两个数是词/字间距，不改 leading", () => {
    const scan = scanContentStream("BT 72 700 Td 16 TL (a) Tj 3 2 (b) \" (c) ' (d) Tj T* (e) Tj ET");
    expect(scan.runs.map((run) => run.operator)).toEqual(["Tj", '"', "'", "Tj", "Tj"]);
    // 双引号那步只换行（700-16），行距仍是 TL 给的 16；后面 T* 也照 16 往下走。
    expect(originsOf(scan)).toEqual([
      { x: 72, y: 700 },
      { x: 72, y: 684 },
      { x: 72, y: 668 },
      { x: 72, y: 668 },
      { x: 72, y: 652 },
    ]);
  });

  it("换栏就换 x：同一列里 y 单调下降，跨列时 x 跳变——这正是第二刀要的判据", () => {
    const scan = scanContentStream(
      "BT 60 760 Td 12 TL (l1) Tj T* (l2) Tj 320 0 Td (r1) Tj T* (r2) Tj ET",
    );
    expect(originsOf(scan)).toEqual([
      { x: 60, y: 760 },
      { x: 60, y: 748 },
      { x: 380, y: 748 },
      { x: 380, y: 736 },
    ]);
  });

  it("cm 复合进 CTM：缩放与平移都要落在原点上", () => {
    const scan = scanContentStream("q 2 0 0 2 0 0 cm BT 10 10 Td (a) Tj ET Q");
    expect(originsOf(scan)).toEqual([{ x: 20, y: 20 }]);
  });

  it("WPS 那种写法：字面字号 439 乘过 Tm 的 0.05 才是生效字号，y 翻转的 cm 也要落在原点上", () => {
    // 取自 tmp/丁长长-高级前端开发2026.pdf 第 1 页标题那一行的原文（页高 841.9）。
    const scan = scanContentStream("q 1 0 0 -1 0 841.9 cm BT /FT8 439 Tf 0.05 0 0 -0.05 253.559 114.89 Tm <1D2C> Tj ET Q");
    expect(scan.runs).toHaveLength(1);
    expect(scan.runs[0]?.origin.x).toBeCloseTo(253.559, 3);
    // 114.89 是翻转空间里距页底的距离，回到 y 朝上的用户空间就是 841.9 - 114.89。
    expect(scan.runs[0]?.origin.y).toBeCloseTo(727.01, 3);
    expect(scan.runs[0]?.fontSizePt).toBeCloseTo(21.95, 3);
  });

  it("q/Q 把图形状态与文字状态一起存回：Q 之后原点与字体都回到压栈那份", () => {
    const scan = scanContentStream("BT /F1 10 Tf 72 720 Td (a) Tj q 1 0 0 1 200 200 cm /F2 20 Tf 5 5 Td (b) Tj Q (c) Tj ET");
    expect(scan.runs.map((run) => [run.fontName, run.fontSizePt])).toEqual([
      ["F1", 10],
      ["F2", 20],
      ["F1", 10],
    ]);
    // 中间那条是 277,925 而不是 205,205：`Td` 是在**行矩阵**上累加（72+5, 720+5），`cm` 只是把结果映射过去。
    expect(originsOf(scan)).toEqual([
      { x: 72, y: 720 },
      { x: 277, y: 925 },
      { x: 72, y: 720 },
    ]);
  });

  it("BT 只清两个矩阵：字体与 leading 跨不过 BT，原点重新从 (0,0) 起算", () => {
    const scan = scanContentStream("BT /F1 9 Tf 72 720 Td 12 TL (a) Tj ET BT (b) Tj 3 -3 Td (c) Tj ET");
    expect(scan.runs.map((run) => run.fontName)).toEqual(["F1", "F1", "F1"]);
    expect(originsOf(scan)).toEqual([
      { x: 72, y: 720 },
      { x: 0, y: 0 },
      { x: 3, y: -3 },
    ]);
  });

  it("TJ 数组里那几个数是字距，不产生新原点：同一条数组的 run 共享原点", () => {
    const scan = scanContentStream("BT 50 60 Td [(a) -250 (b)] TJ ET");
    expect(scan.runs.map((run) => run.operator)).toEqual(["TJ", "TJ"]);
    expect(originsOf(scan)).toEqual([
      { x: 50, y: 60 },
      { x: 50, y: 60 },
    ]);
  });

  it("操作数不够就保持上一个原点，不许拿 0 顶", () => {
    const scan = scanContentStream("BT 72 720 Td (a) Tj 1 2 3 Tm (b) Tj 4 Td (c) Tj ET");
    expect(originsOf(scan)).toEqual([
      { x: 72, y: 720 },
      { x: 72, y: 720 },
      { x: 72, y: 720 },
    ]);
  });

  it("十六进制串与注释、字典混在一起也不抛异常", () => {
    const scan = scanContentStream("% 注释里 (不该被当成串)\nBT <41 42> Tj << /Subtype /Form >> BDC (z) Tj ET");
    expect(scan.runs.map((run) => [run.operand.form, run.origin.x, run.origin.y])).toEqual([
      ["hex", 0, 0],
      ["literal", 0, 0],
    ]);
  });
});
