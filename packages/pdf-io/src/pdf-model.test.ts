/**
 * 逻辑文本行拼接的用例（spec 阶段 7 的 W2 前置）。
 *
 * 钉的是"两份真文件实测出来的两类生产者形态都不能拼错"：
 * ① Word / Google 一行一条 `Tj`（同一行多条 run 时共享同一个原点）；
 * ② WPS 一个字一条 `Tj` + `TD`，**同一行的 y 相同**（实测抖动 0.15pt）、换栏时 y 相同而 x 跳 94…157pt。
 * 所以判据是「y 容差 + 按字号归一的 x 间隔」：行内推进实测 p50=1.00em / p90=1.01em，断栏 ≥9em，
 * 阈值取 4em 落在两类中间。断言里最要紧的一条是**不许串栏**——两栏拼成一句比切成两句更难发现，
 * 因为它会把右栏的文字塞进左栏那一句里，模型照着改就会改到别人家的那条 run。
 */
import { describe, expect, it } from "vitest";
import { pdfTextLinesOf, type PdfTextRun } from "./pdf-model.js";

/** 造一条 run：只填拼行用得上的字段，其余按可读的默认值。 */
function run(overrides: Partial<PdfTextRun> = {}): PdfTextRun {
  return {
    pageIndex: 0,
    streamRef: "6 0 R",
    withinForm: null,
    operator: "Tj",
    form: "hex",
    start: 0,
    encodedBytes: 2,
    tokenLength: 6,
    baseFont: "STSong",
    composite: true,
    fontSizePt: 10.45,
    origin: { x: 90, y: 615 },
    text: "字",
    ...overrides,
  };
}

/** 逐字形排一行：从 x 起、每字一个字号宽（就是 WPS 那种 `TD` 推进）。 */
function glyphLine(text: string, x: number, y: number, size = 10.45): PdfTextRun[] {
  return [...text].map((char, index) => run({ text: char, fontSizePt: size, start: index, origin: { x: x + index * size, y } }));
}

describe("pdfTextLinesOf", () => {
  it("WPS 那种逐字形的 run 拼成一行", () => {
    const lines = pdfTextLinesOf(glyphLine("姓名：丁长长", 90, 615));
    expect(lines).toHaveLength(1);
    expect(lines[0]).toMatchObject({ text: "姓名：丁长长", x: 90, y: 615, fromRun: 0, toRun: 6, hasUndecoded: false });
  });

  it("同一行里 y 的小抖动算同一行（实测 0.15pt），差一个行距算换行", () => {
    const lines = pdfTextLinesOf([
      run({ text: "甲", origin: { x: 90, y: 615 } }),
      run({ text: "乙", origin: { x: 100.45, y: 614.85 } }),
      run({ text: "丙", origin: { x: 90, y: 604.4 } }),
    ]);
    expect(lines.map((line) => line.text)).toEqual(["甲乙", "丙"]);
  });

  it("换栏不许串：右栏与左栏同一个 y，但 x 空出 9em，必须另起一行", () => {
    const lines = pdfTextLinesOf([
      ...glyphLine("姓名：丁长长", 90, 615),
      ...glyphLine("出生日期：2026", 289.2, 615),
    ]);
    expect(lines.map((line) => line.text)).toEqual(["姓名：丁长长", "出生日期：2026"]);
    expect(lines[1]).toMatchObject({ x: 289.2, y: 615, fromRun: 6 });
  });

  it("间隔阈值按字号归一：大字号标题里 22pt 的推进是行内，小字号正文里 30pt 也不算断行、94pt 才算", () => {
    const title = pdfTextLinesOf(glyphLine("个人简历", 253.5, 727, 22));
    expect(title.map((line) => line.text)).toEqual(["个人简历"]);
    const smallGap = pdfTextLinesOf([run({ text: "甲", origin: { x: 90, y: 615 } }), run({ text: "乙", origin: { x: 120, y: 615 } })]);
    expect(smallGap.map((line) => line.text)).toEqual(["甲乙"]);
    const bigGap = pdfTextLinesOf([run({ text: "甲", origin: { x: 90, y: 615 } }), run({ text: "乙", origin: { x: 184, y: 615 } })]);
    expect(bigGap.map((line) => line.text)).toEqual(["甲", "乙"]);
  });

  it("Word 形态：每行一条 run，且同一行里共享原点的多条 run 要合在一起", () => {
    const lines = pdfTextLinesOf([
      run({ text: "SUMMARY", start: 0, origin: { x: 40, y: 655 }, fontSizePt: 12 }),
      run({ text: "A short ", start: 4, origin: { x: 40, y: 635 }, fontSizePt: 10 }),
      run({ text: "professional summary", start: 10, origin: { x: 40, y: 635 }, fontSizePt: 10 }),
      run({ text: "SKILLS", start: 20, origin: { x: 40, y: 440 }, fontSizePt: 12 }),
    ]);
    expect(lines.map((line) => line.text)).toEqual(["SUMMARY", "A short professional summary", "SKILLS"]);
    expect(lines[1]?.toRun).toBe(3);
  });

  it("x 倒退就另起一行：同一行里倒着排的那条不许接进来（它属于另一个块）", () => {
    const lines = pdfTextLinesOf([
      run({ text: "甲", origin: { x: 90, y: 615 } }),
      run({ text: "乙", origin: { x: 100.45, y: 615 } }),
      run({ text: "丙", origin: { x: 90, y: 615 } }),
    ]);
    expect(lines.map((line) => line.text)).toEqual(["甲乙", "丙"]);
  });

  it("跨页 / 跨内容流 / 跨 Form 链的一律不合并：三者的原点不在同一个坐标空间", () => {
    const lines = pdfTextLinesOf([
      run({ text: "甲", pageIndex: 0 }),
      run({ text: "乙", pageIndex: 1 }),
      run({ text: "丙", pageIndex: 1, streamRef: "9 0 R" }),
      run({ text: "丁", pageIndex: 1, streamRef: "9 0 R", withinForm: "Fm0" }),
    ]);
    expect(lines).toHaveLength(4);
    expect(lines.map((line) => [line.pageIndex, line.withinForm])).toEqual([
      [0, null],
      [1, null],
      [1, null],
      [1, "Fm0"],
    ]);
  });

  it("行内有解不出原文的 run 时如实标 hasUndecoded，text 就是缺了字的", () => {
    const lines = pdfTextLinesOf([
      run({ text: "甲", origin: { x: 90, y: 615 } }),
      run({ text: null, origin: { x: 100.45, y: 615 } }),
      run({ text: "丙", origin: { x: 110.9, y: 615 } }),
    ]);
    expect(lines[0]).toMatchObject({ text: "甲丙", hasUndecoded: true });
  });

  it("不变量：所有行连续覆盖入参的每一条 run，一条都不丢也不重复", () => {
    const runs = [...glyphLine("个人简历", 253.5, 727, 22), ...glyphLine("一、基本信息", 90, 649.6, 16)];
    const lines = pdfTextLinesOf(runs);
    expect(lines.at(0)?.fromRun).toBe(0);
    expect(lines.at(-1)?.toRun).toBe(runs.length);
    for (let index = 1; index < lines.length; index += 1) expect(lines[index]?.fromRun).toBe(lines[index - 1]?.toRun);
  });

  it("空输入回空数组", () => {
    expect(pdfTextLinesOf([])).toEqual([]);
  });
});
