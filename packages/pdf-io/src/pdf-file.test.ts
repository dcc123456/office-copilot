/**
 * `.pdf` 读模型与 W1 档改写的用例（spec 阶段 7 的第一档，判据取自 D8）。
 *
 * 钉的是三条红线：
 * ① 两种串形态（literal / hex）都要认得出来，否则同一份文件"看得见却改不动"；
 * ② 改完必须**装得回原字节数**，装不回去就整批不写（§2.4），不许悄悄把内容流撑开；
 * ③ 落地走增量更新：原文件的前缀字节一字不动，只追加 delta（§7.5.6）。
 */
import { describe, expect, it } from "vitest";
import {
  applyPdfTextEdits,
  isRunEditable,
  pdfTargetKey,
  readPdfModel,
  revertPdfTextEdits,
  runEditBlock,
  targetOfRun,
  type PdfTextRun,
  type PdfTextTarget,
} from "./pdf-file.js";
import { simpleWinAnsiFixturePdf, cidIdentityHFixturePdf, cidNoToUnicodeFixturePdf, flateWinAnsiFixturePdf } from "./test-pdf.js";

/** 按解出来的文字找定位键。 */
function targetOfText(runs: PdfTextRun[], text: string): PdfTextTarget {
  const run = runs.find((candidate) => candidate.text === text);
  const target = run ? targetOfRun(run) : null;
  if (!target) throw new Error(`夹具里没有可编辑的「${text}」`);
  return target;
}

describe("readPdfModel：两种串形态与 Form XObject 递归", () => {
  it("第一页的 literal 串与第二页的 hex 串都读出来，并解成文字", async () => {
    const model = await readPdfModel(await simpleWinAnsiFixturePdf());
    expect(model.pageCount).toBe(2);
    expect(model.runs.map((run) => run.text)).toEqual(["Quarterly Report 2026", "Inner Form Text", "Appendix A stays put"]);
    expect(model.runs.filter((run) => run.form === "literal")).toHaveLength(2);
    expect(model.runs.filter((run) => run.form === "hex")).toHaveLength(1);
  });

  it("Form XObject 里那段字带得出资源名链，BaseFont 认得 Helvetica", async () => {
    const model = await readPdfModel(await simpleWinAnsiFixturePdf());
    const inner = model.runs.find((run) => run.text === "Inner Form Text");
    expect(inner?.withinForm).toBe("Fr1");
    expect(inner?.baseFont).toBe("Helvetica");
    expect(inner?.composite).toBe(false);
  });

  it("CID/Identity-H 加 /ToUnicode 时按两个字节一个码解出中文", async () => {
    const model = await readPdfModel(await cidIdentityHFixturePdf());
    const only = model.runs[0];
    expect(only?.text).toBe("中文标题");
    expect(only?.composite).toBe(true);
    expect(only?.baseFont).toBe("SourceHanSansCN-Regular");
  });

  it("内容流是 /FlateDecode（zlib 框）时两种串形态照样读出来", async () => {
    const model = await readPdfModel(await flateWinAnsiFixturePdf());
    expect(model.runs.map((run) => run.text)).toEqual(["Quarterly Report 2026", "Flate Hex Text"]);
    expect(model.runs.every((run) => run.pageIndex === 0)).toBe(true);
  });
});

describe("applyPdfTextEdits：同长度替换 + 增量更新", () => {
  it("literal 串改成同长度的新字，前缀字节逐字不变，只追加 delta", async () => {
    const bytes = await simpleWinAnsiFixturePdf();
    const model = await readPdfModel(bytes);
    const result = await applyPdfTextEdits(bytes, [{ target: targetOfText(model.runs, "Quarterly Report 2026"), text: "Quarterly Budget 2026" }]);
    if (!result.ok) throw new Error("该改成功的：" + JSON.stringify(result.failures));
    const prefixSame = Buffer.compare(Buffer.from(bytes.subarray(0, bytes.length)), Buffer.from(result.bytes.subarray(0, bytes.length))) === 0;
    expect(prefixSame).toBe(true);
    expect(result.bytes.length).toBeGreaterThan(bytes.length);
    expect(result.deltaBytes).toBeLessThan(4096);
    const after = await readPdfModel(result.bytes);
    expect(after.runs.map((run) => run.text)).toEqual(["Quarterly Budget 2026", "Inner Form Text", "Appendix A stays put"]);
  });

  it("hex 串那条也改得动，另一页一个字都不动", async () => {
    const bytes = await simpleWinAnsiFixturePdf();
    const model = await readPdfModel(bytes);
    const result = await applyPdfTextEdits(bytes, [{ target: targetOfText(model.runs, "Appendix A stays put"), text: "Appendix A stays hot" }]);
    if (!result.ok) throw new Error("该改成功的：" + JSON.stringify(result.failures));
    const after = await readPdfModel(result.bytes);
    expect(after.runs.map((run) => run.text)).toEqual(["Quarterly Report 2026", "Inner Form Text", "Appendix A stays hot"]);
  });

  it("压缩内容流里改一条，重新压回去仍是 /FlateDecode，另一条不动", async () => {
    const bytes = await flateWinAnsiFixturePdf();
    const model = await readPdfModel(bytes);
    const result = await applyPdfTextEdits(bytes, [{ target: targetOfText(model.runs, "Quarterly Report 2026"), text: "Quarterly Budget 2026" }]);
    if (!result.ok) throw new Error("该改成功的：" + JSON.stringify(result.failures));
    expect(Buffer.compare(Buffer.from(bytes.subarray(0, bytes.length)), Buffer.from(result.bytes.subarray(0, bytes.length))) === 0).toBe(true);
    const after = await readPdfModel(result.bytes);
    expect(after.runs.map((run) => run.text)).toEqual(["Quarterly Budget 2026", "Flate Hex Text"]);
  });

  it("改 Form XObject 里的字要落到那一层，不是页面主内容流", async () => {
    const bytes = await simpleWinAnsiFixturePdf();
    const model = await readPdfModel(bytes);
    const result = await applyPdfTextEdits(bytes, [{ target: targetOfText(model.runs, "Inner Form Text"), text: "Inner Form Test" }]);
    if (!result.ok) throw new Error("该改成功的：" + JSON.stringify(result.failures));
    const after = await readPdfModel(result.bytes);
    expect(after.runs.map((run) => run.text)).toContain("Inner Form Test");
  });

  it("CID 子集里的四个字能重排，改完仍是四个 CID", async () => {
    const bytes = await cidIdentityHFixturePdf();
    const model = await readPdfModel(bytes);
    const result = await applyPdfTextEdits(bytes, [{ target: targetOfText(model.runs, "中文标题"), text: "标题文中" }]);
    if (!result.ok) throw new Error("该改成功的：" + JSON.stringify(result.failures));
    const after = await readPdfModel(result.bytes);
    expect(after.runs[0]?.text).toBe("标题文中");
  });

  it("新文字变长就整批不写，理由给 length-mismatch（PDF 没有回流）", async () => {
    const bytes = await simpleWinAnsiFixturePdf();
    const model = await readPdfModel(bytes);
    const result = await applyPdfTextEdits(bytes, [{ target: targetOfText(model.runs, "Appendix A stays put"), text: "Appendix A stays put!" }]);
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.failures[0]?.reason).toBe("length-mismatch");
  });

  it("CID 字体没带 /ToUnicode 时读侧给 null、写侧整批拒绝，不许猜一个字上来", async () => {
    const bytes = await cidNoToUnicodeFixturePdf();
    const model = await readPdfModel(bytes);
    const run = model.runs[0];
    expect(run?.text).toBeNull();
    if (!run?.streamRef) throw new Error("夹具的内容流该是间接对象");
    const result = await applyPdfTextEdits(bytes, [{ target: { pageIndex: run.pageIndex, streamRef: run.streamRef, withinForm: run.withinForm, start: run.start }, text: "文中题标" }]);
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.failures.map((failure) => failure.reason)).toEqual(["no-encoding"]);
  });

  it("子集里没有这个字给 char-outside-subset，不许拿别的字顶上", async () => {
    const bytes = await cidIdentityHFixturePdf();
    const model = await readPdfModel(bytes);
    const result = await applyPdfTextEdits(bytes, [{ target: targetOfText(model.runs, "中文标题"), text: "中文测试" }]);
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.failures.map((failure) => failure.reason)).toEqual(["char-outside-subset"]);
  });

  it("一条好一条坏时不做半批：失败项全列出来且不产出字节", async () => {
    const bytes = await simpleWinAnsiFixturePdf();
    const model = await readPdfModel(bytes);
    const result = await applyPdfTextEdits(bytes, [
      { target: targetOfText(model.runs, "Quarterly Report 2026"), text: "Quarterly Budget 2026" },
      { target: { pageIndex: 0, streamRef: "999 0 R", withinForm: null, start: 0 }, text: "不存在的一条" },
    ]);
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.failures.map((failure) => failure.reason)).toEqual(["run-not-found"]);
  });

  it("定位键对不上现场模型时不猜：给 run-not-found 而不是改错地方", async () => {
    const bytes = await simpleWinAnsiFixturePdf();
    const model = await readPdfModel(bytes);
    const target = targetOfText(model.runs, "Quarterly Report 2026");
    const result = await applyPdfTextEdits(bytes, [{ target: { ...target, start: target.start + 1 }, text: "Quarterly Budget 2026" }]);
    expect(result.ok).toBe(false);
  });
});

describe("revertPdfTextEdits：按内容寻址撤销，不牵连后来的编辑（§2.5）", () => {
  it("撤销把那条换回原文，撤销这一次同样是前缀逐字不变 + 只追加 delta", async () => {
    const bytes = await simpleWinAnsiFixturePdf();
    const model = await readPdfModel(bytes);
    const edited = await applyPdfTextEdits(bytes, [{ target: targetOfText(model.runs, "Quarterly Report 2026"), text: "Quarterly Budget 2026" }]);
    if (!edited.ok) throw new Error("该改成功的：" + JSON.stringify(edited.failures));
    expect(edited.applied.map((entry) => [entry.fromText, entry.toText])).toEqual([["Quarterly Report 2026", "Quarterly Budget 2026"]]);

    const reverted = await revertPdfTextEdits(edited.bytes, edited.applied);
    if (!reverted.ok) throw new Error("该撤成功的：" + JSON.stringify(reverted.failures));
    expect(Buffer.compare(Buffer.from(edited.bytes.subarray(0, edited.bytes.length)), Buffer.from(reverted.bytes.subarray(0, edited.bytes.length))) === 0).toBe(true);
    expect(reverted.bytes.length).toBeGreaterThan(edited.bytes.length);
    expect((await readPdfModel(reverted.bytes)).runs.map((run) => run.text)).toEqual(["Quarterly Report 2026", "Inner Form Text", "Appendix A stays put"]);
  });

  it("撤销只回那一条：别处后来的编辑一字不动", async () => {
    const bytes = await simpleWinAnsiFixturePdf();
    const model = await readPdfModel(bytes);
    const first = await applyPdfTextEdits(bytes, [{ target: targetOfText(model.runs, "Quarterly Report 2026"), text: "Quarterly Budget 2026" }]);
    if (!first.ok) throw new Error("该改成功的：" + JSON.stringify(first.failures));
    const afterFirst = await readPdfModel(first.bytes);
    const second = await applyPdfTextEdits(first.bytes, [{ target: targetOfText(afterFirst.runs, "Appendix A stays put"), text: "Appendix A stays hot" }]);
    if (!second.ok) throw new Error("该改成功的：" + JSON.stringify(second.failures));

    const reverted = await revertPdfTextEdits(second.bytes, first.applied);
    if (!reverted.ok) throw new Error("该撤成功的：" + JSON.stringify(reverted.failures));
    expect((await readPdfModel(reverted.bytes)).runs.map((run) => run.text)).toEqual(["Quarterly Report 2026", "Inner Form Text", "Appendix A stays hot"]);
  });

  it("那条又被后面的批次改过时给 text-drifted，整批不写而不是把新改动冲掉", async () => {
    const bytes = await simpleWinAnsiFixturePdf();
    const model = await readPdfModel(bytes);
    const target = targetOfText(model.runs, "Quarterly Report 2026");
    const first = await applyPdfTextEdits(bytes, [{ target, text: "Quarterly Budget 2026" }]);
    if (!first.ok) throw new Error("该改成功的：" + JSON.stringify(first.failures));
    const second = await applyPdfTextEdits(first.bytes, [{ target, text: "Quarterly Profit 2026" }]);
    if (!second.ok) throw new Error("该改成功的：" + JSON.stringify(second.failures));

    const reverted = await revertPdfTextEdits(second.bytes, first.applied);
    expect(reverted.ok).toBe(false);
    if (reverted.ok) return;
    expect(reverted.failures.map((failure) => failure.reason)).toEqual(["text-drifted"]);
  });
});

describe("runEditBlock / pdfTargetKey：能不能改与定位键只有这一份判定", () => {
  /** 取夹具里那条主内容流的 run。 */
  async function mainRun(): Promise<PdfTextRun> {
    const model = await readPdfModel(await simpleWinAnsiFixturePdf());
    const run = model.runs.find((candidate) => candidate.text === "Quarterly Report 2026");
    if (!run) throw new Error("夹具少了主内容流那条");
    return run;
  }

  it("能改的回 null，内联流与解不出原文各回一个码（视图与 AI 都读它，§3.2）", async () => {
    const run = await mainRun();
    expect(runEditBlock(run)).toBeNull();
    expect(isRunEditable(run)).toBe(true);
    expect(runEditBlock({ ...run, streamRef: null })).toBe("inline-stream");
    expect(isRunEditable({ ...run, text: null })).toBe(false);
    expect(runEditBlock({ ...run, text: null })).toBe("undecoded");
  });

  it("内联流抽不出定位键：没有可标记的对象，改不了", async () => {
    const run = await mainRun();
    expect(targetOfRun(run)).toEqual({ pageIndex: 0, streamRef: run.streamRef, withinForm: null, start: run.start });
    expect(targetOfRun({ ...run, streamRef: null })).toBeNull();
  });

  it("定位键四个字段都进串，Form 链为空时留空位——同一条 run 在写与撤两侧拼出同一个键", async () => {
    const run = await mainRun();
    const target = targetOfRun(run);
    if (!target) throw new Error("该抽得出定位键");
    expect(pdfTargetKey(target)).toBe(`0|${run.streamRef}||${run.start}`);
    expect(pdfTargetKey({ ...target, withinForm: "Fm0/Fm1", pageIndex: 1, start: 88 })).toBe(`1|${run.streamRef}|Fm0/Fm1|88`);
  });
});
