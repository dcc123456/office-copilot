/**
 * pdf 那条出入口（`docs/office-parity-spec.md` 阶段 7）。
 *
 * 钉三件事：① 落地方式是**增量更新**——写盘的是"原文件逐字节 + 追加的 delta"，原前缀一字不动（spec D8）；
 * ② 事务语义——任一条改不动就整批不写，磁盘一个字节都不动（§2.4）；③ 撤销的凭据以现场为准，
 * 那条 run 后来又被改过时整条撤不动（§2.5）。引擎侧的解码与错误码本身在 `packages/pdf-io` 测，这里只测出入口。
 * 夹具用 `@office-copilot/pdf-io/testing`，与引擎单测共用（§3.1）。
 */
import { describe, expect, it, vi } from "vitest";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import type { PdfReadModel, PdfTextTarget } from "@office-copilot/pdf-io";
import { simpleWinAnsiFixturePdf } from "@office-copilot/pdf-io/testing";
import { PDF_EXT, editPdfDoc, openPdfDoc, revertPdfDoc } from "./pdfHandlers.js";

vi.mock("electron", () => ({ dialog: {} }));

/** 把夹具落到一个临时文件里，返回路径（测试自己负责清，§7.3）。 */
async function writePdf(name: string): Promise<string> {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "oc-pdf-"));
  const file = path.join(dir, name);
  await fs.writeFile(file, await simpleWinAnsiFixturePdf());
  return file;
}

/** 按当前磁盘上的模型定位「那条显示为 oldText 的 run」，返回改写坐标。 */
async function targetOfText(file: string, oldText: string): Promise<PdfTextTarget> {
  const opened = await openPdfDoc(file);
  if (!opened.ok || !opened.document) throw new Error("打开失败：" + opened.message);
  const run = opened.document.runs.find((candidate) => candidate.text === oldText);
  if (!run || !run.streamRef) throw new Error(`夹具里没有可编辑的「${oldText}」`);
  return { pageIndex: run.pageIndex, streamRef: run.streamRef, withinForm: run.withinForm, start: run.start };
}

/** 重开磁盘那份，问现在有没有这段文字——证明落地的确实是回包里的模型。 */
async function textOnDisk(file: string, text: string): Promise<boolean> {
  const reopened = await openPdfDoc(file);
  if (!reopened.ok || !reopened.document) throw new Error("重开失败：" + reopened.message);
  return hasText(reopened.document, text);
}

/** 读模型里是否原样存在这段文字。 */
function hasText(model: PdfReadModel, text: string): boolean {
  return model.runs.some((run) => run.text === text);
}

describe("pdf 出入口的扩展名分派", () => {
  it("只认 .pdf，大小写都行", () => {
    expect(PDF_EXT.test("报表.pdf")).toBe(true);
    expect(PDF_EXT.test("/tmp/a.PDF")).toBe(true);
    expect(PDF_EXT.test("/tmp/a.docx")).toBe(false);
  });
});

describe("pdf 出入口的落地与撤销", () => {
  it("打开→改文字：磁盘上是新文字、留一份 .bak，且原文件前缀逐字节不变（增量更新）", async () => {
    const file = await writePdf("改字.pdf");
    const opened = await openPdfDoc(file);
    expect(opened.ok && opened.document?.pageCount).toBe(2);
    const base = await fs.readFile(file);

    const edited = await editPdfDoc(file, file, [
      { target: await targetOfText(file, "Quarterly Report 2026"), text: "Quarterly Report 2027" },
    ]);
    expect(edited.ok).toBe(true);
    if (!edited.ok) return;
    expect(hasText(edited.document, "Quarterly Report 2027")).toBe(true);
    expect(await textOnDisk(file, "Quarterly Report 2027")).toBe(true);
    expect(edited.applied[0]?.fromText).toBe("Quarterly Report 2026");
    // 追加式落盘：新文件 = 原字节 + delta，前缀一个字节都没被重排。
    const out = await fs.readFile(file);
    expect(out.subarray(0, base.length).equals(base)).toBe(true);
    expect(out.length).toBe(base.length + edited.deltaBytes);
    await expect(fs.stat(`${file}.bak`)).resolves.toBeTruthy();
  });

  it("撤销写回旧文字：applied 原样送回即可回到改前，同样只追加不改前缀", async () => {
    const file = await writePdf("撤销.pdf");
    const edited = await editPdfDoc(file, file, [
      { target: await targetOfText(file, "Quarterly Report 2026"), text: "Quarterly Report 2027" },
    ]);
    expect(edited.ok && edited.applied).toHaveLength(1);
    const base = await fs.readFile(file);
    const reverted = await revertPdfDoc(file, file, edited.ok ? edited.applied : []);
    expect(reverted.ok && hasText(reverted.document, "Quarterly Report 2026")).toBe(true);
    expect(await textOnDisk(file, "Quarterly Report 2026")).toBe(true);
    if (!reverted.ok) return;
    const out = await fs.readFile(file);
    expect(out.subarray(0, base.length).equals(base)).toBe(true);
  });

  it("撤销的凭据以现场为准：那条 run 后来又被改过时整条撤不动，磁盘一个字不动（§2.5）", async () => {
    const file = await writePdf("撤不动.pdf");
    const first = await editPdfDoc(file, file, [
      { target: await targetOfText(file, "Quarterly Report 2026"), text: "Quarterly Report 2027" },
    ]);
    expect(first.ok).toBe(true);
    const second = await editPdfDoc(file, file, [
      { target: await targetOfText(file, "Quarterly Report 2027"), text: "Quarterly Report 2028" },
    ]);
    expect(second.ok).toBe(true);
    const before = await fs.readFile(file);

    const stale = await revertPdfDoc(file, file, first.ok ? first.applied : []);
    expect(stale.ok).toBe(false);
    if (stale.ok) return;
    expect(stale.failures[0]?.reason).toBe("text-drifted");
    expect((await fs.readFile(file)).equals(before)).toBe(true);
  });

  it("任一条改不动就整批不写盘：等长约束不过时磁盘与错误码一起给到渲染层", async () => {
    const file = await writePdf("整批.pdf");
    const before = await fs.readFile(file);
    const bad = await editPdfDoc(file, file, [
      { target: await targetOfText(file, "Quarterly Report 2026"), text: "Quarterly Report 2026 (draft)" },
    ]);
    expect(bad.ok).toBe(false);
    if (bad.ok) return;
    expect(bad.failures.map((item) => item.reason)).toEqual(["length-mismatch"]);
    expect((await fs.readFile(file)).equals(before)).toBe(true);
    // 失败回包**没有** message：原因码是结构化的，中文文案由渲染层组装（§5.7）。
    expect(bad).not.toHaveProperty("message");
  });

  it("底包不在时给一句跟着动作走的中文，不去创建新文件", async () => {
    const dir = await fs.mkdtemp(path.join(os.tmpdir(), "oc-pdf-"));
    const missing = path.join(dir, "不存在.pdf");
    const result = await editPdfDoc(missing, missing, [{ target: { pageIndex: 0, streamRef: "1 0 R", withinForm: null, start: 0 }, text: "x" }]);
    expect(!result.ok && "message" in result && result.message).toContain("找不到要改的那份文件");
    await expect(fs.stat(missing)).rejects.toBeTruthy();
  });
});
