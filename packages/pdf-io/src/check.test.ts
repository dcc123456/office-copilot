/**
 * `.pdf` 结构体检的用例（spec D8 的 `oc pdf check` 判据）。
 *
 * 钉八条：① 干净夹具 0 条结论；② 缺 `/ToUnicode` 那份只报"改不了"那一档，不报损坏；
 * ③ 不是 PDF 的字节 → 头、尾、解析三条一起报，且**不抛**；④ 只砍掉尾巴 → 只报 `%%EOF` 缺失；
 * ⑤ 只剩开头一节 → `%%EOF` 缺失 + 目录查不到（实测撞出来的：这种文件 `load` 照样成功，`doc.catalog` 却是 undefined，
 * 不判这一条整份体检就抛出去）；⑥ `/Count` 与 `/Kids` 不等 → 页树那条；
 * ⑦ 内容流声明 `/FlateDecode` 却解不出 → 那一页单独报，别页不牵连；⑧ 一段随机二进制也只回结论。
 *
 * 坏文件一律**等长改写**真夹具的字节得到：PDF 的交叉引用表按字节偏移寻址，改一个字节数就会连带一堆假问题，
 * 那样测到的就不是"这一条判定"而是"偏移全乱了"。
 */
import { describe, expect, it } from "vitest";
import { auditPdfStructure } from "./check.js";
import { cidNoToUnicodeFixturePdf, flateWinAnsiFixturePdf, simpleWinAnsiFixturePdf } from "./test-pdf.js";

/** 读侧的 latin1 视图（PDF 的字典与操作数都是 ASCII，latin1 保证一字节一字符、下标即偏移）。 */
function asLatin1(bytes: Uint8Array): string {
  return Buffer.from(bytes).toString("latin1");
}

/**
 * 等长替换一段字节。
 *
 * @param bytes 原文件字节。
 * @param needle 要定位的 ASCII 片段（必须唯一）。
 * @param replacement 与 needle **同长**的替换文本，保证 xref 偏移不位移。
 * @returns 改过的新字节。
 */
function patchSameLength(bytes: Uint8Array, needle: string, replacement: string): Uint8Array {
  const text = asLatin1(bytes);
  const at = text.indexOf(needle);
  if (at < 0) throw new Error(`夹具里找不到 ${JSON.stringify(needle)}，用例前提已经不成立了`);
  if (text.indexOf(needle, at + 1) >= 0) throw new Error(`${JSON.stringify(needle)} 在夹具里出现多次，等长替换会改到别处`);
  if (needle.length !== replacement.length) throw new Error("替换文本必须与原文同长，否则交叉引用偏移全部位移");
  const next = Uint8Array.from(bytes);
  for (let index = 0; index < replacement.length; index += 1) next[at + index] = replacement.charCodeAt(index);
  return next;
}

/**
 * 把第一条 `/FlateDecode` 流的正文换成等长的非法字节（zlib 头都不对，解压必抛）。
 *
 * @param bytes 原文件字节。
 * @returns 改过的字节与那段正文的长度（断言用）。
 */
function breakFirstFlateStream(bytes: Uint8Array): { patched: Uint8Array; bodyLength: number } {
  const text = asLatin1(bytes);
  const dictAt = text.indexOf("/FlateDecode");
  if (dictAt < 0) throw new Error("夹具里没有 FlateDecode 流");
  const streamKeyword = text.indexOf("stream", dictAt);
  const bodyStart = streamKeyword + "stream".length + (text[streamKeyword + 7] === "\r" ? 2 : 1);
  const bodyEnd = text.indexOf("endstream", bodyStart);
  const next = Uint8Array.from(bytes);
  next.fill(0x00, bodyStart, bodyEnd);
  return { patched: next, bodyLength: bodyEnd - bodyStart };
}

describe("auditPdfStructure 的干净面", () => {
  it("两页 WinAnsi 夹具：没有一条结论（既没损坏，也没有改不了的文字）", async () => {
    expect(await auditPdfStructure(await simpleWinAnsiFixturePdf())).toEqual([]);
  });

  it("缺 /ToUnicode 的那份只报 limited 那一档，不谎报损坏", async () => {
    const issues = await auditPdfStructure(await cidNoToUnicodeFixturePdf());
    expect(issues.map((issue) => issue.code)).toEqual(["text-not-editable"]);
    expect(issues[0]?.severity).toBe("limited");
    expect(issues[0]?.detail).toContain("认不出内容");
    expect(issues[0]?.detail).toContain("/ToUnicode");
  });
});

describe("auditPdfStructure 的坏文件面", () => {
  it("根本不是 PDF：头、尾、解析三条都报，且函数不抛", async () => {
    const issues = await auditPdfStructure(new Uint8Array(Buffer.from("hello, this is definitely not a pdf file at all\n", "latin1")));
    expect(issues.map((issue) => issue.code)).toEqual(["header-missing", "eof-missing", "parse-failed"]);
    expect(issues.every((issue) => issue.severity === "broken")).toBe(true);
  });

  it("砍掉尾巴 60 字节：只报 %%EOF 缺失，别的不牵连（pdf-lib 能把对象恢复出来）", async () => {
    const bytes = (await simpleWinAnsiFixturePdf()).slice(0, -60);
    expect((await auditPdfStructure(bytes)).map((issue) => issue.code)).toEqual(["eof-missing"]);
  });

  it("只剩开头一节：%%EOF 缺失 + 目录查不到，且**不抛**（实测：这种文件 `load` 照样成功，`doc.catalog` 却是 undefined）", async () => {
    const bytes = (await simpleWinAnsiFixturePdf()).slice(0, 120);
    expect((await auditPdfStructure(bytes)).map((issue) => issue.code)).toEqual(["eof-missing", "catalog-missing"]);
  });

  it("/Count 与 /Kids 不等：页树那条报出来，数字写在 detail 里", async () => {
    const issues = await auditPdfStructure(patchSameLength(await simpleWinAnsiFixturePdf(), "/Count 2", "/Count 3"));
    const pageIssue = issues.find((issue) => issue.code === "page-count-mismatch");
    expect(pageIssue?.detail).toContain("/Count = 3");
    expect(pageIssue?.detail).toContain("/Kids 实际 2 项");
  });

  it("内容流解不出来：只报那一页，另一页不牵连", async () => {
    const { patched, bodyLength } = breakFirstFlateStream(await flateWinAnsiFixturePdf());
    expect(bodyLength).toBeGreaterThan(8);
    const issues = await auditPdfStructure(patched);
    const streamIssues = issues.filter((issue) => issue.code === "content-stream-undecodable");
    expect(streamIssues.map((issue) => issue.pageNumber)).toEqual([1]);
    expect(streamIssues[0]?.detail).toContain("FlateDecode");
  });

  it("体检不抛：给一段随机二进制也只回结论", async () => {
    const random = Uint8Array.from({ length: 4096 }, (_, index) => (index * 37) % 256);
    await expect(auditPdfStructure(random)).resolves.toBeInstanceOf(Array);
  });
});
