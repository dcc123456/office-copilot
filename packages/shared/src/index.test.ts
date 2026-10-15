import { describe, expect, it } from "vitest";
import { describePdfAuditIssue, describeSlideReadFailure, summarizePdfAudit, type PdfAuditIssue } from "./index.js";

describe("PDF 体检结论的那份文案", () => {
  it("整份文件级别的问题不带页号，页级的那条必须带「第 N 页」", () => {
    expect(
      describePdfAuditIssue({
        code: "header-missing",
        severity: "broken",
        pageNumber: null,
        detail: "开头 8 字节读到的是 89 50 4E 47",
      }),
    ).toBe("文件头没有 %PDF- 版本标记，阅读器可能直接拒开：开头 8 字节读到的是 89 50 4E 47");
    expect(
      describePdfAuditIssue({
        code: "content-stream-undecodable",
        severity: "broken",
        pageNumber: 3,
        detail: "/Filter = /FlateDecode，解压失败（invalid distance too far back）",
      }),
    ).toBe("第 3 页 · 这一页的内容流解不出来，打开后这一页会是空白：/Filter = /FlateDecode，解压失败（invalid distance too far back）");
    expect(
      describePdfAuditIssue({ code: "text-not-editable", severity: "limited", pageNumber: null, detail: "" }),
    ).toBe("有一部分文字我们读不出内容，因而也改不了");
  });

  it("toast 那一档放不下字节原文：只要句子时 detail 得能关掉", () => {
    const issue: PdfAuditIssue = { code: "eof-missing", severity: "broken", pageNumber: null, detail: '结尾 16 字节读到的是 "\\u0007z"' };
    expect(describePdfAuditIssue(issue, false)).toBe("文件尾没有 %%EOF，说明这份是被截断的，末尾的页与交叉引用可能不完整");
  });

  it("头部那句按红色与黄色各自计数，一条没有时是另一句", () => {
    const broken: PdfAuditIssue = { code: "eof-missing", severity: "broken", pageNumber: null, detail: "" };
    const limited: PdfAuditIssue = { code: "text-not-editable", severity: "limited", pageNumber: 2, detail: "181 / 1113" };
    expect(summarizePdfAudit([])).toBe("结构体检没有发现问题：这份 PDF 读得动，显示文字也都可改。");
    expect(summarizePdfAudit([broken, limited, broken])).toBe("结构体检：2 处可能让阅读器打不开或丢内容，1 处能打开但有一部分改不了。");
    // 只有一档时不能把另一档写成「0 处」，那句是给终端与面板读的，不是给机器解析的
    expect(summarizePdfAudit([limited])).toBe("结构体检：1 处能打开但有一部分改不了。");
  });
});

describe("shared：幻灯片那一页读不懂的三种说法", () => {
  it("三种失败码各念各的，且都不说成\"这一页是空的\"", () => {
    const sentences = (["missing-part", "invalid-xml", "no-sp-tree"] as const).map(describeSlideReadFailure);
    expect(new Set(sentences).size).toBe(3);
    for (const sentence of sentences) {
      expect(sentence).not.toContain("没有形状。");
      expect(sentence.length).toBeGreaterThan(0);
    }
    expect(sentences[0]).toContain("不在文件里");
    expect(sentences[1]).toContain("解析失败");
    expect(sentences[2]).toContain("形状树");
  });
});
