import { describe, expect, it } from "vitest";
import { describePdfIssue } from "./pdf.js";

describe("oc pdf 的装配面", () => {
  it("整份文件级别的问题不带页号，页级的那条必须带「第 N 页」", () => {
    expect(
      describePdfIssue({ code: "header-missing", severity: "broken", pageNumber: null, detail: "开头 8 字节读到的是 89 50 4E 47" }),
    ).toBe("文件头没有 %PDF- 版本标记，阅读器可能直接拒开：开头 8 字节读到的是 89 50 4E 47");
    expect(
      describePdfIssue({
        code: "content-stream-undecodable",
        severity: "broken",
        pageNumber: 3,
        detail: "/Filter = /FlateDecode，解压失败（invalid distance too far back）",
      }),
    ).toBe("第 3 页 · 这一页的内容流解不出来，打开后这一页会是空白：/Filter = /FlateDecode，解压失败（invalid distance too far back）");
    expect(
      describePdfIssue({ code: "text-not-editable", severity: "limited", pageNumber: null, detail: "" }),
    ).toBe("有一部分文字我们读不出内容，因而也改不了");
  });
});
