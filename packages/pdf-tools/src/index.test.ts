/**
 * PDF 工具面的入参用例（验收 B9：模型给什么都不许抛，非法的一律换成回给它看的中文串）。
 *
 * 与 `pptx-tools`/`sheet-tools` 那两份同形：坏 JSON、缺字段、多字段、越界、不存在的工具名、
 * 未被提供的改动类型（`set_run_font` 这一类现在根本不开）都只回理由；另外钉住上送的定义与 schema 表键一致。
 */
import { describe, expect, it } from "vitest";
import { PDF_TOOLS, PDF_TOOL_SCHEMAS, parseToolArguments } from "./index.js";

/** 取 `pdf_submit_proposal` 那份入参（成功时用来读 edits）。 */
function proposal(raw: unknown) {
  return parseToolArguments("pdf_submit_proposal", JSON.stringify(raw));
}

describe("pdf_describe / pdf_read_page_text 的入参", () => {
  it("缺省值由 schema 补上：不给 maxPages 就是 50，不给 maxRunsPerPage 就是 50", () => {
    const describeArgs = parseToolArguments("pdf_describe", "{}");
    expect(describeArgs.ok ? describeArgs.args.maxPages : describeArgs.reason).toBe(50);
    const read = parseToolArguments("pdf_read_page_text", '{"pageNumber":3}');
    expect(read.ok ? read.args : read.reason).toEqual({ pageNumber: 3, maxRunsPerPage: 50 });
  });

  it("页号 0 与负数、非整数、超过上限都拒，理由里带上出问题的字段名", () => {
    for (const raw of ['{"pageNumber":0}', '{"pageNumber":-2}', '{"pageNumber":1.5}', '{"pageNumber":501}']) {
      const parsed = parseToolArguments("pdf_read_page_text", raw);
      expect(parsed.ok).toBe(false);
      if (parsed.ok) return;
      expect(parsed.reason).toContain("pageNumber");
    }
  });

  it("多字段与坏 JSON 都只回一句理由，不抛异常", () => {
    expect(parseToolArguments("pdf_describe", '{"maxPages":5,"maxRuns":9}').ok).toBe(false);
    const broken = parseToolArguments("pdf_describe", "{maxPages");
    expect(broken.ok).toBe(false);
    if (broken.ok) return;
    expect(broken.reason).toContain("不是合法 JSON");
  });

  it("模型点了不存在的方法名时回「不认识的工具」（调用方仍不会抛）", () => {
    const parsed = parseToolArguments("pdf_edit_page" as never, "{}");
    expect(parsed.ok).toBe(false);
    if (parsed.ok) return;
    expect(parsed.reason).toContain("不认识的工具");
  });
});

describe("pdf_submit_proposal 的入参", () => {
  it("一条等长替换照收，字段原样交给代理去翻定位键", () => {
    const parsed = proposal({ summary: "把标题年份改成 2027", edits: [{ type: "set_run_text", pageNumber: 1, runIndex: 0, text: "Quarterly Report 2027" }] });
    expect(parsed.ok).toBe(true);
    if (!parsed.ok) return;
    expect(parsed.args.edits).toHaveLength(1);
    expect(parsed.args.edits[0]).toMatchObject({ type: "set_run_text", pageNumber: 1, runIndex: 0 });
  });

  it("现在只有 set_run_text 一类：改字体/颜色/几何那一类条目直接拒", () => {
    for (const type of ["set_run_font", "set_run_color", "set_shape_geometry", "remove_shape", "add_text_box"]) {
      const parsed = proposal({ summary: "试一条不给的改动", edits: [{ type, pageNumber: 1, runIndex: 0, text: "一样长" }] });
      expect(parsed.ok).toBe(false);
    }
  });

  it("runIndex 越界、文字空串、edits 空数组、多塞字段都拒", () => {
    const cases: unknown[] = [
      { summary: "序号超上限", edits: [{ type: "set_run_text", pageNumber: 1, runIndex: 5001, text: "一样长" }] },
      { summary: "空文字", edits: [{ type: "set_run_text", pageNumber: 1, runIndex: 0, text: "" }] },
      { summary: "没有改动", edits: [] },
      { summary: "多字段", edits: [{ type: "set_run_text", pageNumber: 1, runIndex: 0, text: "一样长", bold: true }] },
      { summary: "", edits: [{ type: "set_run_text", pageNumber: 1, runIndex: 0, text: "一样长" }] },
    ];
    for (const raw of cases) expect(proposal(raw).ok).toBe(false);
  });

  it("提案里不给字节数、内容流引用这类引擎坐标：模型想塞也塞不进去", () => {
    const parsed = proposal({ summary: "自己报坐标", edits: [{ type: "set_run_text", pageNumber: 1, runIndex: 0, text: "一样长", streamRef: "6 0 R", start: 32 }] });
    expect(parsed.ok).toBe(false);
    if (parsed.ok) return;
    expect(parsed.reason).toContain("streamRef");
  });
});

describe("上送的工具定义", () => {
  it("三个工具名与 schema 表的键一一对上，且都关着 additionalProperties", () => {
    expect(PDF_TOOLS.map((tool) => tool.function.name)).toEqual(["pdf_describe", "pdf_read_page_text", "pdf_submit_proposal"]);
    expect(Object.keys(PDF_TOOL_SCHEMAS).sort()).toEqual(PDF_TOOLS.map((tool) => tool.function.name).sort());
    for (const tool of PDF_TOOLS) expect(tool.function.parameters.additionalProperties).toBe(false);
  });

  it("写面那份描述里把「字数完全相同」这条硬约束说了出来（模型不看代码只看描述）", () => {
    const submit = PDF_TOOLS.find((tool) => tool.function.name === "pdf_submit_proposal");
    expect(submit?.function.description).toContain("字数完全相同");
    expect(submit?.function.description).toContain("提不了");
  });
});
