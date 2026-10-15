import { describe, expect, it } from "vitest";
import { PptxEdit, PPTX_TOOL_SCHEMAS, PPTX_TOOLS, parseToolArguments } from "./index.js";

describe("pptx-tools：入参解析永不抛异常（验收 B9）", () => {
  it("合法 JSON 通过并套上默认值", () => {
    const r = parseToolArguments("ppt_describe", "{}");
    expect(r.ok).toBe(true);
    if (r.ok) expect(r.args).toEqual({ maxSlides: 50, maxShapesPerSlide: 40 });
  });

  it("不认识的工具名给理由而不是抛异常", () => {
    const r = parseToolArguments("no_such_tool" as never, "{}");
    expect(r.ok).toBe(false);
    if (!r.ok) expect(r.reason).toContain("不认识的工具");
  });

  it("入参不是合法 JSON 给理由", () => {
    const r = parseToolArguments("ppt_read_slide_text", "{不是json");
    expect(r.ok).toBe(false);
    if (!r.ok) expect(r.reason).toContain("不是合法 JSON");
  });

  it("strict：多给的字段被拒", () => {
    const r = parseToolArguments("ppt_read_slide_text", '{"slideNumber":1,"别多写":true}');
    expect(r.ok).toBe(false);
  });

  it("缺必填字段被拒（slideNumber）", () => {
    const r = parseToolArguments("ppt_read_slide_text", "{}");
    expect(r.ok).toBe(false);
    if (!r.ok) expect(r.reason).toContain("slideNumber");
  });

  it("越界：页号 0 被拒（1 基）", () => {
    const r = parseToolArguments("ppt_read_slide_text", '{"slideNumber":0}');
    expect(r.ok).toBe(false);
  });
});

describe("pptx-tools：提案的八类改动", () => {
  const one = (edit: unknown) => parseToolArguments("ppt_submit_proposal", JSON.stringify({ summary: "改一处", edits: [edit] }));

  it("set_run_text 合法入参通过", () => {
    const r = one({ type: "set_run_text", slideNumber: 1, shapeId: 2, paraIndex: 0, runIndex: 0, text: "新标题" });
    expect(r.ok).toBe(true);
  });

  it("set_run_font 什么都不给（bold/sizePt 都缺）被 refine 拒", () => {
    const r = one({ type: "set_run_font", slideNumber: 1, shapeId: 2, paraIndex: 0, runIndex: 0 });
    expect(r.ok).toBe(false);
    if (!r.ok) expect(r.reason).toContain("至少给一个");
  });

  it("set_run_font 只给 sizePt 通过；字号越界（401pt）被拒", () => {
    expect(one({ type: "set_run_font", slideNumber: 1, shapeId: 2, paraIndex: 0, runIndex: 0, sizePt: 24 }).ok).toBe(true);
    expect(one({ type: "set_run_font", slideNumber: 1, shapeId: 2, paraIndex: 0, runIndex: 0, sizePt: 401 }).ok).toBe(false);
  });

  it("set_run_color 认 #rrggbb 与省略 #，拒绝 red 这种写法", () => {
    expect(one({ type: "set_run_color", slideNumber: 1, shapeId: 2, paraIndex: 0, runIndex: 0, color: "#C00000" }).ok).toBe(true);
    expect(one({ type: "set_run_color", slideNumber: 1, shapeId: 2, paraIndex: 0, runIndex: 0, color: "c00000" }).ok).toBe(true);
    const bad = one({ type: "set_run_color", slideNumber: 1, shapeId: 2, paraIndex: 0, runIndex: 0, color: "red" });
    expect(bad.ok).toBe(false);
  });

  it("set_shape_geometry 四个方位一个都不给被 refine 拒；只给 leftPx 通过", () => {
    expect(one({ type: "set_shape_geometry", slideNumber: 1, shapeId: 4 }).ok).toBe(false);
    expect(one({ type: "set_shape_geometry", slideNumber: 1, shapeId: 4, leftPx: 100 }).ok).toBe(true);
  });

  it("set_shape_geometry 宽度非正被拒", () => {
    expect(one({ type: "set_shape_geometry", slideNumber: 1, shapeId: 4, widthPx: 0 }).ok).toBe(false);
  });

  it("move_shape_layer 认四个方向；缺方向或方向拼错都被拒", () => {
    expect(one({ type: "move_shape_layer", slideNumber: 1, shapeId: 4, direction: "top" }).ok).toBe(true);
    expect(one({ type: "move_shape_layer", slideNumber: 1, shapeId: 4, direction: "front" }).ok).toBe(false);
    expect(one({ type: "move_shape_layer", slideNumber: 1, shapeId: 4 }).ok).toBe(false);
  });

  it("add_text_box 四个几何值与文字都要给：缺任一被拒，宽高非正被拒，空文字被拒", () => {
    const box = { type: "add_text_box", slideNumber: 1, leftPx: 96, topPx: 96, widthPx: 480, heightPx: 144, text: "谢谢观看" };
    expect(one(box).ok).toBe(true);
    expect(one({ ...box, text: "" }).ok).toBe(false);
    expect(one({ ...box, widthPx: 0 }).ok).toBe(false);
    expect(one({ ...box, heightPx: -1 }).ok).toBe(false);
    for (const field of ["leftPx", "topPx", "widthPx", "heightPx", "text"] as const) {
      const withoutField = Object.fromEntries(Object.entries(box).filter(([key]) => key !== field));
      expect(one(withoutField).ok).toBe(false);
    }
  });

  it("add_text_box 不许带 shapeId（新形状号由引擎落地时分配，模型引用不到）", () => {
    expect(one({ type: "add_text_box", slideNumber: 1, leftPx: 0, topPx: 0, widthPx: 100, heightPx: 100, text: "甲", shapeId: 9 }).ok).toBe(false);
  });

  it("remove_shape 要页号与形状号；缺形状号或形状号越界被拒", () => {
    expect(one({ type: "remove_shape", slideNumber: 2, shapeId: 4 }).ok).toBe(true);
    expect(one({ type: "remove_shape", slideNumber: 2 }).ok).toBe(false);
    expect(one({ type: "remove_shape", slideNumber: 2, shapeId: 0 }).ok).toBe(false);
  });

  it('set_notes 只要页号与逐段文字：空串段落合法，空数组、缺字段、单段超 2000 字都被拒', () => {
    expect(one({ type: "set_notes", slideNumber: 1, paragraphTexts: ["甲", ""] }).ok).toBe(true);
    expect(one({ type: "set_notes", slideNumber: 1 }).ok).toBe(false);
    expect(one({ type: "set_notes", slideNumber: 1, paragraphTexts: [] }).ok).toBe(false);
    expect(one({ type: "set_notes", slideNumber: 1, paragraphTexts: ["x".repeat(2001)] }).ok).toBe(false);
  });

  it("改文字超过 2000 字被拒", () => {
    const r = one({ type: "set_run_text", slideNumber: 1, shapeId: 2, paraIndex: 0, runIndex: 0, text: "x".repeat(2001) });
    expect(r.ok).toBe(false);
  });

  it("未知改动类型被拒", () => {
    const r = one({ type: "add_shape", slideNumber: 1 });
    expect(r.ok).toBe(false);
  });
});

describe("pptx-tools：导出一致性", () => {
  it("每个工具定义都能在 schema 表里找到同名 schema（不给模型一个校验不了的口子）", () => {
    const names = Object.keys(PPTX_TOOL_SCHEMAS);
    for (const tool of PPTX_TOOLS) {
      expect(names).toContain(tool.function.name);
      expect(tool.function.parameters.additionalProperties).toBe(false);
    }
    expect(PPTX_TOOLS.map((t) => t.function.name).sort()).toEqual(names.sort());
  });

  it("zod 联合里的每一类改动都在给模型的 JSON oneOf 里（只加校验忘了声明，模型就提不了那条）", () => {
    const zodTypes = PptxEdit.options.map((option) => String(option.shape.type.value)).sort();
    const submit = PPTX_TOOLS.find((tool) => tool.function.name === "ppt_submit_proposal");
    if (!submit) throw new Error("没有 ppt_submit_proposal 这条工具定义");
    const branches = (
      submit.function.parameters.properties as { edits: { items: { oneOf: { properties: { type: { const: string } } }[] } } }
    ).edits.items.oneOf;
    expect(branches.map((branch) => branch.properties.type.const).sort()).toEqual(zodTypes);
    expect(zodTypes).toEqual([
      "add_text_box",
      "move_shape_layer",
      "remove_shape",
      "set_notes",
      "set_run_color",
      "set_run_font",
      "set_run_text",
      "set_shape_geometry",
    ]);
  });
});
