import { describe, expect, it } from "vitest";
import { AggregateArgs, DescribeArgs, FindArgs, ReadRangeArgs, SHEET_TOOLS, SHEET_TOOL_SCHEMAS, parseToolArguments, type SheetToolName } from "./index.js";

/**
 * 解析一段入参并只取失败理由，断言"B9：给模型的是错误串"时用。
 *
 * @param name 工具名。
 * @param raw 模型给的 `function.arguments` 原文。
 * @returns `ok:false` 的理由串；成功时返回 null（调用方的断言据此失败）。
 */
function reasonOf(name: SheetToolName, raw: string): string | null {
  const outcome = parseToolArguments(name, raw);
  return outcome.ok ? null : outcome.reason;
}

describe("读工具面的入参校验（spec 阶段 4 · 验收 B9）", () => {
  it("缺省值与 schema 一致：表头默认算、行数默认 50", () => {
    expect(DescribeArgs.parse({})).toEqual({ hasHeader: true, maxColumns: 50 });
    expect(ReadRangeArgs.parse({ range: "A1:D20" })).toEqual({ range: "A1:D20", maxRows: 50 });
    expect(FindArgs.parse({ column: "C", criteria: ">100" })).toEqual({ column: "C", criteria: ">100", hasHeader: true, limit: 50 });
  });

  it("多余的键一律拒收：模型不许把我们没定义的条件塞进工具入参", () => {
    expect(reasonOf("sheet_find", '{"column":"C","criteria":">100","sheetName":"销售"}')).toContain("入参不合");
    expect(reasonOf("sheet_read_range", '{"range":"A1:B2","maxRows":10,"orderBy":"C"}')).toContain("orderBy");
  });

  it("越界与空值给理由，不抛异常", () => {
    expect(reasonOf("sheet_read_range", '{"range":"A1:D20","maxRows":201}')).toContain("maxRows");
    expect(reasonOf("sheet_read_range", '{"range":""}')).toContain("range");
    expect(reasonOf("sheet_aggregate", '{"groupBy":["A","B","C","D"],"aggregates":[{"fn":"sum","column":"C"}]}')).toContain("groupBy");
    expect(reasonOf("sheet_aggregate", '{"groupBy":["A"],"aggregates":[{"fn":"stdev","column":"C"}]}')).toContain("fn");
    expect(reasonOf("sheet_find", '{"column":"C","criteria":""}')).toContain("criteria");
  });

  it("非法 JSON、空串、数组这类怪东西都得是错误串（模型回包不可信）", () => {
    for (const raw of ["", "{", "not json", "[]", "null", '"A1:B2"', '{"range":"A1:B2"},']) {
      expect(reasonOf("sheet_read_range", raw)).not.toBeNull();
    }
  });

  it("不认识的工具名同样只给理由", () => {
    expect(reasonOf("sheet_drop_table" as keyof typeof SHEET_TOOL_SCHEMAS, "{}")).toContain("不认识的工具");
  });

  it("给模型的 JSON 工具定义与 zod schema 不许长歪：名字与参数面逐条对上", () => {
    expect(SHEET_TOOLS.map((tool) => tool.function.name).sort()).toEqual(Object.keys(SHEET_TOOL_SCHEMAS).sort());
    for (const tool of SHEET_TOOLS) {
      const properties = (tool.function.parameters.properties ?? {}) as Record<string, unknown>;
      // 每个 schema 的键都要在工具定义里出现，否则模型永远填不出这个参数
      expect(Object.keys(SHEET_TOOL_SCHEMAS[tool.function.name].shape).sort()).toEqual(Object.keys(properties).sort());
    }
  });

  it("必填项在两边口径一致：schema 里没默认值的键，JSON 定义必须标 required", () => {
    const requiredOf = (name: SheetToolName) => {
      const tool = SHEET_TOOLS.find((item) => item.function.name === name);
      return ((tool?.function.parameters.required ?? []) as string[]).slice().sort();
    };
    expect(requiredOf("sheet_read_range")).toEqual(["range"]);
    expect(requiredOf("sheet_aggregate")).toEqual(["aggregates", "groupBy"]);
    expect(requiredOf("sheet_find")).toEqual(["column", "criteria"]);
    expect(requiredOf("sheet_describe")).toEqual([]);
  });

  it("校验过的聚合入参能直接交给引擎：函数名与列引用原样带过去", () => {
    const outcome = parseToolArguments("sheet_aggregate", '{"groupBy":["区域"],"aggregates":[{"fn":"average","column":"C"}],"hasHeader":true}');
    expect(outcome).toEqual({
      ok: true,
      args: { groupBy: ["区域"], aggregates: [{ fn: "average", column: "C" }], hasHeader: true, limit: 50 },
    });
  });

  it("AggregateArgs 的嵌套项也拒收多余键（模型爱在聚合项里加 weight 这类东西）", () => {
    expect(AggregateArgs.safeParse({ groupBy: ["A"], aggregates: [{ fn: "sum", column: "C", weight: 2 }] }).success).toBe(false);
  });
});
