import { describe, expect, it } from "vitest";
import { parseMetric, pickSheet, splitColumns } from "./sheet.js";
import type { SheetPayload } from "@office-copilot/sheet-io";

/**
 * 造一张最小的表载荷：CLI 的这几条纯函数只认名字与序号，不碰格子内容。
 *
 * @param name 工作表名。
 * @returns 够 `pickSheet` 用的载荷。
 */
function fakeSheet(name: string): SheetPayload {
  return {
    name,
    sheetId: 1,
    partPath: `xl/worksheets/sheet-${name}.xml`,
    lastColIndex: 0,
    lastRowIndex: 0,
    cells: {},
    merges: [],
    columns: [],
    rows: [],
    warnings: [],
  } satisfies SheetPayload;
}

describe("命令行参数怎么翻成引擎的入参", () => {
  it("挑表：认名字，也认 info 里显示的 1 基序号", () => {
    const sheets = [fakeSheet("销售"), fakeSheet("空表")];
    expect(pickSheet(sheets)?.name).toBe("销售");
    expect(pickSheet(sheets, "空表")?.name).toBe("空表");
    expect(pickSheet(sheets, "2")?.name).toBe("空表");
    // 认不出的一律 undefined，由命令层回一句"可选：1=销售、2=空表"，不许默默用第一张表。
    expect(pickSheet(sheets, "3")).toBeUndefined();
    expect(pickSheet(sheets, "汇总")).toBeUndefined();
  });

  it("聚合度量：avg 是 average 的别名，函数名与列引用都能带空格或冒号", () => {
    expect(parseMetric("sum:销量")).toEqual({ fn: "sum", column: "销量" });
    expect(parseMetric("avg C")).toEqual({ fn: "average", column: "C" });
    expect(parseMetric("count:月份")).toEqual({ fn: "count", column: "月份" });
  });

  it("分组键与判定键：逗号分隔，空白丢掉", () => {
    expect(splitColumns(" 月份, 地区 ,")).toEqual(["月份", "地区"]);
    expect(splitColumns("地区")).toEqual(["地区"]);
  });
});
