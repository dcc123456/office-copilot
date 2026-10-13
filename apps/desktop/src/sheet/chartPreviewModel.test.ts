import { describe, expect, it } from "vitest";
import { chartPreviewConfig } from "./chartPreviewModel.js";
import type { ResolvedChart } from "@office-copilot/sheet-io";

/**
 * 一份最小的现场图表：三种类别、一个系列。
 *
 * @param overrides 覆盖默认字段，用于测图型与系列数带来的分支。
 * @returns 引擎解析完的图表（数值与引用串都是引擎会给的形状）。
 */
function chartFixture(overrides: Partial<ResolvedChart> = {}): ResolvedChart {
  return {
    kind: "bar",
    title: "销量对比",
    anchor: { colIndex: 4, rowIndex: 1 },
    categories: ["1月", "2月", "3月"],
    categoryRef: "'销售'!$A$2:$A$4",
    categoriesAreNumeric: false,
    categoryNumbers: [],
    series: [{ name: "销量", nameRef: "'销售'!$C$1", valueRef: "'销售'!$C$2:$C$4", values: [100, 120, 150] }],
    widthPx: 480,
    heightPx: 288,
    ...overrides,
  };
}

describe("图表预览配置（D5 的 app 内那一半）", () => {
  it("类别、系列名与数值逐条搬过去，一个数字都不自己算", () => {
    const config = chartPreviewConfig(chartFixture());
    expect(config.type).toBe("bar");
    expect(config.data.labels).toEqual(["1月", "2月", "3月"]);
    expect(config.data.datasets).toEqual([{ label: "销量", data: [100, 120, 150], borderColor: "#4e79a7", backgroundColor: "#4e79a7" }]);
    expect(config.options.plugins.title).toEqual({ display: true, text: "销量对比" });
  });

  it("图型一一对应，标题为空时不占位", () => {
    expect(chartPreviewConfig(chartFixture({ kind: "line" })).type).toBe("line");
    const pie = chartPreviewConfig(chartFixture({ kind: "pie", title: "" }));
    expect(pie.type).toBe("pie");
    expect(pie.options.plugins.title.display).toBe(false);
  });

  it("单系列柱/线不画图例，饼图与多系列才画；饼图按扇区配色", () => {
    const single = chartFixture();
    const second = { ...single.series[0]!, name: "利润", values: [10, 20, 30] };
    expect(chartPreviewConfig(single).options.plugins.legend.display).toBe(false);
    const twoSeries = chartPreviewConfig(chartFixture({ series: [single.series[0]!, second] }));
    expect(twoSeries.options.plugins.legend.display).toBe(true);
    // 第二条系列换色，多条按调色板轮着取
    expect(twoSeries.data.datasets.map((dataset) => dataset.borderColor)).toEqual(["#4e79a7", "#f28e2b"]);
    const pie = chartPreviewConfig(chartFixture({ kind: "pie" }));
    expect(pie.options.plugins.legend.display).toBe(true);
    expect(pie.data.datasets[0]!.backgroundColor?.length).toBeGreaterThan(1);
  });

  it("数值数组是复制出来的，改预览不会动到引擎那份", () => {
    const chart = chartFixture();
    const config = chartPreviewConfig(chart);
    config.data.datasets[0]!.data[0] = 0;
    config.data.labels.push("4月");
    expect(chart.series[0]!.values).toEqual([100, 120, 150]);
    expect(chart.categories).toEqual(["1月", "2月", "3月"]);
  });
});
