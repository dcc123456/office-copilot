import { describe, expect, it } from "vitest";
import { columnsFor, geometryFor } from "./PageSetupDialog.js";

describe("页面设置的分栏取值", () => {
  it("单栏与新建两栏都是等宽，不带逐栏宽度", () => {
    expect(columnsFor(1, 0, null)).toEqual({
      count: 1,
      gap: 0,
      equalWidth: true,
      widths: [],
    });
    expect(columnsFor(2, 30, null)).toEqual({
      count: 2,
      gap: 30,
      equalWidth: true,
      widths: [],
    });
  });

  it("栏数没变时保留文档原有的不等宽栏", () => {
    // 用户只是改页边距就点确定，不能把"窄左栏 + 宽正文"拍平成两等栏（AGENTS §2.7）。
    const source = { count: 2, gap: 30, equalWidth: false, widths: [120, 452] };
    expect(columnsFor(2, 25, source)).toEqual({
      count: 2,
      gap: 25,
      equalWidth: false,
      widths: [120, 452],
    });
    // 一旦改了栏数，逐栏宽度就没有可对应的栏了，按等宽重建。
    expect(columnsFor(3, 25, source)).toEqual({
      count: 3,
      gap: 25,
      equalWidth: true,
      widths: [],
    });
  });

  it("geometryFor 把分栏原样带进节几何", () => {
    const columns = { count: 2, gap: 30, equalWidth: true, widths: [] };
    const g = geometryFor(
      800,
      1000,
      { top: 90, right: 100, bottom: 90, left: 100 },
      48,
      48,
      columns,
    );
    expect(g.columns).toEqual(columns);
    expect(g.contentWidth).toBe(600);
    // 模板只改页边距的那条调用路径：不传分栏就是不设，不会凭空写出 `w:cols`。
    expect(geometryFor(800, 1000, { top: 90, right: 100, bottom: 90, left: 100 }, 48, 48).columns).toBeUndefined();
  });
});
