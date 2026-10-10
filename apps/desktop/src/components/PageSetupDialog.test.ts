import { describe, expect, it } from "vitest";
import {
  columnsFor,
  geometryFor,
  lineNumbersFor,
  pageBordersFor,
  watermarkFor,
  type WatermarkForm,
} from "./PageSetupDialog.js";
import { emptyPageBorders } from "../editor/pageBorders.js";

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
      { columns },
    );
    expect(g.columns).toEqual(columns);
    expect(g.contentWidth).toBe(600);
    // 模板只改页边距的那条调用路径：不传就是不设，不会凭空写出 `w:cols`。
    expect(
      geometryFor(800, 1000, { top: 90, right: 100, bottom: 90, left: 100 }, 48, 48).columns,
    ).toBeUndefined();
  });

  it("geometryFor 把边框、行号与水印一并带进节几何", () => {
    const pageBorders = { ...emptyPageBorders(), top: { style: "single", sizePt: 1, spacePt: 2, color: null } };
    const lineNumbers = { countBy: 2, distance: 32, start: 5 };
    const watermark = {
      text: "机密",
      font: "黑体",
      sizePt: null,
      color: "#c0c0c0",
      semiTransparent: true,
      layout: "diagonal" as const,
    };
    const g = geometryFor(
      800,
      1000,
      { top: 90, right: 100, bottom: 90, left: 100 },
      48,
      48,
      { pageBorders, lineNumbers, watermark },
    );
    expect(g.pageBorders).toEqual(pageBorders);
    expect(g.lineNumbers).toEqual(lineNumbers);
    expect(g.watermark).toEqual(watermark);
    // 没碰过的几项保持 undefined：docx-io 据此原样保留文件里的 `w:pgBorders`/`w:lnNumType` 与页眉里那个艺术字。
    const untouched = geometryFor(
      800,
      1000,
      { top: 90, right: 100, bottom: 90, left: 100 },
      48,
      48,
      { columns: { count: 1, gap: 30, equalWidth: true, widths: [] } },
    );
    expect(untouched.pageBorders).toBeUndefined();
    expect(untouched.lineNumbers).toBeUndefined();
    expect(untouched.watermark).toBeUndefined();
  });
});

describe("页面设置里边框与行号的取值", () => {
  it("用户没碰过边框时返回 undefined，文件里那条元素原样不动", () => {
    expect(pageBordersFor(false, emptyPageBorders())).toBeUndefined();
    expect(
      pageBordersFor(
        false,
        { ...emptyPageBorders(), top: { style: "single", sizePt: 1, spacePt: 1, color: null } },
      ),
    ).toBeUndefined();
  });

  it("四条边都关掉时返回 null，保存要删掉 `w:pgBorders`", () => {
    expect(pageBordersFor(true, emptyPageBorders())).toBeNull();
  });

  it("改过边框时把四条边原样交出去", () => {
    const draft = {
      offsetFrom: "text",
      top: { style: "double", sizePt: 2, spacePt: 3, color: "#ff0000" },
      left: null,
      bottom: null,
      right: null,
    };
    expect(pageBordersFor(true, draft)).toEqual(draft);
  });

  it("行号：没碰过不写，关掉返回 null，开着换算成 countBy/距离/起始号", () => {
    expect(lineNumbersFor(false, { enabled: true, countBy: 1, start: 1, distancePx: 32 })).toBeUndefined();
    expect(lineNumbersFor(true, { enabled: false, countBy: 1, start: 1, distancePx: 32 })).toBeNull();
    expect(lineNumbersFor(true, { enabled: true, countBy: 2, start: 5, distancePx: 32 })).toEqual({
      countBy: 2,
      distance: 32,
      start: 5,
    });
    // 输入框被清空时按"关掉"处理，而不是写出 `w:countBy="0"` 这种 Word 不接受的值。
    expect(lineNumbersFor(true, { enabled: true, countBy: 0, start: 1, distancePx: 32 })).toBeNull();
  });
});

describe("页面设置里水印的取值", () => {
  /** 对话框草稿的公共底，每条用例只覆盖自己要看的那个字段。 */
  const form = (patch: Partial<WatermarkForm> = {}): WatermarkForm => ({
    enabled: true,
    text: "机密",
    font: "黑体",
    sizePt: 0,
    color: "#ff0000",
    semiTransparent: false,
    layout: "horizontal",
    ...patch,
  });

  it("用户没碰过水印区时返回 undefined，页眉里已有的艺术字原样不动", () => {
    expect(watermarkFor(false, form())).toBeUndefined();
    // 美化排版那条路径同理：模板没有水印角色，就不该替用户动它（AGENTS §2.7）。
    expect(watermarkFor(false, form({ enabled: false, text: "" }))).toBeUndefined();
  });

  it("取消勾选或清空文字时返回 null，保存要把水印删掉", () => {
    expect(watermarkFor(true, form({ enabled: false }))).toBeNull();
    expect(watermarkFor(true, form({ text: "   " }))).toBeNull();
  });

  it("字号留 0 就是 Word 的“自动”，落到模型里是 sizePt: null", () => {
    expect(watermarkFor(true, form())).toEqual({
      text: "机密",
      font: "黑体",
      sizePt: null,
      color: "#ff0000",
      semiTransparent: false,
      layout: "horizontal",
    });
    expect(watermarkFor(true, form({ sizePt: 48 }))).toMatchObject({ sizePt: 48 });
  });

  it("文字与字体去掉首尾空格，字体留空时回到 Word 的缺省宋体", () => {
    expect(watermarkFor(true, form({ text: "  内部资料  " }))).toMatchObject({
      text: "内部资料",
    });
    expect(watermarkFor(true, form({ font: "  " }))).toMatchObject({ font: "宋体" });
  });
});
