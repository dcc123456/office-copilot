import { describe, expect, it } from "vitest";
import {
  ParagraphFormatting,
  sameTabStopPosition,
  withTabStop,
  withoutTabStop,
} from "./ParagraphFormatting.js";

/**
 * 段落属性（制表位、首字下沉）的 DOM 读写规则。粘贴与 `getHTML/setContent(html)` 走的是 PM 的
 * DOMSerializer / DOMParser，只认属性自己的 renderHTML / parseHTML——没有 parseHTML
 * 的属性在粘贴后一律回到 default，一段带点前导符的目录条目就会变成"没建模过"，
 * 导出时用户的 w:tabs 又透传回来，清掉的东西复活。
 *
 * 签名取自 @tiptap/core dist/index.js 的实际调用点：属性级 `renderHTML(attrs)`
 * 首参是 attrs 全体，`parseHTML(el)` 首参是 DOM 元素。
 */

type AttrRule = {
  default: unknown;
  parseHTML: (el: { getAttribute: (name: string) => string | null }) => unknown;
  renderHTML: (attrs: Record<string, unknown>) => Record<string, string>;
};

// addGlobalAttributes 返回的是"每组 types 一份属性表"的数组，段落属性在含
// paragraph 的那一组里。
const rules = (
  ParagraphFormatting.config.addGlobalAttributes as unknown as () => {
    types: string[];
    attributes: { tabStops: AttrRule; dropCap: AttrRule };
  }[]
)().find((group) => group.types.includes("paragraph"))!.attributes;

/** 只实现 getAttribute 的元素替身：parseHTML 规则只用得到它。 */
const elementWith = (values: Record<string, string>) => ({
  getAttribute: (name: string) => values[name] ?? null,
});

describe("首字下沉属性的 DOM 读写", () => {
  it("默认值是 undefined：新建段落不该声明“这一段关掉了 w:framePr”", () => {
    expect(rules.dropCap.default).toBeUndefined();
  });

  it("位置与行数原样进出，行数同时写成 CSS 变量", () => {
    const html = rules.dropCap.renderHTML({ dropCap: { position: "drop", lines: 3 } });
    expect(html["data-drop-cap"]).toBe('{"position":"drop","lines":3}');
    expect(html.class).toBe("oc-drop-cap");
    expect(html.style).toBe("--oc-drop-cap: 3");
    expect(rules.dropCap.parseHTML(elementWith(html))).toEqual({ position: "drop", lines: 3 });
  });

  it("突出到页边距外是两值语法：跨几行但只下沉一行", () => {
    expect(
      rules.dropCap.renderHTML({ dropCap: { position: "margin", lines: 4 } }).style,
    ).toBe("--oc-drop-cap: 4 1");
  });

  it("源文档没写行数时不凭空补一个：CSS 用 Word 的 3 行，data-* 里仍然没有 lines", () => {
    const html = rules.dropCap.renderHTML({ dropCap: { position: "drop" } });
    expect(html["data-drop-cap"]).toBe('{"position":"drop"}');
    expect(html.style).toBe("--oc-drop-cap: 3");
    expect(rules.dropCap.parseHTML(elementWith(html))).toEqual({ position: "drop" });
  });

  it("关掉只留空属性，不挂类：那条 ::first-letter 规则有缺省行数", () => {
    expect(rules.dropCap.renderHTML({ dropCap: null })).toEqual({ "data-drop-cap": "" });
    expect(rules.dropCap.parseHTML(elementWith({ "data-drop-cap": "" }))).toBeNull();
  });

  it("属性不在、JSON 读不懂、位置不是 Word 的取值，一律退回 undefined", () => {
    for (const raw of [undefined, "not json", "null", "[]", '{"position":"raise"}']) {
      const values: Record<string, string> = raw === undefined ? {} : { "data-drop-cap": raw };
      expect(rules.dropCap.parseHTML(elementWith(values))).toBeUndefined();
    }
    // 没建模过时什么都不渲染，段落的 DOM 不会长出空属性。
    expect(rules.dropCap.renderHTML({})).toEqual({});
  });

  it("行数坏掉只丢行数，不把整条下沉判成没建模过", () => {
    // 面板据此显示"下沉"而不是"无"；行数交给 Word 的缺省 3 行，导出不会凭空写 w:lines。
    for (const raw of ['{"position":"drop","lines":0}', '{"position":"drop","lines":2.5}', '{"position":"drop","lines":"3"}']) {
      expect(rules.dropCap.parseHTML(elementWith({ "data-drop-cap": raw }))).toEqual({
        position: "drop",
      });
    }
    expect(
      rules.dropCap.parseHTML(elementWith({ "data-drop-cap": '{"position":"drop","lines":1}' })),
    ).toEqual({ position: "drop", lines: 1 });
  });
});

describe("制表位属性的 DOM 读写", () => {
  it("默认值是 undefined：新建段落不该声明“制表位已全部清除”", () => {
    expect(rules.tabStops.default).toBeUndefined();
  });

  it("清单原样进出，位置保持 pt 不做第二次换算", () => {
    const stops = [
      { positionPt: 20, alignment: "center" },
      { positionPt: 415.3, alignment: "right", leader: "dot" },
    ];
    const html = rules.tabStops.renderHTML({ tabStops: stops });
    expect(JSON.parse(html["data-tab-stops"] as string)).toEqual(stops);
    expect(rules.tabStops.parseHTML(elementWith(html))).toEqual(stops);
  });

  it("空清单照样写出去：它与“没建模过”是两回事", () => {
    expect(rules.tabStops.renderHTML({ tabStops: [] })).toEqual({ "data-tab-stops": "[]" });
    expect(rules.tabStops.parseHTML(elementWith({ "data-tab-stops": "[]" }))).toEqual([]);
    expect(rules.tabStops.parseHTML(elementWith({}))).toBeUndefined();
  });

  it("读不懂的 JSON 退回 undefined，而不是把半截清单当真", () => {
    for (const raw of ["not json", '{"a":1}', '[{"alignment":"right"}]', "[{\"positionPt\":-1}]"]) {
      expect(rules.tabStops.parseHTML(elementWith({ "data-tab-stops": raw }))).toBeUndefined();
    }
    // 没有 tabStops 时什么都不渲染，段落的 DOM 不会长出空属性。
    expect(rules.tabStops.renderHTML({})).toEqual({});
  });
});

describe("制表位清单的并入与移除", () => {
  const dot = { positionPt: 56.7, alignment: "right" as const, leader: "dot" as const };
  const center = { positionPt: 100, alignment: "center" as const };

  it("同位置整条替换，别的位置照旧按升序排", () => {
    expect(withTabStop([center], dot)).toEqual([dot, center]);
    // 只改前导符也是替换：改一次多一条，清单很快挤满重叠的制表位。
    expect(withTabStop([center, { ...dot, leader: "hyphen" as const }], dot)).toEqual([
      dot,
      center,
    ]);
  });

  it("没建模过的段落从空表开始，且并入不改动入参", () => {
    const stops = [center];
    expect(withTabStop(undefined, dot)).toEqual([dot]);
    expect(withTabStop(stops, dot)).toEqual([dot, center]);
    expect(stops).toEqual([center]);
  });

  it("移除只删那一位；没建模过的一律原样退回 undefined", () => {
    expect(withoutTabStop([dot, center], 56.7)).toEqual([center]);
    // 清光了是空表（导出删段落自己的 w:tabs），与"从没建模过"两回事。
    expect(withoutTabStop([dot], 56.7)).toEqual([]);
    expect(withoutTabStop(undefined, 56.7)).toBeUndefined();
  });

  it("按 0.1pt 认位：读侧与 cmToPt 都只给到这一格", () => {
    expect(sameTabStopPosition(56.7, 113.4 / 2)).toBe(true);
    expect(sameTabStopPosition(56.7, 56.8)).toBe(false);
  });
});
