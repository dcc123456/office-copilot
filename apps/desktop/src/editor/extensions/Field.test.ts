import { describe, expect, it } from "vitest";
import { Field } from "./Field.js";

/**
 * 域节点的属性读写规则。复制粘贴走的是 PM 的 DOMSerializer / DOMParser，
 * 它只认属性自己写的 renderHTML / parseHTML——没有 parseHTML 的属性在粘贴后
 * 一律回到 default，一个 `TOC \o "1-3"` 就会变成 PAGE 域。
 * 这里直接调这两个回调，不引 jsdom（本仓没有 DOM 测试环境）。
 *
 * 两个签名都取自 @tiptap/core dist/index.js 的实际调用点：
 * 属性级 `renderHTML(nodeOrMark.attrs)`（首参是 attrs 全体）、
 * `parseHTML(node)`（首参是 DOM 元素）。写成 `{ node }` 解构会在渲染时
 * 抛 TypeError，把整个编辑器炸成白屏。
 */

type AttrRule = {
  default: unknown;
  parseHTML: (el: { getAttribute: (name: string) => string | null }) => string | undefined;
  renderHTML: (attrs: Record<string, unknown>) => Record<string, unknown>;
};

const rules = (
  Field.config.addAttributes as unknown as () => {
    kind: AttrRule;
    instr: AttrRule;
    display: AttrRule;
  }
)();

const render = Field.config.renderHTML as unknown as (
  arg: { HTMLAttributes: Record<string, unknown>; node: { attrs: Record<string, unknown> } },
) => [string, Record<string, unknown>, string];

/** 只实现 getAttribute 的元素替身：parseHTML 规则只用得到它。 */
const elementWith = (values: Record<string, string>) => ({
  getAttribute: (name: string) => values[name] ?? null,
});

/** 复现 PM 的合成步骤：属性级 renderHTML 的返回值并成节点的 HTMLAttributes。 */
const htmlAttributesOf = (attrs: Record<string, unknown>): Record<string, unknown> =>
  [rules.kind, rules.instr, rules.display]
    .map((rule) => rule.renderHTML(attrs))
    .reduce((all, part) => ({ ...all, ...part }), {});

describe("域芯片的 DOM 读写", () => {
  it("指令与含制表符的缓存结果能从 DOM 里读回来", () => {
    const el = elementWith({
      "data-field": "TOC",
      "data-instr": `TOC \\o "1-3" \\h`,
      "data-display": "第一章\t3",
    });
    expect(rules.kind.parseHTML(el)).toBe("TOC");
    expect(rules.instr.parseHTML(el)).toBe(`TOC \\o "1-3" \\h`);
    expect(rules.display.parseHTML(el)).toBe("第一章\t3");
  });

  it("读不到 data-* 时返回 undefined，让 PM 回落到默认值", () => {
    // 返回 null 会被 PM 当成"值就是 null"存进 attrs，把默认值覆盖掉。
    const el = elementWith({});
    expect(rules.instr.parseHTML(el)).toBeUndefined();
    expect(rules.display.parseHTML(el)).toBeUndefined();
  });

  it("写出的是 data-* 属性，不是裸属性名", () => {
    const attrs = { kind: "DATE", instr: "DATE", display: "2026-10-10" };
    expect(htmlAttributesOf(attrs)).toEqual({
      "data-field": "DATE",
      "data-instr": "DATE",
      "data-display": "2026-10-10",
    });

    const [, html] = render({ HTMLAttributes: htmlAttributesOf(attrs), node: { attrs } });
    expect(html).toMatchObject({ class: "pm-field", "data-instr": "DATE", title: "DATE" });
    expect(Object.keys(html)).not.toContain("instr");
    expect(Object.keys(html)).not.toContain("kind");
  });

  it("Word 还没算过结果的域，芯片上显示指令而不是空白", () => {
    const attrs = { kind: "TOC", instr: `TOC \\o "1-3"`, display: "" };
    const [, , text] = render({ HTMLAttributes: htmlAttributesOf(attrs), node: { attrs } });
    expect(text).toBe("TOC");
  });
});
