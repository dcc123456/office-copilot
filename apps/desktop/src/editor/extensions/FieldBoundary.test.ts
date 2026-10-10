import { describe, expect, it } from "vitest";
import { FieldBegin, FieldEnd } from "./FieldBoundary.js";

/**
 * 边界节点的 DOM 读写：复制粘贴一段目录时，begin 上的指令与"有没有 separate"
 * 必须跟着走，否则粘出来的域会丢掉指令（Word 打开就是一串没有域的普通文字）。
 * 签名取自 @tiptap/core dist/index.js 的真实调用点：属性级 parseHTML 收元素，
 * 节点级 renderHTML 收 `{ node, HTMLAttributes }`。
 */

type AttrRule = {
  default: unknown;
  parseHTML: (el: {
    getAttribute: (name: string) => string | null;
    hasAttribute: (name: string) => boolean;
  }) => unknown;
};

type NodeRender = (arg: {
  HTMLAttributes: Record<string, unknown>;
  node: { attrs: Record<string, unknown> };
}) => [string, Record<string, string>, ...unknown[]];

const beginAttrs = FieldBegin.config.addAttributes!.call(FieldBegin as never) as Record<
  string,
  AttrRule
>;
const renderBegin = FieldBegin.config.renderHTML as unknown as NodeRender;
const renderEnd = FieldEnd.config.renderHTML as unknown as NodeRender;

/** 只实现这两个方法的元素替身：属性规则只用得到它们。 */
const elementWith = (values: Record<string, string>) => ({
  getAttribute: (name: string) => values[name] ?? null,
  hasAttribute: (name: string) => name in values,
});

describe("跨段域边界的 DOM 读写", () => {
  it("起点把指令与 separate 记号写进 data-*，终点只留结束记号", () => {
    const [, beginHtml] = renderBegin({
      HTMLAttributes: {},
      node: { attrs: { kind: "TOC", instr: `TOC \\o "1-3"`, isSeparated: true } },
    });
    expect(beginHtml).toMatchObject({
      "data-field-boundary": "begin",
      "data-field-kind": "TOC",
      "data-instr": `TOC \\o "1-3"`,
      class: "pm-field-boundary pm-field-boundary-begin",
    });
    expect(beginHtml).toHaveProperty("data-separated");
    expect(Object.keys(beginHtml)).not.toContain("instr");

    const [, endHtml] = renderEnd({ HTMLAttributes: {}, node: { attrs: {} } });
    expect(endHtml).toMatchObject({ "data-field-boundary": "end" });
    expect(Object.keys(endHtml)).not.toContain("data-instr");
  });

  it("dirty 的域（只到指令就断了）写 data-dirty，读回来还是 isSeparated=false", () => {
    const [, html] = renderBegin({
      HTMLAttributes: {},
      node: { attrs: { kind: "TOC", instr: "TOC", isSeparated: false } },
    });
    expect(html).toHaveProperty("data-dirty");
    expect(html).not.toHaveProperty("data-separated");
    expect(beginAttrs.isSeparated!.parseHTML(elementWith({ "data-dirty": "" }))).toBe(false);
    expect(beginAttrs.isSeparated!.parseHTML(elementWith({ "data-separated": "" }))).toBe(true);
  });

  it("终点节点没有起点属性：读普通元素时全部回落默认值", () => {
    // 缺省必须是 undefined 而不是 false，否则写回会凭空造出一个 separate 记号。
    const el = elementWith({});
    expect(beginAttrs.instr!.parseHTML(el)).toBeUndefined();
    expect(beginAttrs.isSeparated!.parseHTML(el)).toBeUndefined();
  });
});
