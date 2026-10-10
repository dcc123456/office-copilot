import { describe, expect, it } from "vitest";
import { getSchema, type CommandProps } from "@tiptap/core";
import { EditorState, TextSelection } from "@tiptap/pm/state";
import { EDITOR_EXTENSIONS } from "./editorExtensions.js";
import { VertAlign } from "./VertAlign.js";

/**
 * 上标 / 下标标记的 DOM 读写与命令语义，跑在 app 真正装配的 schema 上（AGENTS §5.3）。
 *
 * 两个回调直接调，不引 jsdom（本仓没有 DOM 测试环境）；命令需要一个 `editor.isActive`，
 * 这里按 `@tiptap/core` 的 `isMarkActive` 语义给一个固定答案——它拿 `objectIncludes(attrs,
 * attributes, { strict: false })` 比对属性，所以"同一档再点取消、另一档点是换过去"确实分得开。
 */

const schema = getSchema(EDITOR_EXTENSIONS);

/**
 * 扩展的 `this`：Tiptap 在装配时把标记自身当上下文传进这些回调，命令实现里读的是
 * `this.name`，renderHTML 读 `this.options`。手工调用时必须把这两样喂全，否则
 * `setMark(this.name)` 变成 `setMark(undefined)`，事务里一个标记都不写。
 */
const EXTENSION_CTX = { name: "vertAlign", options: { HTMLAttributes: {} } };

type CommandName = "setVertAlign" | "toggleVertAlign" | "unsetVertAlign";

/**
 * 建一份只有一个段落的最小文档，选区盖住整段文字。
 *
 * @param text 段落文字。
 * @param marks 这段文字一开始带不带标记（JSON 形式，交给真实 schema 校验）。
 * @returns `run` 跑本扩展的命令、`alignVals` 读回每个文字节点的档位。
 */
function harness(text: string, marks?: unknown[]) {
  let state = EditorState.create({
    schema,
    doc: schema.nodeFromJSON({
      type: "doc",
      content: [{ type: "paragraph", content: [{ type: "text", text, ...(marks ? { marks } : {}) }] }],
    }),
  });
  state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 1, 1 + text.length)));

  // 只实现标记用到的三下：事务攒在链上，run() 一次性派发，与 Tiptap 的 chain 同形。
  const chainOf = () => {
    const tr = state.tr;
    const api = {
      setMark: (name: string, attrs?: Record<string, unknown>) => {
        const type = schema.marks[name];
        if (type) tr.addMark(tr.selection.from, tr.selection.to, type.create(attrs));
        return api;
      },
      unsetMark: (name: string) => {
        const type = schema.marks[name];
        if (type) tr.removeMark(tr.selection.from, tr.selection.to, type);
        return api;
      },
      run: () => {
        state = state.apply(tr);
        return true;
      },
    };
    return api;
  };

  const commands = (
    VertAlign.config.addCommands as unknown as (this: unknown) => Record<
      CommandName,
      (arg: { val: string }) => (props: CommandProps) => boolean
    >
  ).call(EXTENSION_CTX);

  /**
   * 跑一条命令。
   *
   * @param name 命令名。
   * @param val 档位，`unsetVertAlign` 不需要。
   * @param isActive 光标处已经是这一档了吗（喂给 `toggleVertAlign` 的那个判断）。
   */
  const run = (name: CommandName, val?: string, isActive = false): boolean =>
    commands[name]({ val: val as never })({
      editor: { isActive: () => isActive },
      state,
      dispatch: (tr: Parameters<typeof state.apply>[0]) => {
        state = state.apply(tr);
      },
      view: { dispatch: (tr: Parameters<typeof state.apply>[0]) => (state = state.apply(tr)) },
      commands: {},
      chain: () => chainOf() as never,
    } as unknown as CommandProps);

  /** 每个文字节点上的档位，没标记的是 undefined——一个 run 只该有一个 vertAlign。 */
  const alignVals = (): (string | undefined)[] => {
    const out: (string | undefined)[] = [];
    state.doc.descendants((node) => {
      if (!node.isText) return true;
      const hits = node.marks.filter((m) => m.type.name === "vertAlign");
      for (const [i, mark] of hits.entries()) out.push(i === 0 ? String(mark.attrs.val) : "重复");
      if (!hits.length) out.push(undefined);
      return true;
    });
    return out;
  };

  return { run, alignVals };
}

const attrRules = (
  VertAlign.config.addAttributes as unknown as () => {
    val: { default: unknown; rendered?: boolean };
  }
)();

const parseRules = (
  VertAlign.config.parseHTML as unknown as (this: unknown) => {
    tag: string;
    getAttrs: (el: unknown) => unknown;
  }[]
).call(EXTENSION_CTX);

const render = (
  VertAlign.config.renderHTML as unknown as (arg: {
    mark?: { attrs: Record<string, unknown> };
    HTMLAttributes: Record<string, unknown>;
  }) => [string, Record<string, unknown>, number]
).bind(EXTENSION_CTX);

describe("上标 / 下标标记的 DOM 读写", () => {
  it("缺省档位是上标", () => {
    expect(attrRules.val.default).toBe("superscript");
  });

  it("<sup> 与 <sub> 各认成一档", () => {
    const sup = parseRules.find((rule) => rule.tag === "sup");
    const sub = parseRules.find((rule) => rule.tag === "sub");
    expect(sup?.getAttrs({ tagName: "SUP" })).toEqual({ val: "superscript" });
    expect(sub?.getAttrs({ tagName: "SUB" })).toEqual({ val: "subscript" });
  });

  it("渲染按档位选标签，认不出的一律走上标，且不往 DOM 挂 val 属性", () => {
    expect(render({ mark: { attrs: { val: "subscript" } }, HTMLAttributes: {} })[0]).toBe("sub");
    expect(render({ mark: { attrs: { val: "superscript" } }, HTMLAttributes: {} })[0]).toBe("sup");
    expect(render({ HTMLAttributes: {} })[0]).toBe("sup");
    // 复制出去的 HTML 只有标准标签：`val` 是 rendered: false 的属性，不该出现在 DOM 上。
    expect(render({ mark: { attrs: { val: "subscript" } }, HTMLAttributes: {} })[1]).not.toHaveProperty("val");
    expect(attrRules.val.rendered).toBe(false);
  });
});

describe("上标 / 下标命令（真实 schema）", () => {
  it("setVertAlign 写的档位就是 attrs 里那个键", () => {
    const h = harness("mc2");
    expect(h.run("setVertAlign", "subscript")).toBe(true);
    expect(h.alignVals()).toEqual(["subscript"]);
  });

  it("同一档再点是取消，另一档再点是换过去", () => {
    const off = harness("mc2");
    expect(off.run("toggleVertAlign", "superscript", false)).toBe(true);
    expect(off.alignVals()).toEqual(["superscript"]);
    const on = harness("mc2", [{ type: "vertAlign", attrs: { val: "superscript" } }]);
    expect(on.run("toggleVertAlign", "superscript", true)).toBe(true);
    expect(on.alignVals()).toEqual([undefined]);
  });

  it("上标与下标共用一个标记：换档只剩一条，不会又上又下", () => {
    const h = harness("H2O", [{ type: "vertAlign", attrs: { val: "subscript" } }]);
    expect(h.run("toggleVertAlign", "superscript", false)).toBe(true);
    expect(h.alignVals()).toEqual(["superscript"]);
  });

  it("unsetVertAlign 两档一起清掉", () => {
    const h = harness("H2O", [{ type: "vertAlign", attrs: { val: "subscript" } }]);
    expect(h.run("unsetVertAlign")).toBe(true);
    expect(h.alignVals()).toEqual([undefined]);
  });
});
