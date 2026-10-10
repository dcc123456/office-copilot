import { describe, expect, it } from "vitest";
import { getSchema, type CommandProps } from "@tiptap/core";
import { pmToOoxml, type PMDoc } from "@office-copilot/docx-io";
import { history, undo } from "@tiptap/pm/history";
import type { Node as PMNode } from "@tiptap/pm/model";
import { EditorState, TextSelection, type Transaction } from "@tiptap/pm/state";
import { EDITOR_EXTENSIONS } from "./editorExtensions.js";
import { TrackChanges, trackChangesPlugin, trackKey, type TrackSignature } from "./TrackChanges.js";

/**
 * 修订开关（排名 4 ⑥）的行为。
 *
 * 本仓没有 DOM 测试环境，所以拿**真实 `EDITOR_EXTENSIONS` 的 schema** 建一个不带视图的
 * `EditorState`，把撤销栈和开关插件挂上去，再按编辑器的原样派发事务：打字、删除、回车合并段落
 * 在 ProseMirror 里都是事务，`appendTransaction` 看到的与浏览器里那条通道完全一致。
 */

const schema = getSchema(EDITOR_EXTENSIONS);

const text = (t: string, marks?: unknown[]) => ({ type: "text", text: t, ...(marks ? { marks } : {}) });
const para = (content: unknown[]) => ({ type: "paragraph", content });
/** 前人留下的一次修订（署名与本轮不同，用来试"不并进别人那一条"）。 */
const othersRevision = (kind: "ins" | "del", id: string) => ({
  type: "revision",
  attrs: { kind, id, author: "李四", date: "2026-01-01T00:00:00Z" },
});

/** 本轮修订的署名：与工具栏那次调用同形。 */
const SIGNATURE: TrackSignature = { author: "张三", date: "2026-10-10T22:00:00Z" };

/**
 * 一个最小工作台：真实 schema + 撤销栈 + 修订开关。
 *
 * @param blocks 段内容（JSON），默认一段"开头 / 要删的 / 结尾"，位置 4..7 正好是"要删的"。
 * @returns 读状态、开关修订、打字 / 删除 / 加粗 / 合并段落 / 撤销 / 放光标的入口。
 */
function harness(blocks: unknown[] = [para([text("开头 "), text("要删的"), text(" 结尾")])]) {
  let state = EditorState.create({
    schema,
    doc: schema.nodeFromJSON({ type: "doc", content: blocks } as never),
    plugins: [history(), trackChangesPlugin()],
  });
  const dispatch = (tr: Transaction) => {
    state = state.apply(tr);
  };
  return {
    get state() {
      return state;
    },
    /** 走扩展自己的命令，不直接 setMeta：命令层与插件层一起被测到。 */
    track(on: boolean, signature: TrackSignature = SIGNATURE) {
      const commands = (
        TrackChanges as unknown as {
          config: { addCommands: () => Record<string, (...args: unknown[]) => (props: CommandProps) => boolean> };
        }
      ).config.addCommands.call({ options: {}, name: "trackChanges" });
      const command = on ? commands.enableTrackChanges : commands.disableTrackChanges;
      if (!command) throw new Error(`修订扩展里没有 ${on ? "enable" : "disable"} 命令`);
      const tr = state.tr;
      // 与 Tiptap 的 CommandManager 同一套约定：命令只改 `tr`（`dispatch` 是个空函数，
      // 只用来告诉命令"这次会落地"），真正派发的是这里。
      if (command(on ? signature : undefined)({ state, tr, dispatch: () => undefined } as unknown as CommandProps)) {
        dispatch(tr);
      }
    },
    /** 打字：把 `through..at` 之间的字换成 `words`（`through` 省略时就是纯插入）。 */
    type(at: number, words: string, through?: number) {
      dispatch(state.tr.insertText(words, at, through ?? at));
    },
    /** 选中一段按 Delete。 */
    remove(from: number, to: number) {
      dispatch(state.tr.delete(from, to));
    },
    /** 给一段文字加粗（格式改动）。 */
    emphasize(from: number, to: number) {
      const bold = schema.marks.bold;
      if (!bold) throw new Error("schema 里没有 bold");
      dispatch(state.tr.addMark(from, to, bold.create()));
    },
    /** 在 `at` 处把前后两段并起来（段首 Backspace）。 */
    joinAt(at: number) {
      dispatch(state.tr.join(at));
    },
    /** ⌘Z。 */
    back() {
      undo(state, dispatch);
    },
    /** 点一下把光标放到某处（删除的落点要跟着走）。 */
    focus(at: number) {
      dispatch(state.tr.setSelection(TextSelection.create(state.doc, at)));
    },
  };
}

/**
 * 文档里带修订标记的文字节点，按文档顺序。
 *
 * @param doc 当前文档。
 * @returns 每段的 `text/kind/id/author`（`author` 为 null 表示文档里没署名）。
 */
function revisedRuns(doc: PMNode): { text: string; kind: string; id: string; author: string | null }[] {
  const out: { text: string; kind: string; id: string; author: string | null }[] = [];
  doc.descendants((node) => {
    const mark = node.marks.find((m) => m.type.name === "revision");
    if (mark && node.isText) {
      out.push({
        text: node.text ?? "",
        kind: String(mark.attrs.kind),
        id: String(mark.attrs.id),
        author: (mark.attrs.author as string | null) ?? null,
      });
    }
    return true;
  });
  return out;
}

describe("修订开关", () => {
  it("关着时普通编辑什么都不记", () => {
    const h = harness();
    h.type(3, "甲");
    expect(revisedRuns(h.state.doc)).toEqual([]);
    expect(h.state.doc.textBetween(1, 4)).toBe("开头甲");
  });

  it("开着时打的字补成插入，署名跟开关那一刻走", () => {
    const h = harness();
    h.track(true);
    h.type(3, "甲");
    expect(revisedRuns(h.state.doc)).toEqual([{ text: "甲", kind: "ins", id: "1", author: "张三" }]);
    expect(trackKey.getState(h.state)?.date).toBe(SIGNATURE.date);
  });

  it("连续几次输入并进同一条修订：Word 也是一次打字一条 w:ins", () => {
    const h = harness();
    h.track(true);
    h.type(3, "甲");
    h.type(4, "乙");
    h.type(5, "丙");
    const runs = revisedRuns(h.state.doc);
    // 连着打的字标记一模一样，PM 顺手把相邻节点并成一个 run——正对应 Word 里一条 `w:ins` 包一个 run。
    expect(runs.map((r) => r.text)).toEqual(["甲乙丙"]);
    expect(runs.map((r) => r.id)).toEqual(["1"]);
  });

  it("删掉的文字插回原位并带删除标记，光标留在删除线右边", () => {
    const h = harness();
    h.focus(4);
    h.track(true);
    h.remove(4, 7);
    expect(revisedRuns(h.state.doc)).toEqual([{ text: "要删的", kind: "del", id: "1", author: "张三" }]);
    expect(h.state.doc.textBetween(1, 10)).toBe("开头 要删的 结尾");
    expect(h.state.selection.from).toBe(7);
  });

  it("选中一段直接打字：旧的在前带 del、新的在后带 ins，各占一个编号", () => {
    const h = harness();
    h.track(true);
    h.type(4, "新的", 7);
    expect(revisedRuns(h.state.doc)).toEqual([
      { text: "要删的", kind: "del", id: "2", author: "张三" },
      { text: "新的", kind: "ins", id: "1", author: "张三" },
    ]);
  });

  it("这次刚插入的字又被删掉：两次改动互相抵消，不留删除线", () => {
    const h = harness();
    h.track(true);
    h.type(3, "甲");
    h.remove(3, 4);
    expect(revisedRuns(h.state.doc)).toEqual([]);
    expect(h.state.doc.textBetween(1, 3)).toBe("开头");
  });

  it("别人插入的字被我删掉：外层换成本次删除，不叠两层标记", () => {
    const h = harness([para([text("开头 "), text("旧见", [othersRevision("ins", "9")]), text(" 结尾")])]);
    h.track(true);
    h.remove(4, 6);
    expect(revisedRuns(h.state.doc)).toEqual([{ text: "旧见", kind: "del", id: "1", author: "张三" }]);
  });

  it("撤销不是新修订：回到改动前，不再多出一条删除", () => {
    const h = harness();
    h.track(true);
    h.type(3, "甲");
    h.back();
    expect(h.state.doc.textBetween(1, 3)).toBe("开头");
    expect(revisedRuns(h.state.doc)).toEqual([]);
  });

  it("改格式不记修订（w:rPrChange 没建模，见规格排名 4 判据）", () => {
    const h = harness();
    h.track(true);
    h.emphasize(4, 7);
    expect(revisedRuns(h.state.doc)).toEqual([]);
    expect(h.state.doc.resolve(5).nodeBefore?.marks.some((m) => m.type.name === "bold")).toBe(true);
  });

  it("段首 Backspace 合并两段：段落标记的修订没建模，一个字都不记", () => {
    const h = harness([para([text("上一段")]), para([text("下一段")])]);
    h.track(true);
    h.joinAt(5);
    expect(revisedRuns(h.state.doc)).toEqual([]);
    expect(h.state.doc.childCount).toBe(1);
  });

  it("关掉开关以后的编辑不再记，已经记下的留在文档里", () => {
    const h = harness();
    h.track(true);
    h.type(3, "甲");
    h.track(false);
    h.type(4, "乙");
    expect(revisedRuns(h.state.doc)).toEqual([{ text: "甲", kind: "ins", id: "1", author: "张三" }]);
  });

  it("文档里已有修订编号 5：新修订从 6 开始，不撞车", () => {
    const h = harness([para([text("前人改的", [othersRevision("ins", "5")])])]);
    h.track(true);
    h.type(1, "甲");
    expect(revisedRuns(h.state.doc)).toEqual([
      { text: "甲", kind: "ins", id: "6", author: "张三" },
      { text: "前人改的", kind: "ins", id: "5", author: "李四" },
    ]);
  });

  it("没填作者就留空：CT_TrackChange 的 w:author 是可选属性，不编名字", () => {
    const h = harness();
    h.track(true, { author: null, date: null });
    h.type(3, "甲");
    expect(revisedRuns(h.state.doc)).toEqual([{ text: "甲", kind: "ins", id: "1", author: null }]);
  });

  it("开出来的修订写回 OOXML：w:ins 与 w:del/w:delText 都在", () => {
    const h = harness();
    h.track(true);
    h.type(3, "甲");
    h.focus(5);
    h.remove(4, 7);
    const xml = pmToOoxml(h.state.doc.toJSON() as PMDoc);
    expect(xml).toContain("<w:ins");
    expect(xml).toContain("<w:del");
    expect(xml).toContain("<w:delText");
    expect(xml).toContain('w:author="张三"');
  });
});
