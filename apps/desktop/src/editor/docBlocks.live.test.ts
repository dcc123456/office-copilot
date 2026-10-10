import { describe, expect, it } from "vitest";
import { getSchema, type CommandProps } from "@tiptap/core";
import { ptToCssLetterSpacing } from "@office-copilot/docx-io";
import type { Editor } from "@tiptap/react";
import type { Node as PMNode } from "@tiptap/pm/model";
import { EditorState, TextSelection } from "@tiptap/pm/state";
import { EDITOR_EXTENSIONS } from "./extensions/editorExtensions.js";
import { ParagraphFormatting } from "./extensions/ParagraphFormatting.js";
import { blocksOf, executeOps } from "./docBlocks.js";
import { planEdit, revert, addressOf, type Block, type CoreOp, type StylePatch } from "../ai/docEdit.js";
import { templateProposal, typographyProposal } from "./formatDocument.js";
import { templateById } from "./templates.js";

/**
 * The formatting pipeline exercised against the schema the app actually builds —
 * no DOM, but real ProseMirror nodes, so position bugs and mark flattening show
 * up here the way they used to show up in the user's document.
 */

const BODY_FONT = '"Microsoft YaHei", "PingFang SC", "微软雅黑", sans-serif';

const schema = getSchema(EDITOR_EXTENSIONS);

function harness(content: unknown): Editor {
  let state = EditorState.create({
    schema,
    doc: schema.nodeFromJSON({ type: "doc", content }),
  });
  return {
    get state() {
      return state;
    },
    view: {
      dispatch: (tr: Parameters<typeof state.apply>[0]) => {
        state = state.apply(tr);
      },
    },
  } as unknown as Editor;
}

const text = (t: string, marks?: unknown[]) => ({ type: "text", text: t, ...(marks ? { marks } : {}) });
const style = (attrs: Partial<Record<string, string | null>>) => [
  {
    type: "textStyle",
    attrs: {
      fontFamily: null,
      fontFamilyWest: null,
      fontSize: null,
      letterSpacing: null,
      color: null,
      backgroundColor: null,
      ...attrs,
    },
  },
];
const bold = () => ({ type: "bold" });
const italic = () => ({ type: "italic" });
const underline = () => ({ type: "underline" });

const para = (content: unknown[]) => ({ type: "paragraph", content });
const heading = (level: number, content: unknown[]) => ({ type: "heading", attrs: { level }, content });
const listItem = (content: unknown[]) => ({ type: "listItem", content: [para(content)] });
const bulletList = (...items: unknown[][]) => ({ type: "bulletList", content: items.map(listItem) });

/** Per-run view of the document: what the user actually sees character-wise. */
function charRuns(doc: PMNode): { text: string; attrs: Record<string, unknown>; bold: boolean }[] {
  const out: { text: string; attrs: Record<string, unknown>; bold: boolean }[] = [];
  doc.descendants((n) => {
    if (n.isText) {
      const ts = n.marks.find((m) => m.type.name === "textStyle");
      out.push({
        text: n.text ?? "",
        attrs: (ts?.attrs as Record<string, unknown>) ?? {},
        bold: n.marks.some((m) => m.type.name === "bold"),
      });
    }
    return true;
  });
  return out;
}

/** The mark names on every run containing `word`. */
function marksOfRun(editor: Editor, word: string): string[] {
  const out: string[] = [];
  editor.state.doc.descendants((n) => {
    if (n.isText && (n.text ?? "").includes(word)) out.push(...n.marks.map((m) => m.type.name));
    return true;
  });
  return out;
}

function run(editor: Editor, word: string) {
  const found = charRuns(editor.state.doc).find((r) => r.text.includes(word));
  if (!found) throw new Error(`no run containing ${word}`);
  return found;
}

/** Plan and execute exactly as the UI does, returning the live editor. */
function applyStyles(editor: Editor, patches: Array<{ index: number; style: StylePatch }>) {
  const blocks = blocksOf(editor.state.doc);
  const plan = planEdit(blocks, { summary: "t", ops: patches.map((p) => ({ op: "setStyle" as const, ...p })) });
  const exec = plan.ops.length ? executeOps(editor, plan.ops, plan.after) : null;
  return { plan, exec };
}

const mixedBoldDoc = [
  para([text("开头 "), text("重点", [bold()]), text(" 结尾")]),
];

/**
 * 用真实 schema 跑一条 ParagraphFormatting 命令。
 *
 * harness 的 editor 是手搓的最小对象，没有 Tiptap 的 `commands` / `chain` 装配，
 * 所以取扩展配置里那份命令实现本身，喂给它同样的 CommandProps，dispatch 用 harness 的 view。
 *
 * @param editor harness 出来的编辑器。
 * @param name 要跑的命令名。
 * @param args 该命令的入参。
 * @param words 要选中的文字：每处文字所在的块都会落进命令的选区。
 * @returns 命令的返回值（选区里有段落或标题才为 true）。
 */
function runFormatCommand(
  editor: Editor,
  name: "setParagraphFormat" | "applyNamedStyle",
  args: unknown[],
  words: string[],
): boolean {
  let from = -1;
  let to = -1;
  editor.state.doc.descendants((node, pos) => {
    if (!node.isText || !words.some((word) => (node.text ?? "").includes(word))) return;
    from = from < 0 ? pos : Math.min(from, pos);
    to = Math.max(to, pos + node.nodeSize);
  });
  if (from < 0) throw new Error(`no run matching ${words.join("/")}`);
  editor.view.dispatch(
    editor.state.tr.setSelection(TextSelection.create(editor.state.doc, from, to)),
  );
  const commands = (
    ParagraphFormatting.config.addCommands as (this: unknown) => Record<
      "setParagraphFormat" | "applyNamedStyle",
      (...args: unknown[]) => (props: CommandProps) => boolean
    >
  ).call({ options: {} });
  return commands[name](...args)({
    editor,
    state: editor.state,
    dispatch: editor.view.dispatch,
    view: editor.view,
    // 选区里没有段落或标题时的兜底路径，用例不依赖它。
    commands: { updateAttributes: () => false },
  } as unknown as CommandProps);
}

describe("reading run-level style", () => {
  it("records keys the runs disagree on as mixed, and keeps them out of style", () => {
    const editor = harness([
      para([text("普通正文 "), text("红字", style({ color: "#ff0000" }))]),
    ]);
    const b = blocksOf(editor.state.doc)[0]!;
    expect(b.mixed).toEqual(["color"]);
    expect(b.style.color).toBeUndefined();
  });

  it("reports a uniform character style as a block style", () => {
    const editor = harness([para([text("全部加粗", [bold()])])]);
    const b = blocksOf(editor.state.doc)[0]!;
    expect(b.mixed).toBeUndefined();
    expect(b.style.bold).toBe(true);
  });
});

describe("block-wide style writes", () => {
  it("unify the font without erasing a red word's colour", () => {
    const editor = harness([
      para([text("普通正文 "), text("红字", style({ color: "#ff0000" })), text(" 结尾")]),
    ]);
    const { exec } = applyStyles(editor, [{ index: 0, style: { fontSize: "14px", fontFamily: BODY_FONT } }]);
    expect(exec?.ok).toBe(true);
    const red = run(editor, "红字");
    expect(red.attrs.color).toBe("#ff0000");
    expect(red.attrs.fontSize).toBe("14px");
    expect(red.attrs.fontFamily).toBe(BODY_FONT);
    expect(run(editor, "普通正文").attrs.fontSize).toBe("14px");
  });

  it("keep highlighting when only the size changes", () => {
    const editor = harness([para([text("荧光", style({ backgroundColor: "#ffff00" }))])]);
    const { exec } = applyStyles(editor, [{ index: 0, style: { fontSize: "18px" } }]);
    expect(exec?.ok).toBe(true);
    expect(run(editor, "荧光").attrs).toMatchObject({ backgroundColor: "#ffff00", fontSize: "18px" });
  });

  it("keep partial bold when unifying size and font", () => {
    const editor = harness(mixedBoldDoc);
    const { exec } = applyStyles(editor, [{ index: 0, style: { fontSize: "14px", fontFamily: BODY_FONT } }]);
    expect(exec?.ok).toBe(true);
    expect(run(editor, "重点").bold).toBe(true);
    expect(run(editor, "开头").bold).toBe(false);
    expect(run(editor, "重点").attrs.fontSize).toBe("14px");
  });

  it("refuse to flatten a block whose runs disagree on bold", () => {
    const editor = harness(mixedBoldDoc);
    const original = editor.state.doc;
    const { plan, exec } = applyStyles(editor, [{ index: 0, style: { bold: true } }]);
    expect(plan.ops).toEqual([]);
    expect(plan.skipped.join()).toContain("加粗");
    expect(exec).toBeNull();
    expect(editor.state.doc.eq(original)).toBe(true);
  });

  it("refuse to flatten a block whose runs disagree on size", () => {
    const editor = harness([
      para([text("正文 "), text("小注", style({ fontSize: "12px" }))]),
    ]);
    const { plan } = applyStyles(editor, [{ index: 0, style: { fontSize: "14px" } }]);
    expect(plan.skipped.join()).toContain("字号");
  });

  it("refuse at the executor too, when ops bypass the planner", () => {
    const editor = harness(mixedBoldDoc);
    const original = editor.state.doc;
    const target = blocksOf(editor.state.doc)[0]!;
    const op: CoreOp = {
      kind: "setStyle",
      addr: { by: "blockText", text: target.text, nth: 0 },
      style: { bold: true },
    };
    const exec = executeOps(editor, [op]);
    expect(exec.ok).toBe(false);
    expect(exec.reason).toContain("抹平");
    expect(editor.state.doc.eq(original)).toBe(true);
  });

  it("set align and line height on a list item without breaking the list", () => {
    const editor = harness([bulletList([text("列表项")])]);
    const { plan, exec } = applyStyles(editor, [{ index: 0, style: { align: "center", lineHeight: "1.5" } }]);
    expect(plan.skipped).toEqual([]);
    expect(exec?.ok).toBe(true);
    expect(() => editor.state.doc.check()).not.toThrow();
    const list = editor.state.doc.firstChild!;
    expect(list.type.name).toBe("bulletList");
    expect(list.firstChild?.type.name).toBe("listItem");
    const p = list.firstChild!.firstChild!;
    expect(p.attrs.textAlign).toBe("center");
    expect(p.attrs.lineHeight).toBe("1.5");
  });

  it("patch an image paragraph's alignment without touching characters", () => {
    const editor = harness([
      { type: "paragraph", content: [{ type: "image", attrs: { src: "a.png", alt: "图" } }] },
    ]);
    const { exec } = applyStyles(editor, [{ index: 0, style: { align: "center" } }]);
    expect(exec?.ok).toBe(true);
    expect(editor.state.doc.firstChild?.attrs.textAlign).toBe("center");
    expect(editor.state.doc.firstChild?.type.name).toBe("paragraph");
  });
});

describe("段落对话框补齐的字段（真实 schema）", () => {
  it("右缩进与悬挂缩进写进段落，并能从 block.style 读回", () => {
    const editor = harness([para([text("正文")])]);
    const { exec } = applyStyles(editor, [
      { index: 0, style: { indent: 21, indentRight: 21, indentFirst: -21 } },
    ]);
    expect(exec?.ok).toBe(true);
    const p = editor.state.doc.firstChild!;
    expect(p.attrs.indentRight).toBe(21);
    // 负的 indentFirst 就是悬挂缩进，编辑器不做第二次单位换算
    expect(p.attrs.indentFirst).toBe(-21);
    expect(blocksOf(editor.state.doc)[0]!.style.indentRight).toBe(21);
  });

  it("分页开关与大纲级别活过 doc JSON 往返，false 不被当成未设置", () => {
    const node = schema.nodeFromJSON({
      type: "paragraph",
      attrs: {
        keepWithNext: true,
        keepLines: true,
        widowControl: false,
        snapToGrid: false,
        outlineLevel: 0,
      },
      content: [text("标题行")],
    })!;
    expect(node.attrs.widowControl).toBe(false);
    expect(node.attrs.outlineLevel).toBe(0);
    const again = schema.nodeFromJSON(node.toJSON())!;
    expect(again.attrs).toEqual(node.attrs);
    expect(() => again.check()).not.toThrow();
  });

  it("新建段落不带任何分页开关，套版不会凭空写出 w:keepNext", () => {
    const p = schema.nodeFromJSON({ type: "paragraph", content: [text("普通段落")] })!;
    for (const key of [
      "keepWithNext",
      "keepLines",
      "widowControl",
      "snapToGrid",
      "outlineLevel",
      "indentRight",
    ]) {
      expect(p.attrs[key]).toBeNull();
    }
    // styleId 反过来必须是"缺失"：null 在 docx-io 里是"用户清了样式"，
    // 新建段落若默认成 null，导出就会连源文档带着的引用一起删。
    expect(p.attrs.styleId).toBeUndefined();
    // tabStops 同一条理由只能是 undefined：[] 在 docx-io 里是"制表位全部清除"，
    // 拿它当默认值会把导入文档每个段落的 w:tabs 都删掉。
    expect(p.attrs.tabStops).toBeUndefined();
  });

  it("制表位清单活过 doc JSON 往返，位置仍是 pt", () => {
    const stops = [{ positionPt: 415.3, alignment: "right", leader: "dot" }];
    const node = schema.nodeFromJSON({
      type: "paragraph",
      attrs: { tabStops: stops },
      content: [text("第一章")],
    })!;
    expect(node.attrs.tabStops).toEqual(stops);
    const again = schema.nodeFromJSON(node.toJSON())!;
    expect(again.attrs.tabStops).toEqual(stops);
    expect(() => again.check()).not.toThrow();
  });
});

describe("具名样式引用 styleId（真实 schema）", () => {
  it("套用 / 清除 / 没建模过三种状态都能活过 doc JSON 往返", () => {
    const applied = schema.nodeFromJSON({
      type: "paragraph",
      attrs: { styleId: "af0" },
      content: [text("标题段")],
    })!;
    expect(applied.attrs.styleId).toBe("af0");
    expect(JSON.parse(JSON.stringify(applied.toJSON())).attrs.styleId).toBe("af0");

    const cleared = schema.nodeFromJSON({
      type: "paragraph",
      attrs: { styleId: null },
      content: [text("原为标题段")],
    })!;
    // null 必须真的留在 JSON 里：丢了它，"清掉样式"会被读成"没建模过"而把引用又写回去。
    expect(JSON.stringify(cleared.toJSON())).toContain('"styleId":null');

    const untouched = schema.nodeFromJSON({ type: "paragraph", content: [text("普通段落")] })!;
    expect(untouched.attrs.styleId).toBeUndefined();
  });

  it("套样式只补 styleId，段里已有的缩进、分页开关与红字都不受牵连", () => {
    const editor = harness([
      {
        type: "paragraph",
        attrs: { indentRight: 21, keepWithNext: true },
        content: [text("正文 "), text("红字", style({ color: "#ff0000" }))],
      },
    ]);
    const p = editor.state.doc.firstChild!;
    editor.view.dispatch(
      editor.state.tr.setNodeMarkup(0, undefined, { ...p.attrs, styleId: "af0" }),
    );
    const after = editor.state.doc.firstChild!;
    expect(after.attrs.styleId).toBe("af0");
    expect(after.attrs.indentRight).toBe(21);
    expect(after.attrs.keepWithNext).toBe(true);
    expect(charRuns(editor.state.doc).map((r) => [r.text, r.attrs.color])).toEqual([
      ["正文 ", undefined],
      ["红字", "#ff0000"],
    ]);
  });

  it("跨段一次改完：两段都写上，各自已有的属性各留各的", () => {
    const editor = harness([
      { type: "paragraph", attrs: { indentRight: 21 }, content: [text("甲段")] },
      { type: "paragraph", attrs: { keepWithNext: true }, content: [text("乙段")] },
    ]);
    expect(
      runFormatCommand(editor, "setParagraphFormat", [{ textAlign: "center" }], ["甲段", "乙段"]),
    ).toBe(true);
    const first = editor.state.doc.child(0);
    const second = editor.state.doc.child(1);
    // 逐块各发一个基于旧 doc 的事务时，第二段起抛 mismatched transaction，只有第一段改上。
    expect([first.attrs.textAlign, first.attrs.indentRight]).toEqual(["center", 21]);
    expect([second.attrs.textAlign, second.attrs.keepWithNext]).toEqual(["center", true]);
  });

  it("选区里的列表项也一起改，且不动列表结构", () => {
    const editor = harness([bulletList([text("列表项")])]);
    expect(runFormatCommand(editor, "setParagraphFormat", [{ lineHeight: "1.5" }], ["列表项"])).toBe(
      true,
    );
    const list = editor.state.doc.firstChild!;
    expect(list.type.name).toBe("bulletList");
    expect(list.firstChild!.firstChild!.attrs.lineHeight).toBe("1.5");
  });

  it("窗格套用标题样式：正文段变成对应级别的标题，红字与加粗都不受牵连", () => {
    const editor = harness([
      {
        type: "paragraph",
        attrs: { indentRight: 21, lineHeight: "1.5" },
        content: [text("第一章 "), text("重点", [bold()]), text(" 结尾")],
      },
    ]);
    expect(runFormatCommand(editor, "applyNamedStyle", ["2", 1], ["第一章"])).toBe(true);
    const h = editor.state.doc.firstChild!;
    expect([h.type.name, h.attrs.level, h.attrs.styleId]).toEqual(["heading", 1, "2"]);
    // 段落自己设过的排版属性是"直接格式"，样式引用不该把它们清掉。
    expect(h.attrs.indentRight).toBe(21);
    expect(h.attrs.lineHeight).toBe("1.5");
    expect(marksOfRun(editor, "重点")).toContain("bold");
  });

  it("窗格清除样式：标题退回正文并留下 null 引用，level 不残留", () => {
    const editor = harness([heading(2, [text("原二级标题")])]);
    expect(
      runFormatCommand(editor, "applyNamedStyle", [null, undefined], ["原二级标题"]),
    ).toBe(true);
    const p = editor.state.doc.firstChild!;
    expect(p.type.name).toBe("paragraph");
    expect(p.attrs.styleId).toBeNull();
    // 留着 heading 却没有标题样式引用，重新导入时这段会被读成正文，级别白丢。
    expect(p.attrs.level).toBeUndefined();
  });

  it("跨段套用一次改完：两段都换，各自的其它属性原样保留", () => {
    const editor = harness([
      { type: "paragraph", attrs: { indentRight: 21 }, content: [text("甲段")] },
      { type: "paragraph", attrs: { keepWithNext: true }, content: [text("乙段")] },
    ]);
    expect(
      runFormatCommand(editor, "applyNamedStyle", ["af8", 2], ["甲段", "乙段"]),
    ).toBe(true);
    const [first, second] = [editor.state.doc.child(0), editor.state.doc.child(1)];
    expect(first.type.name).toBe("heading");
    expect([first.attrs.level, first.attrs.styleId, first.attrs.indentRight]).toEqual([2, "af8", 21]);
    expect([second.attrs.level, second.attrs.styleId, second.attrs.keepWithNext]).toEqual([
      2,
      "af8",
      true,
    ]);
  });

  it("套样式不会碰列表结构：列表项只是多了引用与级别", () => {
    const editor = harness([bulletList([text("列表项")])]);
    expect(
      runFormatCommand(editor, "applyNamedStyle", ["af3", undefined], ["列表项"]),
    ).toBe(true);
    const list = editor.state.doc.firstChild!;
    expect(list.type.name).toBe("bulletList");
    const p = list.firstChild!.firstChild!;
    expect(p.type.name).toBe("paragraph");
    expect(p.attrs.styleId).toBe("af3");
  });
});

describe("automatic typesetting", () => {
  const mixedDocument = [
    heading(1, [text("一级标题")]),
    para([text("普通正文 "), text("红字", style({ color: "#ff0000" })), text(" 结尾")]),
    bulletList([text("列表项")]),
    heading(2, [text("二级标题")]),
  ];

  it("centres H1 and writes sizes as CSS px", () => {
    const editor = harness(mixedDocument);
    const patches = typographyProposal(blocksOf(editor.state.doc)).ops.flatMap((o) =>
      o.op === "setStyle" ? [{ index: o.index, style: o.style }] : [],
    );
    const { exec } = applyStyles(editor, patches);
    expect(exec?.ok).toBe(true);
    const h1 = editor.state.doc.firstChild!;
    expect(h1.attrs.textAlign).toBe("center");
    expect(run(editor, "一级标题").attrs.fontSize).toBe("22px");
    expect(run(editor, "红字").attrs.color).toBe("#ff0000");
    expect(run(editor, "列表项").attrs.fontSize).toBe("14px");
  });

  it("is idempotent on the live document", () => {
    const editor = harness(mixedDocument);
    const first = typographyProposal(blocksOf(editor.state.doc));
    expect(first.ops.length).toBeGreaterThan(0);
    const plan = planEdit(blocksOf(editor.state.doc), first);
    executeOps(editor, plan.ops, plan.after);
    const second = typographyProposal(blocksOf(editor.state.doc));
    expect(second.ops).toEqual([]);
    expect(second.summary).toBe("格式已符合规范");
  });

  it("undoes back to a document identical to the original", () => {
    const editor = harness(mixedDocument);
    const original = editor.state.doc;
    const blocks = blocksOf(editor.state.doc);
    const plan = planEdit(blocks, typographyProposal(blocks));
    const exec = executeOps(editor, plan.ops, plan.after);
    expect(exec.ok).toBe(true);
    expect(editor.state.doc.eq(original)).toBe(false);

    const back = revert(exec.after, plan.inverses);
    expect(back.failed).toEqual([]);
    const undo = executeOps(editor, back.applied, back.after);
    expect(undo.ok).toBe(true);
    expect(editor.state.doc.eq(original)).toBe(true);
  });

  it("leaves hand-set body text alone, and still unifies the rest", () => {
    const editor = harness([
      para([text("手改过的大小", style({ fontSize: "20px" }))]),
      para([text("普通正文")]),
    ]);
    const proposal = typographyProposal(blocksOf(editor.state.doc));
    expect(proposal.ops.flatMap((o) => (o.op === "setStyle" ? [o.index] : []))).toEqual([1]);
  });
});

describe("block snapshots stay consistent with the document", () => {
  it("keeps block types and text in outline order across lists, tables and images", () => {
    const editor = harness([
      heading(1, [text("标题")]),
      para([text("正文")]),
      { type: "bulletList", content: [listItem([text("甲")]), listItem([text("乙")])] },
      {
        type: "table",
        content: [
          {
            type: "tableRow",
            content: [{ type: "tableCell", content: [para([text("单元格")])] }],
          },
        ],
      },
    ]);
    const blocks: Block[] = blocksOf(editor.state.doc);
    expect(blocks.map((b) => [b.type, b.text])).toEqual([
      ["heading", "标题"],
      ["paragraph", "正文"],
      ["paragraph", "甲"],
      ["paragraph", "乙"],
      ["table", "单元格"],
    ]);
    expect(blocks.map((b) => b.index)).toEqual([0, 1, 2, 3, 4]);
  });
});

describe("书签锚点（真实 schema）", () => {
  const start = (id: string, name: string) => ({ type: "bookmarkStart", attrs: { id, name } });
  const end = (id: string) => ({ type: "bookmarkEnd", attrs: { id } });

  it("id 与 name 活过 doc JSON，且不占段落文字", () => {
    const editor = harness([para([start("7", "_Toc9"), text("第一章"), end("7")]), para([text("正文")])]);
    const doc = editor.state.doc;
    // 书签是区间标记不是内容：多算一个字符就会让 AI 的改动块定位错位。
    expect(doc.textContent).toBe("第一章正文");
    expect(blocksOf(doc).map((b) => b.text)).toEqual(["第一章", "正文"]);
    const anchor = doc.firstChild!.child(0);
    expect(anchor.type.name).toBe("bookmarkStart");
    expect(anchor.attrs).toMatchObject({ id: "7", name: "_Toc9" });
  });

  it("跨段配对：起点在上段、终点在下段都照样留在树里", () => {
    const doc = harness([para([start("0", "span"), text("上")]), para([text("下"), end("0")])]).state.doc;
    const json = doc.toJSON() as { content?: { content?: { type: string }[] }[] };
    expect(json.content?.[0]?.content?.[0]?.type).toBe("bookmarkStart");
    expect(json.content?.[1]?.content?.at(-1)?.type).toBe("bookmarkEnd");
  });

  it("只装着书签的段落不会塌成空段", () => {
    const doc = harness([para([start("1", "only"), end("1")]), para([text("正文")])]).state.doc;
    expect(doc.childCount).toBe(2);
    expect(doc.firstChild!.childCount).toBe(2);
  });
});

describe("域芯片（真实 schema）", () => {
  const field = (instr: string, display: string) => ({
    type: "field",
    attrs: { kind: instr.split(" ")[0]!, instr, display },
  });

  it("指令与含制表符的缓存结果活过 doc JSON", () => {
    const doc = harness([para([field(`TOC \\o "1-3" \\h`, "第一章\t3"), text("后面")])]).state.doc;
    const json = JSON.parse(JSON.stringify(doc.toJSON())) as {
      content: { content: { type: string; attrs: Record<string, unknown> }[] }[];
    };
    expect(json.content[0]!.content[0]!.attrs).toMatchObject({
      instr: `TOC \\o "1-3" \\h`,
      display: "第一章\t3",
    });
    // 缓存结果是芯片上的字，不是正文：算进字数就会让改动块定位错位。
    expect(doc.textContent).toBe("后面");
  });

  it("Word 还没算过结果的域仍是一个域节点，不会退化成文字", () => {
    const doc = harness([para([field("DATE", "")])]).state.doc;
    expect(doc.firstChild!.childCount).toBe(1);
    expect(doc.firstChild!.child(0).type.name).toBe("field");
  });
});

describe("跨段域的边界（真实 schema）", () => {
  const begin = (instr: string, isSeparated = true) => ({
    type: "fieldBegin",
    attrs: { kind: instr.split(" ")[0]!, instr, isSeparated },
  });
  const end = () => ({ type: "fieldEnd" });
  const field = (instr: string, display: string) => ({
    type: "field",
    attrs: { kind: instr.split(" ")[0]!, instr, display },
  });

  it("目录那种 begin 在首段、end 在末段的域，两段都留在树里且不占正文", () => {
    const doc = harness([
      para([begin(`TOC \\o "1-3"`), text("第一章\t"), field("PAGEREF _Toc1", "3")]),
      para([text("第二章\t"), field("PAGEREF _Toc2", "5"), end()]),
    ]).state.doc;
    expect(doc.textContent).toBe("第一章\t第二章\t");
    expect(doc.firstChild!.child(0).type.name).toBe("fieldBegin");
    expect(doc.lastChild!.child(2).type.name).toBe("fieldEnd");
    const json = JSON.parse(JSON.stringify(doc.toJSON())) as {
      content: { content: { type: string; attrs?: Record<string, unknown> }[] }[];
    };
    expect(json.content[0]!.content[0]!.attrs).toMatchObject({
      instr: `TOC \\o "1-3"`,
      isSeparated: true,
    });
  });

  it("Word 标成 dirty、还没写 separate 的目录开头也保得住", () => {
    const doc = harness([para([begin("TOC", false), text("旧目录第一条")])]).state.doc;
    expect(doc.firstChild!.child(0).attrs.isSeparated).toBe(false);
    // 边界不是文字：把它算进段落长度会让 AI 的改动块按错位置落地。
    expect(doc.textContent).toBe("旧目录第一条");
  });
});

describe("脚注与尾注的引用点（真实 schema）", () => {
  const fn = (id: string) => ({ type: "footnoteReference", attrs: { id } });
  const en = (id: string) => ({ type: "endnoteReference", attrs: { id } });

  it("id 活过 doc JSON，且不占正文：脚注文字在部件里，不在树上", () => {
    const doc = harness([para([text("正文甲"), fn("1")]), para([text("正文乙"), en("2")])]).state.doc;
    expect(doc.textContent).toBe("正文甲正文乙");
    expect(blocksOf(doc).map((b) => b.text)).toEqual(["正文甲", "正文乙"]);
    expect(doc.firstChild!.child(1).type.name).toBe("footnoteReference");
    const json = JSON.parse(JSON.stringify(doc.toJSON())) as {
      content: { content: { type: string; attrs?: Record<string, unknown> }[] }[];
    };
    expect(json.content[0]!.content[1]).toMatchObject({
      type: "footnoteReference",
      attrs: { id: "1" },
    });
    expect(json.content[1]!.content[1]).toMatchObject({
      type: "endnoteReference",
      attrs: { id: "2" },
    });
  });

  it("一段里两个脚注夹着文字，顺序与各自的 id 都不串位", () => {
    const doc = harness([para([text("前"), fn("3"), text("中"), fn("4"), text("后")])]).state.doc;
    expect(doc.firstChild!.childCount).toBe(5);
    expect(doc.firstChild!.child(1).attrs.id).toBe("3");
    expect(doc.firstChild!.child(3).attrs.id).toBe("4");
    expect(doc.textContent).toBe("前中后");
  });
});

describe("超链接与正文色", () => {
  const GONGWEN = templateById("gongwen")!;
  const linkMark = {
    type: "link",
    attrs: { href: "https://example.com", target: null },
  };
  const setStyleOps = (blocks: Block[]) =>
    templateProposal(GONGWEN, blocks)
      .ops.filter((o) => o.op === "setStyle")
      .map((o) => ({
        index: Number((o as { index: number }).index),
        style: (o as { style: StylePatch }).style,
      }));

  it("paints the body colour around a link without flattening the link", () => {
    // docx-io now hands a link run over with no colour of its own — Word's blue
    // is the document's to paint — so the paragraph is styleable again.
    const editor = harness([para([text("联系方式："), text("官网", [linkMark])])]);
    const blocks = blocksOf(editor.state.doc);
    expect(blocks[0]!.mixed ?? []).not.toContain("color");
    expect(setStyleOps(blocks).some((o) => o.style.color === "#000000")).toBe(true);
    applyStyles(editor, setStyleOps(blocks));
    expect(run(editor, "联系方式").attrs.color).toBe("#000000");
    expect(run(editor, "官网").attrs.color).toBeNull();
  });

  it("keeps a link the user coloured by hand", () => {
    const editor = harness([
      para([
        text("联系方式：", style({ color: "#262626" })),
        text("官网", [linkMark, ...style({ color: "#0563c1" })]),
      ]),
    ]);
    applyStyles(editor, setStyleOps(blocksOf(editor.state.doc)));
    // The prose takes 公文's ink; the link keeps the blue the user chose, and
    // the writer defers to it instead of injecting the template's.
    expect(run(editor, "联系方式").attrs.color).toBe("#000000");
    expect(run(editor, "官网").attrs.color).toBe("#0563c1");
  });

  it("has no colour opinion for a paragraph that is only a link", () => {
    const editor = harness([para([text("https://example.com", [linkMark])])]);
    const blocks = blocksOf(editor.state.doc);
    expect(blocks[0]!.mixed).toContain("color");
    for (const o of setStyleOps(blocks)) expect(o.style.color).toBeUndefined();
  });
});

describe("字符样式维度", () => {
  const GONGWEN = templateById("gongwen")!;
  const bodyOps = (blocks: Block[]) =>
    templateProposal(GONGWEN, blocks)
      .ops.filter((o) => o.op === "setStyle" && (o as { index: number }).index === 0)
      .map((o) => (o as { style: StylePatch }).style);

  it("gives 公文 runs their 西文字体 and 字间距", () => {
    const editor = harness([
      heading(1, [text("关于开展年度检查的通知")]),
      para([text("各有关单位：")]),
    ]);
    const ops = templateProposal(GONGWEN, blocksOf(editor.state.doc))
      .ops.filter((o) => o.op === "setStyle")
      .map((o) => ({ index: (o as { index: number }).index, style: (o as { style: StylePatch }).style }));
    expect(ops).toHaveLength(2);
    applyStyles(editor, ops);
    // 标题：黑体二号、字间距半磅；正文：仿宋三号。两处西文都是 Times。
    expect(run(editor, "关于开展").attrs).toMatchObject({
      fontFamily: "黑体",
      fontFamilyWest: "Times New Roman",
      letterSpacing: ptToCssLetterSpacing(0.5),
    });
    expect(run(editor, "各有关单位").attrs).toMatchObject({
      fontFamily: "仿宋",
      fontFamilyWest: "Times New Roman",
      color: "#000000",
      letterSpacing: null,
    });
  });

  it("keeps a half-underlined sentence underlined and styles the rest", () => {
    // The v1 protection generalised to every new character key: a block-wide
    // write may not flatten what the user marked by hand, but it still fixes
    // the font, the colour and the size of that same paragraph.
    const editor = harness([
      para([text("请参见"), text("附表三", [underline()]), text("办理。")]),
    ]);
    const blocks = blocksOf(editor.state.doc);
    expect(blocks[0]!.mixed).toContain("underline");
    const patch = bodyOps(blocks)[0]!;
    expect(patch.underline).toBeUndefined();
    expect(patch.fontFamily).toBe("仿宋");
    applyStyles(editor, [{ index: 0, style: patch }]);
    expect(marksOfRun(editor, "附表三")).toContain("underline");
    expect(marksOfRun(editor, "办理")).not.toContain("underline");
    expect(run(editor, "附表三").attrs.fontFamily).toBe("仿宋");
    expect(run(editor, "办理").attrs.color).toBe("#000000");
  });

  it("clears italic where 公文 forbids it, when the whole run is italic", () => {
    const editor = harness([para([text("斜体的正文", [italic()])])]);
    const blocks = blocksOf(editor.state.doc);
    expect(blocks[0]!.mixed ?? []).not.toContain("italic");
    expect(blocks[0]!.style.italic).toBe(true);
    applyStyles(editor, [{ index: 0, style: bodyOps(blocks)[0]! }]);
    expect(marksOfRun(editor, "斜体的正文")).not.toContain("italic");
  });
});

describe("书签 op 与整段 JSON 插入（真实 schema）", () => {
  const addrOf = (editor: Editor, index: number) => addressOf(blocksOf(editor.state.doc), index)!;

  it("给整段打书签：起点在文字前、终点在文字后，段落文字一个字节都不动", () => {
    const editor = harness([heading(1, [text("第一章")])]);
    const original = editor.state.doc.toJSON();
    const add: CoreOp[] = [{ kind: "addBookmark", addr: addrOf(editor, 0), id: "5", name: "_Toc5" }];
    expect(executeOps(editor, add).ok).toBe(true);
    const para = editor.state.doc.firstChild!;
    expect(para.childCount).toBe(3);
    // 断言只比 `type.name` 与 `attrs` 这些纯数据：PM 节点实例上的 `type` 是 NodeType，
    // 直接 toMatchObject 会让断言器顺着 schema 图递归，把测试进程吃到 OOM。
    expect(para.child(0).type.name).toBe("bookmarkStart");
    expect(para.child(0).attrs).toMatchObject({ id: "5", name: "_Toc5" });
    expect(para.child(2).type.name).toBe("bookmarkEnd");
    // 书签是原子行内节点，不占文字：AI 的改动块仍按"第一章"寻址。
    expect(editor.state.doc.textContent).toBe("第一章");

    const remove: CoreOp[] = [{ kind: "removeBookmark", addr: addrOf(editor, 0), id: "5", name: "_Toc5" }];
    expect(executeOps(editor, remove).ok).toBe(true);
    expect(editor.state.doc.toJSON()).toEqual(original);
  });

  it("撤一个不存在的书签就拒绝，且一个字节都不写", () => {
    const editor = harness([para([text("正文")])]);
    const before = editor.state.doc.toJSON();
    const r = executeOps(editor, [
      { kind: "removeBookmark", addr: addrOf(editor, 0), id: "9", name: "_Toc9" },
    ]);
    expect(r.ok).toBe(false);
    expect(editor.state.doc.toJSON()).toEqual(before);
  });

  it("带 json 的 insertBlock 建出真实结构：制表位、域芯片与文字各就各位", () => {
    const editor = harness([para([text("正文")])]);
    const entry = {
      type: "paragraph",
      attrs: { tabStops: [{ positionPt: 415.3, alignment: "right", leader: "dot" }] },
      content: [
        { type: "text", text: "第一章\t" },
        { type: "field", attrs: { kind: "PAGEREF", instr: "PAGEREF _Toc1 \\h", display: "3" } },
      ],
    };
    const r = executeOps(editor, [
      { kind: "insertBlock", anchor: { at: "start" }, block: { type: "paragraph", text: "第一章\t", json: entry } },
    ]);
    expect(r.ok).toBe(true);
    const inserted = editor.state.doc.firstChild!;
    expect(inserted.attrs.tabStops).toEqual([{ positionPt: 415.3, alignment: "right", leader: "dot" }]);
    expect(inserted.child(1).type.name).toBe("field");
    expect(editor.state.doc.childCount).toBe(2);
  });

  it("json 建不出段落或文字对不上时整块不写——宁可不写也不写假结构", () => {
    const editor = harness([para([text("正文")])]);
    const wrongText = {
      kind: "insertBlock",
      anchor: { at: "start" as const },
      block: {
        type: "paragraph",
        text: "对不上",
        json: { type: "paragraph", content: [{ type: "text", text: "第一章" }] },
      },
    } satisfies CoreOp;
    expect(executeOps(editor, [wrongText]).ok).toBe(false);
    const notABlock = {
      kind: "insertBlock",
      anchor: { at: "start" as const },
      block: { type: "paragraph", text: "", json: { type: "bookmarkStart", attrs: { id: "1" } } },
    } satisfies CoreOp;
    expect(executeOps(editor, [notABlock]).ok).toBe(false);
    expect(editor.state.doc.childCount).toBe(1);
  });
});
