import { describe, expect, it } from "vitest";
import { getSchema, type CommandProps } from "@tiptap/core";
import { ptToCssLetterSpacing, ooxmlToPm, pmToOoxml } from "@office-copilot/docx-io";
import type { PMDoc } from "@office-copilot/docx-io";
import type { Editor } from "@tiptap/react";
import type { Node as PMNode } from "@tiptap/pm/model";
import { EditorState, TextSelection } from "@tiptap/pm/state";
import { EDITOR_EXTENSIONS } from "./extensions/editorExtensions.js";
import { CommentCommands, commentAnchorRanges } from "./extensions/Comment.js";
import { footnoteRefKey, indexFootnoteRefs } from "./extensions/FootnoteRef.js";
import { ParagraphFormatting } from "./extensions/ParagraphFormatting.js";
import { Revision, selectionTouchesRevision } from "./extensions/Revision.js";
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
const vertAlign = (val: "superscript" | "subscript") => ({ type: "vertAlign", attrs: { val } });

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
 * 把选区放到指定文字上：harness 的 editor 是手搓的最小对象，没有 Tiptap 的 `commands`，
 * 只能自己派发一个设选区的事务。
 *
 * @param editor harness 出来的编辑器。
 * @param words 要选中的文字，所有匹配处一起进选区。
 * @param mode `range` 拉整段文字（默认），`cursor` 只把光标停在第一处前面。
 * @throws 文档里找不到要选的文字——是用例写错了，不该静默通过。
 */
function selectWords(editor: Editor, words: string[], mode: "range" | "cursor" = "range"): void {
  let from = -1;
  let to = -1;
  editor.state.doc.descendants((node, pos) => {
    if (!node.isText || !words.some((word) => (node.text ?? "").includes(word))) return;
    from = from < 0 ? pos : Math.min(from, pos);
    to = Math.max(to, pos + node.nodeSize);
  });
  if (from < 0) throw new Error(`no run matching ${words.join("/")}`);
  editor.view.dispatch(
    editor.state.tr.setSelection(
      mode === "cursor"
        ? TextSelection.create(editor.state.doc, from)
        : TextSelection.create(editor.state.doc, from, to),
    ),
  );
}

/**
 * 用真实 schema 跑一条扩展自己的命令。
 *
 * harness 的 editor 是手搓的最小对象，没有 Tiptap 的 `commands` / `chain` 装配，
 * 所以取扩展配置里那份命令实现本身，喂给它同样的 CommandProps，dispatch 用 harness 的 view。
 *
 * @param editor harness 出来的编辑器。
 * @param name 要跑的命令名。
 * @param args 该命令的入参。
 * @param words 要选中的文字：每处文字所在的块都会落进命令的选区。
 * @param source 命令所在的扩展，默认段落格式那一个（标记类的命令传自己的扩展）。
 * @param mode 选区形状：`range` 拉整段文字（默认），`cursor` 只把光标停在第一处前面
 *   ——面板上的按钮就是这种情况，用户通常没有拉选区。
 * @returns 命令的返回值（选区里能落这个格式才为 true）。
 * @throws 选区里找不到要选的文字，或那份扩展根本没有这条命令——是用例写错了，不该静默通过。
 */
function runFormatCommand(
  editor: Editor,
  name: string,
  args: unknown[],
  words: string[],
  source: { config: { addCommands?: () => Record<string, unknown> } } =
    ParagraphFormatting as unknown as {
      config: { addCommands?: () => Record<string, unknown> };
    },
  mode: "range" | "cursor" = "range",
): boolean {
  selectWords(editor, words, mode);
  const commands = (
    source.config.addCommands as (this: unknown) => Record<
      string,
      (...args: unknown[]) => (props: CommandProps) => boolean
    >
  ).call({
    options: {},
    // 标记类命令实现里读 `this.name`（Tiptap 装配时把扩展自身当上下文），手搓 ctx 要给它。
    name: (source as unknown as { name?: string }).name,
  });
  const command = commands[name];
  if (!command) throw new Error(`扩展里没有命令 ${name}`);
  return command(...args)({
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
    // dropCap 也一样：null 是"用户点了无"，默认成它会把源文档的 w:framePr 全删。
    expect(p.attrs.dropCap).toBeUndefined();
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

describe("制表位命令（真实 schema）", () => {
  /** Word 目录行的右页边制表位：位置是 docx-io `cmToPt` 给的 0.1pt 精度。 */
  const rightDot = { positionPt: 56.7, alignment: "right", leader: "dot" };
  const paraWith = (tabStops: unknown, word: string) => ({
    type: "paragraph",
    attrs: { tabStops },
    content: [text(word)],
  });

  it("「设置」只动这一个位置，别段落在别处的制表位原样留着", () => {
    // 面板上那张清单是从第一段读出来的，整份覆盖会把第二段 100pt 那条抹掉。
    const editor = harness([para([text("目录行")]), paraWith([{ positionPt: 100, alignment: "center" }], "签名行")]);
    expect(runFormatCommand(editor, "setTabStop", [rightDot], ["目录行", "签名行"])).toBe(true);
    expect(editor.state.doc.child(0).attrs.tabStops).toEqual([rightDot]);
    expect(editor.state.doc.child(1).attrs.tabStops).toEqual([
      rightDot,
      { positionPt: 100, alignment: "center" },
    ]);
  });

  it("同位置再设一次是替换，不会堆出两条挨着的制表位", () => {
    const editor = harness([para([text("第一章")])]);
    runFormatCommand(editor, "setTabStop", [rightDot], ["第一章"]);
    runFormatCommand(editor, "setTabStop", [{ positionPt: 56.7, alignment: "right" }], ["第一章"]);
    expect(editor.state.doc.child(0).attrs.tabStops).toEqual([
      { positionPt: 56.7, alignment: "right" },
    ]);
  });

  it("「清除」只删那一位；没建模过制表位的段落不因此变成“全部清除”", () => {
    const editor = harness([
      para([text("甲")]),
      paraWith([{ positionPt: 56.7, alignment: "right", leader: "dot" }], "乙"),
    ]);
    expect(runFormatCommand(editor, "clearTabStop", [56.7], ["甲", "乙"])).toBe(true);
    // 甲从来没建模过：写成 [] 等于替用户声明"这段清空了制表位"，导出会删掉它的 w:tabs。
    expect(editor.state.doc.child(0).attrs.tabStops).toBeUndefined();
    expect(editor.state.doc.child(1).attrs.tabStops).toEqual([]);
  });

  it("「全部清除」把选区里每一段写成空清单", () => {
    const editor = harness([
      paraWith([rightDot], "甲"),
      paraWith([{ positionPt: 100, alignment: "center" }], "乙"),
    ]);
    expect(runFormatCommand(editor, "clearAllTabStops", [], ["甲", "乙"])).toBe(true);
    expect(editor.state.doc.child(0).attrs.tabStops).toEqual([]);
    expect(editor.state.doc.child(1).attrs.tabStops).toEqual([]);
  });

  it("光标停在段落里（没拉选区）也能设制表位，面板才是这么用的", () => {
    const editor = harness([para([text("签名行")])]);
    expect(
      runFormatCommand(editor, "setTabStop", [rightDot], ["签名行"], undefined, "cursor"),
    ).toBe(true);
    expect(editor.state.doc.child(0).attrs.tabStops).toEqual([rightDot]);
  });
});

describe("首字下沉 dropCap（真实 schema 与 docx 往返）", () => {
  /** Word 写首字下沉的形状：`w:pPr/w:framePr` 带 `w:dropCap` 与 `w:lines`。 */
  const dropCapDoc = (pPr: string) =>
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
    `<w:body><w:p><w:pPr>${pPr}</w:pPr><w:r><w:t>正文</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`;

  it("有值 / 显式关掉 / 没建模过三种状态都活过 doc JSON 往返", () => {
    const applied = schema.nodeFromJSON({
      type: "paragraph",
      attrs: { dropCap: { position: "drop", lines: 3 } },
      content: [text("下沉段")],
    })!;
    expect(applied.attrs.dropCap).toEqual({ position: "drop", lines: 3 });
    expect(() => schema.nodeFromJSON(applied.toJSON()).check()).not.toThrow();

    const cleared = schema.nodeFromJSON({
      type: "paragraph",
      attrs: { dropCap: null },
      content: [text("原来下沉")],
    })!;
    // null 必须真的留在 JSON 里：丢了它，"关掉下沉"会被读成"没建模过"而把框架又写回去。
    expect(JSON.stringify(cleared.toJSON())).toContain('"dropCap":null');

    const untouched = schema.nodeFromJSON({ type: "paragraph", content: [text("普通段")] })!;
    expect(untouched.attrs.dropCap).toBeUndefined();
  });

  it("源文档的下沉进真实 schema，一个没动过的段落写回还是那一句", () => {
    const editor = harness(ooxmlToPm(dropCapDoc(`<w:framePr w:dropCap="drop" w:lines="3"/>`)).content);
    expect(editor.state.doc.child(0).attrs.dropCap).toEqual({ position: "drop", lines: 3 });
    // 默认补齐的其它段落属性不能牵连这条框架：透传回来还是同一个元素、同一对属性。
    expect(pmToOoxml(editor.state.doc.toJSON() as PMDoc)).toContain(
      '<w:framePr w:dropCap="drop" w:lines="3"/>',
    );
  });

  it("面板改成突出页边距外、四行：写回换掉 dropCap 与 lines，别的属性不牵动", () => {
    const editor = harness(
      ooxmlToPm(
        dropCapDoc(`<w:ind w:left="300"/><w:framePr w:dropCap="drop" w:lines="2" w:hSpace="180"/>`),
      ).content,
    );
    expect(
      runFormatCommand(
        editor,
        "setParagraphFormat",
        [{ dropCap: { position: "margin", lines: 4 } }],
        ["正文"],
      ),
    ).toBe(true);
    const xml = pmToOoxml(editor.state.doc.toJSON() as PMDoc);
    expect(xml).toContain('<w:framePr w:dropCap="margin" w:lines="4" w:hSpace="180"/>');
    // 段落自己的缩进（Word 写在 framePr 之后）与这条框架各归各的。
    expect(xml).toContain('<w:ind w:left="300"/>');
  });

  it("选「无」删掉整条纯下沉框架：留下空的 w:framePr 会把正文变成匿名文本框", () => {
    const editor = harness(
      ooxmlToPm(dropCapDoc(`<w:framePr w:dropCap="drop" w:lines="3"/>`)).content,
    );
    expect(
      runFormatCommand(editor, "setParagraphFormat", [{ dropCap: null }], ["正文"]),
    ).toBe(true);
    expect(pmToOoxml(editor.state.doc.toJSON() as PMDoc)).not.toContain("framePr");
  });

  it("没建模的段落属性跟着 rawPpr 走过编辑器：真文本框的框架一个字都不动", () => {
    // 这条框架没有 dropCap，docx-io 读成"没建模过"；它必须原样活着，否则经编辑器存一次
    // 就把用户的文本框洗成正文（同一段落上的 w:pBdr 是同一个道理）。
    const frame =
      `<w:framePr w:w="2191" w:h="811" w:hRule="exact" w:wrap="around" ` +
      `w:vAnchor="margin" w:hAnchor="page" w:x="2970" w:y="1" w:anchorLock="1"/>`;
    const editor = harness(
      ooxmlToPm(dropCapDoc(`<w:pBdr><w:bottom w:val="single" w:sz="6"/></w:pBdr>${frame}`)).content,
    );
    expect(editor.state.doc.child(0).attrs.dropCap).toBeUndefined();
    const xml = pmToOoxml(editor.state.doc.toJSON() as PMDoc);
    expect(xml).toContain(frame);
    expect(xml).toContain('<w:pBdr><w:bottom w:val="single" w:sz="6"/></w:pBdr>');
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

  it("编号按正文出现顺序现算，脚注与尾注各数各的", () => {
    const doc = harness([
      para([text("正文甲"), fn("7")]),
      para([text("正文乙"), en("9"), text("续"), fn("4")]),
    ]).state.doc;
    const index = indexFootnoteRefs(doc);
    expect(index[footnoteRefKey("footnote", "7")]?.number).toBe(1);
    // id 是内部键，不是版面上的号：正文里后出现的 "4" 显示为 2。
    expect(index[footnoteRefKey("footnote", "4")]?.number).toBe(2);
    expect(index[footnoteRefKey("endnote", "9")]?.number).toBe(1);
  });

  it("重复引用同一个 id 得到同一个号，pos 记第一次出现的位置", () => {
    const doc = harness([para([text("前"), fn("1")]), para([text("后"), fn("1")])]).state.doc;
    const index = indexFootnoteRefs(doc);
    const key = footnoteRefKey("footnote", "1");
    expect(index[key]).toEqual({ number: 1, pos: 2 });
    expect(doc.nodeAt(2)!.attrs.id).toBe("1");
  });

  it("id 是空串的引用点不进表，芯片回落到部件字样", () => {
    const doc = harness([para([text("正文"), fn("")])]).state.doc;
    expect(indexFootnoteRefs(doc)).toEqual({});
  });
});

/**
 * 排名 4 ②：批注的三个锚点进真实 schema。
 *
 * 这一段只验一件事——节点活着穿过 doc JSON。不注册的话 ProseMirror 会把不认识的内联
 * 节点直接丢掉，`docx-io` 保存时按锚点对齐条目，正文里的批注就整批消失了。
 */
describe("批注锚点（真实 schema）", () => {
  const start = (id: string) => ({ type: "commentRangeStart", attrs: { id } });
  const end = (id: string) => ({ type: "commentRangeEnd", attrs: { id } });
  const ref = (id: string) => ({ type: "commentReference", attrs: { id } });

  it("id 活过 doc JSON，且不占正文：批注文字在部件里，不在树上", () => {
    const editor = harness([
      para([text("正文甲"), start("0"), text("被批注的文字"), end("0"), ref("0")]),
    ]);
    const doc = editor.state.doc;
    expect(doc.textContent).toBe("正文甲被批注的文字");
    expect(blocksOf(doc).map((b) => b.text)).toEqual(["正文甲被批注的文字"]);
    const json = JSON.parse(JSON.stringify(doc.toJSON())) as {
      content: { content: { type: string; attrs?: Record<string, unknown> }[] }[];
    };
    expect(json.content[0]!.content.map((n) => n.type)).toEqual([
      "text",
      "commentRangeStart",
      "text",
      "commentRangeEnd",
      "commentReference",
    ]);
    expect(json.content[0]!.content[1]).toMatchObject({
      type: "commentRangeStart",
      attrs: { id: "0" },
    });
    expect(json.content[0]!.content[4]).toMatchObject({
      type: "commentReference",
      attrs: { id: "0" },
    });
  });

  it("跨段区间：端点分处两个段落也各自留在树里，段落边界不变", () => {
    const doc = harness([para([text("上"), start("1")]), para([end("1"), text("下")])]).state.doc;
    expect(doc.childCount).toBe(2);
    expect(doc.firstChild!.childCount).toBe(2);
    expect(doc.lastChild!.childCount).toBe(2);
    expect(doc.textContent).toBe("上下");
  });

  it("只装锚点的段落不塌掉，也不会多出一个可编辑的空位", () => {
    const doc = harness([para([start("2"), end("2"), ref("2")]), para([text("正文")])]).state.doc;
    expect(doc.childCount).toBe(2);
    expect(doc.firstChild!.childCount).toBe(3);
    expect(blocksOf(doc).map((b) => b.text)).toEqual(["", "正文"]);
  });

  it("批注与书签同 id 并存：两套锚点互不吞并", () => {
    // 注解 id 是整个文档共享的一个编号空间，Word 就把书签和批注写成同一个号。
    const doc = harness([
      para([
        { type: "bookmarkStart", attrs: { id: "0", name: "_Toc1" } },
        start("0"),
        text("正文"),
        end("0"),
        { type: "bookmarkEnd", attrs: { id: "0" } },
      ]),
    ]).state.doc;
    expect(doc.firstChild!.childCount).toBe(5);
    expect(doc.textContent).toBe("正文");
  });

  /**
   * 排名 4 ⑤：侧栏要的两条命令加一份位置索引。
   *
   * 断言都盯着位置：这批命令全靠 PM 位置动手，位置算错就是"批注框住别的字"这种
   * 用户看得见、单测最容易放过的缺陷。
   *
   * @param editor harness 出来的编辑器。
   * @param name 要跑的命令名。
   * @param id 传给命令的 `w:id`。
   * @param words 选区要盖住的文字；传 null 表示"用文档当前选区"，用例自己先派事务。
   * @returns 命令有没有真的动手（返回 false 时不该产生事务）。
   */
  function runComment(
    editor: Editor,
    name: "insertCommentRange" | "deleteComment",
    id: string,
    words: string[] | null,
  ): boolean {
    if (words) selectWords(editor, words);
    const commands = (
      CommentCommands as unknown as {
        config: {
          addCommands: () => Record<string, (...args: unknown[]) => (props: CommandProps) => boolean>;
        };
      }
    ).config.addCommands.call({ options: {}, name: "commentCommands" });
    const command = commands[name];
    if (!command) throw new Error(`批注扩展里没有命令 ${name}`);
    return command(id)({
      editor,
      state: editor.state,
      dispatch: editor.view.dispatch,
      view: editor.view,
    } as unknown as CommandProps);
  }

  it("选区插入：区间起点在字前，终点与引用记号按 Word 的顺序跟在字后", () => {
    const editor = harness([para([text("正文甲")])]);
    expect(runComment(editor, "insertCommentRange", "7", ["正文甲"])).toBe(true);
    const doc = editor.state.doc;
    // 锚点不占正文：被批的字一个字都不能少。
    expect(doc.textContent).toBe("正文甲");
    expect(commentAnchorRanges(doc)).toEqual([{ id: "7", start: 1, end: 5, reference: 6 }]);
    expect(pmToOoxml(doc.toJSON() as PMDoc)).toContain(
      `<w:commentRangeStart w:id="7"/><w:r><w:t xml:space="preserve">正文甲</w:t></w:r><w:commentRangeEnd w:id="7"/><w:r><w:commentReference w:id="7"/></w:r>`,
    );
  });

  it("选区只盖住一部分字：锚点贴在被选的那个 run 两侧，前后的字在区间外", () => {
    // 同格式的相邻 run 会被 ProseMirror 并成一个，中间那段得带个标记才站得住。
    const editor = harness([
      para([text("正文甲"), text("被批注的字", [bold()]), text("正文乙")]),
    ]);
    expect(runComment(editor, "insertCommentRange", "7", ["被批注的字"])).toBe(true);
    const doc = editor.state.doc;
    expect(doc.textContent).toBe("正文甲被批注的字正文乙");
    expect(commentAnchorRanges(doc)).toEqual([{ id: "7", start: 4, end: 10, reference: 11 }]);
    expect(doc.nodeAt(4)!.type.name).toBe("commentRangeStart");
    expect(doc.nodeAt(5)!.text).toBe("被批注的字");
    expect(doc.nodeAt(11)!.type.name).toBe("commentReference");
  });

  it("光标没拉选区时不插入——批注总得有个被批的对象", () => {
    const editor = harness([para([text("正文甲")])]);
    selectWords(editor, ["正文甲"], "cursor");
    const before = editor.state.doc;
    expect(runComment(editor, "insertCommentRange", "0", null)).toBe(false);
    expect(editor.state.doc).toBe(before);
  });

  it("跨段选区：两端各落进自己的段落，段落数与文字都不动", () => {
    const editor = harness([para([text("上段字")]), para([text("下段字")])]);
    expect(runComment(editor, "insertCommentRange", "3", ["上段字", "下段字"])).toBe(true);
    const doc = editor.state.doc;
    expect(doc.childCount).toBe(2);
    expect(doc.textContent).toBe("上段字下段字");
    // 起点落在第一段的内容头上，终点与引用记号落在最后一段的内容尾上。
    expect(commentAnchorRanges(doc)).toEqual([{ id: "3", start: 1, end: 10, reference: 11 }]);
  });

  it("全选（选区起点在文档开头）时锚点仍然在段内，不会把树插坏", () => {
    const editor = harness([para([text("整篇的字")])]);
    editor.view.dispatch(
      editor.state.tr.setSelection(
        TextSelection.create(editor.state.doc, 0, editor.state.doc.content.size),
      ),
    );
    expect(runComment(editor, "insertCommentRange", "5", null)).toBe(true);
    const doc = editor.state.doc;
    expect(doc.childCount).toBe(1);
    expect(doc.firstChild!.firstChild!.type.name).toBe("commentRangeStart");
    expect(commentAnchorRanges(doc)).toEqual([
      { id: "5", start: 1, end: 6, reference: 7 },
    ]);
  });

  it("删除批注：三个锚点一起走，另一条批注不受牵连", () => {
    const editor = harness([
      para([text("甲"), start("0"), text("被批的字"), end("0"), ref("0")]),
      para([start("1"), text("乙"), end("1"), ref("1")]),
    ]);
    expect(runComment(editor, "deleteComment", "0", ["甲"])).toBe(true);
    const doc = editor.state.doc;
    expect(doc.textContent).toBe("甲被批的字乙");
    expect(commentAnchorRanges(doc)).toEqual([{ id: "1", start: 8, end: 10, reference: 11 }]);
  });

  it("删没有的 id 返回 false，不产生空事务", () => {
    const editor = harness([para([start("0"), text("甲"), end("0"), ref("0")])]);
    const before = editor.state.doc;
    expect(runComment(editor, "deleteComment", "9", ["甲"])).toBe(false);
    expect(editor.state.doc).toBe(before);
  });

  it("索引按正文顺序排，空 id 的锚点不进表", () => {
    const doc = harness([
      para([start("2"), text("甲"), end("2")]),
      para([start(""), text("乙"), end("")]),
      para([start("1"), text("丙"), end("1"), ref("1")]),
    ]).state.doc;
    expect(commentAnchorRanges(doc)).toEqual([
      { id: "2", start: 1, end: 3 },
      { id: "1", start: 11, end: 13, reference: 14 },
    ]);
  });
});

describe("修订跟踪（真实 schema）", () => {
  const ins = {
    type: "revision",
    attrs: { kind: "ins", id: "1", author: "张三", date: "2026-01-02T03:04:05Z" },
  };
  const del = { type: "revision", attrs: { kind: "del", id: "2", author: "李四", date: null } };

  /** Word 写修订的形状：`w:ins`/`w:del` 与 `w:r` 平级，被删的文字写成 `w:delText`。 */
  const TRACKED =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
    `<w:body><w:p><w:r><w:t>正文</w:t></w:r>` +
    `<w:ins w:id="1" w:author="张三" w:date="2026-01-02T03:04:05Z"><w:r><w:t>插入的字</w:t></w:r></w:ins>` +
    `<w:del w:id="2" w:author="李四"><w:r><w:delText>删掉的字</w:delText></w:r></w:del></w:p>` +
    `<w:sectPr/></w:body></w:document>`;

  /**
   * 文档里每个修订标记的属性，按出现顺序。
   *
   * @param doc 编辑器当前文档。
   * @returns `revision` 标记的 attrs 快照（PM 会把缺省属性补齐，所以 `date` 是 null 而不是缺失）。
   */
  function revisions(doc: PMNode): unknown[] {
    const out: unknown[] = [];
    doc.descendants((n) => {
      if (n.isText) for (const m of n.marks) if (m.type.name === "revision") out.push(m.attrs);
      return true;
    });
    return out;
  }

  /**
   * 跑修订命令。
   *
   * @param editor harness 出来的编辑器。
   * @param name `acceptRevision` 或 `rejectRevision`。
   * @param words 光标/选区要落在哪些文字上。
   * @returns 命令有没有碰到修订。
   */
  const runRevision = (
    editor: Editor,
    name: "acceptRevision" | "rejectRevision",
    words: string[],
  ): boolean =>
    runFormatCommand(
      editor,
      name,
      [],
      words,
      Revision as unknown as { config: { addCommands?: () => Record<string, unknown> } },
    );

  it("docx-io 读出的修订进得了真实 schema，写回还是 ins/del 加 delText", () => {
    const editor = harness(ooxmlToPm(TRACKED).content);
    // 被删的文字必须在树上：它在 Word 里是可见的（删除线），一次存盘不能把它弄没。
    expect(editor.state.doc.textContent).toBe("正文插入的字删掉的字");
    expect(revisions(editor.state.doc)).toEqual([
      { kind: "ins", id: "1", author: "张三", date: "2026-01-02T03:04:05Z" },
      { kind: "del", id: "2", author: "李四", date: null },
    ]);
    const xml = pmToOoxml(editor.state.doc.toJSON() as PMDoc);
    expect(xml).toContain(`<w:ins w:id="1" w:author="张三" w:date="2026-01-02T03:04:05Z">`);
    expect(xml).toContain(`<w:delText xml:space="preserve">删掉的字</w:delText>`);
    expect(xml).not.toContain(`<w:t xml:space="preserve">删掉的字</w:t>`);
  });

  it("整段刷正文色不把修订抹掉（§2.3：局部格式不许拍平）", () => {
    const editor = harness([para([text("正文"), text("插入的字", [ins])])]);
    applyStyles(editor, [{ index: 0, style: { color: "#000000" } }]);
    expect(revisions(editor.state.doc)).toEqual([
      { kind: "ins", id: "1", author: "张三", date: "2026-01-02T03:04:05Z" },
    ]);
    expect(run(editor, "插入的字").attrs.color).toBe("#000000");
  });

  it("接受：插入的字留下、删掉的字走；拒绝正好反过来", () => {
    const doc = () =>
      harness([para([text("正文"), text("插入的字", [ins]), text("删掉的字", [del])])]);

    const keepInsertion = doc();
    expect(runRevision(keepInsertion, "acceptRevision", ["插入的字"])).toBe(true);
    expect(keepInsertion.state.doc.textContent).toBe("正文插入的字删掉的字");
    // 只处理选区碰到的那条：另一条修订原样留着。
    expect(revisions(keepInsertion.state.doc)).toEqual([
      { kind: "del", id: "2", author: "李四", date: null },
    ]);

    const confirmDeletion = doc();
    expect(runRevision(confirmDeletion, "acceptRevision", ["删掉的字"])).toBe(true);
    expect(confirmDeletion.state.doc.textContent).toBe("正文插入的字");
    expect(revisions(confirmDeletion.state.doc)).toEqual([
      { kind: "ins", id: "1", author: "张三", date: "2026-01-02T03:04:05Z" },
    ]);

    const dropInsertion = doc();
    expect(runRevision(dropInsertion, "rejectRevision", ["插入的字"])).toBe(true);
    expect(dropInsertion.state.doc.textContent).toBe("正文删掉的字");

    const restoreDeletion = doc();
    expect(runRevision(restoreDeletion, "rejectRevision", ["删掉的字"])).toBe(true);
    expect(restoreDeletion.state.doc.textContent).toBe("正文插入的字删掉的字");
    expect(revisions(restoreDeletion.state.doc)).toEqual([
      { kind: "ins", id: "1", author: "张三", date: "2026-01-02T03:04:05Z" },
    ]);
  });

  it("一次修订被 Word 按段落拆开时，接受一次全清", () => {
    // Word 给跨段的一次插入写的是同一个 `w:id` 的两条 `w:ins`；用户只点在第一段里，
    // 接受的也应该是整条修订，而不是那一段。
    const editor = harness([
      para([text("正文"), text("甲段插入", [ins])]),
      para([text("乙段插入", [ins])]),
    ]);
    expect(runRevision(editor, "acceptRevision", ["甲段插入"])).toBe(true);
    expect(editor.state.doc.textContent).toBe("正文甲段插入乙段插入");
    expect(revisions(editor.state.doc)).toEqual([]);
  });

  it("选区没碰到修订时命令返回 false（菜单项据此置灰）", () => {
    const editor = harness([para([text("正文"), text("插入的字", [ins])])]);
    expect(runRevision(editor, "acceptRevision", ["正文"])).toBe(false);
  });

  it("右键菜单的落点事实与命令同一个判据：碰到修订才亮，碰不到就灰", () => {
    const editor = harness([para([text("正文"), text("插入的字", [ins])])]);
    selectWords(editor, ["插入的字"]);
    expect(selectionTouchesRevision(editor)).toBe(true);
    selectWords(editor, ["正文"]);
    expect(selectionTouchesRevision(editor)).toBe(false);
    // 菜单亮 == 命令有活干：两条都走 `revisionsTouching`，不能一个说有一个说无。
    expect(runRevision(editor, "acceptRevision", ["正文"])).toBe(false);
    selectWords(editor, ["插入的字"]);
    expect(runRevision(editor, "acceptRevision", ["插入的字"])).toBe(true);
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

  it("套整段字符样式时不拍平用户手打的下标", () => {
    // 上标/下标不在 AI 与模板的字符样式词表里（`CHAR_STYLE_KEYS`），所以既不参与"这段
    // 是否统一"的投票，也不会被整段 addMark/removeMark 扫到——化学式的角标必须活着。
    const editor = harness([
      para([text("溶解度："), text("H"), text("2", [vertAlign("subscript")]), text("O")]),
    ]);
    const blocks = blocksOf(editor.state.doc);
    expect(blocks[0]!.mixed ?? []).not.toContain("vertAlign");
    applyStyles(editor, [{ index: 0, style: bodyOps(blocks)[0]! }]);
    expect(marksOfRun(editor, "2")).toContain("vertAlign");
    expect(marksOfRun(editor, "溶解度")).not.toContain("vertAlign");
    expect(run(editor, "2").attrs.fontFamily).toBe("仿宋");
  });

  it("真实 schema 收下上标与下标两个档位，档位不会被默认值抹平", () => {
    // 标记的 `val` 缺省是 superscript，所以"下标读回上标"这种错只会在这里露出来。
    const editor = harness([
      para([text("x", [vertAlign("superscript")]), text("y", [vertAlign("subscript")])]),
    ]);
    const vals: (string | undefined)[] = [];
    editor.state.doc.descendants((n) => {
      if (n.isText) {
        for (const m of n.marks) if (m.type.name === "vertAlign") vals.push(m.attrs.val);
      }
      return true;
    });
    expect(vals).toEqual(["superscript", "subscript"]);
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

describe("行内 op：插入与撤掉引用域（真实 schema）", () => {
  const addrOf = (editor: Editor, index: number) => addressOf(blocksOf(editor.state.doc), index)!;
  const refField = (name: string) => ({
    type: "field",
    attrs: { kind: "REF", instr: `REF ${name}`, display: "1" },
  });

  it("偏移按文字计数：域和书签都不占一个字", () => {
    const editor = harness([
      para([text("见"), { type: "field", attrs: { kind: "SEQ", instr: `SEQ 图 \\* ARABIC`, display: "1" } }, text("下图")]),
    ]);
    // 「见下图」共 3 个字，偏移 3 是这一段的末尾——中间那条域不参与计数。
    const r = executeOps(editor, [{ kind: "insertInline", addr: addrOf(editor, 0), at: 3, json: refField("_Ref1") }]);
    expect(r.ok).toBe(true);
    const p = editor.state.doc.firstChild!;
    expect(p.child(p.childCount - 1).type.name).toBe("field");
    expect(p.textContent).toBe("见下图");
  });

  it("偏移落在两段文字中间时，域插进那个缝隙而不是开头结尾", () => {
    const editor = harness([para([text("见"), text("下图")])]);
    const r = executeOps(editor, [{ kind: "insertInline", addr: addrOf(editor, 0), at: 1, json: refField("_Ref1") }]);
    expect(r.ok).toBe(true);
    const p = editor.state.doc.firstChild!;
    expect(p.child(1).type.name).toBe("field");
    expect(p.child(2).textContent).toBe("下图");
  });

  it("撤引用只认指令文字：同段另一个域留着，指令不匹配的整段不写", () => {
    const editor = harness([
      para([
        text("见"),
        { type: "field", attrs: { kind: "REF", instr: `REF _Ref1`, display: "1" } },
        { type: "field", attrs: { kind: "PAGEREF", instr: `PAGEREF _Toc9 \\h`, display: "3" } },
      ]),
    ]);
    const before = editor.state.doc.toJSON();
    const missing = executeOps(editor, [{ kind: "removeInline", addr: addrOf(editor, 0), instr: `REF _Ref7` }]);
    expect(missing.ok).toBe(false);
    expect(editor.state.doc.toJSON()).toEqual(before);

    const r = executeOps(editor, [{ kind: "removeInline", addr: addrOf(editor, 0), instr: `REF _Ref1` }]);
    expect(r.ok).toBe(true);
    const p = editor.state.doc.firstChild!;
    expect(p.childCount).toBe(2);
    expect(p.child(1).attrs.instr).toBe(`PAGEREF _Toc9 \\h`);
  });

  it("偏移在文字之外就拒绝，且一个字节都不写", () => {
    const editor = harness([para([text("见下图")])]);
    const before = editor.state.doc.toJSON();
    const r = executeOps(editor, [{ kind: "insertInline", addr: addrOf(editor, 0), at: 9, json: refField("_Ref1") }]);
    expect(r.ok).toBe(false);
    expect(editor.state.doc.toJSON()).toEqual(before);
  });
});
