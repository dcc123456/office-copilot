import { describe, expect, it } from "vitest";
import { getSchema } from "@tiptap/core";
import type { Editor } from "@tiptap/react";
import type { Node as PMNode } from "@tiptap/pm/model";
import { EditorState, TextSelection } from "@tiptap/pm/state";
import { EDITOR_EXTENSIONS } from "./extensions/editorExtensions.js";
import { executeOps, jsonContentText } from "./docBlocks.js";
import {
  captionJson,
  captionTargetFor,
  captionNumberPatches,
  planCaptionInsertion,
  renumberSequences,
  seqInstr,
  seqLabelOfInstr,
  type CaptionOptions,
} from "./captions.js";

/**
 * 题注跑在 app 真正构建的 schema 上：`SEQ` 域、书签都是行内原子节点，
 * 只有真实 schema 能证明这些 JSON 建得出节点、块模型按文字找得回来。
 */

const schema = getSchema(EDITOR_EXTENSIONS);

function harness(content: unknown, selFrom?: number): Editor {
  const doc = schema.nodeFromJSON({ type: "doc", content });
  let state = EditorState.create({
    schema,
    doc,
    ...(selFrom === undefined ? {} : { selection: TextSelection.create(doc!, selFrom) }),
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

const text = (t: string) => ({ type: "text", text: t });
const para = (content: unknown[]) => ({ type: "paragraph", content });
/** 一条 Word 写好的题注段（书签裹着 SEQ 域，后面跟着说明文字）。 */
const captionPara = (label: string, display: string, note = "") =>
  para([
    text(`${label} `),
    { type: "bookmarkStart", attrs: { id: display, name: `_Ref${display}` } },
    { type: "field", attrs: { kind: "SEQ", instr: seqInstr(label), display } },
    { type: "bookmarkEnd", attrs: { id: display } },
    ...(note ? [text(` ${note}`)] : []),
  ]);

const OPTIONS: CaptionOptions = { label: "图", captionText: "系统架构", position: "below" };

/** 一段里 `SEQ` 域缓存着的号数，用来断言重排结果。 */
function seqNumbers(doc: PMNode): string[] {
  const out: string[] = [];
  doc.descendants((node) => {
    if (node.type.name === "field" && seqLabelOfInstr(String(node.attrs.instr ?? "")))
      out.push(String(node.attrs.display));
  });
  return out;
}

describe("SEQ 指令", () => {
  it("标签拼进指令，读回来是同一个标签", () => {
    expect(seqInstr("图")).toBe(`SEQ 图 \\* ARABIC`);
    expect(seqLabelOfInstr(seqInstr("图"))).toBe("图");
  });

  it("带空格的标签加引号，否则 Word 只认第一个词", () => {
    expect(seqInstr("图 1")).toBe(`SEQ "图 1" \\* ARABIC`);
    expect(seqLabelOfInstr(`SEQ "图 1" \\* ARABIC`)).toBe("图 1");
  });

  it("不是 SEQ 的指令读不出标签", () => {
    expect(seqLabelOfInstr(`PAGEREF _Toc1 \\h`)).toBeNull();
    // 标签是空引号：序列名不存在，重排时该跳过而不是当成一列来数。
    expect(seqLabelOfInstr(`SEQ "" \\* ARABIC`)).toBeNull();
  });
});

describe("题注段落的形状", () => {
  it("书签裹住 SEQ 域，交叉引用点就是这条书签", () => {
    const json = captionJson({ label: "图", captionText: "系统架构" }, 2, {
      id: "7",
      name: "_Ref7",
    });
    expect(json.content?.map((n) => (n as { type: string }).type)).toEqual([
      "text",
      "bookmarkStart",
      "field",
      "bookmarkEnd",
      "text",
    ]);
    // 块模型按纯文本寻址，而域芯片自己的显示文字不属于正文。
    expect(jsonContentText(json)).toBe("图  系统架构");
  });

  it("没有说明文字时只有标签和编号", () => {
    const json = captionJson({ label: "表", captionText: "" }, 1, { id: "1", name: "_Ref1" });
    expect(json.content).toEqual([
      { type: "text", text: "表 " },
      { type: "bookmarkStart", attrs: { id: "1", name: "_Ref1" } },
      { type: "field", attrs: { kind: "SEQ", instr: `SEQ 表 \\* ARABIC`, display: "1" } },
      { type: "bookmarkEnd", attrs: { id: "1" } },
    ]);
  });
});

describe("插入题注", () => {
  it("落在真实 schema 上：编号、书签与说明文字都在段里", () => {
    const editor = harness([para([text("正文")])], 2);
    const plan = planCaptionInsertion(editor, OPTIONS);
    expect(plan.ok).toBe(true);
    expect(plan.target?.number).toBe(1);

    const exec = executeOps(editor, plan.ops, plan.after);
    expect(exec.ok).toBe(true);
    const doc = editor.state.doc;
    expect(doc.childCount).toBe(2);
    expect(doc.child(1).textContent).toBe("图  系统架构");
    expect(seqNumbers(doc)).toEqual(["1"]);
  });

  it("撤销题注把整段原样收回，正文回到插入前", () => {
    const editor = harness([para([text("正文")])], 2);
    const plan = planCaptionInsertion(editor, OPTIONS);
    const beforeJSON = editor.state.doc.toJSON();
    executeOps(editor, plan.ops, plan.after);
    expect(editor.state.doc.childCount).toBe(2);

    const undo = executeOps(editor, plan.inverses, plan.before);
    expect(undo.ok).toBe(true);
    expect(editor.state.doc.toJSON()).toEqual(beforeJSON);
  });

  it("书签 id 与目录的 _Toc 共用一个计数器，不撞号", () => {
    const editor = harness(
      [
        para([
          { type: "bookmarkStart", attrs: { id: "1", name: "_Toc1" } },
          text("第一章"),
          { type: "bookmarkEnd", attrs: { id: "1" } },
        ]),
      ],
      4,
    );
    const plan = planCaptionInsertion(editor, { ...OPTIONS, captionText: "" });
    expect(plan.target?.bookmark).toEqual({ id: "2", name: "_Ref2" });
  });

  it("标签为空或选区给不出块地址时拒绝写入", () => {
    const editor = harness([para([text("正文")])], 2);
    expect(planCaptionInsertion(editor, { ...OPTIONS, label: "  " }).reason).toBe("bad-label");
    // 块索引都落在正文之外（选区被外部改坏了）时，target 求不出来，这次什么都不写。
    expect(captionTargetFor(editor.state.doc, 999, { label: "图", position: "below" })).toBeNull();
  });
});

describe("题注编号跟着正文走", () => {
  it("删掉第一条以后，后面的号数依次前移", () => {
    // 读回来的文档里留着 2、3（第一条刚被删掉）。
    const doc = schema.nodeFromJSON({
      type: "doc",
      content: [captionPara("图", "2", "架构图"), captionPara("图", "3", "流程图")],
    })!;
    expect(seqNumbers(doc)).toEqual(["2", "3"]);

    const state = EditorState.create({ schema, doc });
    const tr = renumberSequences(state);
    expect(tr).not.toBeNull();
    expect(seqNumbers(tr!.doc)).toEqual(["1", "2"]);
    // 派生显示值不进撤销栈，Ctrl+Z 撤的是用户那次删除。
    expect(tr!.getMeta("addToHistory")).toBe(false);
  });

  it("每个标签各数各的，图和表互不打乱", () => {
    const doc = schema.nodeFromJSON({
      type: "doc",
      content: [
        captionPara("图", "9", ""),
        captionPara("表", "9", ""),
        captionPara("图", "9", ""),
      ],
    })!;
    const state = EditorState.create({ schema, doc });
    expect(seqNumbers(renumberSequences(state)!.doc)).toEqual(["1", "1", "2"]);
  });

  it("号数都对时不派发事务", () => {
    const doc = schema.nodeFromJSON({
      type: "doc",
      content: [captionPara("图", "1", ""), captionPara("图", "2", "")],
    })!;
    expect(captionNumberPatches(doc)).toEqual([]);
    expect(renumberSequences(EditorState.create({ schema, doc }))).toBeNull();
  });

  it("插入点之前有几条同标签的题注，决定新题注是几号", () => {
    const doc = schema.nodeFromJSON({
      type: "doc",
      content: [captionPara("图", "1", ""), para([text("正文")]), captionPara("图", "2", "")],
    })!;
    const target = EditorState.create({ schema, doc });
    // 光标在中间那段（块索引 1，文档坐标 7..11）里：上方排在第一条题注之后，下方排在第二条之前。
    expect(captionTargetFor(target.doc, 9, { label: "图", position: "above" })?.number).toBe(2);
    expect(captionTargetFor(target.doc, 9, { label: "图", position: "below" })?.number).toBe(2);
    expect(captionTargetFor(target.doc, 9, { label: "表", position: "below" })?.number).toBe(1);
  });
});
