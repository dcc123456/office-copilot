import { describe, expect, it } from "vitest";
import { getSchema } from "@tiptap/core";
import type { Editor } from "@tiptap/react";
import type { Node as PMNode } from "@tiptap/pm/model";
import { EditorState, TextSelection } from "@tiptap/pm/state";
import { EDITOR_EXTENSIONS } from "./extensions/editorExtensions.js";
import { executeOps } from "./docBlocks.js";
import { seqInstr } from "./captions.js";
import {
  aboveOrBelow,
  collectRefTargets,
  crossRefJson,
  planCrossReference,
  refInstr,
  type CrossRefOptions,
} from "./crossrefs.js";

/**
 * 交叉引用跑在 app 真正构建的 schema 上：`insertInline` 要在既有段落里插一个行内原子，
 * 只有真实 schema 能证明坐标换算与块模型寻址都对得上。
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
/** Word 写好的题注段：`_Ref` 书签只裹住 `SEQ` 域，所以"书签文字"就是编号。 */
const captionPara = (label: string, display: string, note = "") =>
  para([
    text(`${label} `),
    { type: "bookmarkStart", attrs: { id: "7", name: `_Ref${display}` } },
    { type: "field", attrs: { kind: "SEQ", instr: seqInstr(label), display } },
    { type: "bookmarkEnd", attrs: { id: "7" } },
    ...(note ? [text(` ${note}`)] : []),
  ]);
/** 目录条目段：`_Toc` 书签裹住整行标题文字。 */
const tocEntry = (name: string, title: string, page: string) =>
  para([
    { type: "bookmarkStart", attrs: { id: "9", name } },
    text(title),
    { type: "bookmarkEnd", attrs: { id: "9" } },
    text(`\t${page}`),
  ]);

/** 一段里所有引用域的指令与缓存结果，用来断言落盘形状。 */
function refFields(doc: PMNode): { instr: string; display: string }[] {
  const out: { instr: string; display: string }[] = [];
  doc.descendants((node) => {
    if (node.type.name !== "field") return;
    const instr = String(node.attrs.instr ?? "");
    if (/^(REF|PAGEREF)\s/.test(instr)) out.push({ instr, display: String(node.attrs.display ?? "") });
  });
  return out;
}

const OPTIONS: CrossRefOptions = { target: "_Ref1", what: "page", asHyperlink: true };

describe("交叉引用的指令", () => {
  it("三档引用内容各自对应官方列出的开关", () => {
    expect(refInstr("_Ref2", "text", false)).toBe("REF _Ref2");
    expect(refInstr("_Ref2", "page", false)).toBe("PAGEREF _Ref2");
    expect(refInstr("_Ref2", "position", false)).toBe(`REF _Ref2 \\p`);
    // 勾选超链接只加 `\h`：手写 w:hyperlink 会把每条引用都变蓝（同 ③.2 的取舍）。
    expect(refInstr("_Ref2", "page", true)).toBe(`PAGEREF _Ref2 \\h`);
  });

  it("`\\p` 的缓存结果按插入点在目标前后来回翻", () => {
    expect(aboveOrBelow(3, 20)).toBe("以下");
    expect(aboveOrBelow(30, 20)).toBe("以上");
  });

  it("域 JSON 的 kind 取指令的动词，docx-io 按它写回 fldSimple", () => {
    expect(crossRefJson(`PAGEREF _Ref2 \\h`, "3")).toEqual({
      type: "field",
      attrs: { kind: "PAGEREF", instr: `PAGEREF _Ref2 \\h`, display: "3" },
    });
  });
});

describe("可引用的目标", () => {
  it("题注的书签文字是编号，目录条目的是标题文字", () => {
    const doc = schema.nodeFromJSON({
      type: "doc",
      content: [tocEntry("_Toc1", "第一章 绪论", "2"), captionPara("图", "1", "系统架构")],
    })!;
    const targets = collectRefTargets(doc);
    expect(targets.map((t) => [t.kind, t.covered])).toEqual([
      ["toc", "第一章 绪论"],
      ["caption", "1"],
    ]);
    expect(targets[1]!.blockText).toBe("图  系统架构");
  });

  it("落单的半边不列，Word 的对话框也不列没有终点的书签", () => {
    const doc = schema.nodeFromJSON({
      type: "doc",
      content: [para([{ type: "bookmarkStart", attrs: { id: "1", name: "孤儿" } }, text("文字")])],
    })!;
    expect(collectRefTargets(doc)).toEqual([]);
  });
});

describe("插入交叉引用", () => {
  it("落在真实 schema 上：域插在光标处，原文一个字没动", () => {
    // 「见下图」：光标在"见"之后（文档坐标 2 → 段内文字偏移 1）。
    const editor = harness([para([text("见下图")]), captionPara("图", "1", "系统架构")], 2);
    const plan = planCrossReference(editor, OPTIONS);
    expect(plan.ok).toBe(true);
    expect(plan.instr).toBe(`PAGEREF _Ref1 \\h`);

    const exec = executeOps(editor, plan.ops, plan.after);
    expect(exec.ok).toBe(true);
    const first = editor.state.doc.child(0);
    expect(first.textContent).toBe("见下图");
    expect(first.content.child(1).type.name).toBe("field");
    expect(refFields(editor.state.doc)).toEqual([{ instr: `PAGEREF _Ref1 \\h`, display: "" }]);
  });

  it("撤引用不牵连用户后来的手工编辑", () => {
    const editor = harness([para([text("见下图")]), captionPara("图", "1", "系统架构")], 2);
    const plan = planCrossReference(editor, OPTIONS);
    executeOps(editor, plan.ops, plan.after);
    // 用户在别处改了一条：按内容寻址的逆操作不该被它带跑（AGENTS §2.5）。
    executeOps(editor, [
      {
        kind: "setText",
        addr: { by: "blockText", text: "图  系统架构", nth: 0 },
        text: "图  换个说法",
      },
    ]);

    const undo = executeOps(editor, plan.inverses);
    expect(undo.ok).toBe(true);
    expect(refFields(editor.state.doc)).toEqual([]);
    // setText 写入前会把连续空白并成一个（docBlocks 的 normalise），所以换过说明的题注是单空格。
    expect(editor.state.doc.child(1).textContent).toBe("图 换个说法");
  });

  it("引用内容选『书签文字』时，题注取编号、目录条目取标题", () => {
    const editor = harness(
      [para([text("如所示")]), tocEntry("_Toc1", "第一章 绪论", "2"), captionPara("图", "4", "系统架构")],
      4,
    );
    const asText: CrossRefOptions = { target: "_Ref4", what: "text", asHyperlink: false };
    const plan = planCrossReference(editor, asText);
    expect(plan.ok).toBe(true);
    expect(plan.display).toBe("4");

    const tocPlan = planCrossReference(editor, { ...asText, target: "_Toc1" });
    expect(tocPlan.display).toBe("第一章 绪论");
  });

  it("没有书签、书签被删、光标不在正文里都只给拒绝码", () => {
    const empty = harness([para([text("正文")])], 2);
    expect(planCrossReference(empty, OPTIONS).reason).toBe("no-targets");

    const editor = harness([para([text("见下图")]), captionPara("图", "1", "系统架构")], 2);
    expect(planCrossReference(editor, { ...OPTIONS, target: "_Ref9" }).reason).toBe("bad-target");
  });

});

describe("引用编号跟着题注走", () => {
  it("题注重排后，指向它的 REF 域缓存结果一起改", async () => {
    const { captionNumberPatches } = await import("./captions.js");
    const doc = schema.nodeFromJSON({
      type: "doc",
      content: [
        para([text("参见 "), { type: "field", attrs: { kind: "REF", instr: "REF _Ref1", display: "9" } }]),
        captionPara("图", "1", "旧图"),
      ],
    })!;
    const patches = captionNumberPatches(doc);
    // SEQ 那条号数没变（就是 1 条），只有引用被刷成 1。
    expect(patches).toHaveLength(1);
    expect(patches[0]!.attrs.display).toBe("1");
  });
});
