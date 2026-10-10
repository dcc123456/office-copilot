import { describe, expect, it } from "vitest";
import { getSchema } from "@tiptap/core";
import type { Editor } from "@tiptap/react";
import type { Node as PMNode } from "@tiptap/pm/model";
import { EditorState, TextSelection } from "@tiptap/pm/state";
import { EDITOR_EXTENSIONS } from "./extensions/editorExtensions.js";
import { executeOps } from "./docBlocks.js";
import {
  collectTocEntries,
  nextBookmarkId,
  planTocInsertion,
  tocAnchorForSelection,
  tocEntryJson,
  tocLevelOf,
  tocPageNumberRefresh,
  type TocOptions,
} from "./toc.js";

/**
 * 目录生成跑在 app 真正构建的 schema 上：书签、域边界、制表位都是行内原子或
 * 段落属性，只有真实 schema 能证明这些 JSON 建得出节点、位置算得对。
 */

const schema = getSchema(EDITOR_EXTENSIONS);

function harness(content: unknown, selFrom?: number): Editor {
  const doc = schema.nodeFromJSON({ type: "doc", content });
  let state = EditorState.create({
    schema,
    doc,
    // 不给选区就用默认的文档开头，`tocAnchorForSelection` 那条用例要能指定光标。
    ...(selFrom === undefined ? {} : { selection: TextSelection.create(doc!, selFrom) }),
  });
  return {
    get state() {
      return state;
    },
    get isDestroyed() {
      return false;
    },
    view: {
      dispatch: (tr: Parameters<typeof state.apply>[0]) => {
        state = state.apply(tr);
      },
    },
  } as unknown as Editor;
}

const text = (t: string) => ({ type: "text", text: t });
const para = (content: unknown[], attrs?: Record<string, unknown>) => ({
  type: "paragraph",
  ...(attrs ? { attrs } : {}),
  content,
});
const heading = (level: number, t: string) => ({
  type: "heading",
  attrs: { level },
  content: [text(t)],
});
const bookmark = (id: string, name: string) => ({
  type: "bookmarkStart",
  attrs: { id, name },
});
const bookmarkEnd = (id: string) => ({ type: "bookmarkEnd", attrs: { id } });

const OPTIONS: TocOptions = {
  levels: 2,
  showPageNumbers: true,
  leader: "dot",
  // A4 正文区宽度：换算成磅是这次改动里唯一一处 px→pt，走 docx-io 的函数。
  contentWidthPx: 794,
};

describe("目录条目的收集", () => {
  it("标题按 level、正文段按大纲级别进目录，正文与 9 级不进", () => {
    const doc = schema.nodeFromJSON({
      type: "doc",
      content: [
        heading(1, "第一章"),
        para([text("正文")]),
        heading(3, "太深了"),
        para([text("参考文献级别")], { outlineLevel: 0 }),
        para([text("正文级别")], { outlineLevel: 9 }),
      ],
    })!;
    const entries = collectTocEntries(doc, 2, () => 1);
    expect(entries.map((e) => [e.text, e.level])).toEqual([
      ["第一章", 1],
      ["参考文献级别", 1],
    ]);
  });

  it("空标题不进目录，书签 id 从文档已有的后面接着编", () => {
    const doc = schema.nodeFromJSON({
      type: "doc",
      content: [
        {
          type: "heading",
          attrs: { level: 1 },
          content: [bookmark("7", "_Toc7"), text("有书签的标题"), bookmarkEnd("7")],
        },
        // 空标题在 ProseMirror 里是不允许有空 text 节点的，所以一个 text 都不给。
        { type: "heading", attrs: { level: 1 }, content: [] },
        heading(1, "  "),
        heading(2, "第二章"),
      ],
    })!;
    expect(nextBookmarkId(doc)).toBe("8");
    expect(collectTocEntries(doc, 3, () => 2).map((e) => e.bookmark)).toEqual([
      "_Toc8",
      "_Toc9",
    ]);
  });

  it("tocLevelOf 把 Word 的 0 起大纲级别换成 1 起的级别", () => {
    const h2 = schema.nodeFromJSON(heading(2, "x"))!;
    const outline = schema.nodeFromJSON(para([text("x")], { outlineLevel: 2 }))!;
    const body = schema.nodeFromJSON(para([text("x")]))!;
    expect(tocLevelOf(h2)).toBe(2);
    expect(tocLevelOf(outline)).toBe(3);
    expect(tocLevelOf(body)).toBeNull();
  });
});

describe("条目段落的 JSON", () => {
  const entry = {
    index: 0,
    level: 2,
    text: "小节",
    bookmark: "_Toc1",
    id: "1",
    page: 7,
  };

  it("每条都带右对齐前导符制表位，二级再往右让一级", () => {
    const json = tocEntryJson(entry, OPTIONS, { first: true, last: false });
    expect(json.attrs).toMatchObject({
      indent: 20,
      tabStops: [{ positionPt: 595.5, alignment: "right", leader: "dot" }],
    });
  });

  it("一级不写缩进，页码未知的条目也不写假页码", () => {
    const json = tocEntryJson({ ...entry, level: 1, page: 0 }, OPTIONS, {
      first: false,
      last: false,
    });
    expect(json.attrs?.indent).toBeUndefined();
    const field = (json.content ?? [])[1] as { attrs: Record<string, unknown> };
    expect(field.attrs.display).toBe("");
  });

  it("首条挂 TOC 域的 begin，末条挂 end，中间是标题⇥PAGEREF", () => {
    const first = tocEntryJson({ ...entry, level: 1 }, OPTIONS, { first: true, last: false });
    const last = tocEntryJson({ ...entry, level: 1 }, OPTIONS, { first: false, last: true });
    const [begin] = first.content as [Record<string, unknown>];
    expect(begin).toMatchObject({
      type: "fieldBegin",
      attrs: { kind: "TOC", instr: 'TOC \\o "1-2" \\h \\z \\u', isSeparated: true },
    });
    expect((last.content as Record<string, unknown>[]).at(-1)).toEqual({ type: "fieldEnd" });
    expect(first.content?.[1]).toMatchObject({ type: "text", text: "小节\t" });
    expect(first.content?.[2]).toMatchObject({
      type: "field",
      attrs: { kind: "PAGEREF", instr: "PAGEREF _Toc1 \\h", display: "7" },
    });
  });

  it("关掉显示页码就既不写制表符也不写域，连制表位都不留", () => {
    const json = tocEntryJson(entry, { ...OPTIONS, showPageNumbers: false }, {
      first: true,
      last: true,
    });
    expect(json.content).toHaveLength(3); // fieldBegin + 标题 + fieldEnd
    expect(json.attrs?.tabStops).toBeUndefined();
    expect((json.content as Record<string, unknown>[])[1]).toMatchObject({
      type: "text",
      text: "小节",
    });
  });

  it("JSON 建得出真实节点，段落文字与块模型用来寻址的文字一致", () => {
    const json = tocEntryJson(entry, OPTIONS, { first: true, last: true });
    const node = schema.nodeFromJSON(json)!;
    node.check();
    // 域芯片的 display 不是 textContent：块模型只认得到"标题⇥"。
    expect(node.textContent).toBe("小节\t");
  });
});

describe("插入目录的 op 与撤销（真实 schema）", () => {
  const source = [heading(1, "第一章"), para([text("正文")]), heading(2, "小节"), para([text("结尾")])];

  it("条目落在文档最前面，顺序与标题一致，每个标题都有书签", () => {
    const editor = harness(source);
    const plan = planTocInsertion(editor, OPTIONS);
    expect(plan.ok).toBe(true);
    expect(plan.entries.map((e) => e.text)).toEqual(["第一章", "小节"]);
    const exec = executeOps(editor, plan.ops, plan.after);
    expect(exec.ok).toBe(true);

    const doc = editor.state.doc;
    expect(doc.childCount).toBe(source.length + 2);
    expect(doc.child(0).textContent).toBe("第一章\t");
    expect(doc.child(1).textContent).toBe("小节\t");
    expect(doc.child(0).attrs.tabStops).toEqual([
      { positionPt: 595.5, alignment: "right", leader: "dot" },
    ]);
    // fieldBegin 只属于首条，fieldEnd 只属于末条。
    expect(inlineTypes(doc.child(0))).toContain("fieldBegin");
    expect(inlineTypes(doc.child(1))).toContain("fieldEnd");
    expect(inlineTypes(doc.child(1))).not.toContain("fieldBegin");

    const bookmarks = collectBookmarks(doc);
    expect(bookmarks).toEqual([
      { id: "1", name: "_Toc1", in: "第一章" },
      { id: "2", name: "_Toc2", in: "小节" },
    ]);
  });

  it("撤销目录把条目与书签一起撤掉，文档回到插入前", () => {
    const editor = harness(source);
    const beforeJSON = editor.state.doc.toJSON();
    const plan = planTocInsertion(editor, OPTIONS);
    executeOps(editor, plan.ops, plan.after);
    const undo = executeOps(editor, plan.inverses, plan.before);
    expect(undo.ok).toBe(true);
    expect(editor.state.doc.toJSON()).toEqual(beforeJSON);
  });

  it("指定 position 时目录插在那一段之后，而不是文档开头", () => {
    const editor = harness(source);
    const plan = planTocInsertion(editor, {
      ...OPTIONS,
      position: { at: "after", after: { by: "blockText", text: "正文", nth: 0 } },
    });
    expect(plan.ok).toBe(true);
    executeOps(editor, plan.ops, plan.after);
    // 倒序插入的净效果是条目正序紧跟锚点：正文 → 第一章 → 小节。
    expect(texts(editor.state.doc)).toEqual([
      "第一章",
      "正文",
      "第一章\t",
      "小节\t",
      "小节",
      "结尾",
    ]);
  });

  it("没有够格的标题就不写：拒绝原因是给界面的码，不是文案", () => {
    const editor = harness([para([text("只有正文")])]);
    expect(planTocInsertion(editor, OPTIONS).reason).toBe("no-headings");
    const levels = planTocInsertion(editor, { ...OPTIONS, levels: 0 });
    expect(levels.reason).toBe("bad-levels");
    expect(levels.ops).toEqual([]);
  });
});

describe("插入位置取自光标（Word 在光标处插入目录）", () => {
  const source = [heading(1, "第一章"), para([text("正文")]), heading(2, "小节"), para([text("结尾")])];
  // 正文段里第一个字符的位置：第一段 heading 的 nodeSize（1 + 3 字 + 1）再加一。
  const insideBody = schema.nodeFromJSON({ type: "doc", content: source })!.child(0).nodeSize + 1;

  it("光标在第二段里，就得到那一段的内容地址", () => {
    expect(tocAnchorForSelection(harness(source, insideBody))).toEqual({
      at: "after",
      after: { by: "blockText", text: "正文", nth: 0 },
    });
  });

  it("光标在表格段落里也照样给得出锚点（表格按内容寻址）", () => {
    const doc = [
      heading(1, "第一章"),
      {
        type: "table",
        content: [
          {
            type: "tableRow",
            content: [{ type: "tableCell", content: [para([text("甲")])] }],
          },
        ],
      },
    ];
    // 光标要真的落在单元格文字里：表格、行、单元格、段落各占一层开合（+4）。
    const insideTable = schema.nodeFromJSON({ type: "doc", content: doc })!.child(0).nodeSize + 4;
    expect(tocAnchorForSelection(harness(doc, insideTable)).at).toBe("after");
  });

  it("位置一路带到 op：目录确实插在光标那段之后", () => {
    const withSel = harness(source, insideBody);
    const plan = planTocInsertion(withSel, {
      ...OPTIONS,
      position: tocAnchorForSelection(withSel),
    });
    expect(executeOps(withSel, plan.ops, plan.after).ok).toBe(true);
    expect(texts(withSel.state.doc)).toEqual([
      "第一章",
      "正文",
      "第一章\t",
      "小节\t",
      "小节",
      "结尾",
    ]);
  });
});

describe("刷新条目缓存页码", () => {
  it("分页变了就只改 PAGEREF 的 display，文字一个不动", () => {
    const entry = tocEntryJson(
      { index: 0, level: 1, text: "第一章", bookmark: "_Toc1", id: "1", page: 1 },
      OPTIONS,
      { first: true, last: true },
    );
    const editor = harness([
      entry,
      {
        type: "heading",
        attrs: { level: 1 },
        content: [bookmark("1", "_Toc1"), text("第一章"), bookmarkEnd("1")],
      },
    ]);
    const state = editor.state;
    const tr = tocPageNumberRefresh(state, () => 4);
    expect(tr).not.toBeNull();
    const next = state.apply(tr!);
    const tocPara = next.doc.child(0);
    expect(fieldDisplay(tocPara)).toBe("4");
    // 目录段的文字仍只是"标题⇥"：页码是域的缓存值，不进正文。
    expect(tocPara.textContent).toBe("第一章\t");
    expect(next.doc.lastChild!.textContent).toBe("第一章");
    // 再刷一次就没有要改的了。
    expect(tocPageNumberRefresh(next, () => 4)).toBeNull();
  });

  it("求不出页码（分页还没测出来）时不产生事务", () => {
    const editor = harness([
      {
        type: "heading",
        attrs: { level: 1 },
        content: [bookmark("1", "_Toc1"), text("第一章"), bookmarkEnd("1")],
      },
    ]);
    expect(tocPageNumberRefresh(editor.state, () => 0)).toBeNull();
  });
});

/** 一段里的行内节点类型名，用来断言域边界与书签的位置。 */
function inlineTypes(node: PMNode): string[] {
  const out: string[] = [];
  node.forEach((child) => out.push(child.type.name));
  return out;
}

/** 顶层块的纯文本，按文档顺序。 */
function texts(doc: PMNode): string[] {
  return Array.from({ length: doc.childCount }, (_, i) => doc.child(i).textContent);
}

/** 条目段落里 `PAGEREF` 子域缓存的页码；没有域时 undefined。 */
function fieldDisplay(entry: PMNode): string | undefined {
  for (let i = 0; i < entry.childCount; i++) {
    const child = entry.child(i);
    if (child.type.name === "field") return String(child.attrs.display ?? "");
  }
  return undefined;
}

/** 文档里成对出现的书签：id、名字与它所在的段落文字。 */
function collectBookmarks(doc: PMNode): { id: string; name: string; in: string }[] {
  const out: { id: string; name: string; in: string }[] = [];
  doc.forEach((block) => {
    block.forEach((child) => {
      if (child.type.name === "bookmarkStart") {
        out.push({ id: String(child.attrs.id), name: String(child.attrs.name), in: block.textContent });
      }
    });
  });
  return out;
}
