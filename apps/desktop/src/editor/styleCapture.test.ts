import { describe, expect, it } from "vitest";
import { getSchema, type Editor } from "@tiptap/core";
import type { Node as PMNode } from "@tiptap/pm/model";
import { EditorState, NodeSelection, TextSelection } from "@tiptap/pm/state";
import { EDITOR_EXTENSIONS } from "./extensions/editorExtensions.js";
import { captureStyleFromSelection } from "./styleCapture.js";

/**
 * 抓取逻辑跑在 app 真正装配的 schema 上（AGENTS §5.3：只在临时数组里测过的不算数）——
 * 段落 attrs 的默认值、标记的属性名都以 `EDITOR_EXTENSIONS` 为准，写错键名这里就红。
 */

const schema = getSchema(EDITOR_EXTENSIONS);

/** 文字区间（from 与 to 相同即光标）。 */
type TextRange = { from: number; to: number };

/** 要造的选区：文字区间，或整块节点选中。 */
type SelectionSpec = TextRange | { nodeAt: number };

/** 用给定内容建一份真实 schema 的文档，并按 spec 摆好选区。 */
function harness(content: unknown, select: (doc: PMNode) => SelectionSpec): Editor {
  const doc = schema.nodeFromJSON({ type: "doc", content });
  const spec = select(doc);
  const selection =
    "nodeAt" in spec
      ? NodeSelection.create(doc, spec.nodeAt)
      : TextSelection.create(doc, spec.from, spec.to);
  return { state: EditorState.create({ schema, doc, selection }) } as unknown as Editor;
}

/** 第 n 个顶层块的文字区间（不含两侧的分隔符）。 */
function blockRange(doc: PMNode, wanted: number): TextRange {
  if (wanted >= doc.childCount) throw new Error(`没有第 ${wanted} 个块。`);
  let offset = 0;
  for (let i = 0; i < wanted; i++) offset += doc.child(i).nodeSize;
  return { from: offset + 1, to: offset + doc.child(wanted).nodeSize - 1 };
}

/** 块尾的光标（不选中任何文字）。 */
function blockEnd(doc: PMNode, wanted: number): TextRange {
  const { to } = blockRange(doc, wanted);
  return { from: to, to };
}

/** 跨块整选：从第一块的首字到最后一条文字（PM 允许选区带上块之间的分隔符）。 */
function wholeDoc(doc: PMNode): TextRange {
  return { from: blockRange(doc, 0).from, to: blockRange(doc, doc.childCount - 1).to };
}

const textStyle = (attrs: Record<string, string | null>) => ({
  type: "textStyle",
  attrs: { fontFamily: null, fontFamilyWest: null, fontSize: null, letterSpacing: null, color: null, backgroundColor: null, ...attrs },
});

const paragraph = (attrs: Record<string, unknown>, content: unknown[]) => ({ type: "paragraph", attrs, content });

describe("captureStyleFromSelection", () => {
  it("抓到段落格式与行内标记，套用的样式 id 不算抓到的格式", () => {
    const editor = harness(
      [
        paragraph(
          { indent: 36, textAlign: "center", spacingBefore: 12, styleId: "2" },
          [{ type: "text", text: "重点条款", marks: [{ type: "bold" }, textStyle({ fontSize: "24px", color: "#c00000" })] }],
        ),
      ],
      (doc) => blockRange(doc, 0),
    );
    const result = captureStyleFromSelection(editor);
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.edit.paragraph).toEqual({ indent: 36, textAlign: "center", spacingBefore: 12 });
    expect(result.edit.marks).toHaveLength(2);
    expect(result.edit.marks).toContainEqual({ type: "bold" });
    expect(result.edit.marks).toContainEqual({ type: "textStyle", attrs: { fontSize: "24px", color: "#c00000" } });
  });

  it("两段之间段前距不同就拒绝，并说差在哪个属性", () => {
    const editor = harness(
      [
        paragraph({ spacingBefore: 12 }, [{ type: "text", text: "第一段" }]),
        paragraph({ spacingBefore: 6 }, [{ type: "text", text: "第二段" }]),
      ],
      wholeDoc,
    );
    expect(captureStyleFromSelection(editor)).toEqual({
      ok: false,
      code: "mixed-paragraph-format",
      detail: "spacingBefore",
    });
  });

  it("同一段里加粗只覆盖一半文字时拒绝——不许拍平用户的局部格式", () => {
    const editor = harness(
      [
        paragraph(
          { indent: 36 },
          [
            { type: "text", text: "加粗的部分", marks: [{ type: "bold" }] },
            { type: "text", text: "和没加粗的部分" },
          ],
        ),
      ],
      (doc) => blockRange(doc, 0),
    );
    expect(captureStyleFromSelection(editor)).toEqual({ ok: false, code: "mixed-run-format", detail: "bold" });
  });

  it("跨段的字号不同也算冲突，报的是 textStyle 而不是段落属性", () => {
    const editor = harness(
      [
        paragraph({}, [{ type: "text", text: "一号", marks: [textStyle({ fontSize: "16px" })] }]),
        paragraph({}, [{ type: "text", text: "二号", marks: [textStyle({ fontSize: "14px" })] }]),
      ],
      wholeDoc,
    );
    expect(captureStyleFromSelection(editor)).toEqual({ ok: false, code: "mixed-run-format", detail: "textStyle" });
  });

  it("光标状态也能抓：Word 允许不选中就把样式更新成光标处的格式", () => {
    const editor = harness(
      [
        paragraph(
          { lineHeight: "1.5" },
          [
            { type: "text", text: "已经写完的一行", marks: [{ type: "italic" }] },
            { type: "text", text: "。", marks: [{ type: "italic" }] },
          ],
        ),
      ],
      (doc) => blockEnd(doc, 0),
    );
    const result = captureStyleFromSelection(editor);
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.edit.paragraph).toEqual({ lineHeight: "1.5" });
    expect(result.edit.marks).toEqual([{ type: "italic" }]);
  });

  it("选区里没有段落或标题时不给格式，让 UI 提示先选文字", () => {
    const editor = harness([{ type: "horizontalRule" }], () => ({ nodeAt: 0 }));
    expect(captureStyleFromSelection(editor)).toEqual({ ok: false, code: "no-block" });
  });

  it("空段落也能抓：新建样式时行内标记为空是合法结果", () => {
    const editor = harness([paragraph({ indent: 18 }, [])], (doc) => blockRange(doc, 0));
    const result = captureStyleFromSelection(editor);
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.edit).toEqual({ paragraph: { indent: 18 }, marks: [] });
  });

  it("两段一样的制表位清单抓得出来，不会被当成冲突", () => {
    // 制表位是对象数组，PM 里每个节点各持一份——按引用比的话"完全相同"也是冲突。
    const stop = { positionPt: 56.7, alignment: "right", leader: "dot" };
    const editor = harness(
      [
        paragraph({ tabStops: [stop] }, [{ type: "text", text: "甲" }]),
        paragraph({ tabStops: [{ ...stop }] }, [{ type: "text", text: "乙" }]),
      ],
      wholeDoc,
    );
    const result = captureStyleFromSelection(editor);
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.edit.paragraph.tabStops).toEqual([stop]);
  });

  it("清单不一样就拒绝，报的属性是 tabStops", () => {
    const editor = harness(
      [
        paragraph({ tabStops: [{ positionPt: 56.7, alignment: "right", leader: "dot" }] }, [{ type: "text", text: "甲" }]),
        paragraph({ tabStops: [{ positionPt: 415.3, alignment: "center" }] }, [{ type: "text", text: "乙" }]),
      ],
      wholeDoc,
    );
    expect(captureStyleFromSelection(editor)).toEqual({
      ok: false,
      code: "mixed-paragraph-format",
      detail: "tabStops",
    });
  });

  it("「全部清除」那一段也能抓：空清单原样交给样式定义", () => {
    // 空清单是有意义的（样式里也不该留着制表位），不是"没写过"。
    const editor = harness([paragraph({ tabStops: [] }, [{ type: "text", text: "甲" }])], (doc) =>
      blockRange(doc, 0),
    );
    const result = captureStyleFromSelection(editor);
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.edit.paragraph).toEqual({ tabStops: [] });
  });

  it("没建模过制表位的段落不参与 tabStops 的取值", () => {
    const editor = harness(
      [
        paragraph({ tabStops: [{ positionPt: 56.7, alignment: "right" }] }, [{ type: "text", text: "甲" }]),
        paragraph({}, [{ type: "text", text: "乙" }]),
      ],
      wholeDoc,
    );
    const result = captureStyleFromSelection(editor);
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.edit.paragraph.tabStops).toEqual([{ positionPt: 56.7, alignment: "right" }]);
  });

  it("抓下标要带档位：只留类型名会把它写成上标", () => {
    const editor = harness(
      [
        paragraph(
          {},
          [{ type: "text", text: "二氧化碳", marks: [{ type: "vertAlign", attrs: { val: "subscript" } }] }],
        ),
      ],
      (doc) => blockRange(doc, 0),
    );
    const result = captureStyleFromSelection(editor);
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.edit.marks).toEqual([{ type: "vertAlign", attrs: { val: "subscript" } }]);
  });

  it("上标与下标混在同一段里算冲突，报的是 vertAlign", () => {
    // 一条样式定义只有一个 `w:vertAlign`，抓成上标就把用户的下标盖了（AGENTS §2.3）。
    const editor = harness(
      [
        paragraph({}, [
          { type: "text", text: "平方", marks: [{ type: "vertAlign", attrs: { val: "superscript" } }] },
          { type: "text", text: "与立方", marks: [{ type: "vertAlign", attrs: { val: "subscript" } }] },
        ]),
      ],
      (doc) => blockRange(doc, 0),
    );
    expect(captureStyleFromSelection(editor)).toEqual({
      ok: false,
      code: "mixed-run-format",
      detail: "vertAlign",
    });
  });
});
