import { describe, expect, it } from "vitest";
import {
  alignRewrite,
  coerceProposal,
  planEdit,
  revert,
  rewritableBlocks,
  rewriteResume,
  runOps,
  type Block,
  type DiagramResolver,
  type Proposal,
} from "./docEdit.js";

const para = (index: number, text: string, extra: Partial<Block> = {}): Block => ({
  index,
  type: "paragraph",
  text,
  style: {},
  ...extra,
});

const doc = (...texts: string[]): Block[] => texts.map((t, i) => para(i, t));
const lines = (b: Block[]) => b.map((x) => `${x.type}:${x.text}`);

const diagram: DiagramResolver = () => ({ src: "data:image/png;base64,AA", width: 100, height: 60 });

describe("planEdit", () => {
  it("resolves text, style, insert and delete ops against the outline", () => {
    const before = doc("第一节", "第二节", "第三节");
    const proposal: Proposal = {
      summary: "改一版",
      ops: [
        { op: "setText", index: 1, text: "第二节（重写）" },
        { op: "setStyle", index: 0, style: { bold: true, fontSize: "16px" } },
        { op: "insertParagraph", after: 2, text: "结语" },
        { op: "deleteBlock", index: 2 },
      ],
    };
    const plan = planEdit(before, proposal);
    expect(plan.ok).toBe(true);
    expect(plan.skipped).toEqual([]);
    expect(lines(plan.after)).toEqual([
      "paragraph:第一节",
      "paragraph:第二节（重写）",
      "paragraph:结语",
    ]);
    expect(plan.after[0]!.style.bold).toBe(true);
    // Indexes are the ones the model was shown: "第三节" is deleted, not the
    // paragraph the insert pushed after it.
    expect(plan.after.some((b) => b.text === "第三节")).toBe(false);
  });

  it("新建空文档里的连续插入按落地过程中的编号排队，段序不能乱", () => {
    // 模型看到空文档只能自己编号：after=-1、0、1、2…，每个编号指"写到这里为止的第几段"。
    const before = doc("");
    const proposal: Proposal = {
      summary: "写小说",
      ops: [
        { op: "insertParagraph", after: -1, text: "第一章 启程" },
        { op: "insertParagraph", after: 0, text: "夜色沉进河谷。" },
        { op: "insertParagraph", after: 1, text: "他推开木门。" },
        { op: "insertParagraph", after: 2, text: "门后没有灯。" },
      ],
    };
    const plan = planEdit(before, proposal);
    expect(plan.skipped).toEqual([]);
    expect(plan.after.map((b) => b.text)).toEqual([
      "第一章 启程",
      "夜色沉进河谷。",
      "他推开木门。",
      "门后没有灯。",
      "",
    ]);
  });

  it("有文字的文档里，超出大纲的 after 指的是刚插入的那一段", () => {
    const before = doc("第一段", "第二段");
    const proposal: Proposal = {
      summary: "结尾补两段",
      ops: [
        { op: "insertParagraph", after: 1, text: "补在第二段后" },
        { op: "insertParagraph", after: 2, text: "补在刚插入那段后" },
        { op: "insertParagraph", after: 0, text: "插在原第一段后" },
      ],
    };
    const plan = planEdit(before, proposal);
    expect(plan.skipped).toEqual([]);
    expect(plan.after.map((b) => b.text)).toEqual([
      "第一段",
      "插在原第一段后",
      "第二段",
      "补在第二段后",
      "补在刚插入那段后",
    ]);
  });

  it("revert undoes the whole change and nothing else", () => {
    const before = doc("第一节", "第二节", "第三节");
    const plan = planEdit(before, {
      summary: "s",
      ops: [
        { op: "setText", index: 1, text: "第二节（重写）" },
        { op: "insertParagraph", after: 0, text: "插入段" },
        { op: "deleteBlock", index: 2 },
      ],
    });
    const back = revert(plan.after, plan.inverses);
    expect(back.ok).toBe(true);
    expect(lines(back.after)).toEqual(lines(before));
  });

  it("undoes one modification after the user kept editing elsewhere", () => {
    const before = doc("标题", "正文一", "正文二");
    const plan = planEdit(before, {
      summary: "s",
      ops: [{ op: "setText", index: 1, text: "正文一（AI 改）" }],
    });
    // User edits an untouched block afterwards.
    const current = plan.after.map((b) =>
      b.index === 2 ? { ...b, text: "正文二（用户自己改的）" } : b,
    );
    const back = revert(current, plan.inverses);
    expect(back.ok).toBe(true);
    expect(lines(back.after)).toEqual(["paragraph:标题", "paragraph:正文一", "paragraph:正文二（用户自己改的）"]);
  });

  it("reports a conflict instead of guessing when the target itself changed", () => {
    const before = doc("甲", "乙");
    const plan = planEdit(before, {
      summary: "s",
      ops: [{ op: "setText", index: 1, text: "乙（AI）" }],
    });
    const current = plan.after.map((b) => (b.text === "乙（AI）" ? { ...b, text: "乙被大改" } : b));
    const back = revert(current, plan.inverses);
    expect(back.applied).toHaveLength(0);
    expect(back.failed.length).toBeGreaterThan(0);
  });

  it("addresses duplicate paragraphs by occurrence, not just by text", () => {
    const before = doc("重复", "重复", "其它");
    const plan = planEdit(before, {
      summary: "s",
      ops: [{ op: "setText", index: 1, text: "重复（第二处）" }],
    });
    expect(lines(plan.after)).toEqual(["paragraph:重复", "paragraph:重复（第二处）", "paragraph:其它"]);
    expect(lines(revert(plan.after, plan.inverses).after)).toEqual(lines(before));
  });

  it("expands replaceAll into one undoable op per hit", () => {
    const before = doc("合同 甲方 与 甲方", "无关", "乙方 与 甲方");
    const plan = planEdit(before, {
      summary: "s",
      ops: [{ op: "replaceAll", find: "甲方", replace: "受让方" }],
    });
    expect(plan.ops).toHaveLength(2);
    expect(lines(revert(plan.after, plan.inverses).after)).toEqual(lines(before));
  });

  it("skips a diagram when nothing can render it, and inserts it when something can", () => {
    const ops: Proposal["ops"] = [
      {
        op: "insertDiagram",
        after: 0,
        alt: "流程",
        diagram: { layout: "flow-v", nodes: [{ id: "a", label: "开始", shape: "box" }], edges: [] },
      },
    ];
    const before = doc("说明");
    expect(planEdit(before, { summary: "s", ops }).skipped.length).toBe(1);
    const withDiagram = planEdit(before, { summary: "s", ops }, diagram);
    expect(withDiagram.after[1]!.type).toBe("image");
    expect(lines(revert(withDiagram.after, withDiagram.inverses).after)).toEqual(lines(before));
  });

  it("keeps a table insert reversible", () => {
    const before = doc("概述");
    const plan = planEdit(
      before,
      {
        summary: "s",
        ops: [{ op: "insertTable", after: 0, rows: [["项", "值"], ["a", "b"]], headerRow: true }],
      },
      diagram,
    );
    expect(plan.after[1]!.type).toBe("table");
    expect(lines(revert(plan.after, plan.inverses).after)).toEqual(lines(before));
  });
});

describe("runOps", () => {
  it("collects inverses in undo order", () => {
    const before = doc("一", "二");
    const r = runOps(before, [
      { kind: "insertBlock", anchor: { at: "end" }, block: { type: "paragraph", text: "三" } },
      { kind: "setText", addr: { by: "blockText", text: "一", nth: 0 }, text: "一改" },
    ]);
    expect(r.skipped).toEqual([]);
    expect(lines(r.after)).toEqual(["paragraph:一改", "paragraph:二", "paragraph:三"]);
    expect(lines(revert(r.after, r.inverses).after)).toEqual(lines(before));
  });

  it("行内插引用只加域：正文没变，逆操作按指令文字认域", () => {
    const before = doc("见下图", "图 1 系统架构");
    const r = runOps(before, [
      {
        kind: "insertInline",
        addr: { by: "blockText", text: "见下图", nth: 0 },
        at: 3,
        json: { type: "field", attrs: { kind: "PAGEREF", instr: `PAGEREF _Ref1 \\h`, display: "5" } },
      },
    ]);
    expect(r.skipped).toEqual([]);
    // 域是行内原子、不贡献文字，所以块模型看到的还是原文——内容寻址的键没被这次写入改掉（§2.5）。
    expect(lines(r.after)).toEqual(lines(before));
    expect(r.inverses).toEqual([
      {
        kind: "removeInline",
        addr: { by: "blockText", text: "见下图", nth: 0 },
        instr: `PAGEREF _Ref1 \\h`,
      },
    ]);
    expect(revert(r.after, r.inverses).ok).toBe(true);
  });

  it("行内插入的偏移落在段外就不落地", () => {
    const r = runOps(doc("见下图"), [
      {
        kind: "insertInline",
        addr: { by: "blockText", text: "见下图", nth: 0 },
        at: 99,
        json: { type: "field", attrs: { kind: "REF", instr: `REF _Ref1`, display: "" } },
      },
    ]);
    expect(r.applied).toEqual([]);
    expect(r.inverses).toEqual([]);
    expect(r.skipped.join("")).toContain("偏移");
  });
});

describe("coerceProposal", () => {
  it("drops out-of-range indexes and unknown ops with reasons", () => {
    const r = coerceProposal(
      {
        summary: "s",
        ops: [
          { op: "setText", index: 99, text: "x" },
          { op: "nope" },
          { op: "setText", index: 0, text: "可用" },
        ],
      },
      2,
    );
    expect(r.proposal?.ops).toHaveLength(1);
    expect(r.errors).toHaveLength(2);
  });

  it("requires a summary-free payload to still produce a proposal", () => {
    const r = coerceProposal({ ops: [{ op: "deleteBlock", index: 0 }] }, 1);
    expect(r.proposal?.summary).toBeTruthy();
    expect(r.proposal?.ops).toHaveLength(1);
  });

  it("连续插入的编号递增不卡原文档段数", () => {
    // 写一整篇小说时模型一条一段地编号；只按原文档段数校验的话，48 条会被判成越权丢弃。
    const ops = Array.from({ length: 50 }, (_, i) => ({
      op: "insertParagraph",
      after: i - 1,
      text: `第 ${i + 1} 段正文`,
    }));
    const r = coerceProposal({ summary: "写小说", ops }, 1);
    expect(r.errors).toEqual([]);
    expect(r.proposal?.ops).toHaveLength(50);
  });

  it("段落样式接受带单位的磅值", () => {
    const r = coerceProposal(
      {
        summary: "s",
        ops: [
          { op: "insertParagraph", after: 0, text: "正文", style: { indentFirst: "24pt", lineHeight: "1.7" } },
        ],
      },
      1,
    );
    expect(r.errors).toEqual([]);
    expect(r.proposal?.ops[0]).toEqual({
      op: "insertParagraph",
      after: 0,
      text: "正文",
      style: { indentFirst: 24, lineHeight: "1.7" },
    });

    // "2cm" 不是段落格式的单位，宁可不写也不能猜。
    const unit = coerceProposal(
      { summary: "s", ops: [{ op: "setStyle", index: 0, style: { spacingBefore: "12 磅", indentRight: "2cm" } }] },
      1,
    );
    expect(unit.proposal?.ops[0]).toEqual({ op: "setStyle", index: 0, style: { spacingBefore: 12 } });
  });

  it("normalises ragged tables away instead of half-building them", () => {
    const r = coerceProposal(
      { summary: "s", ops: [{ op: "insertTable", after: 0, rows: [["a", "b"], ["c"]] }] },
      1,
    );
    expect(r.proposal).toBeNull();
    expect(r.errors.join()).toContain("rows");
  });
});

describe("alignRewrite：快捷指令没选区时对整篇逐段改写", () => {
  const table = (index: number): Block => ({
    index,
    type: "table",
    text: "甲 | 乙",
    rows: [["甲", "乙"]],
    style: {},
  });

  it("第 i 段落到第 i 个可改写块，表格与空段不参与对齐", () => {
    const blocks = [para(0, "第一段"), table(1), para(2, ""), para(3, "第四段")];
    const r = alignRewrite(blocks, "第一段（润色）\n第四段（润色）");
    expect(r.errors).toEqual([]);
    expect(r.ops).toEqual([
      { op: "setText", index: 0, text: "第一段（润色）" },
      { op: "setText", index: 3, text: "第四段（润色）" },
    ]);
    // 送出去的原文与对齐口径必须同源，否则两边段数永远对不上。
    expect(rewritableBlocks(blocks).map((b) => b.text).join("\n")).toBe("第一段\n第四段");
  });

  it("段数对不上就整批放弃，既不硬塞也不替用户删段", () => {
    const r = alignRewrite(doc("甲", "乙", "丙"), "甲改写\n乙改写");
    expect(r.ops).toEqual([]);
    expect(r.errors.join()).toContain("段数对不上");
  });

  it("文档里没有可改写的文字时给原因，而不是空提案", () => {
    expect(alignRewrite(doc("", "  "), "随便写点什么").errors.join()).toContain("没有可以改写的文字");
  });

  it("逐段 setText 走提案管道：版式不动，撤销能回到原样", () => {
    const blocks = [para(0, "第一段"), table(1), para(2, "第三段")];
    const { ops } = alignRewrite(blocks, "第一段改写\n第三段改写");
    const plan = planEdit(blocks, { summary: "全文润色", ops });
    expect(plan.ok).toBe(true);
    expect(lines(plan.after)).toEqual([
      "paragraph:第一段改写",
      "table:甲 | 乙",
      "paragraph:第三段改写",
    ]);
    expect(lines(revert(plan.after, plan.inverses).after)).toEqual(lines(blocks));
  });
});

describe("rewriteResume：暂停后续写时只补剩下的段落", () => {
  it("断在段中间的那半截丢掉重写，段数才不会对不上", () => {
    const source = "第一段\n第二段\n第三段";
    const r = rewriteResume(source, "第一段改写\n第二段改写\n第三段（半");
    expect(r.kept).toEqual(["第一段改写", "第二段改写"]);
    expect(r.remaining).toBe(1);
  });

  it("半截只有一句残句时无段可留，等于整篇重来", () => {
    const r = rewriteResume("第一段\n第二段", "第一段改写（写到一半");
    expect(r.kept).toEqual([]);
    expect(r.remaining).toBe(2);
  });

  it("半截看起来收满了也照样丢掉末段，宁可重写一段也不多要一段", () => {
    const r = rewriteResume("第一段\n第二段\n第三段", "甲\n乙\n丙");
    expect(r.kept).toEqual(["甲", "乙"]);
    expect(r.remaining).toBe(1);
  });

  it("留下的段落 + 模型补的那一段拼回去，落地时段数正好齐全", () => {
    const source = "第一段\n第二段\n第三段";
    const r = rewriteResume(source, "第一段改写\n第二段改写\n第三段公");
    const merged = `${r.kept.join("\n")}\n第三段改写`;
    const aligned = alignRewrite(doc(...source.split("\n")), merged);
    expect(aligned.errors).toEqual([]);
    expect(aligned.ops).toEqual([
      { op: "setText", index: 0, text: "第一段改写" },
      { op: "setText", index: 1, text: "第二段改写" },
      { op: "setText", index: 2, text: "第三段改写" },
    ]);
    // 反过来看不丢掉残句的接法：段数照样对得上，但第三段是"半截 + 重写"的缝合怪，
    // 会当成新文字写进文档——所以续写必须先丢掉可能被切断的那一段。
    expect(alignRewrite(doc(...source.split("\n")), "第一段改写\n第二段改写\n第三段公第三段改写").ops[2]).toEqual({
      op: "setText",
      index: 2,
      text: "第三段公第三段改写",
    });
  });

  it("半截里有空行也不影响计数：折叠口径与 alignRewrite 同源", () => {
    const r = rewriteResume("第一段\n第二段\n第三段", "第一段改写\n\n   \n第二段改");
    expect(r.kept).toEqual(["第一段改写"]);
    expect(r.remaining).toBe(2);
  });
});
