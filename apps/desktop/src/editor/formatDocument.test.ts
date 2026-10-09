import { describe, expect, it } from "vitest";
import { planEdit, revert, type Block } from "../ai/docEdit.js";
import { typographyProposal } from "./formatDocument.js";

const heading = (index: number, text: string, level: number): Block => ({
  index,
  type: "heading",
  text,
  style: { headingLevel: level },
});
const para = (index: number, text: string, style: Partial<Block["style"]> = {}): Block => ({
  index,
  type: "paragraph",
  text,
  style,
});

describe("typographyProposal", () => {
  it("styles headings and plain body text, and leaves hand-formatted text alone", () => {
    const blocks: Block[] = [
      heading(0, "一级标题", 1),
      para(1, "正文段落"),
      para(2, "手改过的正文", { fontSize: "20px" }),
      heading(3, "二级标题", 2),
    ];
    const p = typographyProposal(blocks);
    expect(p.ops.map((o) => "index" in o ? o.index : -1)).toEqual([0, 1, 3]);
    expect(p.summary).toContain("统一格式");
    expect(planEdit(blocks, p).after[0]!.style.align).toBe("center");
    expect(planEdit(blocks, p).after[3]!.style.fontSize).toBe("18px");
  });

  it("is idempotent: applying it once leaves nothing to change", () => {
    const blocks: Block[] = [heading(0, "标题", 1), para(1, "正文")];
    const first = planEdit(blocks, typographyProposal(blocks));
    const second = typographyProposal(first.after);
    expect(second.ops).toEqual([]);
  });

  it("is undoable as a single modification", () => {
    const blocks: Block[] = [heading(0, "标题", 1), para(1, "正文"), para(2, "另一段")];
    const plan = planEdit(blocks, typographyProposal(blocks));
    const back = revert(plan.after, plan.inverses);
    expect(back.ok).toBe(true);
    expect(back.after.map((b) => b.style)).toEqual(blocks.map((b) => b.style));
    expect(back.after.map((b) => b.text)).toEqual(["标题", "正文", "另一段"]);
  });

  it("never rewrites text", () => {
    const blocks: Block[] = [para(0, "内容"), heading(1, "题", 3)];
    const p = typographyProposal(blocks);
    expect(p.ops.every((o) => o.op === "setStyle")).toBe(true);
  });
});
