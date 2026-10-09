import { describe, it, expect } from "vitest";
import {
  ReplaceTextArgs,
  InsertParagraphArgs,
  InsertTableArgs,
  SubmitProposalArgs,
  editsFromProposal,
  WORD_TOOLS,
} from "./index.js";

describe("word-tools schemas", () => {
  it("accepts valid replace_text args", () => {
    const parsed = ReplaceTextArgs.parse({ find: "甲方", replace: "客户" });
    expect(parsed.regex).toBe(false);
    expect(parsed.scope).toBe("body");
  });

  it("rejects empty find", () => {
    expect(() => ReplaceTextArgs.parse({ find: "", replace: "x" })).toThrow();
  });

  it("defaults insert_paragraph position", () => {
    const p = InsertParagraphArgs.parse({ text: "hello" });
    expect(p.position).toBe("end");
  });

  it("rejects empty table", () => {
    expect(() => InsertTableArgs.parse({ rows: [] })).toThrow();
  });

  it("accepts a well-formed proposal", () => {
    const args = SubmitProposalArgs.parse({
      summary: "替换术语",
      edits: [
        { type: "replace_text", find: "甲方", replace: "客户" },
        { type: "insert_paragraph", text: "附加说明", position: "end" },
        { type: "insert_table", rows: [["A", "B"], ["1", "2"]] },
      ],
    });
    const edits = editsFromProposal(args);
    expect(edits).toHaveLength(3);
    expect(edits[0]!.type).toBe("replace_text");
  });

  it("exposes OpenAI tool definitions", () => {
    expect(WORD_TOOLS).toHaveLength(2);
    expect(WORD_TOOLS[0]!.function.name).toBe("word_list_paragraphs");
  });
});
