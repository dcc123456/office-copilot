import { describe, expect, it } from "vitest";
import { diffBlocks, diffStats, inlineDiff, toGitLines, tokenize } from "./diff.js";

describe("tokenize", () => {
  it("splits Chinese per character and Latin per word", () => {
    expect(tokenize("修订 文档")).toEqual(["修", "订", " ", "文", "档"]);
    expect(tokenize("hello world")).toEqual(["hello", " ", "world"]);
  });
});

describe("inlineDiff", () => {
  it("marks only the changed characters", () => {
    const parts = inlineDiff("我们讨论一下", "我们讨论下");
    expect(parts.filter((p) => p.t !== "eq").map((p) => `${p.t}:${p.s}`).join("|")).toBe("del:一");
    expect(parts.filter((p) => p.t !== "del").map((p) => p.s).join("")).toBe("我们讨论下");
  });

  it("round-trips both sides", () => {
    const before = "本合同自双方签字之日起生效。";
    const after = "本合同自双方签署之日起生效，并加盖公章。";
    const parts = inlineDiff(before, after);
    expect(parts.filter((p) => p.t !== "del").map((p) => p.s).join("")).toBe(after);
    expect(parts.filter((p) => p.t !== "ins").map((p) => p.s).join("")).toBe(before);
  });

  it("degrades to a whole-line diff past the cell budget", () => {
    const a = "甲".repeat(700);
    const b = "乙".repeat(700);
    const parts = inlineDiff(a, b);
    expect(parts).toHaveLength(2);
    expect(parts[0]).toMatchObject({ t: "del", s: a });
    expect(parts[1]).toMatchObject({ t: "ins", s: b });
  });
});

describe("diffBlocks", () => {
  it("folds a rewritten paragraph into one mod row", () => {
    const rows = diffBlocks(["前言", "第二节 旧内容"], ["前言", "第二节 新内容"]);
    expect(rows.map((r) => r.kind)).toEqual(["eq", "mod"]);
    expect(diffStats(rows)).toEqual({ added: 0, removed: 0, changed: 1 });
  });

  it("counts pure additions and deletions", () => {
    const rows = diffBlocks(["一", "二", "三"], ["一", "三", "四"]);
    expect(diffStats(rows)).toEqual({ added: 1, removed: 1, changed: 0 });
  });

  it("reports nothing changed for a format-only edit", () => {
    const rows = diffBlocks(["同"], ["同"]);
    expect(rows).toEqual([{ kind: "eq", before: "同", after: "同" }]);
    expect(diffStats(rows)).toEqual({ added: 0, removed: 0, changed: 0 });
  });
});

describe("toGitLines", () => {
  it("puts the old line before the new one and numbers both sides", () => {
    const before = ["一", "二", "三", "四", "五", "六", "七", "八"];
    const after = ["一", "二", "三", "四", "五", "六", "七改", "八"];
    const lines = toGitLines(before, after, 2);
    expect(lines.map((l) => l.kind)).toEqual(["omit", "ctx", "ctx", "del", "add", "ctx"]);
    expect(lines[0]!.omitted).toBe(4);
    const del = lines.find((l) => l.kind === "del")!;
    const add = lines.find((l) => l.kind === "add")!;
    expect(del.beforeNo).toBe(7);
    expect(del.afterNo).toBeUndefined();
    expect(add.afterNo).toBe(7);
    expect(add.beforeNo).toBeUndefined();
    expect(del.parts.map((p) => p.s).join("")).toBe("七");
    expect(add.parts.map((p) => p.s).join("")).toBe("七改");
  });

  it("keeps a whole run when it is short enough for the context", () => {
    const lines = toGitLines(["一", "二"], ["一", "新增", "二"]);
    expect(lines.map((l) => l.kind)).toEqual(["ctx", "add", "ctx"]);
    expect(lines[1]).toMatchObject({ afterNo: 2, parts: [{ t: "ins", s: "新增" }] });
  });

  it("collapses nothing when the text is identical", () => {
    const lines = toGitLines(["同", "样"], ["同", "样"]);
    expect(lines.map((l) => l.kind)).toEqual(["ctx", "ctx"]);
    expect(lines.map((l) => l.beforeNo)).toEqual([1, 2]);
  });
});
