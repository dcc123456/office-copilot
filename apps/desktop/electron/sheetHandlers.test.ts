/**
 * 保存前那道审计闸门的判定逻辑：只拦"这次改动新引入的"结构问题。
 *
 * 拦下后主进程整批不写盘（`AGENTS.md` §2.4），渲染层把 `message` 弹成红 toast，
 * 所以这里要钉的是"什么该拦"，不是文案。
 */
import { describe, expect, it, vi } from "vitest";
import type { CompatIssue } from "@office-copilot/sheet-io";
import { newlyIntroducedIssues } from "./sheetHandlers.js";

vi.mock("electron", () => ({ dialog: {} }));

/** 造一条问题记录。 */
function issue(code: CompatIssue["code"], part: string, detail = ""): CompatIssue {
  return { code, part, detail };
}

describe("保存前的结构审计只对本次事务负责", () => {
  it("底包本来就有的问题不拦（第三方导出的文件不该让人存不了盘）", () => {
    const base = [issue("part-not-declared", "xl/printerSettings.bin", "bin")];
    expect(newlyIntroducedIssues(base, base)).toEqual([]);
    expect(newlyIntroducedIssues(base, [...base, issue("unresolved-relationship-id", "xl/worksheets/sheet1.xml", "rId7")])).toEqual([
      issue("unresolved-relationship-id", "xl/worksheets/sheet1.xml", "rId7"),
    ]);
  });

  it("同一个码落在不同部件上是两条问题，不许互相抵掉", () => {
    const base = [issue("worksheet-child-order", "xl/worksheets/sheet1.xml", "pageMargins -> dimension")];
    const out = [...base, issue("worksheet-child-order", "xl/worksheets/sheet2.xml", "pageMargins -> dimension")];
    expect(newlyIntroducedIssues(base, out)).toEqual([issue("worksheet-child-order", "xl/worksheets/sheet2.xml", "pageMargins -> dimension")]);
  });

  it("输出干净时什么都不拦", () => {
    expect(newlyIntroducedIssues([issue("chart-missing-chart", "xl/charts/chart1.xml")], [])).toEqual([]);
  });
});
