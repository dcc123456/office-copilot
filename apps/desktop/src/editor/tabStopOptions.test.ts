import { describe, expect, it } from "vitest";
import {
  TAB_ALIGNMENTS,
  TAB_LEADERS,
  alignmentLabelOf,
  leaderLabelOf,
  stopLabelOf,
} from "./tabStopOptions.js";

/**
 * 制表位的候选表与清单一行的文案。位置在段落属性里是 pt，界面按 Word 的口径显示成
 * 厘米——换算出自 docx-io（AGENTS §2.1），这里只负责显示。
 */

describe("制表位候选表", () => {
  it("前导符七档齐全，对齐只给 Word 对话框里点得出的五档", () => {
    expect(TAB_LEADERS.map((o) => o.value)).toEqual([
      "none",
      "dot",
      "hyphen",
      "dash",
      "line",
      "heavy",
      "middleDot",
    ]);
    expect(TAB_ALIGNMENTS.map((o) => o.value)).toEqual([
      "left",
      "center",
      "right",
      "decimal",
      "bar",
    ]);
  });

  it("列表制表位与取消位显示得出，却不在候选里", () => {
    // 编号生成的 `w:val="list"` 在导入文档里很常见：读得出却不显示成空白，
    // 但界面造不出第二条——那是 Word 的编号排版自己写的。
    expect(alignmentLabelOf("list")).toBe("列表制表位");
    expect(alignmentLabelOf("clear")).toBe("取消的制表位");
    expect(TAB_ALIGNMENTS.some((o) => o.value === "list")).toBe(false);
    // 表外值退回原词，导入文档里读不懂的写法不至于显示空白。
    expect(alignmentLabelOf("raised" as never)).toBe("raised");
  });

  it("没写 w:leader 与写着 none 都是 Word 的“无”", () => {
    expect(leaderLabelOf(undefined)).toBe("（无）");
    expect(leaderLabelOf("none")).toBe("（无）");
  });
});

describe("清单里一行的文案", () => {
  it("位置按厘米显示两位小数，后面跟对齐与前导符", () => {
    expect(stopLabelOf({ positionPt: 415.3, alignment: "right", leader: "dot" })).toBe(
      "14.65 厘米 右对齐 点线 …",
    );
    expect(stopLabelOf({ positionPt: 56.7, alignment: "left" })).toBe(
      "2.00 厘米 左对齐 （无）",
    );
  });
});
