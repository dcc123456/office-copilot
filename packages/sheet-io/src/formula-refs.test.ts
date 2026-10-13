import { describe, expect, it } from "vitest";
import { shiftFormula } from "./formula-refs.js";

/**
 * 断言"平移后的公式串"。
 *
 * @param formula 原公式体（不含前导 `=`）。
 * @param axis 动的轴。
 * @param from 0 基插入点/被删带起点。
 * @param delta 位移量。
 * @returns 成功时的新公式；拒绝时返回拒绝理由（用例据此断言）。
 */
function shifted(formula: string, axis: "row" | "col", from: number, delta: number): string {
  const result = shiftFormula(formula, axis, from, delta);
  if (!result.ok) throw new Error(result.reason);
  return result.formula;
}

describe("插入行列时的引用平移（`$` 只钉复制，不钉插删）", () => {
  it("在第 3 行之上插一行：范围下端的相对引用跟着下移", () => {
    // `SUM(C2:C4)` 指第 2–4 行；在第 3 行（0 基 2）插入 → C2 在上半段不动，C4 下移成 C5
    expect(shifted("SUM(C2:C4)", "row", 2, 1)).toBe("SUM(C2:C5)");
    expect(shifted("SUM(C2:C4)", "row", 1, 1)).toBe("SUM(C3:C5)");
  });

  it("插入点之上的绝对引用照样挪，但 `$` 的写法必须逐部分保留", () => {
    // `$C$2` 在插入点之上（第 2 行），不受影响；`C4` 在插入点之下，下移成 C5
    expect(shifted("SUM($C$2:C4)", "row", 2, 1)).toBe("SUM($C$2:C5)");
    // 第 1 行上面插两行：绝对引用指的那两个格本身被推下去了，所以行号一起加
    expect(shifted("$A$1+B2", "row", 0, 2)).toBe("$A$3+B4");
  });

  it("列轴同理：插入点之前的列不动，之后的列右移（混合引用只挪钉不住的那半）", () => {
    expect(shifted("SUM(A1:C1)", "col", 1, 1)).toBe("SUM(A1:D1)");
    expect(shifted("$A$1&B2", "col", 0, 1)).toBe("$B$1&C2");
    // `$A1` 钉的是列，可这次动的是列，所以它照移；行号那半永远不动
    expect(shifted("$A1&B2", "col", 0, 1)).toBe("$B1&C2");
    // 动的是行时，钉住列的 `$A1` 里列号不变、行号照移
    expect(shifted("$A1+B$2", "row", 0, 1)).toBe("$A2+B$3");
  });

  it("函数名与字符串字面量都不许当成引用", () => {
    expect(shifted("LOG10(A2)+ATAN2(A1,B1)", "row", 0, 1)).toBe("LOG10(A3)+ATAN2(A2,B2)");
    expect(shifted('IF(A1="B2",A3,0)', "row", 1, 1)).toBe('IF(A1="B2",A4,0)');
  });

  it("整列引用本来就指整列，插删不改它", () => {
    expect(shifted("SUM(C:C)", "row", 0, 1)).toBe("SUM(C:C)");
  });

  it("没有引用受影响时原串带过（不许顺手改大小写）", () => {
    expect(shifted("sum(c4)", "row", 4, 1)).toBe("sum(c4)");
  });
});

describe("删除行列时的拦下规则（我们不产 #REF!）", () => {
  /**
   * 断言删除路径的结果。
   *
   * @param formula 原公式体。
   * @param axis 动的轴。
   * @param startIndex 被删带起点（0 基）。
   * @param count 删除条数。
   * @returns 成功给新公式，拒绝给理由。
   */
  function deleted(formula: string, axis: "row" | "col", startIndex: number, count: number): string {
    const result = shiftFormula(formula, axis, startIndex, -count, { startIndex, endIndex: startIndex + count - 1 });
    return result.ok ? result.formula : result.reason;
  }

  it("端点落在被删区间里时整条拒绝，包括绝对引用", () => {
    // 第 3 行（0 基 2）正是 `C3` 指的那一格：Excel 这时给 #REF!，我们整条不动
    expect(deleted("SUM(C3:C4)", "row", 2, 1)).toContain("#REF!");
    expect(deleted("SUM($C$3:$C$4)", "row", 2, 1)).toContain("#REF!");
    expect(deleted("B2*C2", "col", 1, 1)).toContain("#REF!");
  });

  it("删掉区间中间一行时区间收缩、带外的引用上移，与 Excel 一致", () => {
    // 两个端点都不在被删带里，只是下端上移一格
    expect(deleted("SUM(C2:C4)", "row", 2, 1)).toBe("SUM(C2:C3)");
    expect(deleted("SUM($C$2:$C$4)", "row", 2, 1)).toBe("SUM($C$2:$C$3)");
    expect(deleted("SUM(C5:C6)", "row", 2, 1)).toBe("SUM(C4:C5)");
    expect(deleted("SUM(C5:C$9)", "row", 2, 1)).toBe("SUM(C4:C$8)");
  });
});

describe("认不出的写法一律拒绝（位移要连表名一起解析，属阶段 4）", () => {
  it.each([
    ["跨表引用", "Sheet2!A1"],
    ["外部工作簿引用", "'[1]Sheet1'!$A$2"],
    ["结构化表引用", "SUM(Table1[金额])"],
    ["数组常量", "SUM({1,2})"],
    ["已经是坏引用", "#REF!+A1"],
  ])("%s", (_label, formula) => {
    expect(shiftFormula(formula, "row", 0, 1)).toMatchObject({ ok: false });
  });

  it("名字尾巴长得像引用时也拒绝，不许当成 A1 挪走", () => {
    expect(shiftFormula("1E5+A1", "row", 0, 1)).toMatchObject({ ok: false });
    expect(shiftFormula("Table1!A1", "row", 0, 1)).toMatchObject({ ok: false });
  });
});
