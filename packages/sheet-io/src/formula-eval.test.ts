import { describe, expect, it } from "vitest";
import { evaluateFormula } from "./formula-eval.js";
import type { SheetCell, SheetLike as SheetPayload } from "./model.js";

/** 求值用的表：数字一列、文本一列、一个"看着像数字"的文本格、一个已有公式格（验递归）。 */
const PART = "xl/worksheets/sheet1.xml";

/**
 * 造一格。
 *
 * @param address 大写 A1 地址。
 * @param value 值；字符串按文本格，数字按数值格。
 * @param formula 公式体（不含前导 `=`）；给出时同时塞一个过期缓存值，验求值不许信缓存。
 * @returns 一个格子。
 */
function cellOf(address: string, value: string | number | null, formula?: string): SheetCell {
  const match = /^([A-Z]+)([0-9]+)$/.exec(address)!;
  const colIndex = match[1]!.split("").reduce((acc, char) => acc * 26 + char.charCodeAt(0) - 64, 0) - 1;
  return {
    address,
    colIndex,
    rowIndex: Number(match[2]) - 1,
    kind: typeof value === "string" ? "string" : "number",
    value,
    ...(formula ? { formula, cachedValue: 999999 } : {}),
  };
}

/** @returns 求值用的最小表载荷。 */
function sheetOf(cells: SheetCell[]): SheetPayload {
  return {
    name: "销售",
    sheetId: 1,
    partPath: PART,
    lastColIndex: cells.reduce((acc, cell) => Math.max(acc, cell.colIndex), -1),
    lastRowIndex: cells.reduce((acc, cell) => Math.max(acc, cell.rowIndex), -1),
    cells: Object.fromEntries(cells.map((cell) => [cell.address, cell])),
    merges: [],
    columns: [],
    rows: [],
    warnings: [],
  };
}

const sheet = sheetOf([
  cellOf("A2", 10),
  cellOf("A3", 20),
  cellOf("A4", 30),
  cellOf("B2", "华东"),
  cellOf("B3", "华北"),
  cellOf("B4", "华东"),
  cellOf("C2", null, "SUM(A2:A4)"),
  cellOf("D2", "15"),
  cellOf("E2", "abc"),
]);

/**
 * 求值并取值。
 *
 * @param formula 公式体。
 * @returns 成功时的值；失败时返回 `FAIL:理由`，让断言直接读出哪条没走通。
 */
function valueOf(formula: string): unknown {
  const result = evaluateFormula(formula, sheet);
  return result.ok ? result.value : `FAIL:${result.reason}`;
}

/**
 * 求值并取拒绝理由。
 *
 * @param formula 公式体。
 * @returns `ok:false` 的理由；算出来了返回 `算出来了：<值>`。
 */
function reasonOf(formula: string): string {
  const result = evaluateFormula(formula, sheet);
  return result.ok ? `算出来了：${String(result.value)}` : result.reason;
}

describe("统计函数（spec D4 白名单）", () => {
  it("SUM/AVERAGE/MIN/MAX 只吃数字，区间里的文本与空格跳过（Excel 口径）", () => {
    expect(valueOf("SUM(A2:A4)")).toBe(60);
    expect(valueOf("AVERAGE(A2:A4)")).toBe(20);
    expect(valueOf("MIN(A2:A4)")).toBe(10);
    expect(valueOf("MAX(A2:A4)")).toBe(30);
    expect(valueOf("SUM(B2:B4,A2:A3)")).toBe(30);
    // "15" 是文本格：Excel 按类型跳过，不因为我们能解析出数字就把它计进来
    expect(valueOf("SUM(A2:A4,D2)")).toBe(60);
    expect(valueOf("MIN(B2:B4)")).toBe(0);
  });

  it("COUNT 数的是数字格数（=Excel 的 COUNT，文本不计）", () => {
    expect(valueOf("COUNT(A2:A4)")).toBe(3);
    expect(valueOf("COUNT(B2:B4)")).toBe(0);
    expect(valueOf("COUNT(A2:A4,D2)")).toBe(3);
  });

  it("COUNTIF 认得比较条件与文本条件，文本比较不区分大小写", () => {
    expect(valueOf('COUNTIF(A2:A4,">15")')).toBe(2);
    expect(valueOf('COUNTIF(A2:A4,">=20")')).toBe(2);
    expect(valueOf('COUNTIF(A2:A4,"<>20")')).toBe(2);
    expect(valueOf('COUNTIF(B2:B4,"华东")')).toBe(2);
    expect(valueOf('COUNTIF(E2,"ABC")')).toBe(1);
  });

  it("SUMIF 第三段缺省时对条件区间自身求和，求和区间里的文本跳过", () => {
    expect(valueOf('SUMIF(A2:A4,">=20")')).toBe(50);
    expect(valueOf('SUMIF(B2:B4,"华东",A2:A4)')).toBe(40);
    expect(valueOf('SUMIF(A2:A4,"99",A2:A4)')).toBe(0);
    expect(valueOf('SUMIF(A2:A4,">5",D2)')).toBe(0);
  });
});

describe("算术与文本函数", () => {
  it("优先级与 Excel 一致：乘除高于加减、^ 右结合且高于乘法、一元负号紧贴操作数", () => {
    expect(valueOf("2+3*4")).toBe(14);
    expect(valueOf("2^3^2")).toBe(512);
    expect(valueOf("-2^2")).toBe(4);
    expect(valueOf("(2+3)*4")).toBe(20);
    expect(valueOf("7/2")).toBe(3.5);
    expect(valueOf("100*50%")).toBe(50);
  });

  it("ROUND 是离零方向的四舍五入，位数可为负；ABS 取绝对值", () => {
    expect(valueOf("ROUND(2.5,0)")).toBe(3);
    expect(valueOf("ROUND(-2.5,0)")).toBe(-3);
    expect(valueOf("ROUND(3.14159,2)")).toBe(3.14);
    expect(valueOf("ROUND(1234,-2)")).toBe(1200);
    expect(valueOf("ABS(-7.5)")).toBe(7.5);
  });

  it("IF 按条件取分支，缺省假分支为 FALSE；比较结果直接当条件用", () => {
    expect(valueOf('IF(A2>5,"大","小")')).toBe("大");
    expect(valueOf('IF(A4>100,"大","小")')).toBe("小");
    expect(valueOf('IF(A2>100,"大")')).toBe(false);
    expect(valueOf('IF("TRUE",1,2)')).toBe(1);
    expect(valueOf("IF(A2,1,2)")).toBe(1);
    // Excel 的排序规则里文本永远大于数字，比较不报 #VALUE!
    expect(valueOf('IF(B2>5,"文本大","数字大")')).toBe("文本大");
  });

  it("CONCAT 与 & 把值拼成文本，布尔显示成 TRUE/FALSE、空格是空串", () => {
    expect(valueOf('CONCAT(B2,"-",A2)')).toBe("华东-10");
    expect(valueOf('CONCATENATE("a",2>1,"b")')).toBe("aTRUEb");
    expect(valueOf('B2&"!"&C2')).toBe("华东!60");
    expect(valueOf("B2&Z9")).toBe("华东");
  });
});

describe("引用取数", () => {
  it("空格是 null、文本按内容取，`$` 绝对引用取数前就剥掉", () => {
    expect(valueOf("A2")).toBe(10);
    expect(valueOf("$A$2+1")).toBe(11);
    expect(valueOf("Z9")).toBeNull();
    expect(valueOf("B2")).toBe("华东");
  });

  it("带公式的格递归求值，不许信文件里的过期缓存值", () => {
    expect(valueOf("C2")).toBe(60);
    expect(valueOf("SUM(A2:A4)+C2")).toBe(120);
  });

  it("自我引用判成循环引用（宁可不写缓存值）", () => {
    expect(valueOf("A2+A3")).toBe(30);
    const circular = evaluateFormula("C2+1", sheet, "C2");
    expect(circular.ok).toBe(false);
    if (!circular.ok) expect(circular.reason).toContain("循环引用");
  });
});

describe("拒绝的结构：一律 {ok:false}，调用方据此留空 `<v>`（验收 A4）", () => {
  it("volatile 函数每次都变，写了缓存值就是写一个过期数字", () => {
    expect(reasonOf("TODAY()")).toContain("每次都变");
    expect(reasonOf("NOW()")).toContain("每次都变");
    expect(reasonOf("RAND()")).toContain("每次都变");
  });

  it("跨表、外部工作簿、数组常量、错误值在切词阶段就拒绝", () => {
    expect(reasonOf("Sheet2!A2+1")).toContain("不解释的结构");
    expect(reasonOf("'[1]Sheet1'!$A$2")).toContain("不解释的结构");
    expect(reasonOf("{1,2,3}")).toContain("不解释的结构");
    expect(reasonOf("#DIV/0!")).toContain("不解释的结构");
  });

  it("白名单外的函数交给 Excel 重算，不假装算得出", () => {
    expect(reasonOf("VLOOKUP(B2,A2:B4,2)")).toContain("还没实现");
    expect(reasonOf("IFS(A2>1,1)")).toContain("还没实现");
    expect(reasonOf("SUMM(A2)")).toContain("SUMM");
  });

  it("文本进算术、除零、空区间求平均，给的是 Excel 那三个错误名的理由", () => {
    expect(reasonOf('A2+"abc"')).toContain("#VALUE!");
    expect(reasonOf("A2/0")).toContain("#DIV/0!");
    expect(reasonOf("AVERAGE(B2:B4)")).toContain("#DIV/0!");
  });

  it("整个区间放到要单个数字的位置上时拒绝，不取第一个凑数", () => {
    expect(reasonOf("A2:A4+1")).toContain("#VALUE!");
    expect(reasonOf("(A2:A4)")).toContain("#VALUE!");
  });

  it("括号没闭合、末尾有多余内容、空实参都算没走通", () => {
    expect(reasonOf("SUM(A2,A3")).toContain("没闭合");
    expect(reasonOf("1+2 3")).toContain("没吃干净");
    expect(reasonOf("SUM(A2:A4,)")).toContain("认不出");
  });

  it("求值器永不抛异常：垃圾输入也得是一个错误串（验收 B9）", () => {
    for (const garbage of ["", "   ", "((((", '"未闭合', "1+", "/*", "@", "IF(", ",,,", "()", "SUM("]) {
      const result = evaluateFormula(garbage, sheet, "F5");
      expect(result.ok).toBe(false);
      if (!result.ok) expect(result.reason.length).toBeGreaterThan(0);
    }
  });
});
