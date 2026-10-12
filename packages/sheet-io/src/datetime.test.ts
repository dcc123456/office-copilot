import { describe, expect, it } from "vitest";
import { isDateFormatCode, serialToUtcDate, utcDateToSerial } from "./datetime.js";

describe("日期序列号换算（实测基准：spec §二 P3）", () => {
  it("实测样本里的 D2 是 46053，对应 2026-01-31", () => {
    expect(serialToUtcDate(46053)?.toISOString().slice(0, 10)).toBe("2026-01-31");
    expect(utcDateToSerial(new Date(Date.UTC(2026, 0, 31)))).toBe(46053);
    expect(utcDateToSerial(new Date(Date.UTC(2026, 1, 28)))).toBe(46081);
  });

  it("1900 闰年 bug：第 60 天是不存在的假日期，59/61 仍连续", () => {
    expect(serialToUtcDate(1)?.toISOString().slice(0, 10)).toBe("1900-01-01");
    expect(serialToUtcDate(59)?.toISOString().slice(0, 10)).toBe("1900-02-28");
    expect(serialToUtcDate(60)).toBeUndefined();
    expect(serialToUtcDate(61)?.toISOString().slice(0, 10)).toBe("1900-03-01");
    expect(utcDateToSerial(new Date(Date.UTC(1900, 0, 1)))).toBe(1);
    expect(utcDateToSerial(new Date(Date.UTC(1900, 1, 28)))).toBe(59);
    expect(utcDateToSerial(new Date(Date.UTC(1900, 2, 1)))).toBe(61);
  });

  it("小数部分是一天内的时间；非法输入不返回猜测的日期", () => {
    expect(serialToUtcDate(46053.5)?.toISOString()).toBe("2026-01-31T12:00:00.000Z");
    expect(serialToUtcDate(Number.NaN)).toBeUndefined();
    expect(serialToUtcDate(-1)).toBeUndefined();
  });

  it("引号里的 yyyy 是字面量，不能判成日期格式", () => {
    expect(isDateFormatCode("yyyy-mm-dd")).toBe(true);
    expect(isDateFormatCode("h:mm")).toBe(true);
    expect(isDateFormatCode('"yyyy"')).toBe(false);
    expect(isDateFormatCode("#,##0")).toBe(false);
    expect(isDateFormatCode(undefined)).toBe(false);
  });
});
