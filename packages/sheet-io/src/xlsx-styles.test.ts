/**
 * `xl/styles.xml` 的只追加改写单测（钉住 `AGENTS.md` §2.3「不许把局部格式拍平」与 §二 P3「位置索引不许重排」）。
 */
import { describe, expect, it } from "vitest";
import { StyleTable } from "./xlsx-styles.js";

/** 一条带字体/底色/边框/对齐的样式，专门用来验证"换数字格式时别的一律不动"。 */
const STYLES_WITH_ALIGNED_XF =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
  '<cellXfs count="2">' +
  '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
  '<xf numFmtId="0" fontId="1" fillId="2" borderId="3" xfId="0" applyFont="1"><alignment horizontal="center"/></xf>' +
  "</cellXfs></styleSheet>";

describe("StyleTable.withNumberFormat（只改数字格式，其余随格带走）", () => {
  it("克隆 base 那条 xf：字体/底色/边框/applyFont 与 `<alignment/>` 子节点全部原样保留", () => {
    const styles = StyleTable.parse(STYLES_WITH_ALIGNED_XF);
    const next = styles.withNumberFormat(1, { numFmtId: 10 });

    expect(next).toBe(2);
    expect(styles.toXml()).toContain(
      '<xf numFmtId="10" fontId="1" fillId="2" borderId="3" xfId="0" applyFont="1" applyNumberFormat="1"><alignment horizontal="center"/></xf>',
    );
    // §二 P3：新条目只能追加在末尾，前两条一个字节都不许变
    expect(styles.cellXfCount).toBe(3);
  });

  it("同一格二次改到同一格式不许再加条目；本来就是该格式时返回自己", () => {
    const styles = StyleTable.parse(STYLES_WITH_ALIGNED_XF);
    const first = styles.withNumberFormat(1, { numFmtId: 10 });
    expect(styles.withNumberFormat(1, { numFmtId: 10 })).toBe(first);
    expect(styles.withNumberFormat(first, { numFmtId: 10 })).toBe(first);
    expect(styles.cellXfCount).toBe(3);
    // base 的 numFmtId 已经等于目标：连样式表都不该被标脏
    const untouched = StyleTable.parse(STYLES_WITH_ALIGNED_XF);
    expect(untouched.withNumberFormat(1, { numFmtId: 0 })).toBe(1);
    expect(untouched.isDirty).toBe(false);
  });

  it("自定义格式码先落到 `<numFmts>`，样式仍是从 base 克隆来的", () => {
    const styles = StyleTable.parse(STYLES_WITH_ALIGNED_XF);
    const next = styles.withNumberFormat(1, { formatCode: '#,##0.0"元"' });

    expect(styles.numberFormatFor(next)).toMatchObject({ formatCode: '#,##0.0"元"' });
    expect(styles.toXml()).toContain('numFmtId="164" fontId="1" fillId="2" borderId="3"');
  });

  it("越界的 `s` 按第 0 条（General）算，不许凭空造出带字体的样式", () => {
    const styles = StyleTable.parse(STYLES_WITH_ALIGNED_XF);
    const next = styles.withNumberFormat(99, { numFmtId: 10 });

    expect(styles.fontFor(next)).toEqual({});
    expect(styles.toXml()).toContain('<xf numFmtId="10" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>');
  });
});
