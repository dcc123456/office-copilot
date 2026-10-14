/**
 * 单位换算的机检（`AGENTS.md` §2.1 / spec D4）：`.pptx` 里 EMU 与半磅 ↔ CSS px 与 pt 只在 `units.ts` 发生。
 *
 * 数字取自本机真实幻灯片：4:3 页面宽 9144000 EMU（= 960px @96DPI）、标题框常见 18pt（= sz 1800）。
 */
import { describe, expect, it } from "vitest";
import { emuToPx, halfPointToPt, pxToEmu, ptToHalfPoint } from "./units.js";

describe("pptx-io 单位换算", () => {
  it("EMU ↔ px 用 9525 这一系数往返一致", () => {
    expect(emuToPx(9144000)).toBe(960); // 4:3 页面宽
    expect(emuToPx(914400)).toBe(96);
    expect(pxToEmu(960)).toBe(9144000);
    expect(pxToEmu(-12)).toBe(-114300); // 允许负偏移（挪出幻灯片）
    expect(pxToEmu(130)).toBe(1238250); // 整 px 的 EMU 值往返一致（130×9525）
  });

  it("pt ↔ 半磅是 ×100 / ÷100，半磅级的小数字号不掉精度", () => {
    expect(halfPointToPt(1800)).toBe(18);
    expect(ptToHalfPoint(18)).toBe(1800);
    expect(ptToHalfPoint(24.5)).toBe(2450); // PowerPoint 允许 0.5pt
    expect(halfPointToPt(ptToHalfPoint(13.3))).toBe(13.3);
  });
});
