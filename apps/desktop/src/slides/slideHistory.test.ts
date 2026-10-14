/**
 * 幻灯片历史记录那一层的用例（`AGENTS.md` §2.5 / §7.1）。
 *
 * 钉两件事：① 三类记录混在同一个 modlog 文件里时，幻灯片那条只认自己的形状，不会把文档或表格的记录
 * 误当成能撤销的批次；② 撤销凭据（整页前/后像）进记录时一条不减、顺序不变。
 */
import { describe, expect, it } from "vitest";
import type { SlideEditOp } from "@office-copilot/pptx-io";
import { isSlideEditRecord, pickLatestSlideEdit, slideEditRecordOf } from "./slideHistory.js";

/** 一条最小的幻灯片 op：正文内容无关紧要，形状对就行。 */
function slideOp(partPath = "ppt/slides/slide1.xml"): SlideEditOp {
  return { partPath, beforeXml: "<p:sld>改前</p:sld>", afterXml: "<p:sld>改后</p:sld>" };
}

describe("幻灯片记录的形状判定", () => {
  it("带整页前后像的 ops 才算幻灯片记录", () => {
    expect(isSlideEditRecord({ id: "1", ts: 1, label: "幻灯片改动", summary: "s", ops: [slideOp()] })).toBe(true);
  });

  it("文档那条 CoreOp 记录与表格那条带 partPath/inverses 的记录都不算", () => {
    // 文档记录：op 是 { kind, addr, ... }，没有 beforeXml。
    expect(isSlideEditRecord({ id: "1", ts: 1, label: "AI 修改", summary: "s", ops: [{ kind: "setText", addr: [1] }] })).toBe(false);
    // 表格记录：顶层带 partPath 与 inverses，但 op 本身没有整页前后像。
    expect(
      isSlideEditRecord({
        id: "1",
        ts: 1,
        label: "表格改动",
        summary: "s",
        partPath: "xl/worksheets/sheet1.xml",
        ops: [{ kind: "setCellValue", partPath: "xl/worksheets/sheet1.xml", target: { row: 0, col: 0 }, before: null, after: null }],
        inverses: [],
      }),
    ).toBe(false);
  });

  it("空 ops、缺字段与不是对象都判 false 而不是抛", () => {
    expect(isSlideEditRecord({ id: "1", ts: 1, label: "幻灯片改动", summary: "s", ops: [] })).toBe(false);
    expect(isSlideEditRecord({ id: "1", ts: 1, ops: [{ partPath: "ppt/slides/slide1.xml" }] })).toBe(false);
    expect(isSlideEditRecord(null)).toBe(false);
    expect(isSlideEditRecord("幻灯片改动")).toBe(false);
  });
});

describe("挑出最近一次可撤销的幻灯片改动", () => {
  it("混着文档与表格记录也只挑幻灯片那条，且按 ts 取最新", () => {
    // ts 给死：两条记录在同一毫秒里造出来时，取谁不该看运气。
    const older = { ...slideEditRecordOf([slideOp()], "早一批"), ts: 1000 };
    const newest = { ...slideEditRecordOf([slideOp("ppt/slides/slide2.xml")], "晚一批"), ts: 2000 };
    const records = [
      { id: "doc", ts: 3000, label: "AI 修改", summary: "文档的", ops: [{ kind: "setText", addr: [0] }] },
      older,
      newest,
    ];
    expect(pickLatestSlideEdit(records)?.summary).toBe("晚一批");
  });

  it("已撤销的不给第二次撤销；一条都没有时回 null", () => {
    const record = slideEditRecordOf([slideOp()], "只一批");
    expect(pickLatestSlideEdit([{ ...record, undone: true }])).toBeNull();
    expect(pickLatestSlideEdit([])).toBeNull();
  });
});

describe("一次落地收成一条记录", () => {
  it("label 固定、summary 由装配层给，ops 原样留住且顺序不变", () => {
    const ops = [slideOp("ppt/slides/slide1.xml"), slideOp("ppt/slides/slide3.xml")];
    const record = slideEditRecordOf(ops, "动了 2 张幻灯片");
    expect(record.label).toBe("幻灯片改动");
    expect(record.summary).toBe("动了 2 张幻灯片");
    expect(record.ops).toEqual(ops);
    expect(record.id).toBeTruthy();
    expect(record.undone).toBeUndefined();
  });
});
