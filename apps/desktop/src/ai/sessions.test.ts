import { beforeEach, describe, expect, it } from "vitest";
import {
  listSessions,
  newSessionId,
  openSession,
  removeSession,
  saveSession,
  titleOf,
} from "./sessions.js";
import type { SessionRecord, SessionSummary, SessionTurn } from "../global.js";

/**
 * 会话门面：只验"界面交出来的形状"——标题怎么起、送进 preload 的载荷长什么样。
 * 落盘与剪枝是 electron/sessionStore.ts 的事，那边按真实临时目录测。
 */

interface Call {
  channel: string;
  arg: unknown;
}

/** 假装是 preload 暴露的那四个方法，记下调用并回给固定结果。 */
function stubBridge(result: { sessions?: SessionSummary[]; record?: SessionRecord | null }): Call[] {
  const calls: Call[] = [];
  (globalThis as unknown as { window: unknown }).window = {
    officeCopilot: {
      listSessions: async () => {
        calls.push({ channel: "list", arg: null });
        return result.sessions ?? [];
      },
      readSession: async (id: string) => {
        calls.push({ channel: "read", arg: id });
        return result.record ?? null;
      },
      saveSession: async (payload: unknown) => {
        calls.push({ channel: "save", arg: payload });
        return { ok: true };
      },
      removeSession: async (id: string) => {
        calls.push({ channel: "remove", arg: id });
        return { ok: true };
      },
    },
  };
  return calls;
}

describe("titleOf", () => {
  const turns: SessionTurn[] = [
    { role: "assistant", text: "这是引导语，不算标题" },
    { role: "user", text: "帮我把 第三节 改成\n表格" },
    { role: "assistant", text: "可以，先给它加个表头。" },
  ];

  it("标题取第一句用户的话，换行压成空格", () => {
    expect(titleOf(turns)).toBe("帮我把 第三节 改成 表格");
  });

  it("一句话没有时给占位标题，不让列表里出现空白行", () => {
    expect(titleOf([{ role: "assistant", text: "只有 AI 说话" }])).toBe("未命名会话");
  });

  it("超长的那句截到 24 字加省略号", () => {
    const long = "一".repeat(40);
    expect(titleOf([{ role: "user", text: long }])).toBe(`${"一".repeat(24)}…`);
  });
});

describe("会话门面与 preload 的接线", () => {
  beforeEach(() => {
    stubBridge({});
  });

  it("保存时把标题现算好一起送出，界面不用另写一份", async () => {
    const turns: SessionTurn[] = [
      { role: "user", text: "帮我改简历" },
      { role: "assistant", text: "先把项目经历写短。" },
    ];
    const calls = stubBridge({});
    await saveSession({ id: "s-1", turns });
    expect(calls[0]).toEqual({
      channel: "save",
      arg: { id: "s-1", title: "帮我改简历", turns },
    });
  });

  it("打开与删除都按编号走一个通道，返回值原样交给界面", async () => {
    const record: SessionRecord = {
      id: "s-2",
      title: "改简历",
      createdAt: 1,
      updatedAt: 2,
      turns: [{ role: "user", text: "帮我改简历" }],
    };
    const calls = stubBridge({ record, sessions: [] });
    expect(await openSession("s-2")).toEqual(record);
    expect(await removeSession("s-2")).toBe(true);
    expect(await listSessions()).toEqual([]);
    expect(calls.map((c) => c.channel)).toEqual(["read", "remove", "list"]);
  });

  it("新会话编号能直接当文件名：主进程只认这个字符集", () => {
    expect(newSessionId()).toMatch(/^[A-Za-z0-9._-]{1,64}$/);
  });
});
