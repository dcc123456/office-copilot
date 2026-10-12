import { describe, expect, it, vi } from "vitest";

/**
 * 续聊时历史轮次的组装用例：只验 `buildMessages` 的拼装顺序与截断，
 * 打真实端点由一次性探针负责（不进主干，AGENTS §6.5）。
 */

vi.mock("electron", () => ({
  app: { getPath: () => "/tmp/oc-ai-handlers-test" },
}));

const { buildMessages } = await import("./aiHandlers.js");

describe("buildMessages：把会话历史夹在系统提示词与本轮之间", () => {
  it("没有历史时只有系统与本轮两条", () => {
    const messages = buildMessages({ action: "chat", text: "帮我想个标题" });
    expect(messages.map((m) => m.role)).toEqual(["system", "user"]);
    expect(messages[1]?.content).toBe("帮我想个标题");
  });

  it("chat 以外的动作把指令拼在原文前面，历史仍然排在中间", () => {
    const messages = buildMessages({
      action: "润色",
      text: "第二段原文",
      instruction: "只改写剩下的 2 段",
      history: [
        { role: "user", text: "第一段问" },
        { role: "assistant", text: "第一段答" },
      ],
    });
    expect(messages.map((m) => m.role)).toEqual(["system", "user", "assistant", "user"]);
    expect(messages[3]?.content).toBe("只改写剩下的 2 段\n\n原文：\n第二段原文");
  });

  it("自定义 system 覆盖内置提示词，角色映射成 OpenAI 形态", () => {
    const messages = buildMessages({
      action: "chat",
      text: "本轮",
      system: "你是文档改写代理",
      history: [{ role: "assistant", text: "上一轮答复" }],
    });
    expect(messages[0]).toEqual({ role: "system", content: "你是文档改写代理" });
    expect(messages[1]).toEqual({ role: "assistant", content: "上一轮答复" });
  });

  it("会话太长只带上最近 12 轮，丢的是最早的轮次", () => {
    const many = Array.from({ length: 30 }, (_, index) => ({
      role: (index % 2 === 0 ? "user" : "assistant") as "user" | "assistant",
      text: `第 ${index + 1} 轮`,
    }));
    const messages = buildMessages({ action: "chat", text: "本轮", history: many });
    const carried = messages.slice(1, -1).map((m) => m.content);
    expect(carried).toHaveLength(12);
    expect(carried[0]).toBe("第 19 轮");
    expect(carried.at(-1)).toBe("第 30 轮");
  });
});
