import { afterEach, describe, expect, it, vi } from "vitest";
import type { IpcMain } from "electron";

/**
 * 主进程这一侧的三条线：续聊历史的组装（`buildMessages`）、工具回路的一轮（`toolTurn` /
 * `toolTurnOf`）、以及暂停是否真的中止了请求。全部打桩 `fetch`，不碰真实端点（AGENTS §7.1）；
 * 真实端点由一次性探针负责（不进主干，§6.5）。
 */

vi.mock("electron", () => ({
  app: { getPath: () => "/tmp/oc-ai-handlers-test" },
}));

/** 服务商配置桩：`loadSettings` 与"当前激活的那个"都指向同一条，测试不去读用户机器上的 config.json。 */
vi.mock("./settingsStore.js", () => {
  const stubProvider = {
    id: "stub",
    name: "桩服务商",
    baseUrl: "https://llm.stub/v1/",
    apiKey: "sk-stub",
    model: "stub-model",
    temperature: 0.1,
    extraBody: { thinking: { type: "disabled" } },
    enabled: true,
  };
  return {
    loadSettings: async () => ({
      version: 1,
      providers: [stubProvider],
      defaultProviderId: "stub",
    }),
    getActiveProviderSync: () => stubProvider,
  };
});

const { buildMessages, registerAIHandlers, toolTurn, toolTurnOf } = await import("./aiHandlers.js");

afterEach(() => {
  vi.unstubAllGlobals();
});

/** 一轮请求的桩回包：`json` 给非流式路径，`text` 给 HTTP 错误路径。 */
function stubResponse(json: unknown): Response {
  return {
    ok: true,
    status: 200,
    headers: new Headers({ "content-type": "application/json" }),
    json: async () => json,
    text: async () => JSON.stringify(json),
  } as unknown as Response;
}

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

describe("toolTurnOf：一轮回包是说话还是要调工具", () => {
  it("只回正文时工具调用是空数组，正文去掉首尾空白", () => {
    expect(toolTurnOf({ choices: [{ message: { content: "  先按销售额排序。  " } }] })).toEqual({
      ok: true,
      content: "先按销售额排序。",
      toolCalls: [],
    });
  });

  it("arguments 原文带回，name 为空的条目丢掉（网关偶发只回 id）", () => {
    const result = toolTurnOf({
      choices: [
        {
          message: {
            content: null,
            tool_calls: [
              { id: "call_1", type: "function", function: { name: "sheet_read_range", arguments: '{"range":"A1:B3"}' } },
              { id: "call_2", type: "function", function: { arguments: "{}" } },
            ],
          },
        },
      ],
    });
    expect(result).toEqual({
      ok: true,
      content: "",
      toolCalls: [{ id: "call_1", name: "sheet_read_range", arguments: '{"range":"A1:B3"}' }],
    });
  });

  it("正文与工具调用都没有时判为返回为空，不许静默当成「没有工具调用」", () => {
    expect(toolTurnOf({ choices: [{ message: { content: "   " } }] })).toEqual({
      ok: false,
      message: "模型返回为空。",
    });
  });
});

describe("toolTurn：一轮请求的组装与失败语义", () => {
  const turnRequest = {
    providerId: "stub",
    messages: [
      { role: "system" as const, content: "你是表格代理" },
      { role: "user" as const, content: "把销售额低于均值的行标出来" },
    ],
    tools: [{ type: "function", function: { name: "sheet_describe" } }],
  };

  it("请求体带上 messages / tools / tool_choice，并把服务商的 extraBody 并进同一次请求", async () => {
    const fetchMock = vi.fn(async () =>
      stubResponse({ choices: [{ message: { content: "读完了。", tool_calls: [] } }] }),
    );
    vi.stubGlobal("fetch", fetchMock);

    const result = await toolTurn(turnRequest);

    expect(result).toEqual({ ok: true, content: "读完了。", toolCalls: [] });
    const [url, init] = fetchMock.mock.calls[0] as [string, RequestInit];
    // baseUrl 结尾的斜杠被吃掉，只补一次 `/chat/completions`。
    expect(url).toBe("https://llm.stub/v1/chat/completions");
    expect((init.headers as Record<string, string>).Authorization).toBe("Bearer sk-stub");
    expect(JSON.parse(String(init.body))).toEqual({
      model: "stub-model",
      temperature: 0.1,
      messages: turnRequest.messages,
      tools: turnRequest.tools,
      tool_choice: "auto",
      thinking: { type: "disabled" },
    });
  });

  it("HTTP 非 2xx 时带回状态码与截断的响应正文", async () => {
    vi.stubGlobal("fetch", async () =>
      ({
        ok: false,
        status: 503,
        headers: new Headers(),
        json: async () => ({}),
        text: async () => "upstream busy",
      }) as unknown as Response,
    );

    expect(await toolTurn(turnRequest)).toEqual({ ok: false, message: "LLM 503: upstream busy" });
  });

  it("没配服务商就不发请求：needsSetup 让渲染层去设置页", async () => {
    const fetchMock = vi.fn(async () => stubResponse({}));
    vi.stubGlobal("fetch", fetchMock);

    expect(await toolTurn({ ...turnRequest, providerId: "not-configured" })).toMatchObject({
      ok: false,
      needsSetup: true,
    });
    expect(fetchMock).not.toHaveBeenCalled();
  });
});

describe("ai:stream-start：暂停要真的中止请求", () => {
  /**
   * 等异步链路把事件发出来——读流被中止要绕过一个微任务与一个宏任务。
   *
   * @param predicate 观察到期望事件就返回 true。
   * @returns 超时（50 个宏任务）仍未看到时返回 false，由用例判失败。
   */
  async function waitUntil(predicate: () => boolean): Promise<boolean> {
    for (let attempt = 0; attempt < 50; attempt += 1) {
      if (predicate()) return true;
      await new Promise((resolve) => setTimeout(resolve, 0));
    }
    return predicate();
  }

  it("cancel 之后挂给 fetch 的 signal 被中止，回抛 stopped 而不是错误", async () => {
    const handlers = new Map<string, (...args: unknown[]) => Promise<unknown>>();
    registerAIHandlers({
      handle: (channel: string, fn: unknown) => {
        handlers.set(channel, fn as (...args: unknown[]) => Promise<unknown>);
      },
    } as unknown as IpcMain);

    let isAborted = false;
    vi.stubGlobal("fetch", (_url: string, init: RequestInit) => {
      // 没接上 signal 的旧实现会在这里当场抛错——正是这条用例要拦的。
      (init.signal as AbortSignal).addEventListener("abort", () => {
        isAborted = true;
      });
      return Promise.resolve({
        ok: true,
        status: 200,
        headers: new Headers({ "content-type": "text/event-stream" }),
        body: new ReadableStream<Uint8Array>({
          start(controller) {
            (init.signal as AbortSignal).addEventListener("abort", () => {
              controller.error(new DOMException("aborted", "AbortError"));
            });
          },
        }),
      } as unknown as Response);
    });

    const sent: unknown[] = [];
    const sender = { isDestroyed: () => false, send: (_ch: string, ev: unknown) => sent.push(ev) };
    const started = await handlers.get("ai:stream-start")!(
      { sender },
      { requestId: "r1", action: "chat", text: "写长一点" },
    );
    expect(started).toEqual({ ok: true });
    expect(isAborted).toBe(false);

    await handlers.get("ai:stream-cancel")!({ sender }, "r1");

    expect(isAborted).toBe(true);
    const kinds = () => sent.map((ev) => ev as { kind?: string; stopped?: boolean });
    const sawStopped = await waitUntil(() =>
      kinds().some((ev) => ev.kind === "done" && ev.stopped === true),
    );
    expect(sawStopped).toBe(true);
    expect(kinds().some((ev) => ev.kind === "error")).toBe(false);
  });
});
