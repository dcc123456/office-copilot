import { describe, it, expect, beforeEach, afterEach, vi } from "vitest";
import { runWordAgent } from "./agent.js";

// Mock the LLM by stubbing global.fetch to return canned chat completions.
function mockCompletion(body: unknown) {
  // @ts-expect-error overriding fetch for the test
  global.fetch = vi.fn(async () => ({
    ok: true,
    json: async () => body,
    text: async () => "",
  }));
}

describe("runWordAgent", () => {
  beforeEach(() => {
    process.env.LLM_API_KEY = "test-key";
  });
  afterEach(() => {
    vi.restoreAllMocks();
  });

  it("returns a proposal when the model calls word_submit_proposal", async () => {
    mockCompletion({
      choices: [
        {
          finish_reason: "tool_calls",
          message: {
            role: "assistant",
            content: "",
            tool_calls: [
              {
                id: "call_1",
                type: "function",
                function: {
                  name: "word_submit_proposal",
                  arguments: JSON.stringify({
                    summary: "把甲方改为客户",
                    edits: [{ type: "replace_text", find: "甲方", replace: "客户" }],
                  }),
                },
              },
            ],
          },
        },
      ],
    });

    const events: string[] = [];
    const proposal = await runWordAgent(
      {
        sessionId: "s1",
        instruction: "把甲方改成客户",
        outline: [
          { index: 0, text: "甲方应在期限内交付。" },
          { index: 1, text: "甲方对质量负责。" },
        ],
      },
      {
        onStatus: (m) => events.push(m),
        onToolCall: (name) => events.push(`tool:${name}`),
      },
    );

    expect(proposal).not.toBeNull();
    expect(proposal!.summary).toContain("甲方");
    expect(proposal!.edits).toHaveLength(1);
    expect(proposal!.edits[0]).toMatchObject({ type: "replace_text", find: "甲方" });
  });

  it("surfaces model text as status when no tool is called", async () => {
    mockCompletion({
      choices: [
        {
          finish_reason: "stop",
          message: { role: "assistant", content: "我可以帮你改文档，告诉我要改什么。" },
        },
      ],
    });
    const statuses: string[] = [];
    const proposal = await runWordAgent(
      { sessionId: "s2", instruction: "你好", outline: [] },
      { onStatus: (m) => statuses.push(m) },
    );
    expect(proposal).toBeNull();
    expect(statuses.join("\n")).toContain("我可以帮你改文档");
  });
});
