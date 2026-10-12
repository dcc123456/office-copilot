import { describe, expect, it } from "vitest";
import { getSchema } from "@tiptap/core";
import { EditorState } from "@tiptap/pm/state";
import type { Editor } from "@tiptap/react";
import { EDITOR_EXTENSIONS } from "../editor/extensions/editorExtensions.js";
import { requestProposal } from "./agent.js";

/**
 * 改文档的 agent 走的是主进程那条 SSE 通道：这里只验"接线"——系统提示词必须送出去、
 * 增量必须逐字回抛给界面、收完的 JSON 必须变成提案。模型本身不进来（AGENTS §7.1）。
 */

const schema = getSchema(EDITOR_EXTENSIONS);

/** 起一个只有段落的最小编辑器，大纲由它现算。 */
function editorWith(paragraphs: string[]): Editor {
  const doc = schema.nodeFromJSON({
    type: "doc",
    content: paragraphs.map((t) => ({
      type: "paragraph",
      ...(t ? { content: [{ type: "text", text: t }] } : {}),
    })),
  });
  return { state: EditorState.create({ schema, doc }) } as unknown as Editor;
}

interface StreamEvent {
  requestId: string;
  kind: "delta" | "done" | "error";
  text?: string;
  message?: string;
  needsSetup?: boolean;
}

/**
 * 假装是 preload 暴露的流式接口：记下送出去的 payload，再把分片按顺序喂回监听器。
 *
 * @param events 要回放的事件序列（delta 的文字、done、或 error）。
 * @returns `sent` 里是最后一次送进 `ai:stream-start` 的 payload。
 */
function stubStream(events: StreamEvent[]): { sent: Record<string, unknown> } {
  const sent: Record<string, unknown> = {};
  const listeners: ((ev: StreamEvent) => void)[] = [];
  (globalThis as unknown as { window: unknown }).window = {
    officeCopilot: {
      onAIStream: (cb: (ev: StreamEvent) => void) => {
        listeners.push(cb);
        return () => {
          const i = listeners.indexOf(cb);
          if (i >= 0) listeners.splice(i, 1);
        };
      },
      aiStreamStart: async (payload: { requestId: string }) => {
        Object.assign(sent, payload);
        for (const ev of events) {
          for (const cb of listeners.slice()) cb({ ...ev, requestId: payload.requestId });
        }
        return { ok: true };
      },
    },
  };
  return { sent };
}

describe("改文档的 agent 走流式通道", () => {
  it("系统提示词随请求送出，增量逐字回抛，收完解析成提案", async () => {
    const { sent } = stubStream([
      { requestId: "", kind: "delta", text: '{"summary":"新增两段","ops":[{"op":"insertParagraph","after":0,"text":"' },
      { requestId: "", kind: "delta", text: '甲段\\n乙段"}]}' },
      { requestId: "", kind: "done" },
    ]);
    const chunks: string[] = [];
    const out = await requestProposal({
      editor: editorWith(["现有的一段"]),
      instruction: "再写两段",
      requestId: "req-1",
      onDelta: (c) => chunks.push(c),
    });

    // 漏传 system 就等于把 agent 的指令词汇表整段丢掉，模型只会回散文。
    expect(String(sent.system)).toContain("文档编辑助手");
    expect(sent.action).toBe("agent");
    expect(String(sent.text)).toContain("现有的一段");
    expect(chunks).toHaveLength(2);
    expect(out.ok).toBe(true);
    expect(out.proposal?.summary).toBe("新增两段");
    expect(out.proposal?.ops).toEqual([
      { op: "insertParagraph", after: 0, text: "甲段\n乙段" },
    ]);
  });

  it("模型回的不是 JSON 就照实说，不写任何改动", async () => {
    stubStream([
      { requestId: "", kind: "delta", text: "我建议你把第三段改短一点。" },
      { requestId: "", kind: "done" },
    ]);
    const out = await requestProposal({
      editor: editorWith(["现有的一段"]),
      instruction: "帮我润色",
    });
    expect(out.ok).toBe(false);
    expect(out.message).toContain("可解析");
  });

  it("还没配模型时把 needsSetup 透出来，界面好拉开设置", async () => {
    stubStream([
      { requestId: "", kind: "error", message: "还没有配置模型。", needsSetup: true },
    ]);
    const out = await requestProposal({
      editor: editorWith(["现有的一段"]),
      instruction: "写个开头",
    });
    expect(out.ok).toBe(false);
    expect(out.needsSetup).toBe(true);
    expect(out.message).toContain("配置模型");
  });
});
