import { config } from "./config.js";

/**
 * Minimal OpenAI-compatible chat-completions client. We deliberately use
 * native fetch so there is no SDK lock-in; DeepSeek, Azure OpenAI, and
 * OpenAI itself all speak the same wire format.
 */

export interface ChatMessage {
  role: "system" | "user" | "assistant" | "tool";
  content: string | null;
  tool_calls?: ToolCall[];
  tool_call_id?: string;
  name?: string;
}

export interface ToolCall {
  id: string;
  type: "function";
  function: { name: string; arguments: string };
}

export interface ToolDef {
  type: "function";
  function: {
    name: string;
    description?: string;
    parameters: Record<string, unknown>;
  };
}

export interface ChatCompletionOpts {
  messages: ChatMessage[];
  tools?: ToolDef[];
  temperature?: number;
  signal?: AbortSignal;
}

export interface ChatCompletion {
  choices: Array<{
    message: ChatMessage;
    finish_reason: string;
  }>;
}

export async function chatCompletion(opts: ChatCompletionOpts): Promise<ChatCompletion> {
  const apiKey = process.env.LLM_API_KEY ?? config.llm.apiKey;
  if (!apiKey) {
    throw new Error("LLM_API_KEY 未配置：请复制 .env.example 为 .env 并填入 API Key。");
  }
  const res = await fetch(`${config.llm.baseUrl}/chat/completions`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${apiKey}`,
    },
    body: JSON.stringify({
      model: config.llm.model,
      messages: opts.messages,
      tools: opts.tools,
      temperature: opts.temperature ?? 0.2,
    }),
    signal: opts.signal,
  });

  if (!res.ok) {
    const body = await res.text().catch(() => "");
    throw new Error(`LLM 请求失败 ${res.status}: ${body.slice(0, 500)}`);
  }
  return (await res.json()) as ChatCompletion;
}
