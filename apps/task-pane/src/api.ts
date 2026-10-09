import type { CommandEvent, CommandRequest, Proposal } from "@office-copilot/shared";
import { API_BASE } from "./config.js";

/**
 * Stream a command to the server and yield parsed SSE events. We don't use
 * EventSource because we need a POST body with the document outline.
 */
export async function* streamCommand(
  body: CommandRequest,
  signal?: AbortSignal,
): AsyncGenerator<CommandEvent> {
  const res = await fetch(`${API_BASE}/api/word/command`, {
    method: "POST",
    headers: { "Content-Type": "application/json", Accept: "text/event-stream" },
    body: JSON.stringify(body),
    signal,
  });

  if (!res.ok || !res.body) {
    const text = await res.text().catch(() => "");
    throw new Error(`服务端错误 ${res.status}: ${text || res.statusText}`);
  }

  const reader = res.body.getReader();
  const decoder = new TextDecoder();
  let buffer = "";

  while (true) {
    const { value, done } = await reader.read();
    if (done) break;
    buffer += decoder.decode(value, { stream: true });

    let sep: number;
    while ((sep = buffer.indexOf("\n\n")) >= 0) {
      const raw = buffer.slice(0, sep);
      buffer = buffer.slice(sep + 2);
      const dataLine = raw
        .split("\n")
        .filter((l) => l.startsWith("data:"))
        .map((l) => l.slice(5).trimStart())
        .join("");
      if (!dataLine) continue;
      try {
        yield JSON.parse(dataLine) as CommandEvent;
      } catch {
        // Ignore malformed keep-alive events.
      }
    }
  }
}

export interface CommandResult {
  proposal?: Proposal;
  status: string[];
  error?: string;
}

export async function runCommand(
  body: CommandRequest,
  handlers: {
    onStatus?: (msg: string) => void;
    onProposal?: (p: Proposal) => void;
    onError?: (msg: string) => void;
  },
  signal?: AbortSignal,
): Promise<CommandResult> {
  const status: string[] = [];
  let proposal: Proposal | undefined;
  let error: string | undefined;

  for await (const evt of streamCommand(body, signal)) {
    switch (evt.event) {
      case "status":
        status.push(evt.message);
        handlers.onStatus?.(evt.message);
        break;
      case "tool_call":
        status.push(`调用工具 ${evt.name}`);
        handlers.onStatus?.(`调用工具 ${evt.name}`);
        break;
      case "proposal":
        proposal = evt.proposal;
        handlers.onProposal?.(evt.proposal);
        break;
      case "error":
        error = evt.message;
        handlers.onError?.(evt.message);
        break;
      case "done":
        break;
    }
  }

  return { proposal, status, error };
}
