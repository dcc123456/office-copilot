import type { IpcMain } from "electron";
import { getActiveProviderSync, loadSettings, type ProviderConfig } from "./settingsStore.js";

interface TransformRequest {
  action: string;
  text: string;
  instruction?: string;
  /** Replace the built-in system prompt (used by the document-editing agent). */
  system?: string;
  /** Use a specific provider id (for switching from the UI); default uses the active one. */
  providerId?: string;
  /** 续聊时带过来的历史轮次；上送前只保留最近 `HISTORY_TURNS` 轮（AGENTS §2.9）。 */
  history?: Array<{ role: "user" | "assistant"; text: string }>;
}

/** 一次请求回填给模型的最近轮数：够接上文，又不至于把整段会话原文上送。 */
const HISTORY_TURNS = 12;

/**
 * 工具回路里的一条消息（OpenAI `/chat/completions` 的 `messages` 形状）。
 *
 * `messages` 由渲染层攒——表格载荷只在渲染层，工具在哪边执行、上下文就在哪边长。
 */
export interface ToolTurnMessage {
  role: "system" | "user" | "assistant" | "tool";
  content?: string;
  /** `role:"tool"` 时回填这条结果属于哪次调用。 */
  tool_call_id?: string;
  /** assistant 这一轮要调哪些工具（续跑时要把原样带回去）。 */
  tool_calls?: Array<{ id: string; type: "function"; function: { name: string; arguments: string } }>;
}

/** 模型要求的一次工具调用；`arguments` 是原文 JSON 串，由渲染层过 zod 校验。 */
export interface ToolCall {
  id: string;
  name: string;
  arguments: string;
}

/** `ai:tool-turn` 的请求：一整段上下文 + 给模型看的工具定义。 */
export interface ToolTurnRequest {
  messages: ToolTurnMessage[];
  tools: unknown[];
  providerId?: string;
}

/** 一轮的回应：要么模型只说话（`content`），要么模型要求调工具（`toolCalls`）。 */
export type ToolTurnResult =
  | { ok: true; content: string; toolCalls: ToolCall[] }
  | { ok: false; message: string; needsSetup?: boolean };

type StreamEvent =
  | { requestId: string; kind: "delta"; text: string }
  /** `stopped` 区分"用户按了暂停"和"模型写完了"：界面据此给出「继续」，半截内容不算失败。 */
  | { requestId: string; kind: "done"; stopped?: boolean }
  | {
      requestId: string;
      kind: "error";
      message: string;
      needsSetup?: boolean;
    };

const SYSTEM_BY_ACTION: Record<string, string> = {
  扩写: "你是一个中文写作助手。把用户给的文字扩写得更详细、更完整，保留原意，只输出改写后的正文，不要解释。",
  缩写: "你是一个中文写作助手。把用户给的文字缩写到一半长度以内，保留关键信息，只输出缩写后的正文，不要解释。",
  润色: "你是一个中文写作助手。润色用户给的文字，让它更通顺、更专业，但不改变意思。只输出润色后的正文。",
  "换语气:正式商务":
    "你是一个商务文书编辑。把用户给的文字改写为正式、严谨的商务语气，保留原意。只输出改写后的正文。",
  "换语气:口语化":
    "你是一个中文写作助手。把用户给的文字改写为自然、口语化但仍然专业的表达，保留原意。只输出改写后的正文。",
  "翻译:英文":
    "You are a professional translator. Translate the user's Chinese text into natural English. Output only the translation, no explanation.",
  "翻译:中文":
    "你是专业译者。把用户给的英文翻译成自然流畅的简体中文。只输出译文，不要解释。",
  "排版:规范标题":
    "你是文档排版助手。把用户给的标题改写为更规范、专业的中文标题，不超过20字，只输出标题。",
  chat: "你是 Office Copilot 里的写作助手。回答简洁、中文，必要时分点。可以直接给出可粘贴到文档的成稿。",
};

/**
 * 组装发给模型的 messages：系统提示词 + 最近若干轮历史 + 本轮用户内容。
 *
 * @param req 渲染层送来的请求；`history` 是续聊带过来的历史轮次（最多取最近 12 轮，AGENTS §2.9）。
 * @returns 按顺序排好的消息数组；无历史时就是「系统 + 本轮」两条。
 */
export function buildMessages(
  req: TransformRequest,
): Array<{ role: string; content: string }> {
  const system =
    req.system ?? SYSTEM_BY_ACTION[req.action] ?? "你是一个中文写作助手。按用户要求改写文字，只输出结果。";
  const userContent =
    req.action === "chat"
      ? req.text
      : req.instruction
        ? `${req.instruction}\n\n原文：\n${req.text}`
        : req.text;
  // 历史轮次夹在系统提示词与本轮之间：模型看得见上文，又不必重读整段会话。
  const history = (req.history ?? []).slice(-HISTORY_TURNS).map((t) => ({
    role: t.role,
    content: t.text,
  }));
  return [
    { role: "system", content: system },
    ...history,
    { role: "user", content: userContent },
  ];
}

function chatUrl(provider: ProviderConfig): string {
  return `${provider.baseUrl.replace(/\/$/, "")}/chat/completions`;
}

/** 一次 LLM 请求的超时：流式与非流式同一口径，超时按"用户按了暂停"以外的失败报。 */
const LLM_TIMEOUT_MS = 120_000;

/**
 * 向服务商发一次 `/chat/completions` 请求（三条调用路径共用这一处，避免第三份 fetch 形状）。
 *
 * @param provider 解析好的服务商配置；`extraBody` 原样并进请求体（各家网关的私有开关）。
 * @param body 请求体里除 `model`/`temperature`/`extraBody` 之外的部分（`messages`、`stream`、`tools` 等）。
 * @param signal 超时/取消信号。
 * @returns 原始 `Response`；HTTP 层错误不在这判，由调用方按流式或非流式各自处理。
 */
async function postChat(
  provider: ProviderConfig,
  body: Record<string, unknown>,
  signal: AbortSignal,
): Promise<Response> {
  return fetch(chatUrl(provider), {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${provider.apiKey ?? ""}`,
    },
    body: JSON.stringify({
      model: provider.model,
      temperature: provider.temperature ?? 0.3,
      ...body,
      ...(provider.extraBody ?? {}),
    }),
    signal,
  });
}

async function resolveProvider(providerId?: string): Promise<ProviderConfig | null> {
  const settings = await loadSettings();
  return providerId
    ? (settings.providers.find((p) => p.id === providerId && p.enabled) ?? null)
    : getActiveProviderSync();
}

const NO_PROVIDER_MESSAGE =
  "还没有配置模型。点右上角⚙️添加一个模型服务商（DeepSeek、火山方舟、GLM、Kimi 等都支持）。";

export function registerAIHandlers(ipc: IpcMain): void {
  ipc.handle("ai:transform", async (_e, req: TransformRequest) => {
    const resolved = await providerFor(req.providerId);
    if (!resolved.ok) return resolved;
    return callLLM(resolved.provider, req);
  });

  // 工具回路按"一轮"走：上下文与工具在哪执行都属于渲染层（表格载荷只在那边），主进程只转发。
  ipc.handle("ai:tool-turn", (_e, req: ToolTurnRequest) => toolTurn(req));

  const inflight = new Map<string, { ctrl: AbortController; canceled: boolean }>();

  ipc.handle(
    "ai:stream-start",
    async (
      e,
      payload: { requestId: string } & TransformRequest,
    ) => {
      const { requestId, ...req } = payload;
      const wc = e.sender;
      const send = (ev: StreamEvent) => {
        if (!wc.isDestroyed()) wc.send("ai:stream", ev);
      };
      const resolved = await providerFor(req.providerId);
      if (!resolved.ok) return resolved;
      const provider = resolved.provider;

      const ctrl = new AbortController();
      const entry = { ctrl, canceled: false };
      inflight.set(requestId, entry);
      const t = setTimeout(() => ctrl.abort(), 180_000);

      // Resolve immediately; progress arrives via the "ai:stream" events.
      void runStream(
        provider,
        req,
        ctrl.signal,
        (text) => send({ requestId, kind: "delta", text }),
      )
        .then(() => send({ requestId, kind: "done" }))
        .catch((err: unknown) => {
          if (entry.canceled) {
            send({ requestId, kind: "done", stopped: true });
            return;
          }
          const aborted = err instanceof Error && err.name === "AbortError";
          const message = aborted
            ? "请求超时（180 秒未完成）。"
            : err instanceof Error
              ? err.message
              : String(err);
          send({ requestId, kind: "error", message });
        })
        .finally(() => {
          clearTimeout(t);
          inflight.delete(requestId);
        });
      return { ok: true };
    },
  );

  ipc.handle("ai:stream-cancel", (_e, requestId: string) => {
    const entry = inflight.get(requestId);
    if (entry) {
      entry.canceled = true;
      entry.ctrl.abort();
    }
    return { ok: true };
  });
}

async function runStream(
  provider: ProviderConfig,
  req: TransformRequest,
  signal: AbortSignal,
  onDelta: (text: string) => void,
): Promise<void> {
  const res = await postChat(provider, { messages: buildMessages(req), stream: true }, signal);
  if (!res.ok) {
    const text = await res.text().catch(() => "");
    throw new Error(`LLM ${res.status}: ${text.slice(0, 500)}`);
  }

  const ctype = res.headers.get("content-type") ?? "";
  // Gateways that ignore `stream:true` answer with a single JSON body.
  if (ctype.includes("application/json")) {
    const data = (await res.json()) as {
      choices?: Array<{ message?: { content?: string } }>;
    };
    const text = data.choices?.[0]?.message?.content?.trim() ?? "";
    if (!text) throw new Error("模型返回为空。");
    onDelta(text);
    return;
  }

  if (!res.body) throw new Error("模型返回为空。");
  const reader = res.body.getReader();
  const decoder = new TextDecoder();
  let buffer = "";
  let received = "";
  let doneByServer = false;
  for (;;) {
    const { value, done } = await reader.read();
    if (done) break;
    buffer += decoder.decode(value, { stream: true });
    buffer = buffer.replace(/\r\n/g, "\n");
    // SSE frames are separated by a blank line.
    let sep: number;
    while ((sep = buffer.indexOf("\n\n")) !== -1) {
      const frame = buffer.slice(0, sep);
      buffer = buffer.slice(sep + 2);
      for (const line of frame.split("\n")) {
        const trimmed = line.trim();
        if (!trimmed.startsWith("data:")) continue;
        const payload = trimmed.slice(5).trim();
        if (payload === "[DONE]") {
          doneByServer = true;
          continue;
        }
        let json: {
          choices?: Array<{ delta?: { content?: string } }>;
        };
        try {
          json = JSON.parse(payload) as typeof json;
        } catch {
          continue; // keepalives / malformed frames
        }
        const chunk = json.choices?.[0]?.delta?.content;
        if (typeof chunk === "string" && chunk) {
          received += chunk;
          onDelta(chunk);
        }
      }
    }
    if (doneByServer) break;
  }
  if (!received.trim()) throw new Error("模型返回为空。");
}

export async function callLLM(
  provider: ProviderConfig,
  req: TransformRequest,
): Promise<{ ok: true; text: string } | { ok: false; message: string; needsSetup?: boolean }> {
  const ctrl = new AbortController();
  const t = setTimeout(() => ctrl.abort(), LLM_TIMEOUT_MS);
  try {
    const res = await postChat(provider, { messages: buildMessages(req) }, ctrl.signal);
    if (!res.ok) {
      const text = await res.text().catch(() => "");
      return { ok: false, message: `LLM ${res.status}: ${text.slice(0, 500)}` };
    }
    const data = (await res.json()) as {
      choices?: Array<{ message?: { content?: string } }>;
    };
    const text = data.choices?.[0]?.message?.content?.trim() ?? "";
    if (!text) return { ok: false, message: "模型返回为空。" };
    return { ok: true, text };
  } catch (err) {
    return { ok: false, message: errorMessage(err) };
  } finally {
    clearTimeout(t);
  }
}

/**
 * 工具回路的一轮请求：整段上下文与工具定义都由渲染层给，主进程只管转发与超时。
 *
 * @param req `messages`（含 `role:"tool"` 的回填）、`tools`、可选 `providerId`。
 * @returns 这一轮的正文与模型点名的工具调用；没配服务商时 `needsSetup`，HTTP/超时都是 `ok:false` + 消息。
 */
export async function toolTurn(req: ToolTurnRequest): Promise<ToolTurnResult> {
  const resolved = await providerFor(req.providerId);
  if (!resolved.ok) return resolved;
  const ctrl = new AbortController();
  const t = setTimeout(() => ctrl.abort(), LLM_TIMEOUT_MS);
  try {
    const res = await postChat(
      resolved.provider,
      { messages: req.messages, tools: req.tools, tool_choice: "auto" },
      ctrl.signal,
    );
    if (!res.ok) {
      const text = await res.text().catch(() => "");
      return { ok: false, message: `LLM ${res.status}: ${text.slice(0, 500)}` };
    }
    return toolTurnOf(await res.json());
  } catch (err) {
    return { ok: false, message: errorMessage(err) };
  } finally {
    clearTimeout(t);
  }
}

/**
 * 拆响应体：这一轮模型是说话还是要调工具。
 *
 * @param data 网关返回的 JSON。
 * @returns 正文（已去首尾空白）与工具调用数组；`function.name` 为空的调用条目丢掉（各家网关偶发只回 id）。
 *          两者都空时判为"模型返回为空"——不许把空回包静默当成"没有工具调用"。
 */
export function toolTurnOf(data: unknown): ToolTurnResult {
  const message = (
    data as {
      choices?: Array<{ message?: { content?: unknown; tool_calls?: unknown } }>;
    }
  )?.choices?.[0]?.message;
  const content = typeof message?.content === "string" ? message.content.trim() : "";
  const rawCalls: Array<{ id?: unknown; function?: { name?: unknown; arguments?: unknown } }> =
    Array.isArray(message?.tool_calls) ? message?.tool_calls : [];
  const toolCalls: ToolCall[] = rawCalls
    .map((call) => ({
      id: typeof call.id === "string" ? call.id : "",
      name: typeof call.function?.name === "string" ? call.function.name : "",
      arguments: typeof call.function?.arguments === "string" ? call.function.arguments : "",
    }))
    .filter((call) => call.name !== "");
  if (content === "" && toolCalls.length === 0) return { ok: false, message: "模型返回为空。" };
  return { ok: true, content, toolCalls };
}

/**
 * 把异常换成给人看的一句话。
 *
 * @param err `fetch`/解析抛出的异常。
 * @returns 超时单独说明（`AbortError` 既来自 120 秒超时，也来自用户按暂停）。
 */
function errorMessage(err: unknown): string {
  if (err instanceof Error && err.name === "AbortError") {
    return `请求超时（${LLM_TIMEOUT_MS / 1000} 秒未完成）。`;
  }
  return err instanceof Error ? err.message : String(err);
}

/** 解析服务商的两种失败要分开：配置读不出来带原因，读出来没配才是"去添加模型服务商"。 */
type ProviderResolution =
  | { ok: true; provider: ProviderConfig }
  | { ok: false; message: string; needsSetup?: boolean };

/**
 * 按 `providerId` 取服务商（省略时取当前激活的那个）。
 *
 * @param providerId 渲染层指定的服务商 id；`undefined` 用设置里的默认。
 * @returns 取到就是 `{ok:true, provider}`；没配是 `needsSetup`；读配置失败是带原因的 `ok:false`。
 */
async function providerFor(providerId?: string): Promise<ProviderResolution> {
  try {
    const provider = await resolveProvider(providerId);
    return provider
      ? { ok: true, provider }
      : { ok: false, needsSetup: true, message: NO_PROVIDER_MESSAGE };
  } catch (err) {
    return { ok: false, message: errorMessage(err) };
  }
}
