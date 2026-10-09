import "dotenv/config";
import { WORD_TOOLS, SubmitProposalArgs, editsFromProposal } from "@office-copilot/word-tools";
import type { DocxParagraph } from "@office-copilot/docx-io";
import type { Proposal } from "@office-copilot/shared";
import { newId } from "@office-copilot/shared";

interface ChatMessage {
  role: "system" | "user" | "assistant" | "tool";
  content: string | null;
  tool_calls?: Array<{
    id: string;
    type: "function";
    function: { name: string; arguments: string };
  }>;
  tool_call_id?: string;
  name?: string;
}

const SYSTEM_PROMPT = `你是一个命令行里的 Word 文档修改助手。用户会给你一个 .docx 文件的段落大纲和一条自然语言指令。你可以：

1. 调用 word_list_paragraphs —— 没有参数。用户请求里通常已经附带完整大纲，只有在你需要更多信息时才调用。
2. 最终调用 word_submit_proposal，把打算做的修改打包成一份提案，命令行会展示 diff 并让用户确认。

规则：
- 只支持 replace_text / insert_paragraph / insert_table 三类编辑。
- 插入段落：默认 position="end"；要在某段前后插入时，position="before"/"after" 并提供 anchor（段落下标，从 0 开始）。
- 替换文字：find 必须是文档里确实存在的原文，不要凭空造句子。
- 改写语气/润色：使用 replace_text，find 填原句、replace 填改写句。
- 插入表格：rows 必须是二维字符串数组，列数一致。
- summary 用一句中文概括要做什么。
- 如果信息不足无法安全修改，调用 word_submit_proposal 时给一个空 edits 数组并在 summary 里说明原因。`;

export interface AgentOptions {
  model?: string;
  baseUrl?: string;
  apiKey?: string;
  maxRounds?: number;
}

export async function runAgent(input: {
  instruction: string;
  paragraphs: DocxParagraph[];
  onStatus?: (msg: string) => void;
  onToolCall?: (name: string, args: unknown) => void;
  signal?: AbortSignal;
}, opts: AgentOptions = {}): Promise<Proposal | null> {
  const baseUrl =
    opts.baseUrl ?? process.env.LLM_BASE_URL ?? "https://api.deepseek.com/v1";
  const apiKey = opts.apiKey ?? process.env.LLM_API_KEY ?? "";
  const model = opts.model ?? process.env.LLM_MODEL ?? "deepseek-chat";
  const maxRounds = opts.maxRounds ?? 6;

  if (!apiKey) {
    throw new Error(
      "未配置 LLM_API_KEY。请运行 `cp .env.example .env` 并填入 Key，或设置环境变量。",
    );
  }

  const messages: ChatMessage[] = [
    { role: "system", content: SYSTEM_PROMPT },
    {
      role: "user",
      content: buildUserPrompt(input.instruction, input.paragraphs),
    },
  ];

  let proposal: Proposal | null = null;

  for (let round = 0; round < maxRounds; round++) {
    input.onStatus?.(round === 0 ? "请求模型…" : `继续推理（第 ${round + 1} 轮）…`);
    const resp = await chatCompletion(messages, { baseUrl, apiKey, model }, input.signal);
    const choice = resp.choices[0];
    if (!choice) throw new Error("模型没有返回结果。");
    const msg = choice.message;
    messages.push({
      role: "assistant",
      content: msg.content ?? "",
      tool_calls: msg.tool_calls,
    });

    if (!msg.tool_calls || msg.tool_calls.length === 0) {
      if (msg.content) input.onStatus?.(msg.content);
      break;
    }

    for (const call of msg.tool_calls) {
      const result = dispatch(call, input.paragraphs, (p) => {
        proposal = p;
      });
      input.onToolCall?.(call.function.name, JSON.parse(call.function.arguments || "{}"));
      messages.push({
        role: "tool",
        tool_call_id: call.id,
        name: call.function.name,
        content: result,
      });
    }
    if (proposal) break;
  }
  return proposal;
}

function dispatch(
  call: { id: string; function: { name: string; arguments: string } },
  paragraphs: DocxParagraph[],
  setProposal: (p: Proposal) => void,
): string {
  let args: unknown = {};
  try {
    args = call.function.arguments ? JSON.parse(call.function.arguments) : {};
  } catch (e) {
    return `参数解析失败：${(e as Error).message}`;
  }

  switch (call.function.name) {
    case "word_list_paragraphs":
      return JSON.stringify({
        count: paragraphs.length,
        paragraphs: paragraphs.map((p) => ({
          index: p.index,
          style: p.style,
          text: p.text,
        })),
      });
    case "word_submit_proposal": {
      let parsed;
      try {
        parsed = SubmitProposalArgs.parse(args);
      } catch (e) {
        return `参数校验失败：${(e as Error).message}`;
      }
      const edits = editsFromProposal(parsed);
      const proposal: Proposal = {
        id: newId("proposal"),
        summary: parsed.summary,
        edits,
        createdAt: new Date().toISOString(),
      };
      setProposal(proposal);
      return JSON.stringify({ ok: true, proposalId: proposal.id, editCount: edits.length });
    }
    default:
      return `未知工具：${call.function.name}`;
  }
}

function buildUserPrompt(instruction: string, paragraphs: DocxParagraph[]): string {
  const lines: string[] = [`用户指令：${instruction}`, ""];
  if (paragraphs.length > 0) {
    lines.push(`当前文档共 ${paragraphs.length} 段：`);
    for (const p of paragraphs) {
      const text = p.text.length > 200 ? p.text.slice(0, 200) + "…" : p.text;
      lines.push(`[${p.index}]${p.style ? `(${p.style})` : ""} ${text}`);
    }
  } else {
    lines.push("（空文档）");
  }
  return lines.join("\n");
}

async function chatCompletion(
  messages: ChatMessage[],
  cfg: { baseUrl: string; apiKey: string; model: string },
  signal?: AbortSignal,
): Promise<{
  choices: Array<{ message: ChatMessage; finish_reason: string }>;
}> {
  const res = await fetch(`${cfg.baseUrl}/chat/completions`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${cfg.apiKey}`,
    },
    body: JSON.stringify({
      model: cfg.model,
      messages,
      tools: WORD_TOOLS,
      temperature: 0.2,
    }),
    signal,
  });
  if (!res.ok) {
    const body = await res.text().catch(() => "");
    throw new Error(`LLM 请求失败 ${res.status}: ${body.slice(0, 500)}`);
  }
  return res.json();
}
