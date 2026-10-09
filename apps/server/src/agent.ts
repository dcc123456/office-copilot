import type { ChatMessage, ToolCall } from "./llm.js";
import { chatCompletion } from "./llm.js";
import {
  WORD_TOOLS,
  SubmitProposalArgs,
  editsFromProposal,
  type SubmitProposalArgsT,
} from "@office-copilot/word-tools";
import type { CommandRequest, ParagraphInfo, Proposal } from "@office-copilot/shared";
import { newId } from "@office-copilot/shared";
import { config } from "./config.js";

const SYSTEM_PROMPT = `你是一个嵌入在 Microsoft Word 里的中文写作助手。用户会给你一条自然语言指令，你可以：

1. 调用 word_list_paragraphs 查看当前文档大纲（你可能已经在请求里拿到了大纲；只有当你需要更完整信息时才调用）。
2. 最终调用 word_submit_proposal 把你打算做的修改打包成一份“提案”。

规则：
- 提案里只能包含 replace_text / insert_paragraph / insert_table 三类编辑，按执行顺序排列。
- 不要声称你已经修改了文档；你只是在提交提案，用户会先看到预览再决定是否接受。
- 替换文字尽量精确，避免误伤（例如“甲方”可能在多处出现，这是符合预期的全局替换）。
- 改写语气类任务：使用 replace_text，find 填原句、replace 填改写后的句子。
- 插入表格：rows 必须是二维字符串数组，列数一致；首行默认加粗。
- 回答要简洁、中文，summary 用一句话概括。
- 如果信息不足（例如用户要求“改第二段”但没有大纲），先调用 word_list_paragraphs。`;

export interface AgentHandlers {
  onStatus?: (message: string) => void;
  onToolCall?: (name: string, args: unknown) => void;
  onProposal?: (proposal: Proposal) => void;
}

export async function runWordAgent(
  request: CommandRequest,
  handlers: AgentHandlers = {},
  signal?: AbortSignal,
): Promise<Proposal | null> {
  const outline = request.outline ?? [];
  const messages: ChatMessage[] = [
    { role: "system", content: SYSTEM_PROMPT },
    {
      role: "user",
      content: buildUserPrompt(request.instruction, outline, request.selection),
    },
  ];

  let proposal: Proposal | null = null;

  for (let round = 0; round < config.maxToolRounds; round++) {
    handlers.onStatus?.(round === 0 ? "正在请求模型…" : `继续推理（第 ${round + 1} 轮）…`);
    const resp = await chatCompletion({
      messages,
      tools: WORD_TOOLS,
      temperature: 0.2,
      signal,
    });

    const choice = resp.choices[0];
    if (!choice) throw new Error("模型没有返回结果。");
    const msg = choice.message;
    messages.push({
      role: "assistant",
      content: msg.content ?? "",
      tool_calls: msg.tool_calls,
    });

    if (!msg.tool_calls || msg.tool_calls.length === 0) {
      // No tool call — if we already have a proposal we're done, otherwise
      // the model is just chatting. We surface its text as a status message.
      if (msg.content) handlers.onStatus?.(msg.content);
      break;
    }

    for (const call of msg.tool_calls) {
      const result = await dispatchTool(call, outline, handlers, (p) => {
        proposal = p;
      });
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

async function dispatchTool(
  call: ToolCall,
  outline: ParagraphInfo[],
  handlers: AgentHandlers,
  setProposal: (p: Proposal) => void,
): Promise<string> {
  let args: unknown = {};
  try {
    args = call.function.arguments ? JSON.parse(call.function.arguments) : {};
  } catch (e) {
    return `参数解析失败：${(e as Error).message}`;
  }
  handlers.onToolCall?.(call.function.name, args);

  switch (call.function.name) {
    case "word_list_paragraphs":
      return JSON.stringify({
        count: outline.length,
        paragraphs: outline.map((p) => ({
          index: p.index,
          style: p.style,
          text: p.text.length > 200 ? p.text.slice(0, 200) + "…" : p.text,
        })),
      });

    case "word_submit_proposal": {
      let parsed: SubmitProposalArgsT;
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
      handlers.onProposal?.(proposal);
      return JSON.stringify({ ok: true, proposalId: proposal.id, editCount: edits.length });
    }

    default:
      return `未知工具：${call.function.name}`;
  }
}

function buildUserPrompt(
  instruction: string,
  outline: ParagraphInfo[],
  selection?: string,
): string {
  const lines: string[] = [];
  lines.push(`用户指令：${instruction}`);
  if (selection && selection.trim()) {
    lines.push("");
    lines.push("当前选区文本：");
    lines.push(
      selection.length > 1000 ? selection.slice(0, 1000) + "…（已截断）" : selection,
    );
  }
  if (outline.length > 0) {
    lines.push("");
    lines.push(`文档大纲（共 ${outline.length} 段，最多展示 200 段）：`);
    for (const p of outline.slice(0, 200)) {
      const text = p.text.length > 120 ? p.text.slice(0, 120) + "…" : p.text;
      lines.push(`[${p.index}]${p.style ? `(${p.style})` : ""} ${text}`);
    }
  } else {
    lines.push("");
    lines.push("（暂未获取到文档大纲。）");
  }
  return lines.join("\n");
}
