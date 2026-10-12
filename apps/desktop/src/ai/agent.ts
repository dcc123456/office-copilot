import type { Editor } from "@tiptap/react";
import {
  addressOf,
  buildOutline,
  coerceProposal,
  normalise,
  paragraphTexts,
  planEdit,
  revert,
  runOps,
  type AiOp,
  type Block,
  type CoreOp,
  type Proposal,
} from "./docEdit.js";
import { describeCoreOp } from "./docEdit.js";
import { diagramResolverFor } from "./diagramCanvas.js";
import { blocksOf, executeOps, locateAll } from "../editor/docBlocks.js";
import type { SessionTurn } from "../global.js";

/**
 * The document-editing agent: ask the model for a JSON edit plan, validate it
 * against the real outline, apply it as one transaction, and record enough to
 * diff and individually undo it later.
 */

export interface ModificationRecord {
  id: string;
  ts: number;
  label: string;
  summary: string;
  ops: CoreOp[];
  inverses: CoreOp[];
  before: Block[];
  after: Block[];
  undone: boolean;
  undoneAt?: number;
  /** 美化排版记录打标，一键还原据此一次找齐。 */
  kind?: "template";
  /** 套模板前的页边距/页眉页脚/超链接配色；它们不走段落 op，只能靠这份留底复原。 */
  docLevel?: Record<string, unknown>;
}

export const SYSTEM_PROMPT = `你是文档编辑助手，可以直接修改用户的文档。你会看到文档大纲，每行格式为：
[索引] 类型 "文字" {当前样式}
其中类型为 P（正文）、H1…H6（标题）、TABLE、IMAGE。

只输出一个 JSON 对象，不要解释、不要代码块围栏。结构：
{"summary":"一句话说明这次改动做了什么","ops":[ 操作列表 ]}

可用操作：
1. {"op":"setText","index":N,"text":"整段新文字"} — 重写第 N 段。扩写、缩写、润色、改错别字、翻译都用它。text 必须是完整的新段落文字，不是差异片段。
2. {"op":"setStyle","index":N,"style":{...}} — 修改第 N 段的样式/格式。style 支持的字段：
   headingLevel(1~6，null 表示降为正文)、align("left"|"center"|"right"|"justify")、
   bold/italic/underline(true|false)、
   fontSize("16px")、fontFamily("宋体")、fontFamilyWest("Times New Roman")、
   letterSpacing("0.67px")、color("#cc0000")、highlight("#ffff00")、lineHeight("1.5")、
   spacingBefore/spacingAfter(单位磅)、indent/indentRight/indentFirst(单位磅，indentFirst 为负表示悬挂缩进)。把不需要的字段设成 null 表示清除。
   段内只有部分文字带某种字符样式（如加粗、颜色）时，整段设置那一项会被拒绝，只设其余字段。
3. {"op":"insertParagraph","after":N,"text":"...","style":{...}} — 在第 N 段后插入。after=-1 插到文首。一次只插一段，text 里不要放换行。
4. {"op":"insertTable","after":N,"rows":[["列1","列2"],["值1","值2"]],"headerRow":true} — 插入表格，rows 必须是等宽的二维数组。
5. {"op":"insertDiagram","after":N,"alt":"图注","diagram":{"title":"图标题","layout":"flow-v|flow-h|grid","nodes":[{"id":"a","label":"步骤一","shape":"box|round|diamond"}],"edges":[{"from":"a","to":"b","label":"可选条件"}]}} — 生成一张 canvas 示意图（流程图/结构图）插入文档，用来增强说明。
6. {"op":"deleteBlock","index":N} — 删除第 N 段。
7. {"op":"replaceAll","find":"旧词","replace":"新词"} — 全文替换一个词。

规则：
- 只输出那个 JSON 对象本身：不要复述大纲、不要写任何解释文字、不要加代码块围栏。
- 已有段落的 index/after 用大纲里的编号。写新内容时每条 insertParagraph 的 after 指向"写到这里为止文档里的段落编号"，也就是刚插入那一段的编号（新文档从 after=-1 开始，接着 0、1、2…依次递增）。
- 自动排版：给标题设 headingLevel/align/bold，给正文设 fontSize、lineHeight、indentFirst 等；排版任务不要改文字。
- 纠正错别字：只改确实有错的段落，并在 summary 里列出改了哪几个词；没有错误就返回 {"summary":"未发现错别字","ops":[]}。
- 保持文档原有语言，除非用户明确要求翻译。
- 写新内容（章节、小说、长文）要一段一条 op：需要几段就发几条 insertParagraph，一条 text 里塞多行会被拆成多段，标题就得另发一条带 headingLevel 的。
- text 里的对白、引用一律用中文引号（“” 或 「」）。绝对不要往 text 里塞没转义的英文双引号 " —— 它会把整份 JSON 写坏，一个字都落不了地。
- 一次改动尽量精简，只做用户要求的事。`;

export function buildUserMessage(instruction: string, blocks: Block[], selection?: string): string {
  const parts = [`文档大纲：\n${buildOutline(blocks)}`];
  if (selection && selection.trim()) {
    parts.push(`用户当前选中的文字：\n${selection.slice(0, 2000)}`);
  }
  parts.push(`指令：${instruction}`);
  return parts.join("\n\n");
}

/**
 * 从模型回复里取出提案 JSON。
 *
 * @param text 模型原文，可能带解释文字或 ```json 围栏，也可能被截断。
 * @returns `value` 是解析出的对象（取不到就是 null）；`truncated` 表示只能抢救到
 *   最后一个完整的 op——输出在中途断了，前面的 op 依然可用。
 */
export function extractJson(text: string): { value: unknown | null; truncated: boolean } {
  const withoutFence = text.replace(/```(?:json)?/gi, "").trim();
  // 严格那份先试：修文本是猜模型的意思，能不改就不改。
  const passes = [pickProposal(withoutFence), pickProposal(repairJsonText(withoutFence))];
  for (const p of passes) if (p.value !== null && !p.truncated) return p;
  for (const p of passes) if (p.value !== null) return p;
  return { value: null, truncated: false };
}

/**
 * 在一段（可能不合法的）文本里找提案对象。
 *
 * @param text 去掉代码块围栏后的文本。
 * @returns 提案对象与是否被截断；取不到时 `value` 为 null。
 */
function pickProposal(text: string): { value: unknown | null; truncated: boolean } {
  const candidates: { body: string; scan: ReturnType<typeof scanBraces> }[] = [];
  for (let i = text.indexOf("{"); i >= 0; i = text.indexOf("{", i + 1)) {
    const body = text.slice(i);
    candidates.push({ body, scan: scanBraces(body) });
  }

  // 提示词里就带着 `{当前样式}` 这种散文花括号，模型复述大纲时它会先被扫到；
  // op 对象本身也能 parse 成 JSON，所以只认长得像提案的那个候选。
  for (const { body, scan } of candidates) {
    if (scan.closed === null) continue;
    const value = parseJsonObject(body.slice(0, scan.closed));
    if (isProposalShape(value)) return { value, truncated: false };
  }
  // 输出被 max_tokens 截断：丢掉最后一条没写完的 op，补齐右括号，能落地多少算多少。
  for (const { body, scan } of candidates) {
    if (scan.memberEnd === null) continue;
    const value = parseJsonObject(`${body.slice(0, scan.memberEnd)}${scan.closers}`);
    if (isProposalShape(value)) return { value, truncated: true };
  }
  return { value: null, truncated: false };
}

/**
 * parse 一个 JSON 对象字面量。
 *
 * @param text 候选片段。
 * @returns 解析出的对象；不是对象或语法不合法时 null。
 */
function parseJsonObject(text: string): Record<string, unknown> | null {
  try {
    const v: unknown = JSON.parse(text);
    return typeof v === "object" && v !== null && !Array.isArray(v) ? (v as Record<string, unknown>) : null;
  } catch {
    return null;
  }
}

/**
 * 把模型写歪的 JSON 修成能 parse 的文本，只在严格解析失败之后才跑。
 *
 * 两类高频毛病：正文里该写 `\n` 却打了裸回车；对白里的英文直双引号没写成 `\"`
 * （`"text":"他说"你好"。"` 这种，一个引号就能让整份提案作废）。
 *
 * @param text 原始文本（可以带 JSON 前后的说明文字）。
 * @returns 转义后的文本；结构字符（`, } ] :`）紧邻的引号才当结束引号。
 */
function repairJsonText(text: string): string {
  let out = "";
  let inString = false;
  let escaped = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i]!;
    if (!inString) {
      if (ch === '"') inString = true;
      out += ch;
      continue;
    }
    if (escaped) {
      escaped = false;
      out += ch;
      continue;
    }
    if (ch === "\\") {
      escaped = true;
      out += ch;
      continue;
    }
    if (ch === '"') {
      // 只有后面紧跟结构符（或到了结尾）才是收束引号，否则它是正文里的引号。
      if (isQuoteTerminator(text, i + 1)) {
        inString = false;
        out += ch;
      } else {
        out += '\\"';
      }
      continue;
    }
    const code = ch.codePointAt(0)!;
    out += code < 0x20 ? escapeControlChar(ch) : ch;
  }
  return out;
}

/**
 * 判断字符串里的这个引号是不是真正的结束引号。
 *
 * @param text 全文。
 * @param next 引号之后第一个字符的下标。
 * @returns 之后（跳过空白）是 `,` `}` `]` `:` 或文本结束时为 true。
 */
function isQuoteTerminator(text: string, next: number): boolean {
  let i = next;
  while (i < text.length && /\s/.test(text[i]!)) i++;
  if (i >= text.length) return true;
  return text[i] === "," || text[i] === "}" || text[i] === "]" || text[i] === ":";
}

/**
 * 把裸控制字符写成 JSON 转义序列。
 *
 * @param ch codePoint 小于 0x20 的字符。
 * @returns 形如 `\n`、`\t`、`\u0001` 的两/六字符转义。
 */
function escapeControlChar(ch: string): string {
  if (ch === "\n") return "\\n";
  if (ch === "\r") return "\\r";
  if (ch === "\t") return "\\t";
  return `\\u${ch.codePointAt(0)!.toString(16).padStart(4, "0")}`;
}

/**
 * 判断一个对象是不是"修改方案"的形状。
 *
 * @param value 候选对象。
 * @returns 带 ops（或其别名）或 summary 之一时为 true；单条 op 之类不算。
 */
function isProposalShape(value: Record<string, unknown> | null): value is Record<string, unknown> {
  if (!value) return false;
  return ["ops", "operations", "操作", "summary", "说明", "title"].some((k) => k in value);
}

/**
 * 单遍扫描 JSON 文本的括号结构，字符串字面量里的括号与转义都不计数。
 *
 * @param body 从某个 `{` 开始的文本，允许结尾被截断。
 * @returns `closed` 根对象配平处的下标（未配平则 null）；`memberEnd` 最后一个完整成员的
 *   结束下标；`closers` 是在 `memberEnd` 处把 JSON 补合法所需的右括号（从里到外）。
 */
function scanBraces(body: string): { closed: number | null; memberEnd: number | null; closers: string } {
  const stack: string[] = [];
  let memberEnd: number | null = null;
  let closers = "";
  let inString = false;
  let escaped = false;
  for (let i = 0; i < body.length; i++) {
    const ch = body[i];
    if (inString) {
      if (escaped) escaped = false;
      else if (ch === "\\") escaped = true;
      else if (ch === '"') inString = false;
      continue;
    }
    if (ch === '"') inString = true;
    else if (ch === "{") stack.push("}");
    else if (ch === "[") stack.push("]");
    else if (ch === "}" || ch === "]") {
      if (stack.pop() !== ch) return { closed: null, memberEnd, closers: "" };
      if (stack.length === 0) return { closed: i + 1, memberEnd, closers: "" };
      // 刚闭合的是一个还能继续往下写的成员（一条 op、一个 style、ops 数组），
      // 在它后面切断再补上栈里剩下的右括号，就还是一份合法 JSON。
      memberEnd = i + 1;
      closers = stack.slice().reverse().join("");
    }
  }
  return { closed: null, memberEnd, closers };
}

export interface ProposalOutcome {
  ok: boolean;
  proposal?: Proposal;
  before: Block[];
  message?: string;
  needsSetup?: boolean;
  /** 用户按了暂停：半截内容不算失败，界面要给「继续」把剩下的补完。 */
  stopped?: boolean;
  warnings: string[];
}

/**
 * 从还没收完的模型输出里取出正文段落，给界面做"边收边显示"的预览。
 *
 * @param raw 目前收到的全部增量。
 * @returns 按文档顺序排好的段落文字；最后一段可能只写了一半，收完前它会长。
 */
export function streamedParagraphs(raw: string): string[] {
  return jsonStrings(raw, ["text", "alt"]).map((f) => f.text);
}

/**
 * 从还没收完的模型输出里取出提案的一句话说明。
 *
 * @param raw 目前收到的全部增量。
 * @returns 说明文字；`"summary"` 还没出现时是空串。
 */
export function streamedSummary(raw: string): string {
  return jsonStrings(raw, ["summary"])[0]?.text ?? "";
}

/**
 * 扫出给定键名下已经开头的字符串值，按它们在文本里出现的位置排序。
 *
 * @param raw 模型输出的全部增量。
 * @param keys 要读的键名（`summary`/`text`/`alt`）。
 * @returns `{ at, text }` 列表，`at` 是键名在 raw 里的下标；只有键名、值没开头的忽略。
 */
function jsonStrings(raw: string, keys: readonly string[]): { at: number; text: string }[] {
  const out: { at: number; text: string }[] = [];
  for (const key of keys) {
    // 前面挂着转义反斜杠的不是键名：正文里写的 "text":" 会被转义成 \"text\":\"。
    const re = new RegExp(`(?<!\\\\)"${key}"\\s*:\\s*"`, "g");
    let m: RegExpExecArray | null;
    while ((m = re.exec(raw))) {
      const value = readJsonString(raw, m.index + m[0].length);
      if (value?.trim()) out.push({ at: m.index, text: value.trim() });
    }
  }
  return out.sort((a, b) => a.at - b.at);
}

/**
 * 从指定位置读一个 JSON 字符串的值（不含首尾引号）。
 *
 * @param text 完整文本。
 * @param from 值起始的下标（正好在开引号之后）。
 * @returns 解码后的文字；引号还没闭合时返回写到哪算哪的半截（流式预览要的就是这个）。
 */
function readJsonString(text: string, from: number): string {
  let i = from;
  let escaped = false;
  while (i < text.length) {
    const ch = text[i];
    if (escaped) escaped = false;
    else if (ch === "\\") escaped = true;
    // 正文里没转义的引号后面挂的是字，不是 `, } ] :`，不能当结束引号——否则一段话
    // 会被模型自己写的引号截成两半显示。
    else if (ch === '"' && isQuoteTerminator(text, i + 1)) break;
    i++;
  }
  const frag = text.slice(from, i);
  if (i < text.length) return decodeJsonString(frag);
  // 引号未闭合：丢掉末尾写了一半的转义起点，其余照原样显示。
  return decodeJsonString(frag.replace(/\\$/, ""));
}

/**
 * 把 JSON 字符串体解码成真实文字。
 *
 * @param frag 引号之间（或截断处）的原始片段。
 * @returns 解码结果；转义序列本身还不完整时退回原文，宁可多显示一个反斜杠也不少一行。
 */
function decodeJsonString(frag: string): string {
  try {
    return JSON.parse(`"${frag}"`);
  } catch {
    return frag;
  }
}

/**
 * 走主进程的流式通道问模型一次，边收边把增量交给 `onDelta`。
 *
 * 界面靠它把"AI 正在输出"变成看得见的东西：loading 态 + 逐字滚出来的正文。
 * 失败一律 reject，错误对象上带 `needsSetup`（还没配模型时界面要拉开设置）。
 *
 * @param opts.action 主进程侧的内置语气键（`chat`/`扩写`…）；给了 `system` 时只用作日志。
 * @param opts.text 送给模型的用户消息（agent 路径就是大纲 + 指令）。
 * @param opts.system 覆盖内置系统提示词，文档编辑 agent 必须传它。
 * @param opts.instruction 拼在原文前面的一段说明，选区改写用。
 * @param opts.providerId 指定服务商；留空用当前生效的那个。
 * @param opts.requestId 本次请求 id，用于把增量认回正确的界面卡片，也让「停止」能取消。
 * @param opts.history 这个会话里已经发生过的轮次，续聊时接上文用（主进程只取最近若干轮）。
 * @param opts.onDelta 每段增量到达时调用（按收到的顺序）。
 * @returns 收完时 `{ stopped: false }`；被「停止」打断时 `{ stopped: true }`（已收到的增量照常留在调用方）；
 *   配置缺失、网络失败、超时都 reject。
 */
export function aiStreamTransform(opts: {
  action: string;
  text: string;
  system?: string;
  instruction?: string;
  providerId?: string | null;
  requestId: string;
  history?: SessionTurn[];
  onDelta: (chunk: string) => void;
}): Promise<{ stopped: boolean }> {
  return new Promise((resolve, reject) => {
    const fail = (message: string, needsSetup?: boolean) => {
      const err = new Error(message) as Error & { needsSetup?: boolean };
      if (needsSetup) err.needsSetup = true;
      reject(err);
    };
    const off = window.officeCopilot.onAIStream((ev) => {
      if (ev.requestId !== opts.requestId) return;
      if (ev.kind === "delta") opts.onDelta(ev.text ?? "");
      else if (ev.kind === "done") {
        off();
        resolve({ stopped: Boolean(ev.stopped) });
      } else {
        off();
        fail(ev.message ?? "调用失败", ev.needsSetup);
      }
    });
    window.officeCopilot
      .aiStreamStart({
        requestId: opts.requestId,
        action: opts.action,
        text: opts.text,
        instruction: opts.instruction,
        system: opts.system,
        providerId: opts.providerId ?? undefined,
        history: opts.history,
      })
      .then((r) => {
        if (!r.ok) {
          off();
          fail(r.message ?? "调用失败", r.needsSetup);
        }
      })
      .catch((e: unknown) => {
        off();
        fail(e instanceof Error ? e.message : String(e));
      });
  });
}

/**
 * 问模型要一份修改方案：大纲 + 指令送出去，回来的 JSON 校验成 Proposal。
 *
 * @param opts.editor 当前编辑器（大纲按它现算）。
 * @param opts.instruction 用户原话。
 * @param opts.providerId 指定服务商。
 * @param opts.selection 当前选区文字，会一并送出。
 * @param opts.onDelta 模型的逐字增量；界面用它显示"正在输出"。
 * @param opts.requestId 本次请求 id，界面用它把「停止」接到这条流上。
 * @param opts.continueFrom 续写：上一次暂停时已收到的改动与那半截原文，本次只补剩下的部分。
 * @returns 方案与送出去时的大纲快照；解析失败或调用失败时 `ok` 为 false 并带原因，
 *   被暂停打断时 `stopped` 为 true（已经收到的那部分照样可以用）。
 */
export async function requestProposal(opts: {
  editor: Editor;
  instruction: string;
  providerId?: string | null;
  selection?: string;
  onDelta?: (chunk: string) => void;
  requestId?: string;
  continueFrom?: { ops: AiOp[]; rawTail: string };
}): Promise<ProposalOutcome> {
  const before = blocksOf(opts.editor.state.doc);
  const instruction = opts.continueFrom
    ? continuationInstruction(opts.instruction, opts.continueFrom, before.length)
    : opts.instruction;
  const userMessage = buildUserMessage(instruction, before, opts.selection);
  let acc = "";
  let stopped = false;
  try {
    const r = await aiStreamTransform({
      action: "agent",
      text: userMessage,
      system: SYSTEM_PROMPT,
      providerId: opts.providerId ?? null,
      requestId: opts.requestId ?? crypto.randomUUID(),
      onDelta: (chunk) => {
        acc += chunk;
        opts.onDelta?.(chunk);
      },
    });
    stopped = r.stopped;
  } catch (e) {
    const err = e as Error & { needsSetup?: boolean };
    return {
      ok: false,
      before,
      message: err.message ?? String(e),
      needsSetup: err.needsSetup,
      warnings: [],
    };
  }
  const parsed = extractJson(acc);
  if (parsed.value === null) {
    if (stopped) {
      return {
        ok: false,
        stopped: true,
        before,
        message: "已暂停生成，这一段还没写出可用的改动。点「继续」让我接着写完。",
        warnings: [],
      };
    }
    // 把模型回的前一小段贴出来：解析失败的原因（它压根没回 JSON、还是回歪了）一眼能看出来。
    const peek = acc.replace(/\s+/g, " ").trim().slice(0, 160);
    return {
      ok: false,
      before,
      message: `模型没有返回可解析的修改方案${peek ? `，它回的是：${peek}` : ""}。请重试或换个说法。`,
      warnings: [],
    };
  }
  // 续写时前面已经插入的段落也算进可引用段数，否则模型接着排的编号会被当成越界丢掉。
  const prior = opts.continueFrom?.ops ?? [];
  const coerced = coerceProposal(parsed.value, before.length + insertedParagraphCount(prior));
  if (!coerced.proposal) {
    const errors = stopped ? [`已暂停生成。${coerced.errors.join("；")}`] : coerced.errors;
    return { ok: false, before, message: errors.join("；"), warnings: errors, stopped };
  }
  const proposal: Proposal = prior.length
    ? { summary: coerced.proposal.summary, ops: [...prior, ...coerced.proposal.ops] }
    : coerced.proposal;
  // 截断或被暂停的输出只抢救到前一半：明说少了多少，别让文档悄悄短一截。
  const warnings = stopped
    ? [
        `已暂停生成，目前收到 ${proposal.ops.length} 条修改，点「继续」把剩下的补完。`,
        ...coerced.errors,
      ]
    : parsed.truncated
      ? [
          `模型输出被截断，只收到前 ${proposal.ops.length} 条修改，可以再说「继续」让它往下写。`,
          ...coerced.errors,
        ]
      : coerced.errors;
  return { ok: true, proposal, before, warnings, stopped };
}

/**
 * 把"上一次停在半路"改写成一条接着写的指令。
 *
 * @param instruction 用户原话，续写时原样带上，模型才知道整体要做什么。
 * @param from 已收到的改动与那半截原文。
 * @param blockCount 送出去时文档的段数。
 * @returns 拼好的指令（仍然要求一份完整 JSON，只是内容只补剩下的部分）。
 */
function continuationInstruction(
  instruction: string,
  from: { ops: AiOp[]; rawTail: string },
  blockCount: number,
): string {
  const tail = from.rawTail.replace(/\s+/g, " ").trim().slice(-400);
  return (
    `${instruction}\n\n` +
    `上一条方案写到一半被打断了，已经收下 ${from.ops.length} 条改动，它停在这句话中间：\n…${tail}\n\n` +
    `不要重复已经给出的改动，只输出剩下的部分，仍然是一份完整的 {"summary":"…","ops":[…]}。` +
    `新增段落的 after 从 ${nextInsertAfter(from.ops, blockCount)} 往上接着排。`
  );
}

/** 续写的下一条插入该用哪个 after：排在已给出插入编号的后面，没插入过就是文档末尾。 */
function nextInsertAfter(ops: AiOp[], blockCount: number): number {
  let last = blockCount - 1;
  for (const o of ops) if (o.op === "insertParagraph" && o.after > last) last = o.after;
  return last + 1;
}

/**
 * 全文改写时追加的约束：落地走 `alignRewrite`，段数对不上就整批放弃，
 * 所以必须一开始就说清"一段对一段"，别让模型自己合并段落。
 */
export const FULL_DOC_REWRITE_NOTE =
  "下面这份是整篇文档。请按原有段落逐段改写：保持段落数量与顺序，一段都不合并、不新增、不删除；只用换行分段，不要编号、不要解释、不要代码块围栏。";

/**
 * 组装快捷指令改写时的附加说明（主进程把它拼在原文前面）。
 *
 * @param scope document 时要求段数一一对应；selection 不加额外约束。
 * @param resume 暂停后的续写进度（docEdit 的 `rewriteResume`）：`kept` 是留下不动的段落，
 *   `remaining` 还要模型补写的段数。给了它，任务就从"改写整篇"变成"只补第 kept+1 段起的剩下几段"。
 * @returns 说明文字；没有要补的约束时返回 undefined，走主进程内置语气原样。
 */
export function rewriteInstruction(
  scope: "selection" | "document",
  resume?: { kept: string[]; remaining: number },
): string | undefined {
  const rule = scope === "document" ? FULL_DOC_REWRITE_NOTE : "";
  if (!resume) return rule || undefined;
  const done = resume.kept.length;
  const head = done ? `前 ${done} 段已经改好、我会自己留着：\n${resume.kept.join("\n")}\n\n` : "";
  return `${rule}\n\n${head}上一次改写在某一段中间被打断，那段残句不算数，我从头重发原文。请只改写剩下的 ${resume.remaining} 段（从第 ${done + 1} 段起），正好 ${resume.remaining} 行，每段一行；不要重复上面已经改好的段落，也不要写解释和代码块围栏。`.trim();
}

/** 这些改动会新增多少段（`after` 的可引用段数要跟着长）。 */
function insertedParagraphCount(ops: AiOp[]): number {
  return ops.reduce((n, o) => (o.op === "insertParagraph" ? n + Math.max(1, paragraphTexts(o.text).length) : n), 0);
}

export interface CommitResult {
  ok: boolean;
  message?: string;
  record?: ModificationRecord;
}

/**
 * Apply already-resolved ops as one transaction and log the modification so it
 * can be diffed and undone on its own later.
 */
export async function commitOps(
  editor: Editor,
  opts: {
    ops: CoreOp[];
    before: Block[];
    after: Block[];
    inverses: CoreOp[];
    label: string;
    summary: string;
    docKey: string;
    kind?: "template";
    docLevel?: Record<string, unknown>;
  },
): Promise<CommitResult> {
  if (!opts.ops.length) return { ok: false, message: "没有需要写入的修改。" };
  const exec = executeOps(editor, opts.ops, opts.after);
  if (!exec.ok) return { ok: false, message: exec.reason ?? "写入失败。" };

  const record: ModificationRecord = {
    id: crypto.randomUUID(),
    ts: Date.now(),
    label: opts.label,
    summary: opts.summary,
    ops: opts.ops,
    inverses: opts.inverses,
    before: opts.before,
    after: exec.after,
    undone: false,
    ...(opts.kind ? { kind: opts.kind } : null),
    ...(opts.docLevel ? { docLevel: opts.docLevel } : null),
  };
  const saved = await window.officeCopilot.appendModification({ docKey: opts.docKey, record });
  if (!saved.ok) {
    // The edit is in the document; losing the history entry is not worth
    // rolling the user's document back, so surface it instead.
    return { ok: true, record, message: `修改已写入，但历史记录保存失败：${saved.message ?? "未知错误"}` };
  }
  return { ok: true, record };
}

export interface PlanOutcome {
  ok: boolean;
  message?: string;
  ops?: CoreOp[];
  before?: Block[];
  after?: Block[];
  inverses?: CoreOp[];
  skipped: string[];
}

/** Validate + simulate a proposal against the live document without writing. */
export function simulateProposal(editor: Editor, proposal: Proposal): PlanOutcome {
  const before = blocksOf(editor.state.doc);
  const plan = planEdit(before, proposal, (spec, alt) => diagramResolverFor(spec, alt));
  if (!plan.ok) {
    return { ok: false, message: plan.skipped.join("；") || "提案无法执行。", skipped: plan.skipped };
  }
  return {
    ok: true,
    ops: plan.ops,
    before: plan.before,
    after: plan.after,
    inverses: plan.inverses,
    skipped: plan.skipped,
  };
}

export async function applyProposal(
  editor: Editor,
  proposal: Proposal,
  docKey: string,
): Promise<CommitResult> {
  const plan = simulateProposal(editor, proposal);
  if (!plan.ok) return { ok: false, message: plan.message };
  return commitOps(editor, {
    ops: plan.ops!,
    before: plan.before!,
    after: plan.after!,
    inverses: plan.inverses!,
    label: "AI 修改",
    summary: proposal.summary,
    docKey,
  });
}

/**
 * Roll back one recorded modification. Later unrelated edits do not block it:
 * each inverse is matched by content. Ops whose target the user has since
 * changed are reported and skipped rather than guessed at.
 */
export async function undoModification(
  editor: Editor,
  record: ModificationRecord,
  docKey: string,
): Promise<CommitResult & { failed: string[] }> {
  if (record.undone) {
    return { ok: false, message: "这一条修改已经撤销过了。", failed: [] };
  }
  const current = blocksOf(editor.state.doc);
  const r = revert(current, record.inverses);
  if (!r.applied.length) {
    return {
      ok: false,
      message: `无法撤销：这些段落自那次修改后已被改动。\n${r.failed.join("\n")}`,
      failed: r.failed,
    };
  }
  const exec = executeOps(editor, r.applied, r.ok ? r.after : undefined);
  if (!exec.ok) {
    return { ok: false, message: exec.reason ?? "撤销失败。", failed: r.failed };
  }
  await window.officeCopilot.patchModification({
    docKey,
    id: record.id,
    patch: { undone: true, undoneAt: Date.now() },
  });
  const message = r.ok
    ? undefined
    : `部分撤销（${r.applied.length}/${record.inverses.length}）：${r.failed.join("；")}`;
  return { ok: true, message, failed: r.failed };
}

/* ------------------------------------------------------------------ */
/* Non-agent edits that still need to land in the history              */
/* ------------------------------------------------------------------ */

/**
 * Rewrite the current selection with `rewritten` as block-level setText ops, so
 * a bubble rewrite is undoable like any other AI edit.
 */
export async function commitSelectionRewrite(
  editor: Editor,
  opts: { from: number; to: number; rewritten: string; label: string; docKey: string },
): Promise<CommitResult> {
  const before = blocksOf(editor.state.doc);
  const located = locateAll(editor.state.doc);
  const touched = located.filter(
    (l) => l.inner && l.inner.from < opts.to && l.inner.to > opts.from,
  );
  if (!touched.length) return { ok: false, message: "选区不在可改写的段落内。" };

  const clean = normalise(opts.rewritten);
  const ops: CoreOp[] = [];
  touched.forEach((l, i) => {
    const inner = l.inner!;
    const blockIndex = l.block.index;
    const addr = addressOf(before, blockIndex);
    if (!addr) return;
    const prefix = i === 0 ? l.block.text.slice(0, Math.max(0, opts.from - inner.from)) : "";
    const suffix = i === touched.length - 1 ? l.block.text.slice(Math.max(0, opts.to - inner.from)) : "";
    const text = i === 0 ? normalise(prefix + clean + suffix) : normalise(suffix);
    ops.push({ kind: "setText", addr, text });
  });
  if (!ops.length) return { ok: false, message: "选区不在可改写的段落内。" };

  const run = runOps(before, ops);
  return commitOps(editor, {
    ops: run.applied,
    before,
    after: run.after,
    inverses: run.inverses,
    label: opts.label,
    summary: `${opts.label}：${ops.map(describeCoreOp).join("，")}`,
    docKey: opts.docKey,
  });
}

/** Insert assistant text as a real modification instead of a raw paste. */
export async function commitInsertedText(
  editor: Editor,
  opts: { text: string; docKey: string; at: "cursor" | "end"; label?: string },
): Promise<CommitResult> {
  const before = blocksOf(editor.state.doc);
  const paragraphs = opts.text.split(/\n+/).map((s) => s.trim()).filter(Boolean);
  if (!paragraphs.length) return { ok: false, message: "没有可插入的内容。" };

  let anchorIndex = before.length - 1;
  if (opts.at === "cursor") {
    const pos = editor.state.selection.from;
    const hit = locateAll(editor.state.doc).find((l) => l.inner && l.inner.from <= pos && l.inner.to >= pos);
    if (hit) anchorIndex = hit.block.index;
  }
  // Sequential inserts must stack: each one follows the previous index.
  const aiOps: AiOp[] = paragraphs.map((text, i): AiOp => ({
    op: "insertParagraph",
    after: anchorIndex + i,
    text,
  }));
  const proposal: Proposal = {
    summary: opts.label ?? "AI 插入文案",
    ops: aiOps,
  };
  return applyProposal(editor, proposal, opts.docKey);
}
