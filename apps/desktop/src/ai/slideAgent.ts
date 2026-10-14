/**
 * 幻灯片的 AI 代理（spec 阶段 5 / 决策 D6）：把"改标题 / 统一字号 / 把某个框挪位"这类诉求跑成一串工具回合，
 * 落地只认 `ppt_submit_proposal` 交出的那批改动，并且**只交提案，不写文件**（`AGENTS.md` §2.8）。
 *
 * 与 `sheetAgent.ts` 同一条形状，三条约束照抄：
 * ① 页/形状/段/run 这些事实由 `@office-copilot/pptx-io` 从现场模型读（`describeDeck`/`readSlideText`），这里只搬运（§2.11）；
 * ② 上送内容默认截断（§2.9）：页数、每页形状数、每段 run 数、单条 run 文字长度都有上限，超出的部分引擎会标出来；
 * ③ 模型给的非法入参与越界页号都换成回给它的一句话，永不抛异常（§7.1 的 mock 用例逐条钉住）。
 *
 * 交出去的 `edits` 就是视图"保存"用的那批 `PptxEditRequest`（`pptx-io/edit.ts`）：确认后原样送 `ppt:edit`，
 * 主进程逐条现做带前像的 op——所以这里不需要、也不许再长第二套落地或撤销通道（§3.2）。
 */
import {
  describeDeck,
  readSlideText,
  type PptxEditRequest,
  type PresentationModel,
  type SlideShape,
} from "@office-copilot/pptx-io";
import {
  parseToolArguments,
  PPTX_TOOLS,
  type PptxToolName,
  type PptxEditT,
  type SubmitProposalArgsT,
} from "@office-copilot/pptx-tools";
import type { ToolCall, ToolTurnMessage, ToolTurnRequest, ToolTurnResult } from "../global.js";

/** 一次任务里最多转多少轮：模型无限要工具就该停下来问人，而不是无限烧 token。 */
const MAX_TOOL_TURNS = 8;

/** 首条用户消息里那份画像的上限：给模型看得见全貌，但不许把整份 deck 摊上去（§2.9）。 */
const CONTEXT_MAX_SLIDES = 20;
const CONTEXT_MAX_SHAPES_PER_SLIDE = 20;

const SLIDE_SYSTEM = `你是演示文稿（.pptx）代理，看得见的是这份文稿的结构画像，不是整份原文。
- 动手前先调 ppt_describe 拿页号与 shapeId；要改某段文字再调 ppt_read_slide_text 确认 paraIndex / runIndex，不许凭印象数，也不许自己编形状号。
- 要改只能调 ppt_submit_proposal，条目只有 set_run_text / set_run_font / set_run_color / set_shape_geometry / move_shape_layer 五种，
  前三条指到"某一页某个形状的某段某 run"，后两条指到形状（叠放方向是 up/down/top/bottom）。增删形状、改版式与母式这里做不了，遇到就直说做不了。
- 一段里如果有多条 run（read_slide_text 会标出来），必须逐 run 改，整段覆盖会拍平那一截的粗体与字号——这是禁止的。
- 字号单位是 pt，几何单位是屏幕 px；EMU、半磅这些换算由引擎做，你不许换算数字往字段里塞。
- 颜色只认 #rrggbb 六位十六进制。
- 纯分析的问题直接回答，用中文，带上你实际读到的页号、形状名与文字，不要为了回答去提提案。`;

/** `runSlideAgent` 的输入。 */
export interface SlideAgentInput {
  /** 主进程读出来的当前模型（只读：代理不改它，落地由上层确认后走 `ppt:edit`）。 */
  deck: PresentationModel;
  /** 用户那句诉求原文。 */
  instruction: string;
  /** 已经跑过的 transcript：暂停后继续时原样带回。 */
  messages?: ToolTurnMessage[];
  /** 指定服务商；省略即用设置里激活的那个。 */
  providerId?: string;
  /** 发一轮请求的通道，默认走 preload 的 `ai:tool-turn`，单测里换成桩。 */
  ask?: (payload: ToolTurnRequest) => Promise<ToolTurnResult>;
  /** 用户按了暂停：下一轮请求发出前停下，transcript 原样带回。 */
  isPaused?: () => boolean;
}

/** 一次幻灯片任务的结果。`messages` 每支都带，UI 靠它接着往下跑。 */
export type SlideAgentOutcome =
  | {
      kind: "proposal";
      summary: string;
      /** 翻成引擎坐标的那批改动，确认后原样送 `ppt:edit` 即可落地并拿到逆操作（§2.5）。 */
      edits: PptxEditRequest[];
      /** 提案里没能变成改动的条目，逐条中文理由，不许静默丢掉（§2.5）。 */
      notes: string[];
      /** 模型原样的提案条目（给用户看"它到底提了什么"，单位是模型视角）。 */
      proposed: PptxEditT[];
      messages: ToolTurnMessage[];
    }
  | { kind: "answer"; text: string; messages: ToolTurnMessage[] }
  | { kind: "paused"; messages: ToolTurnMessage[] }
  | { kind: "failed"; message: string; needsSetup?: boolean; messages: ToolTurnMessage[] };

/**
 * 组装上送的文稿画像：页数 + 每页形状清单与文字概要（由引擎截断，这里不再二次裁剪）。
 *
 * @param deck 当前模型。
 * @returns 多行中文文本，作为本轮用户消息的前缀。
 */
export function slideContextOf(deck: PresentationModel): string {
  const outline = describeDeck(deck, { maxSlides: CONTEXT_MAX_SLIDES, maxShapesPerSlide: CONTEXT_MAX_SHAPES_PER_SLIDE });
  const lines = [
    `演示文稿：${outline.slideCount} 页，页面 ${outline.slideSizePx.widthPx}×${outline.slideSizePx.heightPx} px。`,
    `画像：${JSON.stringify(outline.slides)}`,
  ];
  if (outline.slidesTruncated) lines.push(`页太多，只描述了前 ${outline.slides.length} 页，剩下的页要用 ppt_describe 加大 maxSlides 才看得见。`);
  return lines.join("\n");
}

/**
 * 把提案里的条目翻成引擎认的跨界改动（部件路径 + 形状号 + 段/run 序号）。
 *
 * 页号是模型视角的 1 基，部件路径由现场模型算——这是唯一一处做这个映射（§3.3）。
 * 翻不动的（页号越界、该页没有这个形状、几何一条属性都没给）逐条记进 `notes`，
 * 不许交出一批主进程必然拒绝、整批不落地的改动（§2.4 的事务语义）。
 *
 * @param deck 提案要落在这份文稿上（页号→部件路径按它算）。
 * @param edits `ppt_submit_proposal` 校验过的条目。
 * @returns `edits` 按提案顺序排好；`notes` 是没变成改动的条目的中文说明。
 */
export function slideRequestsOfProposal(deck: PresentationModel, edits: PptxEditT[]): { requests: PptxEditRequest[]; notes: string[] } {
  const requests: PptxEditRequest[] = [];
  const notes: string[] = [];
  for (const edit of edits) {
    const slide = deck.slides.find((entry) => entry.order + 1 === edit.slideNumber);
    if (!slide) {
      notes.push(`没有第 ${edit.slideNumber} 页（这份文稿共 ${deck.slides.length} 页），这条跳过。`);
      continue;
    }
    const shape = slide.shapes.find((entry) => entry.shapeId === edit.shapeId);
    if (!shape) {
      notes.push(`第 ${edit.slideNumber} 页没有 shapeId=${edit.shapeId} 这个形状，这条跳过。`);
      continue;
    }
    if (edit.type === "set_shape_geometry") {
      // 组合形状的几何引擎会拒（子偏移语义不同）：这里就别交，否则主进程整批不落地（§2.4）。
      if (isGroupShape(shape)) {
        notes.push(`第 ${edit.slideNumber} 页的「${shape.name}」是组合形状，几何改不了（子偏移语义不同），这条跳过。`);
        continue;
      }
      requests.push({
        kind: "geometry",
        target: { partPath: slide.partPath, shapeId: shape.shapeId },
        box: { leftPx: edit.leftPx, topPx: edit.topPx, widthPx: edit.widthPx, heightPx: edit.heightPx },
      });
      continue;
    }
    if (edit.type === "move_shape_layer") {
      // 已在某一端时引擎会整批拒绝（§2.4），所以按该页形状顺序先判一刀再决定交不交。
      const edgeNote = layerEdgeNote(slide.shapes, shape, edit.direction);
      if (edgeNote) {
        notes.push(`第 ${edit.slideNumber} 页「${shape.name}」那条改不了：${edgeNote}`);
        continue;
      }
      requests.push({ kind: "layer", target: { partPath: slide.partPath, shapeId: shape.shapeId }, direction: edit.direction });
      continue;
    }
    const outOfRange = runRangeNote(shape, edit.paraIndex, edit.runIndex);
    if (outOfRange) {
      notes.push(`第 ${edit.slideNumber} 页「${shape.name}」那条改不了：${outOfRange}`);
      continue;
    }
    const target = { partPath: slide.partPath, shapeId: shape.shapeId, paraIndex: edit.paraIndex, runIndex: edit.runIndex };
    if (edit.type === "set_run_text") requests.push({ kind: "text", target, text: edit.text });
    if (edit.type === "set_run_color") requests.push({ kind: "color", target, hex: edit.color });
    if (edit.type === "set_run_font") requests.push({ kind: "font", target, format: { bold: edit.bold, sizePt: edit.sizePt } });
  }
  return { requests, notes };
}

/**
 * 判断一个形状是不是组合形状（几何 op 对它没意义，引擎会拒）。
 *
 * @param shape 形状快照。
 * @returns true 表示 `kind === "group"`。
 */
function isGroupShape(shape: SlideShape): boolean {
  return shape.kind === "group";
}

/**
 * 叠放方向是不是已经到头。
 *
 * `slide.shapes` 就是 `<p:spTree>` 里形状的文档顺序（读侧按它列），而那个顺序就是叠放次序：
 * 第一个在最底、最后一个在最顶。所以只要比对首尾，不必再解析一次 XML。
 *
 * @param shapes 该页全部形状（文档顺序）。
 * @param shape 本次要挪的那个。
 * @param direction 模型给的方向。
 * @returns 一句中文理由（该页只有一个形状、或已在要去的那一端）；能挪时返回 null。
 */
function layerEdgeNote(
  shapes: SlideShape[],
  shape: SlideShape,
  direction: "up" | "down" | "top" | "bottom",
): string | null {
  const at = shapes.indexOf(shape);
  if (shapes.length < 2) return "这一页只有这一个形状，没有可叠放的邻居";
  if ((direction === "up" || direction === "top") && at === shapes.length - 1) return "它已经在最顶层";
  if ((direction === "down" || direction === "bottom") && at === 0) return "它已经在最底层";
  return null;
}

/**
 * 段号与 run 号对不对得上现场模型。
 *
 * @param shape 该形状快照（段数与每段 run 数都来自它）。
 * @param paraIndex 模型给的段号（0 基）。
 * @param runIndex 模型给的 run 号（0 基）。
 * @returns 一句中文理由；对得上时返回 null。
 */
function runRangeNote(shape: SlideShape, paraIndex: number, runIndex: number): string | null {
  if (shape.paragraphTexts.length === 0) return "这个形状没有文字体（图片或图形框），改不了文字";
  if (paraIndex >= shape.paragraphTexts.length) return `段落序号 ${paraIndex} 越界（只有 ${shape.paragraphTexts.length} 段）`;
  const runCount = shape.paragraphRunCounts[paraIndex] ?? 0;
  if (runCount === 0) return `第 ${paraIndex} 段的文字来自字段（页码 / 日期），不在可改范围`;
  if (runIndex >= runCount) return `第 ${paraIndex} 段只有 ${runCount} 个 run，没有第 ${runIndex} 个`;
  return null;
}

/**
 * 跑一次工具调用：入参先过 `parseToolArguments`，再交给 `pptx-io` 算。
 *
 * @param deck 当前模型（只读）。
 * @param call 模型点名的工具与入参原文。
 * @returns 回给模型的文本；`proposal` 非空表示模型交了提案，回路该收尾。
 */
function executeToolCall(deck: PresentationModel, call: ToolCall): { content: string; proposal?: SubmitProposalArgsT } {
  switch (call.name as PptxToolName) {
    case "ppt_describe": {
      const parsed = parseToolArguments("ppt_describe", call.arguments);
      return { content: parsed.ok ? JSON.stringify(describeDeck(deck, parsed.args)) : parsed.reason };
    }
    case "ppt_read_slide_text": {
      const parsed = parseToolArguments("ppt_read_slide_text", call.arguments);
      if (!parsed.ok) return { content: parsed.reason };
      const read = readSlideText(deck, parsed.args.slideNumber, parsed.args.maxRunsPerParagraph);
      return { content: read.ok ? JSON.stringify(read.report) : read.reason };
    }
    case "ppt_submit_proposal": {
      const parsed = parseToolArguments("ppt_submit_proposal", call.arguments);
      if (!parsed.ok) return { content: parsed.reason };
      return {
        content: `提案已收下（${parsed.args.edits.length} 条改动），等用户确认后才会落地。要继续核对就接着调读工具，别重复提交。`,
        proposal: parsed.args,
      };
    }
    default:
      return { content: `不认识的工具：${call.name}。可用面只有 ppt_describe / ppt_read_slide_text / ppt_submit_proposal。` };
  }
}

/**
 * 跑一次幻灯片任务：转工具回合，直到模型交提案、直接作答、被暂停或出错。
 *
 * @param input 见 `SlideAgentInput`；`deck` 全程只读。
 * @returns `SlideAgentOutcome`。没配服务商、HTTP 失败、超过 8 轮都走 `failed`，不抛异常。
 */
export async function runSlideAgent(input: SlideAgentInput): Promise<SlideAgentOutcome> {
  const ask = input.ask ?? ((payload: ToolTurnRequest) => window.officeCopilot.aiToolTurn(payload));
  const messages: ToolTurnMessage[] = input.messages?.length
    ? input.messages.map((message) => ({ ...message }))
    : [
        { role: "system", content: SLIDE_SYSTEM },
        { role: "user", content: `${slideContextOf(input.deck)}\n\n诉求：${input.instruction}` },
      ];

  for (let turn = 0; turn < MAX_TOOL_TURNS; turn += 1) {
    if (input.isPaused?.()) return { kind: "paused", messages };
    const result = await ask({ messages, tools: PPTX_TOOLS, providerId: input.providerId });
    if (!result.ok) return { kind: "failed", message: result.message, needsSetup: result.needsSetup, messages };
    // 模型没点工具就是说完话了：这一轮的正文即为答复（纯分析类任务走这里，不出提案）。
    if (!result.toolCalls.length) return { kind: "answer", text: result.content, messages };

    messages.push({
      role: "assistant",
      ...(result.content ? { content: result.content } : {}),
      tool_calls: result.toolCalls.map((call) => ({
        id: call.id,
        type: "function" as const,
        function: { name: call.name, arguments: call.arguments },
      })),
    });
    for (const call of result.toolCalls) {
      const executed = executeToolCall(input.deck, call);
      messages.push({ role: "tool", tool_call_id: call.id, content: executed.content });
      if (!executed.proposal) continue;
      const { requests, notes } = slideRequestsOfProposal(input.deck, executed.proposal.edits);
      return { kind: "proposal", summary: executed.proposal.summary, edits: requests, notes, proposed: executed.proposal.edits, messages };
    }
  }
  return { kind: "failed", message: `工具回合到了 ${MAX_TOOL_TURNS} 轮上限，模型还没交提案或给出结论，先停下。`, messages };
}
