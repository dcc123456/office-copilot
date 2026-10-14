/**
 * 表格的 AI 代理（spec 决策 D6）：把"分析 / 整理 / 计算 / 出图"的诉求跑成一串工具回合，落地只认
 * `sheet_submit_proposal` 交出的那批改动与 `sheet_create_chart` 交出的那一张图表，并且**只交提案，不写表**（`AGENTS.md` §2.8）。
 *
 * 三条约束决定了这里的形状：
 * ① 取值、统计、算公式一律在 `@office-copilot/sheet-io` 里做，这里只搬运结构化结果（验收 B10）；
 * ② 上送给模型的内容默认截断——结构摘要 ≤50 列、样本 ≤50 行、单格文本 ≤120 字（§2.9、验收 B11）；
 * ③ 模型给的非法入参与引擎的失败理由都换成回给它的一句话，永不抛异常（验收 B9）。
 *
 * 工具回路放在渲染层：表格载荷只在这里，主进程只负责转发一轮请求（`ai:tool-turn`）。
 */
import {
  aggregateSheet,
  deleteTracksOp,
  describeSheet,
  duplicateGroups,
  findRows,
  formatA1,
  insertTracksOp,
  numberFormatColumnOps,
  numberFormatFromCode,
  readRange,
  resolveChartSpec,
  resolveColumnRef,
  setCellFormulaOp,
  setCellTextOp,
  setColumnWidthOp,
  sortRowsOp,
  type ResolvedChart,
  type SheetCoreOp,
  type SheetPayload,
} from "@office-copilot/sheet-io";
import {
  parseToolArguments,
  SHEET_TOOLS,
  type SheetEditT,
  type SheetToolName,
  type SubmitProposalArgsT,
} from "@office-copilot/sheet-tools";
import type { ToolCall, ToolTurnMessage, ToolTurnRequest, ToolTurnResult } from "../global.js";

/** 一次任务里最多转多少轮：模型无限要工具就该停下来问人，而不是无限烧 token。 */
const MAX_TOOL_TURNS = 8;

const SHEET_SYSTEM = `你是 Excel 表格代理，看得见的是当前工作表的结构摘要与前若干行样本，不是整张表。
- 要数字就先调工具：sheet_describe / sheet_read_range / sheet_aggregate / sheet_find / sheet_duplicates。不许凭样本心算整表结论。
- 要改表只能调 sheet_submit_proposal，条目只有 set_cell、set_column_width、insert_rows、delete_rows、sort_rows、set_number_format 六种；
  行号与列引用必须来自前面工具的回包，不许自己数行号。提案经用户确认才会落地。
- 要出图调 sheet_create_chart，只报区域（类别轴 + 每条系列的数值区域 + 落点那一格），数字一律由引擎读表得到；
  一张表只能挂一个图表，区域取不到数时引擎会回一句理由，改区域再调。经用户确认后才会随文件保存。
- 排序用 sort_rows：只交数据行区域（不含表头那一行）与键列，新的行序由引擎算，不许自己报次序。
  区域必须盖住一条记录的全部格子；带里有公式、或这些行在区域外还有内容时引擎会拒绝，照它给的理由改区域或直说做不了。
- 整列统一数字格式用 set_number_format：只交列与格式码（如 0.00%），逐格怎么改由引擎按现场判定，文本格与格式已对的格不会动。
  格式码不在我们认得的表里会被拒，照它给的理由换码。
- 拆合列这些这里做不了，遇到就直说做不了，别绕着改坏用户的表。
- 纯分析的问题直接回答，不要为了回答去提提案。回答用中文，带上你实际查到的数字。`;

/** `runSheetAgent` 的输入。 */
export interface SheetAgentInput {
  /** 当前这张表的结构化载荷（只读：代理不改它，落地由上层用 op 做）。 */
  sheet: SheetPayload;
  /** 用户那句诉求原文。 */
  instruction: string;
  /** 首行是否当表头，默认 true；上送摘要与列引用解析共用这一个口径。 */
  hasHeader?: boolean;
  /** 已经跑过的 transcript：暂停后继续时原样带回，模型不必重读上文。 */
  messages?: ToolTurnMessage[];
  /** 指定服务商；省略即用设置里激活的那个。 */
  providerId?: string;
  /** 发一轮请求的通道，默认走 preload 的 `ai:tool-turn`，单测里换成桩。 */
  ask?: (payload: ToolTurnRequest) => Promise<ToolTurnResult>;
  /** 用户按了暂停：下一轮请求发出前停下，transcript 原样带回。 */
  isPaused?: () => boolean;
}

/** 一次任务的结果。`messages` 每支都带，UI 靠它接着往下跑。 */
export type SheetAgentOutcome =
  | {
      kind: "proposal";
      summary: string;
      edits: SheetEditT[];
      /** 提案里能被现有 op 表达的部分；空数组意味着这份提案落不了地。 */
      ops: SheetCoreOp[];
      /** 没能变成 op 的条目原因，逐条给人看，不许静默丢掉（§2.5）。 */
      notes: string[];
      messages: ToolTurnMessage[];
    }
  | {
      /** 一张图表提案（spec 决策 D5）：数据已由引擎从现场表取好，落地只等用户确认。 */
      kind: "chart";
      chart: ResolvedChart;
      messages: ToolTurnMessage[];
    }
  | { kind: "answer"; text: string; messages: ToolTurnMessage[] }
  | { kind: "paused"; messages: ToolTurnMessage[] }
  | { kind: "failed"; message: string; needsSetup?: boolean; messages: ToolTurnMessage[] };

/**
 * 组装上送的工作表上下文：结构摘要 + 一小段样本，两处都带"被截断"的说明。
 *
 * @param sheet 当前表的载荷。
 * @param hasHeader 首行是否当表头。
 * @returns 多行中文文本，作为本轮用户消息的前缀（长度已由 `sheet-io` 的上限约束，不再二次裁剪）。
 */
export function sheetContextOf(sheet: SheetPayload, hasHeader: boolean): string {
  const profile = describeSheet(sheet, { hasHeader });
  const sample = profile.usedRange ? readRange(sheet, { range: profile.usedRange }) : null;
  const lines = [`工作表「${profile.sheetName}」：${profile.rowCount} 行 × ${profile.columnCount} 列${profile.usedRange ? `，已用区域 ${profile.usedRange}` : "（空表）"}。`];
  if (profile.headerWarnings.length) lines.push(`表头问题：${profile.headerWarnings.join("；")}`);
  if (profile.columnsTruncated) lines.push(`列太多，只描述前 ${profile.columns.length} 列，列数按 ${profile.columnCount} 算。`);
  lines.push(`列画像：${JSON.stringify(profile.columns)}`);
  if (sample?.ok) {
    lines.push(`样本（${sample.ref}，共 ${sample.rowCount} 行 × ${sample.colCount} 列${sample.truncated ? "，超出上限的部分没上送" : ""}）：${JSON.stringify(sample.rows)}`);
    if (sample.truncated) lines.push(`样本只到第 ${sample.rows.at(-1)?.rowNumber ?? 0} 行，后面的行要用 sheet_read_range 接着读。`);
  }
  return lines.join("\n");
}

/**
 * 跑一次工具调用：入参先过 `parseToolArguments`，再交给 `sheet-io` 算。
 *
 * @param sheet 当前表的载荷（只读）。
 * @param call 模型点名的工具与入参原文。
 * @returns 回给模型的文本；`proposal` 非空表示模型交了改动提案，`chart` 非空表示它出了图表，两者都让回路收尾。
 */
function executeToolCall(
  sheet: SheetPayload,
  call: ToolCall,
): { content: string; proposal?: SubmitProposalArgsT; chart?: ResolvedChart } {
  switch (call.name as SheetToolName) {
    case "sheet_describe": {
      const parsed = parseToolArguments("sheet_describe", call.arguments);
      return { content: parsed.ok ? JSON.stringify(describeSheet(sheet, parsed.args)) : parsed.reason };
    }
    case "sheet_read_range": {
      const parsed = parseToolArguments("sheet_read_range", call.arguments);
      return { content: parsed.ok ? JSON.stringify(readRange(sheet, parsed.args)) : parsed.reason };
    }
    case "sheet_aggregate": {
      const parsed = parseToolArguments("sheet_aggregate", call.arguments);
      return { content: parsed.ok ? JSON.stringify(aggregateSheet(sheet, parsed.args)) : parsed.reason };
    }
    case "sheet_find": {
      const parsed = parseToolArguments("sheet_find", call.arguments);
      return { content: parsed.ok ? JSON.stringify(findRows(sheet, parsed.args)) : parsed.reason };
    }
    case "sheet_duplicates": {
      const parsed = parseToolArguments("sheet_duplicates", call.arguments);
      return { content: parsed.ok ? JSON.stringify(duplicateGroups(sheet, parsed.args)) : parsed.reason };
    }
    case "sheet_create_chart": {
      const parsed = parseToolArguments("sheet_create_chart", call.arguments);
      if (!parsed.ok) return { content: parsed.reason };
      // 数值与引用串一律由引擎从现场表取；取不到就把理由回给模型，让它改引用再调（验收 B10）。
      const resolved = resolveChartSpec(sheet, parsed.args);
      if (!resolved.ok) return { content: `${resolved.reason}请改区域引用后再调一次 sheet_create_chart。` };
      const { chart } = resolved;
      const seriesLines = chart.series.map((item) => `${item.name}=${item.valueRef}`).join("，");
      return {
        content: `图表已解析：${chart.kind}｜类别 ${chart.categoryRef}（${chart.categories.length} 项）｜系列 ${seriesLines}｜落点 ${formatA1(chart.anchor)}。等用户确认后才会随文件保存。`,
        chart,
      };
    }
    case "sheet_submit_proposal": {
      const parsed = parseToolArguments("sheet_submit_proposal", call.arguments);
      if (!parsed.ok) return { content: parsed.reason };
      return {
        content: `提案已收下（${parsed.args.edits.length} 条改动），等用户确认后才会落地。要继续核对就接着调读工具，别重复提交。`,
        proposal: parsed.args,
      };
    }
    default:
      return { content: `不认识的工具：${call.name}。可用面只有 sheet_describe / sheet_read_range / sheet_aggregate / sheet_find / sheet_duplicates / sheet_create_chart / sheet_submit_proposal。` };
  }
}

/**
 * 把提案里的条目翻成自带逆操作的 op。
 *
 * 只翻这五类：它们各自有安全的撤销路径（§2.5）。翻不动的（列认不出、地址非法、公式造不出 op、
 * 排序被引擎的硬约束拒绝）逐条记进 `notes`，不许静默丢掉半份提案。
 *
 * @param sheet 提案要落在这张表上（按现场构造前像）。
 * @param edits `sheet_submit_proposal` 校验过的条目。
 * @param hasHeader 首行是否当表头，影响列引用的解析。
 * @returns `ops` 按提案顺序排好；`notes` 是没变成 op 的条目的中文说明。
 */
export function sheetOpsOfProposal(
  sheet: SheetPayload,
  edits: SheetEditT[],
  hasHeader = true,
): { ops: SheetCoreOp[]; notes: string[] } {
  const ops: SheetCoreOp[] = [];
  const notes: string[] = [];
  for (const edit of edits) {
    switch (edit.type) {
      case "set_cell": {
        const address = edit.address.trim().toUpperCase();
        // 与网格里手输同一条分派：= 开头当公式，公式自己算缓存值（spec 决策 D4）。
        const formula = edit.text.trim().startsWith("=") ? setCellFormulaOp(sheet, address, edit.text) : null;
        const op = formula?.op ?? setCellTextOp(sheet, address, edit.text);
        if (op) ops.push(op);
        else notes.push(`${address} 那一格写不了，已跳过。`);
        break;
      }
      case "set_column_width": {
        const colIndex = resolveColumnRef(sheet, edit.column, hasHeader);
        if (colIndex === null) {
          notes.push(`认不出列「${edit.column}」，那条列宽跳过。`);
          break;
        }
        // 宽度已经一样时 op 是 null：Excel 上改了也看不出，不算漏改。
        const op = setColumnWidthOp(sheet, colIndex, edit.widthPx);
        if (op) ops.push(op);
        break;
      }
      case "insert_rows": {
        const op = insertTracksOp(sheet, "row", edit.row - 1, edit.count);
        if (op) ops.push(op);
        else notes.push(`第 ${edit.row} 行插不了，已跳过。`);
        break;
      }
      case "delete_rows": {
        // 从大到小删：先删下面的行，上面那些行的行号还没被挪走，内容寻址更容易对上。
        for (const rowNumber of [...edit.rows].sort((left, right) => right - left)) {
          const op = deleteTracksOp(sheet, "row", rowNumber - 1);
          if (op) ops.push(op);
          else notes.push(`第 ${rowNumber} 行删不了，已跳过。`);
        }
        break;
      }
      case "sort_rows": {
        // 行序由引擎算（§2.11）：这里只把区域与键列交过去，被拒绝时把引擎那句理由原样报给上层。
        const made = sortRowsOp(sheet, { range: edit.range, column: edit.column, order: edit.order, hasHeader });
        if (made.ok) ops.push(made.op);
        else notes.push(`这条排序没落地：${made.reason}`);
        break;
      }
      case "set_number_format": {
        // 逐格的展开在引擎（§3.3、§2.11）：这里只把列与格式码交过去，认不出的码整条拒。
        const colIndex = resolveColumnRef(sheet, edit.column, hasHeader);
        if (colIndex === null) {
          notes.push(`认不出列「${edit.column}」，那条数字格式跳过。`);
          break;
        }
        const format = numberFormatFromCode(edit.format);
        if (!format) {
          notes.push(`格式码「${edit.format}」我们不认得，这条数字格式跳过（只认 General、0、0.00、#,##0、#,##0.00、0%、0.00%、m/d/yyyy、yyyy-mm-dd 这些写法）。`);
          break;
        }
        const made = numberFormatColumnOps(sheet, colIndex, format);
        if (!made.ops.length) notes.push(`${edit.column} 列没有需要改数字格式的格子（文本格与格式已对的格都不动）。`);
        ops.push(...made.ops);
        break;
      }
    }
  }
  return { ops, notes };
}

/**
 * 跑一次表格任务：转工具回合，直到模型交提案或图表、直接作答、被暂停或出错。
 *
 * @param input 见 `SheetAgentInput`；`sheet` 全程只读。
 * @returns `SheetAgentOutcome`。没配服务商、HTTP 失败、超过 8 轮都走 `failed`，不抛异常。
 */
export async function runSheetAgent(input: SheetAgentInput): Promise<SheetAgentOutcome> {
  const hasHeader = input.hasHeader ?? true;
  const ask = input.ask ?? ((payload: ToolTurnRequest) => window.officeCopilot.aiToolTurn(payload));
  const messages: ToolTurnMessage[] = input.messages?.length
    ? input.messages.map((message) => ({ ...message }))
    : [
        { role: "system", content: SHEET_SYSTEM },
        { role: "user", content: `${sheetContextOf(input.sheet, hasHeader)}\n\n诉求：${input.instruction}` },
      ];

  for (let turn = 0; turn < MAX_TOOL_TURNS; turn += 1) {
    if (input.isPaused?.()) return { kind: "paused", messages };
    const result = await ask({ messages, tools: SHEET_TOOLS, providerId: input.providerId });
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
      const executed = executeToolCall(input.sheet, call);
      messages.push({ role: "tool", tool_call_id: call.id, content: executed.content });
      if (executed.chart) return { kind: "chart", chart: executed.chart, messages };
      if (!executed.proposal) continue;
      const { ops, notes } = sheetOpsOfProposal(input.sheet, executed.proposal.edits, hasHeader);
      return { kind: "proposal", summary: executed.proposal.summary, edits: executed.proposal.edits, ops, notes, messages };
    }
  }
  return { kind: "failed", message: `工具回合到了 ${MAX_TOOL_TURNS} 轮上限，模型还没交提案或给出结论，先停下。`, messages };
}
