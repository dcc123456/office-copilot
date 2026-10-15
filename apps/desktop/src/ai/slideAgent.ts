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
  notesFlattenNote,
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
import { layerEdgeNote } from "../slides/slideDeckModel.js";

/** 一次任务里最多转多少轮：模型无限要工具就该停下来问人，而不是无限烧 token。 */
const MAX_TOOL_TURNS = 8;

/** 首条用户消息里那份画像的上限：给模型看得见全貌，但不许把整份 deck 摊上去（§2.9）。 */
const CONTEXT_MAX_SLIDES = 20;
const CONTEXT_MAX_SHAPES_PER_SLIDE = 20;

const SLIDE_SYSTEM = `你是演示文稿（.pptx）代理，看得见的是这份文稿的结构画像，不是整份原文。
- 动手前先调 ppt_describe 拿页号与 shapeId；要改某段文字再调 ppt_read_slide_text 确认 paraIndex / runIndex，不许凭印象数，也不许自己编形状号。
- 要改只能调 ppt_submit_proposal，条目只有 set_run_text / set_table_cell_text / set_run_font / set_run_color / set_shape_geometry / move_shape_layer / add_text_box / remove_shape / set_notes / create_notes 十种，
  前四条指到"某一页某个形状的某段某 run"（表格那条还要给行与列），中间两条指到形状（叠放方向是 up/down/top/bottom），后四条是结构或整篇改动：
  add_text_box 只给页号与四个几何值加文字（新形状号由引擎落地时分配，同一份提案里引用不到它），remove_shape 给页号与形状号，
  set_notes 给页号与逐段文字，把那一页**已存在**的演讲者备注整篇换掉（不是追加）；
  create_notes 同样给页号与逐段文字，用在画像里 notesPartPath 为 null 的那一页（**给它新建一份备注页**）——两档各管一半，别给已有备注页的页提 create，也别给没有备注页的页提 set。
  画像顶层的 hasNotesMaster=false 时这两档都提不了（这份文件没有备注母版，新建的备注页没有可继承的版式）。
  同一批里删掉的形状就别再改它的文字或字体（落地按顺序执行，那条必然失败并整批不写）。改版式与母版这里做不了，遇到就直说做不了。
- 表格住在图形框（kind=graphicFrame）里：那个形状自己没有文字体，ppt_read_slide_text 会改成逐格列出（row / col 都是 0 基，shapeId 用图形框那个）。
  改表格里的字就提 set_table_cell_text；改那一格某一截的字号/粗体/颜色就提 set_run_font / set_run_color 并**把 row 与 col 一起带上**（只给一个会被拒）；合并区里被并进来的那一格（列表里标了续格）不画自己的字，提了会被拒，要改就改合并起点那一格。
- set_shape_geometry 想只改一个轴（比如只把宽改一下），那个形状必须自己写过几何；画像里 hasOwnGeometry=false 的形状（标题/正文那类占位符，值从版式继承）要改就把四个值一起给。
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
 * 翻不动的（页号越界、该页没有这个形状、几何一条属性都没给、同一批里那个形状已被删掉、备注那一档提错了对象
 * （有备注页却提 create / 没有备注页却提 set）、这份文件没有备注母版、同一批里那一页已建过备注页、
 * 备注带多种格式、图形框里不是表、表格的行列/段/run 越界或落在合并续格）逐条记进 `notes`，
 * 不许交出一批主进程必然拒绝、整批不落地的改动（§2.4 的事务语义）。
 *
 * @param deck 提案要落在这份文稿上（页号→部件路径按它算）。
 * @param edits `ppt_submit_proposal` 校验过的条目。
 * @returns `edits` 按提案顺序排好；`notes` 是没变成改动的条目的中文说明。
 */
export function slideRequestsOfProposal(deck: PresentationModel, edits: PptxEditT[]): { requests: PptxEditRequest[]; notes: string[] } {
  const requests: PptxEditRequest[] = [];
  const notes: string[] = [];
  // 同一批里已经删掉的形状（`部件路径#形状号`）：按顺序落地时它们后面那条必然定位失败，整批就不写了（§2.4）。
  const removedInBatch = new Set<string>();
  // 同一批里已经给它建过备注页的那几页（按幻灯片部件路径记）：第二条会在引擎那儿撞上"已经有备注页了"，整批就不写了。
  const createdNotesInBatch = new Set<string>();
  for (const edit of edits) {
    const slide = deck.slides.find((entry) => entry.order + 1 === edit.slideNumber);
    if (!slide) {
      notes.push(`没有第 ${edit.slideNumber} 页（这份文稿共 ${deck.slides.length} 页），这条跳过。`);
      continue;
    }
    if (edit.type === "add_text_box") {
      // 加文本框只定位到整页：新形状的形状号是引擎落地时按该页现有最大值 +1 分配的，模型这一批里引用不到。
      requests.push({
        kind: "add_text_box",
        target: { partPath: slide.partPath },
        box: { leftPx: edit.leftPx, topPx: edit.topPx, widthPx: edit.widthPx, heightPx: edit.heightPx },
        text: edit.text,
      });
      continue;
    }
    if (edit.type === "set_notes") {
      // 备注住在**另一个部件**里，坐标是这一页读模型给的 `notes.partPath`（引擎坐标不上送给模型，与 add_text_box 同一条理由）。
      // 这一页没有备注部件、或现有备注带多种格式时，引擎那条必然拒绝并让整批不落地（§2.4），所以这里先判一刀再决定交不交。
      if (!slide.notes) {
        notes.push(`第 ${edit.slideNumber} 页没有备注页（这份文件里那一页不存在备注部件），这一档只能改已有的备注，这条跳过。`);
        continue;
      }
      const flattenNote = notesFlattenNote(slide.notes.paragraphRunCounts, slide.notes.styledRunCount);
      if (flattenNote) {
        notes.push(`第 ${edit.slideNumber} 页的备注改不了：${flattenNote}`);
        continue;
      }
      requests.push({ kind: "set_notes", target: { notesPartPath: slide.notes.partPath }, paragraphTexts: edit.paragraphTexts });
      continue;
    }
    if (edit.type === "create_notes") {
      // 新建与改已有各管一半，所以这里先判三刀，理由与引擎那三条一一对上（§2.4：交出一条引擎必拒的就是整批不落地）：
      // ① 这一页已经有备注部件——那条路是 set_notes 的，提 create 会造出第二份备注页；
      // ② 这份文件没有备注母版——新建的备注页没有可继承的版式，绝不自己造一份（§2.7）；
      // ③ 同一批里这一页已经建过一次——落地按顺序执行，第二条在引擎那儿必然看到"已经有备注页了"。
      if (slide.notes) {
        notes.push(`第 ${edit.slideNumber} 页已经有备注页了，那一份要改用 set_notes，这条新建跳过。`);
        continue;
      }
      if (!deck.hasNotesMaster) {
        notes.push(`这份文稿没有备注母版（notesMaster），新建的备注页没有可继承的版式，第 ${edit.slideNumber} 页这条跳过。`);
        continue;
      }
      if (createdNotesInBatch.has(slide.partPath)) {
        notes.push(`第 ${edit.slideNumber} 页的备注页在同一批里已经建过一次，这条是多余的，跳过。`);
        continue;
      }
      createdNotesInBatch.add(slide.partPath);
      requests.push({ kind: "create_notes", target: { partPath: slide.partPath }, paragraphTexts: edit.paragraphTexts });
      continue;
    }
    const shape = slide.shapes.find((entry) => entry.shapeId === edit.shapeId);
    if (!shape) {
      notes.push(`第 ${edit.slideNumber} 页没有 shapeId=${edit.shapeId} 这个形状，这条跳过。`);
      continue;
    }
    if (removedInBatch.has(`${slide.partPath}#${shape.shapeId}`)) {
      notes.push(`第 ${edit.slideNumber} 页「${shape.name}」在同一批里已经被删掉，这条改不了，跳过。`);
      continue;
    }
    if (edit.type === "remove_shape") {
      removedInBatch.add(`${slide.partPath}#${shape.shapeId}`);
      requests.push({ kind: "remove_shape", target: { partPath: slide.partPath, shapeId: shape.shapeId } });
      continue;
    }
    if (edit.type === "set_shape_geometry") {
      // 组合与组内形状的几何引擎会拒（子偏移语义不同）：这里就别交，否则主进程整批不落地（§2.4）。
      const blockedNote = geometryBlockedNote(shape);
      if (blockedNote) {
        notes.push(`第 ${edit.slideNumber} 页的「${shape.name}」几何改不了（${blockedNote}），这条跳过。`);
        continue;
      }
      // 占位符那一类自己没写 `<a:xfrm>`：`<a:off>` 的 x、y 与 `<a:ext>` 的 cx、cy 在 OOXML 里都是必填，
      // 只给一轴时引擎写出来的文件会被 PowerPoint 判损坏，所以这条同样别交（理由见 pptx-io 的 missingPairValues）。
      if (!shape.boxPx && !hasFullBox(edit)) {
        notes.push(
          `第 ${edit.slideNumber} 页「${shape.name}」的位置尺寸是从版式继承的（它自己没写），要改就得把 leftPx、topPx、widthPx、heightPx 四个一起给，这条跳过。`,
        );
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
    if (edit.type === "set_table_cell_text") {
      // 表格那一格住在图形框自己的 `<a:tbl>` 里，不在形状的 `p:txBody`，所以坐标要多带行列；
      // 引擎那条会把**整批**拒掉（§2.4），因此四层（表、行列、段、run）都按现场模型先判一刀再决定交不交。
      const cellNote = tableCellRangeNote(shape, edit.row, edit.col, edit.paraIndex, edit.runIndex);
      if (cellNote) {
        notes.push(`第 ${edit.slideNumber} 页「${shape.name}」那一格改不了：${cellNote}`);
        continue;
      }
      requests.push({
        kind: "table_text",
        target: { partPath: slide.partPath, shapeId: shape.shapeId, row: edit.row, col: edit.col, paraIndex: edit.paraIndex, runIndex: edit.runIndex },
        text: edit.text,
      });
      continue;
    }
    if ((edit.type === "set_run_font" || edit.type === "set_run_color") && edit.row !== undefined && edit.col !== undefined) {
      // 带行列就是改表格里那一格：引擎那两条 op 认得这个坐标（`RunTarget`），四层预拒与 set_table_cell_text 共用同一份判定（§3.2）。
      const cellNote = tableCellRangeNote(shape, edit.row, edit.col, edit.paraIndex, edit.runIndex);
      if (cellNote) {
        notes.push(`第 ${edit.slideNumber} 页「${shape.name}」那一格改不了：${cellNote}`);
        continue;
      }
      const cellTarget = { partPath: slide.partPath, shapeId: shape.shapeId, row: edit.row, col: edit.col, paraIndex: edit.paraIndex, runIndex: edit.runIndex };
      if (edit.type === "set_run_color") requests.push({ kind: "color", target: cellTarget, hex: edit.color });
      else requests.push({ kind: "font", target: cellTarget, format: { bold: edit.bold, sizePt: edit.sizePt } });
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
 * 这个形状的几何是不是压根改不了（组合本身，或装在组合里的那一格——引擎两条都会整条拒）。
 *
 * @param shape 形状快照。
 * @returns 一句中文理由；能改几何时返回 null。
 */
function geometryBlockedNote(shape: SlideShape): string | null {
  if (shape.kind === "group") return "它是组合形状，子偏移语义不同";
  if (shape.groupShapeId !== undefined) return "它装在某个组合里，坐标写的是组内子坐标系，换算回页面位置那一档还没做";
  return null;
}

/**
 * 这条几何提案是不是把四个值都给齐了（占位符那种只能整对给，见 `slideRequestsOfProposal` 里的用法）。
 *
 * @param edit 已判定为 `set_shape_geometry` 的那条提案。
 * @returns true 表示 `leftPx`、`topPx`、`widthPx`、`heightPx` 全部给了数字。
 */
function hasFullBox(edit: Extract<PptxEditT, { type: "set_shape_geometry" }>): boolean {
  return edit.leftPx !== undefined && edit.topPx !== undefined && edit.widthPx !== undefined && edit.heightPx !== undefined;
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
  if (shape.paragraphTexts.length === 0) return "这个形状没有文字体（图片、图形框或连接线），改不了文字";
  if (paraIndex >= shape.paragraphTexts.length) return `段落序号 ${paraIndex} 越界（只有 ${shape.paragraphTexts.length} 段）`;
  const runCount = shape.paragraphRunCounts[paraIndex] ?? 0;
  if (runCount === 0) return `第 ${paraIndex} 段的文字来自字段（页码 / 日期），不在可改范围`;
  if (runIndex >= runCount) return `第 ${paraIndex} 段只有 ${runCount} 个 run，没有第 ${runIndex} 个`;
  return null;
}

/**
 * 表格那一格的坐标对不对得上现场模型（表、行列、段、run 四层逐层查）。
 *
 * 引擎的表格那三条 op（改字、改字体、改颜色）一旦拒就是整批不落地（§2.4），所以这里拿读模型先判。
 * 合并续格单独点名：它在 XML 里确实存在、也确实有一段空 `<a:p/>`，但 PowerPoint 不画它的字，写了等于静默没写。
 *
 * @param shape 该形状快照（表格来自它的 `table`，没有就是图形框里装的不是表）。
 * @param row 模型给的行号（0 基）。
 * @param col 模型给的列号（0 基，合并起点往后照数）。
 * @param paraIndex 模型给的段号（0 基，该格自己的段）。
 * @param runIndex 模型给的 run 号（0 基，该段自己的 run）。
 * @returns 一句中文理由；四层都对得上时返回 null。
 */
function tableCellRangeNote(shape: SlideShape, row: number, col: number, paraIndex: number, runIndex: number): string | null {
  const table = shape.table;
  if (!table) return "那个形状里没有表格（图形框里装的可能是图表，或者它压根是图片），那一格什么都改不了";
  if (row >= table.rows.length) return `行号 ${row} 越界（这张表有 ${table.rows.length} 行）`;
  const cells = table.rows[row] ?? [];
  if (col >= cells.length) return `列号 ${col} 越界（第 ${row} 行有 ${cells.length} 格）`;
  const cell = cells[col]!;
  if (cell.isMergeContinuation) return `第 ${row} 行第 ${col} 列是合并区里被并进来的那一格，它不画自己的字，要改的是合并起点那一格`;
  if (paraIndex >= cell.paragraphRunCounts.length) return `第 ${row} 行第 ${col} 列只有 ${cell.paragraphRunCounts.length} 段，没有第 ${paraIndex} 段`;
  const runCount = cell.paragraphRunCounts[paraIndex] ?? 0;
  if (runCount === 0) return `第 ${row} 行第 ${col} 列的第 ${paraIndex} 段没有 run（空段或字段），那一截什么都改不了`;
  if (runIndex >= runCount) return `第 ${row} 行第 ${col} 列的第 ${paraIndex} 段只有 ${runCount} 个 run，没有第 ${runIndex} 个`;
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
