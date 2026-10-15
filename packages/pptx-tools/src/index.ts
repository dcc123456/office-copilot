/**
 * 幻灯片的 AI 工具面（spec 决策 D6，pptx 一线）：给模型看的 JSON 工具定义 + zod 入参校验。
 *
 * 与 `sheet-tools`/`word-tools` 同构，守同一条红线：模型只产出**入参与提案**，定位、落地、撤销全部由
 * `@office-copilot/pptx-io` 做（`AGENTS.md` §2.8、§2.11）。写面只有 `ppt_submit_proposal` 一个出口，
 * 它的九类改动（改 run 文字 / 改 run 字体 / 改 run 颜色 / 改表格某格的一条 run / 移动缩放形状 / 调整叠放次序 / 加文本框 / 删形状 / 整篇换备注）与
 * `pptx-io/edit.ts` 里**自带逆操作**的九个 op 一一对应——表达不了的（改版式与母版、给没备注页的那页造备注）宁可不提供，
 * 也不发一批没有逆操作的直写（§2.5）。
 * 按 spec §三 的包边界，本包只依赖 `zod` 与 `@office-copilot/shared`（那条入参判定共用一份）、**不 import `pptx-io`**——它只管"这套入参合不合法"，
 * 执行由调用方（桌面 slideAgent）拿校验过的入参去调引擎函数：把 `slideNumber` 翻成部件路径、
 * 把 `{shapeId,paraIndex,runIndex}` 交给 `setShapeRunText`/`setRunFont`/`setRunColor`、
 * 把表格那格的 `{shapeId,row,col,paraIndex,runIndex}` 交给 `setTableCellRunText`、
 * 把 `{shapeId}` 交给 `setShapeGeometry`/`moveShapeLayer`/`removeShape`、把整页路径交给 `addTextBox`。
 * 坐标一律是"模型视角"：幻灯片 1 基（ppt_describe 报第几页）、形状号取自 ppt_describe 给的 shapeId、
 * 段/run 与表格的行/列 0 基（与 ppt_read_slide_text 列出的序号一致）。单位说在描述里：字号 pt、几何 px（§2.1 换算在引擎边）。
 * 非法入参一律换成给模型看的错误串，永不抛异常（验收 B9）：入口只有 `parseToolArguments`。
 */
import { z } from "zod";
import { parseToolArguments as parseToolArgumentsAgainst, type ToolArgsOutcome } from "@office-copilot/shared";

/** 幻灯片页号，1 基（模型数页从 1 开始，和界面上"第 N 页"一致）。 */
const SlideNumber = z.number().int().min(1).max(500).describe("幻灯片页号，1 基（第 1 页就是 1）。来自 ppt_describe 列出的页号，别自己猜。");

/** 形状号 `<p:cNvPr @id>`，取自 ppt_describe 给每个形状报的 shapeId。 */
const ShapeId = z.number().int().min(1).max(1_000_000).describe("形状号（shapeId），来自 ppt_describe 里该页每个形状标的那个 id，别自己编。");

/** 段落序号，0 基（该形状文本体里的第几段）。 */
const ParaIndex = z.number().int().min(0).max(2000).describe("段落序号，0 基（该形状里第几段）。来自 ppt_read_slide_text 列出的段号。");

/** run 序号，0 基（该段里的第几个文字片段）。 */
const RunIndex = z.number().int().min(0).max(2000).describe("run 序号，0 基（该段里第几个文字片段）。一段可能被拆成多段不同格式的 run。");

/** 字体颜色：`#rrggbb`（编辑器内一律 hex，§2.2），允许省略前导 `#`，归一化由引擎做。 */
const RunColor = z
  .string()
  .regex(/^#?[0-9a-fA-F]{6}$/)
  .describe("文字颜色，六位十六进制 #rrggbb（如 #C00000）。只认这个形式，别的颜色写法（red、rgb()、8 位带透明度）会被拒。");

/**
 * `ppt_describe` 的入参：读整份演示文稿的页/形状画像，让模型先知道有哪些页、每页有哪些可点形状。
 *
 * 数值与形状清单一律由引擎从现场 `.pptx` 读（§2.11），这里只约束翻页/翻形状的上限。
 */
export const DescribeArgs = z
  .object({
    maxSlides: z.number().int().min(1).max(200).default(50).describe("最多描述多少页，默认 50；剩下的页只报总数。"),
    maxShapesPerSlide: z.number().int().min(1).max(200).default(40).describe("每页最多列多少个形状，默认 40；超出的只报形状总数。"),
  })
  .strict();

/** `ppt_read_slide_text` 的入参：把某一页每个文本形状的段/run 摊开，好让模型填对 paraIndex/runIndex。 */
export const ReadSlideTextArgs = z
  .object({
    slideNumber: SlideNumber,
    maxRunsPerParagraph: z.number().int().min(1).max(200).default(50).describe("每段最多列多少个 run，默认 50。"),
  })
  .strict();

/**
 * 表格那一格的行号（`<a:tr>` 位置，0 基，与读模型 `SlideShape.table.rows` 同口径）。
 * 边界只写这一份：`set_table_cell_text` 必填它，`set_run_font`/`set_run_color` 可给它。
 */
const TableRow = z.number().int().min(0).max(2000);

/** 表格那一格的列号（该行内 `<a:tc>` 位置，0 基；合并起点也算第 0 列往后数）。 */
const TableCol = z.number().int().min(0).max(500);

/** 一条改 run 文字的请求。 */
const SetRunTextEdit = z
  .object({
    type: z.literal("set_run_text"),
    slideNumber: SlideNumber,
    shapeId: ShapeId,
    paraIndex: ParaIndex,
    runIndex: RunIndex,
    text: z.string().max(2000).describe("这个 run 要换成的新文字。只动这一个 run，同段其它 run 的格式与文字不受影响。"),
  })
  .strict();

/**
 * 一条改**幻灯片表格里某一格**的请求。
 *
 * 表格挂在图形框（`p:graphicFrame`）里，那一格的文字体不是形状自己的 `p:txBody`，所以定位要多给行与列；
 * 两者都是 0 基，取自 ppt_read_slide_text 列出的行列号。合并区里被并进来的那一格（续格）不画自己的字，提了会被引擎拒。
 */
const SetTableCellTextEdit = z
  .object({
    type: z.literal("set_table_cell_text"),
    slideNumber: SlideNumber,
    shapeId: ShapeId,
    row: TableRow.describe("表格行号，0 基。来自 ppt_read_slide_text 里这张表列出的行号，别自己数。"),
    col: TableCol.describe("表格列号，0 基（合并起点也算第 0 列往后数）。越界或落在合并续格上都会被拒。"),
    paraIndex: ParaIndex,
    runIndex: RunIndex,
    text: z.string().max(2000).describe("这一格该 run 要换成的新文字。只动这一个 run，同格同段其它 run 的格式与文字不受影响。"),
  })
  .strict();

/** 一条改 run 字体（粗体 / 字号）的请求；bold 与 sizePt 至少给一个由 `SubmitProposalArgs` 顶层校验。 */
const SetRunFontEdit = z
  .object({
    type: z.literal("set_run_font"),
    slideNumber: SlideNumber,
    shapeId: ShapeId,
    row: TableRow.optional().describe("给了行号（连同列号）就是改**幻灯片表格里那一格**的那条 run；省略就是改普通形状自己的文字体。"),
    col: TableCol.optional().describe("给了列号就是改表格里那一格；与行号一起给，0 基。落在合并续格上会被拒。"),
    paraIndex: ParaIndex,
    runIndex: RunIndex,
    bold: z.boolean().optional().describe("是否加粗：true 加粗、false 取消加粗。省略表示不改这一项。"),
    sizePt: z.number().min(1).max(400).optional().describe("字号，单位磅（pt，不是半磅也不是 px）；换算成 OOXML 的 sz 由引擎做。省略表示不改字号。"),
  })
  .strict();

/** 一条改 run 颜色的请求。 */
const SetRunColorEdit = z
  .object({
    type: z.literal("set_run_color"),
    slideNumber: SlideNumber,
    shapeId: ShapeId,
    row: TableRow.optional().describe("给了行号（连同列号）就是改**幻灯片表格里那一格**的那条 run；省略就是改普通形状自己的文字体。"),
    col: TableCol.optional().describe("给了列号就是改表格里那一格；与行号一起给，0 基。落在合并续格上会被拒。"),
    paraIndex: ParaIndex,
    runIndex: RunIndex,
    color: RunColor,
  })
  .strict();

/**
 * 一条移动 / 缩放形状的请求；四个方位至少给一个由 `SubmitProposalArgs` 顶层校验。几何单位是屏幕 px（不是 EMU、不是 pt），换算由引擎做（§2.1）。
 * 自己没写 `<a:xfrm>` 的形状（占位符那类）必须四个一起给，因为 `<a:off>` 的 x、y 与 `<a:ext>` 的 cx、cy 在 OOXML 里都是必填（校验在 `pptx-io` 的 `setShapeGeometry`）。
 */
const SetShapeGeometryEdit = z
  .object({
    type: z.literal("set_shape_geometry"),
    slideNumber: SlideNumber,
    shapeId: ShapeId,
    leftPx: z.number().min(-5000).max(20000).optional().describe("形状左缘横坐标，单位屏幕像素（px）；可为负表示挪出画面。省略表示不改横坐标。"),
    topPx: z.number().min(-5000).max(20000).optional().describe("形状上缘纵坐标，单位 px。省略表示不改纵坐标。"),
    widthPx: z.number().min(1).max(20000).optional().describe("形状宽度，单位 px，须为正。省略表示不改宽度。"),
    heightPx: z.number().min(1).max(20000).optional().describe("形状高度，单位 px，须为正。省略表示不改高度。"),
  })
  .strict();

/** 一条沿叠放次序移动形状的请求；`up`/`down` 相邻挪一层，`top`/`bottom` 到两端。 */
const MoveShapeLayerEdit = z
  .object({
    type: z.literal("move_shape_layer"),
    slideNumber: SlideNumber,
    shapeId: ShapeId,
    direction: z
      .enum(["up", "down", "top", "bottom"])
      .describe("叠放方向：up 上移一层、down 下移一层、top 置于顶层、bottom 置于底层。该形状已在指定端时这条会被拒。"),
  })
  .strict();

/** 一条加文本框的请求：四个几何值都要给（新形状没有可"不改"的原值），单位同样是屏幕 px。 */
const AddTextBoxEdit = z
  .object({
    type: z.literal("add_text_box"),
    slideNumber: SlideNumber,
    leftPx: z.number().min(-5000).max(20000).describe("新文本框左缘横坐标，单位屏幕像素（px）。"),
    topPx: z.number().min(-5000).max(20000).describe("新文本框上缘纵坐标，单位 px。"),
    widthPx: z.number().min(1).max(20000).describe("新文本框宽度，单位 px，须为正。"),
    heightPx: z.number().min(1).max(20000).describe("新文本框高度，单位 px，须为正。"),
    text: z
      .string()
      .min(1)
      .max(2000)
      .describe(
        "文本框里的文字，必填。落地后它落在该页最上层，形状号由引擎当场分配，所以**这一份提案里不能引用它**（要再改字号颜色得等下一轮重新 describe）。",
      ),
  })
  .strict();

/** 一条删形状的请求：整个形状连同它的文字一起没，逆操作是这一页的完整前像（§2.5）。 */
const RemoveShapeEdit = z
  .object({
    type: z.literal("remove_shape"),
    slideNumber: SlideNumber,
    shapeId: ShapeId,
  })
  .strict();

/**
 * 一条改某页演讲者备注的请求：**整篇换掉**这一页已经存在的备注，不是追加。
 *
 * 模型只交页号，`notesSlideN.xml` 那个路径由装配层按现场模型算（引擎坐标不上送，与 `add_text_box` 同一条理由）。
 */
const SetNotesEdit = z
  .object({
    type: z.literal("set_notes"),
    slideNumber: SlideNumber,
    paragraphTexts: z
      .array(z.string().max(2000))
      .min(1)
      .max(200)
      .describe(
        "备注的逐段文字，一段一个字符串（空串表示那一行没字），1~200 段。这是**整篇替换**不是追加；" +
          "这一页本来没有备注页时这条会被拒（新建备注部件现在做不了），画像里 notesPartPath 为 null 的页就别提。",
      ),
  })
  .strict();

/** 提案里的一条改动。只给有安全逆操作的这几类（AGENTS.md §2.5），结构类到增删形状与整篇换备注为止。 */
export const PptxEdit = z.discriminatedUnion("type", [
  SetRunTextEdit,
  SetTableCellTextEdit,
  SetRunFontEdit,
  SetRunColorEdit,
  SetShapeGeometryEdit,
  MoveShapeLayerEdit,
  AddTextBoxEdit,
  RemoveShapeEdit,
  SetNotesEdit,
]);

/**
 * `ppt_submit_proposal` 的入参。
 *
 * **不在这里**的是改版式/母版与删整页（还没建成带逆操作的 op），模型只能提这九项。
 * "字体/几何至少给一项"这种跨字段约束不能放进 `discriminatedUnion`（会把成员包成 ZodEffects 而非法），所以在外层校验。
 */
export const SubmitProposalArgs = z
  .object({
    summary: z.string().min(1).max(200).describe("一句话中文说明这份提案要做什么。"),
    edits: z.array(PptxEdit).min(1).max(200).describe("按顺序落地的改动，1~200 条。"),
  })
  .strict()
  .superRefine((val, ctx) => {
    val.edits.forEach((edit, index) => {
      if (edit.type === "set_run_font" && edit.bold === undefined && edit.sizePt === undefined) {
        ctx.addIssue({ code: z.ZodIssueCode.custom, path: ["edits", index], message: "set_run_font 的 bold 与 sizePt 至少给一个，否则这条什么都没改。" });
      }
      if ((edit.type === "set_run_font" || edit.type === "set_run_color") && (edit.row === undefined) !== (edit.col === undefined)) {
        // 只给一个下标时这一条到底指普通形状还是表格里那一格都判不出来；引擎按行与列两个下标才走表格那条链。
        ctx.addIssue({ code: z.ZodIssueCode.custom, path: ["edits", index], message: "set_run_font / set_run_color 要改表格里那一格时，row 与 col 必须一起给；改普通形状的 run 就两个都省略。" });
      }
      if (
        edit.type === "set_shape_geometry" &&
        edit.leftPx === undefined &&
        edit.topPx === undefined &&
        edit.widthPx === undefined &&
        edit.heightPx === undefined
      ) {
        ctx.addIssue({ code: z.ZodIssueCode.custom, path: ["edits", index], message: "set_shape_geometry 的 leftPx / topPx / widthPx / heightPx 至少给一个。" });
      }
    });
  });

/** 工具名 → 入参 schema；`parseToolArguments` 按这个名字分派。 */
export const PPTX_TOOL_SCHEMAS = {
  ppt_describe: DescribeArgs,
  ppt_read_slide_text: ReadSlideTextArgs,
  ppt_submit_proposal: SubmitProposalArgs,
} as const;

/** 工具名（模型回调时能出现的合法集合，用来拦住不存在的方法名）。 */
export type PptxToolName = keyof typeof PPTX_TOOL_SCHEMAS;

export type DescribeArgsT = z.infer<typeof DescribeArgs>;
export type ReadSlideTextArgsT = z.infer<typeof ReadSlideTextArgs>;
export type SubmitProposalArgsT = z.infer<typeof SubmitProposalArgs>;
export type PptxEditT = z.infer<typeof PptxEdit>;

/**
 * 解析模型给的入参：JSON 语法错、缺字段、多字段、越界都换成给模型看的错误串，**永不抛异常**（验收 B9）。
 *
 * 判定本身在 `@office-copilot/shared`——几个工具面用的是同一条，抽出来只留一份（`AGENTS.md` §3.2）；
 * 这里只把本包那张 schema 表接上，让调用方拿到按工具名收窄过的入参类型。
 *
 * @param name 工具名；不在 `PPTX_TOOL_SCHEMAS` 里时给"不认识的工具"。
 * @param rawArguments 模型回包里的 `function.arguments` 原文（JSON 字符串）。
 * @returns 校验过的入参，或给模型看的失败理由。
 */
export function parseToolArguments<T extends PptxToolName>(name: T, rawArguments: string): ToolArgsOutcome<z.infer<(typeof PPTX_TOOL_SCHEMAS)[T]>> {
  return parseToolArgumentsAgainst(PPTX_TOOL_SCHEMAS, name, rawArguments);
}

/** 一条工具定义（OpenAI function-calling 形状，与 `SHEET_TOOLS`/`WORD_TOOLS` 同形）。 */
export interface PptxToolDefinition {
  type: "function";
  function: {
    name: PptxToolName;
    description: string;
    /** JSON Schema 的极小子集：我们只用 `properties` / `required` / `additionalProperties` 这三样。 */
    parameters: { type: "object"; properties?: Record<string, unknown>; required?: string[]; additionalProperties: false };
  };
}

/** 一个"目标 run"的公共参数（除几何外三类 run 编辑都用它），内联进 JSON Schema 时复用。 */
const RUN_TARGET_JSON = {
  slideNumber: { type: "integer", minimum: 1, maximum: 500 },
  shapeId: { type: "integer", minimum: 1, maximum: 1_000_000 },
  paraIndex: { type: "integer", minimum: 0, maximum: 2000 },
  runIndex: { type: "integer", minimum: 0, maximum: 2000 },
};

/** 表格那一格行列下标的 JSON 形态：边界与 `TableRow`/`TableCol` 那两条 zod 一致，只写一份。 */
const TABLE_CELL_POS_JSON = {
  row: { type: "integer", minimum: 0, maximum: 2000 },
  col: { type: "integer", minimum: 0, maximum: 500 },
};

/** 给模型的读工具与写工具定义，直接进请求的 `tools` 数组。 */
export const PPTX_TOOLS: PptxToolDefinition[] = [
  {
    type: "function",
    function: {
      name: "ppt_describe",
      description:
        "读出当前演示文稿的结构画像：一共几页、幻灯片尺寸、每页有哪些形状（形状名、shapeId、是占位符还是文本框/图片、是否自带几何、每页的文字概要）。" +
        "要做任何编辑之前先调用它拿到 slideNumber 与 shapeId。形状/页的数量由引擎从现场文件读，你不许自己编。",
      parameters: {
        type: "object",
        properties: {
          maxSlides: { type: "integer", minimum: 1, maximum: 200, default: 50 },
          maxShapesPerSlide: { type: "integer", minimum: 1, maximum: 200, default: 40 },
        },
        additionalProperties: false,
      },
    },
  },
  {
    type: "function",
    function: {
      name: "ppt_read_slide_text",
      description:
        "把某一页里每个文本形状的段落与 run 摊开列出（含 paraIndex、runIndex、该 run 现文本、现有的粗体/字号/颜色）。" +
        "图形框里装的若是表格，会按行列逐格列出（含 row、col、该格现文本与逐段 run 数）；合并区里被并进来的那一格会标出来，它不画自己的字。" +
        "提交 run 级编辑前用它确认要改的 paraIndex/runIndex 与行列号，别凭印象数。",
      parameters: {
        type: "object",
        properties: {
          slideNumber: { type: "integer", minimum: 1, maximum: 500 },
          maxRunsPerParagraph: { type: "integer", minimum: 1, maximum: 200, default: 50 },
        },
        required: ["slideNumber"],
        additionalProperties: false,
      },
    },
  },
  {
    type: "function",
    function: {
      name: "ppt_submit_proposal",
      description:
        "把你打算对幻灯片做的修改打包成提案返回。用户会看到逐条改动并确认后才落地，一次回答只提交一份提案。" +
        "九类改动：set_run_text 改某个 run 的文字（只动这一个 run，不拍平同段其它格式）；" +
        "set_table_cell_text 改表格里某一格的某个 run（除 shapeId 还要给 row/col，都是 0 基，取自 ppt_read_slide_text 列出的行列；" +
        "shapeId 给的是那个**图形框**自己的 id，不是格子的 id；落在合并续格上会被拒，那一格不画自己的字，要改就改合并起点那一格）；" +
        "set_run_font 改某个 run 的粗体/字号（bold 与 sizePt 至少给一个，字号单位 pt）；" +
        "set_run_color 改某个 run 的颜色（#rrggbb）；这两条一起给 row/col 时改的就是**幻灯片表格里那一格**的那条 run（行列都是 0 基，取自 ppt_read_slide_text；续格一样会被拒），省略就是改普通形状自己的文字体；" +
        "set_shape_geometry 移动/缩放某个形状（leftPx/topPx/widthPx/heightPx 至少一个，单位 px；" +
        "画像是占位符/母版形状那种（没有自己的几何、值从版式继承）时四个必须一起给，只给一轴会被拒；" +
        "组合形状自己、以及画像里 inGroup=true 的那一格（坐标写在组内子坐标系）几何都改不了，提了只会换回一句理由）；" +
        "move_shape_layer 改某个形状的叠放次序（up/down 相邻挪一层，top/bottom 到两端）；" +
        "add_text_box 在某页新加一个文本框（四个几何值都要给，文字必填，形状号由引擎在落地时分配，这一份提案里引用不到它）；" +
        "remove_shape 删掉某个形状（连同它的文字一起没，可撤销）。" +
        "set_notes 把某一页的演讲者备注整篇换成你给的几段（一段一个字符串，不是追加；" +
        "只对**本来就有备注页**的页有效，画像里 notesPartPath 为 null 的页提了会被拒，现有备注带多种格式时也会被拒，因为整篇覆盖会拍平那些格式）。" +
        "改版式与母版、删整页、给没有备注页的那一页新建备注现在都提不了。",
      parameters: {
        type: "object",
        properties: {
          summary: { type: "string", minLength: 1, maxLength: 200 },
          edits: {
            type: "array",
            minItems: 1,
            maxItems: 200,
            items: {
              type: "object",
              oneOf: [
                {
                  type: "object",
                  properties: { type: { const: "set_run_text" }, ...RUN_TARGET_JSON, text: { type: "string", maxLength: 2000 } },
                  required: ["type", "slideNumber", "shapeId", "paraIndex", "runIndex", "text"],
                  additionalProperties: false,
                },
                {
                  type: "object",
                  properties: {
                    type: { const: "set_table_cell_text" },
                    ...RUN_TARGET_JSON,
                    ...TABLE_CELL_POS_JSON,
                    text: { type: "string", maxLength: 2000 },
                  },
                  required: ["type", "slideNumber", "shapeId", "row", "col", "paraIndex", "runIndex", "text"],
                  additionalProperties: false,
                },
                {
                  type: "object",
                  properties: {
                    type: { const: "set_run_font" },
                    ...RUN_TARGET_JSON,
                    ...TABLE_CELL_POS_JSON,
                    bold: { type: "boolean" },
                    sizePt: { type: "number", minimum: 1, maximum: 400 },
                  },
                  required: ["type", "slideNumber", "shapeId", "paraIndex", "runIndex"],
                  additionalProperties: false,
                },
                {
                  type: "object",
                  properties: {
                    type: { const: "set_run_color" },
                    ...RUN_TARGET_JSON,
                    ...TABLE_CELL_POS_JSON,
                    color: { type: "string", pattern: "^#?[0-9a-fA-F]{6}$" },
                  },
                  required: ["type", "slideNumber", "shapeId", "paraIndex", "runIndex", "color"],
                  additionalProperties: false,
                },
                {
                  type: "object",
                  properties: {
                    type: { const: "set_shape_geometry" },
                    slideNumber: { type: "integer", minimum: 1, maximum: 500 },
                    shapeId: { type: "integer", minimum: 1, maximum: 1_000_000 },
                    leftPx: { type: "number", minimum: -5000, maximum: 20000 },
                    topPx: { type: "number", minimum: -5000, maximum: 20000 },
                    widthPx: { type: "number", minimum: 1, maximum: 20000 },
                    heightPx: { type: "number", minimum: 1, maximum: 20000 },
                  },
                  required: ["type", "slideNumber", "shapeId"],
                  additionalProperties: false,
                },
                {
                  type: "object",
                  properties: {
                    type: { const: "move_shape_layer" },
                    slideNumber: { type: "integer", minimum: 1, maximum: 500 },
                    shapeId: { type: "integer", minimum: 1, maximum: 1_000_000 },
                    direction: { enum: ["up", "down", "top", "bottom"] },
                  },
                  required: ["type", "slideNumber", "shapeId", "direction"],
                  additionalProperties: false,
                },
                {
                  type: "object",
                  properties: {
                    type: { const: "add_text_box" },
                    slideNumber: { type: "integer", minimum: 1, maximum: 500 },
                    leftPx: { type: "number", minimum: -5000, maximum: 20000 },
                    topPx: { type: "number", minimum: -5000, maximum: 20000 },
                    widthPx: { type: "number", minimum: 1, maximum: 20000 },
                    heightPx: { type: "number", minimum: 1, maximum: 20000 },
                    text: { type: "string", minLength: 1, maxLength: 2000 },
                  },
                  required: ["type", "slideNumber", "leftPx", "topPx", "widthPx", "heightPx", "text"],
                  additionalProperties: false,
                },
                {
                  type: "object",
                  properties: {
                    type: { const: "remove_shape" },
                    slideNumber: { type: "integer", minimum: 1, maximum: 500 },
                    shapeId: { type: "integer", minimum: 1, maximum: 1_000_000 },
                  },
                  required: ["type", "slideNumber", "shapeId"],
                  additionalProperties: false,
                },
                {
                  type: "object",
                  properties: {
                    type: { const: "set_notes" },
                    slideNumber: { type: "integer", minimum: 1, maximum: 500 },
                    paragraphTexts: {
                      type: "array",
                      minItems: 1,
                      maxItems: 200,
                      items: { type: "string", maxLength: 2000 },
                    },
                  },
                  required: ["type", "slideNumber", "paragraphTexts"],
                  additionalProperties: false,
                },
              ],
            },
          },
        },
        required: ["summary", "edits"],
        additionalProperties: false,
      },
    },
  },
];
