/**
 * 表格的 AI 工具面（spec 决策 D6）：给模型看的 JSON 工具定义 + zod 入参校验。
 *
 * 与 `word-tools` 同构，守同一条红线：模型只产出**入参与提案**，取值、统计、落地全部由
 * `@office-copilot/sheet-io` 做（`AGENTS.md` §2.8、spec §零.4，验收 B10）。
 * 写面只有 `sheet_submit_proposal` 一个出口，它的五类改动（改格 / 列宽 / 插行 / 删行 / 按列重排整行）
 * 与 `sheet-edit.ts` 里**自带逆操作**的五类 op 一一对应——表达不了的（拆合列、统一数字格式）
 * 宁可不提供，也不发一批没有逆操作的直写（§2.5，spec §七"整理写侧缺 op"；排序按决策 D7 立项后才有 op）。
 * 按 spec §三 的包边界，本包只依赖 `zod`、**不 import `sheet-io`**——它只管"这套入参合不合法"，
 * 执行由调用方（桌面 `sheetAgent`，提交 13）拿校验过的入参去调引擎函数。
 * 非法入参一律换成给模型看的错误串，永不抛异常（验收 B9）：入口只有 `parseToolArguments`。
 */
import { z } from "zod";

/**
 * 列引用：列字母（`C`）或表头原文（`销售额`）；同名表头用 `销售额#2` 取第二个。
 * 真正的解析口径在 `sheet-io` 的 `resolveColumnRef`，这里只约束长度，避免出现第二套判定（§3.1）。
 */
const ColumnRef = z.string().min(1).max(60).describe("列引用：列字母（C）或表头原文（销售额）。同名表头写成「表头#2」。");

/** 首行是否当表头；所有读工具共用这一个口径。 */
const HasHeader = z.boolean().default(true).describe("首行是不是表头。为 true 时首行不进统计、也不参与行号定位。");

/** `sheet_describe` 的入参。 */
export const DescribeArgs = z
  .object({
    hasHeader: HasHeader,
    maxColumns: z
      .number()
      .int()
      .min(1)
      .max(200)
      .default(50)
      .describe("最多描述多少列，默认 50。宽表只回前若干列的画像，其余只报列数。"),
  })
  .strict();

/** `sheet_read_range` 的入参。 */
export const ReadRangeArgs = z
  .object({
    range: z.string().min(1).max(20).describe("A1 区域写法，例如 A1:D20；也可以写单格 C2。"),
    maxRows: z.number().int().min(1).max(200).default(50).describe("最多读多少行，默认 50；剩下的行只报总数。"),
  })
  .strict();

/** 一个聚合：函数 + 作用列（口径与 `sheet-io` 的 `AggregateSpec` 一致）。 */
const AggregateSpec = z
  .object({
    fn: z.enum(["count", "sum", "average", "min", "max"]).describe("聚合函数；只允许这五个，其它写法请落成 Excel 公式。"),
    column: ColumnRef,
  })
  .strict();

/** `sheet_aggregate` 的入参。 */
export const AggregateArgs = z
  .object({
    groupBy: z.array(ColumnRef).min(1).max(3).describe("分组列，1~3 个。再多就不是分析而是把整表搬过来。"),
    aggregates: z.array(AggregateSpec).min(1).max(5).describe("要算的聚合，1~5 个；数值一律由引擎算出。"),
    hasHeader: HasHeader,
    limit: z.number().int().min(1).max(100).default(50).describe("最多回多少个分组，默认 50；分组总数不受它影响。"),
  })
  .strict();

/** `sheet_find` 的入参。 */
export const FindArgs = z
  .object({
    column: ColumnRef,
    criteria: z
      .string()
      .min(1)
      .max(80)
      .describe("COUNTIF 风格的条件：>100、<=2、<>北京、直接写要等于的值。文本比较不区分大小写；日期列按序列号比较。"),
    hasHeader: HasHeader,
    limit: z.number().int().min(1).max(200).default(50).describe("最多回多少条命中，默认 50；命中总数不受它影响。"),
  })
  .strict();

/** `sheet_duplicates` 的入参（去重前先数清楚"能安全删掉几行"）。 */
export const DuplicatesArgs = z
  .object({
    keys: z.array(ColumnRef).min(1).max(5).describe("判定重复的键列，1~5 个；要求这些列在同一行上完全相同。"),
    hasHeader: HasHeader,
    limit: z.number().int().min(1).max(200).default(50).describe("最多回多少组，默认 50；重复组总数与多余行数不受它影响。"),
  })
  .strict();

/** 一条改格请求：地址 + 要写进去的文本。 */
const SetCellEdit = z
  .object({
    type: z.literal("set_cell"),
    address: z.string().min(1).max(12).describe("单元格地址，如 C5（1 基，与网格里看到的一致）。"),
    text: z
      .string()
      .max(2000)
      .describe("要写的文本。以 = 开头当公式（由引擎算缓存值）；日期写 2026-01-31 这种完整形式，`1-2` 会被当文本。"),
  })
  .strict();

/** 一条列宽请求。 */
const SetColumnWidthEdit = z
  .object({
    type: z.literal("set_column_width"),
    column: ColumnRef,
    /** 单位说在描述里：模型只会给"多宽"，换算成 Excel 的字符宽度由引擎做（AGENTS.md §2.1）。 */
    widthPx: z.number().min(20).max(2000).describe("列宽，单位是屏幕像素（px，不是 pt 也不是字符数）；换算由表格引擎做。"),
  })
  .strict();

/** 一条插入整行的请求。 */
const InsertRowsEdit = z
  .object({
    type: z.literal("insert_rows"),
    row: z.number().int().min(1).max(1_048_576).describe("在这一行**之前**插入（1 基行号）。"),
    count: z.number().int().min(1).max(500).default(1).describe("插入几行，默认 1。"),
  })
  .strict();

/** 一条删除整行的请求。 */
const DeleteRowsEdit = z
  .object({
    type: z.literal("delete_rows"),
    rows: z
      .array(z.number().int().min(1).max(1_048_576))
      .min(1)
      .max(500)
      .describe("要删除的行号（1 基，来自 sheet_find / sheet_duplicates，别自己数）。"),
  })
  .strict();

/** 一条数据系列：数值区域 + 可选的系列名单格引用（spec 决策 D5）。 */
const ChartSeries = z
  .object({
    values: z.string().min(1).max(20).describe("数值区域，**一列或一行**，如 C2:C4；二维区域会被拒。"),
    name: z
      .string()
      .min(1)
      .max(12)
      .optional()
      .describe("系列名的单格地址，通常是表头那一格（如 C1）。省略时用列字母当名。"),
  })
  .strict();

/**
 * `sheet_create_chart` 的入参：只有 A1 引用，没有任何数值。
 *
 * 数值、引用串、类别显示串一律由 `sheet-io` 的 `resolveChartSpec` 从现场表取（验收 B10），
 * 这里只约束形状；取不到数时引擎回一句理由，模型改引用再来。
 */
export const CreateChartArgs = z
  .object({
    kind: z.enum(["bar", "line", "pie"]).describe("图型：柱状 bar / 折线 line / 饼图 pie，只有这三种。"),
    categories: z.string().min(1).max(20).describe("类别轴区域，一列或一行，如 A2:A4；条数必须与每条系列的数值个数一致。"),
    series: z.array(ChartSeries).min(1).max(8).describe("要画的数据系列，1~8 条；数值由引擎读表得到，你不许自己填数。"),
    anchor: z.string().min(1).max(12).describe("图表左上角落在哪个格，如 E2（单格地址，不是区域）。"),
    title: z.string().max(60).optional().describe("图表标题；省略时用第一条系列的名字。"),
    widthPx: z.number().int().min(120).max(2000).optional().describe("图表框宽度，单位屏幕像素（px，不是 pt），默认 480。"),
    heightPx: z.number().int().min(90).max(2000).optional().describe("图表框高度，单位屏幕像素（px），默认 288。"),
  })
  .strict();

/**
 * 一条按一列重排整行的请求（spec 决策 D7）。
 *
 * 行序由引擎从现场表算（§2.11）：这里只交区域、键列与升降，模型不许自己报一个新行序，
 * 也不许交"排序后的值"。
 */
const SortRowsEdit = z
  .object({
    type: z.literal("sort_rows"),
    range: z
      .string()
      .min(3)
      .max(20)
      .describe("参与重排的数据行区域（不含表头那一行），如 A2:D10；起止都要写全，整列写法（A:C）会被拒。列范围必须盖住一条记录的所有格子。"),
    column: ColumnRef,
    order: z.enum(["asc", "desc"]).default("asc").describe("升序 asc / 降序 desc，默认升序。"),
  })
  .strict();

/**
 * 一条把整列数字格式统一成某个格式的请求（spec 决策 D10）。
 *
 * 逐格的展开由引擎的 `numberFormatColumnOps` 完成，这里只交列与格式码；模型不许逐格报"这格该显示成什么"（§2.11）。
 */
const SetNumberFormatEdit = z
  .object({
    type: z.literal("set_number_format"),
    column: ColumnRef,
    format: z
      .string()
      .min(1)
      .max(30)
      .describe(
        "目标数字格式码，只能是我们认得的这些：General、0、0.00、#,##0、#,##0.00、0%、0.00%、m/d/yyyy，" +
          "小一位数的 0.0 与 0.0%，以及日期时间的自定义写法（如 yyyy-mm-dd）。表外的码引擎会整条拒掉并说明原因。",
      ),
  })
  .strict();

/** 提案里的一条改动。刻意只留这六类：它们各自有安全的逆操作（AGENTS.md §2.5）。 */
export const SheetEdit = z
  .discriminatedUnion("type", [SetCellEdit, SetColumnWidthEdit, InsertRowsEdit, DeleteRowsEdit, SortRowsEdit, SetNumberFormatEdit]);

/**
 * `sheet_submit_proposal` 的入参。
 *
 * 排序后来按 spec 决策 D7 单独立项补成了 `sort_rows`，统一数字格式按 **D10** 补成了 `set_number_format`；
 * **仍不在这里**的是拆合列（D9 第 3 条要求它展开成既有 op 的复合批，代码未起），模型只能提这六项。
 */
export const SubmitProposalArgs = z
  .object({
    summary: z.string().min(1).max(200).describe("一句话中文说明这份提案要做什么。"),
    edits: z.array(SheetEdit).min(1).max(500).describe("按顺序落地的改动，1~500 条。"),
  })
  .strict();

/** 工具名 → 入参 schema；`parseToolArguments` 按这个名字分派。 */
export const SHEET_TOOL_SCHEMAS = {
  sheet_describe: DescribeArgs,
  sheet_read_range: ReadRangeArgs,
  sheet_aggregate: AggregateArgs,
  sheet_find: FindArgs,
  sheet_duplicates: DuplicatesArgs,
  sheet_create_chart: CreateChartArgs,
  sheet_submit_proposal: SubmitProposalArgs,
} as const;

/** 工具名（模型回调时能出现的合法集合，用来拦住不存在的方法名）。 */
export type SheetToolName = keyof typeof SHEET_TOOL_SCHEMAS;

export type DescribeArgsT = z.infer<typeof DescribeArgs>;
export type ReadRangeArgsT = z.infer<typeof ReadRangeArgs>;
export type AggregateArgsT = z.infer<typeof AggregateArgs>;
export type FindArgsT = z.infer<typeof FindArgs>;
export type DuplicatesArgsT = z.infer<typeof DuplicatesArgs>;
export type CreateChartArgsT = z.infer<typeof CreateChartArgs>;
export type SubmitProposalArgsT = z.infer<typeof SubmitProposalArgs>;
export type SheetEditT = z.infer<typeof SheetEdit>;

/** 入参解析结果：`ok:false` 时 `reason` 是直接回给模型的错误串（验收 B9）。 */
export type ToolArgsOutcome<T> = { ok: true; args: T } | { ok: false; reason: string };

/**
 * 解析模型给的入参：JSON 语法错、缺字段、多字段、越界都变成错误串，**永不抛异常**。
 *
 * @param name 工具名；不在 `SHEET_TOOL_SCHEMAS` 里时给"不认识的工具"。
 * @param rawArguments 模型回包里的 `function.arguments` 原文（JSON 字符串）。
 * @returns 校验过的入参，或给模型看的失败理由。
 */
export function parseToolArguments<T extends SheetToolName>(name: T, rawArguments: string): ToolArgsOutcome<z.infer<(typeof SHEET_TOOL_SCHEMAS)[T]>> {
  const schema = SHEET_TOOL_SCHEMAS[name];
  if (!schema) return { ok: false, reason: `不认识的工具：${String(name)}` };
  let parsedJson: unknown;
  try {
    parsedJson = JSON.parse(rawArguments);
  } catch {
    return { ok: false, reason: `工具 ${name} 的入参不是合法 JSON，请按 schema 重写。` };
  }
  const result = schema.safeParse(parsedJson);
  if (result.success) return { ok: true, args: result.data as z.infer<(typeof SHEET_TOOL_SCHEMAS)[T]> };
  const details = result.error.issues.map((issue) => `${issue.path.join(".") || "(根对象)"}：${issue.message}`).join("；");
  return { ok: false, reason: `工具 ${name} 的入参不合：${details}` };
}

/** 一条工具定义（OpenAI function-calling 形状，与 `WORD_TOOLS` 同形）。 */
export interface SheetToolDefinition {
  type: "function";
  function: {
    name: SheetToolName;
    description: string;
    /** JSON Schema 的极小子集：我们只用 `properties` / `required` / `additionalProperties` 这三样。 */
    parameters: { type: "object"; properties?: Record<string, unknown>; required?: string[]; additionalProperties: false };
  };
}

/** 给模型的读工具定义，直接进请求的 `tools` 数组。 */
export const SHEET_TOOLS: SheetToolDefinition[] = [
  {
    type: "function",
    function: {
      name: "sheet_describe",
      description: "读出当前表的结构画像：列数、数据行数、每列的类型推断、缺失数、众数、数值统计与异常值。要做任何分析或整理之前先调用它。",
      parameters: {
        type: "object",
        properties: {
          hasHeader: { type: "boolean", default: true },
          maxColumns: { type: "integer", minimum: 1, maximum: 200, default: 50 },
        },
        additionalProperties: false,
      },
    },
  },
  {
    type: "function",
    function: {
      name: "sheet_read_range",
      description: "读一块区域的显示文本（与网格里看到的一致：日期、千分位、百分比都按单元格格式显示）。需要看具体数据长什么样时用，别整表上送。",
      parameters: {
        type: "object",
        properties: {
          range: { type: "string", minLength: 1, maxLength: 20 },
          maxRows: { type: "integer", minimum: 1, maximum: 200, default: 50 },
        },
        required: ["range"],
        additionalProperties: false,
      },
    },
  },
  {
    type: "function",
    function: {
      name: "sheet_aggregate",
      description: "按 1~3 列分组，对指定列算 count/sum/average/min/max。**数值由表格引擎算出，不要自己心算**。",
      parameters: {
        type: "object",
        properties: {
          groupBy: { type: "array", items: { type: "string" }, minItems: 1, maxItems: 3 },
          aggregates: {
            type: "array",
            items: {
              type: "object",
              properties: {
                fn: { type: "string", enum: ["count", "sum", "average", "min", "max"] },
                column: { type: "string" },
              },
              required: ["fn", "column"],
              additionalProperties: false,
            },
            minItems: 1,
            maxItems: 5,
          },
          hasHeader: { type: "boolean", default: true },
          limit: { type: "integer", minimum: 1, maximum: 100, default: 50 },
        },
        required: ["groupBy", "aggregates"],
        additionalProperties: false,
      },
    },
  },
  {
    type: "function",
    function: {
      name: "sheet_find",
      description: "按条件定位数据行，返回行号与单元格地址。条件写法与 Excel 的 COUNTIF 一致（>100、<=2、<>北京、直接写要等于的值）。",
      parameters: {
        type: "object",
        properties: {
          column: { type: "string" },
          criteria: { type: "string", minLength: 1, maxLength: 80 },
          hasHeader: { type: "boolean", default: true },
          limit: { type: "integer", minimum: 1, maximum: 200, default: 50 },
        },
        required: ["column", "criteria"],
        additionalProperties: false,
      },
    },
  },
  {
    type: "function",
    function: {
      name: "sheet_duplicates",
      description: "按若干键列找完全重复的行，返回每组的行号与「能安全删掉几行」。要去重之前先调它数清楚，别凭印象删行。",
      parameters: {
        type: "object",
        properties: {
          keys: { type: "array", items: { type: "string" }, minItems: 1, maxItems: 5 },
          hasHeader: { type: "boolean", default: true },
          limit: { type: "integer", minimum: 1, maximum: 200, default: 50 },
        },
        required: ["keys"],
        additionalProperties: false,
      },
    },
  },
  {
    type: "function",
    function: {
      name: "sheet_create_chart",
      description:
        "提议在当前工作表插入一张**原生图表**（不是贴图）。只交区域引用：数值、类别显示文字、引用串全部由表格引擎从现场表读，" +
        "你不许自己填数也不许报数字。类别轴与每条系列的条数必须一致；系列名通常填表头那一格的地址（如 C1）。" +
        "取不到数（二维区域、条数对不上、格里是文本）引擎会回一句理由，改引用再调。用户确认后才会随文件保存。",
      parameters: {
        type: "object",
        properties: {
          kind: { type: "string", enum: ["bar", "line", "pie"] },
          categories: { type: "string", minLength: 1, maxLength: 20 },
          series: {
            type: "array",
            minItems: 1,
            maxItems: 8,
            items: {
              type: "object",
              properties: {
                values: { type: "string", minLength: 1, maxLength: 20 },
                name: { type: "string", minLength: 1, maxLength: 12 },
              },
              required: ["values"],
              additionalProperties: false,
            },
          },
          anchor: { type: "string", minLength: 1, maxLength: 12 },
          title: { type: "string", maxLength: 60 },
          widthPx: { type: "integer", minimum: 120, maximum: 2000 },
          heightPx: { type: "integer", minimum: 90, maximum: 2000 },
        },
        required: ["kind", "categories", "series", "anchor"],
        additionalProperties: false,
      },
    },
  },
  {
    type: "function",
    function: {
      name: "sheet_submit_proposal",
      description:
        "把你打算对表格做的修改打包成提案返回。用户会看到逐条改动与跳过原因并确认后才落地。一次回答只提交一份提案。改格用 set_cell（文本以 = 开头当公式），列宽用 set_column_width（px），整行用 insert_rows / delete_rows，按一列重排整行用 sort_rows（只交区域与键列，行序由引擎算），整列统一数字格式用 set_number_format（只交列与格式码，逐格由引擎按现场判定展开）；拆合列现在提不了。",
      parameters: {
        type: "object",
        properties: {
          summary: { type: "string", minLength: 1, maxLength: 200 },
          edits: {
            type: "array",
            minItems: 1,
            maxItems: 500,
            items: {
              type: "object",
              oneOf: [
                {
                  type: "object",
                  properties: { type: { const: "set_cell" }, address: { type: "string" }, text: { type: "string" } },
                  required: ["type", "address", "text"],
                  additionalProperties: false,
                },
                {
                  type: "object",
                  properties: { type: { const: "set_column_width" }, column: { type: "string" }, widthPx: { type: "number" } },
                  required: ["type", "column", "widthPx"],
                  additionalProperties: false,
                },
                {
                  type: "object",
                  properties: { type: { const: "insert_rows" }, row: { type: "integer" }, count: { type: "integer", default: 1 } },
                  required: ["type", "row"],
                  additionalProperties: false,
                },
                {
                  type: "object",
                  properties: { type: { const: "delete_rows" }, rows: { type: "array", items: { type: "integer" } } },
                  required: ["type", "rows"],
                  additionalProperties: false,
                },
                {
                  type: "object",
                  properties: {
                    type: { const: "sort_rows" },
                    range: { type: "string" },
                    column: { type: "string" },
                    order: { type: "string", enum: ["asc", "desc"], default: "asc" },
                  },
                  required: ["type", "range", "column"],
                  additionalProperties: false,
                },
                {
                  type: "object",
                  properties: { type: { const: "set_number_format" }, column: { type: "string" }, format: { type: "string" } },
                  required: ["type", "column", "format"],
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
