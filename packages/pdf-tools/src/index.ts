/**
 * PDF 的 AI 工具面（spec 阶段 7 / 决策 D6，pdf 一线）：给模型看的 JSON 工具定义 + zod 入参校验。
 *
 * 与 `sheet-tools`/`word-tools`/`pptx-tools` 同构，守同一条红线：模型只产出**入参与提案**，定位、落地、撤销
 * 全部由 `@office-copilot/pdf-io` 做（`AGENTS.md` §2.8、§2.11）。写面只有 `pdf_submit_proposal` 一个出口，
 * 而它现在只容得下**两类改动**——`set_run_text`（同位置、同字数的整条替换）与 `set_line_text`（同字数的整行替换，
 * 由引擎拆回逐 run，spec D8 的 W1.5）。
 * 改字体、改颜色、改几何、删页这些还没建成自带逆操作的 op，宁可不提供，也不发一批没有逆操作的直写（§2.5）。
 *
 * 按 spec §三 的包边界，本包只依赖 `zod` 与 `@office-copilot/shared`（入参判定共用一份）、**不 import `pdf-io`**：
 * 它只管"这套入参合不合法"，执行由调用方（桌面 `pdfAgent`）拿校验过的入参去调引擎函数——
 * 把 `{pageNumber, runIndex}` 或 `{pageNumber, lineIndex}` 按现场模型翻成 `PdfTextTarget`（内容流引用 + Form 链 + 起始下标），
 * 再交给 `applyPdfTextEdits`。那几个引擎内部坐标跨不出这道边，模型既看不到也不许编。
 *
 * 坐标一律是"模型视角"：页 1 基（`pdf_describe` 报第几页）、页内 run 序号与页内行号 0 基（与 `pdf_read_page_text`
 * 列出的 `runIndex`/`lineIndex` 一致）。等长这条规则说在描述里，但**判定在引擎编码时**做（§2.1：字节换算只发生在 `pdf-io` 进出口）。
 * 非法入参一律换成给模型看的错误串，永不抛异常（验收 B9）：入口只有 `parseToolArguments`。
 */
import { z } from "zod";
import { parseToolArguments as parseToolArgumentsAgainst, type ToolArgsOutcome } from "@office-copilot/shared";

/** PDF 页号，1 基（模型数页从 1 开始，和界面上"第 N 页"一致）。 */
const PageNumber = z.number().int().min(1).max(500).describe("PDF 页号，1 基（第 1 页就是 1）。来自 pdf_describe 列出的页号，别自己猜。");

/** 页内 run 序号，0 基（该页第几条显示文字；上限取宽一点，长表格页的 run 本来就多）。 */
const RunIndex = z.number().int().min(0).max(5000).describe("该页内的显示文字条序号，0 基。必须与 pdf_read_page_text 列出的 runIndex 一致，别凭印象数。");

/** 页内逻辑行号，0 基（该行由几条 run 拼成都不影响它——引擎按现场模型算，模型只管报第几行）。 */
const LineIndex = z.number().int().min(0).max(5000).describe("该页内的逻辑文本行号，0 基。必须与 pdf_read_page_text 列出的 lineIndex 一致，别凭印象数。");

/**
 * `pdf_describe` 的入参：读整份 PDF 的页概要，让模型先知道有几页、每页大概念什么。
 *
 * 页数与文字概要一律由引擎从现场 `.pdf` 读（§2.11），这里只约束翻页的上限。
 */
export const DescribeArgs = z
  .object({
    maxPages: z.number().int().min(1).max(200).default(50).describe("最多描述多少页，默认 50；剩下的页只报总数。"),
  })
  .strict();

/** `pdf_read_page_text` 的入参：把某一页每条显示文字摊开，好让模型填对 runIndex 并写出等长的新文字。 */
export const ReadPageTextArgs = z
  .object({
    pageNumber: PageNumber,
    maxRunsPerPage: z.number().int().min(1).max(200).default(50).describe("每页最多列多少条 run，默认 50；超出会标 runsTruncated，加大它才能看到后面那些。"),
  })
  .strict();

/**
 * 一条改某条显示文字的请求。W1 档只做"同位置、同字数"的替换：
 * PDF 没有回流，字节数一变就装不回原来的位置，引擎会整批不写（spec D8）。
 */
const SetRunTextEdit = z
  .object({
    type: z.literal("set_run_text"),
    pageNumber: PageNumber,
    runIndex: RunIndex,
    text: z
      .string()
      .min(1)
      .max(2000)
      .describe(
        "这条 run 要换成的新文字。**字数必须与 pdf_read_page_text 报的 charCount 完全相同**（不等长就整批不写），" +
          "而且只许用这一页里已经出现过的字——文件里的字体是子集，没带的字补不出来。只动这一条 run，同页其它条不受影响。",
      ),
  })
  .strict();

/**
 * 一条改**整行**的请求（spec D8 的 W1.5）：中文产物一个字发一条 run，"改一句话"要发几十条 `set_run_text`，
 * 所以这一类由引擎把那行拆回逐 run 的等长替换——**每条 run 分到它原来那几个字**，硬约束仍然是整行字数不变。
 */
const SetLineTextEdit = z
  .object({
    type: z.literal("set_line_text"),
    pageNumber: PageNumber,
    lineIndex: LineIndex,
    text: z
      .string()
      .min(1)
      .max(2000)
      .describe(
        "这一行要换成的新文字。**总字数必须与 pdf_read_page_text 报的该行 charCount 完全相同**（不等长整行不动），" +
          "而且只许用这一页里已经出现过的字——文件里的字体是子集，没带的字补不出来。" +
          "引擎会按每条 run 原有的字数依次切开，所以你只管报第几行与整句新文字，不用管那行由几条 run 拼成。",
      ),
  })
  .strict();

/** 提案里的一条改动。只给有安全逆操作的那一类（AGENTS.md §2.5），其余等引擎补上带逆操作的 op 再开。 */
export const PdfEdit = z.discriminatedUnion("type", [SetRunTextEdit, SetLineTextEdit]);

/**
 * `pdf_submit_proposal` 的入参。
 *
 * **不在这里**的是字体、颜色、几何、删页与整段合并（都还没建成带逆操作的 op）。
 * 一份提案里同一条 run 出现两次是允许的（落地按顺序执行，后面那条覆盖前面那条），代理不做合并。
 */
export const SubmitProposalArgs = z
  .object({
    summary: z.string().min(1).max(200).describe("一句话中文说明这份提案要做什么。"),
    edits: z.array(PdfEdit).min(1).max(200).describe("按顺序落地的改动，1~200 条。"),
  })
  .strict();

/** 工具名 → 入参 schema；`parseToolArguments` 按这个名字分派。 */
export const PDF_TOOL_SCHEMAS = {
  pdf_describe: DescribeArgs,
  pdf_read_page_text: ReadPageTextArgs,
  pdf_submit_proposal: SubmitProposalArgs,
} as const;

/** 工具名（模型回调时能出现的合法集合，用来拦住不存在的方法名）。 */
export type PdfToolName = keyof typeof PDF_TOOL_SCHEMAS;

export type DescribeArgsT = z.infer<typeof DescribeArgs>;
export type ReadPageTextArgsT = z.infer<typeof ReadPageTextArgs>;
export type SubmitProposalArgsT = z.infer<typeof SubmitProposalArgs>;
export type PdfEditT = z.infer<typeof PdfEdit>;

/**
 * 解析模型给的入参：JSON 语法错、缺字段、多字段、越界都换成给模型看的错误串，**永不抛异常**（验收 B9）。
 *
 * 判定本身在 `@office-copilot/shared`——几个工具面用的是同一条，抽出来只留一份（`AGENTS.md` §3.2）；
 * 这里只把本包那张 schema 表接上，让调用方拿到按工具名收窄过的入参类型。
 *
 * @param name 工具名；不在 `PDF_TOOL_SCHEMAS` 里时给"不认识的工具"。
 * @param rawArguments 模型回包里的 `function.arguments` 原文（JSON 字符串）。
 * @returns 校验过的入参，或给模型看的失败理由。
 */
export function parseToolArguments<T extends PdfToolName>(name: T, rawArguments: string): ToolArgsOutcome<z.infer<(typeof PDF_TOOL_SCHEMAS)[T]>> {
  return parseToolArgumentsAgainst(PDF_TOOL_SCHEMAS, name, rawArguments);
}

/** 一条工具定义（OpenAI function-calling 形状，与 `SHEET_TOOLS`/`WORD_TOOLS`/`PPTX_TOOLS` 同形）。 */
export interface PdfToolDefinition {
  type: "function";
  function: {
    name: PdfToolName;
    description: string;
    /** JSON Schema 的极小子集：我们只用 `properties` / `required` / `additionalProperties` 这三样。 */
    parameters: { type: "object"; properties?: Record<string, unknown>; required?: string[]; additionalProperties: false };
  };
}

/** 给模型的读工具与写工具定义，直接进请求的 `tools` 数组。 */
export const PDF_TOOLS: PdfToolDefinition[] = [
  {
    type: "function",
    function: {
      name: "pdf_describe",
      description:
        "读出当前 PDF 的结构概要：一共几页、每页有几条显示文字、其中几条能改、这一页念什么（概要文字超长会截并标记）。" +
        "要做任何编辑之前先调用它拿到 pageNumber；run 级序号要靠 pdf_read_page_text 列出，不许自己编。" +
        "页数与文字一律由引擎从现场文件读，你不许换算字节、不许猜内容流位置。",
      parameters: {
        type: "object",
        properties: {
          maxPages: { type: "integer", minimum: 1, maximum: 200, default: 50 },
        },
        additionalProperties: false,
      },
    },
  },
  {
    type: "function",
    function: {
      name: "pdf_read_page_text",
      description:
        "把某一页摊开列出两种粒度：**逐条显示文字**（runIndex、现文字、charCount 原文字数、encodedBytes、字体与是否 CID、" +
        "在哪个 Form 里，以及改不了的话卡在哪一类：内联内容流 inline-stream / 认不出原文 undecoded）与" +
        "**逐逻辑行**（lineIndex、整行文字、整行 charCount、这一行由第几条到第几条 run 拼成）。" +
        "提交 set_run_text 前照着 run 的 charCount 写等长的新文字，提交 set_line_text 前照着该行的 charCount 写等长的整句；" +
        "两个序号都来自这一份现场模型，别凭印象数。",
      parameters: {
        type: "object",
        properties: {
          pageNumber: { type: "integer", minimum: 1, maximum: 500 },
          maxRunsPerPage: { type: "integer", minimum: 1, maximum: 200, default: 50 },
        },
        required: ["pageNumber"],
        additionalProperties: false,
      },
    },
  },
  {
    type: "function",
    function: {
      name: "pdf_submit_proposal",
      description:
        "把你打算对 PDF 做的修改打包成提案返回。用户会看到逐条改动并确认后才落地，一次回答只提交一份提案。" +
        "现在有两类改动：set_run_text 换掉某页某条显示文字（同位置整条替换，可撤销）；" +
        "set_line_text 换掉某页某**逻辑文本行**的整句文字（引擎按该行原来每条 run 的字数依次切开，同样可撤销）。" +
        "**中文产物往往一个字就是一条 run**，改一句话要几十条 set_run_text，那种情形直接用 set_line_text 报整行，别一条条拼。" +
        "硬约束是**新文字与原文字数完全相同**（PDF 没有回流，字节数变了引擎会整批不写并报长度不符），" +
        "并且新文字里每个字都得是这一页已经用过的字（文件里的字体是子集，缺的字补不出来）。" +
        "改字体、改颜色、改几何、删页现在都提不了，遇到就直说做不了。",
      parameters: {
        type: "object",
        properties: {
          summary: { type: "string", minLength: 1, maxLength: 200 },
          edits: {
            type: "array",
            minItems: 1,
            maxItems: 200,
            items: {
              anyOf: [
                {
                  type: "object",
                  properties: {
                    type: { const: "set_run_text" },
                    pageNumber: { type: "integer", minimum: 1, maximum: 500 },
                    runIndex: { type: "integer", minimum: 0, maximum: 5000 },
                    text: { type: "string", minLength: 1, maxLength: 2000 },
                  },
                  required: ["type", "pageNumber", "runIndex", "text"],
                  additionalProperties: false,
                },
                {
                  type: "object",
                  properties: {
                    type: { const: "set_line_text" },
                    pageNumber: { type: "integer", minimum: 1, maximum: 500 },
                    lineIndex: { type: "integer", minimum: 0, maximum: 5000 },
                    text: { type: "string", minLength: 1, maxLength: 2000 },
                  },
                  required: ["type", "pageNumber", "lineIndex", "text"],
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
