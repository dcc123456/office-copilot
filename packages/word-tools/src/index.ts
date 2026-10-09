import { z } from "zod";
import type { ProposedEdit } from "@office-copilot/shared";

/**
 * Tool schemas exposed to the LLM. We deliberately keep these narrow: the model
 * proposes edits, it does NOT execute them — the UI asks the user to confirm,
 * then the task pane calls into @office-copilot/word-client.
 */

export const ListParagraphsArgs = z.object({}).strict();

export const ReplaceTextArgs = z
  .object({
    find: z.string().min(1).describe("Literal text to find."),
    replace: z.string().describe("Replacement text. Use '' to delete."),
    regex: z.boolean().default(false).describe("If true, treat `find` as a JavaScript regex."),
    scope: z.enum(["selection", "body"]).default("body"),
  })
  .strict();

export const InsertParagraphArgs = z
  .object({
    text: z.string().min(1),
    position: z.enum(["before", "after", "start", "end"]).default("end"),
    anchor: z
      .number()
      .int()
      .nonnegative()
      .optional()
      .describe("Required when position is 'before'/'after': the 0-based paragraph index from word_list_paragraphs."),
    style: z.string().optional().describe("Optional Word style name, e.g. 'Heading 1'."),
  })
  .strict();

export const InsertTableArgs = z
  .object({
    rows: z
      .array(z.array(z.string()))
      .min(1)
      .describe("2D array of cell values; first row is treated as header when headerRow=true."),
    headerRow: z.boolean().default(true),
  })
  .strict();

export const SubmitProposalArgs = z
  .object({
    summary: z
      .string()
      .min(1)
      .describe("One-sentence Chinese summary of what this proposal will do."),
    edits: z
      .array(
        z.discriminatedUnion("type", [
          z.object({
            type: z.literal("replace_text"),
            find: z.string().min(1),
            replace: z.string(),
            regex: z.boolean().optional(),
            scope: z.enum(["selection", "body"]).default("body"),
          }),
          z.object({
            type: z.literal("insert_paragraph"),
            text: z.string().min(1),
            position: z.enum(["before", "after", "start", "end"]).default("end"),
            anchor: z.number().int().nonnegative().optional(),
            style: z.string().optional(),
          }),
          z.object({
            type: z.literal("insert_table"),
            rows: z.array(z.array(z.string())).min(1),
            headerRow: z.boolean().default(true),
          }),
        ]),
      )
      .min(1)
      .describe("The list of edits to apply, in order. Image insertion is not LLM-driven."),
  })
  .strict();

export type ReplaceTextArgsT = z.infer<typeof ReplaceTextArgs>;
export type InsertParagraphArgsT = z.infer<typeof InsertParagraphArgs>;
export type InsertTableArgsT = z.infer<typeof InsertTableArgs>;
export type SubmitProposalArgsT = z.infer<typeof SubmitProposalArgs>;

/** OpenAI-style tool definitions consumed by the LLM. */
export const WORD_TOOLS = [
  {
    type: "function" as const,
    function: {
      name: "word_list_paragraphs",
      description:
        "列出当前 Word 文档的段落大纲（编号 + 文本预览）。在需要定位段落时先调用此工具。",
      parameters: { type: "object", properties: {}, additionalProperties: false },
    },
  },
  {
    type: "function" as const,
    function: {
      name: "word_submit_proposal",
      description:
        "把你打算对文档做的修改打包成一份提案返回。用户会看到 diff 并确认后才会真正写入。一次回答只提交一份提案。",
      parameters: {
        type: "object",
        properties: {
          summary: { type: "string" },
          edits: {
            type: "array",
            items: {
              type: "object",
              oneOf: [
                {
                  type: "object",
                  properties: {
                    type: { const: "replace_text" },
                    find: { type: "string" },
                    replace: { type: "string" },
                    regex: { type: "boolean" },
                    scope: { type: "string", enum: ["selection", "body"] },
                  },
                  required: ["type", "find", "replace"],
                  additionalProperties: false,
                },
                {
                  type: "object",
                  properties: {
                    type: { const: "insert_paragraph" },
                    text: { type: "string" },
                    position: { type: "string", enum: ["before", "after", "start", "end"] },
                    style: { type: "string" },
                  },
                  required: ["type", "text"],
                  additionalProperties: false,
                },
                {
                  type: "object",
                  properties: {
                    type: { const: "insert_table" },
                    rows: { type: "array", items: { type: "array", items: { type: "string" } } },
                    headerRow: { type: "boolean" },
                  },
                  required: ["type", "rows"],
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

export const TOOL_NAMES = new Set(WORD_TOOLS.map((t) => t.function.name));

/**
 * Convert validated submit-proposal args into shared ProposedEdit[]; caller
 * (the task pane) is responsible for computing replace-text previews against
 * the live document before showing them to the user.
 */
export function editsFromProposal(args: SubmitProposalArgsT): ProposedEdit[] {
  return args.edits.map((e): ProposedEdit => {
    switch (e.type) {
      case "replace_text":
        return {
          type: "replace_text",
          find: e.find,
          replace: e.replace,
          regex: e.regex,
          scope: e.scope ?? "body",
          previews: [],
        };
      case "insert_paragraph":
        return {
          type: "insert_paragraph",
          text: e.text,
          position: e.position ?? "end",
          anchor: e.anchor,
          style: e.style,
        };
      case "insert_table":
        return {
          type: "insert_table",
          rows: e.rows,
          headerRow: e.headerRow ?? true,
        };
    }
  });
}
