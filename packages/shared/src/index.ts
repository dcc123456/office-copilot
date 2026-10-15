// Shared types and protocols used by the task pane and the server，外加各 IO 包与工具包共用的极小工具。
import { z } from "zod";

/** 一条上送给模型看的文字：截断后的正文与"有没有被截"的标记。 */
export interface ClippedText {
  /** 上送的文字；超上限时是前 `limit` 个字加一个省略号。 */
  text: string;
  /** 是否截掉了后半段（§2.9：截断必须说出来，别让模型以为看到全部）。 */
  truncated: boolean;
}

/**
 * 把一条过长的文字截成上送给模型看的长度。
 *
 * @param value 现场读出的原文。
 * @param limit 上限（字数，按 UTF-16 码元算，与各 IO 包的 `*_TEXT_LIMIT` 常量同一个口径）。
 * @returns 不超时原样且 `truncated:false`；超了取前 `limit` 个字并补 `…`。
 */
export function clipText(value: string, limit: number): ClippedText {
  return value.length <= limit ? { text: value, truncated: false } : { text: `${value.slice(0, limit)}…`, truncated: true };
}

export type InsertPosition = "before" | "after" | "start" | "end";
export type EditScope = "selection" | "body";

export interface ParagraphInfo {
  index: number;
  text: string;
  style?: string;
}

export interface SelectionInfo {
  text: string;
  /** Length hint (characters). */
  length: number;
}

/** A single edit the assistant proposes; the user must confirm before it runs. */
export type ProposedEdit =
  | {
      type: "replace_text";
      find: string;
      replace: string;
      regex?: boolean;
      scope: EditScope;
      /** Per-match preview: [before, after]. */
      previews: Array<[string, string]>;
    }
  | {
      type: "insert_paragraph";
      text: string;
      position: InsertPosition;
      /** 0-based paragraph index, required when position is before/after. */
      anchor?: number;
      style?: string;
    }
  | {
      type: "insert_table";
      rows: string[][];
      headerRow: boolean;
    }
  | {
      type: "insert_image";
      /** Data URL (base64) — never sent to the LLM. */
      dataUrl: string;
      width?: number;
      height?: number;
    };

export interface Proposal {
  id: string;
  summary: string;
  edits: ProposedEdit[];
  /** ISO timestamp. */
  createdAt: string;
}

export interface Snapshot {
  id: string;
  /** OOXML fragment captured before an edit batch, used to undo. */
  ooxml: string;
  createdAt: string;
}

/** Request body for POST /api/word/command */
export interface CommandRequest {
  sessionId: string;
  instruction: string;
  /** Current selection text, if any. */
  selection?: string;
  /** Paragraph outline (index + text, truncated) for context. */
  outline?: ParagraphInfo[];
}

/** Server-sent event shapes for /api/word/command. */
export type CommandEvent =
  | { event: "status"; message: string }
  | { event: "tool_call"; name: string; args: unknown }
  | { event: "proposal"; proposal: Proposal }
  | { event: "error"; message: string }
  | { event: "done" };

export function newId(prefix = "id"): string {
  const rand = Math.random().toString(36).slice(2, 10);
  return `${prefix}_${Date.now().toString(36)}_${rand}`;
}

/** 工具名 → 入参 schema 的表（每个 `*-tools` 包自己那份，键就是模型能点名的工具）。 */
export type ToolSchemaMap = Record<string, z.ZodTypeAny>;

/** 入参解析结果：`ok:false` 时 `reason` 是直接回给模型的错误串（验收 B9）。 */
export type ToolArgsOutcome<TArgs> = { ok: true; args: TArgs } | { ok: false; reason: string };

/**
 * 解析模型给的入参：JSON 语法错、缺字段、多字段、越界都变成错误串，**永不抛异常**。
 *
 * 表格 / 幻灯 / PDF 三个工具面用的是同一条判定，所以这里只有一份（`AGENTS.md` §3.2）；
 * 各包仍各自导出一个 `parseToolArguments`，把**自己那张 schema 表**传进来，模型点错工具名时也回中文而不是抛。
 *
 * @param schemas 该包的工具名 → zod schema 表。
 * @param name 模型点名的工具；不在表里时回"不认识的工具"。
 * @param rawArguments 模型回包里的 `function.arguments` 原文（应为 JSON 字符串）。
 * @returns 校验过的入参，或给模型看的失败理由。
 */
export function parseToolArguments<TSchemas extends ToolSchemaMap, TName extends keyof TSchemas & string>(
  schemas: TSchemas,
  name: TName,
  rawArguments: string,
): ToolArgsOutcome<z.infer<TSchemas[TName]>> {
  const schema = schemas[name];
  if (!schema) return { ok: false, reason: `不认识的工具：${String(name)}` };
  let parsedJson: unknown;
  try {
    parsedJson = JSON.parse(rawArguments);
  } catch {
    return { ok: false, reason: `工具 ${String(name)} 的入参不是合法 JSON，请按 schema 重写。` };
  }
  const result = schema.safeParse(parsedJson);
  if (result.success) return { ok: true, args: result.data as z.infer<TSchemas[TName]> };
  const details = result.error.issues.map((issue) => `${issue.path.join(".") || "(根对象)"}：${issue.message}`).join("；");
  return { ok: false, reason: `工具 ${String(name)} 的入参不合：${details}` };
}
