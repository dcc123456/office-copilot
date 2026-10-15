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

/** `.pdf` 结构体检的严重程度：`broken` = 办公软件可能打不开或丢内容；`limited` = 打得开，但我们这一档改不动。 */
export type PdfAuditSeverity = "broken" | "limited";

/**
 * `.pdf` 体检的问题码集合（判定住在 `@office-copilot/pdf-io`，这里只有码与文案）。
 *
 * 放这一层的原因是 `oc pdf check` 与桌面打开时那句提示念的是**同一套话**（§3.2 同一逻辑第二次就抽公共层），
 * 而文案不许沉进引擎（§5.7）；依赖方向仍是 `pdf-io → shared`（§5.1）。
 */
export type PdfAuditCode =
  | "parse-failed"
  | "header-missing"
  | "eof-missing"
  | "catalog-missing"
  | "page-count-mismatch"
  | "content-stream-undecodable"
  | "text-not-editable";

/** 一条 `.pdf` 体检结论。 */
export interface PdfAuditIssue {
  /** 问题码。 */
  code: PdfAuditCode;
  /** 严重程度。 */
  severity: PdfAuditSeverity;
  /** 出问题的页号（**1 基**，与 PDF 读模型的坐标同一口径）；整份文件级别的问题是 null。 */
  pageNumber: number | null;
  /** 引擎看到的数字或原文片段，装配层直接念给人听。 */
  detail: string;
}

/** 每个问题码一句中文；键类型是 `PdfAuditCode`，新增码忘了配文案会编译不过。 */
const PDF_AUDIT_TEXT: Record<PdfAuditCode, string> = {
  "parse-failed": "解析器读不动这份文件（交叉引用表或对象流坏了），办公软件多半也打不开",
  "header-missing": "文件头没有 %PDF- 版本标记，阅读器可能直接拒开",
  "eof-missing": "文件尾没有 %%EOF，说明这份是被截断的，末尾的页与交叉引用可能不完整",
  "catalog-missing": "目录（/Root 或 /Pages）在文件里查不到，阅读器根本定位不到页",
  "page-count-mismatch": "页树的数量对不上，阅读器会按自己那套猜页数",
  "content-stream-undecodable": "这一页的内容流解不出来，打开后这一页会是空白",
  "text-not-editable": "有一部分文字我们读不出内容，因而也改不了",
};

/**
 * 把一条 `.pdf` 体检结论翻成一句中文。
 *
 * @param issue 引擎给的那条结论（`pageNumber` 是 1 基页号或 null）。
 * @param withDetail 带不带 `：` 后面那串数字/字节原文，默认带。toast 那一档放不下（一句就够长），传 `false`。
 * @returns 「第 N 页 · 」前缀（整份文件级别的问题没有）+ 文案（+ `：` 后的 detail）；detail 为空时不带冒号。
 */
export function describePdfAuditIssue(issue: PdfAuditIssue, withDetail = true): string {
  const where = issue.pageNumber === null ? "" : `第 ${issue.pageNumber} 页 · `;
  const detail = withDetail && issue.detail ? `：${issue.detail}` : "";
  return where + PDF_AUDIT_TEXT[issue.code] + detail;
}

/**
 * 一组 `.pdf` 体检结论的头部那句，给"打开时提示"用。
 *
 * @param issues 全部结论，顺序按引擎报出的顺序。
 * @returns 空数组时那句"没发现问题"；否则点名红色（可能打不开）与黄色（改不动）各有几条。
 */
export function summarizePdfAudit(issues: PdfAuditIssue[]): string {
  if (issues.length === 0) return "结构体检没有发现问题：这份 PDF 读得动，显示文字也都可改。";
  const broken = issues.filter((issue) => issue.severity === "broken").length;
  const limited = issues.length - broken;
  const parts = [broken > 0 ? `${broken} 处可能让阅读器打不开或丢内容` : null, limited > 0 ? `${limited} 处能打开但有一部分改不了` : null].filter(Boolean);
  return `结构体检：${parts.join("，")}。`;
}
