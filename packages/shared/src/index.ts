// Shared types and protocols used by the task pane and the server.

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
