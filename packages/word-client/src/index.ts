import type {
  ParagraphInfo,
  ProposedEdit,
  Snapshot,
  InsertPosition,
  EditScope,
} from "@office-copilot/shared";
import { newId } from "@office-copilot/shared";

/**
 * Thin, Promise-style wrapper around the Word JavaScript API.
 *
 * Every function opens its own Word.run so callers don't have to manage the
 * request context. For batching (e.g. applying a whole proposal), use
 * {@link applyEdits} which runs all edits in one transaction and snapshots the
 * document beforehand.
 *
 * The Office.js runtime is provided by the host (Word desktop/web) — see
 * https://learn.microsoft.com/javascript/api/word for the full surface.
 */

export interface ReplaceResult {
  count: number;
  previews: Array<[string, string]>;
}

// Office.js declares two overlapping overload sets for every insert method:
// one accepting Word.InsertLocation enum values and another accepting the
// PascalCase string literals ("Before", "After", ...). At runtime Office.js
// expects the PascalCase strings, and the enum types don't widen cleanly
// through our union — so we map once and cast at the call site.
const LOC: Record<InsertPosition, "Before" | "After" | "Start" | "End"> = {
  before: "Before",
  after: "After",
  start: "Start",
  end: "End",
};

async function selectionRange(context: Word.RequestContext): Promise<Word.Range | null> {
  const sel = context.document.getSelection();
  sel.load("text");
  await context.sync();
  return sel.text.length > 0 ? sel : null;
}

async function anchorFor(
  context: Word.RequestContext,
  position: InsertPosition,
): Promise<Word.Range> {
  if (position === "start" || position === "end") {
    return context.document.body.getRange("Whole");
  }
  const sel = await selectionRange(context);
  return sel ?? context.document.body.getRange("Whole");
}

export async function listParagraphs(): Promise<ParagraphInfo[]> {
  return Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load(["items"]);
    await context.sync();
    for (const p of paragraphs.items) {
      p.load(["text", "style"]);
    }
    await context.sync();
    return paragraphs.items.map((p, index) => ({
      index,
      text: p.text,
      style: p.style,
    }));
  });
}

export async function getSelection(): Promise<{ text: string }> {
  return Word.run(async (context) => {
    const sel = context.document.getSelection();
    sel.load("text");
    await context.sync();
    return { text: sel.text };
  });
}

export async function replaceText(opts: {
  find: string;
  replace: string;
  regex?: boolean;
  scope?: EditScope;
}): Promise<ReplaceResult> {
  const { find, replace, regex = false, scope = "body" } = opts;

  if (regex) {
    // Word's native search does not speak JavaScript regex. We support the
    // literal path only in this first cut; callers should surface this as a
    // friendly error rather than silently doing the wrong thing.
    throw new Error("regex 替换暂未支持，请使用普通文字查找。");
  }

  return Word.run(async (context) => {
    const target: Word.Range =
      scope === "selection"
        ? context.document.getSelection()
        : context.document.body.getRange("Whole");

    const matches = target.search(find, { matchCase: false, ignorePunct: false } as Word.SearchOptions);
    matches.load(["items"]);
    await context.sync();

    const previews: Array<[string, string]> = [];
    for (const range of matches.items) {
      range.load("text");
      await context.sync();
      previews.push([range.text, replace]);
      range.insertText(replace, "Replace");
    }
    await context.sync();
    return { count: matches.items.length, previews };
  });
}

export async function insertParagraph(opts: {
  text: string;
  position: InsertPosition;
  style?: string;
}): Promise<void> {
  const { text, position, style } = opts;
  return Word.run(async (context) => {
    let para: Word.Paragraph;
    if (position === "start" || position === "end") {
      para = context.document.body.insertParagraph(text, LOC[position] as "Start" | "End");
    } else {
      const anchor = await anchorFor(context, position);
      para = anchor.insertParagraph(text, LOC[position] as "Before" | "After");
    }
    if (style) {
      para.style = style;
    }
    await context.sync();
  });
}

export async function insertTable(opts: {
  rows: string[][];
  headerRow?: boolean;
  position?: InsertPosition;
}): Promise<void> {
  const { rows, headerRow = true, position = "end" } = opts;
  if (rows.length === 0) throw new Error("表格至少需要一行。");
  const colCount = rows[0]!.length;
  if (colCount === 0) throw new Error("表格至少需要一列。");

  return Word.run(async (context) => {
    let table: Word.Table;
    if (position === "start" || position === "end") {
      table = context.document.body.insertTable(
        rows.length,
        colCount,
        LOC[position] as "Start" | "End",
        rows,
      );
    } else {
      const anchor = await anchorFor(context, position);
      table = anchor.insertTable(
        rows.length,
        colCount,
        LOC[position] as "Before" | "After",
        rows,
      );
    }
    if (headerRow) {
      const firstRow = table.rows.getFirst();
      firstRow.load("font");
      await context.sync();
      firstRow.font.bold = true;
    }
    await context.sync();
  });
}

export async function insertImage(opts: {
  base64: string;
  width?: number;
  height?: number;
}): Promise<void> {
  const { base64, width, height } = opts;
  return Word.run(async (context) => {
    const sel = await selectionRange(context);
    const range = sel ?? context.document.body.getRange("End");
    const pic = range.insertInlinePictureFromBase64(base64, "End");
    if (typeof width === "number") pic.width = width;
    if (typeof height === "number") pic.height = height;
    await context.sync();
  });
}

/** Capture the body OOXML so we can restore it on undo. */
export async function getSnapshot(): Promise<Snapshot> {
  return Word.run(async (context) => {
    const range = context.document.body.getRange("Whole");
    const ooxml = range.getOoxml();
    await context.sync();
    return { id: newId("snap"), ooxml: ooxml.value, createdAt: new Date().toISOString() };
  });
}

export async function restoreSnapshot(snapshot: Snapshot): Promise<void> {
  return Word.run(async (context) => {
    const range = context.document.body.getRange("Whole");
    range.insertOoxml(snapshot.ooxml, "Replace");
    await context.sync();
  });
}

/**
 * Apply a batch of edits in one Word.run and return the snapshot taken before.
 * Replace-text edits are enriched with `previews` matching the live document.
 */
export async function applyEdits(edits: ProposedEdit[]): Promise<Snapshot> {
  return Word.run(async (context) => {
    const bodyRange = context.document.body.getRange("Whole");
    const ooxml = bodyRange.getOoxml();
    await context.sync();
    const snapshot: Snapshot = {
      id: newId("snap"),
      ooxml: ooxml.value,
      createdAt: new Date().toISOString(),
    };

    for (const edit of edits) {
      switch (edit.type) {
        case "replace_text": {
          const target: Word.Range =
            edit.scope === "selection"
              ? context.document.getSelection()
              : context.document.body.getRange("Whole");
          const matches = target.search(edit.find, {
            matchCase: false,
            ignorePunct: false,
          } as Word.SearchOptions);
          matches.load(["items"]);
          await context.sync();
          if (matches.items.length === 0) {
            throw new Error(`未找到文字：${edit.find}`);
          }
          for (const r of matches.items) {
            r.insertText(edit.replace, "Replace");
          }
          await context.sync();
          break;
        }
        case "insert_paragraph": {
          const para =
            edit.position === "start" || edit.position === "end"
              ? context.document.body.insertParagraph(
                  edit.text,
                  LOC[edit.position] as "Start" | "End",
                )
              : (await anchorFor(context, edit.position)).insertParagraph(
                  edit.text,
                  LOC[edit.position] as "Before" | "After",
                );
          if (edit.style) para.style = edit.style;
          await context.sync();
          break;
        }
        case "insert_table": {
          const rows = edit.rows;
          const cols = rows[0]!.length;
          const table = context.document.body.insertTable(rows.length, cols, "End", rows);
          if (edit.headerRow) {
            const firstRow = table.rows.getFirst();
            firstRow.load("font");
            await context.sync();
            firstRow.font.bold = true;
          }
          await context.sync();
          break;
        }
        case "insert_image": {
          const b64 = edit.dataUrl.includes(",")
            ? edit.dataUrl.slice(edit.dataUrl.indexOf(",") + 1)
            : edit.dataUrl;
          const range = context.document.body;
          const pic = range.insertInlinePictureFromBase64?.(b64, "End") ?? context.document.body.getRange("End").insertInlinePictureFromBase64(b64, "After");
          if (edit.width) pic.width = edit.width;
          if (edit.height) pic.height = edit.height;
          await context.sync();
          break;
        }
      }
    }
    return snapshot;
  });
}

export const wordClient = {
  listParagraphs,
  getSelection,
  replaceText,
  insertParagraph,
  insertTable,
  insertImage,
  getSnapshot,
  restoreSnapshot,
  applyEdits,
};
