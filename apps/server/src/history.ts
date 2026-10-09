import Database from "better-sqlite3";
import fs from "node:fs";
import path from "node:path";
import { config } from "./config.js";

/**
 * Per-session history: instructions, proposals, and OOXML snapshots. The
 * task pane only needs the most recent snapshot for undo, but persisting the
 * full record lets us later add a "versions" panel and replay.
 */

export interface HistoryRecord {
  id: number;
  session_id: string;
  instruction: string;
  summary: string;
  ooxml_before: string | null;
  ooxml_after: string | null;
  created_at: string;
}

let db: Database.Database | null = null;

export function getDb(): Database.Database {
  if (db) return db;
  fs.mkdirSync(config.dataDir, { recursive: true });
  db = new Database(path.join(config.dataDir, "history.sqlite"));
  db.pragma("journal_mode = WAL");
  db.exec(`
    CREATE TABLE IF NOT EXISTS history (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      session_id TEXT NOT NULL,
      instruction TEXT NOT NULL,
      summary TEXT NOT NULL,
      ooxml_before TEXT,
      ooxml_after TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now'))
    );
    CREATE INDEX IF NOT EXISTS idx_history_session ON history(session_id, id DESC);
  `);
  return db;
}

export function recordHistory(input: {
  sessionId: string;
  instruction: string;
  summary: string;
  ooxmlBefore?: string | null;
  ooxmlAfter?: string | null;
}): HistoryRecord {
  const stmt = getDb().prepare(
    `INSERT INTO history (session_id, instruction, summary, ooxml_before, ooxml_after)
     VALUES (?, ?, ?, ?, ?)
     RETURNING *`,
  );
  return stmt.get(
    input.sessionId,
    input.instruction,
    input.summary,
    input.ooxmlBefore ?? null,
    input.ooxmlAfter ?? null,
  ) as HistoryRecord;
}

export function listHistory(sessionId: string, limit = 50): HistoryRecord[] {
  return getDb()
    .prepare(`SELECT * FROM history WHERE session_id = ? ORDER BY id DESC LIMIT ?`)
    .all(sessionId, limit) as HistoryRecord[];
}
