/**
 * Diff used by the modification-history panel: block-level rows plus an inline
 * word/character diff for the blocks that changed. Written from scratch (LCS) so
 * the app gains no new dependency.
 */

export interface InlinePart {
  t: "eq" | "ins" | "del";
  s: string;
}

export interface DiffRow {
  kind: "eq" | "add" | "del" | "mod";
  before?: string;
  after?: string;
  inline?: InlinePart[];
}

/** Split into comparable tokens: CJK by character, latin by word. */
export function tokenize(s: string): string[] {
  const out: string[] = [];
  let buf = "";
  for (const ch of s) {
    const c = ch.codePointAt(0) ?? 0;
    const isCjk = c > 0x2e80 && !/\s/.test(ch);
    if (isCjk) {
      if (buf) {
        out.push(buf);
        buf = "";
      }
      out.push(ch);
    } else if (/\s/.test(ch)) {
      if (buf) {
        out.push(buf);
        buf = "";
      }
      out.push(ch);
    } else {
      buf += ch;
    }
  }
  if (buf) out.push(buf);
  return out;
}

function lcsOps<T>(a: readonly T[], b: readonly T[]): Array<{ t: "eq" | "ins" | "del"; v: T }> {
  const n = a.length;
  const m = b.length;
  // dp[i][j] = length of LCS of a[i:], b[j:]
  const dp: Uint32Array[] = Array.from({ length: n + 1 }, () => new Uint32Array(m + 1));
  for (let i = n - 1; i >= 0; i--) {
    const row = dp[i]!;
    const next = dp[i + 1]!;
    for (let j = m - 1; j >= 0; j--) {
      row[j] = a[i] === b[j] ? next[j + 1]! + 1 : Math.max(next[j]!, row[j + 1]!);
    }
  }
  const ops: Array<{ t: "eq" | "ins" | "del"; v: T }> = [];
  let i = 0;
  let j = 0;
  while (i < n && j < m) {
    if (a[i] === b[j]) {
      ops.push({ t: "eq", v: a[i]! });
      i++;
      j++;
    } else if (dp[i + 1]![j]! >= dp[i]![j + 1]!) {
      ops.push({ t: "del", v: a[i]! });
      i++;
    } else {
      ops.push({ t: "ins", v: b[j]! });
      j++;
    }
  }
  while (i < n) ops.push({ t: "del", v: a[i++]! });
  while (j < m) ops.push({ t: "ins", v: b[j++]! });
  return ops;
}

const CELL_BUDGET = 400_000;

/** Trim the shared prefix/suffix first, then run LCS on what is left. */
function sequenceDiff<T>(a: readonly T[], b: readonly T[]): Array<{ t: "eq" | "ins" | "del"; v: T }> {
  let start = 0;
  while (start < a.length && start < b.length && a[start] === b[start]) start++;
  let endA = a.length;
  let endB = b.length;
  while (endA > start && endB > start && a[endA - 1] === b[endB - 1]) {
    endA--;
    endB--;
  }
  const head = a.slice(0, start).map((v) => ({ t: "eq" as const, v }));
  const tail = a.slice(endA).map((v) => ({ t: "eq" as const, v }));
  const midA = a.slice(start, endA);
  const midB = b.slice(start, endB);
  if (midA.length * midB.length > CELL_BUDGET) {
    return [
      ...head,
      ...midA.map((v) => ({ t: "del" as const, v })),
      ...midB.map((v) => ({ t: "ins" as const, v })),
      ...tail,
    ];
  }
  return [...head, ...lcsOps(midA, midB), ...tail];
}

export function inlineDiff(before: string, after: string): InlinePart[] {
  const ops = sequenceDiff(tokenize(before), tokenize(after));
  const parts: InlinePart[] = [];
  for (const op of ops) {
    const s = op.v;
    const last = parts[parts.length - 1];
    if (last && last.t === (op.t === "eq" ? "eq" : op.t)) last.s += s;
    else parts.push({ t: op.t === "eq" ? "eq" : (op.t as "ins" | "del"), s });
  }
  return parts;
}

/**
 * Block-level diff: equal blocks stay, runs of removals/additions that line up
 * are folded into a "mod" row with an inline diff, so a rewritten paragraph
 * reads as one changed line instead of two noisy ones.
 */
export function diffBlocks(before: readonly string[], after: readonly string[]): DiffRow[] {
  const ops = sequenceDiff(before, after);
  const rows: DiffRow[] = [];
  let i = 0;
  while (i < ops.length) {
    const op = ops[i]!;
    if (op.t === "eq") {
      rows.push({ kind: "eq", before: op.v, after: op.v });
      i++;
      continue;
    }
    const dels: string[] = [];
    const adds: string[] = [];
    while (i < ops.length && ops[i]!.t === "del") dels.push(ops[i++]!.v);
    while (i < ops.length && ops[i]!.t === "ins") adds.push(ops[i++]!.v);

    const pair = Math.min(dels.length, adds.length);
    for (let k = 0; k < pair; k++) {
      rows.push({
        kind: "mod",
        before: dels[k],
        after: adds[k],
        inline: inlineDiff(dels[k]!, adds[k]!),
      });
    }
    for (let k = pair; k < dels.length; k++) rows.push({ kind: "del", before: dels[k] });
    for (let k = pair; k < adds.length; k++) rows.push({ kind: "add", after: adds[k] });
  }
  return rows;
}

export function diffStats(rows: readonly DiffRow[]): { added: number; removed: number; changed: number } {
  let added = 0;
  let removed = 0;
  let changed = 0;
  for (const r of rows) {
    if (r.kind === "add") added++;
    else if (r.kind === "del") removed++;
    else if (r.kind === "mod") changed++;
  }
  return { added, removed, changed };
}

/**
 * Git-style unified lines: one row per old/new line with both line numbers,
 * `-` before `+`, and unchanged stretches collapsed into an "omit" marker.
 * A folded `mod` row expands into the old line (removed spans struck) followed
 * by the new line (inserted spans marked), like `git diff --word-diff`.
 */
export interface GitLine {
  kind: "ctx" | "add" | "del" | "omit";
  beforeNo?: number;
  afterNo?: number;
  parts: InlinePart[];
  omitted?: number;
}

const CONTEXT_LINES = 2;

export function toGitLines(
  before: readonly string[],
  after: readonly string[],
  context = CONTEXT_LINES,
): GitLine[] {
  const rows = diffBlocks(before, after);
  const raw: GitLine[] = [];
  let bn = 1;
  let an = 1;
  const one = (t: "eq" | "ins" | "del", parts: InlinePart[]): InlinePart[] =>
    parts.filter((p) => p.t === "eq" || p.t === t);

  for (const r of rows) {
    if (r.kind === "eq") {
      raw.push({ kind: "ctx", beforeNo: bn++, afterNo: an++, parts: [{ t: "eq", s: r.after ?? "" }] });
    } else if (r.kind === "del") {
      raw.push({ kind: "del", beforeNo: bn++, parts: [{ t: "del", s: r.before ?? "" }] });
    } else if (r.kind === "add") {
      raw.push({ kind: "add", afterNo: an++, parts: [{ t: "ins", s: r.after ?? "" }] });
    } else {
      const parts = r.inline ?? [];
      raw.push({ kind: "del", beforeNo: bn++, parts: one("del", parts) });
      raw.push({ kind: "add", afterNo: an++, parts: one("ins", parts) });
    }
  }

  if (!raw.some((l) => l.kind !== "ctx")) return raw;

  const keep = raw.map(() => false);
  raw.forEach((l, i) => {
    if (l.kind === "ctx") return;
    for (let j = Math.max(0, i - context); j <= Math.min(raw.length - 1, i + context); j++) keep[j] = true;
  });

  const out: GitLine[] = [];
  let omitted = 0;
  raw.forEach((l, i) => {
    if (keep[i]) {
      if (omitted) {
        out.push({ kind: "omit", parts: [], omitted });
        omitted = 0;
      }
      out.push(l);
    } else {
      omitted += 1;
    }
  });
  if (omitted) out.push({ kind: "omit", parts: [], omitted });
  return out;
}

