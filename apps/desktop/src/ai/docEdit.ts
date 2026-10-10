import { canonColor } from "@office-copilot/docx-io";
import type { DiagramNodeSpec, DiagramShape, DiagramSpec } from "./diagramSpec.js";

/**
 * Pure document-edit model shared by the AI agent, the TipTap adapter and the
 * history panel.
 *
 * The AI speaks in *block indexes* (`AiOp`); those resolve once into *content
 * addresses* (`CoreOp`) that stay valid after unrelated edits — which is what
 * makes undoing a single old modification possible. Every CoreOp produces the
 * inverse that undoes it, so a modification record can be rolled back on demand.
 */

export type BlockType = "paragraph" | "heading" | "table" | "image" | "other";

export interface StylePatch {
  headingLevel?: number | null;
  align?: "left" | "center" | "right" | "justify" | null;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  fontSize?: string | null;
  /** 中文字体（OOXML `w:eastAsia`）。 */
  fontFamily?: string | null;
  /** 西文字体（`w:ascii`/`w:hAnsi`）；缺省 = 与中文字体相同。 */
  fontFamilyWest?: string | null;
  /** 字间距，CSS px 字符串，与 `fontSize` 同单位。 */
  letterSpacing?: string | null;
  color?: string | null;
  /** 高亮/底纹色，canonical `#rrggbb`。 */
  highlight?: string | null;
  lineHeight?: string | null;
  spacingBefore?: number | null;
  spacingAfter?: number | null;
  indent?: number | null;
  /** 右缩进（Word 段落对话框的"右侧"），单位磅。 */
  indentRight?: number | null;
  indentFirst?: number | null;
}

/**
 * 段落级的数值格式键：单位一律磅（pt），docx-io 在进出口换算成 twips。
 * 这条清单同时被 `parseStylePatch` 的取值校验和 `docBlocks` 的读写用，
 * 加一个字段只改这里一次。
 */
export const PARAGRAPH_NUM_KEYS = [
  "spacingBefore",
  "spacingAfter",
  "indent",
  "indentRight",
  "indentFirst",
] as const;

export type ParagraphNumKey = (typeof PARAGRAPH_NUM_KEYS)[number];

export const STYLE_KEYS: readonly (keyof StylePatch)[] = [
  "headingLevel",
  "align",
  "lineHeight",
  "spacingBefore",
  "spacingAfter",
  "indent",
  "indentRight",
  "indentFirst",
  "bold",
  "italic",
  "underline",
  "fontSize",
  "fontFamily",
  "fontFamilyWest",
  "letterSpacing",
  "color",
  "highlight",
];

/**
 * Style keys that live on the runs, not on the block. A block whose runs
 * disagree has no single value for them, so it lists them in `mixed` instead and
 * a block-wide write would silently flatten, say, one red word.
 */
export const CHAR_STYLE_KEYS = [
  "bold",
  "italic",
  "underline",
  "fontSize",
  "fontFamily",
  "fontFamilyWest",
  "letterSpacing",
  "color",
  "highlight",
] as const;
export type CharStyleKey = (typeof CHAR_STYLE_KEYS)[number];

export const CHAR_STYLE_LABELS: Record<CharStyleKey, string> = {
  bold: "加粗",
  italic: "斜体",
  underline: "下划线",
  fontSize: "字号",
  fontFamily: "字体",
  fontFamilyWest: "西文字体",
  letterSpacing: "字间距",
  color: "字色",
  highlight: "底色",
};

export interface Block {
  index: number;
  type: BlockType;
  /** Paragraph/heading text, table cells joined by " | ", image alt text. */
  text: string;
  /** Only for tables. */
  rows?: string[][];
  /** Only for images — kept so undo can re-insert the picture. */
  src?: string;
  style: StylePatch;
  /**
   * Run-level keys the block's own runs disagree on; `style` deliberately omits
   * them so a block-wide write cannot flatten a bold or coloured word.
   */
  mixed?: CharStyleKey[];
}

/** Identifies a block by its content, so it survives unrelated later edits. */
export type Address =
  | { by: "blockText"; text: string; nth: number }
  | { by: "table"; firstRow: string; rowCount: number; nth: number }
  | { by: "image"; alt: string; nth: number };

export type Anchor =
  | { at: "start"; after?: undefined }
  | { at: "end"; after?: undefined }
  | { at: "after"; after: Address };

export type NewBlock =
  | { type: "paragraph"; text: string; style?: StylePatch }
  | { type: "table"; rows: string[][]; headerRow?: boolean }
  | { type: "image"; src: string; alt: string; width: number; height: number };

export type CoreOp =
  | { kind: "setText"; addr: Address; text: string }
  | { kind: "setStyle"; addr: Address; style: StylePatch }
  | { kind: "insertBlock"; anchor: Anchor; block: NewBlock }
  | { kind: "deleteBlock"; addr: Address };

/** What the model is allowed to ask for. Indexes refer to the outline it was shown. */
export type AiOp =
  | { op: "setText"; index: number; text: string }
  | { op: "setStyle"; index: number; style: StylePatch }
  | { op: "insertParagraph"; after: number; text: string; style?: StylePatch }
  | { op: "insertTable"; after: number; rows: string[][]; headerRow?: boolean }
  | { op: "insertDiagram"; after: number; diagram: DiagramSpec; alt?: string }
  | { op: "deleteBlock"; index: number }
  | { op: "replaceAll"; find: string; replace: string };

export interface Proposal {
  summary: string;
  ops: AiOp[];
}

export type DiagramResolver = (spec: DiagramSpec, alt: string) => {
  src: string;
  width: number;
  height: number;
};

/* ------------------------------------------------------------------ */
/* Outline for the model                                               */
/* ------------------------------------------------------------------ */

const OUTLINE_PREVIEW = 160;

/** Compact, model-facing view of the document. `[n]` == block index. */
export function buildOutline(blocks: Block[]): string {
  if (!blocks.length) return "(空文档)";
  return blocks
    .map((b) => {
      const head =
        b.type === "heading"
          ? `H${b.style.headingLevel ?? 1}`
          : b.type === "table"
            ? `TABLE ${b.rows?.length ?? 0}x${b.rows?.[0]?.length ?? 0}`
            : b.type === "image"
              ? "IMAGE"
              : "P";
      const text = b.text.length > OUTLINE_PREVIEW ? `${b.text.slice(0, OUTLINE_PREVIEW)}…` : b.text;
      return `[${b.index}] ${head} ${JSON.stringify(text)}${describeStyle(b.style)}`;
    })
    .join("\n");
}

function describeStyle(s: StylePatch): string {
  const bits: string[] = [];
  if (s.align && s.align !== "left") bits.push(`align=${s.align}`);
  if (s.fontSize) bits.push(`size=${s.fontSize}`);
  if (s.fontFamily) bits.push(`font=${s.fontFamily.split(",")[0]?.trim()}`);
  if (s.fontFamilyWest) bits.push(`fontWest=${s.fontFamilyWest}`);
  if (s.letterSpacing) bits.push(`spacing=${s.letterSpacing}`);
  if (s.color) bits.push(`color=${s.color}`);
  if (s.highlight) bits.push(`highlight=${s.highlight}`);
  if (s.italic) bits.push("italic");
  if (s.underline) bits.push("underline");
  if (s.lineHeight) bits.push(`lineHeight=${s.lineHeight}`);
  if (s.indentFirst != null) bits.push(`firstIndent=${s.indentFirst}`);
  return bits.length ? ` {${bits.join(", ")}}` : "";
}

/* ------------------------------------------------------------------ */
/* Validation / coercion                                              */
/* ------------------------------------------------------------------ */

function isObj(v: unknown): v is Record<string, unknown> {
  return typeof v === "object" && v !== null && !Array.isArray(v);
}

function str(v: unknown): string | null {
  return typeof v === "string" ? v : null;
}

const ALIGN = ["left", "center", "right", "justify"] as const;

type NumStyleKey = ParagraphNumKey;
type StrStyleKey =
  | "fontSize"
  | "fontFamily"
  | "fontFamilyWest"
  | "letterSpacing"
  | "color"
  | "highlight";
type BoolStyleKey = "bold" | "italic" | "underline";

export function parseStylePatch(raw: unknown): StylePatch | null {
  if (!isObj(raw)) return null;
  const out: StylePatch = {};
  const set = (key: keyof StylePatch, value: unknown): void => {
    (out as Record<string, unknown>)[key] = value;
  };

  for (const key of STYLE_KEYS) {
    if (!(key in raw)) continue;
    const v = raw[key];
    if (v === null) {
      set(key, null);
      continue;
    }
    if (key === "headingLevel") {
      if (typeof v === "number" && v >= 1 && v <= 6) set(key, Math.round(v));
      continue;
    }
    if (key === "align") {
      if ((ALIGN as readonly string[]).includes(String(v))) set(key, v);
      continue;
    }
    if (isBoolStyleKey(key)) {
      if (typeof v === "boolean") set(key, v);
      continue;
    }
    if (key === "lineHeight") {
      if (typeof v === "number" && Number.isFinite(v)) set(key, String(v));
      else if (typeof v === "string" && v) set(key, v);
      continue;
    }
    if (isNumStyleKey(key)) {
      if (typeof v === "number" && Number.isFinite(v)) set(key, v);
      continue;
    }
    if (COLOR_STYLE_KEYS.has(key)) {
      const c = canonColor(v);
      if (c) set(key, c);
      continue;
    }
    if (isStrStyleKey(key) && typeof v === "string" && v) set(key, v);
  }
  return Object.keys(out).length ? out : null;
}

function isNumStyleKey(k: keyof StylePatch): k is NumStyleKey {
  return (PARAGRAPH_NUM_KEYS as readonly string[]).includes(k);
}

function isBoolStyleKey(k: keyof StylePatch): k is BoolStyleKey {
  return k === "bold" || k === "italic" || k === "underline";
}

function isStrStyleKey(k: keyof StylePatch): k is StrStyleKey {
  return (
    k === "fontSize" ||
    k === "fontFamily" ||
    k === "fontFamilyWest" ||
    k === "letterSpacing" ||
    k === "color" ||
    k === "highlight"
  );
}

export interface CoercionResult {
  proposal: Proposal | null;
  errors: string[];
}

/** Turn whatever the model produced into a validated proposal against `blockCount`. */
export function coerceProposal(raw: unknown, blockCount: number): CoercionResult {
  const errors: string[] = [];
  if (!isObj(raw)) return { proposal: null, errors: ["提案不是有效的 JSON 对象。"] };

  const summary = (str(raw.summary) ?? str(raw.说明) ?? str(raw.title) ?? "AI 修改").trim();
  const opsRaw = raw.ops ?? raw.operations ?? raw.操作;
  if (!Array.isArray(opsRaw)) return { proposal: null, errors: ["提案缺少 ops 数组。"] };

  const ops: AiOp[] = [];
  opsRaw.forEach((entry, i) => {
    const parsed = coerceOp(entry, blockCount);
    if (parsed.error) errors.push(`第 ${i + 1} 条操作被忽略：${parsed.error}`);
    else if (parsed.op) ops.push(parsed.op);
  });

  if (!ops.length) errors.push("没有任何可执行的修改。");
  return { proposal: ops.length ? { summary: summary || "AI 修改", ops } : null, errors };
}

function coerceOp(entry: unknown, blockCount: number): { op?: AiOp; error?: string } {
  if (!isObj(entry)) return { error: "格式不正确" };
  const kind = str(entry.op) ?? str(entry.action) ?? str(entry.kind) ?? "";

  const inRange = (v: unknown, allowEnd: boolean): v is number =>
    typeof v === "number" &&
    Number.isInteger(v) &&
    v >= -1 &&
    v <= (allowEnd ? blockCount : blockCount - 1);

  switch (kind) {
    case "setText": {
      if (!inRange(entry.index, false)) return { error: "setText 的 index 超出范围" };
      const text = str(entry.text);
      if (!text) return { error: "setText 缺少 text" };
      return { op: { op: "setText", index: entry.index, text } };
    }
    case "setStyle": {
      if (!inRange(entry.index, false)) return { error: "setStyle 的 index 超出范围" };
      const style = parseStylePatch(entry.style ?? entry.format);
      if (!style) return { error: "setStyle 没有可识别的样式字段" };
      return { op: { op: "setStyle", index: entry.index, style } };
    }
    case "insertParagraph": {
      const after = entry.after ?? entry.index;
      if (!inRange(after, true)) return { error: "insertParagraph 的 after 超出范围" };
      const text = str(entry.text);
      if (!text) return { error: "insertParagraph 缺少 text" };
      const style = parseStylePatch(entry.style);
      return { op: { op: "insertParagraph", after, text, ...(style ? { style } : {}) } };
    }
    case "insertTable": {
      const after = entry.after ?? entry.index;
      if (!inRange(after, true)) return { error: "insertTable 的 after 超出范围" };
      const rows = toMatrix(entry.rows);
      if (!rows) return { error: "insertTable 的 rows 必须是等宽的二维字符串数组" };
      return {
        op: {
          op: "insertTable",
          after,
          rows,
          headerRow: entry.headerRow === undefined ? true : Boolean(entry.headerRow),
        },
      };
    }
    case "insertDiagram": {
      const after = entry.after ?? entry.index;
      if (!inRange(after, true)) return { error: "insertDiagram 的 after 超出范围" };
      const diagram = parseDiagramSpec(entry.diagram ?? entry.drawing ?? entry.图);
      if (!diagram) return { error: "insertDiagram 的 diagram 需要非空的 nodes 数组" };
      return {
        op: {
          op: "insertDiagram",
          after,
          diagram,
          alt: str(entry.alt) ?? diagram.title ?? "示意图",
        },
      };
    }
    case "deleteBlock": {
      if (!inRange(entry.index, false)) return { error: "deleteBlock 的 index 超出范围" };
      return { op: { op: "deleteBlock", index: entry.index } };
    }
    case "replaceAll": {
      const find = str(entry.find);
      if (!find) return { error: "replaceAll 缺少 find" };
      return { op: { op: "replaceAll", find, replace: str(entry.replace) ?? "" } };
    }
    default:
      return { error: `未知操作类型 "${kind}"` };
  }
}

function toMatrix(raw: unknown): string[][] | null {
  if (!Array.isArray(raw) || raw.length === 0) return null;
  const rows: string[][] = [];
  let cols = 0;
  for (const r of raw) {
    if (!Array.isArray(r) || r.length === 0) return null;
    const row = r.map((c) => (c === null || c === undefined ? "" : String(c)));
    if (cols === 0) cols = row.length;
    else if (row.length !== cols) return null;
    rows.push(row);
  }
  return rows.length ? rows : null;
}

/** Loose parsing of the diagram spec; strict shape lives in diagramSpec.ts. */
function parseDiagramSpec(raw: unknown): DiagramSpec | null {
  if (!isObj(raw)) return null;
  const nodesRaw = raw.nodes ?? raw.节点;
  if (!Array.isArray(nodesRaw) || nodesRaw.length === 0) return null;
  const nodes: DiagramNodeSpec[] = nodesRaw
    .map((n, i): DiagramNodeSpec => {
      const id = (isObj(n) && str(n.id)) || `n${i + 1}`;
      const label = (isObj(n) ? str(n.label) ?? str(n.text) : String(n ?? "")) || id;
      const shapeRaw = isObj(n) ? str(n.shape) : null;
      const shape: DiagramShape =
        shapeRaw === "diamond" || shapeRaw === "round" || shapeRaw === "box" ? shapeRaw : "box";
      return { id, label, shape };
    })
    .filter((n, i, arr) => arr.findIndex((x) => x.id === n.id) === i);
  if (!nodes.length) return null;

  const ids = new Set(nodes.map((n) => n.id));
  const edgesRaw = raw.edges ?? raw.连线;
  const edges = Array.isArray(edgesRaw)
    ? edgesRaw.flatMap((e) => {
        if (!isObj(e)) return [];
        const from = str(e.from);
        const to = str(e.to);
        if (!from || !to || !ids.has(from) || !ids.has(to)) return [];
        const label = str(e.label) ?? str(e.文字);
        return [{ from, to, ...(label ? { label } : {}) }];
      })
    : [];

  const layoutRaw = str(raw.layout);
  const layout = layoutRaw === "flow-h" || layoutRaw === "grid" ? layoutRaw : "flow-v";
  const title = str(raw.title) ?? str(raw.标题);
  return {
    layout,
    ...(title ? { title } : {}),
    nodes,
    edges,
    ...(typeof raw.width === "number" ? { width: raw.width } : {}),
  };
}

/* ------------------------------------------------------------------ */
/* Addressing                                                           */
/* ------------------------------------------------------------------ */

function tableKey(rows: string[][]): { firstRow: string; rowCount: number } {
  return { firstRow: (rows[0] ?? []).join(" | "), rowCount: rows.length };
}

function isTextual(b: Block): boolean {
  return b.type === "paragraph" || b.type === "heading";
}

function countMatching(blocks: Block[], before: number, pred: (b: Block) => boolean): number {
  let n = 0;
  for (let i = 0; i < before; i++) {
    const b = blocks[i];
    if (b && pred(b)) n += 1;
  }
  return n;
}

/** Address of the block at `index` within the snapshot, with nth-occurrence disambiguation. */
export function addressOf(blocks: Block[], index: number): Address | null {
  const b = blocks[index];
  if (!b) return null;
  if (b.type === "table") {
    const key = tableKey(b.rows ?? []);
    return {
      by: "table",
      ...key,
      nth: countMatching(blocks, index, (o) => o.type === "table" && matchesTable(o, key.firstRow, key.rowCount)),
    };
  }
  if (b.type === "image") {
    return {
      by: "image",
      alt: b.text,
      nth: countMatching(blocks, index, (o) => o.type === "image" && o.text === b.text),
    };
  }
  return {
    by: "blockText",
    text: b.text,
    nth: countMatching(blocks, index, (o) => isTextual(o) && o.text === b.text),
  };
}

function matchesTable(b: Block, firstRow: string, rowCount: number): boolean {
  if (b.type !== "table") return false;
  const key = tableKey(b.rows ?? []);
  return key.firstRow === firstRow && key.rowCount === rowCount;
}

export function matchesAddress(b: Block, addr: Address): boolean {
  switch (addr.by) {
    case "blockText":
      return isTextual(b) && b.text === addr.text;
    case "table":
      return matchesTable(b, addr.firstRow, addr.rowCount);
    case "image":
      return b.type === "image" && b.text === addr.alt;
  }
}

/** Position of the nth match, or -1. */
export function findAddress(blocks: Block[], addr: Address): number {
  let seen = 0;
  for (let i = 0; i < blocks.length; i++) {
    const b = blocks[i];
    if (b && matchesAddress(b, addr)) {
      if (seen === addr.nth) return i;
      seen += 1;
    }
  }
  return -1;
}

function findAnchor(blocks: Block[], anchor: Anchor): number {
  if (anchor.at === "start") return -1;
  if (anchor.at === "end") return blocks.length - 1;
  return findAddress(blocks, anchor.after);
}

/* ------------------------------------------------------------------ */
/* Style helpers                                                        */
/* ------------------------------------------------------------------ */

function patchStyle(current: StylePatch, patch: StylePatch): StylePatch {
  const next: StylePatch = { ...current };
  for (const key of STYLE_KEYS) {
    if (!(key in patch)) continue;
    const v = patch[key];
    if (v === null || v === undefined) delete next[key];
    else (next as Record<string, unknown>)[key] = v;
  }
  return next;
}

/**
 * Style keys holding a CSS colour. They compare through `canonColor`, because
 * the same colour reaches us as `#C00000` from a template, `#c00000` from a
 * .docx and `rgb(192, 0, 0)` from the DOM.
 */
const COLOR_STYLE_KEYS = new Set<string>(["color", "highlight"]);

function canonStyleValue(key: keyof StylePatch, value: unknown): string | number | boolean | undefined {
  if (value === null || value === undefined) return undefined;
  if (COLOR_STYLE_KEYS.has(key)) return canonColor(value) ?? undefined;
  return value as string | number | boolean;
}

/**
 * Would applying `patch` actually change anything? Templates ask this before
 * emitting an op so running the same template twice is a no-op — the same
 * question `commitOps` answers with an empty diff, answered once up front.
 */
export function styleUnchanged(current: StylePatch, patch: StylePatch): boolean {
  const next = patchStyle(current, patch);
  for (const key of STYLE_KEYS) {
    const a = canonStyleValue(key, current[key]);
    const b = canonStyleValue(key, next[key]);
    if (a === undefined && b === undefined) continue;
    if (a === undefined || b === undefined) return false;
    if (typeof a === "number" && typeof b === "number") {
      if (a !== b) return false;
    } else if (String(a) !== String(b)) {
      return false;
    }
  }
  return true;
}

/** Previous values of exactly the keys `patch` touches — that is the inverse. */
function prevStyleFor(current: StylePatch, patch: StylePatch): StylePatch {
  const prev: Record<string, unknown> = {};
  for (const key of STYLE_KEYS) {
    if (!(key in patch)) continue;
    prev[key] = current[key] === undefined ? null : current[key];
  }
  return prev as StylePatch;
}

/* ------------------------------------------------------------------ */
/* The engine                                                           */
/* ------------------------------------------------------------------ */

export interface ApplyResult {
  ok: boolean;
  reason?: string;
  inverse?: CoreOp;
  insertedIndex?: number;
}

function reindex(blocks: Block[]): void {
  blocks.forEach((b, i) => {
    b.index = i;
  });
}

function blockFromNew(nb: NewBlock): Block {
  if (nb.type === "table") {
    return {
      index: 0,
      type: "table",
      text: nb.rows.map((r) => r.join(" | ")).join("\n"),
      rows: nb.rows.map((r) => [...r]),
      style: {},
    };
  }
  if (nb.type === "image") {
    return { index: 0, type: "image", text: nb.alt, src: nb.src, style: {} };
  }
  const style = nb.style ?? {};
  return {
    index: 0,
    type: style.headingLevel ? "heading" : "paragraph",
    text: nb.text,
    style: { ...style },
  };
}

function newFromBlock(b: Block): NewBlock {
  if (b.type === "table") return { type: "table", rows: b.rows ?? [[]] };
  if (b.type === "image") {
    return { type: "image", src: b.src ?? "", alt: b.text, width: 0, height: 0 };
  }
  return { type: "paragraph", text: b.text, style: { ...b.style } };
}

/**
 * Apply one CoreOp to a mutable block array. The returned inverse rolls exactly
 * this op back, assuming nothing else touched the target meanwhile.
 */
export function applyCore(blocks: Block[], op: CoreOp): ApplyResult {
  switch (op.kind) {
    case "setText": {
      const i = findAddress(blocks, op.addr);
      if (i < 0) return { ok: false, reason: "找不到目标段落" };
      const b = blocks[i]!;
      if (!isTextual(b)) return { ok: false, reason: "该块不支持文字替换" };
      const beforeText = b.text;
      blocks[i] = { ...b, text: op.text };
      const addr = addressOf(blocks, i)!;
      return { ok: true, inverse: { kind: "setText", addr, text: beforeText } };
    }
    case "setStyle": {
      const i = findAddress(blocks, op.addr);
      if (i < 0) return { ok: false, reason: "找不到目标段落" };
      const b = blocks[i]!;
      const clash = (b.mixed ?? []).filter((k) => k in op.style);
      if (clash.length) {
        return {
          ok: false,
          reason: `「${clip(b.text)}」段内${clash.map((k) => CHAR_STYLE_LABELS[k]).join("、")}只有部分文字带格式，整段设置会把它抹平；请在文档里选中那段文字再改`,
        };
      }
      const prev = prevStyleFor(b.style, op.style);
      const style = patchStyle(b.style, op.style);
      let type = b.type;
      if ("headingLevel" in op.style) {
        type = style.headingLevel ? "heading" : "paragraph";
      }
      blocks[i] = { ...b, style, type };
      return { ok: true, inverse: { kind: "setStyle", addr: addressOf(blocks, i)!, style: prev } };
    }
    case "insertBlock": {
      const at = findAnchor(blocks, op.anchor);
      if (at < 0 && op.anchor.at !== "start") return { ok: false, reason: "找不到插入位置" };
      const position = at + 1;
      blocks.splice(position, 0, blockFromNew(op.block));
      reindex(blocks);
      return {
        ok: true,
        insertedIndex: position,
        inverse: { kind: "deleteBlock", addr: addressOf(blocks, position)! },
      };
    }
    case "deleteBlock": {
      const i = findAddress(blocks, op.addr);
      if (i < 0) return { ok: false, reason: "找不到目标段落" };
      const removed = blocks[i]!;
      blocks.splice(i, 1);
      reindex(blocks);
      const anchor: Anchor =
        blocks.length === 0 || i === 0 ? { at: "start" } : { at: "after", after: addressOf(blocks, i - 1)! };
      return { ok: true, inverse: { kind: "insertBlock", anchor, block: newFromBlock(removed) } };
    }
  }
}

export interface Plan {
  /** CoreOps to execute against the live document, in order. */
  ops: CoreOp[];
  before: Block[];
  after: Block[];
  /** Inverse ops, in the order they must run to roll `ops` back. */
  inverses: CoreOp[];
  /** Reasons for ops the model asked for but that could not be honoured. */
  skipped: string[];
  ok: boolean;
}

export function emptyBlocks(): Block[] {
  return [];
}

function copyBlocks(blocks: Block[]): Block[] {
  return blocks.map((b) => ({
    ...b,
    style: { ...b.style },
    rows: b.rows ? b.rows.map((r) => [...r]) : undefined,
  }));
}

/**
 * Run already-resolved CoreOps over a copy of the snapshot, collecting the
 * inverse of each step. Used for edits the UI builds directly (selection
 * rewrites) and as the reference semantics for undo.
 */
export function runOps(start: Block[], ops: readonly CoreOp[]): RunOpsResult {
  const blocks = copyBlocks(start);
  const applied: CoreOp[] = [];
  const inverses: CoreOp[] = [];
  const skipped: string[] = [];
  for (const op of ops) {
    const r = applyCore(blocks, op);
    if (!r.ok) {
      skipped.push(`${describeCoreOp(op)}：${r.reason ?? "失败"}`);
      continue;
    }
    applied.push(op);
    if (r.inverse) inverses.push(r.inverse);
  }
  return { after: blocks, applied, skipped, inverses: inverses.reverse() };
}

export interface RunOpsResult {
  after: Block[];
  applied: CoreOp[];
  inverses: CoreOp[];
  skipped: string[];
}

/** Resolve AiOps into CoreOps and simulate them over a copy of the snapshot. */
export function planEdit(blocks: Block[], proposal: Proposal, diagrams?: DiagramResolver): Plan {
  const skipped: string[] = [];
  const working: Block[] = copyBlocks(blocks);
  const ops: CoreOp[] = [];
  const inverses: CoreOp[] = [];

  // The model was shown the outline as it was *before* any of these ops ran, so
  // its indexes mean original indexes. Inserts and deletes shift later blocks,
  // hence this translation table: pos[original] == current position (-1 == gone).
  const pos: number[] = working.map((_, i) => i);

  for (const ai of proposal.ops) {
    for (const core of resolveAiOp(ai, working, pos, diagrams, skipped)) {
      const deleting = core.kind === "deleteBlock" ? findAddress(working, core.addr) : -1;
      const r = applyCore(working, core);
      if (!r.ok) {
        skipped.push(r.reason ?? "失败");
        continue;
      }
      ops.push(core);
      if (r.inverse) inverses.push(r.inverse);

      if (core.kind === "insertBlock" && r.insertedIndex != null) {
        const insertedAt = r.insertedIndex;
        pos.forEach((p, i) => {
          if (p >= insertedAt) pos[i] = p + 1;
        });
      } else if (deleting >= 0) {
        const gone = pos.indexOf(deleting);
        if (gone >= 0) pos[gone] = -1;
        pos.forEach((p, i) => {
          if (p > deleting) pos[i] = p - 1;
        });
      }
    }
  }

  return {
    ops,
    before: blocks,
    after: working,
    inverses: inverses.reverse(),
    skipped,
    ok: ops.length > 0,
  };
}

function resolveAiOp(
  ai: AiOp,
  blocks: Block[],
  pos: readonly number[],
  diagrams: DiagramResolver | undefined,
  skipped: string[],
): CoreOp[] {
  const originalIndex = (index: number, verb: string): Block | null => {
    const cur = pos[index];
    if (cur === undefined) {
      skipped.push(`${verb}：第 ${index + 1} 段不存在`);
      return null;
    }
    if (cur < 0) {
      skipped.push(`${verb}：第 ${index + 1} 段已被本次改动删除`);
      return null;
    }
    const b = blocks[cur];
    if (!b) {
      skipped.push(`${verb}：第 ${index + 1} 段不存在`);
      return null;
    }
    return b;
  };

  const anchorAfter = (after: number): Anchor | null => {
    if (after < 0) return { at: "start" };
    if (after >= pos.length) return { at: "end" };
    const cur = pos[after]!;
    if (cur < 0) {
      skipped.push(`插入位置的第 ${after + 1} 段已被本次改动删除`);
      return null;
    }
    const addr = addressOf(blocks, cur);
    return addr ? { at: "after", after: addr } : null;
  };

  switch (ai.op) {
    // Expanded per block so every replacement stays individually undoable.
    case "replaceAll": {
      const out: CoreOp[] = [];
      for (const b of blocks) {
        if (!isTextual(b) || !b.text.includes(ai.find)) continue;
        const addr = addressOf(blocks, b.index);
        if (addr) out.push({ kind: "setText", addr, text: b.text.split(ai.find).join(ai.replace) });
      }
      if (!out.length) skipped.push(`replaceAll：文档中没有「${ai.find}」`);
      return out;
    }
    case "setText": {
      const target = originalIndex(ai.index, "setText");
      if (!target) return [];
      const addr = addressOf(blocks, target.index);
      return addr ? [{ kind: "setText", addr, text: ai.text }] : [];
    }
    case "setStyle": {
      const target = originalIndex(ai.index, "setStyle");
      if (!target) return [];
      const addr = addressOf(blocks, target.index);
      return addr ? [{ kind: "setStyle", addr, style: ai.style }] : [];
    }
    case "insertParagraph": {
      const anchor = anchorAfter(ai.after);
      return anchor ? [{ kind: "insertBlock", anchor, block: { type: "paragraph", text: ai.text, ...(ai.style ? { style: ai.style } : {}) } }] : [];
    }
    case "insertTable": {
      const anchor = anchorAfter(ai.after);
      return anchor ? [{ kind: "insertBlock", anchor, block: { type: "table", rows: ai.rows, headerRow: ai.headerRow } }] : [];
    }
    case "insertDiagram": {
      const anchor = anchorAfter(ai.after);
      if (!anchor) return [];
      if (!diagrams) {
        skipped.push("insertDiagram：当前环境无法绘制图片");
        return [];
      }
      let rendered: { src: string; width: number; height: number };
      try {
        rendered = diagrams(ai.diagram, ai.alt ?? "示意图");
      } catch (e) {
        skipped.push(`insertDiagram：绘制失败（${e instanceof Error ? e.message : String(e)}）`);
        return [];
      }
      return [
        {
          kind: "insertBlock",
          anchor,
          block: {
            type: "image",
            src: rendered.src,
            alt: ai.alt ?? ai.diagram.title ?? "示意图",
            width: rendered.width,
            height: rendered.height,
          },
        },
      ];
    }
    case "deleteBlock": {
      const target = originalIndex(ai.index, "deleteBlock");
      if (!target) return [];
      const addr = addressOf(blocks, target.index);
      return addr ? [{ kind: "deleteBlock", addr }] : [];
    }
  }
}

export interface RevertResult extends RunOpsResult {
  ok: boolean;
  failed: string[];
}

/**
 * Roll a recorded modification back out of `from` (the document as it is now).
 * Each inverse is matched by content, so unrelated edits made afterwards do not
 * break the undo; a target the user has since changed is reported rather than
 * guessed at.
 */
export function revert(from: Block[], inverses: CoreOp[]): RevertResult {
  const r = runOps(from, inverses);
  return {
    ...r,
    ok: r.skipped.length === 0,
    failed: r.skipped,
  };
}

/* ------------------------------------------------------------------ */
/* Human-readable labels                                               */
/* ------------------------------------------------------------------ */

function clip(s: string): string {
  return s.length > 24 ? `${s.slice(0, 24)}…` : s;
}

function addrLabel(addr: Address): string {
  switch (addr.by) {
    case "blockText":
      return clip(addr.text) || "空段落";
    case "table":
      return `表格「${clip(addr.firstRow)}」`;
    case "image":
      return `图片「${clip(addr.alt)}」`;
  }
}

export function describeCoreOp(op: CoreOp): string {
  switch (op.kind) {
    case "setText":
      return `改写「${addrLabel(op.addr)}」`;
    case "setStyle":
      return `调整「${addrLabel(op.addr)}」样式（${Object.keys(op.style).join("/")}）`;
    case "insertBlock":
      return `插入${op.block.type === "table" ? "表格" : op.block.type === "image" ? "图片" : "段落"}`;
    case "deleteBlock":
      return `删除「${addrLabel(op.addr)}」`;
  }
}

export function describeAiOp(ai: AiOp): string {
  switch (ai.op) {
    case "setText":
      return `重写第 ${ai.index + 1} 段 → ${clip(ai.text)}`;
    case "setStyle":
      return `第 ${ai.index + 1} 段样式：${Object.entries(ai.style)
        .map(([k, v]) => `${k}=${v ?? "清除"}`)
        .join(", ")}`;
    case "insertParagraph":
      return `新增段落：${clip(ai.text)}`;
    case "insertTable":
      return `新增表格 ${ai.rows.length}×${ai.rows[0]?.length ?? 0}`;
    case "insertDiagram":
      return `新增示意图：${ai.diagram.title ?? `${ai.diagram.nodes.length} 个节点`}`;
    case "deleteBlock":
      return `删除第 ${ai.index + 1} 段`;
    case "replaceAll":
      return `全局替换「${ai.find}」→「${ai.replace}」`;
  }
}
