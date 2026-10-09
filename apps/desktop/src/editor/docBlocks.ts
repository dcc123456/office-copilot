import { canonColor } from "@office-copilot/docx-io";
import type { Editor } from "@tiptap/react";
import type { Mark, Node as PMNode, NodeType, Schema } from "@tiptap/pm/model";
import type { Transaction } from "@tiptap/pm/state";
import {
  CHAR_STYLE_KEYS,
  CHAR_STYLE_LABELS,
  matchesAddress,
  type Address,
  type Anchor,
  type Block,
  type CharStyleKey,
  type CoreOp,
  type NewBlock,
  type StylePatch,
} from "../ai/docEdit.js";

/**
 * Bridge between the pure edit model (`ai/docEdit.ts`) and the live TipTap
 * document. Blocks are matched by content, never by a stored position, so an op
 * resolved at proposal time can still be applied — or rolled back — after the
 * user kept typing.
 */

const TEXTUAL = new Set(["paragraph", "heading"]);

/** textStyle attrs whose value is a CSS colour. */
const COLOR_ATTR_KEYS = new Set<string>(["color", "backgroundColor"]);

/** Character style keys carried as marks rather than `textStyle` attributes. */
const MARK_STYLE_KEYS = ["bold", "italic", "underline"] as const;
type MarkStyleKey = (typeof MARK_STYLE_KEYS)[number];
/** The remaining character style keys live inside the `textStyle` mark. */
type AttrStyleKey = Exclude<CharStyleKey, MarkStyleKey>;

function isMarkStyleKey(key: CharStyleKey): key is MarkStyleKey {
  return (MARK_STYLE_KEYS as readonly string[]).includes(key);
}

/** Where each character-level style key lives inside the `textStyle` mark. */
const TEXT_STYLE_ATTRS: Record<AttrStyleKey, string> = {
  fontSize: "fontSize",
  fontFamily: "fontFamily",
  fontFamilyWest: "fontFamilyWest",
  letterSpacing: "letterSpacing",
  color: "color",
  highlight: "backgroundColor",
};

interface Located {
  block: Block;
  /** Node span: what a delete removes and what an insert follows. */
  from: number;
  to: number;
  /** Position of the paragraph/heading itself — inside a list that is not `from`. */
  textPos: number;
  /** Run-level keys the block's own runs disagree on. */
  mixed: CharStyleKey[];
  /** Range of the block's text content, when it can be rewritten. */
  inner: { from: number; to: number } | null;
  /** The paragraph/heading node holding `inner`. */
  textNode: PMNode | null;
  /** Enclosing list span, when this block came from a list item. */
  listSpan: { from: number; to: number } | null;
}

function texts(node: PMNode): PMNode[] {
  const out: PMNode[] = [];
  node.descendants((n) => {
    if (n.isText) out.push(n);
    return !n.isText;
  });
  return out;
}

function styleOf(node: PMNode): StylePatch {
  const attrs = node.attrs as Record<string, unknown>;
  const s: StylePatch = {};
  if (node.type.name === "heading") s.headingLevel = Number(attrs.level ?? 1);
  if (typeof attrs.textAlign === "string") s.align = attrs.textAlign as StylePatch["align"];
  if (attrs.lineHeight != null) s.lineHeight = String(attrs.lineHeight);
  for (const key of ["spacingBefore", "spacingAfter", "indent", "indentFirst"] as const) {
    if (typeof attrs[key] === "number") s[key] = attrs[key];
  }
  return s;
}

function runValue(run: PMNode, key: CharStyleKey): string | boolean | null {
  // A mark is either on the run or not: the value is its *presence*, which is
  // why these keys read as true/null and never as false — "not bold" is the
  // absence of the mark, not a colour-like value a run could carry.
  if (isMarkStyleKey(key)) {
    return run.marks?.some((m) => m.type.name === key) ? true : null;
  }
  const attr = TEXT_STYLE_ATTRS[key];
  for (const m of run.marks ?? []) {
    if (m.type.name !== "textStyle") continue;
    const v = (m.attrs as Record<string, unknown>)[attr];
    if (typeof v === "string" && v) {
      // Same red, two spellings is not a disagreement between two runs.
      return COLOR_ATTR_KEYS.has(attr) ? (canonColor(v) ?? v) : v;
    }
  }
  return null;
}

function isLinkRun(run: PMNode): boolean {
  return (run.marks ?? []).some((m) => m.type.name === "link");
}

/**
 * Run-level style of a block. Keys the runs disagree on go to `mixed` and are
 * left out of `style`: a block-wide write must not flatten one red word.
 */
function charStyleOf(node: PMNode): { style: StylePatch; mixed: CharStyleKey[] } {
  const style: StylePatch = {};
  const mixed: CharStyleKey[] = [];
  const runs = texts(node);
  if (!runs.length) return { style, mixed };
  for (const key of CHAR_STYLE_KEYS) {
    // A hyperlink's colour belongs to the document, not to the run: the
    // template carries a LinkStyle, painted in the editor through a CSS
    // variable and in the file through the link run's own rPr. Left in the
    // vote, Word's blue would disagree with the sentence around it and make
    // every paragraph holding a link unstyleable.
    const voters = key === "color" ? runs.filter((r) => !isLinkRun(r)) : runs;
    if (!voters.length) {
      mixed.push(key);
      continue;
    }
    const values = voters.map((run) => runValue(run, key));
    if (values.some((v) => v !== values[0])) {
      mixed.push(key);
      continue;
    }
    const v = values[0];
    if (v !== null) (style as Record<string, unknown>)[key] = v;
  }
  return { style, mixed };
}

function tableRows(table: PMNode): { rows: string[][]; header: boolean } {
  const rows: string[][] = [];
  let header = false;
  let r = 0;
  table.forEach((row) => {
    const cells: string[] = [];
    row.forEach((cell) => {
      cells.push(cell.textContent);
      if (r === 0 && cell.type.name === "tableHeader") header = true;
    });
    r += 1;
    if (cells.length) rows.push(cells);
  });
  return { rows, header };
}

/** A paragraph holding nothing but one picture is an image block. */
function loneImage(p: PMNode): PMNode | null {
  if (p.childCount !== 1) return null;
  const only = p.firstChild;
  return only && only.type.name === "image" ? only : null;
}

function imageBlock(src: string, alt: string, index: number): Block {
  return { index, type: "image", text: alt, src, style: {} };
}

/** Flatten the document into addressable blocks; array order == outline order. */
export function locateAll(doc: PMNode): Located[] {
  const out: Located[] = [];
  let index = 0;

  const textBlock = (
    node: PMNode,
    from: number,
    to: number,
    textPos: number,
    inner: { from: number; to: number },
    listSpan: { from: number; to: number } | null,
  ): void => {
    const img = loneImage(node);
    if (img) {
      const attrs = img.attrs as Record<string, unknown>;
      out.push({
        block: imageBlock(String(attrs.src ?? ""), String(attrs.alt ?? ""), index++),
        from,
        to,
        textPos,
        mixed: [],
        inner: null,
        textNode: node,
        listSpan,
      });
      return;
    }
    const { style, mixed } = charStyleOf(node);
    out.push({
      block: {
        index: index++,
        type: node.type.name === "heading" ? "heading" : "paragraph",
        text: node.textContent,
        style: { ...styleOf(node), ...style },
        ...(mixed.length ? { mixed } : {}),
      },
      from,
      to,
      textPos,
      mixed,
      inner,
      textNode: node,
      listSpan,
    });
  };

  doc.forEach((node, pos) => {
    const name = node.type.name;

    if (name === "bulletList" || name === "orderedList") {
      const listSpan = { from: pos, to: pos + node.nodeSize };
      node.forEach((item, offset) => {
        const itemFrom = pos + 1 + offset;
        const para = item.firstChild;
        if (para && TEXTUAL.has(para.type.name)) {
          const paraFrom = itemFrom + 1;
          textBlock(
            para,
            itemFrom,
            itemFrom + item.nodeSize,
            paraFrom,
            { from: paraFrom + 1, to: paraFrom + para.nodeSize - 1 },
            listSpan,
          );
        }
      });
      return;
    }

    if (TEXTUAL.has(name)) {
      textBlock(node, pos, pos + node.nodeSize, pos, { from: pos + 1, to: pos + node.nodeSize - 1 }, null);
      return;
    }

    if (name === "table") {
      const { rows } = tableRows(node);
      out.push({
        block: { index: index++, type: "table", text: rows.map((r) => r.join(" | ")).join("\n"), rows, style: {} },
        from: pos,
        to: pos + node.nodeSize,
        textPos: pos,
        mixed: [],
        inner: null,
        textNode: node,
        listSpan: null,
      });
      return;
    }

    if (name === "image") {
      const attrs = node.attrs as Record<string, unknown>;
      out.push({
        block: imageBlock(String(attrs.src ?? ""), String(attrs.alt ?? ""), index++),
        from: pos,
        to: pos + node.nodeSize,
        textPos: pos,
        mixed: [],
        inner: null,
        textNode: node,
        listSpan: null,
      });
      return;
    }

    out.push({
      block: { index: index++, type: "other", text: node.textContent, style: {} },
      from: pos,
      to: pos + node.nodeSize,
      textPos: pos,
      mixed: [],
      inner: null,
      textNode: node,
      listSpan: null,
    });
  });

  return out;
}

export function blocksOf(doc: PMNode): Block[] {
  return locateAll(doc).map((l) => l.block);
}

function locateAddress(doc: PMNode, addr: Address): Located | null {
  let seen = 0;
  for (const l of locateAll(doc)) {
    if (matchesAddress(l.block, addr)) {
      if (seen === addr.nth) return l;
      seen += 1;
    }
  }
  return null;
}

function locateAnchor(doc: PMNode, anchor: Anchor): Located | null {
  if (anchor.at === "start") {
    const first = locateAll(doc)[0];
    if (!first) return null;
    return { ...first, from: 0, to: 0 };
  }
  if (anchor.at === "end") {
    const all = locateAll(doc);
    return all[all.length - 1] ?? null;
  }
  return locateAddress(doc, anchor.after);
}

/** Collapse newlines: one op rewrites one block, so block boundaries stay stable. */
export function normalise(text: string): string {
  return text.replace(/\s*\n\s*/g, " ").replace(/[ \t]{2,}/g, " ").trim();
}

function marksOf(node: PMNode | null) {
  const first = node ? texts(node)[0] : undefined;
  return first?.marks?.length ? first.marks : undefined;
}

/* ------------------------------------------------------------------ */
/* Building new nodes                                                   */
/* ------------------------------------------------------------------ */

function buildBlock(schema: Schema, nb: NewBlock): PMNode | null {
  if (nb.type === "paragraph") return buildTextBlock(schema, nb.text, nb.style ?? {});

  if (nb.type === "image") {
    const image = schema.nodes.image;
    const paraType = schema.nodes.paragraph;
    if (!image || !paraType) return null;
    const node = image.create({
      src: nb.src,
      alt: nb.alt,
      title: nb.alt,
      name: nb.alt,
      width: nb.width,
      height: nb.height,
      layout: "inline",
    });
    return paraType.create(null, [node]);
  }

  const table = schema.nodes.table;
  const rowType = schema.nodes.tableRow;
  const cellType = schema.nodes.tableCell;
  const headType = schema.nodes.tableHeader;
  const paraType = schema.nodes.paragraph;
  if (!table || !rowType || !cellType || !paraType) return null;

  const rows = nb.rows.map((cells, r) => {
    const isHead = Boolean(nb.headerRow) && r === 0 && Boolean(headType);
    const cellNodes = cells.map((value) => {
      const type: NodeType = isHead && headType ? headType : cellType;
      const para = value
        ? paraType.create(null, [schema.text(value)])
        : paraType.createAndFill();
      return type.create(null, para ? [para] : []);
    });
    return rowType.create(null, cellNodes);
  });
  return table.create(null, rows);
}

function buildTextBlock(schema: Schema, text: string, style: StylePatch): PMNode | null {
  const paraType = schema.nodes.paragraph;
  if (!paraType) return null;

  const attrs: Record<string, unknown> = {};
  if (style.align) attrs.textAlign = style.align;
  if (style.lineHeight) attrs.lineHeight = style.lineHeight;
  for (const key of ["spacingBefore", "spacingAfter", "indent", "indentFirst"] as const) {
    if (style[key] != null) attrs[key] = style[key];
  }

  const marks: Mark[] = [];
  const textStyleType = schema.marks.textStyle;
  if (
    textStyleType &&
    (style.fontSize ||
      style.fontFamily ||
      style.fontFamilyWest ||
      style.letterSpacing ||
      style.color ||
      style.highlight)
  ) {
    marks.push(
      textStyleType.create({
        fontSize: style.fontSize ?? null,
        fontFamily: style.fontFamily ?? null,
        fontFamilyWest: style.fontFamilyWest ?? null,
        letterSpacing: style.letterSpacing ?? null,
        color: style.color ? (canonColor(style.color) ?? style.color) : null,
        backgroundColor: style.highlight ? (canonColor(style.highlight) ?? style.highlight) : null,
      }),
    );
  }
  for (const key of MARK_STYLE_KEYS) {
    const mark = schema.marks[key];
    if (style[key] && mark) marks.push(mark.create());
  }

  const clean = normalise(text);
  const content = clean ? [schema.text(clean, marks)] : [];
  if (style.headingLevel && schema.nodes.heading) {
    return schema.nodes.heading.create({ level: style.headingLevel, ...attrs }, content);
  }
  return paraType.create(attrs, content);
}

/* ------------------------------------------------------------------ */
/* Executing                                                            */
/* ------------------------------------------------------------------ */

export interface ExecResult {
  ok: boolean;
  reason?: string;
  before: Block[];
  after: Block[];
  beforeJSON: unknown;
  afterJSON: unknown;
}

/**
 * Run ops as a single transaction. Nothing is dispatched unless every op landed
 * and — when `expected` is supplied — the result matches the plan the user was
 * shown. That keeps a half-applied AI edit out of the document and out of history.
 */
export function executeOps(editor: Editor, ops: CoreOp[], expected?: Block[]): ExecResult {
  const state = editor.state;
  const before = blocksOf(state.doc);
  const beforeJSON = state.doc.toJSON();
  let tr = state.tr;

  for (const op of ops) {
    const r = runOp(tr, op);
    if (!r.ok) {
      return { ok: false, reason: r.reason, before, after: before, beforeJSON, afterJSON: beforeJSON };
    }
    tr = r.tr;
  }

  const after = blocksOf(tr.doc);
  if (expected && !blocksAgree(expected, after)) {
    return {
      ok: false,
      reason: "文档实际变化与提案不一致，已放弃写入。",
      before,
      after,
      beforeJSON,
      afterJSON: beforeJSON,
    };
  }

  editor.view.dispatch(tr.setMeta("aiEdit", true));
  return { ok: true, before, after, beforeJSON, afterJSON: tr.doc.toJSON() };
}

interface RunResult {
  ok: boolean;
  tr: Transaction;
  reason?: string;
}

function runOp(tr: Transaction, op: CoreOp): RunResult {
  const schema = tr.doc.type.schema;

  switch (op.kind) {
    case "setText": {
      const target = locateAddress(tr.doc, op.addr);
      if (!target) return { ok: false, tr, reason: "找不到要改写的段落" };
      if (!target.inner || !target.textNode) return { ok: false, tr, reason: "该元素不支持文字改写" };
      const text = normalise(op.text);
      const marks = marksOf(target.textNode);
      const { from, to } = target.inner;
      if (!text) tr.delete(from, to);
      else tr.replaceWith(from, to, schema.text(text, marks));
      return { ok: true, tr };
    }

    case "setStyle": {
      const target = locateAddress(tr.doc, op.addr);
      if (!target) return { ok: false, tr, reason: "找不到要设置样式的段落" };
      return applyStyle(tr, target, op.style);
    }

    case "insertBlock": {
      const node = buildBlock(schema, op.block);
      if (!node) return { ok: false, tr, reason: "当前文档不支持插入该元素" };
      if (tr.doc.childCount === 0) {
        tr.insert(0, node);
        return { ok: true, tr };
      }
      const at = locateAnchor(tr.doc, op.anchor);
      if (!at) return { ok: false, tr, reason: "找不到插入位置" };
      const insertPos = op.anchor.at === "start" ? at.from : at.to;
      tr.insert(insertPos, node);
      return { ok: true, tr };
    }

    case "deleteBlock": {
      const target = locateAddress(tr.doc, op.addr);
      if (!target) return { ok: false, tr, reason: "找不到要删除的段落" };
      // A list item is removed as its whole <li> span.
      tr.delete(target.from, target.to);
      return { ok: true, tr };
    }
  }
}

function applyStyle(tr: Transaction, target: Located, style: StylePatch): RunResult {
  const schema = tr.doc.type.schema;
  const node = target.textNode;
  if (!node || !TEXTUAL.has(node.type.name)) return { ok: false, tr, reason: "该元素不支持样式修改" };

  if (target.mixed.length) {
    const clash = target.mixed.filter((k) => k in style);
    if (clash.length) {
      return {
        ok: false,
        tr,
        reason: `段内${clash.map((k) => CHAR_STYLE_LABELS[k]).join("、")}只有部分文字带格式，整段设置会把它抹平`,
      };
    }
  }

  const attrKeys = ["align", "lineHeight", "spacingBefore", "spacingAfter", "indent", "indentFirst", "headingLevel"] as const;
  const touchesAttrs = attrKeys.some((k) => style[k] !== undefined);

  if (touchesAttrs) {
    const attrs: Record<string, unknown> = { ...node.attrs };
    if (style.align !== undefined) attrs.textAlign = style.align ?? null;
    if (style.lineHeight !== undefined) attrs.lineHeight = style.lineHeight ?? null;
    for (const key of ["spacingBefore", "spacingAfter", "indent", "indentFirst"] as const) {
      if (style[key] !== undefined) attrs[key] = style[key] ?? null;
    }
    let type: NodeType = node.type;
    if (style.headingLevel !== undefined) {
      if (style.headingLevel === null) {
        type = schema.nodes.paragraph ?? type;
        delete attrs.level;
      } else {
        type = schema.nodes.heading ?? type;
        attrs.level = style.headingLevel;
      }
    }
    try {
      // textPos, not `from`: for a list item those differ, and rewriting the
      // <li> as a paragraph is what wrecked whole documents.
      tr.setNodeMarkup(target.textPos, type, attrs);
    } catch (e) {
      return { ok: false, tr, reason: e instanceof Error ? e.message : "样式设置失败" };
    }
  }

  const charKeys = (Object.keys(TEXT_STYLE_ATTRS) as AttrStyleKey[]).filter(
    (k) => style[k] !== undefined,
  );
  const touchedMarkKeys = MARK_STYLE_KEYS.filter((k) => style[k] !== undefined);
  const touchesChars = touchedMarkKeys.length > 0 || charKeys.length > 0;
  if (!touchesChars) return { ok: true, tr };
  if (!target.inner) return { ok: false, tr, reason: "该元素不支持字符样式" };

  const { from, to } = target.inner;

  for (const key of touchedMarkKeys) {
    const mark = schema.marks[key];
    if (!mark) return { ok: false, tr, reason: `文档不支持${CHAR_STYLE_LABELS[key]}` };
    if (style[key]) tr.addMark(from, to, mark.create());
    else tr.removeMark(from, to, mark);
  }

  const textStyleType = schema.marks.textStyle;
  if (textStyleType && charKeys.length) {
    // Patch each run's own mark: a block-wide mark built from the patch alone
    // would drop the attributes the patch never mentioned — colour, highlight.
    const doc = tr.doc;
    doc.nodesBetween(from, to, (child, pos) => {
      if (!child.isText) return true;
      const existing = child.marks.find((m) => m.type === textStyleType);
      const next: Record<string, unknown> = { ...(existing?.attrs as Record<string, unknown> | null ?? {}) };
      const isLink = child.marks.some((m) => m.type.name === "link");
      for (const key of charKeys) {
        // The document owns a link's colour (see charStyleOf); painting the
        // body colour over it would erase what the template is styling.
        if (key === "color" && isLink) continue;
        const attr = TEXT_STYLE_ATTRS[key];
        // Colours are stored in one canonical form, or the same red reads back
        // as `rgb(192, 0, 0)` and every later comparison disagrees with itself.
        next[attr] = COLOR_ATTR_KEYS.has(attr)
          ? (canonColor(style[key]) ?? style[key] ?? null)
          : (style[key] ?? null);
      }
      const start = pos < from ? from : pos;
      const end = pos + child.nodeSize > to ? to : pos + child.nodeSize;
      if (Object.values(next).some((v) => v !== null && v !== undefined && v !== "")) {
        tr.addMark(start, end, textStyleType.create(next));
      } else {
        tr.removeMark(start, end, textStyleType);
      }
      return false;
    });
  }

  return { ok: true, tr };
}

/** Textual agreement used to guard the dispatch above. */
export function blocksAgree(a: readonly Block[], b: readonly Block[]): boolean {
  if (a.length !== b.length) return false;
  return a.every((x, i) => {
    const y = b[i];
    return y ? x.type === y.type && x.text === y.text : false;
  });
}
