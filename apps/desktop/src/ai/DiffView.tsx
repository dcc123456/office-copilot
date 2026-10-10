import { useMemo } from "react";
import { inlineDiff, toGitLines, type GitLine, type InlinePart } from "./diff.js";
import {
  PARAGRAPH_NUM_KEYS,
  STYLE_KEYS,
  type Block,
  type StylePatch,
} from "./docEdit.js";

/** Tables and images take part in the diff, but say what they are. */
function lineText(b: Block): string {
  if (b.type === "table") return `[表格] ${(b.rows ?? []).map((r) => r.join(" | ")).join(" / ") || b.text}`;
  if (b.type === "image") return `[图] ${b.text}`;
  return b.text;
}

const STYLE_LABELS: Record<keyof StylePatch, string> = {
  headingLevel: "级别",
  align: "对齐",
  bold: "加粗",
  italic: "斜体",
  underline: "下划线",
  fontSize: "字号",
  fontFamily: "字体",
  fontFamilyWest: "西文字体",
  letterSpacing: "字间距",
  color: "颜色",
  highlight: "底色",
  lineHeight: "行距",
  spacingBefore: "段前",
  spacingAfter: "段后",
  indent: "左缩进",
  indentRight: "右缩进",
  indentFirst: "首行缩进",
};

function formatStyle(style: StylePatch): string {
  const bits: string[] = [];
  for (const key of STYLE_KEYS) {
    const v = style[key];
    if (v === undefined || v === null || v === false) continue;
    if (key === "bold" || key === "italic" || key === "underline") bits.push(STYLE_LABELS[key]);
    else if (key === "headingLevel") bits.push(`级别 ${v}`);
    else if (isParagraphNumKey(key)) bits.push(`${STYLE_LABELS[key]} ${v}pt`);
    else bits.push(`${STYLE_LABELS[key]} ${v}`);
  }
  return bits.join(" · ") || "默认样式";
}

/**
 * 段落级的数值格式（段前后距、缩进）在 diff 里统一带 pt 后缀。
 *
 * @param key STYLE_KEYS 里的一个键。
 * @returns 属于 PARAGRAPH_NUM_KEYS 时 true；其余 false。
 */
function isParagraphNumKey(key: keyof StylePatch): boolean {
  return (PARAGRAPH_NUM_KEYS as readonly (keyof StylePatch)[]).includes(key);
}

/** Style-only edits leave the text untouched, so show them as `-旧样式 +新样式`. */
function styleLines(before: readonly Block[], after: readonly Block[]): GitLine[] {
  const out: GitLine[] = [];
  const n = Math.min(before.length, after.length);
  for (let i = 0; i < n; i++) {
    const a = before[i]!;
    const b = after[i]!;
    if (a.type !== b.type || a.text !== b.text) continue;
    const from = formatStyle(a.style);
    const to = formatStyle(b.style);
    if (from === to) continue;
    out.push({ kind: "del", beforeNo: i + 1, parts: [{ t: "del", s: `样式：${from}` }] });
    out.push({ kind: "add", afterNo: i + 1, parts: [{ t: "ins", s: `样式：${to}` }] });
  }
  return out;
}

function Parts({ parts }: { parts: InlinePart[] }) {
  return (
    <>
      {parts.map((p, i) =>
        p.t === "eq" ? (
          <span key={i}>{p.s}</span>
        ) : p.t === "ins" ? (
          <ins key={i}>{p.s}</ins>
        ) : (
          <del key={i}>{p.s}</del>
        ),
      )}
    </>
  );
}

function GitRows({ lines }: { lines: GitLine[] }) {
  return (
    <>
      {lines.map((l, i) =>
        l.kind === "omit" ? (
          <div key={i} className="diff-omit">
            ⋯ 未改动 {l.omitted} 行 ⋯
          </div>
        ) : (
          <div key={i} className={`diff-line ${l.kind}`}>
            <span className="diff-ln">{l.beforeNo ?? ""}</span>
            <span className="diff-ln">{l.afterNo ?? ""}</span>
            <span className="diff-sign">{l.kind === "add" ? "+" : l.kind === "del" ? "-" : " "}</span>
            <span className="diff-code">
              <Parts parts={l.parts} />
            </span>
          </div>
        ),
      )}
    </>
  );
}

/**
 * Unified, git-style diff over the block text, with the style changes of
 * untouched paragraphs listed on their own — enough to see exactly what moved.
 */
export function DiffView({
  before,
  after,
}: {
  before: readonly Block[];
  after: readonly Block[];
}) {
  const lines = useMemo(
    () => toGitLines(before.map(lineText), after.map(lineText)),
    [before, after],
  );
  const styles = useMemo(() => styleLines(before, after), [before, after]);

  const added = lines.filter((l) => l.kind === "add").length;
  const removed = lines.filter((l) => l.kind === "del").length;

  if (!added && !removed && !styles.length) {
    return <div className="diff-empty">没有可见变化。</div>;
  }

  return (
    <div className="diff-view">
      <div className="diff-summary">
        <span className="diff-title">
          {before.length ? `文档原 ${before.length} 段` : "空文档"} → 修改后 {after.length} 段
        </span>
        <span className="diff-stat">
          <span className="diff-add">+{added}</span>
          <span className="diff-del">-{removed}</span>
          {styles.length ? <span className="diff-mod">样式 ×{styles.length / 2}</span> : null}
        </span>
      </div>
      {styles.length > 0 && (
        <>
          <div className="diff-file">段落格式变化</div>
          <GitRows lines={styles} />
        </>
      )}
      {(added > 0 || removed > 0) && (
        <>
          <div className="diff-file">正文变化</div>
          <GitRows lines={lines} />
        </>
      )}
    </div>
  );
}

/** Word-level diff for a single rewritten passage (selection rewrites). */
export function InlineDiff({ before, after }: { before: string; after: string }) {
  const parts = useMemo(() => inlineDiff(before, after), [before, after]);
  return (
    <div className="diff-view">
      <div className="diff-line mod">
        <span className="diff-sign">~</span>
        <span className="diff-code">
          <Parts parts={parts} />
        </span>
      </div>
    </div>
  );
}
