import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import type { Editor } from "@tiptap/react";

/**
 * In-document find/replace bar (⌘F). Uses ProseMirror positions rather
 * than string-HTML hacking so formatting is preserved: we walk all text
 * nodes, find substring matches, then issue replace/replaceAll via
 * insertTextAt / insertContentAt, which keeps marks intact for the
 * surrounding text.
 *
 * Supports:
 *   - Enter = find next, Shift+Enter = find previous
 *   - ⌘G / ⌘⇧G next/previous match
 *   - Replace one / Replace all
 *   - Case-sensitive toggle
 *   - Match count "X / Y"
 */

interface Match {
  from: number;
  to: number;
  text: string;
}

export function FindBar({
  editor,
  onClose,
}: {
  editor: Editor;
  onClose: () => void;
}) {
  const [query, setQuery] = useState("");
  const [replacement, setReplacement] = useState("");
  const [caseSensitive, setCaseSensitive] = useState(false);
  const [index, setIndex] = useState(0);
  const findRef = useRef<HTMLInputElement>(null);

  const matches = useMemo<Match[]>(() => {
    if (!query) return [];
    const out: Match[] = [];
    const q = caseSensitive ? query : query.toLowerCase();
    editor.state.doc.descendants((node, pos) => {
      if (!node.isText || !node.text) return;
      const hay = caseSensitive ? node.text : node.text.toLowerCase();
      let from = 0;
      let at = hay.indexOf(q, from);
      while (at >= 0) {
        out.push({
          from: pos + at,
          to: pos + at + q.length,
          text: node.text.slice(at, at + q.length),
        });
        from = at + q.length;
        at = hay.indexOf(q, from);
      }
    });
    return out;
  }, [query, caseSensitive, editor.state.doc]);

  useEffect(() => {
    findRef.current?.focus();
    findRef.current?.select();
  }, []);

  useEffect(() => {
    setIndex((i) => Math.min(i, Math.max(0, matches.length - 1)));
  }, [matches.length]);

  // Highlight + scroll to the current match.
  useEffect(() => {
    if (matches.length === 0) return;
    const m = matches[Math.min(index, matches.length - 1)]!;
    editor.commands.setTextSelection({ from: m.from, to: m.to });
    try {
      const dom = editor.view.domAtPos(m.from).node as HTMLElement | undefined;
      (dom ?? editor.view.dom).scrollIntoView({ block: "center" });
    } catch {
      // ignore positions that don't map to DOM
    }
  }, [index, matches, editor]);

  const go = useCallback(
    (dir: 1 | -1) => {
      if (matches.length === 0) return;
      setIndex((i) => (i + dir + matches.length) % matches.length);
    },
    [matches.length],
  );

  const replaceOne = useCallback(() => {
    if (matches.length === 0) return;
    const i = Math.min(index, matches.length - 1);
    const m = matches[i]!;
    editor.chain().focus().insertContentAt(
      { from: m.from, to: m.to },
      replacement,
    ).run();
    // After replacing, matches re-computes via editor.state.doc change.
    setIndex((n) => Math.min(n, Math.max(0, matches.length - 2)));
  }, [editor, index, matches, replacement]);

  const replaceAll = useCallback(() => {
    if (matches.length === 0) return;
    // Replace from the end so positions don't shift for earlier matches.
    const chain = editor.chain().focus();
    for (let i = matches.length - 1; i >= 0; i--) {
      const m = matches[i]!;
      chain.insertContentAt({ from: m.from, to: m.to }, replacement);
    }
    chain.run();
    onClose();
  }, [editor, matches, replacement, onClose]);

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Escape") {
        e.preventDefault();
        onClose();
      } else if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === "g") {
        e.preventDefault();
        go(e.shiftKey ? -1 : 1);
      }
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [go, onClose]);

  return (
    <div className="find-bar">
      <input
        ref={findRef}
        placeholder="查找"
        value={query}
        onChange={(e) => setQuery(e.target.value)}
        onKeyDown={(e) => {
          if (e.key === "Enter") {
            e.preventDefault();
            go(e.shiftKey ? -1 : 1);
          }
        }}
      />
      <input
        placeholder="替换为"
        value={replacement}
        onChange={(e) => setReplacement(e.target.value)}
        onKeyDown={(e) => {
          if (e.key === "Enter") {
            e.preventDefault();
            replaceOne();
          }
        }}
      />
      <label
        className="find-opt"
        title="区分大小写"
      >
        <input
          type="checkbox"
          checked={caseSensitive}
          onChange={(e) => setCaseSensitive(e.target.checked)}
        />
        Aa
      </label>
      <span className="find-count">
        {matches.length
          ? `${Math.min(index + 1, matches.length)} / ${matches.length}`
          : query
            ? "无匹配"
            : ""}
      </span>
      <button onClick={() => go(-1)} disabled={matches.length === 0} title="上一个 (⌘⇧G)">
        ↑
      </button>
      <button onClick={() => go(1)} disabled={matches.length === 0} title="下一个 (⌘G)">
        ↓
      </button>
      <button onClick={replaceOne} disabled={matches.length === 0}>
        替换
      </button>
      <button onClick={replaceAll} disabled={matches.length === 0}>
        全部替换
      </button>
      <button onClick={onClose}>关闭</button>
    </div>
  );
}
