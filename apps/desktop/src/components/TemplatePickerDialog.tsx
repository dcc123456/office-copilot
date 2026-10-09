import { useEffect, useRef, useState, type CSSProperties } from "react";
import { X } from "lucide-react";
import type { StylePatch } from "../ai/docEdit.js";
import { TEMPLATES, completeRoles, fontCss, type DocTemplate } from "../editor/templates.js";

/**
 * Template gallery for 美化排版. Own component rather than a `ui/Dialog.js`
 * request: that one is a promise singleton for confirm/prompt and can only
 * resolve a boolean or a string, which cannot carry ten styled cards.
 */

function ptToPx(pt: number): number {
  return Math.round(pt * (96 / 72) * 10) / 10;
}

function previewStyle(patch?: StylePatch): CSSProperties {
  if (!patch) return {};
  const lh = patch.lineHeight;
  const points = lh ? /^(?:pt|atLeast):(\d+(?:\.\d+)?)$/.exec(lh) : null;
  // 西文在前：CSS 字体栈按字符回退，汉字落到中文字体上，正是 Word 里
  // ascii/eastAsia 两槽的分工。
  const east = fontCss(patch.fontFamily ?? undefined);
  const stack = [patch.fontFamilyWest, east].filter(Boolean).join(", ") || undefined;
  return {
    fontFamily: stack,
    fontSize: patch.fontSize ?? undefined,
    letterSpacing: patch.letterSpacing ?? undefined,
    color: patch.color ?? undefined,
    backgroundColor: patch.highlight ?? undefined,
    fontWeight: patch.bold === undefined ? undefined : patch.bold ? 700 : 400,
    fontStyle: patch.italic === undefined ? undefined : patch.italic ? "italic" : "normal",
    textDecoration: patch.underline === undefined ? undefined : patch.underline ? "underline" : "none",
    textAlign: patch.align ?? undefined,
    lineHeight: points ? ptToPx(Number(points[1])) : lh ?? undefined,
    textIndent: patch.indentFirst ? ptToPx(patch.indentFirst) : undefined,
  };
}

export function TemplatePickerDialog({
  open,
  lastTemplateId,
  onClose,
  onPick,
}: {
  open: boolean;
  lastTemplateId?: string;
  onClose: () => void;
  onPick: (t: DocTemplate) => void;
}) {
  const [active, setActive] = useState(0);
  const rootRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (open) setActive(Math.max(0, TEMPLATES.findIndex((t) => t.id === lastTemplateId)));
  }, [open, lastTemplateId]);

  useEffect(() => {
    if (!open) return;
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Escape") {
        e.stopPropagation();
        onClose();
      } else if (e.key === "ArrowDown" || e.key === "ArrowRight") {
        e.preventDefault();
        setActive((i) => Math.min(TEMPLATES.length - 1, i + 1));
      } else if (e.key === "ArrowUp" || e.key === "ArrowLeft") {
        e.preventDefault();
        setActive((i) => Math.max(0, i - 1));
      } else if (e.key === "Enter") {
        e.preventDefault();
        const t = TEMPLATES[active];
        if (t) onPick(t);
      }
    };
    window.addEventListener("keydown", onKey, true);
    return () => window.removeEventListener("keydown", onKey, true);
  }, [open, active, onClose, onPick]);

  useEffect(() => {
    if (!open) return;
    rootRef.current
      ?.querySelector("[data-tp-active]")
      ?.scrollIntoView({ block: "nearest" });
  }, [open, active]);

  if (!open) return null;

  return (
    <div className="modal-backdrop dialog-backdrop" onClick={onClose}>
      <div
        ref={rootRef}
        className="modal template-picker"
        role="dialog"
        aria-modal="true"
        aria-label="美化排版 · 选择模板"
        onClick={(e) => e.stopPropagation()}
      >
        <div className="settings-header">
          <strong>美化排版</strong>
          <span className="tp-keys">↑↓ 选择 · Enter 应用 · Esc 关闭</span>
          <button
            type="button"
            onClick={onClose}
            className="icon"
            aria-label="关闭"
            style={{ marginLeft: "auto" }}
          >
            <X size={16} />
          </button>
        </div>

        <div className="tp-grid">
          {TEMPLATES.map((t, i) => {
            // Preview the role patches as the document will receive them, so
            // 西文字体 and the derived levels show in the card too.
            const roles = completeRoles(t);
            return (
              <button
                key={t.id}
                type="button"
                className={"tp-card" + (i === active ? " active" : "")}
                {...(i === active ? { "data-tp-active": "" } : {})}
                onMouseEnter={() => setActive(i)}
                onClick={() => onPick(t)}
              >
                <div className="tp-card-head">
                  <span className="theme-swatch-dot" style={{ background: t.swatch }} />
                  <span className="tp-label">{t.label}</span>
                  {t.id === lastTemplateId ? (
                    <span className="tp-badge">上次使用</span>
                  ) : null}
                </div>
                <div className="tp-preview">
                  <div className="tp-line" style={previewStyle(roles.title)}>
                    {t.preview.title}
                  </div>
                  <div className="tp-line" style={previewStyle(roles.h1)}>
                    {t.preview.h1}
                  </div>
                  <div className="tp-line" style={previewStyle(roles.body)}>
                    {t.preview.body}
                  </div>
                </div>
                <div className="tp-spec">{t.spec}</div>
                <div className="tp-hint">{t.hint}</div>
              </button>
            );
          })}
        </div>
      </div>
    </div>
  );
}
