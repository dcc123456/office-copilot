import { BubbleMenu, type Editor } from "@tiptap/react";
import { useState } from "react";
import { Check, X, Sparkles, Loader2 } from "lucide-react";
import { InlineDiff } from "./DiffView.js";
import { commitSelectionRewrite } from "./agent.js";

type SuggestionState =
  | { status: "idle" }
  | { status: "loading"; action: string }
  | {
      status: "ready";
      action: string;
      before: string;
      after: string;
      from: number;
      to: number;
    }
  | { status: "error"; message: string }
  | { status: "saving" };

const ACTIONS: Array<{ key: string; label: string }> = [
  { key: "扩写", label: "扩写" },
  { key: "缩写", label: "缩写" },
  { key: "润色", label: "润色" },
  { key: "换语气:正式商务", label: "更正式" },
  { key: "换语气:口语化", label: "口语化" },
  { key: "翻译:英文", label: "译 EN" },
  { key: "翻译:中文", label: "译 中" },
];

export function AISuggestionBubble({
  editor,
  docKey,
  onApplied,
}: {
  editor: Editor;
  docKey: string;
  onApplied: () => void;
}) {
  const [state, setState] = useState<SuggestionState>({ status: "idle" });

  const run = async (action: string) => {
    const { from, to } = editor.state.selection;
    if (from === to) return;
    const text = editor.state.doc.textBetween(from, to, "\n");
    if (!text.trim()) return;
    setState({ status: "loading", action });
    try {
      const r = await window.officeCopilot.aiTransform({ action, text });
      if (!r.ok) {
        if (r.needsSetup) {
          window.dispatchEvent(new CustomEvent("open-settings"));
        }
        throw new Error(r.message ?? "调用失败");
      }
      setState({ status: "ready", action, before: text, after: r.text!, from, to });
    } catch (e) {
      setState({
        status: "error",
        message: e instanceof Error ? e.message : String(e),
      });
    }
  };

  const accept = async () => {
    if (state.status !== "ready") return;
    setState({ status: "saving" });
    const r = await commitSelectionRewrite(editor, {
      from: state.from,
      to: state.to,
      rewritten: state.after,
      label: "AI 改写",
      docKey,
    });
    if (r.ok) {
      onApplied();
      setState({ status: "idle" });
    } else {
      setState({ status: "error", message: r.message ?? "写入失败。" });
    }
  };

  const reject = () => setState({ status: "idle" });

  return (
    <BubbleMenu
      editor={editor}
      shouldShow={({ editor: e, state: sel }) => {
        const { empty } = sel.selection;
        return !empty && e.isEditable && e.isActive("text") === false ? false : !empty;
      }}
      tippyOptions={{
        placement: "bottom",
        zIndex: 40,
        maxWidth: 420,
        onHide: () => {
          // Avoid leaving a loading/ready state if the menu closes while busy.
          setState((s) => (s.status === "loading" ? { status: "idle" } : s));
        },
      }}
    >
      <div className="ai-bubble">
        {state.status === "idle" && (
          <div className="ai-bubble-row">
            <Sparkles size={14} style={{ color: "var(--accent)" }} />
            {ACTIONS.map((a) => (
              <button key={a.key} type="button" onClick={() => run(a.key)}>
                {a.label}
              </button>
            ))}
          </div>
        )}
        {state.status === "loading" && (
          <div className="ai-bubble-row" style={{ gap: 8 }}>
            <Loader2 size={14} className="spin" />
            AI {state.action}中…
          </div>
        )}
        {state.status === "error" && (
          <div className="ai-bubble-col">
            <div style={{ color: "var(--danger)" }}>{state.message}</div>
            <div className="ai-bubble-row" style={{ justifyContent: "flex-end" }}>
              <button onClick={() => setState({ status: "idle" })}>关闭</button>
            </div>
          </div>
        )}
        {state.status === "ready" && (
          <div className="ai-bubble-col">
            <div style={{ fontSize: 11, color: "var(--text-muted)" }}>
              {state.action} · 选中文字的变化
            </div>
            <div className="ai-diff">
              <InlineDiff before={state.before} after={state.after} />
            </div>
            <div className="ai-bubble-row" style={{ justifyContent: "flex-end" }}>
              <button onClick={reject} title="拒绝（Esc）">
                <X size={12} /> 拒绝
              </button>
              <button className="primary" onClick={() => void accept()} title="接受（Enter）">
                <Check size={12} /> 接受替换
              </button>
            </div>
          </div>
        )}
        {state.status === "saving" && (
          <div className="ai-bubble-row" style={{ gap: 8 }}>
            <Loader2 size={14} className="spin" />
            正在写入文档…
          </div>
        )}
      </div>
    </BubbleMenu>
  );
}
