import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { Button, Textarea, Spinner, Badge } from "@fluentui/react-components";
import type { ParagraphInfo, Proposal, Snapshot } from "@office-copilot/shared";
import { newId } from "@office-copilot/shared";
import { wordClient } from "@office-copilot/word-client";
import { runCommand } from "./api.js";
import { ProposalCard } from "./components/ProposalCard.js";

type ChatMessage =
  | { id: string; role: "user"; text: string }
  | { id: string; role: "assistant"; text: string; proposal?: Proposal }
  | { id: string; role: "system"; text: string; isError?: boolean };

export function App() {
  const [messages, setMessages] = useState<ChatMessage[]>([
    {
      id: newId("m"),
      role: "system",
      text:
        "你好，我是 Office 智能助手。可以试试：\n" +
        "• 把文中所有“甲方”改成“客户”\n" +
        "• 在文档末尾插入一个 3 行 2 列的表格\n" +
        "• 在开头加一段正式的前言",
    },
  ]);
  const [input, setInput] = useState("");
  const [busy, setBusy] = useState(false);
  const [hostStatus, setHostStatus] = useState<"loading" | "ok" | "err">("loading");
  const [hostLabel, setHostLabel] = useState("正在连接 Word…");
  const [outline, setOutline] = useState<ParagraphInfo[]>([]);
  const [selectionText, setSelectionText] = useState("");
  const [pendingProposal, setPendingProposal] = useState<Proposal | null>(null);
  const [undoStack, setUndoStack] = useState<Snapshot[]>([]);
  const sessionId = useMemo(() => newId("sess"), []);
  const abortRef = useRef<AbortController | null>(null);
  const fileRef = useRef<HTMLInputElement | null>(null);

  const refreshOutline = useCallback(async () => {
    try {
      const paras = await wordClient.listParagraphs();
      setOutline(paras);
    } catch (e) {
      // Not fatal; just leave outline empty.
      console.warn("listParagraphs failed", e);
    }
  }, []);

  const refreshSelection = useCallback(async () => {
    try {
      const { text } = await wordClient.getSelection();
      setSelectionText(text);
    } catch {
      setSelectionText("");
    }
  }, []);

  useEffect(() => {
    Office.onReady(async (info) => {
      if (info.host === Office.HostType.Word) {
        setHostStatus("ok");
        setHostLabel("已连接 Word");
        await refreshOutline();
        await refreshSelection();
      } else {
        setHostStatus("err");
        setHostLabel("未检测到 Word 宿主，UI 仅作演示");
      }
    });
    // Re-read selection whenever the task pane regains focus.
    const onFocus = () => {
      void refreshSelection();
    };
    window.addEventListener("focus", onFocus);
    return () => window.removeEventListener("focus", onFocus);
  }, [refreshOutline, refreshSelection]);

  const appendMessage = useCallback((m: ChatMessage) => {
    setMessages((prev) => [...prev, m]);
  }, []);

  const sendInstruction = useCallback(
    async (text: string) => {
      if (!text.trim() || busy) return;
      setBusy(true);
      appendMessage({ id: newId("m"), role: "user", text });
      setInput("");

      const controller = new AbortController();
      abortRef.current = controller;
      try {
        await refreshSelection();
        await refreshOutline();
        const result = await runCommand(
          {
            sessionId,
            instruction: text,
            selection: selectionText || undefined,
            outline: outline.slice(0, 200),
          },
          {
            onProposal: (p) => {
              setPendingProposal(p);
              appendMessage({
                id: newId("m"),
                role: "assistant",
                text: p.summary,
                proposal: p,
              });
            },
            onError: (msg) => {
              appendMessage({ id: newId("m"), role: "system", text: msg, isError: true });
            },
          },
          controller.signal,
        );
        if (!result.proposal && !result.error) {
          appendMessage({
            id: newId("m"),
            role: "assistant",
            text: "我没有生成可执行的修改，可以换一种说法再试一次。",
          });
        }
      } catch (e) {
        appendMessage({
          id: newId("m"),
          role: "system",
          text: e instanceof Error ? e.message : String(e),
          isError: true,
        });
      } finally {
        setBusy(false);
        abortRef.current = null;
      }
    },
    [appendMessage, busy, outline, refreshOutline, refreshSelection, selectionText, sessionId],
  );

  const acceptProposal = useCallback(
    async (proposal: Proposal) => {
      setBusy(true);
      try {
        const snapshot = await wordClient.applyEdits(proposal.edits);
        setUndoStack((s) => [...s, snapshot]);
        setPendingProposal(null);
        appendMessage({
          id: newId("m"),
          role: "system",
          text: `已应用 ${proposal.edits.length} 项修改，可点“撤销”回退。`,
        });
        await refreshOutline();
      } catch (e) {
        appendMessage({
          id: newId("m"),
          role: "system",
          text: e instanceof Error ? e.message : String(e),
          isError: true,
        });
      } finally {
        setBusy(false);
      }
    },
    [appendMessage, refreshOutline],
  );

  const rejectProposal = useCallback(() => {
    setPendingProposal(null);
    appendMessage({ id: newId("m"), role: "system", text: "已拒绝该提案。" });
  }, [appendMessage]);

  const undo = useCallback(async () => {
    const snap = undoStack.at(-1);
    if (!snap) return;
    setBusy(true);
    try {
      await wordClient.restoreSnapshot(snap);
      setUndoStack((s) => s.slice(0, -1));
      appendMessage({ id: newId("m"), role: "system", text: "已撤销上一次修改。" });
      await refreshOutline();
    } catch (e) {
      appendMessage({
        id: newId("m"),
        role: "system",
        text: e instanceof Error ? e.message : String(e),
        isError: true,
      });
    } finally {
      setBusy(false);
    }
  }, [appendMessage, refreshOutline, undoStack]);

  const insertImage = useCallback(
    async (file: File) => {
      setBusy(true);
      try {
        const dataUrl = await new Promise<string>((resolve, reject) => {
          const r = new FileReader();
          r.onload = () => resolve(String(r.result));
          r.onerror = () => reject(r.error);
          r.readAsDataURL(file);
        });
        const b64 = dataUrl.slice(dataUrl.indexOf(",") + 1);
        await wordClient.insertImage({ base64: b64, width: 200 });
        appendMessage({
          id: newId("m"),
          role: "system",
          text: `已插入图片 ${file.name}`,
        });
      } catch (e) {
        appendMessage({
          id: newId("m"),
          role: "system",
          text: e instanceof Error ? e.message : String(e),
          isError: true,
        });
      } finally {
        setBusy(false);
      }
    },
    [appendMessage],
  );

  const quickInsertParagraph = useCallback(async () => {
    const text = prompt("要插入的段落内容？");
    if (!text) return;
    setBusy(true);
    try {
      await wordClient.insertParagraph({ text, position: "end" });
      await refreshOutline();
    } catch (e) {
      appendMessage({
        id: newId("m"),
        role: "system",
        text: e instanceof Error ? e.message : String(e),
        isError: true,
      });
    } finally {
      setBusy(false);
    }
  }, [appendMessage, refreshOutline]);

  return (
    <div className="app">
      <header className="header">
        <h1>Office 智能助手 · Word</h1>
        <span className={`status ${hostStatus === "ok" ? "ok" : hostStatus === "err" ? "err" : ""}`}>
          {hostStatus === "loading" && <Spinner size="extra-tiny" />} {hostLabel}
        </span>
      </header>

      <div className="messages">
        {messages.map((m) => {
          const isError = m.role === "system" && m.isError;
          const proposal = m.role === "assistant" ? m.proposal : undefined;
          const bubbleClass =
            m.role === "user" ? "user" : isError ? "error" : "assistant";
          return (
            <div key={m.id} className={`bubble ${bubbleClass}`}>
              {m.text}
              {proposal && (
                <ProposalCard
                  proposal={proposal}
                  disabled={busy}
                  onAccept={() => acceptProposal(proposal)}
                  onReject={rejectProposal}
                />
              )}
            </div>
          );
        })}
        {busy && !pendingProposal && (
          <div className="bubble assistant">
            <Spinner size="extra-tiny" /> 正在思考…
          </div>
        )}
      </div>

      <div className="composer">
        <div className="quick">
          <Button size="small" onClick={() => void refreshOutline()}>
            刷新大纲
          </Button>
          <Button size="small" onClick={quickInsertParagraph} disabled={busy}>
            末尾插段落
          </Button>
          <Button size="small" onClick={() => fileRef.current?.click()} disabled={busy}>
            插入图片
          </Button>
          <input
            ref={fileRef}
            type="file"
            accept="image/png,image/jpeg"
            style={{ display: "none" }}
            onChange={(e) => {
              const f = e.target.files?.[0];
              if (f) void insertImage(f);
              e.target.value = "";
            }}
          />
          <Button size="small" onClick={() => void undo()} disabled={busy || undoStack.length === 0}>
            撤销 ({undoStack.length})
          </Button>
          {selectionText && (
            <Badge appearance="tint" color="informative">
              已选 {selectionText.length} 字
            </Badge>
          )}
        </div>

        <Textarea
          value={input}
          onChange={(e, d) => setInput(d.value)}
          onKeyDown={(e) => {
            if (e.key === "Enter" && (e.metaKey || e.ctrlKey)) {
              e.preventDefault();
              void sendInstruction(input);
            }
          }}
          placeholder="用自然语言告诉助手怎么改文档，⌘/Ctrl+Enter 发送"
          rows={3}
        />
        <div className="row" style={{ justifyContent: "flex-end" }}>
          <Button
            appearance="primary"
            onClick={() => void sendInstruction(input)}
            disabled={busy || !input.trim()}
          >
            发送
          </Button>
        </div>
        {outline.length > 0 && (
          <details>
            <summary>文档大纲（{outline.length} 段）</summary>
            <ol style={{ margin: "6px 0 0 18px", padding: 0, fontSize: 12, maxHeight: 120, overflow: "auto" }}>
              {outline.map((p) => (
                <li key={p.index}>
                  <span style={{ color: "#605e5c" }}>#{p.index}</span>{" "}
                  {p.text.slice(0, 80) || <em style={{ color: "#a19f9d" }}>（空段落）</em>}
                </li>
              ))}
            </ol>
          </details>
        )}
      </div>
    </div>
  );
}
