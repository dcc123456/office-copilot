import { useCallback, useEffect, useMemo, useState } from "react";
import type { Editor } from "@tiptap/react";
import { X, History, RotateCcw, Sparkles, ListChecks, FileClock } from "lucide-react";
import { confirmDialog } from "./ui/Dialog.js";
import { toast } from "./ui/Toast.js";
import { describeCoreOp } from "../ai/docEdit.js";
import { DiffView } from "../ai/DiffView.js";
import { undoModification, type ModificationRecord } from "../ai/agent.js";

interface SnapshotEntry {
  path: string;
  mtime: number;
  size: number;
}

function fmt(t: number): string {
  const d = new Date(t);
  const pad = (n: number) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
}

function reasonFromName(p: string): string {
  const m = p.match(/\d{8}-\d{6}-(.+)\.(docx|json)$/);
  return m && m[1] ? m[1].replace(/_/g, " ") : "自动保存";
}

export function HistoryPanel({
  open,
  onClose,
  editor,
  docKey,
  onChanged,
}: {
  open: boolean;
  onClose: () => void;
  editor: Editor | null;
  docKey: string;
  onChanged: () => void;
}) {
  const [tab, setTab] = useState<"ai" | "snapshots">("ai");
  const [records, setRecords] = useState<ModificationRecord[]>([]);
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const [snapshots, setSnapshots] = useState<SnapshotEntry[]>([]);
  const [loading, setLoading] = useState(false);
  const [undoing, setUndoing] = useState(false);

  const reload = useCallback(async () => {
    setLoading(true);
    const [mods, snaps] = await Promise.all([
      window.officeCopilot.listModifications(docKey) as Promise<ModificationRecord[]>,
      window.officeCopilot.listSnapshots({ originalPath: docKey === "未命名文档" ? null : docKey }),
    ]);
    setRecords(Array.isArray(mods) ? mods : []);
    setSnapshots(snaps);
    setSelectedId((cur) => cur ?? mods[0]?.id ?? null);
    setLoading(false);
  }, [docKey]);

  useEffect(() => {
    if (!open) return;
    void reload();
  }, [open, reload]);

  const selected = useMemo(
    () => records.find((r) => r.id === selectedId) ?? records[0] ?? null,
    [records, selectedId],
  );

  if (!open) return null;

  const undo = async () => {
    if (!selected || !editor) return;
    const ok = await confirmDialog({
      title: "撤销这次修改",
      message: `只回滚「${selected.summary}」这一次 AI 改动，之后你自己做的编辑会保留。继续？`,
      okLabel: "撤销",
    });
    if (!ok) return;
    setUndoing(true);
    const r = await undoModification(editor, selected, docKey);
    setUndoing(false);
    if (!r.ok) {
      toast(r.message ?? "撤销失败。", "error");
      return;
    }
    toast(r.message ? `已部分撤销：${r.message}` : "已撤销这次修改。", r.message ? undefined : "success");
    onChanged();
    await reload();
    setSelectedId(null);
  };

  const restoreSnapshot = async (p: string) => {
    const r = await window.officeCopilot.readSnapshot(p);
    if (!r.ok) {
      toast("读取失败：" + r.message, "error");
      return;
    }
    const ok = await confirmDialog({
      title: "恢复整份快照",
      message: "用此快照覆盖当前文档？当前未保存的修改会丢失。",
      okLabel: "恢复",
      danger: true,
    });
    if (!ok) return;
    try {
      const parsed = JSON.parse(new TextDecoder().decode(r.bytes));
      editor?.commands.setContent(parsed as never);
      toast("已从历史快照恢复。", "success");
      onChanged();
    } catch (e) {
      toast("快照解析失败：" + (e instanceof Error ? e.message : String(e)), "error");
    }
  };

  return (
    <div className="modal-backdrop" onClick={onClose}>
      <div
        className="modal history-modal"
        role="dialog"
        aria-modal="true"
        aria-label="修改历史"
        onClick={(e) => e.stopPropagation()}
      >
        <header className="history-header">
          <History size={16} />
          <strong>修改历史</strong>
          <div className="history-tabs">
            <button
              className={"history-tab" + (tab === "ai" ? " active" : "")}
              onClick={() => setTab("ai")}
            >
              <Sparkles size={12} /> AI 修改（{records.length}）
            </button>
            <button
              className={"history-tab" + (tab === "snapshots" ? " active" : "")}
              onClick={() => setTab("snapshots")}
            >
              <FileClock size={12} /> 整份快照（{snapshots.length}）
            </button>
          </div>
          <button className="icon" onClick={onClose} title="关闭">
            <X size={14} />
          </button>
        </header>

        {tab === "ai" ? (
          <div className="history-body">
            {loading && <div className="history-loading">加载中…</div>}
            {!loading && !records.length && (
              <div className="history-empty">
                还没有 AI 修改记录。让 AI 直接改文档（排版、纠错、扩写、加表格、加示意图）之后，这里会列出每一次改动，可看 diff 并单独撤销。
              </div>
            )}
            <div className="history-list">
              {records.map((r) => (
                <div
                  key={r.id}
                  className={"history-row" + (selected?.id === r.id ? " active" : "")}
                  onClick={() => setSelectedId(r.id)}
                >
                  <div className="history-row-info">
                    <div className="history-row-name">
                      {r.undone ? <span className="undone-tag">已撤销</span> : null}
                      {r.label}：{r.summary}
                    </div>
                    <div className="history-row-meta">
                      {fmt(r.ts)} · {r.ops.length} 处改动
                    </div>
                  </div>
                </div>
              ))}
            </div>
            {selected && (
              <div className="history-detail">
                <div className="history-detail-head">
                  <ListChecks size={13} />
                  <span>改动内容</span>
                  <button
                    className="small-btn"
                    onClick={undo}
                    disabled={undoing || selected.undone || !editor}
                    title="只回滚这一次 AI 修改"
                  >
                    <RotateCcw size={12} /> {undoing ? "撤销中…" : "撤销这次修改"}
                  </button>
                  <button className="small-btn" onClick={onClose}>
                    关闭
                  </button>
                </div>
                <ul className="op-list">
                  {selected.ops.map((op, i) => (
                    <li key={i}>{describeCoreOp(op)}</li>
                  ))}
                </ul>
                <DiffView before={selected.before} after={selected.after} />
              </div>
            )}
          </div>
        ) : (
          <div className="history-body">
            <div className="history-list">
              {snapshots.map((it) => (
                <div key={it.path} className="history-row">
                  <div className="history-row-info">
                    <div className="history-row-name">{reasonFromName(it.path)}</div>
                    <div className="history-row-meta" title={it.path}>
                      {fmt(it.mtime)} · {(it.size / 1024).toFixed(1)} KB
                    </div>
                  </div>
                  <button className="small-btn" onClick={() => void restoreSnapshot(it.path)}>
                    <RotateCcw size={12} /> 恢复整份
                  </button>
                </div>
              ))}
              {!snapshots.length && (
                <div className="history-empty">还没有整份快照。</div>
              )}
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
