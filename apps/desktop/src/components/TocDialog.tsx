import { useEffect, useMemo, useRef, useState } from "react";
import type { Editor } from "@tiptap/react";
import type { TabStopLeader } from "@office-copilot/docx-io";
import {
  collectTocEntries,
  tocAnchorForSelection,
  type TocOptions,
  type TocReject,
} from "../editor/toc.js";

/**
 * 「插入目录」对话框，字段对着 Word 侧栏里那个目录面板：显示级别、是否显示页码、
 * 制表符前导符，再加上"插在哪"。
 *
 * 这里只做装配：算条目、造 op、落文档、写修改历史全在 `editor/toc.ts`，
 * 本组件连一份选项对象都不解释（AGENTS §3.3 / §5.7）。失败原因码由宿主回传，
 * 中文文案在这里拼。
 */

/** Word 前导符下拉框里的七项，`w:leader` 的取值。 */
const LEADERS: { value: TabStopLeader; label: string }[] = [
  { value: "none", label: "（无）" },
  { value: "dot", label: "点线 …" },
  { value: "hyphen", label: "短划线 - - -" },
  { value: "dash", label: "虚线 — — —" },
  { value: "line", label: "实线 ———" },
  { value: "heavy", label: "重实线 ———" },
  { value: "middleDot", label: "中间点 · · ·" },
];

/** 拒绝码 → 界面文案。引擎只给码，不给话（§5.7）。 */
const REJECT_TEXT: Record<TocReject, string> = {
  "no-headings": "文档里没有够格的标题或大纲级别段落，无法生成目录。",
  "bad-levels": "显示级别要在 1 到 9 之间。",
  "ops-failed": "插入位置已经失效（那一段被改过或删掉了），这次没有写入任何内容。",
};

type Where = "docStart" | "atCursor";

export function TocDialog({
  open,
  editor,
  contentWidthPx,
  onClose,
  onInsert,
}: {
  open: boolean;
  editor: Editor;
  /** 正文区宽度，CSS px（`SectionGeometry.contentWidth`），用来定右对齐制表位。 */
  contentWidthPx: number;
  onClose: () => void;
  /**
   * 落地这次插入。成功时返回 null（对话框随即关闭），失败时返回引擎的拒绝码，
   * 由本组件显示对应中文提示。
   */
  onInsert: (opts: TocOptions) => Promise<TocReject | null>;
}) {
  const [levels, setLevels] = useState(3);
  const [showPageNumbers, setShowPageNumbers] = useState(true);
  const [leader, setLeader] = useState<TabStopLeader>("dot");
  const [where, setWhere] = useState<Where>("atCursor");
  const [error, setError] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);
  const rootRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!open) return;
    setError(null);
    setBusy(false);
    rootRef.current?.focus();
  }, [open]);

  useEffect(() => {
    if (!open) return;
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Escape") onClose();
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [open, onClose]);

  // 每次改级别都重新数一遍条目，让"这一档会收进几条"在点确定之前就看得见。
  // 页码这里不参与（只在对话框里数条目），所以 pageOf 一律给 0。
  const entryCount = useMemo(
    () => collectTocEntries(editor.state.doc, levels, () => 0).length,
    [editor, levels, open],
  );

  if (!open) return null;

  const confirm = async () => {
    setBusy(true);
    const reason = await onInsert({
      levels,
      showPageNumbers,
      leader,
      contentWidthPx,
      // 光标位置在点确定的这一刻才取，免得开着对话框又去别处点了光标。
      position: where === "atCursor" ? tocAnchorForSelection(editor) : { at: "start" },
    });
    setBusy(false);
    if (reason) return setError(REJECT_TEXT[reason]);
    onClose();
  };

  return (
    <div className="modal-backdrop" onClick={onClose}>
      <div
        className="modal"
        role="dialog"
        aria-label="插入目录"
        tabIndex={-1}
        ref={rootRef}
        onClick={(e) => e.stopPropagation()}
      >
        <div className="modal-header">
          <strong>插入目录</strong>
        </div>
        <div className="modal-body">
          <div className="ps-row">
            <label>显示级别</label>
            <select value={levels} onChange={(e) => setLevels(Number(e.target.value))}>
              {[1, 2, 3, 4, 5, 6, 7, 8, 9].map((n) => (
                <option key={n} value={n}>
                  {n}
                </option>
              ))}
            </select>
          </div>
          <div className="ps-row">
            <label>制表符前导符</label>
            <select
              value={leader}
              disabled={!showPageNumbers}
              onChange={(e) => setLeader(e.target.value as TabStopLeader)}
            >
              {LEADERS.map((l) => (
                <option key={l.value} value={l.value}>
                  {l.label}
                </option>
              ))}
            </select>
          </div>
          <div className="ps-row">
            <label>页码</label>
            <div className="ps-seg">
              <button
                type="button"
                className={showPageNumbers ? "active" : ""}
                onClick={() => setShowPageNumbers(true)}
              >
                显示
              </button>
              <button
                type="button"
                className={showPageNumbers ? "" : "active"}
                onClick={() => setShowPageNumbers(false)}
              >
                不显示
              </button>
            </div>
          </div>
          <div className="ps-row">
            <label>插入位置</label>
            <div className="ps-seg">
              <button
                type="button"
                className={where === "docStart" ? "active" : ""}
                onClick={() => setWhere("docStart")}
              >
                文档开头
              </button>
              <button
                type="button"
                className={where === "atCursor" ? "active" : ""}
                onClick={() => setWhere("atCursor")}
              >
                光标所在段之后
              </button>
            </div>
          </div>
          <div className="ps-hint">
            {entryCount > 0
              ? `收录 ${entryCount} 个标题${showPageNumbers ? "" : "（不显示页码）"}`
              : "这个级别下没有可收录的标题"}
          </div>
          {error && <div className="ps-hint">{error}</div>}
        </div>
        <div className="modal-footer">
          <button onClick={onClose}>取消</button>
          <button className="primary" disabled={busy || entryCount === 0} onClick={() => void confirm()}>
            {busy ? "插入中…" : "确定"}
          </button>
        </div>
      </div>
    </div>
  );
}
