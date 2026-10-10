import { useEffect, useMemo, useRef, useState } from "react";
import type { Editor } from "@tiptap/react";
import {
  CAPTION_LABELS,
  captionTargetFor,
  type CaptionOptions,
  type CaptionPosition,
  type CaptionReject,
} from "../editor/captions.js";

/**
 * 「插入题注」对话框，字段对着 Word 侧栏里的「题注」：标签、位置、题注文字。
 *
 * 只做装配：算号数、造 op、落文档、写修改历史全在 `editor/captions.ts`，本组件连一份
 * 选项对象都不解释（AGENTS §3.3 / §5.7）。失败原因码由宿主回传，中文文案在这里拼。
 */

/** 预设标签之外的第三项：自己填。 */
const CUSTOM = "其他…";

/** 拒绝码 → 界面文案。引擎只给码，不给话（§5.7）。 */
const REJECT_TEXT: Record<CaptionReject, string> = {
  "bad-label": "标签不能空着：没有序列名的 SEQ 域，Word 也算不出编号。",
  "no-target": "先把光标放到要加题注的那一段（或那张图）里。",
  "ops-failed": "插入位置已经失效（那一段被改过或删掉了），这次没有写入任何内容。",
};

/**
 * @param editor 主编辑器，用来现算"这条该是几号"。
 * @param onInsert 落地这次插入；成功返回 null（对话框随即关闭），失败返回引擎的拒绝码。
 */
export function CaptionDialog({
  open,
  editor,
  onClose,
  onInsert,
}: {
  open: boolean;
  editor: Editor;
  onClose: () => void;
  onInsert: (opts: CaptionOptions) => Promise<CaptionReject | null>;
}) {
  const [preset, setPreset] = useState<string>(CAPTION_LABELS[0]!);
  const [custom, setCustom] = useState("");
  const [captionText, setCaptionText] = useState("");
  const [position, setPosition] = useState<CaptionPosition>("below");
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

  const label = preset === CUSTOM ? custom.trim() : preset;

  // 每次改标签或位置都重数一遍号，让"会插成图 3"在点确定之前就看得见。
  const preview = useMemo(
    () => captionTargetFor(editor.state.doc, editor.state.selection.from, { label, position }),
    [editor, label, position, open],
  );

  if (!open) return null;

  const confirm = async () => {
    setBusy(true);
    const reason = await onInsert({ label, captionText: captionText.trim(), position });
    setBusy(false);
    if (reason) return setError(REJECT_TEXT[reason]);
    onClose();
  };

  return (
    <div className="modal-backdrop" onClick={onClose}>
      <div
        className="modal"
        role="dialog"
        aria-label="插入题注"
        tabIndex={-1}
        ref={rootRef}
        onClick={(e) => e.stopPropagation()}
      >
        <div className="modal-header">
          <strong>插入题注</strong>
        </div>
        <div className="modal-body">
          <div className="ps-row">
            <label>标签</label>
            <select value={preset} onChange={(e) => setPreset(e.target.value)}>
              {[...CAPTION_LABELS, CUSTOM].map((l) => (
                <option key={l} value={l}>
                  {l}
                </option>
              ))}
            </select>
          </div>
          {preset === CUSTOM && (
            <div className="ps-row">
              <label>新标签</label>
              <input
                value={custom}
                placeholder="例如：附则"
                onChange={(e) => setCustom(e.target.value)}
              />
            </div>
          )}
          <div className="ps-row">
            <label>位置</label>
            <div className="ps-seg">
              <button
                type="button"
                className={position === "above" ? "active" : ""}
                onClick={() => setPosition("above")}
              >
                选中项上方
              </button>
              <button
                type="button"
                className={position === "below" ? "active" : ""}
                onClick={() => setPosition("below")}
              >
                选中项下方
              </button>
            </div>
          </div>
          <div className="ps-row">
            <label>说明文字</label>
            <input
              value={captionText}
              placeholder="可留空"
              onChange={(e) => setCaptionText(e.target.value)}
            />
          </div>
          <div className="ps-hint">
            {!label
              ? "填一个标签，正文里的编号按标签各自排"
              : preview
                ? `将插入「${label} ${preview.number}${captionText.trim() ? ` ${captionText.trim()}` : ""}」`
                : "先把光标放到要加题注的那一段里"}
          </div>
          {error && <div className="ps-hint">{error}</div>}
        </div>
        <div className="modal-footer">
          <button onClick={onClose}>取消</button>
          <button
            className="primary"
            disabled={busy || !label || !preview}
            onClick={() => void confirm()}
          >
            {busy ? "插入中…" : "确定"}
          </button>
        </div>
      </div>
    </div>
  );
}
