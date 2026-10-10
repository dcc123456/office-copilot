import { useEffect, useMemo, useState } from "react";
import type { Editor } from "@tiptap/react";
import {
  collectRefTargets,
  refInstr,
  type CrossRefOptions,
  type CrossRefReject,
  type RefTargetKind,
  type RefWhat,
} from "../editor/crossrefs.js";

/**
 * 「插入交叉引用」对话框，字段对着 Word 侧栏里那个引用面板：引用哪个书签、引用它的
 * 什么（文字/页码/以上以下）、是否做成超链接。
 *
 * 这里只做装配：挑目标、算缓存结果、造 op 全在 `editor/crossrefs.ts`（AGENTS §3.3 / §5.7），
 * 本组件只把引擎给的指令文本显示出来。
 */

/** 拒绝码 → 界面文案。 */
const REJECT_TEXT: Record<CrossRefReject, string> = {
  "no-targets": "文档里还没有可引用的书签——先插入目录或题注，再来引用它们。",
  "bad-target": "选中的目标已经不在文档里了（被删掉或改过），这次没有写入任何内容。",
  "no-cursor": "把光标放进一段正文里再插入引用。",
  "ops-failed": "插入位置已经失效（那一段被改过或删掉了），这次没有写入任何内容。",
};

/** Word 对话框"引用类型"那一列的中文。 */
const KIND_LABEL: Record<RefTargetKind, string> = {
  caption: "题注",
  toc: "目录条目",
  bookmark: "书签",
};

/** 引用内容三档，标签照 Word 的说法。 */
const WHATS: { value: RefWhat; label: string }[] = [
  { value: "text", label: "书签文字" },
  { value: "page", label: "页码" },
  { value: "position", label: "以上或下方" },
];

export function CrossRefDialog({
  open,
  editor,
  onClose,
  onInsert,
}: {
  open: boolean;
  editor: Editor;
  onClose: () => void;
  /**
   * 落地这次插入。成功返回 null（对话框据此关闭），失败返回引擎的拒绝码。
   */
  onInsert: (opts: CrossRefOptions) => Promise<CrossRefReject | null>;
}) {
  const [target, setTarget] = useState("");
  const [what, setWhat] = useState<RefWhat>("page");
  const [asHyperlink, setAsHyperlink] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);

  useEffect(() => {
    if (!open) return;
    setError(null);
    setBusy(false);
  }, [open]);

  useEffect(() => {
    if (!open) return;
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Escape") onClose();
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [open, onClose]);

  // 每次打开都重新扫一遍书签：对话框开着的时候用户可能刚插了一条题注。
  const targets = useMemo(() => (open ? collectRefTargets(editor.state.doc) : []), [open, editor]);
  const chosen = targets.find((t) => t.name === target) ?? targets[0];
  useEffect(() => {
    if (open && chosen && chosen.name !== target) setTarget(chosen.name);
  }, [open, chosen, target]);

  if (!open) return null;

  const instr = chosen ? refInstr(chosen.name, what, asHyperlink) : "";

  const confirm = async () => {
    if (!chosen) return setError(REJECT_TEXT["no-targets"]);
    setBusy(true);
    const reason = await onInsert({ target: chosen.name, what, asHyperlink });
    setBusy(false);
    if (reason) return setError(REJECT_TEXT[reason]);
    onClose();
  };

  return (
    <div className="modal-backdrop" onClick={onClose}>
      <div className="modal" role="dialog" aria-label="插入交叉引用" onClick={(e) => e.stopPropagation()}>
        <div className="modal-header">
          <strong>插入交叉引用</strong>
        </div>
        <div className="modal-body">
          <div className="ps-row">
            <label>引用目标</label>
            <select value={chosen?.name ?? ""} onChange={(e) => setTarget(e.target.value)}>
              {targets.map((t) => (
                <option key={`${t.name}:${t.blockPos}`} value={t.name}>
                  {`${KIND_LABEL[t.kind]}：${t.blockText || t.name}`}
                </option>
              ))}
            </select>
          </div>
          <div className="ps-row">
            <label>引用内容</label>
            <div className="ps-seg">
              {WHATS.map((w) => (
                <button
                  key={w.value}
                  type="button"
                  className={what === w.value ? "active" : ""}
                  onClick={() => setWhat(w.value)}
                >
                  {w.label}
                </button>
              ))}
            </div>
          </div>
          <div className="ps-row">
            <label>超链接</label>
            <div className="ps-seg">
              <button
                type="button"
                className={asHyperlink ? "active" : ""}
                onClick={() => setAsHyperlink(true)}
              >
                插入为超链接
              </button>
              <button
                type="button"
                className={asHyperlink ? "" : "active"}
                onClick={() => setAsHyperlink(false)}
              >
                纯文字
              </button>
            </div>
          </div>
          <div className="ps-hint">
            {targets.length === 0
              ? REJECT_TEXT["no-targets"]
              : `域指令：{ ${instr} }${
                  chosen && what === "text" && chosen.covered
                    ? `，当前内容为「${chosen.covered}」`
                    : ""
                }`}
          </div>
          {error && <div className="ps-hint">{error}</div>}
        </div>
        <div className="modal-footer">
          <button onClick={onClose}>取消</button>
          <button
            className="primary"
            disabled={busy || !targets.length}
            onClick={() => void confirm()}
          >
            {busy ? "插入中…" : "确定"}
          </button>
        </div>
      </div>
    </div>
  );
}
