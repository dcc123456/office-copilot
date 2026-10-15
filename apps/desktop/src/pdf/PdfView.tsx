/**
 * PDF 视图（`docs/office-parity-spec.md` 阶段 7）：按页列出显示文字，改其中若干条，再一批落地。
 *
 * 这里只做装配：哪些 run 能改、草稿怎么攒、错误码怎么回给人听，全在 `pdfRuns.ts`；
 * 改写与撤销都只有 `App.tsx` 那一个落地口，组件不碰内容流原文、也不碰 Node 能力（§3.3 / §5.2）。
 * 字节预算不在这里判——编码规则属于 `pdf-io`，视图再算一遍就是第二套换算（§2.1），
 * 所以改长了交给引擎回 `length-mismatch`，这里只把它翻成中文（§5.7）。
 */
import { useCallback, useMemo, useState } from "react";
import type { PdfReadModel, PdfTextEdit } from "@office-copilot/pdf-io";
import type { PdfEditOutcome } from "../global.js";
import { toast } from "../components/ui/Toast.js";
import {
  byteBudgetNote,
  isRunEditable,
  pdfFailureText,
  pdfTargetKey,
  pendingPdfEdits,
  targetOfRun,
  readOnlyReason,
  runsOfPage,
  type PdfDrafts,
} from "./pdfRuns.js";
import "../styles/pdf.css";

export interface PdfViewProps {
  /** 主进程读出来的 PDF 模型；落地成功后换成主进程重读磁盘那份。 */
  document: PdfReadModel;
  /** 这份文件的历史里还有没撤销的 pdf 批次时，撤销按钮才给点（凭据只对那份文件有效）。 */
  canRevert: boolean;
  /** 落地这一批改动：与以后 AI 提案共用同一个写盘口，回包带刷新后的模型。 */
  onCommit: (edits: PdfTextEdit[], summary: string) => Promise<PdfEditOutcome | null>;
  /** 撤销最近一次落地的那批改动（旧文字换回去），回包同样带刷新后的模型。 */
  onRevert: () => Promise<PdfEditOutcome | null>;
}

/**
 * 页导航旁边的那句状态。
 *
 * @param runCount 当前页的 run 条数。
 * @param editableCount 其中可改的条数（内联流与解不出原文的不算）。
 * @returns 一句中文；这一页一条文字都没有时说实话，不显示"0 段"糊弄用户。
 */
function pageSummary(runCount: number, editableCount: number): string {
  if (runCount === 0) return "这一页没有可定位的显示文字";
  return `显示文字 ${runCount} 段，可改 ${editableCount} 段`;
}

export function PdfView({ document, canRevert, onCommit, onRevert }: PdfViewProps): React.ReactElement {
  const [pageIndex, setPageIndex] = useState(0);
  const [drafts, setDrafts] = useState<PdfDrafts>({});
  const [busy, setBusy] = useState(false);

  const page = Math.min(pageIndex, Math.max(0, document.pageCount - 1));
  const pageRuns = useMemo(() => runsOfPage(document, page), [document, page]);
  const pending = useMemo(() => pendingPdfEdits(document, drafts), [document, drafts]);

  const setDraft = useCallback((key: string, text: string) => {
    setDrafts((prev) => {
      const next = { ...prev };
      if (text === "") delete next[key];
      else next[key] = text;
      return next;
    });
  }, []);

  const handleSave = useCallback(async () => {
    if (!pending.length || busy) return;
    setBusy(true);
    try {
      const outcome = await onCommit(pending, `改 ${pending.length} 处文字`);
      if (outcome?.ok) {
        // 落地只追加了一段 delta，原文件的字节一个没重排——这句话是本切片唯一的"写对了"证据。
        setDrafts({});
        toast(`已写入：追加 ${outcome.deltaBytes ?? 0} 字节增量，原来的字节没重排。`);
        return;
      }
      if (outcome?.failures?.length) toast(`保存已取消（磁盘上的文件一个字都没动）：${pdfFailureText(outcome.failures)}`, "error");
      else if (outcome?.message) toast(outcome.message, "error");
    } finally {
      setBusy(false);
    }
  }, [busy, onCommit, pending]);

  const handleRevert = useCallback(async () => {
    if (busy) return;
    setBusy(true);
    try {
      const outcome = await onRevert();
      if (outcome?.ok) {
        setDrafts({});
        toast(`已撤销：又追加 ${outcome.deltaBytes ?? 0} 字节，把旧文字换回去。`);
        return;
      }
      if (outcome?.failures?.length) toast(`撤销已取消：${pdfFailureText(outcome.failures)}`, "error");
      else if (outcome?.message) toast(outcome.message, "error");
    } finally {
      setBusy(false);
    }
  }, [busy, onRevert]);

  return (
    <div className="pdf-view">
      <div className="pdf-view-bar">
        <button type="button" disabled={page <= 0 || busy} onClick={() => setPageIndex(page - 1)}>
          上一页
        </button>
        <span className="pdf-view-page">
          第 {page + 1} / {document.pageCount} 页
        </span>
        <button type="button" disabled={page >= document.pageCount - 1 || busy} onClick={() => setPageIndex(page + 1)}>
          下一页
        </button>
        <span className="pdf-view-count">{pageSummary(pageRuns.length, pageRuns.filter(isRunEditable).length)}</span>
      </div>

      <ul className="pdf-runs">
        {pageRuns.map((run) => {
          const target = targetOfRun(run);
          const key = target ? pdfTargetKey(target) : "";
          const editable = isRunEditable(run);
          const changed = editable && drafts[key] !== undefined && drafts[key] !== run.text;
          return (
            <li key={`${run.pageIndex}-${run.streamRef}-${run.start}`} className={`pdf-run${changed ? " changed" : ""}`}>
              <div className="pdf-run-meta">
                <span>{run.baseFont ?? "字体认不出"}</span>
                {run.composite ? <span className="pdf-run-tag">CID</span> : null}
                {run.withinForm ? <span className="pdf-run-tag">{`表单 ${run.withinForm}`}</span> : null}
                {editable ? <span className="pdf-run-budget">{byteBudgetNote(run)}</span> : null}
              </div>
              {editable ? (
                <input
                  className="pdf-run-input"
                  type="text"
                  value={drafts[key] ?? run.text ?? ""}
                  disabled={busy}
                  spellCheck={false}
                  onChange={(e) => setDraft(key, e.target.value)}
                />
              ) : (
                <div className="pdf-run-note">{`${run.text ?? "（看不出这段念作什么）"} · ${readOnlyReason(run)}`}</div>
              )}
            </li>
          );
        })}
      </ul>

      <div className="pdf-view-foot">
        <span className="pdf-view-count">{pending.length ? `${pending.length} 处待保存` : "没有待保存的改动"}</span>
        <button type="button" disabled={!Object.keys(drafts).length || busy} onClick={() => setDrafts({})}>
          丢弃草稿
        </button>
        <button type="button" className="pdf-save" disabled={!pending.length || busy} onClick={() => void handleSave()}>
          {busy ? "写入中…" : "保存"}
        </button>
        <button type="button" disabled={!canRevert || busy} onClick={() => void handleRevert()}>
          撤销上一次保存
        </button>
      </div>
    </div>
  );
}
