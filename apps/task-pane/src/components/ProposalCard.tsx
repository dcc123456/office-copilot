import { Button } from "@fluentui/react-components";
import type { Proposal, ProposedEdit } from "@office-copilot/shared";

export interface ProposalCardProps {
  proposal: Proposal;
  disabled?: boolean;
  onAccept: () => void;
  onReject: () => void;
}

/**
 * Very small line-based diff: shows the literal `find` and `replace` strings
 * side-by-side. This does not run against the live document; it reflects the
 * LLM's declared intent. The real match count is surfaced after apply.
 */
function ReplacePreview({ edit }: { edit: Extract<ProposedEdit, { type: "replace_text" }> }) {
  return (
    <div className="edit">
      <div className="edit-type">替换文字 {edit.scope === "selection" ? "（选区）" : "（全文）"}</div>
      <div className="diff-line before">- {edit.find}</div>
      <div className="diff-line after">+ {edit.replace || "（删除）"}</div>
    </div>
  );
}

function ParagraphPreview({
  edit,
}: {
  edit: Extract<ProposedEdit, { type: "insert_paragraph" }>;
}) {
  return (
    <div className="edit">
      <div className="edit-type">
        插入段落（{edit.position}
        {edit.style ? `，样式：${edit.style}` : ""}）
      </div>
      <div className="diff-line after">+ {edit.text}</div>
    </div>
  );
}

function TablePreview({ edit }: { edit: Extract<ProposedEdit, { type: "insert_table" }> }) {
  return (
    <div className="edit">
      <div className="edit-type">
        插入表格 {edit.headerRow ? "（首行加粗）" : ""} · {edit.rows.length} 行 ×{" "}
        {edit.rows[0]?.length ?? 0} 列
      </div>
      <pre className="table-preview">
        {edit.rows.map((r, i) => r.join(" | ") + (i < edit.rows.length - 1 ? "\n" : "")).join("")}
      </pre>
    </div>
  );
}

function ImagePreview({ edit }: { edit: Extract<ProposedEdit, { type: "insert_image" }> }) {
  return (
    <div className="edit">
      <div className="edit-type">插入图片</div>
      <img
        src={edit.dataUrl}
        alt="preview"
        style={{ maxWidth: "100%", borderRadius: 4, border: "1px solid #edebe9" }}
      />
    </div>
  );
}

export function ProposalCard({ proposal, disabled, onAccept, onReject }: ProposalCardProps) {
  return (
    <div className="proposal" style={{ marginTop: 8 }}>
      <h3>修改提案</h3>
      <div style={{ color: "#605e5c" }}>{proposal.summary}</div>
      {proposal.edits.map((edit, i) => {
        switch (edit.type) {
          case "replace_text":
            return <ReplacePreview key={i} edit={edit} />;
          case "insert_paragraph":
            return <ParagraphPreview key={i} edit={edit} />;
          case "insert_table":
            return <TablePreview key={i} edit={edit} />;
          case "insert_image":
            return <ImagePreview key={i} edit={edit} />;
        }
      })}
      <div className="actions">
        <Button size="small" onClick={onReject} disabled={disabled}>
          拒绝
        </Button>
        <Button size="small" appearance="primary" onClick={onAccept} disabled={disabled}>
          接受并写入
        </Button>
      </div>
    </div>
  );
}
