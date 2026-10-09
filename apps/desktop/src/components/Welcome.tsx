import { useEffect, useState } from "react";
import { FileText, FolderOpen, FilePlus2, Download } from "lucide-react";

/**
 * Welcome screen shown when no document is open. Lists recent files
 * (pulled from the main process) and offers the basic file actions.
 */
export function Welcome({
  onNew,
  onOpen,
  onOpenRecent,
}: {
  onNew: () => void;
  onOpen: () => void;
  onOpenRecent: (path: string) => void;
}) {
  const [recent, setRecent] = useState<string[]>([]);

  useEffect(() => {
    window.officeCopilot?.listRecent().then((paths) => setRecent(paths));
    const off = window.officeCopilot?.onRecentChanged((paths) => setRecent(paths));
    return () => off?.();
  }, []);

  return (
    <div className="welcome">
      <div className="welcome-card">
        <h1>Office Copilot</h1>
        <p className="welcome-sub">高保真 .docx 编辑器 + AI 写作助手</p>
        <div className="welcome-actions">
          <button className="primary" onClick={onNew}>
            <FilePlus2 size={16} /> 新建空白文档
          </button>
          <button onClick={onOpen}>
            <FolderOpen size={16} /> 打开…
          </button>
        </div>
        {recent.length > 0 && (
          <div className="welcome-recent">
            <h2>最近打开</h2>
            <ul>
              {recent.map((p) => (
                <li key={p} onClick={() => onOpenRecent(p)} title={p}>
                  <FileText size={14} />
                  <span className="welcome-recent-name">
                    {p.split(/[\\/]/).pop()}
                  </span>
                  <span className="welcome-recent-path">{p}</span>
                </li>
              ))}
            </ul>
          </div>
        )}
        <div className="welcome-tip">
          <Download size={12} />
          提示：⌘N 新建 · ⌘O 打开 · ⌘S 保存 · ⌘P 导出 PDF · 也可直接把 .docx 拖进窗口
        </div>
      </div>
    </div>
  );
}
