import { useEffect, useState } from "react";
import { ZoomIn, ZoomOut } from "lucide-react";
import { useUiStore } from "../uiStore.js";

export function StatusBar({
  wordCount,
  charCount,
  editorEl,
}: {
  wordCount: number;
  charCount: number;
  editorEl: HTMLElement | null;
}) {
  const [page, setPage] = useState(1);
  const [total, setTotal] = useState(1);
  const zoom = useUiStore((s) => s.zoom);
  const setZoom = useUiStore((s) => s.setZoom);

  useEffect(() => {
    if (!editorEl) return;
    const onPage = (e: Event) => {
      const d = (e as CustomEvent<{ page: number; total: number }>).detail;
      setPage(d.page);
      setTotal(d.total);
    };
    editorEl.addEventListener("oc:pagechange", onPage);
    return () => editorEl.removeEventListener("oc:pagechange", onPage);
  }, [editorEl]);

  const step = (d: number) =>
    setZoom(Math.round((zoom + d) * 10) / 10);

  return (
    <div className="statusbar">
      <span>第 {page} / {total} 页</span>
      <span>字数：{wordCount}</span>
      <span>字符数：{charCount}</span>
      {import.meta.env.DEV && (
        <span
          className="dev-stamp"
          title="开发构建标记：窗口里看不到 UI·r3 说明运行的不是当前代码（重启 pnpm dev + ⌘R）"
        >
          UI·r3
        </span>
      )}
      <span className="right zoom-ctl">
        <button
          type="button"
          className="zoom-btn"
          title="缩小"
          onClick={() => step(-0.1)}
        >
          <ZoomOut size={13} />
        </button>
        <button
          type="button"
          className="zoom-label"
          title="重置 100%"
          onClick={() => setZoom(1)}
        >
          {Math.round(zoom * 100)}%
        </button>
        <button
          type="button"
          className="zoom-btn"
          title="放大"
          onClick={() => step(0.1)}
        >
          <ZoomIn size={13} />
        </button>
      </span>
    </div>
  );
}
