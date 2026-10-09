import { useEffect, useRef, useState } from "react";
import type { Editor } from "@tiptap/react";

/**
 * Right-click context menu for tables. Offers the same row/column/merge
 * commands as the bubble menu plus cell alignment and cell background
 * color (which TipTap's table extension persists as `background` attr
 * on tableCell/tableHeader).
 *
 * We don't try to replace the browser's native context menu for text;
 * only when the right-click target is inside a <td>/<th>.
 */

interface MenuState {
  x: number;
  y: number;
}

export function TableContextMenu({ editor }: { editor: Editor }) {
  const [pos, setPos] = useState<MenuState | null>(null);
  const ref = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const onContext = (e: MouseEvent) => {
      const target = e.target as HTMLElement;
      const cell = target.closest("td, th");
      if (!cell) return;
      e.preventDefault();
      setPos({ x: e.clientX, y: e.clientY });
    };
    const onClick = () => setPos(null);
    const onKey = (e: KeyboardEvent) => e.key === "Escape" && setPos(null);
    editor.view.dom.addEventListener("contextmenu", onContext);
    window.addEventListener("click", onClick);
    window.addEventListener("keydown", onKey);
    return () => {
      editor.view.dom.removeEventListener("contextmenu", onContext);
      window.removeEventListener("click", onClick);
      window.removeEventListener("keydown", onKey);
    };
  }, [editor]);

  if (!pos) return null;

  const run = (fn: () => void) => () => {
    fn();
    setPos(null);
  };

  const setCellBg = (color: string | null) => () => {
    const chain = editor.chain().focus();
    if (color) {
      // updateAttributes on the active cell sets background; TipTap's
      // tableCell spec accepts a `background` attribute (color hex).
      chain.updateAttributes("tableCell", { background: color });
      chain.updateAttributes("tableHeader", { background: color });
    } else {
      chain.updateAttributes("tableCell", { background: null });
      chain.updateAttributes("tableHeader", { background: null });
    }
    chain.run();
    setPos(null);
  };

  return (
    <div
      ref={ref}
      className="ctx-menu"
      style={{ left: pos.x, top: pos.y }}
      onMouseDown={(e) => e.preventDefault()}
    >
      <CtxItem onClick={run(() => editor.chain().focus().addRowBefore().run())}>
        ↑ 上方插入行
      </CtxItem>
      <CtxItem onClick={run(() => editor.chain().focus().addRowAfter().run())}>
        ↓ 下方插入行
      </CtxItem>
      <CtxItem
        onClick={run(() => editor.chain().focus().addColumnBefore().run())}
      >
        ← 左侧插入列
      </CtxItem>
      <CtxItem
        onClick={run(() => editor.chain().focus().addColumnAfter().run())}
      >
        → 右侧插入列
      </CtxItem>
      <div className="ctx-sep" />
      <CtxItem onClick={run(() => editor.chain().focus().deleteRow().run())}>
        删除行
      </CtxItem>
      <CtxItem
        onClick={run(() => editor.chain().focus().deleteColumn().run())}
      >
        删除列
      </CtxItem>
      <CtxItem
        onClick={run(() => editor.chain().focus().deleteTable().run())}
        danger
      >
        删除表格
      </CtxItem>
      <div className="ctx-sep" />
      <CtxItem
        onClick={run(() => editor.chain().focus().mergeCells().run())}
      >
        合并单元格
      </CtxItem>
      <CtxItem
        onClick={run(() => editor.chain().focus().splitCell().run())}
      >
        拆分单元格
      </CtxItem>
      <CtxItem
        onClick={run(() => editor.chain().focus().toggleHeaderRow().run())}
      >
        切换表头行
      </CtxItem>
      <div className="ctx-sep" />
      <CtxItem
        onClick={run(() =>
          editor.chain().focus().setCellAttribute("verticalAlign", "top").run(),
        )}
      >
        顶部对齐
      </CtxItem>
      <CtxItem
        onClick={run(() =>
          editor
            .chain()
            .focus()
            .setCellAttribute("verticalAlign", "center")
            .run(),
        )}
      >
        垂直居中
      </CtxItem>
      <CtxItem
        onClick={run(() =>
          editor
            .chain()
            .focus()
            .setCellAttribute("verticalAlign", "bottom")
            .run(),
        )}
      >
        底部对齐
      </CtxItem>
      <div className="ctx-sep" />
      <div className="ctx-subhead">单元格底色</div>
      <div className="ctx-swatches">
        {[
          null,
          "#ffffff",
          "#fff2cc",
          "#fce4d6",
          "#e2efda",
          "#deebf7",
          "#d9e1f2",
          "#ead1dc",
          "#d9d9d9",
          "#000000",
        ].map((c) => (
          <button
            key={c ?? "none"}
            className="ctx-swatch"
            style={{
              background: c ?? "transparent",
              border: c ? "1px solid #999" : "1px dashed #999",
            }}
            onClick={setCellBg(c)}
            title={c ? c : "无填充"}
          />
        ))}
      </div>
    </div>
  );
}

function CtxItem({
  children,
  onClick,
  danger,
}: {
  children: React.ReactNode;
  onClick: () => void;
  danger?: boolean;
}) {
  return (
    <button
      type="button"
      className={`ctx-item ${danger ? "danger" : ""}`}
      onClick={onClick}
    >
      {children}
    </button>
  );
}
