import { useEffect, useState } from "react";
import type { Editor } from "@tiptap/react";
import { PanelLeftClose, PanelLeftOpen, ListTree } from "lucide-react";

interface OutlineItem {
  level: number;
  text: string;
  pos: number;
}

export function Outline({
  editor,
  collapsed,
  onToggle,
}: {
  editor: Editor;
  collapsed: boolean;
  onToggle: () => void;
}) {
  const [items, setItems] = useState<OutlineItem[]>([]);

  useEffect(() => {
    const update = () => {
      const result: OutlineItem[] = [];
      editor.state.doc.descendants((node, pos) => {
        if (node.type.name === "heading") {
          result.push({
            level: Number(node.attrs.level ?? 1),
            text: node.textContent || "（空标题）",
            pos,
          });
        }
      });
      setItems(result);
    };
    update();
    editor.on("update", update);
    return () => {
      editor.off("update", update);
    };
  }, [editor]);

  if (collapsed) {
    return (
      <aside className="sidebar sidebar-collapsed" onClick={onToggle} title="展开大纲（⌘B）">
        <div className="sidebar-collapsed-rail">
          <button
            className="icon"
            title="展开"
            onClick={(e) => {
              e.stopPropagation();
              onToggle();
            }}
          >
            <PanelLeftOpen size={16} />
          </button>
          <ListTree size={18} />
          <span className="sidebar-collapsed-label">大纲</span>
        </div>
      </aside>
    );
  }

  return (
    <aside className="sidebar">
      <div className="panel-header">
        <h2>大纲</h2>
        <button
          className="icon"
          title="收起侧栏（⌘B）"
          onClick={onToggle}
        >
          <PanelLeftClose size={14} />
        </button>
      </div>
      {items.length === 0 ? (
        <div className="outline-empty">
          还没有标题。在工具栏把段落改成"标题 1~6"即可自动生成大纲。
        </div>
      ) : (
        items.map((item, i) => (
          <div
            key={i}
            className={`outline-item h${item.level}`}
            onClick={() => {
              // Move the caret first, then ask PageChrome to scroll to the
              // column containing the heading. Direct scrollIntoView in a
              // multi-column layout scrolls vertically and doesn't land on
              // the right page.
              editor.chain().focus().setTextSelection(item.pos).run();
              window.dispatchEvent(
                new CustomEvent("oc:scroll-to-pos", { detail: item.pos }),
              );
            }}
          >
            {item.text}
          </div>
        ))
      )}
    </aside>
  );
}
