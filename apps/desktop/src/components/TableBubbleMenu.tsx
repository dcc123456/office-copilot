import { BubbleMenu, type Editor } from "@tiptap/react";
import {
  Plus,
  Trash2,
  Columns2,
  Rows2,
  Combine,
  Split,
} from "lucide-react";

/**
 * Floating action bar that appears when the caret is inside a table.
 * Uses the commands provided by @tiptap/extension-table.
 */
export function TableBubbleMenu({ editor }: { editor: Editor }) {
  return (
    <BubbleMenu
      editor={editor}
      shouldShow={({ editor: e }) => e.isActive("table")}
      tippyOptions={{
        placement: "top",
        offset: [0, 6],
        zIndex: 40,
      }}
    >
      <div
        style={{
          display: "flex",
          gap: 2,
          background: "#fff",
          border: "1px solid #c8c6c4",
          borderRadius: 6,
          padding: 4,
          boxShadow: "0 4px 12px rgba(0,0,0,0.12)",
        }}
      >
        <TBtn
          title="在上方插入行"
          onClick={() => editor.chain().focus().addRowBefore().run()}
        >
          <Rows2 size={14} /> <Plus size={10} style={{ marginLeft: -6 }} />
        </TBtn>
        <TBtn
          title="在下方插入行"
          onClick={() => editor.chain().focus().addRowAfter().run()}
        >
          <Plus size={10} style={{ marginRight: -6 }} /> <Rows2 size={14} />
        </TBtn>
        <TBtn
          title="在左侧插入列"
          onClick={() => editor.chain().focus().addColumnBefore().run()}
        >
          <Columns2 size={14} /> <Plus size={10} style={{ marginLeft: -6 }} />
        </TBtn>
        <TBtn
          title="在右侧插入列"
          onClick={() => editor.chain().focus().addColumnAfter().run()}
        >
          <Plus size={10} style={{ marginRight: -6 }} /> <Columns2 size={14} />
        </TBtn>
        <span className="divider" />
        <TBtn
          title="删除行"
          onClick={() => editor.chain().focus().deleteRow().run()}
        >
          <Trash2 size={14} />
        </TBtn>
        <TBtn
          title="删除列"
          onClick={() => editor.chain().focus().deleteColumn().run()}
        >
          <Columns2 size={14} />
        </TBtn>
        <TBtn
          title="删除表格"
          onClick={() => editor.chain().focus().deleteTable().run()}
        >
          <Trash2 size={14} color="#a4262c" />
        </TBtn>
        <span className="divider" />
        <TBtn
          title="合并单元格"
          onClick={() => editor.chain().focus().mergeCells().run()}
        >
          <Combine size={14} />
        </TBtn>
        <TBtn
          title="拆分单元格"
          onClick={() => editor.chain().focus().splitCell().run()}
        >
          <Split size={14} />
        </TBtn>
        <TBtn
          title="切换表头行"
          onClick={() => editor.chain().focus().toggleHeaderRow().run()}
        >
          TH
        </TBtn>
      </div>
    </BubbleMenu>
  );
}

function TBtn({
  children,
  onClick,
  title,
}: {
  children: React.ReactNode;
  onClick: () => void;
  title: string;
}) {
  return (
    <button
      type="button"
      title={title}
      onMouseDown={(e) => e.preventDefault()}
      onClick={onClick}
      style={{
        border: "none",
        background: "transparent",
        padding: "4px 6px",
        borderRadius: 3,
        display: "inline-flex",
        alignItems: "center",
        gap: 2,
        cursor: "pointer",
      }}
      onMouseEnter={(e) => (e.currentTarget.style.background = "#f3f2f1")}
      onMouseLeave={(e) => (e.currentTarget.style.background = "transparent")}
    >
      {children}
    </button>
  );
}
