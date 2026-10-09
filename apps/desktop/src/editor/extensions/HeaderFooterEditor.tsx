import { useEffect, type ReactNode } from "react";
import { EditorContent, useEditor, type Editor } from "@tiptap/react";
import StarterKit from "@tiptap/starter-kit";
import Underline from "@tiptap/extension-underline";
import TextAlign from "@tiptap/extension-text-align";
import Link from "@tiptap/extension-link";
import { TextStyle, TextStyleCommands } from "./TextStyle.js";
import { Field } from "./Field.js";
import { Hash, Calendar, AlignLeft, AlignCenter, AlignRight } from "lucide-react";

const EXTENSIONS = [
  StarterKit.configure({ heading: { levels: [1, 2, 3, 4, 5, 6] } }),
  Underline,
  TextAlign.configure({ types: ["heading", "paragraph"] }),
  TextStyle,
  TextStyleCommands,
  Field,
  Link.configure({ openOnClick: false }),
];

/**
 * Compact TipTap editor for header/footer content. Same schema as the
 * main editor but without tables, images, pagination, or drag, so it
 * renders as a single line/band at the top or bottom of every page.
 */
export function HeaderFooterEditor({
  kind,
  initialDoc,
  onUpdate,
  currentPage,
  pageCount,
  toolbarExtras,
}: {
  kind: "header" | "footer";
  initialDoc: Record<string, unknown> | null;
  onUpdate: (doc: Record<string, unknown>) => void;
  currentPage?: number;
  pageCount?: number;
  toolbarExtras?: ReactNode;
}) {
  const editor = useEditor({
    extensions: EXTENSIONS,
    content: (initialDoc as never) ?? {
      type: "doc",
      content: [{ type: "paragraph" }],
    },
    editorProps: { attributes: { class: `hf-editor hf-${kind}` } },
    onUpdate: ({ editor: e }) => {
      onUpdate(e.getJSON() as Record<string, unknown>);
    },
  });

  // Live-update PAGE / NUMPAGES / DATE field values as pagination changes.
  useEffect(() => {
    if (!editor || currentPage === undefined || pageCount === undefined) return;
    const { state, view } = editor;
    const tr = state.tr;
    let touched = false;
    state.doc.descendants((node, pos) => {
      if (node.type.name !== "field") return;
      const k = String(node.attrs.kind).toUpperCase();
      let display = node.attrs.display as string;
      if (k === "PAGE") display = String(currentPage);
      else if (k === "NUMPAGES") display = String(pageCount);
      else if (k === "DATE") display = new Date().toLocaleDateString();
      else return;
      if (display !== node.attrs.display) {
        tr.setNodeMarkup(pos, undefined, { ...node.attrs, display });
        touched = true;
      }
    });
    if (touched) {
      tr.setMeta("addToHistory", false);
      view.dispatch(tr);
    }
  }, [editor, currentPage, pageCount]);

  if (!editor) return null;

  return (
    <div className="hf-root" onMouseDown={(e) => e.stopPropagation()}>
      <div className="hf-toolbar">
        <strong style={{ fontSize: 11, marginRight: 8 }}>
          {kind === "header" ? "页眉" : "页脚"}
        </strong>
        <button
          type="button"
          title="左对齐"
          onClick={() => editor.chain().focus().setTextAlign("left").run()}
        >
          <AlignLeft size={12} />
        </button>
        <button
          type="button"
          title="居中"
          onClick={() => editor.chain().focus().setTextAlign("center").run()}
        >
          <AlignCenter size={12} />
        </button>
        <button
          type="button"
          title="右对齐"
          onClick={() => editor.chain().focus().setTextAlign("right").run()}
        >
          <AlignRight size={12} />
        </button>
        <span className="divider" />
        <button
          type="button"
          title="插入页码字段"
          onClick={() =>
            (
              editor.chain().focus() as unknown as {
                insertField: (k: string, i?: string) => { run: () => void };
              }
            )
              .insertField("page", "PAGE")
              .run()
          }
        >
          <Hash size={12} /> 页码
        </button>
        <button
          type="button"
          title="插入总页数字段"
          onClick={() =>
            (
              editor.chain().focus() as unknown as {
                insertField: (k: string, i?: string) => { run: () => void };
              }
            )
              .insertField("numPages", "NUMPAGES")
              .run()
          }
        >
          共 X 页
        </button>
        <button
          type="button"
          title="插入日期字段"
          onClick={() =>
            (
              editor.chain().focus() as unknown as {
                insertField: (k: string, i?: string) => { run: () => void };
              }
            )
              .insertField("date", "DATE")
              .run()
          }
        >
          <Calendar size={12} /> 日期
        </button>
        {toolbarExtras}
      </div>
      <EditorContent editor={editor} />
    </div>
  );
}
