import { EditorContent, useEditor, type Editor } from "@tiptap/react";
import type { CommentBody, PMDoc } from "@office-copilot/docx-io";
import { X } from "lucide-react";
import { commentAnchorRanges } from "../editor/extensions/Comment.js";
import { NESTED_EXTENSIONS } from "../editor/extensions/nestedExtensions.js";
import { deleteCommentById } from "../editor/comments.js";
import { bodyPreview } from "../editor/docBlocks.js";
import "../styles/commentPane.css";

/**
 * 批注侧栏（排名 4 ⑤）。
 *
 * 与脚注窗格同一个分工：正文只装锚点，批注文字是 `word/comments.xml` 的条目，窗格改完
 * 只把条目表交回上层 state，落盘仍走保存链路里那个唯一的 `DocxFile`。
 * 序号按正文里区间起点的先后排，也就是 Word 侧栏的顺序；正文里已经没有锚点的条目排在最后，
 * 保存时由 `syncCommentIds` 清掉，所以提前标出来。
 *
 * @param editor 主编辑器；跳转与删除都要经它。
 * @param bodies 当前正文的批注条目。
 * @param onBodies 条目改动交回上层，保存时随 payload 送进主进程。
 * @param selectedId 下方编辑器里的那条；由上层拿着，新建批注时才会直接聚焦到新条目。
 * @param onSelect 点某一行时改聚焦。
 * @param onClose 关闭窗格。
 */
export function CommentPane({
  editor,
  bodies,
  onBodies,
  selectedId,
  onSelect,
  onClose,
}: {
  editor: Editor;
  bodies: CommentBody[];
  onBodies: (next: CommentBody[]) => void;
  selectedId: string | null;
  onSelect: (id: string | null) => void;
  onClose: () => void;
}) {
  // 位置索引每次渲染现算：正文改过（打字、撤销）以后锚点就挪了，缓存会跳到别处。
  const anchors = commentAnchorRanges(editor.state.doc);
  const orderOf = new Map(anchors.map((range, i) => [range.id, i]));
  const rows = [...bodies].sort((a, b) => {
    const ia = orderOf.get(a.id) ?? Number.MAX_SAFE_INTEGER;
    const ib = orderOf.get(b.id) ?? Number.MAX_SAFE_INTEGER;
    return ia - ib;
  });

  /**
   * 把正文里那条批注锚点包住的范围选中，并滚过去。
   *
   * @param id 条目的 `w:id`。
   */
  const jumpToAnchor = (id: string) => {
    const range = anchors.find((item) => item.id === id);
    if (!range) return;
    // 锚点本身各占一位，正文从起点后一个字开始、到终点前一个字结束。
    const from = (range.start ?? range.reference ?? 0) + 1;
    const to = Math.max(from, range.end ?? from);
    editor.chain().focus().setTextSelection({ from, to }).scrollIntoView().run();
  };

  /**
   * 删掉一条批注：正文里的三个锚点与部件条目一起走。
   *
   * 删的正是聚焦行时接着聚焦相邻那条——"删除批注"只出现在聚焦行的编辑器下方，
   * 不留下一条就等于把这条命令一起删了；Word 删完也是把焦点落到邻近批注上。
   *
   * @param id 条目的 `w:id`。
   */
  const remove = (id: string) => {
    onBodies(deleteCommentById(editor, bodies, id));
    if (selectedId !== id) return;
    const index = rows.findIndex((entry) => entry.id === id);
    const remaining = rows.filter((entry) => entry.id !== id);
    const next = remaining[index] ?? remaining[remaining.length - 1] ?? null;
    onSelect(next?.id ?? null);
  };

  const current = rows.find((entry) => entry.id === selectedId) ?? null;

  return (
    <aside className="sidebar comment-pane">
      <div className="panel-header">
        <span>批注</span>
        <button type="button" className="icon" title="关闭" onClick={onClose}>
          <X size={14} />
        </button>
      </div>
      <div className="cp-list">
        {rows.length === 0 && <div className="cp-empty">在正文里选中文字，右键「新建批注」。</div>}
        {rows.map((entry) => {
          const order = orderOf.get(entry.id);
          const isActive = current?.id === entry.id;
          return (
            <button
              key={entry.id}
              type="button"
              className={`cp-item${isActive ? " active" : ""}${order === undefined ? " orphan" : ""}`}
              title={
                order === undefined
                  ? "正文里已无锚点，保存时删除"
                  : `跳到正文里的批注区间（id ${entry.id}）`
              }
              onClick={() => {
                if (order === undefined) return;
                onSelect(entry.id);
                jumpToAnchor(entry.id);
              }}
            >
              <sup className="cp-number">{order === undefined ? "?" : order + 1}</sup>
              <span className="cp-body">
                <span className="cp-preview">{bodyPreview(entry.doc)}</span>
                <span className="cp-meta">
                  {entry.author}
                  {entry.date ? ` · ${entry.date.slice(0, 16).replace("T", " ")}` : ""}
                </span>
              </span>
            </button>
          );
        })}
      </div>
      {current && (
        <div className="cp-editor">
          <div className="cp-editor-head">
            <span className="cp-editor-author">{current.author}</span>
            <button type="button" className="cp-action" onClick={() => remove(current.id)}>
              删除批注
            </button>
          </div>
          <CommentBodyEditor
            key={current.id}
            id={current.id}
            initialDoc={current.doc}
            onChange={(doc) =>
              onBodies(
                bodies.map((entry) => (entry.id === current.id ? { ...entry, doc } : entry)),
              )
            }
          />
        </div>
      )}
    </aside>
  );
}

/**
 * 一条批注正文的嵌套编辑器。
 *
 * `key` 由调用方给的 id 决定：换条目编辑时必须重建实例，否则上一份正文会留在状态里，
 * 保存时写进另一条批注。
 *
 * @param id 条目的 `w:id`，进 deps 让编辑器跟着换。
 * @param initialDoc 这条批注的当前正文。
 * @param onChange 每次编辑交回整份正文 JSON。
 */
function CommentBodyEditor({
  id,
  initialDoc,
  onChange,
}: {
  id: string;
  initialDoc: PMDoc;
  onChange: (doc: PMDoc) => void;
}) {
  const editor = useEditor(
    {
      extensions: NESTED_EXTENSIONS,
      content: initialDoc as never,
      editorProps: { attributes: { class: "hf-editor cp-editor" } },
      onUpdate: ({ editor: e }) => onChange(e.getJSON() as PMDoc),
    },
    [id],
  );
  if (!editor) return null;
  return <EditorContent editor={editor} />;
}
