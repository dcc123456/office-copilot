import { useState } from "react";
import { EditorContent, useEditor, type Editor } from "@tiptap/react";
import {
  nextFootnoteId,
  type FootnoteBodies,
  type FootnoteKind,
  type PMDoc,
} from "@office-copilot/docx-io";
import { X } from "lucide-react";
import { FOOTNOTE_REF_KEY, footnoteRefKey } from "../editor/extensions/FootnoteRef.js";
import { NESTED_EXTENSIONS } from "../editor/extensions/nestedExtensions.js";
import "../styles/footnotePane.css";

/** 两种部件的界面名。 */
const KIND_LABEL: Record<FootnoteKind, string> = { footnote: "脚注", endnote: "尾注" };

/** 新建的脚注还没有正文时空着，与 docx-io 写部件时的空条目一致。 */
const EMPTY_BODY: PMDoc = { type: "doc", content: [{ type: "paragraph" }] };

/** 一条条目在窗格里的位置：部件 + `w:id`。 */
interface SelectedRef {
  kind: FootnoteKind;
  id: string;
}

/**
 * 脚注与尾注的正文窗格。
 *
 * 列表按 Word 的编号排（编号来自正文里引用点的出现顺序，见 `FootnoteRef.ts` 的插件），
 * 选中一条就在下方开一个嵌套编辑器改正文。正文改完只交给上层 state，落盘仍走保存链路
 * 里那个唯一的 `DocxFile`——窗格不写文件，也不自己解析部件。
 *
 * @param editor 主编辑器；插入与跳转引用点都要经它。
 * @param bodies 当前正文的脚注与尾注正文。
 * @param onBodies 正文改动交回上层，保存时随 payload 送进主进程。
 * @param onClose 关闭窗格。
 */
export function FootnotePane({
  editor,
  bodies,
  onBodies,
  onClose,
}: {
  editor: Editor;
  bodies: FootnoteBodies;
  onBodies: (next: FootnoteBodies) => void;
  onClose: () => void;
}) {
  const [selected, setSelected] = useState<SelectedRef | null>(null);
  const index = FOOTNOTE_REF_KEY.getState(editor.state) ?? {};

  /**
   * 一个部件的条目，按 Word 编号排序；正文里已经没有引用的排在最后。
   *
   * @param kind 脚注还是尾注。
   */
  const rowsOf = (kind: FootnoteKind) =>
    [...(bodies[kind] ?? [])].sort((a, b) => {
      const na = index[footnoteRefKey(kind, a.id)]?.number ?? Number.MAX_SAFE_INTEGER;
      const nb = index[footnoteRefKey(kind, b.id)]?.number ?? Number.MAX_SAFE_INTEGER;
      return na - nb;
    });

  /**
   * 在光标处插一个新引用点，并把窗格的编辑焦点切到它。
   *
   * @param kind 脚注还是尾注。
   */
  const insert = (kind: FootnoteKind) => {
    const id = nextFootnoteId(bodies[kind] ?? []);
    editor
      .chain()
      .focus()
      .insertContent({
        type: kind === "footnote" ? "footnoteReference" : "endnoteReference",
        attrs: { id },
      })
      .run();
    onBodies({ ...bodies, [kind]: [...(bodies[kind] ?? []), { id, doc: EMPTY_BODY }] });
    setSelected({ kind, id });
  };

  /**
   * 删掉正文里的引用点。**不删正文**：用户接着 Ctrl+Z 把芯片撤回来时，
   * 刚写的字还在；真正没人引用的条目由保存时的 `syncFootnoteIds` 清掉。
   *
   * @param kind 脚注还是尾注。
   * @param id 条目的 `w:id`。
   */
  const removeRef = (kind: FootnoteKind, id: string) => {
    const pos = index[footnoteRefKey(kind, id)]?.pos;
    if (pos === undefined) return;
    editor.chain().focus().deleteRange({ from: pos, to: pos + 1 }).run();
  };

  /**
   * 把光标放到引用点后面，让正文跟着滚过去。
   *
   * @param kind 脚注还是尾注。
   * @param id 条目的 `w:id`。
   */
  const jumpToRef = (kind: FootnoteKind, id: string) => {
    const pos = index[footnoteRefKey(kind, id)]?.pos;
    if (pos === undefined) return;
    editor.chain().focus().setTextSelection(pos + 1).run();
  };

  const current =
    selected && (bodies[selected.kind] ?? []).some((e) => e.id === selected.id)
      ? selected
      : null;
  const currentDoc = current
    ? (bodies[current.kind] ?? []).find((e) => e.id === current.id)?.doc ?? EMPTY_BODY
    : null;

  return (
    <aside className="sidebar footnote-pane">
      <div className="panel-header">
        <span>脚注与尾注</span>
        <button type="button" className="icon" title="关闭" onClick={onClose}>
          <X size={14} />
        </button>
      </div>
      <div className="fn-actions">
        {(Object.keys(KIND_LABEL) as FootnoteKind[]).map((kind) => (
          <button key={kind} type="button" className="fn-action" onClick={() => insert(kind)}>
            插入{KIND_LABEL[kind]}
          </button>
        ))}
      </div>
      <div className="fn-list">
        {(Object.keys(KIND_LABEL) as FootnoteKind[]).map((kind) => (
          <div key={kind} className="fn-group">
            <div className="fn-group-title">{KIND_LABEL[kind]}</div>
            {rowsOf(kind).length === 0 && <div className="fn-empty">暂无</div>}
            {rowsOf(kind).map((entry) => {
              const info = index[footnoteRefKey(kind, entry.id)];
              const isActive = current?.kind === kind && current.id === entry.id;
              return (
                <button
                  key={entry.id}
                  type="button"
                  className={`fn-item${isActive ? " active" : ""}${info ? "" : " orphan"}`}
                  title={info ? `跳到正文里的引用点（id ${entry.id}）` : "正文里已无引用，保存时删除"}
                  onClick={() => {
                    if (!info) return;
                    setSelected({ kind, id: entry.id });
                    jumpToRef(kind, entry.id);
                  }}
                >
                  <sup className="fn-number">{info?.number ?? "?"}</sup>
                  <span className="fn-preview">{previewOf(entry.doc)}</span>
                </button>
              );
            })}
          </div>
        ))}
      </div>
      {current && currentDoc && (
        <div className="fn-editor">
          <div className="fn-editor-head">
            <span>
              {KIND_LABEL[current.kind]} {index[footnoteRefKey(current.kind, current.id)]?.number ?? "?"}
            </span>
            <button
              type="button"
              className="fn-action"
              onClick={() => removeRef(current.kind, current.id)}
            >
              删除引用
            </button>
          </div>
          <FootnoteBodyEditor
            kind={current.kind}
            id={current.id}
            initialDoc={currentDoc}
            onChange={(doc) => {
              const list = (bodies[current.kind] ?? []).map((entry) =>
                entry.id === current.id ? { ...entry, doc } : entry,
              );
              onBodies({ ...bodies, [current.kind]: list });
            }}
          />
        </div>
      )}
    </aside>
  );
}

/**
 * 一条脚注正文的嵌套编辑器。
 *
 * `key` 由调用方给的 `kind + id` 决定：换条目编辑时必须重建实例，否则上一份正文
 * 会留在状态里，保存时写进另一条脚注。
 *
 * @param kind 脚注还是尾注，只用于样式类名。
 * @param id 条目的 `w:id`，进 deps 让编辑器跟着换。
 * @param initialDoc 这条脚注的当前正文。
 * @param onChange 每次编辑交回整份正文 JSON。
 */
function FootnoteBodyEditor({
  kind,
  id,
  initialDoc,
  onChange,
}: {
  kind: FootnoteKind;
  id: string;
  initialDoc: PMDoc;
  onChange: (doc: PMDoc) => void;
}) {
  const editor = useEditor(
    {
      extensions: NESTED_EXTENSIONS,
      content: initialDoc as never,
      editorProps: { attributes: { class: `hf-editor fn-editor-${kind}` } },
      onUpdate: ({ editor: e }) => onChange(e.getJSON() as PMDoc),
    },
    [kind, id],
  );
  if (!editor) return null;
  return <EditorContent editor={editor} />;
}

/**
 * 列表上一行显示的那句正文。
 *
 * @param doc 这条脚注的正文。
 * @returns 第一个非空块的纯文本；新建的脚注还没有字，就是空串。
 */
function previewOf(doc: PMDoc): string {
  for (const block of (doc.content ?? []) as { content?: { text?: string }[] }[]) {
    const text = (block.content ?? []).map((inline) => inline.text ?? "").join("");
    if (text) return text;
  }
  return "";
}
