import { useEffect, useState } from "react";
import type { Editor } from "@tiptap/react";
import { AlignLeft, AlignCenter, AlignRight, AlignJustify, X } from "lucide-react";
import { useAnchoredMenu } from "../hooks/useAnchoredMenu.js";

/**
 * 段落设置面板。工具栏按钮挂的是同一份内容：
 *   - 对齐方式
 *   - 行距 (single / 1.15 / 1.5 / double / fixed)
 *   - 段前 / 段后 (pt)
 *   - 首行缩进 / 左右缩进 (pt)
 *
 * 数值经自定义命令 `setParagraphFormat` 写进当前选中的段落/标题，落在
 * editor/extensions/ParagraphFormatting.ts，再由 docx-io 往返成 OOXML 的
 * <w:spacing> 与 <w:ind>。右键菜单的「段落(P)…」直接复用 ParagraphFields。
 */

interface Attrs {
  textAlign?: string;
  lineHeight?: string;
  spacingBefore?: number;
  spacingAfter?: number;
  indent?: number;
  indentFirst?: number;
}

const LINE_OPTIONS = new Set([
  "",
  "1",
  "1.15",
  "1.5",
  "2",
  "1.35",
  "pt:20",
  "pt:22",
  "pt:29",
  "atLeast:18",
  "atLeast:20",
  "atLeast:22",
]);

/** A 美化排版 template may set a line rule outside the preset list. */
function lineHeightLabel(value: string): string {
  const fixed = /^pt:(\d+(?:\.\d+)?)$/.exec(value);
  if (fixed?.[1]) return `固定 ${fixed[1]} 磅`;
  const min = /^atLeast:(\d+(?:\.\d+)?)$/.exec(value);
  return min?.[1] ? `最小 ${min[1]} 磅` : value;
}

export function ParagraphPopover({ editor }: { editor: Editor }) {
  const { anchorRef, rect, open, toggle, hide } = useAnchoredMenu(280);

  useEffect(() => {
    if (!open) return;
    const onDown = (e: MouseEvent) => {
      if (anchorRef.current && !anchorRef.current.contains(e.target as Node)) hide();
    };
    document.addEventListener("mousedown", onDown);
    return () => document.removeEventListener("mousedown", onDown);
  }, [open, anchorRef, hide]);

  return (
    <div className="paragraph-popover-wrap" ref={anchorRef}>
      <button
        type="button"
        className={`icon ${open ? "active" : ""}`}
        title="段落设置"
        aria-expanded={open}
        onMouseDown={(e) => e.preventDefault()}
        onClick={toggle}
      >
        <AlignLeft size={16} />
        <span className="caret">▾</span>
      </button>
      {open && rect && (
        <div className="paragraph-popover" style={{ left: rect.left, top: rect.top }}>
          <ParagraphFields editor={editor} onClose={hide} />
        </div>
      )}
    </div>
  );
}

/**
 * 段落设置的实际字段，挂载在任意容器里（工具栏浮层或右键菜单的「段落(P)…」面板）。
 *
 * @param editor 目标编辑器；改动只作用于当前选区涉及的块。
 * @param onClose 关闭回调，标题栏的 × 用。
 */
export function ParagraphFields({
  editor,
  onClose,
}: {
  editor: Editor;
  onClose: () => void;
}) {
  const [attrs, setAttrs] = useState<Attrs>({});

  useEffect(() => {
    const read = () => {
      const a: Attrs = {};
      a.textAlign = editor.getAttributes("paragraph").textAlign as
        | string
        | undefined;
      const anyNode =
        editor.getAttributes("paragraph") ||
        editor.getAttributes("heading") ||
        {};
      a.lineHeight = (anyNode.lineHeight as string | undefined) ?? "";
      a.spacingBefore = anyNode.spacingBefore as number | undefined;
      a.spacingAfter = anyNode.spacingAfter as number | undefined;
      a.indent = anyNode.indent as number | undefined;
      a.indentFirst = anyNode.indentFirst as number | undefined;
      setAttrs(a);
    };
    read();
    editor.on("selectionUpdate", read);
    editor.on("update", read);
    return () => {
      editor.off("selectionUpdate", read);
      editor.off("update", read);
    };
  }, [editor]);

  const apply = (patch: Partial<Attrs>) => {
    setAttrs((a) => ({ ...a, ...patch }));
    // 不 chain().focus()：输入框正在被打字，抢回焦点会让第二个字符落进文档。
    // 命令用的还是 state 里的选区，右键菜单/工具栏打开面板时它已经就位。
    const cmd = editor.chain() as unknown as {
      setParagraphFormat: (p: Partial<Attrs>) => { run: () => boolean };
    };
    cmd.setParagraphFormat(patch).run();
  };

  const align = (a: string) => editor.chain().focus().setTextAlign(a).run();

  return (
    <div>
      <div className="pp-header">
        <strong>段落</strong>
        <button type="button" className="pp-close" onClick={onClose} title="关闭">
          <X size={12} />
        </button>
      </div>

      <Row label="对齐方式">
        <div className="pp-seg">
          <SegBtn
            active={!attrs.textAlign || attrs.textAlign === "left"}
            onClick={() => align("left")}
            title="左对齐"
          >
            <AlignLeft size={13} />
          </SegBtn>
          <SegBtn
            active={attrs.textAlign === "center"}
            onClick={() => align("center")}
            title="居中"
          >
            <AlignCenter size={13} />
          </SegBtn>
          <SegBtn
            active={attrs.textAlign === "right"}
            onClick={() => align("right")}
            title="右对齐"
          >
            <AlignRight size={13} />
          </SegBtn>
          <SegBtn
            active={attrs.textAlign === "justify"}
            onClick={() => align("justify")}
            title="两端对齐"
          >
            <AlignJustify size={13} />
          </SegBtn>
        </div>
      </Row>

      <Row label="行距">
        <select
          value={attrs.lineHeight ?? ""}
          onChange={(e) => apply({ lineHeight: e.target.value })}
        >
          <option value="">默认</option>
          <option value="1">单倍</option>
          <option value="1.15">1.15</option>
          <option value="1.5">1.5</option>
          <option value="2">2 倍</option>
          <option value="1.35">1.35 (Word 中文)</option>
          <optgroup label="固定值">
            <option value="pt:20">固定 20 磅</option>
            <option value="pt:22">固定 22 磅</option>
            <option value="pt:29">固定 29 磅（公文）</option>
          </optgroup>
          <optgroup label="最小值">
            <option value="atLeast:18">最小 18 磅（书籍）</option>
            <option value="atLeast:20">最小 20 磅（论文）</option>
            <option value="atLeast:22">最小 22 磅</option>
          </optgroup>
          {LINE_OPTIONS.has(attrs.lineHeight ?? "") ? null : (
            <option value={attrs.lineHeight ?? ""}>
              {lineHeightLabel(attrs.lineHeight ?? "")}
            </option>
          )}
        </select>
      </Row>

      <div className="pp-grid">
        <Row label="段前 (pt)">
          <Num
            value={attrs.spacingBefore}
            placeholder="0"
            onChange={(v) => apply({ spacingBefore: v })}
          />
        </Row>
        <Row label="段后 (pt)">
          <Num
            value={attrs.spacingAfter}
            placeholder="0"
            onChange={(v) => apply({ spacingAfter: v })}
          />
        </Row>
        <Row label="左缩进 (pt)">
          <Num
            value={attrs.indent}
            placeholder="0"
            onChange={(v) => apply({ indent: v })}
          />
        </Row>
        <Row label="首行缩进 (pt)">
          <Num
            value={attrs.indentFirst}
            placeholder="0"
            onChange={(v) => apply({ indentFirst: v })}
          />
        </Row>
      </div>

      <div className="pp-hint">提示：首行缩进 21pt 约等于中文 2 字符。</div>
    </div>
  );
}

function Row({ label, children }: { label: string; children: React.ReactNode }) {
  return (
    <label className="pp-row">
      <span>{label}</span>
      {children}
    </label>
  );
}

function SegBtn({
  active,
  onClick,
  title,
  children,
}: {
  active?: boolean;
  onClick: () => void;
  title: string;
  children: React.ReactNode;
}) {
  return (
    <button
      type="button"
      className={`pp-seg-btn ${active ? "active" : ""}`}
      title={title}
      onMouseDown={(e) => e.preventDefault()}
      onClick={onClick}
    >
      {children}
    </button>
  );
}

function Num({
  value,
  placeholder,
  onChange,
}: {
  value: number | undefined;
  placeholder?: string;
  onChange: (v: number | undefined) => void;
}) {
  return (
    <input
      type="number"
      min={0}
      step={1}
      value={value ?? ""}
      placeholder={placeholder}
      onChange={(e) => {
        const v = e.target.value;
        onChange(v === "" ? undefined : Math.max(0, Number(v)));
      }}
    />
  );
}
