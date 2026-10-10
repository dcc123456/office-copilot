import { useEffect, useState } from "react";
import type { Editor } from "@tiptap/react";
import { Bold, Italic, Underline, Strikethrough, X } from "lucide-react";
import {
  FONT_COLORS,
  FONT_FAMILIES,
  FONT_SIZES,
  HIGHLIGHT_COLORS,
} from "../editor/fontOptions.js";

/**
 * 「字体(F)…」对话框的内容，字段与 Word 的字体对话框对齐：
 * 中文字体 / 西文字体 / 字号 / 字形 / 颜色 / 字符间距。
 *
 * 写入一律走 textStyle 标记的 toggleTextStyle（字号是 CSS px，非半磅；
 * 字符间距是 CSS px，导出时由 docx-io 换算成 OOXML 的 w:spacing）。
 * 段落对象上的样式不受影响，所以清除字体只会去掉字符层格式。
 */

interface Styles {
  fontFamily?: string;
  fontFamilyWest?: string;
  fontSize?: string;
  letterSpacing?: string;
  color?: string;
  backgroundColor?: string;
}

export function FontFields({ editor, onClose }: { editor: Editor; onClose: () => void }) {
  const [style, setStyle] = useState<Styles>({});

  useEffect(() => {
    const read = () =>
      setStyle(editor.getAttributes("textStyle") as Styles);
    read();
    editor.on("selectionUpdate", read);
    editor.on("update", read);
    return () => {
      editor.off("selectionUpdate", read);
      editor.off("update", read);
    };
  }, [editor]);

  /** 只改给出的那一项，其余 textStyle 属性由 toggleTextStyle 自己保留。 */
  const patch = (key: keyof Styles, value: string | null) => {
    setStyle((s) => ({ ...s, [key]: value }));
    // 不 chain().focus()：字符间距是输入框，抢焦点会把后续按键落进文档。
    const chain = editor.chain() as unknown as {
      toggleTextStyle: (attrs: Record<string, string | null>) => { run: () => boolean };
    };
    if (!value && key === "fontFamily") {
      editor.chain().unsetMark("textStyle").run();
      return;
    }
    chain.toggleTextStyle({ [key]: value }).run();
  };

  return (
    <div>
      <div className="pp-header">
        <strong>字体</strong>
        <button type="button" className="pp-close" onClick={onClose} title="关闭">
          <X size={12} />
        </button>
      </div>

      <Row label="中文字体">
        <select
          value={style.fontFamily ?? ""}
          onChange={(e) => patch("fontFamily", e.target.value || null)}
        >
          {FONT_FAMILIES.map((f) => (
            <option key={f.label} value={f.value}>
              {f.label}
            </option>
          ))}
        </select>
      </Row>
      <Row label="西文字体">
        <select
          value={style.fontFamilyWest ?? ""}
          onChange={(e) => patch("fontFamilyWest", e.target.value || null)}
        >
          {FONT_FAMILIES.map((f) => (
            <option key={f.label} value={f.value}>
              {f.label}
            </option>
          ))}
        </select>
      </Row>
      <Row label="字号">
        <select
          value={style.fontSize ?? ""}
          onChange={(e) => patch("fontSize", e.target.value || null)}
        >
          {FONT_SIZES.map((s) => (
            <option key={s.label} value={s.value}>
              {s.label}
            </option>
          ))}
        </select>
      </Row>

      <Row label="字形">
        <div className="pp-seg">
          <SegBtn
            active={editor.isActive("bold")}
            title="加粗"
            onClick={() => editor.chain().focus().toggleBold().run()}
          >
            <Bold size={13} />
          </SegBtn>
          <SegBtn
            active={editor.isActive("italic")}
            title="倾斜"
            onClick={() => editor.chain().focus().toggleItalic().run()}
          >
            <Italic size={13} />
          </SegBtn>
          <SegBtn
            active={editor.isActive("underline")}
            title="下划线"
            onClick={() => editor.chain().focus().toggleUnderline().run()}
          >
            <Underline size={13} />
          </SegBtn>
          <SegBtn
            active={editor.isActive("strike")}
            title="删除线"
            onClick={() => editor.chain().focus().toggleStrike().run()}
          >
            <Strikethrough size={13} />
          </SegBtn>
        </div>
      </Row>

      <Row label="字体颜色">
        <Swatches
          colors={FONT_COLORS}
          value={style.color ?? ""}
          onPick={(c) => patch("color", c || null)}
        />
      </Row>
      <Row label="突出显示">
        <Swatches
          colors={HIGHLIGHT_COLORS}
          value={style.backgroundColor ?? ""}
          onPick={(c) => patch("backgroundColor", c || null)}
        />
      </Row>

      <Row label="字符间距 (px)">
        <input
          type="number"
          step={0.5}
          value={/^-?[\d.]+px$/.test(style.letterSpacing ?? "")
            ? Number(style.letterSpacing!.replace("px", ""))
            : ""}
          placeholder="0"
          onChange={(e) =>
            patch(
              "letterSpacing",
              e.target.value === "" ? null : `${Number(e.target.value)}px`,
            )
          }
        />
      </Row>

      <div className="pp-hint">作用于选区；未选中文字时只影响从这里继续输入的内容。</div>
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
  title,
  onClick,
  children,
}: {
  active?: boolean;
  title: string;
  onClick: () => void;
  children: React.ReactNode;
}) {
  return (
    <button
      type="button"
      className={`pp-seg-btn ${active ? "active" : ""}`}
      title={title}
      onClick={onClick}
    >
      {children}
    </button>
  );
}

/**
 * 一行色块 + 自定义取色器；空串代表清除该颜色。
 *
 * @param colors 预设色值（#rrggbb 或空串）。
 */
function Swatches({
  colors,
  value,
  onPick,
}: {
  colors: string[];
  value: string;
  onPick: (color: string) => void;
}) {
  return (
    <div className="ctx-swatches font-swatches">
      {colors.map((c) => (
        <button
          key={c || "none"}
          type="button"
          className={`ctx-swatch ${value && c && value.toLowerCase() === c.toLowerCase() ? "selected" : ""}`}
          title={c || "清除"}
          style={{
            background:
              c || "repeating-conic-gradient(#ccc 0% 25%, #fff 0% 50%) 50% / 8px 8px",
          }}
          onClick={() => onPick(c)}
        />
      ))}
      <input
        type="color"
        className="ctx-color-input"
        value={/^#[0-9a-f]{6}$/i.test(value) ? value : "#000000"}
        onChange={(e) => onPick(e.target.value)}
        title="自定义颜色"
      />
    </div>
  );
}
