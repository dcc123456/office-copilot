import type { Editor } from "@tiptap/react";
import {
  Bold,
  Italic,
  Underline as UnderlineIcon,
  Strikethrough,
  List,
  ListOrdered,
  Quote,
  Undo2,
  Redo2,
  AlignLeft,
  AlignCenter,
  AlignRight,
  AlignJustify,
  Table as TableIcon,
  Image as ImageIcon,
  Link as LinkIcon,
  Highlighter,
  Palette,
  CornerDownRight,
  CornerDownLeft,
  Paintbrush,
  FilePlus2,
  FolderOpen,
  Save,
  Download,
  Printer,
  Search,
  Minus,
  Settings2,
  SwatchBook,
  Pilcrow,
  Wand2,
  RotateCcw,
} from "lucide-react";

import { ParagraphPopover } from "./ParagraphPopover.js";
import { useAnchoredMenu } from "../hooks/useAnchoredMenu.js";
import { promptSetLink } from "./linkPrompt.js";
import { hasStoredFormat } from "../editor/contextMenu.js";
import {
  FONT_COLORS,
  FONT_FAMILIES,
  FONT_SIZES,
  HIGHLIGHT_COLORS,
} from "../editor/fontOptions.js";

interface ToolbarProps {
  editor: Editor;
  onInsertImage: () => void;
  onInsertTable: () => void;
  onAction: (action: ToolbarAction) => void;
  dirty?: boolean;
  /** 本文档还有可还原的美化排版记录时，「还原排版」才可点。 */
  canRevert?: boolean;
  /** 样式窗格是否已打开，决定「样式」按钮的高亮。 */
  stylesOpen?: boolean;
}

export type ToolbarAction =
  | "new"
  | "open"
  | "save"
  | "save-as"
  | "print"
  | "find-replace"
  | "page-setup"
  | "styles"
  | "beautify"
  | "revert-beautify";

function Group({ children }: { children: React.ReactNode }) {
  return <div className="tb-set">{children}</div>;
}

function Btn({
  active,
  disabled,
  onClick,
  title,
  children,
  className,
}: {
  active?: boolean;
  disabled?: boolean;
  onClick: () => void;
  title: string;
  children: React.ReactNode;
  className?: string;
}) {
  return (
    <button
      type="button"
      className={`icon ${active ? "active" : ""} ${className ?? ""}`}
      title={title}
      disabled={disabled}
      onMouseDown={(e) => e.preventDefault()}
      onClick={onClick}
    >
      {children}
    </button>
  );
}

export function Toolbar({
  editor,
  onInsertImage,
  onInsertTable,
  onAction,
  dirty,
  canRevert,
  stylesOpen,
}: ToolbarProps) {
  const setLink = () => void promptSetLink(editor);

  const painterActive = hasStoredFormat(editor);

  return (
    <div className="toolbar">
      <div className="tb-band">
        <Group>
          <Btn title="新建  ⌘N" onClick={() => onAction("new")}>
          <FilePlus2 size={16} />
        </Btn>
        <Btn title="打开  ⌘O" onClick={() => onAction("open")}>
          <FolderOpen size={16} />
        </Btn>
        <Btn
          title="保存  ⌘S"
          onClick={() => onAction("save")}
          className={dirty ? "dirty" : undefined}
        >
          <Save size={16} />
          {dirty ? <span className="dirty-dot" aria-hidden /> : null}
        </Btn>
        <Btn title="另存为…  ⌘⇧S" onClick={() => onAction("save-as")}>
          <Download size={16} />
        </Btn>
        <Btn title="打印 / 导出 PDF  ⌘P" onClick={() => onAction("print")}>
          <Printer size={16} />
        </Btn>
      </Group>

        <Group>
          <Btn title="撤销" onClick={() => editor.chain().focus().undo().run()} disabled={!editor.can().undo()}>
          <Undo2 size={16} />
        </Btn>
        <Btn title="重做" onClick={() => editor.chain().focus().redo().run()} disabled={!editor.can().redo()}>
          <Redo2 size={16} />
        </Btn>
        <Btn title="查找与替换  ⌘F" onClick={() => onAction("find-replace")}>
          <Search size={16} />
        </Btn>
        </Group>

        <div className="tb-spacer" />

        <Group>
          <Btn title="美化排版（模板）" onClick={() => onAction("beautify")}>
            <Wand2 size={16} />
          </Btn>
          <Btn
            title={canRevert ? "还原排版（撤销全部美化排版，含页边距/页眉页脚）" : "没有可还原的美化排版记录"}
            disabled={!canRevert}
            onClick={() => onAction("revert-beautify")}
          >
            <RotateCcw size={16} />
          </Btn>
          <Btn
            title="样式窗格（套用本文档的具名样式）"
            active={stylesOpen}
            onClick={() => onAction("styles")}
          >
            <SwatchBook size={16} />
          </Btn>
          <Btn title="页面设置" onClick={() => onAction("page-setup")}>
            <Settings2 size={16} />
          </Btn>
        </Group>
      </div>

      <div className="tb-band">
        <Group>
        <select
          className="tb-select tb-block-select"
          title="段落样式"
          value={
            editor.isActive("heading")
              ? `h${Number(editor.getAttributes("heading").level ?? 1)}`
              : "p"
          }
          onChange={(e) => {
            const v = e.target.value;
            if (v === "p") editor.chain().focus().setParagraph().run();
            else
              editor
                .chain()
                .focus()
                .toggleHeading({ level: Number(v.slice(1)) as 1 | 2 | 3 | 4 | 5 | 6 })
                .run();
          }}
        >
          <option value="p">正文</option>
          {[1, 2, 3, 4, 5, 6].map((l) => (
            <option key={l} value={`h${l}`}>
              标题 {l}
            </option>
          ))}
        </select>
        <select
          className="tb-select tb-font-select"
          title="字体"
          value={
            (editor.getAttributes("textStyle").fontFamily as string | undefined) ?? ""
          }
          onChange={(e) => {
            const v = e.target.value;
            const cmd = editor.chain().focus();
            if (!v) cmd.unsetMark("textStyle").run();
            else cmd.toggleTextStyle({ fontFamily: v }).run();
          }}
        >
          {FONT_FAMILIES.map((f) => (
            <option key={f.label} value={f.value}>
              {f.label}
            </option>
          ))}
        </select>
        <select
          className="tb-select tb-size-select"
          title="字号"
          value={
            (editor.getAttributes("textStyle").fontSize as string | undefined) ?? ""
          }
          onChange={(e) => {
            const v = e.target.value;
            if (!v) editor.chain().focus().unsetMark("textStyle").run();
            else editor.chain().focus().toggleTextStyle({ fontSize: v }).run();
          }}
        >
          {FONT_SIZES.map((s) => (
            <option key={s.label} value={s.value}>
              {s.label}
            </option>
          ))}
        </select>
        <ColorMenu
          icon={<Palette size={14} />}
          title="字体颜色"
          customLabel="颜色"
          colors={FONT_COLORS}
          value={(editor.getAttributes("textStyle").color as string) ?? ""}
          onPick={(c) =>
            c
              ? editor.chain().focus().toggleTextStyle({ color: c }).run()
              : editor.chain().focus().toggleTextStyle({ color: null }).run()
          }
        />
        <ColorMenu
          icon={<Highlighter size={14} />}
          title="文字底色"
          customLabel="底色"
          colors={HIGHLIGHT_COLORS}
          value={
            (editor.getAttributes("textStyle").backgroundColor as string) ?? ""
          }
          onPick={(c) =>
            c
              ? editor.chain().focus().toggleTextStyle({ backgroundColor: c }).run()
              : editor
                  .chain()
                  .focus()
                  .toggleTextStyle({ backgroundColor: null })
                  .run()
          }
        />
      </Group>

        <Group>
          <Btn title="加粗 (⌘B)" active={editor.isActive("bold")} onClick={() => editor.chain().focus().toggleBold().run()}>
          <Bold size={16} />
        </Btn>
        <Btn title="斜体" active={editor.isActive("italic")} onClick={() => editor.chain().focus().toggleItalic().run()}>
          <Italic size={16} />
        </Btn>
        <Btn title="下划线" active={editor.isActive("underline")} onClick={() => editor.chain().focus().toggleUnderline().run()}>
          <UnderlineIcon size={16} />
        </Btn>
        <Btn title="删除线" active={editor.isActive("strike")} onClick={() => editor.chain().focus().toggleStrike().run()}>
          <Strikethrough size={16} />
        </Btn>
        <Btn
          title="格式刷：复制当前样式，选中目标粘贴"
          onClick={() => {
            if (hasStoredFormat(editor)) editor.commands.cancelFormatPaint();
            else editor.commands.pickFormat();
          }}
          className={painterActive ? "active" : undefined}
        >
          <Paintbrush size={16} />
        </Btn>
        </Group>

        <Group>
          <Btn title="左对齐" active={editor.isActive({ textAlign: "left" })} onClick={() => editor.chain().focus().setTextAlign("left").run()}>
          <AlignLeft size={16} />
        </Btn>
        <Btn title="居中" active={editor.isActive({ textAlign: "center" })} onClick={() => editor.chain().focus().setTextAlign("center").run()}>
          <AlignCenter size={16} />
        </Btn>
        <Btn title="右对齐" active={editor.isActive({ textAlign: "right" })} onClick={() => editor.chain().focus().setTextAlign("right").run()}>
          <AlignRight size={16} />
        </Btn>
        <Btn title="两端对齐" active={editor.isActive({ textAlign: "justify" })} onClick={() => editor.chain().focus().setTextAlign("justify").run()}>
          <AlignJustify size={16} />
        </Btn>
        <Btn title="无序列表" active={editor.isActive("bulletList")} onClick={() => editor.chain().focus().toggleBulletList().run()}>
          <List size={16} />
        </Btn>
        <Btn title="有序列表" active={editor.isActive("orderedList")} onClick={() => editor.chain().focus().toggleOrderedList().run()}>
          <ListOrdered size={16} />
        </Btn>
        <Btn title="引用" active={editor.isActive("blockquote")} onClick={() => editor.chain().focus().toggleBlockquote().run()}>
          <Quote size={16} />
        </Btn>
        <Btn
          title="减少缩进  Shift+Tab"
          onClick={() => editor.chain().focus().liftListItem("listItem").run()}
          disabled={!editor.isActive("listItem")}
        >
          <CornerDownLeft size={16} />
        </Btn>
        <Btn
          title="增加缩进  Tab"
          onClick={() => editor.chain().focus().sinkListItem("listItem").run()}
          disabled={!editor.isActive("listItem")}
        >
          <CornerDownRight size={16} />
        </Btn>
        <ParagraphPopover editor={editor} />
        </Group>

        <Group>
          <Btn title="插入表格" onClick={onInsertTable}>
          <TableIcon size={16} />
        </Btn>
        <Btn title="插入图片" onClick={onInsertImage}>
          <ImageIcon size={16} />
        </Btn>
        <Btn title="插入链接" active={editor.isActive("link")} onClick={setLink}>
          <LinkIcon size={16} />
        </Btn>
        <Btn
          title="插入分页符"
          onClick={() =>
            editor
              .chain()
              .focus()
              .insertContent({ type: "pageBreak" })
              .run()
          }
        >
          <Pilcrow size={16} />
        </Btn>
        <Btn
          title="分割线"
          onClick={() => editor.chain().focus().setHorizontalRule().run()}
        >
          <Minus size={16} />
        </Btn>
        </Group>
      </div>
    </div>
  );
}

function ColorMenu({
  icon,
  title,
  customLabel,
  colors,
  value,
  onPick,
}: {
  icon: React.ReactNode;
  title: string;
  customLabel: string;
  colors: string[];
  value: string;
  onPick: (color: string) => void;
}) {
  const { anchorRef, rect, open, toggle, hide } = useAnchoredMenu(150);
  return (
    <div
      className="color-menu-wrap"
      ref={anchorRef}
      onBlur={(e) => {
        if (!e.currentTarget.contains(e.relatedTarget)) hide();
      }}
    >
      <button
        type="button"
        className="icon color-menu-btn"
        title={title}
        aria-expanded={open}
        onMouseDown={(e) => e.preventDefault()}
        onClick={toggle}
      >
        {icon}
        {value ? (
          <span className="color-underline" style={{ background: value }} />
        ) : null}
      </button>
      {open && rect && (
        <div className="color-menu" style={{ left: rect.left, top: rect.top }}>
          <div className="color-grid">
            {colors.map((c) => (
              <button
                key={c || "none"}
                type="button"
                className={
                  "color-dot" +
                  (value && c && sameColor(value, c) ? " selected" : "")
                }
                title={c || "清除"}
                onMouseDown={(e) => e.preventDefault()}
                onClick={() => {
                  onPick(c);
                  hide();
                }}
                style={{
                  background:
                    c ||
                    "repeating-conic-gradient(#ccc 0% 25%, #fff 0% 50%) 50% / 8px 8px",
                }}
              />
            ))}
          </div>
          {/* Keeps the document selection alive while the OS picker opens. */}
          <label className="color-custom" onMouseDown={(e) => e.preventDefault()}>
            <input
              type="color"
              value={toHexInput(value)}
              onChange={(e) => onPick(e.target.value)}
            />
            自定义{customLabel}
          </label>
        </div>
      )}
    </div>
  );
}

/** CSS colours arrive as #rrggbb from .docx and as rgb() from pasted HTML. */
function sameColor(a: string, b: string): boolean {
  return toHexInput(a).toLowerCase() === toHexInput(b).toLowerCase();
}

function toHexInput(value: string): string {
  const hex = /^#([0-9a-f]{3}|[0-9a-f]{6})$/i.exec(value);
  if (hex?.[1]) {
    const h = hex[1];
    return `#${h.length === 3 ? h.split("").map((c) => c + c).join("") : h.toLowerCase()}`;
  }
  const rgb = /^rgba?\(\s*(\d+)[\s,]+(\d+)[\s,]+(\d+)/i.exec(value);
  if (rgb) {
    return `#${rgb
      .slice(1, 4)
      .map((n) => Math.min(255, Number(n)).toString(16).padStart(2, "0"))
      .join("")}`;
  }
  return "#000000";
}
