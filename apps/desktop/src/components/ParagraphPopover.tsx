import { useEffect, useState } from "react";
import type { Editor } from "@tiptap/react";
import { AlignLeft, AlignCenter, AlignRight, AlignJustify, X } from "lucide-react";
import { useAnchoredMenu } from "../hooks/useAnchoredMenu.js";

/**
 * 段落设置面板，按 Word 段落对话框的分组排：
 *   - 常规：对齐方式、大纲级别
 *   - 缩进：左/右缩进 (pt)、「特殊格式 + 磅值」（首行缩进 / 悬挂缩进）
 *   - 间距：段前 / 段后 (pt)、行距 (single / 1.15 / 1.5 / double / fixed)
 *   - 分页与换行：孤行控制、与下段同页、段中不分页、段前分页、对齐到网格
 *
 * 数值经自定义命令 `setParagraphFormat` 写进当前选中的段落/标题，落在
 * editor/extensions/ParagraphFormatting.ts，再由 docx-io 往返成 OOXML 的
 * <w:spacing>、<w:ind>、<w:outlineLvl> 与 <w:keepNext/> 一族。右键菜单的
 * 「段落(P)…」直接复用 ParagraphFields。
 */

interface Attrs {
  textAlign?: string;
  outlineLevel?: number;
  lineHeight?: string;
  spacingBefore?: number;
  spacingAfter?: number;
  indent?: number;
  indentRight?: number;
  indentFirst?: number;
  keepWithNext?: boolean;
  keepLines?: boolean;
  widowControl?: boolean;
  snapToGrid?: boolean;
  pageBreakBefore?: boolean;
}

/** Word 段落对话框「特殊格式」的三个取值。 */
type SpecialFormat = "none" | "firstLine" | "hanging";

/** 段落属性里属于「分页与换行」的开关（对应 OOXML 的 CT_OnOff 元素）。 */
type SwitchKey =
  | "widowControl"
  | "keepWithNext"
  | "keepLines"
  | "pageBreakBefore"
  | "snapToGrid";

/** 开关的按钮文案沿用 Word 的原词，tip 是 hover 说明。 */
const SWITCH_FIELDS: { key: SwitchKey; label: string; tip: string }[] = [
  { key: "widowControl", label: "孤行控制", tip: "禁止把一段的首末单行留在上一页/下一页" },
  { key: "keepWithNext", label: "与下段同页", tip: "本段与下一段排在同一页" },
  { key: "keepLines", label: "段中不分页", tip: "本段的所有行排在同一页" },
  { key: "pageBreakBefore", label: "段前分页", tip: "在本段之前强制换页" },
  { key: "snapToGrid", label: "对齐到网格", tip: "如果定义了文档网格，则对齐到网格" },
];

/** Word「大纲级别」下拉。段落属性存 OOXML 的原值：0-8 = 1-9 级，9 = 正文文本。 */
const OUTLINE_OPTIONS: { value: number; label: string }[] = [
  ...[0, 1, 2, 3, 4, 5, 6, 7, 8].map((value) => ({
    value,
    label: `${value + 1} 级`,
  })),
  { value: 9, label: "正文文本" },
];

/** 特殊格式刚打开时给磅值一个初值：中文 2 字符（五号 10.5pt × 2）。 */
const DEFAULT_SPECIAL_PT = 21;

/**
 * 把 indentFirst 属性读成 Word 的「特殊格式」。
 *
 * @param indentFirst 首行缩进磅值；负数就是悬挂缩进（OOXML 的 w:hanging）。
 * @returns 无 / 首行缩进 / 悬挂缩进。
 */
export function specialFormatOf(indentFirst: number | undefined): SpecialFormat {
  if (!indentFirst) return "none";
  return indentFirst < 0 ? "hanging" : "firstLine";
}

/**
 * 把「特殊格式 + 磅值」这对外部字段写回单个 indentFirst 属性。
 *
 * @param kind 下拉选中的特殊格式。
 * @param magnitude 磅值输入框的正数（pt）；空值按 0 处理。
 * @returns 存进段落属性的值：无 = undefined，首行为正，悬挂为负。
 */
export function indentFirstFor(
  kind: SpecialFormat,
  magnitude: number | undefined,
): number | undefined {
  if (kind === "none") return undefined;
  const pt = magnitude && magnitude > 0 ? magnitude : DEFAULT_SPECIAL_PT;
  return kind === "hanging" ? -pt : pt;
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
      // getAttributes 对不活跃的类型返回 {}（真值），所以必须按 isActive 选类型，
      // 不能写 `getAttributes("paragraph") || getAttributes("heading")`：那样选中
      // 标题时读到的一直是空的段落属性，面板永远是空白。
      const type = editor.isActive("heading") ? "heading" : "paragraph";
      const a: Attrs = editor.getAttributes(type) as Attrs;
      setAttrs({
        textAlign: a.textAlign,
        outlineLevel: a.outlineLevel,
        lineHeight: a.lineHeight ?? "",
        spacingBefore: a.spacingBefore,
        spacingAfter: a.spacingAfter,
        indent: a.indent,
        indentRight: a.indentRight,
        indentFirst: a.indentFirst,
        keepWithNext: a.keepWithNext,
        keepLines: a.keepLines,
        widowControl: a.widowControl,
        snapToGrid: a.snapToGrid,
        pageBreakBefore: a.pageBreakBefore,
      });
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

  const special = specialFormatOf(attrs.indentFirst);
  const magnitude = attrs.indentFirst == null ? undefined : Math.abs(attrs.indentFirst);

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

      <Row label="大纲级别">
        <select
          value={attrs.outlineLevel ?? 9}
          onChange={(e) => apply({ outlineLevel: Number(e.target.value) })}
        >
          {OUTLINE_OPTIONS.map((o) => (
            <option key={o.value} value={o.value}>
              {o.label}
            </option>
          ))}
        </select>
      </Row>

      <div className="pp-hint">缩进（磅）</div>
      <div className="pp-grid">
        <Row label="左侧 (pt)">
          <Num
            value={attrs.indent}
            placeholder="0"
            onChange={(v) => apply({ indent: v })}
          />
        </Row>
        <Row label="右侧 (pt)">
          <Num
            value={attrs.indentRight}
            placeholder="0"
            onChange={(v) => apply({ indentRight: v })}
          />
        </Row>
        <Row label="特殊格式">
          <select
            value={special}
            onChange={(e) =>
              apply({
                indentFirst: indentFirstFor(
                  e.target.value as SpecialFormat,
                  magnitude,
                ),
              })
            }
          >
            <option value="none">无</option>
            <option value="firstLine">首行缩进</option>
            <option value="hanging">悬挂缩进</option>
          </select>
        </Row>
        <Row label="磅值 (pt)">
          <Num
            value={magnitude}
            placeholder="0"
            onChange={(v) => apply({ indentFirst: indentFirstFor(special, v) })}
          />
        </Row>
      </div>

      <div className="pp-hint">间距（磅）</div>
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
      </div>

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

      <div className="pp-hint">分页与换行</div>
      <div className="pp-grid">
        {SWITCH_FIELDS.map((sw) => (
          <button
            key={sw.key}
            type="button"
            className={`pp-seg-btn ${attrs[sw.key] ? "active" : ""}`}
            title={`${sw.tip}（点击切换）`}
            onMouseDown={(e) => e.preventDefault()}
            onClick={() => apply({ [sw.key]: !attrs[sw.key] } as Partial<Attrs>)}
          >
            {sw.label}
          </button>
        ))}
      </div>

      <div className="pp-hint">
        提示：磅值 21pt 约等于中文 2 字符；悬挂缩进会让首行向左凸出。
      </div>
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
