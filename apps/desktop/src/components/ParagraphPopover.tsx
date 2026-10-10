import { useEffect, useState } from "react";
import type { Editor } from "@tiptap/react";
import type {
  DropCap,
  DropCapPosition,
  TabStop,
  TabStopAlignment,
  TabStopLeader,
} from "@office-copilot/docx-io";
import { cmToPt, ptToCm } from "@office-copilot/docx-io";
import { AlignLeft, AlignCenter, AlignRight, AlignJustify, X } from "lucide-react";
import { useAnchoredMenu } from "../hooks/useAnchoredMenu.js";
import {
  TAB_ALIGNMENTS,
  TAB_LEADERS,
  alignmentLabelOf,
  stopLabelOf,
} from "../editor/tabStopOptions.js";
import {
  DEFAULT_DROP_CAP_LINES,
  sameTabStopPosition,
} from "../editor/extensions/ParagraphFormatting.js";
import "../styles/tabStops.css";

/**
 * 段落设置面板，按 Word 段落对话框的分组排：
 *   - 常规：对齐方式、大纲级别
 *   - 缩进：左/右缩进 (pt)、「特殊格式 + 磅值」（首行缩进 / 悬挂缩进）
 *   - 间距：段前 / 段后 (pt)、行距 (single / 1.15 / 1.5 / double / fixed)
 *   - 分页与换行：孤行控制、与下段同页、段中不分页、段前分页、对齐到网格
 *   - 首字下沉：位置（无 / 下沉 / 首字突出到页边距外）+ 行数
 *   - 制表位：清单 + 位置/对齐/前导符 + 设置/清除/全部清除（Word 的"制表位…"按钮）
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
  tabStops?: TabStop[];
  dropCap?: DropCap | null;
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

/** Word「首字下沉」对话框的三个位置，取值直接沿用 `w:dropCap`。 */
type DropCapOption = "none" | DropCapPosition;

const DROP_CAP_OPTIONS: { value: DropCapOption; label: string }[] = [
  { value: "none", label: "无" },
  { value: "drop", label: "下沉" },
  { value: "margin", label: "突出到页边距外" },
];

/**
 * 把段落属性读成下拉的当前项。
 *
 * @param cap 首字下沉属性；undefined 是"这一段没建模过"，null 是用户选了"无"。
 * @returns 下拉的取值。两种"没有下沉"都显示成"无"，界面不需要再分一层。
 */
export function dropCapOptionOf(cap: DropCap | null | undefined): DropCapOption {
  return cap == null ? "none" : cap.position;
}

/**
 * 把「位置 + 行数」这对控件的值写回段落属性。
 *
 * @param option 下拉选中的位置。
 * @param lines 行数输入框（≥1 的整数）；空或小于 1 时按 Word 的缺省 3 行。
 * @returns 存进段落属性的值："无"是 null（导出删掉那条纯下沉的 `w:framePr`），
 *   另外两个一定带显式行数——`w:lines` 的 Schema 缺省是 1 行，留空会让 Word 摆一个一行的首字。
 */
export function dropCapFor(
  option: DropCapOption,
  lines: number | undefined,
): DropCap | null {
  if (option === "none") return null;
  const rows = lines && lines >= 1 ? Math.floor(lines) : DEFAULT_DROP_CAP_LINES;
  return { position: option, lines: rows };
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
  // Word 把制表位放在另一个对话框里；这里折进面板底部展开——多开一层窗口会抢走
  // 编辑器选区，而制表位恰恰要求用户看得见自己正在给哪一段定位。
  const [tabStopsOpen, setTabStopsOpen] = useState(false);

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
        tabStops: a.tabStops,
        dropCap: a.dropCap,
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
  const dropCapOption = dropCapOptionOf(attrs.dropCap);

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

      <div className="pp-hint">首字下沉</div>
      <div className="pp-grid">
        <Row label="位置">
          <select
            value={dropCapOption}
            onChange={(e) =>
              apply({
                dropCap: dropCapFor(
                  e.target.value as DropCapOption,
                  attrs.dropCap?.lines,
                ),
              })
            }
          >
            {DROP_CAP_OPTIONS.map((o) => (
              <option key={o.value} value={o.value}>
                {o.label}
              </option>
            ))}
          </select>
        </Row>
        <Row label="行数">
          <Num
            value={attrs.dropCap?.lines}
            placeholder={String(DEFAULT_DROP_CAP_LINES)}
            disabled={dropCapOption === "none"}
            onChange={(v) => apply({ dropCap: dropCapFor(dropCapOption, v) })}
          />
        </Row>
      </div>

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

      <div className="pp-hint">制表位</div>
      <div className="pp-grid">
        <button
          type="button"
          className="pp-seg-btn"
          title="展开/收起制表位编辑区"
          onMouseDown={(e) => e.preventDefault()}
          onClick={() => setTabStopsOpen((open) => !open)}
        >
          {tabStopsOpen ? "收起" : "制表位…"}
        </button>
      </div>
      {tabStopsOpen && (
        <TabStopSection
          stops={attrs.tabStops}
          onSet={(stop) => editor.commands.setTabStop(stop)}
          onClear={(positionPt) => editor.commands.clearTabStop(positionPt)}
          onClearAll={() => editor.commands.clearAllTabStops()}
        />
      )}

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
  disabled,
  onChange,
}: {
  value: number | undefined;
  placeholder?: string;
  /** Word 在"无"这一档把行数框置灰，面板保持同样的形状而不是把行藏掉。 */
  disabled?: boolean;
  onChange: (v: number | undefined) => void;
}) {
  return (
    <input
      type="number"
      min={0}
      step={1}
      disabled={disabled}
      value={value ?? ""}
      placeholder={placeholder}
      onChange={(e) => {
        const v = e.target.value;
        onChange(v === "" ? undefined : Math.max(0, Number(v)));
      }}
    />
  );
}

/**
 * 把「位置」录入框里的厘米换成段落属性用的磅。
 *
 * @param raw 厘米原文（Word 的对话框按两位小数录入）；空串、负数、非数字都不算一个位置。
 * @returns 磅值（0.1pt 精度，换算出自 docx-io）；读不出来时 null，界面据此禁用「设置」。
 */
export function tabStopPositionPt(raw: string): number | null {
  const text = raw.trim();
  if (!text) return null;
  const cm = Number(text);
  if (!Number.isFinite(cm) || cm < 0) return null;
  return cmToPt(cm);
}

/**
 * 「制表位…」编辑区：清单 + Word 对话框的位置/对齐/前导符 + 设置/清除/全部清除。
 *
 * @param stops 当前段落的制表位清单；undefined 是"这一段没建模过"，清单显示成空。
 * @param onSet 设一条：同位置整条替换，这一段在别处的制表位照旧留着。
 * @param onClear 按位置清一条。
 * @param onClearAll 全部清除：写入空清单，导出因此删掉段落自己的 `w:tabs`。
 */
function TabStopSection({
  stops,
  onSet,
  onClear,
  onClearAll,
}: {
  stops: TabStop[] | undefined;
  onSet: (stop: TabStop) => void;
  onClear: (positionPt: number) => void;
  onClearAll: () => void;
}) {
  const [position, setPosition] = useState("");
  const [alignment, setAlignment] = useState<TabStopAlignment>("left");
  const [leader, setLeader] = useState<TabStopLeader>("none");

  const list = stops ?? [];
  const positionPt = tabStopPositionPt(position);
  // Word 也只在真有这一位时才让「清除」可点，按下去没东西可清会让人以为面板坏了。
  const canClear =
    positionPt !== null &&
    list.some((stop) => sameTabStopPosition(stop.positionPt, positionPt));

  /** 载入清单里的一条到录入框（Word：点列表项等于把它摆回上面三个控件）。 */
  const pick = (stop: TabStop) => {
    setPosition(ptToCm(stop.positionPt).toFixed(2));
    setAlignment(stop.alignment);
    setLeader(stop.leader ?? "none");
  };

  return (
    <div className="ts-section">
      <div className="ts-stops">
        {list.length ? (
          list.map((stop) => (
            <button
              key={stop.positionPt}
              type="button"
              className="ts-stop-row"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => pick(stop)}
            >
              {stopLabelOf(stop)}
            </button>
          ))
        ) : (
          <div className="ts-empty">（这一段没有制表位）</div>
        )}
      </div>

      <Row label="位置 (cm)">
        <input
          type="number"
          min={0}
          step={0.01}
          value={position}
          placeholder="0.00"
          onChange={(e) => setPosition(e.target.value)}
        />
      </Row>

      <Row label="对齐方式">
        <select
          value={alignment}
          onChange={(e) => setAlignment(e.target.value as TabStopAlignment)}
        >
          {TAB_ALIGNMENTS.map((o) => (
            <option key={o.value} value={o.value}>
              {o.label}
            </option>
          ))}
          {/* 编号生成的列表制表位：显示得出来，但界面不给造一条（§排名 5 取舍）。 */}
          {TAB_ALIGNMENTS.some((o) => o.value === alignment) ? null : (
            <option value={alignment} disabled>
              {alignmentLabelOf(alignment)}
            </option>
          )}
        </select>
      </Row>

      <Row label="前导符">
        <select value={leader} onChange={(e) => setLeader(e.target.value as TabStopLeader)}>
          {TAB_LEADERS.map((o) => (
            <option key={o.value} value={o.value}>
              {o.label}
            </option>
          ))}
        </select>
      </Row>

      <div className="ts-actions">
        <button
          type="button"
          className="pp-seg-btn"
          disabled={positionPt === null}
          onMouseDown={(e) => e.preventDefault()}
          onClick={() => {
            if (positionPt === null) return;
            onSet({
              positionPt,
              alignment,
              // 「无」就是不写 w:leader：Word 自己也是省略的，写成 "none" 会让
              // 同一份文档两次保存的字节对不上。
              ...(leader === "none" ? {} : { leader }),
            });
          }}
        >
          设置
        </button>
        <button
          type="button"
          className="pp-seg-btn"
          disabled={!canClear}
          onMouseDown={(e) => e.preventDefault()}
          onClick={() => {
            if (positionPt !== null) onClear(positionPt);
          }}
        >
          清除
        </button>
        <button
          type="button"
          className="pp-seg-btn"
          disabled={!list.length}
          onMouseDown={(e) => e.preventDefault()}
          onClick={onClearAll}
        >
          全部清除
        </button>
      </div>

      <div className="pp-hint">
        位置按厘米录入、以磅存进段落属性，作用于选区里的每一段。编辑器里的 Tab 仍按浏览器
        默认宽度显示，制表位与前导符在导出的 .docx 里生效。
      </div>
    </div>
  );
}
