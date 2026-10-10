import { useEffect, useMemo, useState } from "react";
import type { Editor } from "@tiptap/react";
import {
  fitStyleId,
  headingLookupFromPaneStyles,
  nextStyleId,
  type PaneStyle,
  type StyleEdit,
} from "@office-copilot/docx-io";
import { Eraser, X } from "lucide-react";
import {
  captureStyleFromSelection,
  type CaptureFailureCode,
} from "../editor/styleCapture.js";
import { toast } from "./ui/Toast.js";
import "../styles/stylePane.css";

/** Word 内建样式的英文名 → 界面显示名。样式表里已写中文名（"标题 1"）的用原值。 */
const BUILTIN_LABELS: Record<string, string> = {
  normal: "正文",
  "heading 1": "标题 1",
  "heading 2": "标题 2",
  "heading 3": "标题 3",
  "heading 4": "标题 4",
  "heading 5": "标题 5",
  "heading 6": "标题 6",
  title: "标题",
  subtitle: "副标题",
  quote: "引用",
  "intense quote": "明显引用",
  caption: "题注",
  "list paragraph": "列表段落",
  "list bullet": "列表项目符号",
  "list number": "列表编号",
};

/**
 * 窗格上一行该显示什么名字。
 *
 * @param style `paneStyles()` 的一条。
 * @returns 内建样式的中文名；其余（用户自己建的样式）用 `w:name` 原值。
 */
function labelOf(style: PaneStyle): string {
  return BUILTIN_LABELS[style.name.toLowerCase()] ?? style.name;
}

/**
 * 套用或清除一条具名样式，只作用于当前选区涉及的块。
 *
 * 规则本身在 `ParagraphFormatting` 的 `applyNamedStyle` 命令里（引擎层），窗格只装配：
 * 样式引用 + 正文/标题级别的转换。直接格式（红字、加粗、缩进）永远压过样式，
 * 把样式解析成具体属性再盖到段落上会毁掉用户的直接格式，所以只写引用。
 *
 * @param editor 目标编辑器。
 * @param style 要套用的样式；undefined = 清除样式（不写引用，Word 里等同该文档的默认段落样式）。
 */
function applyStyle(editor: Editor, style: PaneStyle | undefined): void {
  editor
    .chain()
    .focus()
    .applyNamedStyle(style?.styleId ?? null, style?.headingLevel)
    .run();
}

/**
 * 「新建 / 更新样式」抓到格式之后，把样式引用挂到选区那些块上，**不改块类型**。
 *
 * 不能用 `applyNamedStyle`：那条命令按 Word 的套用规则会把标题段变成正文段
 * （headingLevel 缺省 = paragraph）。这里只是"这段用的就是这条样式"，
 * 段落在编辑器里是什么还是什么。
 *
 * @param editor 目标编辑器。
 * @param styleId 要写进段落 attrs 的样式 id。
 */
function applyStyleReference(editor: Editor, styleId: string): void {
  editor.chain().focus().setParagraphFormat({ styleId }).run();
}

/**
 * `captureStyleFromSelection` 的 detail（段落 attrs 键 / 标记类型名）→ Word 对话框里的词。
 * 键名与 `CAPTURABLE_PARAGRAPH_KEYS`、`SWITCH_MARK_TYPES` 一一对应。
 */
const FORMAT_LABELS: Record<string, string> = {
  textAlign: "对齐方式",
  pageBreakBefore: "段前分页",
  keepWithNext: "与下段同页",
  keepLines: "段中不分页",
  widowControl: "孤行控制",
  snapToGrid: "对齐到网格",
  outlineLevel: "大纲级别",
  spacingBefore: "段前距",
  spacingAfter: "段后距",
  indent: "左缩进",
  indentRight: "右缩进",
  indentFirst: "首行缩进",
  lineHeight: "行距",
  bold: "加粗",
  italic: "斜体",
  underline: "下划线",
  strike: "删除线",
  code: "代码",
  textStyle: "字符格式",
};

/**
 * 抓取失败时的人话（AGENTS §5.7：引擎只给码与属性名，文案在渲染层拼）。
 *
 * @param code `captureStyleFromSelection` 的失败码。
 * @param detail 冲突的那个段落属性键或标记类型名，可能没有。
 * @returns 提示给用户的中文句子。
 */
function refusalText(code: CaptureFailureCode, detail?: string): string {
  const label = detail ? (FORMAT_LABELS[detail] ?? detail) : "格式";
  if (code === "no-block") return "选区里没有段落或标题，抓不到能写成样式的格式。";
  if (code === "mixed-paragraph-format")
    return `选区里各段的「${label}」不一致，一条样式定义只能记一份，请先只选格式相同的段落。`;
  return `选区里的「${label}」只覆盖部分文字（或两处取值不同），样式定义不能把它拍平成整段生效。`;
}

export function StylePane({
  editor,
  styles,
  onEdit,
  onClose,
}: {
  editor: Editor;
  /**
   * 本文档声明的段落样式清单：打开时来自 styles.xml，之后由主进程的预览投影刷新，
   * 所以窗格里一定包含刚新建的那条。空数组表示文档还没有样式表。
   */
  styles: PaneStyle[];
  /**
   * 交出一条攒着的新建/修改编辑。落盘发生在保存时（`SavePayload.styles`），
   * 预览由调用方走主进程同一个合并函数，窗格自己不再拼一份样式解析。
   */
  onEdit: (edit: StyleEdit) => void;
  onClose: () => void;
}) {
  const lookup = useMemo(() => headingLookupFromPaneStyles(styles), [styles]);
  const [appliedId, setAppliedId] = useState<string | null | undefined>(undefined);
  // 新建样式的名字输入框；清空表示没有要建的意思，按钮随之禁用。
  const [newName, setNewName] = useState("");

  useEffect(() => {
    const read = () => {
      // 与段落设置面板同一个坑：getAttributes 对不活跃的类型返回 {}（真值），
      // 必须按 isActive 选类型，否则选中标题时读到的是空的段落属性。
      const isHeading = editor.isActive("heading");
      const attrs = editor.getAttributes(isHeading ? "heading" : "paragraph") as {
        styleId?: string | null;
        level?: number;
      };
      // 高亮跟写侧同一条规则：段落刚改过级别时，attrs 里的引用还是上一级的。
      setAppliedId(
        fitStyleId(attrs.styleId, isHeading ? Number(attrs.level ?? 1) : undefined, lookup),
      );
    };
    read();
    editor.on("selectionUpdate", read);
    editor.on("update", read);
    return () => {
      editor.off("selectionUpdate", read);
      editor.off("update", read);
    };
  }, [editor, lookup]);

  // 没有引用（清除过 / 没建模过）时 Word 用的就是默认段落样式，窗格照着高亮它。
  const defaultId = styles.find((style) => style.isDefault)?.styleId;
  const shownId = appliedId == null ? defaultId : appliedId;
  const shown = styles.find((style) => style.styleId === shownId);

  /**
   * 抓当前选区的格式，失败就提示并交回 undefined（Word 在这种情况下直接拒绝更新样式）。
   *
   * @returns 抓到的小 `StyleEdit`（只有段落与标记，名字与 id 由调用方补）。
   */
  function captureSelection(): Omit<StyleEdit, "name" | "styleId"> | undefined {
    const captured = captureStyleFromSelection(editor);
    if (captured.ok) return captured.edit;
    toast(refusalText(captured.code, captured.detail), "error");
    return undefined;
  }

  /** Word 的「新建样式」：起名 → 抓选区格式 → 分配 id → 把引用挂到选区 → 交出一条编辑。 */
  function createStyle(): void {
    const name = newName.trim();
    if (!name) return;
    const captured = captureSelection();
    if (!captured) return;
    const styleId = nextStyleId(
      new Set(styles.map((style) => style.styleId)),
      name,
    );
    applyStyleReference(editor, styleId);
    onEdit({ ...captured, styleId, name });
    setNewName("");
    toast(`已新建样式「${name}」，保存后写进文档。`, "success");
  }

  /**
   * Word 的「更新样式以匹配所选内容」：改的是**已声明的那条**定义，名字不动。
   *
   * 高亮落在默认样式上时（选区没写引用）也会走到这里，等于把正文段落收编进那条样式——
   * 与 Word 同一个行为，按钮文案里带上样式名让用户看得见改的是谁。
   */
  function updateShownStyle(): void {
    if (!shown) return;
    const captured = captureSelection();
    if (!captured) return;
    applyStyleReference(editor, shown.styleId);
    onEdit({ ...captured, styleId: shown.styleId, name: shown.name });
    toast(`已按所选内容更新样式「${labelOf(shown)}」，保存后写进文档。`, "success");
  }

  return (
    <aside className="sidebar style-pane">
      <div className="panel-header">
        <h2>样式</h2>
        <button
          type="button"
          className="icon"
          title="清除样式（退回正文）"
          onClick={() => applyStyle(editor, undefined)}
        >
          <Eraser size={14} />
        </button>
        <button type="button" className="icon" title="关闭样式窗格" onClick={onClose}>
          <X size={14} />
        </button>
      </div>

      {styles.length === 0 ? (
        <div className="outline-empty">
          这份文档还没有声明具名样式。选中要模仿的格式，在下面起个名字就能新建一条。
        </div>
      ) : (
        <div className="style-list">
          {styles.map((style) => (
            <button
              key={style.styleId}
              type="button"
              className={`style-item${shownId === style.styleId ? " active" : ""}`}
              title={`${style.name} · ${style.styleId}`}
              onClick={() => applyStyle(editor, style)}
            >
              {labelOf(style)}
            </button>
          ))}
        </div>
      )}

      <div className="style-create">
        <input
          className="style-name-field"
          value={newName}
          placeholder="新样式名"
          aria-label="新样式名"
          onChange={(e) => setNewName(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === "Enter") createStyle();
          }}
        />
        <button
          type="button"
          className="style-action"
          disabled={!newName.trim()}
          onClick={createStyle}
        >
          新建样式
        </button>
        <button
          type="button"
          className="style-action"
          disabled={!shown}
          title={shown ? `把所选内容的格式写进「${labelOf(shown)}」的定义` : "先选中一条样式"}
          onClick={updateShownStyle}
        >
          {shown ? `更新「${labelOf(shown)}」以匹配所选内容` : "更新样式以匹配所选内容"}
        </button>
      </div>
    </aside>
  );
}
