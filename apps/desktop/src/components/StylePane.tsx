import { useEffect, useMemo, useState } from "react";
import type { Editor } from "@tiptap/react";
import {
  fitStyleId,
  headingLookupFromPaneStyles,
  type PaneStyle,
} from "@office-copilot/docx-io";
import { Eraser, X } from "lucide-react";
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

export function StylePane({
  editor,
  styles,
  onClose,
}: {
  editor: Editor;
  /** 本文档声明的段落样式清单，来自打开时的 styles.xml；空数组表示文档没有样式表。 */
  styles: PaneStyle[];
  onClose: () => void;
}) {
  const lookup = useMemo(() => headingLookupFromPaneStyles(styles), [styles]);
  const [appliedId, setAppliedId] = useState<string | null | undefined>(undefined);

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
          这份文档没有声明具名样式。样式来自 .docx 里的 styles.xml，新建的空白文档没有这一份。
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
    </aside>
  );
}
