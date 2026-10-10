import type { Editor } from "@tiptap/core";
import type { Mark, Node as PMNode } from "@tiptap/pm/model";
import type { PMMark, PPrAttrs, StyleEdit } from "@office-copilot/docx-io";

/**
 * 「从选区抓格式」——Word 的"新建样式 / 更新样式以匹配所选内容"的前半段
 * （`docs/word-parity-spec.md` 排名 2 ⑥）。
 *
 * 抓出来的是**编辑器那套结构化数据**（`PPrAttrs` + `PMMark[]`），不是 XML：单位与颜色的
 * 换算只允许发生在 docx-io 进出口（AGENTS §2.1），样式定义的 XML 由主进程那边的
 * `DocxFile.applyStyleEdits` 生成。
 *
 * 与格式刷（`extensions/FormatPainter.ts`）的取舍不同：格式刷按 Word 的宽松规则取第一个
 * 字符，这里要往文档里写一条长期存在的样式定义，所以选区内任何一处不一致都必须拒绝并
 * 说明差在哪个属性上（AGENTS §2.3——静默拍平会盖掉用户只给部分文字加的红字与加粗）。
 */

/** 能从段落 attrs 里取走的建模键；`styleId` 与 `rawPpr` 故意不在其中。 */
const CAPTURABLE_PARAGRAPH_KEYS = [
  "textAlign",
  "pageBreakBefore",
  "keepWithNext",
  "keepLines",
  "widowControl",
  "snapToGrid",
  "outlineLevel",
  "spacingBefore",
  "spacingAfter",
  "indent",
  "indentRight",
  "indentFirst",
  "lineHeight",
  "tabStops",
] as const satisfies readonly (keyof PPrAttrs)[];

/** 无属性的开关标记；`link` 不在这里——超链接是内容，Word 也不把它长在段落样式上。 */
const SWITCH_MARK_TYPES = ["bold", "italic", "underline", "strike", "code"] as const;

/** 样式表建模得了的标记类型名，其余（高亮之类留给后面排名）不参与抓取。 */
const CAPTURABLE_MARK_NAMES = new Set<string>([
  ...SWITCH_MARK_TYPES,
  "textStyle",
  "vertAlign",
]);

/** 抓取失败的原因码；渲染层按码组装中文文案（AGENTS §5.7：引擎不说人话）。 */
export type CaptureFailureCode = "no-block" | "mixed-paragraph-format" | "mixed-run-format";

/** 成功时是这次抓到的格式（够拼一条 `StyleEdit`），失败时是原因码与冲突的属性/标记名。 */
export type CaptureResult =
  | { ok: true; edit: Omit<StyleEdit, "name" | "styleId"> }
  | { ok: false; code: CaptureFailureCode; detail?: string };

/** 只有段落与标题承载段落格式；表格、图片之类不参与。 */
function isFormatBlock(node: PMNode): boolean {
  return node.type.name === "paragraph" || node.type.name === "heading";
}

/**
 * 把一个块的 attrs 收成可抓取的子集。
 *
 * @param node 段落或标题节点。
 * @returns 只含写过且值不是 undefined/null 的建模键（段前段后与缩进是 pt，行距是倍数或
 *   `pt:N` / `atLeast:N`，对齐是 CSS 词）；空对象表示这个块没有可抓的段落格式。
 */
function paragraphFormatOf(node: PMNode): PPrAttrs {
  const captured: Record<string, unknown> = {};
  const attrs = node.attrs as Record<string, unknown>;
  for (const key of CAPTURABLE_PARAGRAPH_KEYS) {
    const value = attrs[key];
    if (value !== undefined && value !== null) captured[key] = value;
  }
  return captured as PPrAttrs;
}

/**
 * 把一条标记转成可比较的签名。
 *
 * @param mark ProseMirror 标记。
 * @returns `类型{属性按名字排序}`；去掉 undefined/null 属性，这样"没写颜色"与
 *   "颜色为 null"算同一条，不会因为键序或空值误判成冲突。
 */
function markSignature(mark: Mark): string {
  const attrs = mark.attrs as Record<string, unknown>;
  const kept = Object.keys(attrs)
    .filter((key) => attrs[key] !== undefined && attrs[key] !== null)
    .sort()
    .map((key) => `${key}=${String(attrs[key])}`);
  return `${mark.type.name}{${kept.join(",")}}`;
}

/**
 * 把 ProseMirror 标记还原成 docx-io 认识的标记对象。
 *
 * @param mark 选区里的标记，调用方保证类型名在 `CAPTURABLE_MARK_NAMES` 里。
 * @returns 开关标记只留类型名；`textStyle` 带字符串属性（颜色 #rrggbb、字号 CSS px，
 *   都是编辑器里本来的单位）。
 */
function toPmMark(mark: Mark): PMMark {
  if (mark.type.name === "textStyle") {
    const attrs = mark.attrs as Record<string, unknown>;
    const kept: Record<string, string> = {};
    for (const key of ["fontFamily", "fontFamilyWest", "fontSize", "letterSpacing", "color", "backgroundColor"]) {
      const value = attrs[key];
      if (typeof value === "string" && value.trim()) kept[key] = value;
    }
    return { type: "textStyle", attrs: kept };
  }
  // 上标与下标共用一个标记，样式定义里得记住是哪一档，只留类型名会把它写成上标。
  if (mark.type.name === "vertAlign") {
    const val = mark.attrs.val;
    return {
      type: "vertAlign",
      attrs: { val: val === "subscript" ? "subscript" : "superscript" },
    };
  }
  return { type: mark.type.name as (typeof SWITCH_MARK_TYPES)[number] };
}

/**
 * 两段文字的标记哪里不一样。
 *
 * @param first 已抓取那段文字的签名列表。
 * @param second 当前这段文字的签名列表。
 * @returns 冲突的标记类型名，用于告诉用户"这段里加粗只覆盖了一部分文字"。
 */
function conflictingMarkName(first: readonly string[], second: readonly string[]): string {
  const inFirst = new Set(first);
  const inSecond = new Set(second);
  const odd = [...first, ...second].find(
    (signature) => !inSecond.has(signature) || !inFirst.has(signature),
  );
  return (odd ?? first[0] ?? "").replace(/\{.*$/, "");
}

/**
 * 收集选区文字上的标记，并要求它们处处一致。
 *
 * @param editor 当前编辑器；空选区按光标处的标记算（Word 也允许光标状态下更新样式）。
 * @returns `{ marks }` 一致时的标记列表（可为空，表示这段文字没有行内格式）；
 *   `{ conflict }` 表示同一段里有的字加了这个标记、有的字没有，或两处属性值不同。
 */
function collectMarks(editor: Editor): { marks?: PMMark[]; conflict?: string } {
  const { from, to, empty } = editor.state.selection;
  const signaturesOf = (marks: readonly Mark[]): string[] =>
    marks.filter((mark) => CAPTURABLE_MARK_NAMES.has(mark.type.name)).map(markSignature).sort();

  let captured: PMMark[] | undefined;
  let capturedSignatures: string[] | undefined;
  let conflict: string | undefined;

  const consider = (marks: readonly Mark[]): boolean => {
    const signatures = signaturesOf(marks);
    if (capturedSignatures && signatures.join("|") !== capturedSignatures.join("|")) {
      conflict = conflictingMarkName(capturedSignatures, signatures);
      return false;
    }
    capturedSignatures = signatures;
    captured = marks
      .filter((mark) => CAPTURABLE_MARK_NAMES.has(mark.type.name))
      .map(toPmMark);
    return true;
  };

  if (empty) {
    consider(editor.state.selection.$from.marks());
  } else {
    editor.state.doc.nodesBetween(from, to, (node) => {
      if (conflict) return false;
      if (!node.isText) return true;
      return consider(node.marks);
    });
  }
  return conflict ? { conflict } : { marks: captured ?? [] };
}

/**
 * 两处段落属性的取值算不算同一个。
 *
 * 标量直接比值；制表位是对象数组，PM 每个节点各持一份，按引用比会把两条一模一样
 * 的清单判成冲突，所以数组按内容比。
 *
 * @param first 已抓到的值。
 * @param second 当前这个块在同一键上的值。
 * @returns 是否一致（不一致就要拒绝并报告是哪个属性）。
 */
function sameFormatValue(first: unknown, second: unknown): boolean {
  return first === second || JSON.stringify(first) === JSON.stringify(second);
}

/**
 * 从选区抓出一份可以写成样式定义的格式。
 *
 * @param editor 装配了 `EDITOR_EXTENSIONS` 的编辑器实例；选区决定作用范围。
 * @returns `{ ok: true, edit }`：`edit.paragraph` 是段落 attrs（只含真写过的键，不含 styleId
 *   与 rawPpr），`edit.marks` 是行内标记；`{ ok: false, code, detail }`：`no-block` = 选区里
 *   没有段落或标题，`mixed-paragraph-format` = 多个块的同一个段落属性取值不同（detail 是哪个
 *   属性），`mixed-run-format` = 标记只覆盖部分文字或两处属性不同（detail 是哪个标记）。
 */
export function captureStyleFromSelection(editor: Editor): CaptureResult {
  const blocks: PMNode[] = [];
  const { from, to, empty } = editor.state.selection;
  if (empty) {
    const parent = editor.state.selection.$from.parent;
    if (isFormatBlock(parent)) blocks.push(parent);
  } else {
    editor.state.doc.nodesBetween(from, to, (node) => {
      if (isFormatBlock(node)) blocks.push(node);
    });
  }
  if (!blocks.length) return { ok: false, code: "no-block" };

  const merged: Record<string, unknown> = {};
  for (const node of blocks) {
    for (const [key, value] of Object.entries(paragraphFormatOf(node))) {
      if (merged[key] !== undefined && !sameFormatValue(merged[key], value)) {
        return { ok: false, code: "mixed-paragraph-format", detail: key };
      }
      merged[key] = value;
    }
  }

  const marks = collectMarks(editor);
  if (marks.conflict) return { ok: false, code: "mixed-run-format", detail: marks.conflict };
  return { ok: true, edit: { paragraph: merged as PPrAttrs, marks: marks.marks ?? [] } };
}
