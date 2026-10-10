import { lineHeightCss, ptToPx } from "./extensions/ParagraphFormatting.js";
import { runStyleDeclarations } from "./extensions/TextStyle.js";
import type { ResolvedStyle, ResolvedStyleFormat } from "@office-copilot/docx-io";

/**
 * 文档作用域的样式表（`docs/word-parity-spec.md` 排名 2 ⑤ 的渲染半）。
 *
 * `docx-io` 已经把样式链按 Word 的优先级合并成编辑器那套词汇（pt / 倍数 / CSS px /
 * #rrggbb），这里只把它们写成 CSS 规则：段间距与缩进的 pt→px、行距的 `pt:`/`atLeast:`
 * 映射都复用段落直接格式那一份函数，因此不产生第二套换算（AGENTS §2.1）。
 *
 * 选择器是 `.ProseMirror [data-style-id="…"]`（特异度 0,2,0），压得过应用基线的
 * `.ProseMirror p` / `.ProseMirror h1`（0,1,1）；而用户做过的直接格式是行内 style，
 * 行内胜过样式表，正好停在 Word 的"直接格式最高"那一层。
 */

/** Word / WPS 实际使用的 styleId 字符集；不合规的 id 不进选择器。 */
const SAFE_STYLE_ID = /^[A-Za-z0-9_-]{1,64}$/;

/**
 * 文档默认段落样式该落在哪些元素上。
 *
 * @param declaredIds 这一份样式表里已经单独成规的 id。
 * @returns 选择器组：没写 `data-style-id` 的正文、用户清除过样式的段落（空串），
 *   以及**引用了本文档没声明的 id** 的段落——Word 对未定义引用就是按默认样式排的
 *   （`StyleMap.byId` 的告诫），真实文档里很常见：导出的简历写 `w:pStyle w:val="Heading1"`，
 *   styles.xml 里声明的却是 `1`。
 *   特异度固定 0,2,0（`:is()` 取的是参数里最高的那条，与条数无关），排在具名规则之前，
 *   于是同特异度时后者获胜，声明过的样式不会被这份兜底规则盖住。
 */
function fallbackSelector(declaredIds: string[]): string {
  const unstyled = '.ProseMirror p:not([data-style-id]), .ProseMirror [data-style-id=""]';
  if (!declaredIds.length) return `${unstyled}, .ProseMirror [data-style-id]`;
  const declared = declaredIds.map((id) => `[data-style-id="${id}"]`).join(", ");
  return `${unstyled}, .ProseMirror [data-style-id]:not(:is(${declared}))`;
}

/**
 * 校验一个来自 .docx 的字符串值能否原样进 CSS。
 *
 * @param value 待写进声明的值；undefined / null 表示这一层没写过。
 * @param pattern 这一类值的形状白名单。
 * @returns 通过则原值，否则 undefined（丢弃）。样式表的正文会写进 `<style>` 元素，
 *   含 `<`、`}` 之类字符的值能把规则截断，所以这份数据按系统边界对待。
 */
function safeValue(value: string | null | undefined, pattern: RegExp): string | undefined {
  return typeof value === "string" && pattern.test(value) ? value : undefined;
}

const CSS_LENGTH = /^-?\d+(?:\.\d+)?px$/;
const FONT_NAME = /^[^<>"'\\{};()]{1,64}$/;
const HEX_COLOR = /^#[0-9a-fA-F]{6}$/;
const LINE_HEIGHT_VALUE = /^(?:atLeast:|pt:)?\d+(?:\.\d+)?$/;

/**
 * 只保留形状可预期的字符格式，其余按"这份文档没写过"处理。
 *
 * @param format 一条样式的解析结果。
 * @returns 可以安全拼进 CSS 的同名字段；字体栈与颜色因此和 `TextStyle` 标记共用一份映射。
 */
function checkedRunFormat(format: ResolvedStyleFormat): ResolvedStyleFormat {
  return {
    ...format,
    fontFamily: safeValue(format.fontFamily, FONT_NAME),
    fontFamilyWest: safeValue(format.fontFamilyWest, FONT_NAME),
    fontSize: safeValue(format.fontSize, CSS_LENGTH),
    letterSpacing: safeValue(format.letterSpacing, CSS_LENGTH),
    color: safeValue(format.color, HEX_COLOR),
    backgroundColor: safeValue(format.backgroundColor, HEX_COLOR),
  };
}

/**
 * 段落属性 → CSS 声明。
 *
 * @param format 解析结果（spacingBefore/after、indent 等一律是 pt，lineHeight 是倍数
 *   或 `pt:N`/`atLeast:N`）。
 * @returns 声明列表，不含分号；这一层没写过的属性不出现。
 */
function paragraphDeclarations(format: ResolvedStyleFormat): string[] {
  const declarations: string[] = [];
  if (format.textAlign) declarations.push(`text-align: ${format.textAlign}`);
  if (format.spacingBefore != null)
    declarations.push(`margin-top: ${ptToPx(format.spacingBefore)}px`);
  if (format.spacingAfter != null)
    declarations.push(`margin-bottom: ${ptToPx(format.spacingAfter)}px`);
  if (format.indent != null)
    declarations.push(`margin-left: ${ptToPx(format.indent)}px`);
  if (format.indentRight != null)
    declarations.push(`margin-right: ${ptToPx(format.indentRight)}px`);
  if (format.indentFirst != null)
    declarations.push(`text-indent: ${ptToPx(format.indentFirst)}px`);
  const lineHeight = safeValue(format.lineHeight, LINE_HEIGHT_VALUE);
  if (lineHeight) declarations.push(`line-height: ${lineHeightCss(lineHeight)}`);
  return declarations;
}

/**
 * 字符格式里的加粗 / 倾斜 / 下划线 / 删除线 → CSS。
 *
 * @param format 解析结果；这些字段是三态，undefined = 样式链没提过。
 * @returns 声明列表。false 只有加粗与倾斜落成 `normal`：它们会盖到 `<h1>` 这类基线本身就是
 *   粗体的元素上，不写就等于没取消。下划线/删除线的 false 什么都不写就够了——装饰只可能来自
 *   这条规则本身或子元素的 `<u>`/`<s>`，两种情况都不需要在这里取消。
 */
function emphasisDeclarations(format: ResolvedStyleFormat): string[] {
  const declarations: string[] = [];
  if (format.bold !== undefined)
    declarations.push(`font-weight: ${format.bold ? "bold" : "normal"}`);
  if (format.italic !== undefined)
    declarations.push(`font-style: ${format.italic ? "italic" : "normal"}`);
  const decorationLines = [
    format.underline === true ? "underline" : null,
    format.strike === true ? "line-through" : null,
  ].filter((line): line is string => line !== null);
  if (decorationLines.length)
    declarations.push(`text-decoration: ${decorationLines.join(" ")}`);
  return declarations;
}

/**
 * 一条具名样式对应的选择器。
 *
 * @param styleId styles.xml 声明的 id（null 走 `fallbackSelector`，不从这里走）。
 * @returns 选择器；id 不合 `SAFE_STYLE_ID` 时 undefined，这条样式因此整条不产出，
 *   引用它的段落由兜底规则接管。
 */
function selectorFor(styleId: string): string | undefined {
  return SAFE_STYLE_ID.test(styleId)
    ? `.ProseMirror [data-style-id="${styleId}"]`
    : undefined;
}

/**
 * 一条解析结果对应的 CSS 声明。
 *
 * @param format 段落 + 字符的合并结果（单位见模块头）。
 * @returns 声明列表；值的形状不认识时逐条丢弃，其余照常生效。
 */
function declarationsFor(format: ResolvedStyleFormat): string[] {
  const checked = checkedRunFormat(format);
  return [
    ...paragraphDeclarations(checked),
    ...runStyleDeclarations(checked),
    ...emphasisDeclarations(checked),
  ];
}

/**
 * 拼一条规则。
 *
 * @param selector 选择器（可以是一组，逗号分隔）。
 * @param declarations 声明列表，不含分号。
 * @returns 多行 CSS 规则；空声明返回空串，由调用方过滤。
 */
function rule(selector: string, declarations: string[]): string {
  if (!declarations.length) return "";
  return `${selector} {\n  ${declarations.join(";\n  ")};\n}`;
}

/**
 * 把解析好的样式表拼成一段 CSS。
 *
 * @param styles `docx-io` 的 `PmLoadResult.styleFormats`；空数组（新建文档、或样式链
 *   什么都没写过）得到空串。
 * @returns 兜底规则在最前，具名规则按入参顺序跟在后面（同特异度时后者获胜，见
 *   `fallbackSelector`）。渲染层只管把这段文本放进 `<style>`，不含单位换算与 Word 语义判断。
 */
export function docStyleSheet(styles: ResolvedStyle[]): string {
  const named: { styleId: string; selector: string; declarations: string[] }[] = [];
  for (const entry of styles) {
    if (entry.styleId === null) continue;
    const selector = selectorFor(entry.styleId);
    if (!selector) continue;
    named.push({ styleId: entry.styleId, selector, declarations: declarationsFor(entry.format) });
  }
  const rules: string[] = [];
  const defaultEntry = styles.find((entry) => entry.styleId === null);
  if (defaultEntry) {
    rules.push(
      rule(
        fallbackSelector(named.map((style) => style.styleId)),
        declarationsFor(defaultEntry.format),
      ),
    );
  }
  for (const style of named) rules.push(rule(style.selector, style.declarations));
  return rules.filter(Boolean).join("\n\n");
}
