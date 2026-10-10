import { canonColor } from "@office-copilot/docx-io";
import { Extension, Mark, mergeAttributes } from "@tiptap/core";

interface TextStyleAttrs {
  fontFamily?: string | null;
  fontFamilyWest?: string | null;
  fontSize?: string | null;
  letterSpacing?: string | null;
  color?: string | null;
  backgroundColor?: string | null;
}

/**
 * 字符格式 → CSS 声明。
 *
 * @param attrs 编辑器存的字符格式：字号与字距是 CSS px 字符串，颜色是 #rrggbb，
 *   null / undefined 都表示"没设过"。
 * @returns 按 OOXML 语义排好的声明（不含分号）；标记渲染与文档作用域样式表共用这一份映射，
 *   于是"西文字体在前"这类约定不会在两处各写一边（AGENTS §3.2）。
 */
export function runStyleDeclarations(attrs: TextStyleAttrs): string[] {
  // 西文字体排在前：CSS 的字体栈是按字符回退的，拉丁字母取前者、汉字取后者，
  // 正是 OOXML 的 ascii/eastAsia 分工。
  const stack = [attrs.fontFamilyWest, attrs.fontFamily].filter(Boolean).join(", ");
  return [
    stack ? `font-family: ${stack}` : null,
    attrs.fontSize ? `font-size: ${attrs.fontSize}` : null,
    attrs.letterSpacing ? `letter-spacing: ${attrs.letterSpacing}` : null,
    attrs.color ? `color: ${attrs.color}` : null,
    attrs.backgroundColor ? `background-color: ${attrs.backgroundColor}` : null,
  ].filter((declaration): declaration is string => declaration !== null);
}

const TEXT_STYLE_DATA_ATTRS = [
  "data-font-family",
  "data-font-family-west",
  "data-font-size",
  "data-letter-spacing",
  "data-color",
  "data-background-color",
] as const;

declare module "@tiptap/core" {
  interface Commands<ReturnType> {
    textStyleFormatting: {
      setTextStyle: (attrs: TextStyleAttrs) => ReturnType;
      unsetTextStyle: () => ReturnType;
      toggleTextStyle: (attrs: Partial<TextStyleAttrs>) => ReturnType;
    };
  }
}

/**
 * A single "textStyle" mark carrying font family, font size (CSS px string),
 * text color and background color. This maps directly onto our OOXML
 * converter.
 *
 * Each attribute round-trips through its own `data-*` and reads that first,
 * exactly like the paragraph formatting attrs. Without it the browser rewrites
 * `color: #c00000` into `rgb(192, 0, 0)` on the way back in, so "is this run
 * already this colour?" answered wrongly and re-applying a template kept
 * producing edits that changed nothing.
 */
export const TextStyle = Mark.create({
  name: "textStyle",

  inclusive: true,

  addOptions() {
    return { HTMLAttributes: {} };
  },

  addAttributes() {
    const data = (el: HTMLElement, name: string) => el.getAttribute(name) || null;
    return {
      fontFamily: {
        default: null,
        parseHTML: (el) => data(el as HTMLElement, "data-font-family") || (el as HTMLElement).style.fontFamily || null,
        renderHTML: (attrs) =>
          attrs.fontFamily ? { "data-font-family": String(attrs.fontFamily) } : {},
      },
      fontFamilyWest: {
        default: null,
        parseHTML: (el) => data(el as HTMLElement, "data-font-family-west"),
        renderHTML: (attrs) =>
          attrs.fontFamilyWest ? { "data-font-family-west": String(attrs.fontFamilyWest) } : {},
      },
      fontSize: {
        default: null,
        parseHTML: (el) => data(el as HTMLElement, "data-font-size") || (el as HTMLElement).style.fontSize || null,
        renderHTML: (attrs) =>
          attrs.fontSize ? { "data-font-size": String(attrs.fontSize) } : {},
      },
      letterSpacing: {
        default: null,
        parseHTML: (el) => data(el as HTMLElement, "data-letter-spacing"),
        renderHTML: (attrs) =>
          attrs.letterSpacing ? { "data-letter-spacing": String(attrs.letterSpacing) } : {},
      },
      color: {
        default: null,
        parseHTML: (el) =>
          canonColor(data(el as HTMLElement, "data-color") || (el as HTMLElement).style.color),
        renderHTML: (attrs) => (attrs.color ? { "data-color": String(attrs.color) } : {}),
      },
      backgroundColor: {
        default: null,
        parseHTML: (el) =>
          canonColor(
            data(el as HTMLElement, "data-background-color") ||
              (el as HTMLElement).style.backgroundColor,
          ),
        renderHTML: (attrs) =>
          attrs.backgroundColor
            ? { "data-background-color": String(attrs.backgroundColor) }
            : {},
      },
    };
  },

  parseHTML() {
    return [
      { tag: "span[data-text-style]" },
      {
        tag: "span",
        getAttrs: (el) => {
          const node = el as HTMLElement;
          if (TEXT_STYLE_DATA_ATTRS.some((name) => node.getAttribute(name))) return null;
          const s = node.style;
          if (!s) return false;
          return s.fontFamily || s.fontSize || s.color || s.backgroundColor ? null : false;
        },
      },
    ];
  },

  renderHTML({ mark, HTMLAttributes }) {
    // `HTMLAttributes` now holds the rendered `data-*` map; the CSS goes to the
    // stored attrs, which printing reads through getHTML().
    const attrs = ((mark as { attrs?: TextStyleAttrs } | undefined)?.attrs ??
      (HTMLAttributes as TextStyleAttrs)) as TextStyleAttrs;
    const style = runStyleDeclarations(attrs).join("; ");
    return [
      "span",
      mergeAttributes(
        this.options.HTMLAttributes as Record<string, unknown>,
        HTMLAttributes as Record<string, string>,
        {
          "data-text-style": "",
          ...(style ? { style } : {}),
        },
      ),
      0,
    ];
  },

  addCommands() {
    return {
      setTextStyle:
        (attrs: TextStyleAttrs) =>
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        ({ chain }: any) =>
          chain().setMark(this.name, attrs).run(),
      unsetTextStyle:
        () =>
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        ({ chain }: any) =>
          chain().unsetMark(this.name).run(),
    } as never;
  },
});

/**
 * Helper that sets only the provided keys while preserving existing ones.
 */
export const TextStyleCommands = Extension.create({
  name: "textStyleCommands",
  addCommands() {
    return {
      toggleTextStyle:
        (attrs: Partial<TextStyleAttrs>) =>
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        ({ editor, chain }: any) => {
          const current = (editor.getAttributes("textStyle") ?? {}) as TextStyleAttrs;
          const next: TextStyleAttrs = {
            fontFamily: current.fontFamily ?? null,
            fontFamilyWest: current.fontFamilyWest ?? null,
            fontSize: current.fontSize ?? null,
            letterSpacing: current.letterSpacing ?? null,
            color: current.color ?? null,
            backgroundColor: current.backgroundColor ?? null,
            ...attrs,
          };
          return chain().setMark("textStyle", next).run();
        },
    } as never;
  },
});
