/**
 * 右键菜单的数据层：把"右击落在什么对象上"翻译成一张 Word 风格的命令表。
 *
 * Word 的右键菜单会随目标变化（文字 / 图片 / 表格各给不同命令），并且把当前
 * 不可用的命令置灰而不是隐藏。这里只产出结构化数据与禁用条件，不含任何 DOM 与
 * React，方便单测覆盖；渲染层（components/EditorContextMenu.tsx）负责装配。
 */

import type { Editor } from "@tiptap/core";

/** 右击落点。table 交给 TableContextMenu，本模块不产出菜单（返回 null）。 */
export type CtxTarget = "text" | "image" | "table";

/** 剪贴板内容快照；两者都空时「粘贴」置灰，与 Word 一致。 */
export interface ClipboardSnapshot {
  text: string;
  html: string;
}

export interface CtxFacts {
  target: CtxTarget;
  /** 有实质选区（拖选的文字或被选中的图片对象），决定剪切/复制是否可用。 */
  hasSelection: boolean;
  isLink: boolean;
  linkHref: string | null;
  isListItem: boolean;
  isBulletList: boolean;
  isOrderedList: boolean;
  /** 光标所在块的标题级别，null 表示正文。 */
  headingLevel: number | null;
  /** 被右击图片当前的文字环绕方式，非图片目标为 null。 */
  imageWrap: "none" | "square" | "tight" | "through" | "topAndBottom" | null;
  /** 格式刷是否已取好样式（正在待刷状态）。 */
  painterActive: boolean;
  clipboard: ClipboardSnapshot;
}

export type CtxItemId =
  | "cut"
  | "copy"
  | "paste"
  | "paste-source"
  | "paste-merge"
  | "paste-text"
  | "format-painter"
  | "clear-format"
  | "bullet-list"
  | "numbered-list"
  | "indent-more"
  | "indent-less"
  | "block"
  | "wrap"
  | "font-dialog"
  | "paragraph-dialog"
  | "link-insert"
  | "link-open"
  | "link-remove"
  | "delete-object"
  | `block:${string}`
  | `wrap:${string}`
  | `ai:${string}`;

export interface CtxItem {
  id: CtxItemId;
  /** 菜单文字，与 Word 的中文界面用词一致。 */
  label: string;
  /** 右侧灰色快捷键提示；只写真实存在的绑定，不虚构。 */
  hint?: string;
  disabled?: boolean;
  checked?: boolean;
  danger?: boolean;
  children?: CtxItem[];
}

export type CtxRow = CtxItem | { separator: true };

const WORD_WRAP_LABELS: Record<NonNullable<CtxFacts["imageWrap"]>, string> = {
  none: "紧密型环绕（嵌入文字下方）",
  square: "方形",
  tight: "紧密",
  through: "穿越",
  topAndBottom: "上下型环绕",
};

const PASTE_ITEMS: CtxItem[] = [
  { id: "paste-source", label: "保留源格式" },
  { id: "paste-merge", label: "合并格式" },
  { id: "paste-text", label: "只保留文本" },
];

/**
 * 由落点事实生成菜单行。
 *
 * @param facts 渲染层从编辑器状态 + DOM + 剪贴板采集到的事实（字号单位不参与）。
 * @returns 菜单行数组；目标在表格内时返回 null，表示"这张菜单不归我们"。
 */
export function buildContextMenu(facts: CtxFacts): CtxRow[] | null {
  if (facts.target === "table") return null;
  const rows: CtxRow[] = [];
  const canEdit = facts.hasSelection;

  if (facts.target === "image") {
    pushClipboardRows(rows, facts, /* withPaste */ true);
    rows.push({ separator: true });
    rows.push({
      id: "wrap",
      label: "文字环绕",
      children: (Object.keys(WORD_WRAP_LABELS) as Array<keyof typeof WORD_WRAP_LABELS>).map(
        (kind) => ({
          id: `wrap:${kind}` as CtxItemId,
          label: WORD_WRAP_LABELS[kind],
          checked: facts.imageWrap === kind,
        }),
      ),
    });
    rows.push({ separator: true });
    rows.push({ id: "delete-object", label: "删除图片", danger: true });
    return rows;
  }

  pushClipboardRows(rows, facts, /* withPaste */ true);
  rows.push({ separator: true });
  rows.push({
    id: "format-painter",
    label: "格式刷",
    checked: facts.painterActive,
  });
  rows.push({ id: "clear-format", label: "清除所有格式", disabled: !canEdit });
  rows.push({ separator: true });
  rows.push({
    id: "block",
    label: "段落样式",
    children: [
      {
        id: "block:paragraph",
        label: "正文",
        checked: facts.headingLevel === null,
      },
      ...[1, 2, 3, 4, 5, 6].map((level) => ({
        id: `block:h${level}` as CtxItemId,
        label: `标题 ${level}`,
        checked: facts.headingLevel === level,
      })),
    ],
  });
  rows.push({
    id: "bullet-list",
    label: "项目符号",
    checked: facts.isBulletList,
  });
  rows.push({
    id: "numbered-list",
    label: "编号",
    checked: facts.isOrderedList,
  });
  rows.push({
    id: "indent-more",
    label: "增加缩进",
    hint: "Tab",
    disabled: !facts.isListItem,
  });
  rows.push({
    id: "indent-less",
    label: "减少缩进",
    hint: "Shift+Tab",
    disabled: !facts.isListItem,
  });
  rows.push({ separator: true });
  rows.push({ id: "font-dialog", label: "字体(F)…" });
  rows.push({ id: "paragraph-dialog", label: "段落(P)…" });

  if (facts.isLink) {
    rows.push({ separator: true });
    rows.push({ id: "link-open", label: "打开超链接" });
    rows.push({ id: "link-insert", label: "编辑超链接(E)…" });
    rows.push({ id: "link-remove", label: "删除超链接" });
  } else {
    rows.push({ id: "link-insert", label: "超链接(K)…" });
  }

  rows.push({ separator: true });
  rows.push({ id: "ai:polish", label: "AI：润色", disabled: !canEdit });
  rows.push({ id: "ai:expand", label: "AI：扩写", disabled: !canEdit });
  rows.push({ id: "ai:shorten", label: "AI：缩写", disabled: !canEdit });
  return rows;
}

/**
 * 追加 Word 顶部的剪贴板分组：剪切/复制按选区置灰，粘贴按剪贴板置灰。
 *
 * @param withPaste 为 false 时只给剪切/复制（表格外的空档等场景目前都用 true）。
 */
function pushClipboardRows(
  rows: CtxRow[],
  facts: CtxFacts,
  withPaste: boolean,
): void {
  rows.push({ id: "cut", label: "剪切", hint: "⌘X", disabled: !facts.hasSelection });
  rows.push({ id: "copy", label: "复制", hint: "⌘C", disabled: !facts.hasSelection });
  if (!withPaste) return;
  const hasClipboard = Boolean(facts.clipboard.text || facts.clipboard.html);
  rows.push({
    id: "paste",
    label: "粘贴",
    hint: "⌘V",
    disabled: !hasClipboard,
    children: PASTE_ITEMS.map((item) =>
      item.id === "paste-text" && !facts.clipboard.text
        ? { ...item, disabled: true }
        : item,
    ),
  });
}

/**
 * 把菜单摆到鼠标点上，并保证整块菜单留在窗口内（Word 会在贴边时上翻/左移）。
 *
 * @param point 右键命中点，视口坐标 px。
 * @param size 菜单自身宽高，px。
 * @param viewport 视口宽高，px。
 * @returns 菜单的 fixed 定位左上角，px。
 */
export function menuRect(
  point: { x: number; y: number },
  size: { width: number; height: number },
  viewport: { width: number; height: number },
): { left: number; top: number } {
  const gap = 4;
  const flipUp = point.y + size.height + gap > viewport.height;
  const flipLeft = point.x + size.width + gap > viewport.width;
  return {
    left: Math.max(gap, flipLeft ? point.x - size.width : point.x),
    top: Math.max(gap, flipUp ? viewport.height - size.height - gap : point.y),
  };
}

/**
 * 子菜单贴在父项右侧的位置，父项贴近视口右缘时翻到左侧。
 *
 * @param parent 父菜单矩形（左上角 + 宽高，px）。
 * @param itemTop 被悬停项相对父菜单顶部的偏移，px。
 * @param size 子菜单自身宽高，px。
 * @param viewport 视口宽高，px。
 */
export function submenuRect(
  parent: { left: number; top: number; width: number },
  itemTop: number,
  size: { width: number; height: number },
  viewport: { width: number; height: number },
): { left: number; top: number } {
  const flipLeft = parent.left + parent.width + size.width > viewport.width;
  const top = Math.min(
    Math.max(4, parent.top + itemTop),
    Math.max(4, viewport.height - size.height - 4),
  );
  return {
    left: flipLeft
      ? Math.max(4, parent.left - size.width)
      : Math.min(parent.left + parent.width, viewport.width - size.width - 4),
    top,
  };
}

/**
 * 格式刷当前是否已经存好一份样式（待刷状态）。
 *
 * @param editor 桌面编辑器实例；格式刷存在 extension storage 里，TipTap 没有公开类型。
 */
export function hasStoredFormat(editor: Editor): boolean {
  return Boolean(
    (editor as unknown as { storage: { formatPainter?: { stored: unknown } } }).storage
      .formatPainter?.stored,
  );
}

