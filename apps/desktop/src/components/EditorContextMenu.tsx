import { useCallback, useEffect, useLayoutEffect, useRef, useState } from "react";
import type { Editor } from "@tiptap/react";
import { NodeSelection } from "@tiptap/pm/state";
import {
  buildContextMenu,
  hasStoredFormat,
  menuRect,
  submenuRect,
  type CtxFacts,
  type CtxItem,
  type CtxItemId,
  type CtxRow,
} from "../editor/contextMenu.js";
import { selectionTouchesRevision } from "../editor/extensions/Revision.js";
import { ParagraphFields } from "./ParagraphPopover.js";
import { FontFields } from "./FontFields.js";
import { promptSetLink, safeHref } from "./linkPrompt.js";

/**
 * 正文的右键编辑菜单：Word 风格的命令表 + 置灰规则 + 字体/段落对话框。
 *
 * 表格内的右击仍由 TableContextMenu 负责（它只认 td/th），两者监听同一个
 * DOM、互不重叠。菜单纯数据来自 editor/contextMenu.ts，这里只做装配与落地。
 */

/** .ctx-menu 的实测宽度，见 styles.css 的 .ctx-menu/.ctx-item。 */
const MENU_W = 208;
const SUB_W = 186;
const PANEL_W = 280;
/** 一行菜单 / 一条分隔线的高度估计，px；真实高度以量出来的为准。 */
const ROW_H = 31;
const SEP_H = 9;

/**
 * 浮层的真实高度：第一帧用估计值定位，量到以后立刻校正。
 *
 * 菜单位置要按真实高度决定"贴近视口下缘时上翻"，纯估计会溢出窗口。
 *
 * @param fallback 首帧使用的高度估计，px。
 * @returns 挂到浮层上的 ref 与当前高度，px。
 */
function useMeasuredHeight(fallback: number, resetKey: unknown) {
  const ref = useRef<HTMLDivElement>(null);
  const [height, setHeight] = useState(fallback);
  const lastKey = useRef(resetKey);
  if (lastKey.current !== resetKey) {
    lastKey.current = resetKey;
    setHeight(fallback);
  }
  const measure = useCallback(() => {
    const el = ref.current;
    if (el && Math.abs(el.offsetHeight - height) > 2) setHeight(el.offsetHeight);
  }, [height]);
  useLayoutEffect(measure);
  return { ref, height };
}

interface OpenMenu {
  x: number;
  y: number;
  facts: CtxFacts;
  /** 右击落在图片节点上时的位置，其它目标为 null。 */
  imagePos: number | null;
}

interface Panel {
  kind: "font" | "paragraph";
  x: number;
  y: number;
}

export function EditorContextMenu({ editor }: { editor: Editor }) {
  const [open, setOpen] = useState<OpenMenu | null>(null);
  const [panel, setPanel] = useState<Panel | null>(null);
  const [sub, setSub] = useState<{ id: CtxItemId; top: number } | null>(null);
  // 子菜单的"延时收"：斜着走向子菜单的路上必然擦过下面几行叶子项，立刻收就永远
  // 走不到目的地；给 260ms 宽限，进入菜单/子菜单就撤销，停手不走了才收。
  const closeSubTimer = useRef<number | null>(null);
  const cancelCloseSub = useCallback(() => {
    if (closeSubTimer.current !== null) {
      window.clearTimeout(closeSubTimer.current);
      closeSubTimer.current = null;
    }
  }, []);
  const scheduleCloseSub = useCallback(() => {
    cancelCloseSub();
    closeSubTimer.current = window.setTimeout(() => setSub(null), 260);
  }, [cancelCloseSub]);
  useEffect(() => cancelCloseSub, [cancelCloseSub]);

  useEffect(() => {
    const dom = editor.view.dom;
    const onContextMenu = (e: MouseEvent) => {
      const hit = e.target as HTMLElement;
      if (hit.closest("td, th")) return;
      if (!editor.isEditable) return;
      e.preventDefault();
      void openAt(editor, e.clientX, e.clientY, hit).then((state) => {
        if (state) {
          setPanel(null);
          setSub(null);
          setOpen(state);
        }
      });
    };
    dom.addEventListener("contextmenu", onContextMenu);
    return () => dom.removeEventListener("contextmenu", onContextMenu);
  }, [editor]);

  // 点别处、按 Esc、窗口尺寸变化或页面滚动都要收掉菜单（滚动后坐标就失效了）。
  useEffect(() => {
    if (!open && !panel) return;
    const close = () => {
      cancelCloseSub();
      setOpen(null);
      setPanel(null);
      setSub(null);
    };
    const onDown = (e: MouseEvent) => {
      if (e.button !== 0) return;
      const t = e.target as Node;
      if (menuRef.current?.contains(t) || panelRef.current?.contains(t)) return;
      close();
    };
    const onKey = (e: KeyboardEvent) => e.key === "Escape" && close();
    window.addEventListener("mousedown", onDown, true);
    window.addEventListener("keydown", onKey);
    window.addEventListener("resize", close);
    window.addEventListener("scroll", close, true);
    return () => {
      window.removeEventListener("mousedown", onDown, true);
      window.removeEventListener("keydown", onKey);
      window.removeEventListener("resize", close);
      window.removeEventListener("scroll", close, true);
    };
  }, [open, panel]);

  const rows = open ? buildContextMenu(open.facts) : null;
  const view = { width: window.innerWidth, height: window.innerHeight };
  const menuBox = useMeasuredHeight(
    open ? estimateHeight(rows ?? []) : 0,
    open ? `${open.x}:${open.y}` : null,
  );
  const panelBox = useMeasuredHeight(380, panel ? `${panel.kind}:${panel.x}` : null);
  const subBox = useMeasuredHeight(120, sub ? `${sub.id}:${sub.top}` : null);
  const menuRef = menuBox.ref;
  const panelRef = panelBox.ref;
  const menuPos = open
    ? menuRect({ x: open.x, y: open.y }, { width: MENU_W, height: menuBox.height }, view)
    : null;
  const panelPos = panel
    ? menuRect({ x: panel.x, y: panel.y }, { width: PANEL_W, height: panelBox.height }, view)
    : null;

  const act = (id: CtxItemId) => {
    const anchor = open;
    if (!anchor) return;
    cancelCloseSub();
    if (id === "font-dialog" || id === "paragraph-dialog") {
      setOpen(null);
      setSub(null);
      setPanel({ kind: id === "font-dialog" ? "font" : "paragraph", x: anchor.x, y: anchor.y });
      return;
    }
    setOpen(null);
    setPanel(null);
    setSub(null);
    void applyAction(editor, id, anchor);
  };

  return (
    <>
      {open && rows && menuPos && (
        <div
          ref={menuRef}
          className="ctx-menu editor-ctx-menu"
          style={{ left: menuPos.left, top: menuPos.top, width: MENU_W }}
          onMouseDown={(e) => e.preventDefault()}
          // 子菜单只在指针彻底离开"菜单 + 子菜单"这个整体时才收掉：mouseleave 按
          // DOM 后代判定，子菜单是本容器的后代，所以路过父项右侧那 5px 内边距、或
          // 斜着走向子菜单的路上都不会误关；真正的离开还要过 260ms 延时才生效。
          onMouseEnter={cancelCloseSub}
          onMouseLeave={scheduleCloseSub}
          // 手慢的时候，斜着走会擦过下面几行叶子项，光是"进入叶子"就排程收掉的话，
          // 用户永远走不到子菜单上。所以：只要在菜单内继续移动就撤销这个排程，真正
          // 停下或离开整体才收。
          onMouseMove={cancelCloseSub}
        >
          {rows.map((row, i) =>
            "separator" in row ? (
              <div key={`sep-${i}`} className="ctx-sep" />
            ) : row.children ? (
              // 子菜单挂在父项的 DOM 之内，让它成为父项的"后代命中区"：斜着移动进
              // 子菜单时不会走到"谁都不属于"的空白上；收掉的动作交给容器（见上）。
              <div key={row.id + String(i)} className="ctx-row">
                <CtxButton
                  item={row}
                  onHover={(top) => {
                    cancelCloseSub();
                    setSub({ id: row.id, top });
                  }}
                  onLeave={() => {}}
                  onRun={() => {}}
                />
                {sub?.id === row.id &&
                  (() => {
                    const pos = submenuRect(
                      { left: menuPos.left, top: menuPos.top, width: MENU_W },
                      sub.top,
                      { width: SUB_W, height: subBox.height },
                      view,
                    );
                    return (
                      <div
                        ref={subBox.ref}
                        className="ctx-menu ctx-submenu"
                        style={{ left: pos.left, top: pos.top, width: SUB_W }}
                        // 从叶子项绕进子菜单时指针一直没离开外层容器，容器的
                        // mouseenter 不会再触发，所以这里自己撤销待执行的收。
                        onMouseEnter={cancelCloseSub}
                        onMouseDown={(e) => e.preventDefault()}
                      >
                        {row.children.map((child) => (
                          <CtxButton
                            key={child.id}
                            item={child}
                            onHover={() => {}}
                            onLeave={() => {}}
                            onRun={() => act(child.id)}
                          />
                        ))}
                      </div>
                    );
                  })()}
              </div>
            ) : (
              <CtxButton
                key={row.id + String(i)}
                item={row}
                // 移到没有子菜单的项上：延时收掉还开着的子菜单（立刻收的话，斜着
                // 走向子菜单的路上擦过这一行就把目的地关掉了）。
                onHover={scheduleCloseSub}
                onLeave={() => {}}
                onRun={() => act(row.id)}
              />
            ),
          )}
        </div>
      )}

      {panel && panelPos && (
        <div
          ref={panelRef}
          className="paragraph-popover ctx-panel"
          style={{ left: panelPos.left, top: panelPos.top }}
          onMouseDown={(e) => {
            // 输入框/下拉必须能拿到焦点，否则打字全落回编辑器；其余空白处阻止默认，
            // 保住文档里的选区（改动仍作用在用户右键前选中的文字上）。
            if ((e.target as HTMLElement).closest("input, select, textarea")) return;
            e.preventDefault();
          }}
        >
          {panel.kind === "font" ? (
            <FontFields editor={editor} onClose={() => setPanel(null)} />
          ) : (
            <ParagraphFields editor={editor} onClose={() => setPanel(null)} />
          )}
        </div>
      )}
    </>
  );
}

function CtxButton({
  item,
  onHover,
  onLeave,
  onRun,
}: {
  item: CtxItem;
  onHover: (offsetTop: number) => void;
  onLeave: () => void;
  onRun: () => void;
}) {
  return (
    <button
      type="button"
      className={`ctx-item ${item.danger ? "danger" : ""} ${item.checked ? "checked" : ""}`}
      disabled={item.disabled}
      onMouseEnter={(e) => onHover((e.currentTarget as HTMLElement).offsetTop)}
      onMouseLeave={onLeave}
      onClick={() => !item.disabled && !item.children && onRun()}
    >
      <span className="ctx-check">{item.checked ? "✓" : ""}</span>
      <span className="ctx-label">{item.label}</span>
      {item.hint ? <span className="ctx-hint">{item.hint}</span> : null}
      {item.children ? <span className="ctx-arrow">›</span> : null}
    </button>
  );
}

/**
 * 执行一条菜单命令。菜单关闭后编辑器的选区仍在（面板上的 mousedown 都被 preventDefault）。
 *
 * @param id 命令标识，来自 editor/contextMenu.ts。
 * @param anchor 打开菜单那一刻的落点事实与图片位置。
 */
async function applyAction(
  editor: Editor,
  id: CtxItemId,
  anchor: OpenMenu,
): Promise<void> {
  switch (id) {
    case "cut":
      document.execCommand("cut");
      return;
    case "copy":
      document.execCommand("copy");
      return;
    case "paste":
    case "paste-source":
      return pasteFromClipboard(editor, "source");
    case "paste-merge":
      return pasteFromClipboard(editor, "merge");
    case "paste-text":
      return pasteFromClipboard(editor, "text");
    case "format-painter":
      if (hasStoredFormat(editor)) editor.commands.cancelFormatPaint();
      else editor.commands.pickFormat();
      return;
    case "clear-format":
      editor
        .chain()
        .focus()
        .unsetMark("bold")
        .unsetMark("italic")
        .unsetMark("underline")
        .unsetMark("strike")
        .unsetTextStyle()
        .run();
      return;
    case "bullet-list":
      editor.chain().focus().toggleBulletList().run();
      return;
    case "numbered-list":
      editor.chain().focus().toggleOrderedList().run();
      return;
    case "indent-more":
      editor.chain().focus().sinkListItem("listItem").run();
      return;
    case "indent-less":
      editor.chain().focus().liftListItem("listItem").run();
      return;
    case "font-dialog":
    case "paragraph-dialog":
      return;
    case "link-insert":
      return promptSetLink(editor);
    case "link-open": {
      const href = anchor.facts.linkHref && safeHref(anchor.facts.linkHref);
      if (href) window.open(href, "_blank", "noopener");
      return;
    }
    case "link-remove":
      editor.chain().focus().extendMarkRange("link").unsetLink().run();
      return;
    case "revision-accept":
      editor.chain().focus().acceptRevision().run();
      return;
    case "revision-reject":
      editor.chain().focus().rejectRevision().run();
      return;
    case "delete-object":
      if (anchor.imagePos === null) return;
      editor.chain().focus().setNodeSelection(anchor.imagePos).deleteSelection().run();
      return;
    case "block:paragraph":
      editor.chain().focus().setParagraph().run();
      return;
    default:
      break;
  }

  if (id.startsWith("block:h")) {
    const level = Number(id.slice("block:h".length)) as 1 | 2 | 3 | 4 | 5 | 6;
    editor.chain().focus().toggleHeading({ level }).run();
    return;
  }
  if (id.startsWith("wrap:")) {
    const wrap = id.slice("wrap:".length) as NonNullable<CtxFacts["imageWrap"]>;
    if (anchor.imagePos === null) return;
    editor
      .chain()
      .focus()
      .setNodeSelection(anchor.imagePos)
      .updateAttributes("image", { wrap })
      .run();
    return;
  }
  if (id.startsWith("ai:")) {
    const action = { polish: "润色", expand: "扩写", shorten: "缩写" }[
      id.slice("ai:".length)
    ];
    if (action) window.dispatchEvent(new CustomEvent("ai:run", { detail: action }));
  }
}

/**
 * 从系统剪贴板粘贴，三种 Word 粘贴方式的语义：
 * 保留源格式＝原样进文档；合并格式＝保留结构但丢弃直接字符格式（字体/字号/颜色）；
 * 只保留文本＝纯文本，多行按段落切开。
 *
 * @param mode 粘贴方式。
 */
async function pasteFromClipboard(
  editor: Editor,
  mode: "source" | "merge" | "text",
): Promise<void> {
  const clip = await window.officeCopilot.readClipboard();
  const html = mode === "text" ? "" : clip.html;
  if (!html) {
    insertPlainText(editor, clip.text);
    return;
  }
  const from = editor.state.selection.from;
  editor.chain().focus().insertContent(html).run();
  if (mode !== "merge") return;
  const to = editor.state.selection.to;
  if (to > from) editor.chain().setTextSelection({ from, to }).unsetTextStyle().run();
}

/**
 * 纯文本粘贴：一行按插入文字处理，多行按段落插入（Word 的"只保留文本"）。
 *
 * @param text 剪贴板纯文本，可能带 \r\n。
 */
function insertPlainText(editor: Editor, text: string): void {
  const lines = text.replace(/\r\n?/g, "\n").split("\n");
  if (lines.length <= 1) {
    // 用节点 JSON 而不是 HTML 字符串：insertContent 收到字符串会当 HTML 解析，
    // 而"只保留文本"必须原样插入，包括形如 <b> 的字面量。
    if (!lines[0]) return;
    editor.chain().focus().insertContent({ type: "text", text: lines[0] }).run();
    return;
  }
  editor.chain().focus().insertContent(
    lines.map((line) =>
      line ? { type: "paragraph", content: [{ type: "text", text: line }] } : { type: "paragraph" },
    ),
  ).run();
}

/**
 * 采集右击落点的事实，并按 Word 的规矩处理选区：
 * 右击落在现有选区内保持选区，落在别处则把光标移到右击处。
 *
 * @param x 视口坐标 px。
 * @param y 视口坐标 px。
 * @param hit 右击命中的 DOM 元素。
 * @returns 打开菜单所需的状态；落点不在文档内时仍出菜单，只是命令按现有选区执行。
 */
async function openAt(
  editor: Editor,
  x: number,
  y: number,
  hit: HTMLElement,
): Promise<OpenMenu | null> {
  const where = editor.view.posAtCoords({ left: x, top: y });
  const imagePos = imagePosAt(editor, where?.inside ?? null, hit);
  const selection = editor.state.selection;
  const keepSelection =
    !selection.empty && where !== null && where.pos >= selection.from && where.pos <= selection.to;
  if (where && !keepSelection && imagePos === null) {
    editor.commands.setTextSelection(where.pos);
  }

  const state = editor.state;
  const nodeSelection = state.selection instanceof NodeSelection;
  const $pos = state.selection.$from;
  let headingLevel: number | null = null;
  let isListItem = false;
  let inBulletList = false;
  let inOrderedList = false;
  for (let depth = $pos.depth; depth > 0; depth -= 1) {
    const ancestor = $pos.node(depth);
    if (ancestor.type.name === "heading" && headingLevel === null)
      headingLevel = Number(ancestor.attrs.level);
    if (ancestor.type.name === "listItem") isListItem = true;
    if (ancestor.type.name === "bulletList") inBulletList = true;
    if (ancestor.type.name === "orderedList") inOrderedList = true;
  }

  const linkMark = state.selection.$from
    .marks()
    .find((mark) => mark.type.name === "link");
  const href =
    hit.closest<HTMLAnchorElement>("a[href]")?.getAttribute("href") ??
    (linkMark?.attrs.href as string | undefined) ??
    null;
  const imageNode = imagePos === null ? null : state.doc.nodeAt(imagePos);
  const clipboard = await window.officeCopilot.readClipboard();

  return {
    x,
    y,
    imagePos,
    facts: {
      target: imagePos === null ? "text" : "image",
      hasSelection: !state.selection.empty,
      isLink: Boolean(href),
      linkHref: href,
      isListItem,
      isBulletList: inBulletList,
      isOrderedList: inOrderedList,
      headingLevel,
      imageWrap: (imageNode?.attrs.wrap as CtxFacts["imageWrap"]) ?? null,
      painterActive: hasStoredFormat(editor),
      hasRevision: selectionTouchesRevision(editor),
      clipboard,
    },
  };
}

/**
 * 找出右击点落在哪个图片节点上。
 *
 * @param inside posAtCoords 给出的最近内层节点位置，可能为 null。
 * @param hit 右击命中的 DOM 元素，用于覆盖 NodeSelection 的情况。
 * @returns 图片节点的位置 px→文档位置；不是图片时返回 null。
 */
function imagePosAt(
  editor: Editor,
  inside: number | null,
  hit: HTMLElement,
): number | null {
  if (inside !== null && editor.state.doc.nodeAt(inside)?.type.name === "image")
    return inside;
  const selection = editor.state.selection;
  const onImageDom = Boolean(hit.closest("img, [data-drawing]"));
  if (
    onImageDom &&
    selection instanceof NodeSelection &&
    selection.node.type.name === "image"
  )
    return selection.from;
  return null;
}

/** 用行数估菜单高度，避免渲染后再测二次定位。 */
function estimateHeight(rows: CtxRow[]): number {
  let height = 8;
  for (const row of rows) height += "separator" in row ? SEP_H : ROW_H;
  return height;
}
