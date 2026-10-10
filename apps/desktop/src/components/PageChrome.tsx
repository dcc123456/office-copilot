import { useCallback, useEffect, useMemo, useRef, useState, type CSSProperties, type ReactNode } from "react";
import { EditorContent, type Editor } from "@tiptap/react";
import { Pencil } from "lucide-react";
import type { SectionGeometry } from "@office-copilot/docx-io";
import { HeaderFooterEditor } from "../editor/extensions/HeaderFooterEditor.js";
import { renderHeaderFooterHtml } from "../editor/renderHeaderFooter.js";
import { borderLinesFor, borderSideToCss } from "../editor/pageBorders.js";
import { useUiStore } from "../uiStore.js";
import "../styles/pageColumns.css";
import "../styles/pageBorders.css";
import "../styles/dropCap.css";

/**
 * Paginated editing stage that visually mirrors Microsoft Word:
 *
 *   .editor-wrap                   (scrolls vertically)
 *     .page-sizer
 *       .page-stage (display:flex; flex-wrap:wrap)
 *         .page-slot (width = pageWidth; per slot)
 *           .page-card (white A4 sheet, contains padding/margins)
 *             .page-chrome (header/footer overlays)
 *             .page-content (ProseMirror is a single element, but we
 *                             use PagedLayout to insert .oc-page-gap
 *                             widgets at boundaries, giving the visual
 *                             effect of multiple cards)
 *
 * The PM content lives in ONE .page-card stack. We render "fake" cards
 * behind each `--oc-page-gap` so the editor appears to span multiple
 * A4 pages, with the gray editor canvas visible between them. The
 * outer flex-wrap places this stack in the centre and lets the canvas
 * show additional pages side-by-side when the viewport is wide.
 */

type HfDoc = Record<string, unknown> | null;

const DEFAULT_GEOM: SectionGeometry = {
  pageWidth: 794,
  pageHeight: 1123,
  marginTop: 96,
  marginRight: 96,
  marginBottom: 96,
  marginLeft: 96,
  header: 48,
  footer: 48,
  contentWidth: 602,
  contentHeight: 931,
};

const PAGE_GAP = 24;
const OUTER_PAD = 32;
const ROW_GAP = 24;
const COL_GAP = 32;

export function PageChrome({
  editor,
  headerDoc,
  footerDoc,
  section,
  onHeaderChange,
  onFooterChange,
  children,
}: {
  editor: Editor;
  headerDoc: HfDoc;
  footerDoc: HfDoc;
  section?: SectionGeometry | null;
  onHeaderChange: (doc: HfDoc) => void;
  onFooterChange: (doc: HfDoc) => void;
  children?: ReactNode;
}) {
  const g = section ?? DEFAULT_GEOM;
  const editorEl = editor.view.dom as HTMLElement;
  const zoom = useUiStore((s) => s.zoom);
  const [pageCount, setPageCount] = useState(1);
  const [currentPage, setCurrentPage] = useState(1);
  const [editing, setEditing] = useState<null | "header" | "footer">(null);
  const wrapRef = useRef<HTMLDivElement>(null);
  const sizerRef = useRef<HTMLDivElement>(null);
  const stageRef = useRef<HTMLDivElement>(null);

  /**
   * 栏边界在正文区里的横向偏移（CSS px，相对页面左缘要再加 marginLeft）。
   *
   * 编辑视图不真的把文字排成多栏——`PagedLayout` 按块高度算分页，CSS columns
   * 会把块拆断让量测失效——所以这里只算导引线的位置：等宽按
   * (正文宽 - 间距*(栏数-1))/栏数，不等宽用文件里读到的逐栏宽度。
   * 单栏或没有 `w:cols` 时返回空数组，画不出任何东西。
   */
  const columnGuides = useMemo(() => {
    const cols = g.columns;
    if (!cols || cols.count < 2) return [];
    const bounds: number[] = [];
    if (!cols.equalWidth && cols.widths.length === cols.count) {
      let x = 0;
      for (let i = 0; i < cols.widths.length - 1; i++) {
        x += cols.widths[i]! + cols.gap;
        bounds.push(x);
      }
      return bounds;
    }
    const colWidth = (g.contentWidth - cols.gap * (cols.count - 1)) / cols.count;
    for (let i = 1; i < cols.count; i++) bounds.push(i * (colWidth + cols.gap));
    return bounds;
  }, [g.columns, g.contentWidth]);

  /**
   * 页面边框（`w:pgBorders`）在画布上的四条线：每条带着自己在页面坐标里的位置与
   * CSS 线型。偏移基准（距页面 / 距正文）由 docx-io 读进来的 `offsetFrom` 决定，
   * 与 Word 画的位置一致；文档没有边框时数组为空。
   */
  const borderLines = useMemo(
    () => (g.pageBorders ? borderLinesFor(g.pageBorders, g) : []),
    [g],
  );

  // Push page geometry into CSS variables for the PagedLayout plugin
  // and the editor styles to consume.
  useEffect(() => {
    editorEl.style.setProperty("--oc-content-width", `${g.contentWidth}px`);
    editorEl.style.setProperty("--oc-content-height", `${g.contentHeight}px`);
    editorEl.style.setProperty("--oc-margin-top", `${g.marginTop}px`);
    editorEl.style.setProperty("--oc-margin-bottom", `${g.marginBottom}px`);
    editorEl.style.setProperty("--oc-margin-left", `${g.marginLeft}px`);
    editorEl.style.setProperty("--oc-margin-right", `${g.marginRight}px`);
    editorEl.style.setProperty("--oc-page-gap", `${PAGE_GAP}px`);
    editorEl.style.setProperty("--oc-page-width", `${g.pageWidth}px`);
    editorEl.style.setProperty("--oc-page-height", `${g.pageHeight}px`);
    editorEl.style.setProperty("--oc-header", `${g.header}px`);
    editorEl.style.setProperty("--oc-footer", `${g.footer}px`);
    // The document's own flow metrics: Word lays body lines on its grid and
    // leaves no gap when the default style asks for none. Emptying a variable
    // falls back to the app default in styles.css.
    editorEl.style.setProperty(
      "--oc-line-pitch",
      g.linePitch ? `${g.linePitch}px` : "",
    );
    editorEl.style.setProperty(
      "--oc-para-gap",
      g.paraSpacingAfter === undefined ? "" : `${g.paraSpacingAfter}px`,
    );
    editorEl.style.width = `${g.contentWidth}px`;
  }, [
    editorEl,
    g.contentWidth,
    g.contentHeight,
    g.marginTop,
    g.marginBottom,
    g.marginLeft,
    g.marginRight,
    g.pageWidth,
    g.pageHeight,
    g.header,
    g.footer,
    g.linePitch,
    g.paraSpacingAfter,
  ]);

  // Listen for page count updates from the plugin.
  useEffect(() => {
    const handler = (e: Event) => {
      const n = (e as CustomEvent<{ pages: number }>).detail.pages;
      setPageCount(Math.max(1, n));
    };
    editorEl.addEventListener("oc:pagecount", handler);
    return () => editorEl.removeEventListener("oc:pagecount", handler);
  }, [editorEl]);

  // Listen for paginate events (legacy) — re-measure.
  useEffect(() => {
    const handler = () => editorEl.dispatchEvent(new CustomEvent("oc:pagecount"));
    editorEl.addEventListener("oc:paginate", handler);
    return () => editorEl.removeEventListener("oc:paginate", handler);
  }, [editorEl]);

  // Total stack height including page gaps.
  const totalHeight =
    pageCount * g.pageHeight + (pageCount - 1) * PAGE_GAP;

  // Auto-fit columns based on wrap width.
  const [cols, setCols] = useState(1);
  useEffect(() => {
    const compute = () => {
      const wrap = wrapRef.current;
      if (!wrap) return;
      const avail = wrap.clientWidth - OUTER_PAD * 2;
      const c = Math.max(
        1,
        Math.floor((avail + COL_GAP) / (g.pageWidth + COL_GAP)),
      );
      setCols(c);
    };
    compute();
    const ro = new ResizeObserver(compute);
    if (wrapRef.current) ro.observe(wrapRef.current);
    window.addEventListener("resize", compute);
    return () => {
      ro.disconnect();
      window.removeEventListener("resize", compute);
    };
  }, [g.pageWidth]);

  // When cols>1, we don't actually split the editor into multiple
  // columns; we center the stack and let the canvas show multiple
  // "shadow" pages side-by-side for the Word "two pages" look.
  const stageWidth = useMemo(() => {
    // We render `cols` page-width columns; the editor stack itself
    // lives in the first column. The remaining columns are paper
    // shadows for visual effect.
    return cols * g.pageWidth + (cols - 1) * COL_GAP;
  }, [cols, g.pageWidth]);

  const rows = Math.max(1, Math.ceil(pageCount / 1));
  const stageHeight = rows * g.pageHeight + (rows - 1) * ROW_GAP;

  // Update stage sizing.
  useEffect(() => {
    if (stageRef.current) {
      stageRef.current.style.width = `${Math.max(stageWidth, g.pageWidth)}px`;
      stageRef.current.style.height = `${Math.max(stageHeight, totalHeight)}px`;
    }
    if (sizerRef.current) {
      sizerRef.current.style.width = `${Math.max(stageWidth, g.pageWidth)}px`;
      sizerRef.current.style.height = `${Math.max(stageHeight, totalHeight)}px`;
    }
  }, [stageWidth, stageHeight, g.pageWidth, totalHeight]);

  // Zoom scale.
  useEffect(() => {
    const stage = stageRef.current;
    const sizer = sizerRef.current;
    if (!stage || !sizer) return;
    stage.style.transform = `scale(${zoom})`;
    stage.style.transformOrigin = "0 0";
    const r = stage.getBoundingClientRect();
    sizer.style.width = `${r.width * zoom}px`;
    sizer.style.height = `${r.height * zoom}px`;
  }, [zoom, stageWidth, stageHeight, g.pageWidth, totalHeight]);

  // Current page from scroll position.
  const updateCurrentPage = useCallback(() => {
    const wrap = wrapRef.current;
    if (!wrap) return;
    const top = wrap.scrollTop;
    const p = Math.max(1, Math.floor((top + g.pageHeight / 2) / (g.pageHeight + PAGE_GAP)) + 1);
    setCurrentPage(Math.min(pageCount, p));
  }, [g.pageHeight, pageCount]);

  useEffect(() => {
    const wrap = wrapRef.current;
    wrap?.addEventListener("scroll", updateCurrentPage, { passive: true });
    updateCurrentPage();
    return () => wrap?.removeEventListener("scroll", updateCurrentPage);
  }, [updateCurrentPage]);

  // Broadcast to status bar.
  useEffect(() => {
    editorEl.dispatchEvent(
      new CustomEvent("oc:pagechange", {
        bubbles: true,
        detail: { page: currentPage, total: pageCount },
      }),
    );
  }, [currentPage, pageCount, editorEl]);

  // Outline click → scroll to the row containing the target.
  useEffect(() => {
    const handler = (e: Event) => {
      const pos = (e as CustomEvent<number>).detail;
      if (typeof pos !== "number") return;
      const wrap = wrapRef.current;
      if (!wrap) return;
      const raw: Node | null =
        editor.view.nodeDOM(pos) ?? editor.view.domAtPos(pos).node;
      let block: HTMLElement | null =
        ((raw?.nodeType === Node.ELEMENT_NODE ? raw : raw?.parentNode) ??
          null) as HTMLElement | null;
      while (
        block &&
        block !== editorEl &&
        !/^(P|H1|H2|H3|UL|OL|PRE|BLOCKQUOTE|TABLE|LI)$/.test(block.tagName)
      ) {
        block = block.parentElement;
      }
      const target = block ?? editorEl;
      // Both rects are post-transform screen px and the sizer is sized in
      // screen px, so this offset holds at any zoom level.
      const delta =
        target.getBoundingClientRect().top - wrap.getBoundingClientRect().top;
      wrap.scrollTo({
        top: wrap.scrollTop + delta - OUTER_PAD - 8,
        behavior: "smooth",
      });
      editor.commands.setTextSelection(pos);
    };
    window.addEventListener("oc:scroll-to-pos", handler);
    return () => window.removeEventListener("oc:scroll-to-pos", handler);
  }, [editor, editorEl]);

  // The pages for the header/footer overlay — one per page, stacked
  // vertically at the editor position.
  const pages = useMemo(
    () =>
      Array.from({ length: pageCount }, (_, i) => ({
        i,
        y: i * (g.pageHeight + PAGE_GAP),
      })),
    [pageCount, g.pageHeight],
  );

  const headerTop = g.header;
  const headerHeight = g.marginTop - g.header;
  const footerHeight = g.marginBottom - g.footer;
  const footerTop = g.pageHeight - g.footer - footerHeight;
  const chromeLeft = g.marginLeft;
  const chromeWidth = g.contentWidth;

  return (
    <>
      <div
        ref={wrapRef}
        className="editor-wrap"
        style={{ padding: `${OUTER_PAD}px` }}
      >
        <div ref={sizerRef} className="page-sizer">
          <div ref={stageRef} className="page-stage paged-stage">
            {/* Shadow pages for multi-column visual effect when cols>1 */}
            {cols > 1 &&
              Array.from({ length: (cols - 1) * rows }).map((_, i) => {
                const col = (i % (cols - 1)) + 1;
                const row = Math.floor(i / (cols - 1));
                return (
                  <div
                    key={`shadow-${i}`}
                    className="page-card page-card-shadow"
                    style={{
                      left: col * (g.pageWidth + COL_GAP),
                      top: row * (g.pageHeight + ROW_GAP),
                      width: g.pageWidth,
                      height: g.pageHeight,
                    }}
                  />
                );
              })}
            {/* The single real page stack containing the editor. */}
            {pages.map((p) => (
              <div
                key={`card-${p.i}`}
                className="page-card page-card-real"
                style={{
                  left: 0,
                  top: p.y,
                  width: g.pageWidth,
                  height: g.pageHeight,
                }}
              />
            ))}
            {/* 分栏导引线：每页一组，位置见 columnGuides。 */}
            {columnGuides.map((offset, boundary) =>
              pages.map((p) => (
                <div
                  key={`col-${boundary}-${p.i}`}
                  className="oc-column-guide"
                  style={{
                    left: g.marginLeft + offset,
                    top: p.y + g.marginTop,
                    height: g.contentHeight,
                  }}
                />
              )),
            )}
            {/* 页面边框：每条边一个零尺寸盒子，线型/线宽/颜色经 CSS 变量交给 pageBorders.css 画。 */}
            {borderLines.map((line) =>
              pages.map((p) => {
                const css = borderSideToCss(line.side);
                return (
                  <div
                    key={`border-${line.key}-${p.i}`}
                    className="oc-page-border"
                    style={
                      {
                        left: line.x,
                        top: p.y + line.y,
                        width: line.horizontal ? line.length : undefined,
                        height: line.horizontal ? undefined : line.length,
                        "--oc-border-width": `${css.widthPx}px`,
                        "--oc-border-style": css.style,
                        "--oc-border-color": css.color,
                      } as CSSProperties
                    }
                  />
                );
              }),
            )}
            <div
              className="prosemirror-host"
              style={{
                position: "absolute",
                left: g.marginLeft,
                top: g.marginTop,
                width: g.contentWidth,
              }}
            >
              <EditorContent editor={editor} />
            </div>
            {children}
            {pages.map((p) => (
              <div
                key={`h-${p.i}`}
                className="page-chrome page-header"
                style={{
                  left: chromeLeft,
                  top: p.y + headerTop,
                  width: chromeWidth,
                  minHeight: headerHeight,
                }}
                onDoubleClick={() => setEditing("header")}
                title="双击编辑页眉"
              >
                <div
                  className="hf-static"
                  dangerouslySetInnerHTML={{
                    __html:
                      renderHeaderFooterHtml(
                        headerDoc as never,
                        p.i + 1,
                        pageCount,
                      ) || `<span class="hf-placeholder">双击添加页眉</span>`,
                  }}
                />
                <Pencil size={10} className="hf-edit-hint" />
              </div>
            ))}
            {pages.map((p) => (
              <div
                key={`f-${p.i}`}
                className="page-chrome page-footer"
                style={{
                  left: chromeLeft,
                  top: p.y + footerTop,
                  width: chromeWidth,
                  minHeight: footerHeight,
                }}
                onDoubleClick={() => setEditing("footer")}
                title="双击编辑页脚"
              >
                <div
                  className="hf-static"
                  dangerouslySetInnerHTML={{
                    __html:
                      renderHeaderFooterHtml(
                        footerDoc as never,
                        p.i + 1,
                        pageCount,
                      ) || `<span class="hf-placeholder">双击添加页脚</span>`,
                  }}
                />
                <Pencil size={10} className="hf-edit-hint" />
              </div>
            ))}
            {editing && (
              <div
                className="hf-popover"
                style={{
                  left: 16,
                  top: 16,
                  width: Math.min(520, g.pageWidth - 32),
                  zIndex: 50,
                }}
              >
                <HeaderFooterEditor
                  kind={editing}
                  initialDoc={editing === "header" ? headerDoc : footerDoc}
                  currentPage={currentPage}
                  pageCount={pageCount}
                  onUpdate={(doc) => {
                    if (editing === "header") onHeaderChange(doc);
                    else onFooterChange(doc);
                  }}
                  toolbarExtras={
                    <button
                      type="button"
                      style={{ marginLeft: "auto" }}
                      onClick={() => setEditing(null)}
                    >
                      完成
                    </button>
                  }
                />
              </div>
            )}
          </div>
        </div>
      </div>
    </>
  );
}
