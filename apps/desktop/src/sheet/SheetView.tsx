/**
 * 表格视图（spec 阶段 2）：虚拟滚动的网格、选区、就地编辑、列宽拖动与表签。
 *
 * 这里只做装配：几何、落地与单位换算分别在 `sheetGridModel` 与 `@office-copilot/sheet-io`，
 * 组件里不出现第二套换算，也不碰 Node 能力（`AGENTS.md` §3.3 / §5.2）。
 */
import { useCallback, useEffect, useMemo, useRef, useState, type KeyboardEvent, type PointerEvent as ReactPointerEvent } from "react";
import { formatA1, indexToColName, rangeContains, readOnlyReason, type A1Address, type A1Range, type SheetPayload, type WorkbookPayload } from "@office-copilot/sheet-io";
import {
  COLUMN_HEADER_HEIGHT_PX,
  OVERSCAN_ROWS,
  ROW_HEADER_WIDTH_PX,
  addressAtPoint,
  buildSheetLayout,
  displayTextOf,
  moveAddress,
  selectionLabel,
  selectionRectOf,
  visibleSlotRange,
  withCellText,
  withColumnWidth,
} from "./sheetGridModel.js";
import { startPointerDrag } from "../hooks/usePointerDrag.js";
import "../styles/sheet.css";

export interface SheetViewProps {
  /** 打开的工作簿载荷；表的数量与顺序在保存时不许变。 */
  workbook: WorkbookPayload;
  /** 当前表签。 */
  activeIndex: number;
  onSelectSheet: (index: number) => void;
  /** 当前表被改过后的新载荷。 */
  onChangeSheet: (sheet: SheetPayload) => void;
}

/** 表格里没选中任何格时的初始位置。 */
const ORIGIN: A1Address = { colIndex: 0, rowIndex: 0 };

/** 列宽热区宽度（px），只影响命中，不改变文档里的宽度。 */
const RESIZE_HOTZONE_PX = 6;

/**
 * 表格视图。
 *
 * @param props 见 `SheetViewProps`。
 * @returns 一张表的网格；工作簿里没有任何表时给出空态而不是崩溃。
 */
export function SheetView({ workbook, activeIndex, onSelectSheet, onChangeSheet }: SheetViewProps) {
  const sheet = workbook.sheets[activeIndex];
  const layout = useMemo(() => (sheet ? buildSheetLayout(sheet) : undefined), [sheet]);
  /** 大表只读（spec D19）：阈值与理由都在 `sheet-io`，这里只按它拦住所有写入路径，不做第二处状态。 */
  const readOnlyHint = useMemo(() => readOnlyReason(workbook), [workbook]);
  const isReadOnly = readOnlyHint !== null;

  const scrollRef = useRef<HTMLDivElement | null>(null);
  const colHeaderRef = useRef<HTMLDivElement | null>(null);
  const rowHeaderRef = useRef<HTMLDivElement | null>(null);
  const [rowSlot, setRowSlot] = useState<[number, number]>([0, 30]);
  const [colSlot, setColSlot] = useState<[number, number]>([0, 40]);
  const [viewport, setViewport] = useState({ width: 800, height: 480 });
  const [active, setActive] = useState<A1Address>(ORIGIN);
  const [anchor, setAnchor] = useState<A1Address>(ORIGIN);
  const [editing, setEditing] = useState<{ address: string; text: string } | null>(null);
  const [resizing, setResizing] = useState<{ colIndex: number; widthPx: number } | null>(null);

  const range = selectionRectOf(anchor, active);

  /**
   * 按滚动距离重算要渲染的行列槽位，并让两条表头跟着滚。
   *
   * @param scrollLeft 横向滚动距离（px）。
   * @param scrollTop 纵向滚动距离（px）。
   */
  const applyScroll = useCallback(
    (scrollLeft: number, scrollTop: number) => {
      if (colHeaderRef.current) colHeaderRef.current.style.transform = `translateX(${-scrollLeft}px)`;
      if (rowHeaderRef.current) rowHeaderRef.current.style.transform = `translateY(${-scrollTop}px)`;
      if (!layout) return;
      const nextRows = visibleSlotRange(layout.rowOffsetPx, scrollTop, viewport.height, OVERSCAN_ROWS);
      const nextCols = visibleSlotRange(layout.colOffsetPx, scrollLeft, viewport.width, OVERSCAN_ROWS);
      setRowSlot((previous) => (previous[0] === nextRows[0] && previous[1] === nextRows[1] ? previous : nextRows));
      setColSlot((previous) => (previous[0] === nextCols[0] && previous[1] === nextCols[1] ? previous : nextCols));
    },
    [layout, viewport.height, viewport.width],
  );

  // 视口尺寸变化（含首次挂载）要重算窗口，否则多滚两下就露白。
  useEffect(() => {
    const el = scrollRef.current;
    if (!el) return;
    const measure = () => {
      setViewport({ width: el.clientWidth, height: el.clientHeight });
      applyScroll(el.scrollLeft, el.scrollTop);
    };
    measure();
    const observer = new ResizeObserver(measure);
    observer.observe(el);
    return () => observer.disconnect();
  }, [applyScroll]);

  const scrollToAddress = useCallback((address: A1Address) => {
    const el = scrollRef.current;
    if (!el || !layout) return;
    const slotCol = layout.colIndexes.indexOf(address.colIndex);
    const slotRow = layout.rowIndexes.indexOf(address.rowIndex);
    if (slotCol < 0 || slotRow < 0) return;
    const x = layout.colOffsetPx[slotCol] ?? 0;
    const y = layout.rowOffsetPx[slotRow] ?? 0;
    if (x < el.scrollLeft || x + (layout.colWidthPx[slotCol] ?? 0) > el.scrollLeft + el.clientWidth) {
      el.scrollLeft = Math.max(0, x - el.clientWidth + (layout.colWidthPx[slotCol] ?? 0) + ROW_HEADER_WIDTH_PX);
    }
    if (y < el.scrollTop || y + (layout.rowHeightPx[slotRow] ?? 0) > el.scrollTop + el.clientHeight) {
      el.scrollTop = Math.max(0, y - el.clientHeight + (layout.rowHeightPx[slotRow] ?? 0) + COLUMN_HEADER_HEIGHT_PX);
    }
  }, [layout]);

  /**
   * 把编辑框里的文字落到当前格。
   *
   * @param text 输入框内容。
   * @param move `down` 回车后到下一行，`right` Tab 后到下一列，`none` 原地不动。
   */
  const commitText = useCallback(
    (text: string, move: "down" | "right" | "none") => {
      if (!sheet || !layout || isReadOnly) return;
      const address = formatA1(active);
      onChangeSheet(withCellText(sheet, address, text));
      setEditing(null);
      const next = moveAddress(
        active,
        move === "right" ? 1 : 0,
        move === "down" ? 1 : 0,
        layout.colIndexes[layout.colIndexes.length - 1] ?? 0,
        layout.rowIndexes[layout.rowIndexes.length - 1] ?? 0,
      );
      setActive(next);
      setAnchor(next);
      scrollToAddress(next);
    },
    [active, isReadOnly, layout, onChangeSheet, scrollToAddress, sheet],
  );

  /**
   * 让选区挪到一格。
   *
   * @param next 目标格。
   * @param extend 为 true 时保留锚点（拖选 / Shift+方向键），否则锚点跟着走。
   */
  const moveSelection = useCallback(
    (next: A1Address, extend: boolean) => {
      setActive(next);
      if (!extend) setAnchor(next);
      scrollToAddress(next);
    },
    [scrollToAddress],
  );

  /**
   * 打开当前格的编辑框（初值是格里的现文本）。
   *
   * 只读工作簿里什么都不做——所有写入路径都汇到 `commitText`，这里只是提前拦住无谓的编辑态。
   */
  const startEditing = () => {
    if (isReadOnly) return;
    setEditing({ address: formatA1(active), text: cellTextAt(sheet, active) });
  };

  /**
   * 网格的键盘导航。
   *
   * @param event React 键盘事件；编辑框打开时不处理，交给输入框自己。
   */
  const onKeyDown = (event: KeyboardEvent<HTMLDivElement>) => {
    if (!layout || editing) return;
    const maxCol = (layout.colIndexes.length || 1) - 1;
    const maxRow = (layout.rowIndexes.length || 1) - 1;
    const steps: Record<string, [number, number]> = {
      ArrowLeft: [-1, 0],
      ArrowRight: [1, 0],
      ArrowUp: [0, -1],
      ArrowDown: [0, 1],
      Tab: [1, 0],
      Enter: [0, 1],
    };
    const step = steps[event.key];
    if (step && !(event.ctrlKey || event.metaKey)) {
      event.preventDefault();
      moveSelection(moveAddress(active, step[0], step[1], maxCol, maxRow), event.shiftKey);
      if (event.key === "Enter") startEditing();
      return;
    }
    if (event.key === "F2") {
      event.preventDefault();
      startEditing();
      return;
    }
    if (event.key === "Delete" || event.key === "Backspace") {
      event.preventDefault();
      commitText("", "none");
      return;
    }
    if (event.key === "Home") {
      event.preventDefault();
      moveSelection(event.ctrlKey || event.metaKey ? ORIGIN : { colIndex: 0, rowIndex: active.rowIndex }, false);
      return;
    }
    if (event.key === "PageDown" || event.key === "PageUp") {
      event.preventDefault();
      const span = Math.max(1, rowSlot[1] - rowSlot[0] - OVERSCAN_ROWS * 2);
      moveSelection(moveAddress(active, 0, event.key === "PageDown" ? span : -span, maxCol, maxRow), false);
      return;
    }
    // 可打印字符直接进编辑框，并把这个字符作为初值，省一次点击。
    if (event.key.length === 1 && !isReadOnly && !event.ctrlKey && !event.metaKey && !event.altKey) {
      event.preventDefault();
      setEditing({ address: formatA1(active), text: event.key });
    }
  };

  /**
   * 在画布上按下指针：选中该格并开始拖选。
   *
   * @param event 指针事件，坐标相对画布。
   */
  const onCanvasPointerDown = (event: ReactPointerEvent<HTMLDivElement>) => {
    if (!layout) return;
    const rect = event.currentTarget.getBoundingClientRect();
    const address = addressAtPoint(layout, event.clientX - rect.left, event.clientY - rect.top);
    setEditing(null);
    setAnchor(address);
    setActive(address);
    const target = event.currentTarget;
    startPointerDrag(target, event.pointerId, {
      move: (moveEvent) => {
        const moved = addressAtPoint(layout, moveEvent.clientX - rect.left, moveEvent.clientY - rect.top);
        setActive(moved);
      },
    });
  };

  if (!sheet || !layout) {
    return <div className="sheet-empty">这份工作簿里没有可读的工作表。</div>;
  }

  const activeAddress = formatA1(active);
  const activeCell = sheet.cells[activeAddress];
  const [rowStart, rowEnd] = rowSlot;
  const [colStart, colEnd] = colSlot;
  const cells: React.ReactNode[] = [];
  for (let slotRow = rowStart; slotRow < rowEnd; slotRow += 1) {
    const rowIndex = layout.rowIndexes[slotRow];
    if (rowIndex === undefined) continue;
    for (let slotCol = colStart; slotCol < colEnd; slotCol += 1) {
      const colIndex = layout.colIndexes[slotCol];
      if (colIndex === undefined) continue;
      const address = formatA1({ colIndex, rowIndex });
      const cell = sheet.cells[address];
      const isHead = active.colIndex === colIndex && active.rowIndex === rowIndex;
      const widthPx = slotWidth(resizing, colIndex, layout.colWidthPx[slotCol] ?? 64);
      cells.push(
        <div
          key={address}
          className={[
            "sheet-cell",
            cell?.kind === "number" ? "sheet-cell--num" : "",
            cell?.formula ? "sheet-cell--formula" : "",
            rangeContains(range, { colIndex, rowIndex }) ? "is-selected" : "",
            isHead ? "is-head" : "",
          ].join(" ")}
          style={{
            left: layout.colOffsetPx[slotCol] ?? 0,
            top: layout.rowOffsetPx[slotRow] ?? 0,
            width: widthPx,
            height: layout.rowHeightPx[slotRow] ?? 20,
          }}
        >
          {editing?.address === address ? null : (
            <span className="sheet-cell-text">{displayTextOf(cell)}</span>
          )}
        </div>,
      );
    }
  }

  const mergeOverlays = sheet.merges.map((merge) => {
    const first = merge.range.start;
    const last = merge.range.end;
    const slotLeft = layout.colIndexes.indexOf(first.colIndex);
    const slotTop = layout.rowIndexes.indexOf(first.rowIndex);
    const slotRight = layout.colIndexes.indexOf(last.colIndex);
    const slotBottom = layout.rowIndexes.indexOf(last.rowIndex);
    if (slotLeft < 0 || slotTop < 0 || slotRight < 0 || slotBottom < 0) return null;
    // 覆盖层不透明，把区间内部的格线与空格一起盖掉，只留左上角那一格的文字。
    const anchorCell = sheet.cells[formatA1(first)];
    return (
      <div
        key={merge.ref}
        className={`sheet-merge${anchorCell?.kind === "number" ? " sheet-cell--num" : ""}`}
        style={{
          left: layout.colOffsetPx[slotLeft] ?? 0,
          top: layout.rowOffsetPx[slotTop] ?? 0,
          width: (layout.colOffsetPx[slotRight + 1] ?? 0) - (layout.colOffsetPx[slotLeft] ?? 0),
          height: (layout.rowOffsetPx[slotBottom + 1] ?? 0) - (layout.rowOffsetPx[slotTop] ?? 0),
        }}
      >
        <span className="sheet-cell-text">{displayTextOf(anchorCell)}</span>
      </div>
    );
  });

  const selectionSize = { cols: range.end.colIndex - range.start.colIndex + 1, rows: range.end.rowIndex - range.start.rowIndex + 1 };

  return (
    <div className="sheet-view">
      <div className="sheet-namebox">
        <span className="sheet-namebox-address">{activeAddress}</span>
        <input
          className="sheet-formulabar"
          readOnly
          value={activeCell?.formula ? `=${activeCell.formula}` : displayTextOf(activeCell)}
          aria-label="编辑栏"
        />
        {activeCell?.formula ? <span className="sheet-recalc-hint">公式显示的是文件里的缓存值，需在 Excel / WPS 里重算</span> : null}
        {isReadOnly ? (
          <span className="sheet-readonly-hint" title={readOnlyHint ?? ""}>
            只读
          </span>
        ) : null}
      </div>

      <div className="sheet-grid">
        <div className="sheet-corner" />
        <div className="sheet-colheader">
          <div className="sheet-colheader-track" ref={colHeaderRef}>
            {layout.colIndexes.map((colIndex, slot) => {
              const widthPx = slotWidth(resizing, colIndex, layout.colWidthPx[slot] ?? 64);
              const selected = colIndex >= range.start.colIndex && colIndex <= range.end.colIndex;
              return (
                <div
                  key={colIndex}
                  className={`sheet-colhead${selected ? " is-selected" : ""}`}
                  style={{ left: layout.colOffsetPx[slot] ?? 0, width: widthPx }}
                >
                  {indexToColName(colIndex)}
                  {isReadOnly ? null : (
                    <span
                      className="sheet-col-resize"
                      style={{ right: -RESIZE_HOTZONE_PX / 2 }}
                      onPointerDown={(event) => startColumnResize(event, colIndex, widthPx)}
                      title="拖动调整列宽"
                    />
                  )}
                </div>
              );
            })}
          </div>
        </div>
        <div className="sheet-rowheader">
          <div className="sheet-rowheader-track" ref={rowHeaderRef}>
            {layout.rowIndexes.map((rowIndex, slot) => {
              const selected = rowIndex >= range.start.rowIndex && rowIndex <= range.end.rowIndex;
              return (
                <div
                  key={rowIndex}
                  className={`sheet-rowhead${selected ? " is-selected" : ""}`}
                  style={{ top: layout.rowOffsetPx[slot] ?? 0, height: layout.rowHeightPx[slot] ?? 20 }}
                >
                  {rowIndex + 1}
                </div>
              );
            })}
          </div>
        </div>
        <div
          className="sheet-scroll"
          ref={scrollRef}
          tabIndex={0}
          onKeyDown={onKeyDown}
          onScroll={(event) => {
            const el = event.currentTarget;
            applyScroll(el.scrollLeft, el.scrollTop);
          }}
        >
          <div className="sheet-canvas" style={{ width: layout.totalWidthPx, height: layout.totalHeightPx }} onPointerDown={onCanvasPointerDown}>
            {cells}
            {mergeOverlays}
            <div
              className="sheet-selection"
              style={{
                left: layout.colOffsetPx[layout.colIndexes.indexOf(range.start.colIndex)] ?? 0,
                top: layout.rowOffsetPx[layout.rowIndexes.indexOf(range.start.rowIndex)] ?? 0,
                width: selectionWidthPx(layout, range),
                height: selectionHeightPx(layout, range),
              }}
            />
            {editing ? (
              <input
                className="sheet-editor"
                autoFocus
                value={editing.text}
                style={{
                  left: layout.colOffsetPx[layout.colIndexes.indexOf(active.colIndex)] ?? 0,
                  top: layout.rowOffsetPx[layout.rowIndexes.indexOf(active.rowIndex)] ?? 0,
                  width: Math.max(80, layout.colWidthPx[layout.colIndexes.indexOf(active.colIndex)] ?? 64),
                  height: layout.rowHeightPx[layout.rowIndexes.indexOf(active.rowIndex)] ?? 20,
                }}
                onChange={(event) => setEditing({ address: editing.address, text: event.target.value })}
                onBlur={() => commitText(editing.text, "none")}
                onKeyDown={(event) => {
                  if (event.key === "Escape") {
                    event.preventDefault();
                    setEditing(null);
                    scrollRef.current?.focus();
                    return;
                  }
                  if (event.key === "Enter" || event.key === "Tab") {
                    event.preventDefault();
                    commitText(editing.text, event.key === "Tab" ? "right" : "down");
                    scrollRef.current?.focus();
                  }
                }}
              />
            ) : null}
          </div>
        </div>
      </div>

      <div className="sheet-tabs">
        {workbook.sheets.map((tab, index) => (
          <button
            key={tab.partPath}
            type="button"
            className={`sheet-tab${index === activeIndex ? " is-active" : ""}`}
            onClick={() => {
              setEditing(null);
              onSelectSheet(index);
              scrollRef.current?.focus();
            }}
          >
            {tab.name}
          </button>
        ))}
        <span className="sheet-tabinfo">
          {selectionLabel(range)}
          {selectionSize.cols > 1 || selectionSize.rows > 1 ? ` · ${selectionSize.cols}×${selectionSize.rows}` : ""}
        </span>
      </div>
    </div>
  );

  /**
   * 开始拖列宽：热区上按下后由指针捕获跟随。
   *
   * @param event 热区上的指针事件。
   * @param colIndex 真实列号。
   * @param startWidthPx 起拖时的宽度（px）。
   */
  function startColumnResize(event: ReactPointerEvent<HTMLSpanElement>, colIndex: number, startWidthPx: number) {
    event.preventDefault();
    event.stopPropagation();
    const startX = event.clientX;
    setResizing({ colIndex, widthPx: startWidthPx });
    const el = event.currentTarget;
    startPointerDrag(el, event.pointerId, {
      move: (moveEvent) => {
        setResizing({ colIndex, widthPx: Math.max(20, startWidthPx + (moveEvent.clientX - startX)) });
      },
      end: () => {
        setResizing((current) => {
          if (current && sheet) onChangeSheet(withColumnWidth(sheet, colIndex, current.widthPx));
          return null;
        });
      },
    });
  }
}

/** 当前格的文字（编辑框初值）：优先原值，日期格给显示串。 */
function cellTextAt(sheet: SheetPayload | undefined, address: A1Address): string {
  if (!sheet) return "";
  const cell = sheet.cells[formatA1(address)];
  if (!cell) return "";
  if (cell.formula) return `=${cell.formula}`;
  return displayTextOf(cell);
}

/** 拖动中用预览宽度，否则用文档里的宽度。 */
function slotWidth(resizing: { colIndex: number; widthPx: number } | null, colIndex: number, fallbackPx: number): number {
  return resizing && resizing.colIndex === colIndex ? resizing.widthPx : fallbackPx;
}

/** 选区宽（px）：跨列的范围按前缀和相减。 */
function selectionWidthPx(layout: ReturnType<typeof buildSheetLayout>, range: A1Range): number {
  const leftSlot = layout.colIndexes.indexOf(range.start.colIndex);
  const rightSlot = layout.colIndexes.indexOf(range.end.colIndex);
  if (leftSlot < 0 || rightSlot < 0) return 0;
  return (layout.colOffsetPx[rightSlot + 1] ?? 0) - (layout.colOffsetPx[leftSlot] ?? 0);
}

/** 选区高（px）：跨行的范围按行前缀和相减。 */
function selectionHeightPx(layout: ReturnType<typeof buildSheetLayout>, range: A1Range): number {
  const topSlot = layout.rowIndexes.indexOf(range.start.rowIndex);
  const bottomSlot = layout.rowIndexes.indexOf(range.end.rowIndex);
  if (topSlot < 0 || bottomSlot < 0) return 0;
  return (layout.rowOffsetPx[bottomSlot + 1] ?? 0) - (layout.rowOffsetPx[topSlot] ?? 0);
}
