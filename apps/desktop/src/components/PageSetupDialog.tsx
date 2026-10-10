import { useEffect, useState } from "react";
import type { SectionColumns, SectionGeometry } from "@office-copilot/docx-io";

/**
 * Page Setup dialog (Word: 文件 → 页面设置). Lets the user change:
 *   - paper size (A4 / Letter / A3 / Legal / custom)
 *   - orientation (portrait / landscape) — auto-swaps W/H
 *   - margins (top/right/bottom/left) and header/footer distances
 *   - columns (栏数 + 栏间距)
 *
 * On save we produce a complete SectionGeometry object in CSS px at
 * 96 DPI. The app applies it immediately by replacing the section
 * state and writes it into the .docx's <w:sectPr> on save.
 *
 * Values are shown in centimetres (Word on Chinese-locale systems uses
 * cm) for familiarity; the conversion is cm -> px via (cm * 96 / 2.54).
 */

/** Word 的栏数上限（页面设置对话框里「栏数」输入框能填到的最大数）。 */
const MAX_COLUMNS = 13;
/** 文档没设过分栏时对话框显示的栏间距，Word 的新建多栏默认值 1.27 cm。 */
const DEFAULT_COL_GAP_CM = 1.27;

/**
 * Page geometry in CSS px at 96 DPI. Shared with 美化排版 so a template's
 * margins produce exactly the object the 页面设置 dialog would have.
 *
 * @param columns 分栏设置；不传就是"这份文档没设过分栏"，保存时不会写出 `w:cols`。
 */
export function geometryFor(
  pageWidth: number,
  pageHeight: number,
  margins: { top: number; right: number; bottom: number; left: number },
  header: number,
  footer: number,
  columns?: SectionColumns,
): SectionGeometry {
  return {
    pageWidth,
    pageHeight,
    marginTop: margins.top,
    marginRight: margins.right,
    marginBottom: margins.bottom,
    marginLeft: margins.left,
    header,
    footer,
    contentWidth: pageWidth - margins.left - margins.right,
    contentHeight: pageHeight - margins.top - margins.bottom,
    columns,
  };
}

/**
 * 把对话框里的分栏选择收成 `SectionColumns`。
 *
 * @param count 栏数，1 就是单栏（Word 也会老老实实写 `w:num="1"`）。
 * @param gapPx 栏间距，CSS px。
 * @param source 文档里原本那份分栏。只有用户没改栏数时才把逐栏宽度原样带回，
 *   免得"只改了页边距"就把不等宽栏拍平（AGENTS §2.7）。
 * @returns 要写进节属性的分栏设置。
 */
export function columnsFor(
  count: number,
  gapPx: number,
  source: SectionColumns | null | undefined,
): SectionColumns {
  const keepWidths = source && source.count === count && !source.equalWidth;
  return {
    count,
    gap: gapPx,
    equalWidth: !keepWidths,
    widths: keepWidths ? source!.widths : [],
  };
}

const CM = 96 / 2.54;
const pxToCm = (px: number) => Math.round((px / CM) * 100) / 100;
const cmToPx = (cm: number) => Math.round(cm * CM * 10) / 10;

interface Paper {
  id: string;
  label: string;
  w: number; // px at 96dpi
  h: number;
}
const PAPERS: Paper[] = [
  { id: "A4", label: "A4", w: 793.7, h: 1122.5 },
  { id: "A3", label: "A3", w: 1122.5, h: 1587.4 },
  { id: "A5", label: "A5", w: 559.4, h: 793.7 },
  { id: "Letter", label: "Letter (US)", w: 816, h: 1056 },
  { id: "Legal", label: "Legal (US)", w: 816, h: 1344 },
];

export function PageSetupDialog({
  open,
  section,
  onClose,
  onApply,
}: {
  open: boolean;
  section: SectionGeometry | null;
  onClose: () => void;
  onApply: (s: SectionGeometry) => void;
}) {
  const [paper, setPaper] = useState<string>("A4");
  const [custom, setCustom] = useState(false);
  const [orientation, setOrientation] = useState<"portrait" | "landscape">(
    "portrait",
  );
  const [w, setW] = useState(793.7);
  const [h, setH] = useState(1122.5);
  const [mt, setMt] = useState(2.54);
  const [mr, setMr] = useState(3.18);
  const [mb, setMb] = useState(2.54);
  const [ml, setMl] = useState(3.18);
  const [hd, setHd] = useState(1.5);
  const [ft, setFt] = useState(1.75);
  const [colNum, setColNum] = useState(1);
  const [colGapCm, setColGapCm] = useState(DEFAULT_COL_GAP_CM);
  // 文档原本的分栏，用来在"用户没碰栏数"时把不等宽的逐栏宽度原样带回。
  const [sourceColumns, setSourceColumns] = useState<SectionColumns | null>(null);

  useEffect(() => {
    if (!open) return;
    const g: SectionGeometry = section ?? {
      pageWidth: 793.7,
      pageHeight: 1122.5,
      marginTop: cmToPx(2.54),
      marginRight: cmToPx(3.18),
      marginBottom: cmToPx(2.54),
      marginLeft: cmToPx(3.18),
      header: cmToPx(1.5),
      footer: cmToPx(1.75),
      contentWidth: 601.7,
      contentHeight: 930.5,
    };
    const matched =
      PAPERS.find(
        (p) =>
          (Math.abs(p.w - g.pageWidth) < 1 && Math.abs(p.h - g.pageHeight) < 1) ||
          (Math.abs(p.w - g.pageHeight) < 1 && Math.abs(p.h - g.pageWidth) < 1),
      ) ?? null;
    setPaper(matched?.id ?? "custom");
    setCustom(!matched);
    setOrientation(g.pageHeight >= g.pageWidth ? "portrait" : "landscape");
    setW(g.pageWidth);
    setH(g.pageHeight);
    setMt(pxToCm(g.marginTop));
    setMr(pxToCm(g.marginRight));
    setMb(pxToCm(g.marginBottom));
    setMl(pxToCm(g.marginLeft));
    setHd(pxToCm(g.header));
    setFt(pxToCm(g.footer));
    setColNum(g.columns?.count ?? 1);
    setColGapCm(g.columns ? pxToCm(g.columns.gap) : DEFAULT_COL_GAP_CM);
    setSourceColumns(g.columns ?? null);
  }, [open, section]);

  if (!open) return null;

  const changePaper = (id: string) => {
    setPaper(id);
    if (id === "custom") {
      setCustom(true);
      return;
    }
    setCustom(false);
    const p = PAPERS.find((x) => x.id === id)!;
    if (orientation === "portrait") {
      setW(p.w);
      setH(p.h);
    } else {
      setW(p.h);
      setH(p.w);
    }
  };
  const changeOrientation = (o: "portrait" | "landscape") => {
    setOrientation(o);
    if (paper !== "custom") {
      const p = PAPERS.find((x) => x.id === paper)!;
      setW(o === "portrait" ? p.w : p.h);
      setH(o === "portrait" ? p.h : p.w);
    } else {
      setW(h);
      setH(w);
    }
  };

  const apply = () => {
    onApply(
      geometryFor(
        w,
        h,
        {
          top: cmToPx(mt),
          right: cmToPx(mr),
          bottom: cmToPx(mb),
          left: cmToPx(ml),
        },
        cmToPx(hd),
        cmToPx(ft),
        columnsFor(colNum, cmToPx(colGapCm), sourceColumns),
      ),
    );
    onClose();
  };

  // 正文区与每栏宽度都用同一份 cm 值，避免 `w - ml - mr` 那种再换算一遍的写法。
  const bodyWidthCm = pxToCm(w - cmToPx(ml) - cmToPx(mr));
  const bodyHeightCm = pxToCm(h - cmToPx(mt) - cmToPx(mb));
  const columnWidthCm =
    colNum > 1 ? (bodyWidthCm - colGapCm * (colNum - 1)) / colNum : bodyWidthCm;

  return (
    <div className="modal-backdrop" onClick={onClose}>
      <div className="modal" onClick={(e) => e.stopPropagation()}>
        <div className="modal-header">
          <strong>页面设置</strong>
        </div>
        <div className="modal-body">
          <div className="ps-row">
            <label>纸张大小</label>
            <select
              value={paper}
              onChange={(e) => changePaper(e.target.value)}
            >
              {PAPERS.map((p) => (
                <option key={p.id} value={p.id}>
                  {p.label}
                </option>
              ))}
              <option value="custom">自定义…</option>
            </select>
          </div>
          <div className="ps-row">
            <label>方向</label>
            <div className="ps-seg">
              <button
                type="button"
                className={orientation === "portrait" ? "active" : ""}
                onClick={() => changeOrientation("portrait")}
              >
                纵向
              </button>
              <button
                type="button"
                className={orientation === "landscape" ? "active" : ""}
                onClick={() => changeOrientation("landscape")}
              >
                横向
              </button>
            </div>
          </div>
          {custom && (
            <div className="ps-grid">
              <label>
                宽度 (cm)
                <input
                  type="number"
                  step="0.1"
                  min="1"
                  value={pxToCm(w)}
                  onChange={(e) => {
                    setW(cmToPx(Number(e.target.value)));
                    setPaper("custom");
                  }}
                />
              </label>
              <label>
                高度 (cm)
                <input
                  type="number"
                  step="0.1"
                  min="1"
                  value={pxToCm(h)}
                  onChange={(e) => {
                    setH(cmToPx(Number(e.target.value)));
                    setPaper("custom");
                  }}
                />
              </label>
            </div>
          )}
          <div className="ps-section-title">页边距 (cm)</div>
          <div className="ps-grid four">
            <label>
              上
              <input
                type="number"
                step="0.1"
                value={mt}
                onChange={(e) => setMt(Number(e.target.value))}
              />
            </label>
            <label>
              下
              <input
                type="number"
                step="0.1"
                value={mb}
                onChange={(e) => setMb(Number(e.target.value))}
              />
            </label>
            <label>
              左
              <input
                type="number"
                step="0.1"
                value={ml}
                onChange={(e) => setMl(Number(e.target.value))}
              />
            </label>
            <label>
              右
              <input
                type="number"
                step="0.1"
                value={mr}
                onChange={(e) => setMr(Number(e.target.value))}
              />
            </label>
            <label>
              页眉距边
              <input
                type="number"
                step="0.1"
                value={hd}
                onChange={(e) => setHd(Number(e.target.value))}
              />
            </label>
            <label>
              页脚距边
              <input
                type="number"
                step="0.1"
                value={ft}
                onChange={(e) => setFt(Number(e.target.value))}
              />
            </label>
          </div>
          <div className="ps-section-title">分栏</div>
          <div className="ps-grid">
            <label>
              栏数
              <input
                type="number"
                min="1"
                max={MAX_COLUMNS}
                step="1"
                value={colNum}
                onChange={(e) =>
                  setColNum(
                    Math.max(
                      1,
                      Math.min(MAX_COLUMNS, Math.round(Number(e.target.value) || 1)),
                    ),
                  )
                }
              />
            </label>
            <label>
              栏间距 (cm)
              <input
                type="number"
                min="0"
                step="0.05"
                value={colGapCm}
                disabled={colNum < 2}
                onChange={(e) => setColGapCm(Math.max(0, Number(e.target.value) || 0))}
              />
            </label>
          </div>
          <div className="ps-hint">
            当前正文区域 ≈ {bodyWidthCm.toFixed(2)} cm × {bodyHeightCm.toFixed(2)} cm
            {colNum > 1 && `，每栏 ≈ ${columnWidthCm.toFixed(2)} cm`}
          </div>
        </div>
        <div className="modal-footer">
          <button onClick={onClose}>取消</button>
          <button className="primary" onClick={apply}>
            确定
          </button>
        </div>
      </div>
    </div>
  );
}
