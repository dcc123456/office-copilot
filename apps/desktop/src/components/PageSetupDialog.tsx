import { useEffect, useState } from "react";

/**
 * Page Setup dialog (Word: 文件 → 页面设置). Lets the user change:
 *   - paper size (A4 / Letter / A3 / Legal / custom)
 *   - orientation (portrait / landscape) — auto-swaps W/H
 *   - margins (top/right/bottom/left) and header/footer distances
 *
 * On save we produce a complete SectionGeometry object in CSS px at
 * 96 DPI. The app applies it immediately by replacing the section
 * state and writes it into the .docx's <w:sectPr> on save.
 *
 * Values are shown in centimetres (Word on Chinese-locale systems uses
 * cm) for familiarity; the conversion is cm -> px via (cm * 96 / 2.54).
 */

interface SectionGeometryLike {
  pageWidth: number;
  pageHeight: number;
  marginTop: number;
  marginRight: number;
  marginBottom: number;
  marginLeft: number;
  header: number;
  footer: number;
  contentWidth: number;
  contentHeight: number;
}

/**
 * Page geometry in CSS px at 96 DPI. Shared with 美化排版 so a template's
 * margins produce exactly the object the 页面设置 dialog would have.
 */
export function geometryFor(
  pageWidth: number,
  pageHeight: number,
  margins: { top: number; right: number; bottom: number; left: number },
  header: number,
  footer: number,
): SectionGeometryLike {
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
  section: SectionGeometryLike | null;
  onClose: () => void;
  onApply: (s: SectionGeometryLike) => void;
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

  useEffect(() => {
    if (!open) return;
    const g = section ?? {
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
      ),
    );
    onClose();
  };

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
          <div className="ps-hint">
            当前正文区域 ≈ {pxToCm(w - cmToPx(ml) - cmToPx(mr)).toFixed(2)} cm ×{" "}
            {pxToCm(h - cmToPx(mt) - cmToPx(mb)).toFixed(2)} cm
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
