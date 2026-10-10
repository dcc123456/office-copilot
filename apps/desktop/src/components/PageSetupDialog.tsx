import { Fragment, useEffect, useState } from "react";
import type {
  PageBorderSide,
  SectionColumns,
  SectionGeometry,
  SectionLineNumbers,
  SectionPageBorders,
  SectionWatermark,
  WatermarkLayout,
} from "@office-copilot/docx-io";
import {
  BORDER_STYLE_OPTIONS,
  defaultBorderSide,
  emptyPageBorders,
  hasAnyBorderSide,
  type BorderSideKey,
} from "../editor/pageBorders.js";
import "../styles/pageBorders.css";

/**
 * Page Setup dialog (Word: 文件 → 页面设置). Lets the user change:
 *   - paper size (A4 / Letter / A3 / Legal / custom)
 *   - orientation (portrait / landscape) — auto-swaps W/H
 *   - margins (top/right/bottom/left) and header/footer distances
 *   - columns (栏数 + 栏间距)
 *   - page borders (四条边的线型 / 线宽 / 间距 / 颜色) and line numbering
 *   - watermark (文字 / 字体 / 字号 / 颜色 / 半透明 / 版式)
 *
 * On save we produce a complete SectionGeometry object in CSS px at
 * 96 DPI. The app applies it immediately by replacing the section
 * state and writes it into the .docx's <w:sectPr> on save.
 *
 * Values are shown in centimetres (Word on Chinese-locale systems uses
 * cm) for familiarity; the conversion is cm -> px via (cm * 96 / 2.54).
 * 边框的线宽与间距保持 pt —— 那是 `w:sz`/`w:space` 自己的单位（§17.3.4），
 * 不再换算成 cm。
 */

/** Word 的栏数上限（页面设置对话框里「栏数」输入框能填到的最大数）。 */
const MAX_COLUMNS = 13;
/** 文档没设过分栏时对话框显示的栏间距，Word 的新建多栏默认值 1.27 cm。 */
const DEFAULT_COL_GAP_CM = 1.27;
/** 行号没设过时对话框显示的"距正文"，Word 的 0.25 英寸。 */
const DEFAULT_LN_DISTANCE_CM = 0.635;
/** 四条边的中文名，顺序与 `w:pgBorders` 写子元素的顺序一致。 */
const BORDER_SIDES: readonly { key: BorderSideKey; label: string }[] = [
  { key: "top", label: "上" },
  { key: "left", label: "左" },
  { key: "bottom", label: "下" },
  { key: "right", label: "右" },
];

/**
 * Page geometry in CSS px at 96 DPI. Shared with 美化排版 so a template's
 * margins produce exactly the object the 页面设置 dialog would have.
 *
 * @param extras 文档级设置里"这一层管不着"的那几项（分栏 / 页面边框 / 行号 / 水印）。
 *   传 undefined 的字段表示这次保存不碰它——docx-io 会原样保留文件里的元素
 *   （AGENTS §2.7：模板只改页边距，不许顺手动用户的边框与行号）。
 */
export function geometryFor(
  pageWidth: number,
  pageHeight: number,
  margins: { top: number; right: number; bottom: number; left: number },
  header: number,
  footer: number,
  extras?: Pick<
    SectionGeometry,
    "columns" | "pageBorders" | "lineNumbers" | "watermark"
  >,
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
    columns: extras?.columns,
    pageBorders: extras?.pageBorders,
    lineNumbers: extras?.lineNumbers,
    watermark: extras?.watermark,
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

/**
 * 页面边框要写进节属性的值。
 *
 * @param edited 用户在对话框里动过边框没有。没动过就返回 `undefined`，
 *   让 docx-io 保持文件里原样（包括我们没建模的 `w:display`、`w:shadow`）。
 * @param draft 对话框当前的四条边。
 * @returns `undefined` 不碰，`null` 删掉 `w:pgBorders`，否则这条边框的设置。
 */
export function pageBordersFor(
  edited: boolean,
  draft: SectionPageBorders,
): SectionPageBorders | null | undefined {
  if (!edited) return undefined;
  return hasAnyBorderSide(draft) ? draft : null;
}

/**
 * 行号要写进节属性的值。
 *
 * @param edited 用户在对话框里动过行号没有；语义同 `pageBordersFor`。
 * @param form 对话框里的行号四个字段，`distancePx` 是 CSS px。
 * @returns `undefined` 不碰，`null` 关掉行号（删 `w:lnNumType`），否则写这条元素。
 */
export function lineNumbersFor(
  edited: boolean,
  form: { enabled: boolean; countBy: number; start: number; distancePx: number },
): SectionLineNumbers | null | undefined {
  if (!edited) return undefined;
  if (!form.enabled || form.countBy < 1) return null;
  return {
    countBy: Math.round(form.countBy),
    distance: form.distancePx,
    start: Math.max(1, Math.round(form.start)),
  };
}

/** 对话框里水印区的一组控件值。 */
export interface WatermarkForm {
  enabled: boolean;
  /** 水印文字，Word 的艺术字版式里唯一必填项。 */
  text: string;
  /** 中文字体名，直接进 `v:textpath` 的 `font-family`。 */
  font: string;
  /** 字号，pt；0 是 Word 的"自动"（`fitshape`），会写 `sizePt: null`。 */
  sizePt: number;
  /** `#rrggbb`，`type="color"` 保证不会是别的形状（§2.2）。 */
  color: string;
  semiTransparent: boolean;
  layout: WatermarkLayout;
}

/** 文档里无水印时对话框显示的初值，与 Word 插入水印的缺省一致：灰、半透明、对角。 */
const DEFAULT_WATERMARK_FORM: WatermarkForm = {
  enabled: false,
  text: "草稿",
  font: "宋体",
  sizePt: 0,
  color: "#c0c0c0",
  semiTransparent: true,
  layout: "diagonal",
};

/**
 * 把水印区收成要写进页眉部件的值。
 *
 * @param edited 用户在对话框里动过水印没有；语义同 `pageBordersFor`——
 *   没动过返回 `undefined`，docx-io 会原样保留页眉里已有的艺术字形状。
 * @param form 对话框里的那七个字段。
 * @returns `undefined` 不碰，`null` 删掉水印，否则要写进 `v:shape` 的那份设置。
 */
export function watermarkFor(
  edited: boolean,
  form: WatermarkForm,
): SectionWatermark | null | undefined {
  if (!edited) return undefined;
  const text = form.text.trim();
  if (!form.enabled || !text) return null;
  return {
    text,
    font: form.font.trim() || DEFAULT_WATERMARK_FORM.font,
    sizePt: form.sizePt > 0 ? form.sizePt : null,
    color: form.color,
    semiTransparent: form.semiTransparent,
    layout: form.layout,
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
  const [borders, setBorders] = useState<SectionPageBorders>(emptyPageBorders());
  const [bordersEdited, setBordersEdited] = useState(false);
  const [lnEnabled, setLnEnabled] = useState(false);
  const [lnCountBy, setLnCountBy] = useState(1);
  const [lnStart, setLnStart] = useState(1);
  const [lnDistanceCm, setLnDistanceCm] = useState(DEFAULT_LN_DISTANCE_CM);
  const [lnEdited, setLnEdited] = useState(false);
  const [wm, setWm] = useState<WatermarkForm>(DEFAULT_WATERMARK_FORM);
  const [wmEdited, setWmEdited] = useState(false);

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
    // 边框、行号与水印的"没碰过"状态用 edited 标记表示：只有用户真的改过才把值送进节属性，
    // 否则交给 docx-io 原样保留文件里的元素（含我们没建模的 w:display / w:restart）。
    setBordersEdited(false);
    setLnEdited(false);
    setWmEdited(false);
    setBorders(g.pageBorders ? { ...g.pageBorders } : emptyPageBorders());
    setLnEnabled((g.lineNumbers?.countBy ?? 0) > 0);
    setLnCountBy(g.lineNumbers?.countBy ?? 1);
    setLnStart(g.lineNumbers?.start ?? 1);
    setLnDistanceCm(pxToCm(g.lineNumbers?.distance ?? cmToPx(DEFAULT_LN_DISTANCE_CM)));
    setWm(
      g.watermark
        ? { ...g.watermark, enabled: true, sizePt: g.watermark.sizePt ?? 0 }
        : DEFAULT_WATERMARK_FORM,
    );
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

  /**
   * 改一条边的一个字段。
   *
   * @param key 上/左/下/右之一。
   * @param patch 要覆盖的字段（线型、线宽 pt、间距 pt 或颜色）。
   */
  const patchBorderSide = (
    key: BorderSideKey,
    patch: Partial<PageBorderSide>,
  ) => {
    setBordersEdited(true);
    setBorders((prev) => {
      // 文档原本没有这条边时，用户第一次动它就是"新建一条边框"，取 Word 的缺省值起步。
      const base = prev[key] ?? defaultBorderSide();
      const style = patch.style ?? base.style;
      return { ...prev, [key]: style === "none" ? null : { ...base, ...patch } };
    });
  };

  /**
   * 改水印区的一个字段。
   *
   * @param patch 要覆盖的字段。同时置上 edited——"没碰过"（保留原样）与
   *   "碰过但清空"（删掉）就靠它区分。
   */
  const patchWm = (patch: Partial<WatermarkForm>) => {
    setWmEdited(true);
    setWm((prev) => ({ ...prev, ...patch }));
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
        {
          columns: columnsFor(colNum, cmToPx(colGapCm), sourceColumns),
          pageBorders: pageBordersFor(bordersEdited, borders),
          lineNumbers: lineNumbersFor(lnEdited, {
            enabled: lnEnabled,
            countBy: lnCountBy,
            start: lnStart,
            distancePx: cmToPx(lnDistanceCm),
          }),
          watermark: watermarkFor(wmEdited, wm),
        },
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
          <div className="ps-section-title">页面边框</div>
          <div className="ps-border-grid">
            <span />
            <span>线型</span>
            <span>线宽 (pt)</span>
            <span>间距 (pt)</span>
            <span>颜色</span>
            {BORDER_SIDES.map(({ key, label }) => {
              const side = borders[key];
              return (
                <Fragment key={key}>
                  <span className="ps-border-name">{label}</span>
                  <select
                    value={side?.style ?? "none"}
                    onChange={(e) => patchBorderSide(key, { style: e.target.value })}
                  >
                    {BORDER_STYLE_OPTIONS.map((o) => (
                      <option key={o.val} value={o.val}>
                        {o.label}
                      </option>
                    ))}
                  </select>
                  <input
                    type="number"
                    min="0.25"
                    step="0.25"
                    value={side?.sizePt ?? 0.5}
                    disabled={!side}
                    onChange={(e) =>
                      patchBorderSide(key, { sizePt: Math.max(0.25, Number(e.target.value) || 0.25) })
                    }
                  />
                  <input
                    type="number"
                    min="0"
                    step="1"
                    value={side?.spacePt ?? 1}
                    disabled={!side}
                    onChange={(e) =>
                      patchBorderSide(key, { spacePt: Math.max(0, Number(e.target.value) || 0) })
                    }
                  />
                  <input
                    type="color"
                    value={side?.color ?? "#000000"}
                    disabled={!side}
                    onChange={(e) => patchBorderSide(key, { color: e.target.value })}
                  />
                </Fragment>
              );
            })}
          </div>
          <div className="ps-section-title">行号</div>
          <div className="ps-grid">
            <label>
              启用行号
              <input
                type="checkbox"
                checked={lnEnabled}
                onChange={(e) => {
                  setLnEdited(true);
                  setLnEnabled(e.target.checked);
                }}
              />
            </label>
            <label>
              每几行一个号
              <input
                type="number"
                min="1"
                step="1"
                value={lnCountBy}
                disabled={!lnEnabled}
                onChange={(e) => {
                  setLnEdited(true);
                  setLnCountBy(Math.max(1, Math.round(Number(e.target.value) || 1)));
                }}
              />
            </label>
            <label>
              起始编号
              <input
                type="number"
                min="1"
                step="1"
                value={lnStart}
                disabled={!lnEnabled}
                onChange={(e) => {
                  setLnEdited(true);
                  setLnStart(Math.max(1, Math.round(Number(e.target.value) || 1)));
                }}
              />
            </label>
            <label>
              距正文 (cm)
              <input
                type="number"
                min="0"
                step="0.05"
                value={lnDistanceCm}
                disabled={!lnEnabled}
                onChange={(e) => {
                  setLnEdited(true);
                  setLnDistanceCm(Math.max(0, Number(e.target.value) || 0));
                }}
              />
            </label>
          </div>
          <div className="ps-section-title">水印</div>
          <div className="ps-grid">
            <label>
              启用水印
              <input
                type="checkbox"
                checked={wm.enabled}
                onChange={(e) => patchWm({ enabled: e.target.checked })}
              />
            </label>
            <label>
              半透明
              <input
                type="checkbox"
                checked={wm.semiTransparent}
                disabled={!wm.enabled}
                onChange={(e) => patchWm({ semiTransparent: e.target.checked })}
              />
            </label>
            <label>
              文字
              <input
                type="text"
                value={wm.text}
                disabled={!wm.enabled}
                onChange={(e) => patchWm({ text: e.target.value })}
              />
            </label>
            <label>
              字体
              <input
                type="text"
                value={wm.font}
                disabled={!wm.enabled}
                onChange={(e) => patchWm({ font: e.target.value })}
              />
            </label>
            <label>
              字号 (pt，0 为自动)
              <input
                type="number"
                min="0"
                step="1"
                value={wm.sizePt}
                disabled={!wm.enabled}
                onChange={(e) =>
                  patchWm({ sizePt: Math.max(0, Number(e.target.value) || 0) })
                }
              />
            </label>
            <label>
              颜色
              <input
                type="color"
                value={wm.color}
                disabled={!wm.enabled}
                onChange={(e) => patchWm({ color: e.target.value })}
              />
            </label>
          </div>
          <div className="ps-row">
            <label>版式</label>
            <div className="ps-seg">
              <button
                type="button"
                className={wm.layout === "diagonal" ? "active" : ""}
                disabled={!wm.enabled}
                onClick={() => patchWm({ layout: "diagonal" })}
              >
                对角
              </button>
              <button
                type="button"
                className={wm.layout === "horizontal" ? "active" : ""}
                disabled={!wm.enabled}
                onClick={() => patchWm({ layout: "horizontal" })}
              >
                横向
              </button>
            </div>
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
