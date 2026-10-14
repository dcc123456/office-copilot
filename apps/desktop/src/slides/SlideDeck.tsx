/**
 * 幻灯片视图（`docs/office-parity-spec.md` 阶段 3）：按页列出形状，改文字、字号、颜色、形状的位置尺寸与叠放次序，
 * 并在这一页加一个文本框、删掉一个形状。
 *
 * 这里只做装配：草稿哪些能落地由 `slideDeckModel.ts` 说一不二，段落里有几条 run、形状有没有自己的
 * `<a:xfrm>` 都是 `@office-copilot/pptx-io` 读模型给的事实，改动一律收成 `PptxEditRequest` 交给主进程现做 op
 * （组件不碰幻灯片原文，也不碰 Node 能力，`AGENTS.md` §3.3 / §5.2）。
 * 两条守卫在模型层：一段里有多条 run 时**不许**整段覆盖（会拍平局部格式，§2.3）；几何继承自版式时**不写死**（D3）。
 * 单位一律 CSS pt / px（读侧已按 `pptx-io/units.ts` 换好），视图里不再乘 100 或 9525（§2.1）。
 */
import { useCallback, useMemo, useState } from "react";
import type { PptxEditRequest, PresentationModel, ShapeBoxPx, SlideLayerDirection, SlideShape } from "@office-copilot/pptx-io";
import type { PptxEditOutcome } from "../global.js";
import { toast } from "../components/ui/Toast.js";
import {
  boxDraftKeyOf,
  draftKeyOf,
  dropShapeDrafts,
  LAYER_DIRECTION_LABEL,
  layerEdgeNote,
  pendingEdits,
  SHAPE_BOX_FIELD_LABEL,
  type SlideDraft,
} from "./slideDeckModel.js";
import "../styles/slides.css";

export interface SlideDeckProps {
  /** 主进程读出来的演示文稿模型；保存成功后换成主进程重读磁盘算出的那份。 */
  deck: PresentationModel;
  /** 上一次落地的那批改动还属于当前这份文件，撤销按钮才给点（前像只对那份文件有效）。 */
  canRevert: boolean;
  /** 落地这一批改动：与 AI 提案共用 `App.tsx` 那一个写盘口，主进程回包里带刷新后的模型。 */
  onCommit: (edits: PptxEditRequest[], summary: string) => Promise<PptxEditOutcome | null>;
  /** 撤销上一次落地的那批改动（逆序送回前像），回包同样带刷新后的模型。 */
  onRevert: () => Promise<PptxEditOutcome | null>;
}

/** 颜色选择框在"这一截字的颜色继承自版式"时的显示值（不代表现场真有这个色）。 */
const INHERITED_COLOR = "#000000";

/**
 * 形状在列表上的身份说明。
 *
 * @param shape 一个形状快照。
 * @returns 一句中文：占位符带它的版式角色，其余按大类点名。
 */
function shapeKindLabel(shape: SlideShape): string {
  if (shape.isPlaceholder) return `占位符 ${shape.placeholderRef ?? "?"}`;
  if (shape.kind === "picture") return "图片";
  if (shape.kind === "graphicFrame") return "图形框（表格 / 图表）";
  if (shape.kind === "group") return "组合形状";
  return shape.paragraphTexts.length ? "文本框" : "形状（无文字）";
}

/**
 * 多条 run 的段落为什么只能看不能改。
 *
 * @param runCount 该段的 `<a:r>` 条数。
 * @returns 一句中文理由；`runCount` 为 0 时说明这段的文字来自字段（页码、日期）而非 run。
 */
function lockedParagraphHint(runCount: number): string {
  if (runCount === 0) return "这段的文字来自字段（页码 / 日期等），不在可改范围";
  return `这段有 ${runCount} 截不同格式的文字，整段覆盖会拍平粗体与字号，本视图只读`;
}

/**
 * 数字输入框的取值。
 *
 * @param raw 输入框原始串。
 * @returns 解析出的数值；空串或非法数字回 undefined，含义是"这条不交出"（不是 0，也不是"改回继承值"）。
 */
function numberOrUndefined(raw: string): number | undefined {
  const trimmed = raw.trim();
  const value = Number(trimmed);
  return trimmed === "" || !Number.isFinite(value) ? undefined : value;
}

/**
 * 主进程的回包折成一句失败理由（三个写盘入口共用，§3.2 不写第三份）。
 *
 * @param outcome 跨界写盘的回包；`null` 表示那份演示文稿已经不在了。
 * @param fallback 回包没带中文理由时用的那句。
 * @returns 失败时要 toast 的那句；成功（回包带刷新后的模型）返回 null。
 */
function failureNote(outcome: PptxEditOutcome | null, fallback: string): string | null {
  if (outcome?.ok && outcome.presentation) return null;
  return outcome?.message ?? fallback;
}

/**
 * 形状在列表与改动说明里的称呼。
 *
 * @param shape 一个形状快照。
 * @returns `<p:cNvPr @name>`；名字是空串时用形状号兜底，别让改动说明里出现「」。
 */
function shapeTitle(shape: SlideShape): string {
  return shape.name || `形状 ${shape.shapeId}`;
}

/** 叠放那四个按钮的排列顺序（照 `LAYER_DIRECTION_LABEL` 的键序：相邻两层先、两端后）。 */
const LAYER_DIRECTIONS = Object.keys(LAYER_DIRECTION_LABEL) as SlideLayerDirection[];

/**
 * 新建文本框的缺省位置与尺寸（CSS px @96DPI）。
 *
 * 本视图没有画布，点不出坐标，所以给一个固定的默认框，用户要挪位置就用行里那四个数字输入框
 * （几何 op 已在）。这里只出现 px，不出现 EMU/914400 —— 换算仍在 `pptx-io/units.ts` 那一处（§2.1）。
 */
const NEW_TEXT_BOX_PX: ShapeBoxPx = { leftPx: 96, topPx: 96, widthPx: 480, heightPx: 144 };

/**
 * 新建文本框里的第一句文字。
 *
 * 空文本框在 OOXML 里是 `<a:p/>`，读模型给出 0 条 run，那一行就成了只读行（多 run 守卫走的是同一条），
 * 用户加了框却打不了字。所以缺省给一句可改的文字，文案归渲染层组装（§5.7）。
 */
const NEW_TEXT_BOX_TEXT = "新文本框";

export function SlideDeck({ deck, canRevert, onCommit, onRevert }: SlideDeckProps) {
  /** 未保存的草稿：run 级字段键见 `draftKeyOf`，几何键见 `boxDraftKeyOf`。 */
  const [drafts, setDrafts] = useState<Record<string, SlideDraft>>({});
  const [isBusy, setIsBusy] = useState(false);

  /** 把一处 run 级草稿字段并进去（同一定位键上文字/字号/颜色各存各的，互不覆盖）。 */
  const patchDraft = useCallback((key: string, patch: SlideDraft) => {
    setDrafts((current) => ({ ...current, [key]: { ...current[key], ...patch } }));
  }, []);

  /**
   * 改几何草稿里的单个字段。
   *
   * @param key 该形状的几何键。
   * @param field `leftPx` / `topPx` / `widthPx` / `heightPx` 之一。
   * @param raw 输入框原始串；清空会把这一项从草稿里删掉（= 不写这个字段）。
   */
  const patchBox = useCallback((key: string, field: keyof ShapeBoxPx, raw: string) => {
    const value = numberOrUndefined(raw);
    setDrafts((current) => {
      const box = { ...current[key]?.box };
      if (value === undefined) delete box[field];
      else box[field] = value;
      return { ...current, [key]: { ...current[key], box } };
    });
  }, []);

  /** 待保存的那批改动：落地范围与顺序由模型层决定，视图这里只拿来显示条数和送主进程。 */
  const edits = useMemo(() => pendingEdits(deck, drafts), [deck, drafts]);

  const saveEdits = useCallback(async () => {
    if (edits.length === 0) return;
    setIsBusy(true);
    const saved = await onCommit(edits, `手工改了 ${edits.length} 处`);
    setIsBusy(false);
    const note = failureNote(saved, "这份演示文稿已经不在了，改动没有写入。");
    if (note) {
      toast(note, "error");
      return;
    }
    setDrafts({});
    toast(`已保存，动了 ${(saved?.changedParts ?? []).length} 张幻灯片`);
  }, [edits, onCommit]);

  /**
   * 撤销上一次落地的那批改动。
   *
   * **不清草稿**：撤销只回已落盘的那批，用户还没保存的输入是他的另一件事（点一次层级按钮就会
   * 产生一个已落盘的批次，撤销是常事，牵连未保存的输入等于吞字）。草稿存的是绝对值，
   * 回滚后 `pendingEdits` 会对着新现场重算，恰好等于现场值的那条自己就不交了。
   */
  const revertLastBatch = useCallback(async () => {
    setIsBusy(true);
    const reverted = await onRevert();
    setIsBusy(false);
    const note = failureNote(reverted, "没有可撤销的批次，或这份文件已经换了。");
    if (note) {
      toast(note, "error");
      return;
    }
    toast("已撤销上一次保存的那批改动");
  }, [onRevert]);

  /**
   * 立刻落地一条**结构性**改动（叠放 / 加文本框 / 删形状）——这几条不进草稿。
   *
   * 理由是它们描述的都不是"某个字段的绝对值"：`up` / `down` 相对的是当前顺序，删掉的形状不会再出现，
   * 攒进草稿点第二次就退化成一次（PowerPoint 功能区同样即点即动）。落地后主进程回包里带重读磁盘的模型，
   * 形状列表按新的文档顺序重排；未保存的文字/几何草稿仍留在 `drafts` 里，不受牵连。
   *
   * @param edit 那一条改动。
   * @param summary 写进那份 modlog 的改动说明（文案在这里组装，§5.7）。
   * @param failFallback 回包没带中文理由时的那句失败文案。
   * @param doneToast 成功后的那句提示。
   */
  const commitNow = useCallback(
    async (edit: PptxEditRequest, summary: string, failFallback: string, doneToast: string): Promise<boolean> => {
      setIsBusy(true);
      const done = await onCommit([edit], summary);
      setIsBusy(false);
      const note = failureNote(done, failFallback);
      if (note) {
        toast(note, "error");
        return false;
      }
      toast(doneToast);
      return true;
    },
    [onCommit],
  );

  /**
   * @param partPath 该页的幻灯片部件路径。
   * @param shape 要挪的那个形状。
   * @param direction 四个方向之一；到头的那些按钮是禁用的，不会走到引擎整批拒绝那条路（§2.4）。
   */
  const commitLayer = useCallback(
    (partPath: string, shape: SlideShape, direction: SlideLayerDirection) =>
      void commitNow(
        { kind: "layer", target: { partPath, shapeId: shape.shapeId }, direction },
        `把「${shapeTitle(shape)}」${LAYER_DIRECTION_LABEL[direction]}`,
        "这份演示文稿已经不在了，叠放次序没有改动。",
        `已把「${shapeTitle(shape)}」${LAYER_DIRECTION_LABEL[direction]}`,
      ),
    [commitNow],
  );

  /**
   * 在这一页末尾加一个文本框（挂在形状子序列末尾 = 最顶层）。
   *
   * @param partPath 该页的幻灯片部件路径。
   * @param slideOrder 该页序号（0 基，只用于改动说明里那句"第几页"）。
   */
  const addTextBox = useCallback(
    (partPath: string, slideOrder: number) =>
      void commitNow(
        { kind: "add_text_box", target: { partPath }, box: NEW_TEXT_BOX_PX, text: NEW_TEXT_BOX_TEXT },
        `在第 ${slideOrder + 1} 页加一个文本框`,
        "这份演示文稿已经不在了，文本框没有加上。",
        "已加一个文本框，改完文字点保存",
      ),
    [commitNow],
  );

  /**
   * 删掉一个形状（整页前像留在历史里，点「撤销上一次保存」能连它一起回来，§2.5）。
   *
   * 删成功后把**那个形状**的未保存草稿一并丢掉：新文本框的形状号是"这一页已有 `@id` 最大值 +1"，
   * 删掉的正好是最大号时，下一次新建会把同一个号发给新形状——留着旧草稿就等于把用户给旧框打的字
   * 静默写进新框。丢掉的是这一行还没保存的输入，形状本身在撤销里回得来。
   *
   * @param partPath 该页的幻灯片部件路径。
   * @param shape 要删的那个形状。
   */
  const deleteShape = useCallback(
    async (partPath: string, shape: SlideShape) => {
      const done = await commitNow(
        { kind: "remove_shape", target: { partPath, shapeId: shape.shapeId } },
        `删掉「${shapeTitle(shape)}」`,
        "这份演示文稿已经不在了，形状没有删除。",
        `已删掉「${shapeTitle(shape)}」，要回退点「撤销上一次保存」`,
      );
      if (done) setDrafts((current) => dropShapeDrafts(current, partPath, shape.shapeId));
    },
    [commitNow],
  );

  return (
    <div className="slide-deck">
      <div className="slide-deck-bar">
        <span className="slide-deck-count">
          {deck.slides.length} 页 · {edits.length} 处待保存
        </span>
        <button type="button" onClick={() => void saveEdits()} disabled={isBusy || edits.length === 0}>
          保存
        </button>
        <button type="button" onClick={() => void revertLastBatch()} disabled={isBusy || !canRevert}>
          撤销上一次保存
        </button>
        {edits.length > 0 ? (
          <button type="button" onClick={() => setDrafts({})} disabled={isBusy}>
            放弃改动
          </button>
        ) : null}
      </div>
      <div className="slide-deck-list">
        {deck.slides.map((slide) => (
          <section className="slide-card" key={slide.partPath}>
            <h3 className="slide-card-title">
              第 {slide.order + 1} 页
              <span className="slide-card-part">{slide.partPath}</span>
              <button
                type="button"
                className="slide-add-textbox"
                onClick={() => addTextBox(slide.partPath, slide.order)}
                disabled={isBusy}
              >
                加文本框
              </button>
            </h3>
            {slide.shapes.length === 0 ? <p className="slide-para-locked">这张幻灯片没有可读出的形状。</p> : null}
            {slide.shapes.map((shape) => {
              const boxKey = boxDraftKeyOf(slide.partPath, shape.shapeId);
              return (
                <div className="slide-shape" key={`${slide.partPath}#${shape.shapeId}`}>
                  <div className="slide-shape-meta">
                    <strong>{shapeTitle(shape)}</strong>
                    <span className="slide-shape-kind">{shapeKindLabel(shape)}</span>
                    <span className="slide-layer-buttons">
                      {LAYER_DIRECTIONS.map((direction) => {
                        const note = layerEdgeNote(slide.shapes, shape, direction);
                        return (
                          <button
                            type="button"
                            key={direction}
                            onClick={() => void commitLayer(slide.partPath, shape, direction)}
                            disabled={isBusy || note !== null}
                            title={note ?? LAYER_DIRECTION_LABEL[direction]}
                          >
                            {LAYER_DIRECTION_LABEL[direction]}
                          </button>
                        );
                      })}
                    </span>
                    <button
                      type="button"
                      className="slide-delete-shape"
                      onClick={() => void deleteShape(slide.partPath, shape)}
                      disabled={isBusy}
                      title={`删掉「${shapeTitle(shape)}」，之后可点「撤销上一次保存」回退`}
                    >
                      删形状
                    </button>
                  </div>
                  {shape.boxPx ? (
                    <div className="slide-box-fields">
                      {(["leftPx", "topPx", "widthPx", "heightPx"] as (keyof ShapeBoxPx)[]).map((field) => (
                        <label className="slide-field" key={field}>
                          <span>{SHAPE_BOX_FIELD_LABEL[field]}</span>
                          <input
                            type="number"
                            value={drafts[boxKey]?.box?.[field] ?? shape.boxPx?.[field] ?? ""}
                            onChange={(event) => patchBox(boxKey, field, event.target.value)}
                          />
                        </label>
                      ))}
                    </div>
                  ) : (
                    <p className="slide-box-note">位置与尺寸继承自版式，本视图不写死（写死就切断母版联动，spec D3）。</p>
                  )}
                  {shape.paragraphTexts.map((paragraphText, paraIndex) => {
                    const runCount = shape.paragraphRunCounts[paraIndex] ?? 0;
                    const key = draftKeyOf(slide.partPath, shape.shapeId, paraIndex);
                    if (runCount !== 1) {
                      return (
                        <div className="slide-para-locked" key={key}>
                          <span>{paragraphText || "（空段）"}</span>
                          <em>{lockedParagraphHint(runCount)}</em>
                        </div>
                      );
                    }
                    const draft = drafts[key];
                    const currentFormat = shape.paragraphRunFormats[paraIndex]?.[0] ?? {};
                    return (
                      <div className="slide-para-row" key={key}>
                        <input
                          className="slide-para-input"
                          type="text"
                          value={draft?.text ?? paragraphText}
                          onChange={(event) => patchDraft(key, { text: event.target.value })}
                          spellCheck={false}
                        />
                        <label className="slide-field">
                          <span>字号</span>
                          <input
                            type="number"
                            min={1}
                            step={0.5}
                            placeholder={currentFormat.sizePt === undefined ? "继承" : ""}
                            value={draft?.sizePt ?? currentFormat.sizePt ?? ""}
                            onChange={(event) => patchDraft(key, { sizePt: numberOrUndefined(event.target.value) })}
                          />
                        </label>
                        <label className="slide-field">
                          <span>颜色</span>
                          <input
                            type="color"
                            title={currentFormat.hex ?? "当前颜色继承自版式"}
                            value={draft?.hex ?? currentFormat.hex ?? INHERITED_COLOR}
                            onChange={(event) => patchDraft(key, { hex: event.target.value })}
                          />
                        </label>
                      </div>
                    );
                  })}
                </div>
              );
            })}
          </section>
        ))}
      </div>
    </div>
  );
}
