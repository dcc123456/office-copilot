/**
 * 幻灯片视图（`docs/office-parity-spec.md` 阶段 3）：按页列出形状与文字，单条 run 的段落可就地改文字。
 *
 * 这里只做装配：草稿哪些能落地由 `slideDeckModel.ts` 说一不二，段落里有几条 run 是 `@office-copilot/pptx-io`
 * 读模型给的事实，改动一律收成 `PptxEditRequest` 交给主进程现做 op（组件不碰幻灯片原文，也不碰 Node 能力，
 * `AGENTS.md` §3.3 / §5.2）。
 * 一段里有多条 run 时**不许**整段覆盖——那会把粗体/字号这些局部格式拍平（§2.3），那种段落只读。
 */
import { useCallback, useMemo, useState } from "react";
import type { PresentationModel, SlideEditOp, SlideShape } from "@office-copilot/pptx-io";
import { toast } from "../components/ui/Toast.js";
import { draftKeyOf, pendingTextEdits } from "./slideDeckModel.js";
import "../styles/slides.css";

export interface SlideDeckProps {
  /** 主进程读出来的演示文稿模型；保存成功后换成主进程重读磁盘算出的那份。 */
  deck: PresentationModel;
  /** 当前文档路径（幻灯片还不支持新建，所以进来时一定有路径）。 */
  path: string;
  /** 用主进程送回的新模型刷新视图，避免再发一次读请求。 */
  onDeck: (deck: PresentationModel) => void;
}

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

export function SlideDeck({ deck, path, onDeck }: SlideDeckProps) {
  /** 未保存的段落草稿：键见 `draftKeyOf`，值是用户改后的文字。 */
  const [drafts, setDrafts] = useState<Record<string, string>>({});
  /** 上一次落地那批 op 的前像；送回复用同一条撤销通道（§2.5）。 */
  const [lastOps, setLastOps] = useState<SlideEditOp[]>([]);
  const [isBusy, setIsBusy] = useState(false);

  /** 待保存的那批改动：落地范围与顺序由模型层决定，视图这里只拿来显示条数和送主进程。 */
  const pendingEdits = useMemo(() => pendingTextEdits(deck, drafts), [deck, drafts]);

  const saveEdits = useCallback(async () => {
    if (pendingEdits.length === 0) return;
    setIsBusy(true);
    const saved = await window.officeCopilot.editPptx({ path, edits: pendingEdits });
    setIsBusy(false);
    if (!saved.ok || !saved.presentation) {
      if (saved.message) toast(saved.message, "error");
      return;
    }
    setDrafts({});
    setLastOps(saved.ops ?? []);
    onDeck(saved.presentation);
    toast(`已保存，动了 ${(saved.changedParts ?? []).length} 张幻灯片`);
  }, [onDeck, path, pendingEdits]);

  const revertLastBatch = useCallback(async () => {
    if (lastOps.length === 0) return;
    setIsBusy(true);
    // 逆序送回：后落地的那条先写回前像，别的形状与幻灯片不受牵连（§2.5 内容寻址）。
    const reverted = await window.officeCopilot.revertPptx({ path, ops: [...lastOps].reverse() });
    setIsBusy(false);
    if (!reverted.ok || !reverted.presentation) {
      if (reverted.message) toast(reverted.message, "error");
      return;
    }
    setLastOps([]);
    setDrafts({});
    onDeck(reverted.presentation);
    toast("已撤销上一次保存的那批改动");
  }, [lastOps, onDeck, path]);

  return (
    <div className="slide-deck">
      <div className="slide-deck-bar">
        <span className="slide-deck-count">
          {deck.slides.length} 页 · {pendingEdits.length} 处待保存
        </span>
        <button type="button" onClick={() => void saveEdits()} disabled={isBusy || pendingEdits.length === 0}>
          保存
        </button>
        <button type="button" onClick={() => void revertLastBatch()} disabled={isBusy || lastOps.length === 0}>
          撤销上一次保存
        </button>
        {pendingEdits.length > 0 ? (
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
            </h3>
            {slide.shapes.length === 0 ? <p className="slide-para-locked">这张幻灯片没有可读出的形状。</p> : null}
            {slide.shapes.map((shape) => (
              <div className="slide-shape" key={`${slide.partPath}#${shape.shapeId}`}>
                <div className="slide-shape-meta">
                  <strong>{shape.name || `形状 ${shape.shapeId}`}</strong>
                  <span className="slide-shape-kind">{shapeKindLabel(shape)}</span>
                </div>
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
                  return (
                    <input
                      className="slide-para-input"
                      key={key}
                      type="text"
                      value={drafts[key] ?? paragraphText}
                      onChange={(event) => setDrafts((current) => ({ ...current, [key]: event.target.value }))}
                      spellCheck={false}
                    />
                  );
                })}
              </div>
            ))}
          </section>
        ))}
      </div>
    </div>
  );
}
