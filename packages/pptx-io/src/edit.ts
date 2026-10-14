/**
 * `.pptx` 结构化编辑 op：改某个形状某段某 run 的文字，产出可撤销的前/后像。
 *
 * 与 `sheet-io` 的 `runSheetOps`/`revertSheetOps` 同一套契约（`AGENTS.md` §2.8、§3 不许第二套撤销栈）：
 * 引擎只**产出提案**（带 `beforeXml`/`afterXml` 的 op），落地与撤销由上层在用户确认后调用 `apply`/`revert`。
 * op 存的是**整张幻灯片部件的前像与后像**——撤销按内容寻址（写回前像），不受后来对其他幻灯片的编辑牵连（§2.5）。
 *
 * 为什么只改到 run 粒度（§2.3 不许拍平局部格式）：一段里可能有多个带不同 `rPr`（粗体、字号、颜色）的 run，
 * 整段替换会把这些格式压平。这里只替换目标 run 的 `<a:t>` 文本，其它 run 与 `rPr` 一个字节都不动。
 * 实测：`serialize` 对本机真实幻灯片的往返**逐字一致**（仅声明后换行差异），所以重序列化不会误伤同段其它格式。
 */
import { XML_DECLARATION, children, firstChild, findChildren, isText, parseDocument, serialize, type Element } from "@office-copilot/opc";
import { PptxFile } from "./pptx-file.js";

/** 一次幻灯片文字编辑的前/后像（部件粒度）。 */
export interface SlideTextEditOp {
  /** 幻灯片部件路径，如 `ppt/slides/slide1.xml`。 */
  partPath: string;
  /** 编辑前该部件的完整 XML 文本（撤销就写回它）。 */
  beforeXml: string;
  /** 编辑后该部件的完整 XML 文本（落地就写回它）。 */
  afterXml: string;
}

export type SlideEditResult = { ok: true; op: SlideTextEditOp } | { ok: false; reason: string };

/** 目标 run 的定位坐标。 */
export interface ShapeRunTarget {
  /** 幻灯片部件路径。 */
  partPath: string;
  /** 形状 `<p:cNvPr @id>`。 */
  shapeId: number;
  /** 段落序号（`<p:txBody>` 内 `<a:p>`，0 基）。 */
  paraIndex: number;
  /** run 序号（该段内 `<a:r>`，0 基）。 */
  runIndex: number;
}

const SHAPE_HOST_TAGS = ["p:sp", "p:pic", "p:graphicFrame", "p:grpSp"];

/**
 * 产出一条"把指定 run 的文字改成 newText"的编辑提案，**不落地**。
 *
 * @param file 已打开的读写器。
 * @param target run 定位坐标。
 * @param newText 新的纯文本（会被 XML 转义后写入 `<a:t>`）。
 * @returns `{ ok, op }`：定位失败（部件缺失、形状/段落/run 越界、该 run 无可改文本节点）时 `{ ok:false, reason }`。
 */
export async function setShapeRunText(file: PptxFile, target: ShapeRunTarget, newText: string): Promise<SlideEditResult> {
  const beforeXml = await file.partText(target.partPath);
  if (beforeXml === undefined) return { ok: false, reason: `幻灯片部件不存在：${target.partPath}` };
  let root: Element;
  try {
    root = parseDocument(beforeXml);
  } catch {
    return { ok: false, reason: "幻灯片部件解析失败。" };
  }
  const shape = findShapeById(root, target.shapeId);
  if (!shape) return { ok: false, reason: `找不到形状 id=${target.shapeId}` };
  const txBody = firstChild(shape, "p:txBody");
  if (!txBody) return { ok: false, reason: `形状 id=${target.shapeId} 没有文本体（图片/图形框不可改文字）` };
  const paras = findChildren(txBody, "a:p");
  const para = paras[target.paraIndex];
  if (!para) return { ok: false, reason: `段落序号越界：${target.paraIndex}（共 ${paras.length} 段）` };
  const runs = findChildren(para, "a:r");
  const run = runs[target.runIndex];
  if (!run) return { ok: false, reason: `run 序号越界：${target.runIndex}（该段有 ${runs.length} 个 run）` };
  const aT = firstChild(run, "a:t");
  if (!aT) return { ok: false, reason: "该 run 没有 <a:t> 文本节点。" };
  // 只接受"单一文本子节点"的 <a:t>：里面还嵌了元素（罕见的富文本）时拒绝，改整段会拍平格式（§2.3）。
  if (children(aT).length !== 0) return { ok: false, reason: "该 run 的 <a:t> 含嵌套元素，形状不受支持，拒绝整段替换以免拍平格式。" };
  const textNode = (aT.children ?? []).find(isText);
  if (!textNode) return { ok: false, reason: "该 run 的 <a:t> 没有可改的文本节点。" };
  textNode.text = newText;
  const afterXml = `${XML_DECLARATION}${serialize(root)}`;
  return { ok: true, op: { partPath: target.partPath, beforeXml, afterXml } };
}

/**
 * 落地一条编辑提案：把后像写回部件。
 *
 * @param file 读写器。
 * @param op 由 `setShapeRunText` 产出的 op。
 */
export function applySlideEdit(file: PptxFile, op: SlideTextEditOp): void {
  file.writePartText(op.partPath, op.afterXml);
}

/**
 * 撤销一条编辑提案：把前像写回部件（内容寻址，不受其它部件后续编辑牵连）。
 *
 * @param file 读写器。
 * @param op 之前已 `apply` 的 op。
 */
export function revertSlideEdit(file: PptxFile, op: SlideTextEditOp): void {
  file.writePartText(op.partPath, op.beforeXml);
}

/**
 * 在幻灯片根元素里按 `<p:cNvPr @id>` 找形状元素。
 *
 * @param root 幻灯片根（`<p:sld>`）。
 * @param shapeId 形状号。
 * @returns 形状元素；找不到返回 undefined。
 */
function findShapeById(root: Element, shapeId: number): Element | undefined {
  const cSld = firstChild(root, "p:cSld");
  const spTree = cSld ? firstChild(cSld, "p:spTree") : undefined;
  if (!spTree) return undefined;
  for (const el of children(spTree)) {
    if (!SHAPE_HOST_TAGS.includes(el.name)) continue;
    const host = firstChild(el, "p:nvSpPr") ?? firstChild(el, "p:nvPicPr") ?? firstChild(el, "p:nvGraphicFramePr") ?? firstChild(el, "p:nvGrpSpPr");
    const cNvPr = host ? firstChild(host, "p:cNvPr") : undefined;
    if (cNvPr && Number(cNvPr.attributes["id"]) === shapeId) return el;
  }
  return undefined;
}
