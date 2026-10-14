/**
 * `.pptx` 结构化编辑 op：改某个形状某段某 run 的文字、字体、颜色，改形状的位置/尺寸与叠放次序，产出可撤销的前/后像。
 *
 * 与 `sheet-io` 的 `runSheetOps`/`revertSheetOps` 同一套契约（`AGENTS.md` §2.8、§3 不许第二套撤销栈）：
 * 引擎只**产出提案**（带 `beforeXml`/`afterXml` 的 op），落地与撤销由上层在用户确认后调用 `apply`/`revert`。
 * op 存的是**整张幻灯片部件的前像与后像**——撤销按内容寻址（写回前像），不受后来对其他幻灯片的编辑牵连（§2.5）。
 *
 * 为什么只改到 run 粒度（§2.3 不许拍平局部格式）：一段里可能有多个带不同 `rPr`（粗体、字号、颜色）的 run，
 * 整段替换会把这些格式压平。这里只碰目标 run 自己的 `<a:t>` 文本或 `<a:rPr>` 属性，其它 run 一个字节都不动。
 * 实测：`serialize` 对本机真实幻灯片的往返**逐字一致**（仅声明后换行差异），所以重序列化不会误伤同段其它格式。
 */
import { XML_DECLARATION, children, firstChild, findChildren, isElement, isText, parseDocument, serialize, type Element } from "@office-copilot/opc";
import { PptxFile } from "./pptx-file.js";
import { ptToHalfPoint, pxToEmu, type ShapeGeometryPatch } from "./units.js";

/** 一次幻灯片编辑的前/后像（部件粒度）：文字与字体 op 共用这一份形状。 */
export interface SlideEditOp {
  /** 幻灯片部件路径，如 `ppt/slides/slide1.xml`。 */
  partPath: string;
  /** 编辑前该部件的完整 XML 文本（撤销就写回它）。 */
  beforeXml: string;
  /** 编辑后该部件的完整 XML 文本（落地就写回它）。 */
  afterXml: string;
}

export type SlideEditResult = { ok: true; op: SlideEditOp } | { ok: false; reason: string };

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

/** 目标 run 的字体属性（都是可选，给了哪个改哪个）。 */
export interface RunFontFormat {
  /** 粗体；true 写 `b="1"`，false 写 `b="0"`。 */
  bold?: boolean;
  /** 字号（CSS pt，非半磅）；写进 `sz` 时按 `units.ts` 换成 pt×100（§2.1 的唯一换算点）。 */
  sizePt?: number;
}

const SHAPE_HOST_TAGS = ["p:sp", "p:pic", "p:graphicFrame", "p:grpSp"];

/** 目标形状的定位坐标（几何与叠放两条 op 只到形状粒度，不需要段/run）。 */
export interface ShapeGeometryTarget {
  /** 幻灯片部件路径。 */
  partPath: string;
  /** 形状 `<p:cNvPr @id>`。 */
  shapeId: number;
}

/** 叠放方向：`up`/`down` 相邻挪一层，`top`/`bottom` 直接到两端（PowerPoint 功能区的那四个按钮）。 */
export type SlideLayerDirection = "up" | "down" | "top" | "bottom";

/**
 * 一条要落地的幻灯片改动：`kind` 与本文件那五个 op 生产者一一对应，坐标原样交给它们。
 *
 * 这是**跨界契约**（渲染层 → 主进程）放在引擎侧的理由：两边都按它构造与解析，
 * 写第二份就会有人改一边忘一边（`AGENTS.md` §3.2）。里面只有坐标与值，没有幻灯片原文，
 * 也没有给人看的文案（§5.7：文案在渲染层组装）。
 */
export type PptxEditRequest =
  | { kind: "text"; target: ShapeRunTarget; text: string }
  | { kind: "font"; target: ShapeRunTarget; format: RunFontFormat }
  | { kind: "color"; target: ShapeRunTarget; hex: string }
  | { kind: "geometry"; target: ShapeGeometryTarget; box: ShapeGeometryPatch }
  | { kind: "layer"; target: ShapeGeometryTarget; direction: SlideLayerDirection };

/** 保存请求：只带坐标与值，底包由主进程自己从磁盘重读。 */
export interface PptxEditPayload {
  /** 当前文档路径；null 表示还没落过盘。 */
  path: string | null;
  /** 本次要落地的一批改动，按数组顺序应用。 */
  edits: PptxEditRequest[];
  /** 另存为时对话框用的缺省文件名。 */
  defaultName?: string;
}

/** 撤销请求：把之前落地时送回来的那批 op **逆序**送回即可回滚（§2.5 内容寻址）。 */
export interface PptxRevertPayload {
  /** 当前文档路径。 */
  path: string;
  /** 之前那批改动落地时的 op（带前像）。 */
  ops: SlideEditOp[];
}

/** `<a:rPr>` 的填充组（ECMA-376 Part 1 §21.1.2.x）：同一 run 只能有一个，改色时先把旧填充删净再放新的。 */
const FILL_GROUP_TAGS = ["a:noFill", "a:solidFill", "a:gradFill", "a:blipFill", "a:pattFill", "a:grpFill"];

/**
 * 产出一条"把指定 run 的文字改成 newText"的编辑提案，**不落地**。
 *
 * @param file 已打开的读写器。
 * @param target run 定位坐标。
 * @param newText 新的纯文本（会被 XML 转义后写入 `<a:t>`）。
 * @returns `{ ok, op }`：定位失败（部件缺失、形状/段落/run 越界、该 run 无可改文本节点）时 `{ ok:false, reason }`。
 */
export async function setShapeRunText(file: PptxFile, target: ShapeRunTarget, newText: string): Promise<SlideEditResult> {
  const located = await locateRun(file, target);
  if (!located.ok) return located;
  const aT = firstChild(located.run, "a:t");
  if (!aT) return { ok: false, reason: "该 run 没有 <a:t> 文本节点。" };
  // 只接受"单一文本子节点"的 <a:t>：里面还嵌了元素（罕见的富文本）时拒绝，改整段会拍平格式（§2.3）。
  if (children(aT).length !== 0) return { ok: false, reason: "该 run 的 <a:t> 含嵌套元素，形状不受支持，拒绝整段替换以免拍平格式。" };
  const textNode = (aT.children ?? []).find(isText);
  if (!textNode) return { ok: false, reason: "该 run 的 <a:t> 没有可改的文本节点。" };
  textNode.text = newText;
  return { ok: true, op: toOp(target.partPath, located.beforeXml, located.root) };
}

/**
 * 产出一条"改指定 run 的字体（粗体 / 字号）"的编辑提案，**不落地**。
 *
 * @param file 已打开的读写器。
 * @param target run 定位坐标。
 * @param format 要改的字体属性；`bold` 与 `sizePt` 至少给一个，否则视为无操作拒绝。
 * @returns `{ ok, op }`：定位失败或属性非法时 `{ ok:false, reason }`。rPr 不存在时会新建一个 `<a:rPr/>` 插到 `<a:r>` 首位。
 */
export async function setRunFont(file: PptxFile, target: ShapeRunTarget, format: RunFontFormat): Promise<SlideEditResult> {
  if (format.bold === undefined && format.sizePt === undefined) {
    return { ok: false, reason: "没有要改的字体属性（bold / sizePt 至少给一个）。" };
  }
  if (format.sizePt !== undefined && !(format.sizePt > 0)) {
    return { ok: false, reason: `字号必须是正数 pt：${format.sizePt}` };
  }
  const located = await locateRun(file, target);
  if (!located.ok) return located;
  const rPr = getOrCreateRunPr(located.run);
  if (format.bold !== undefined) rPr.attributes["b"] = format.bold ? "1" : "0";
  if (format.sizePt !== undefined) rPr.attributes["sz"] = String(ptToHalfPoint(format.sizePt));
  return { ok: true, op: toOp(target.partPath, located.beforeXml, located.root) };
}

/**
 * 产出一条"改指定 run 的字体颜色"的编辑提案，**不落地**。
 *
 * @param file 已打开的读写器。
 * @param target run 定位坐标。
 * @param hex 颜色，`#rrggbb`（编辑器内一律 hex，§2.2）；写进 OOXML 时去 `#` 存进 `<a:srgbClr val>`，是这一层唯一的颜色形态转换。
 * @returns `{ ok, op }`：定位失败或颜色非法时 `{ ok:false, reason }`。会替换该 run 已有的填充色，不动其它 run。
 */
export async function setRunColor(file: PptxFile, target: ShapeRunTarget, hex: string): Promise<SlideEditResult> {
  const val = normalizeHexColor(hex);
  if (!val) return { ok: false, reason: `颜色必须是 #rrggbb：${hex}` };
  const located = await locateRun(file, target);
  if (!located.ok) return located;
  const rPr = getOrCreateRunPr(located.run);
  // 替换现有填充：删掉填充组里的其它成员，只留新的 solidFill（同一 run 只能有一个填充）。
  const kept = (rPr.children ?? []).filter((c) => !isElement(c) || !FILL_GROUP_TAGS.includes(c.name));
  const solidFill = newSolidFillElement(val);
  // OOXML rPr 子序：`<a:ln>` 之后紧接填充组。有 ln 就插其后，否则插到最前。
  const lnIndex = kept.findIndex((c) => isElement(c) && c.name === "a:ln");
  kept.splice(lnIndex + 1, 0, solidFill);
  rPr.children = kept;
  return { ok: true, op: toOp(target.partPath, located.beforeXml, located.root) };
}

/**
 * 产出一条"移动 / 缩放指定形状"的编辑提案，**不落地**。
 *
 * @param file 已打开的读写器。
 * @param target 形状定位坐标（部件路径 + 形状号）。
 * @param box 位置/尺寸（CSS px @96DPI）；给了哪个改哪个，left/top 允许负值，width/height 须为正。
 * @returns `{ ok, op }`：无属性、尺寸非法、定位失败或形状是组合（子偏移语义不同，暂不支持）时 `{ ok:false, reason }`。
 *   px→EMU 换算走 `units.ts`，那是全仓唯一的一处（§2.1）。对占位符写显式 `<a:xfrm>` 是用户明示的覆盖，不是读侧烘焙（D3 只禁止读侧把继承几何固化）。
 */
export async function setShapeGeometry(file: PptxFile, target: ShapeGeometryTarget, box: ShapeGeometryPatch): Promise<SlideEditResult> {
  if (box.leftPx === undefined && box.topPx === undefined && box.widthPx === undefined && box.heightPx === undefined) {
    return { ok: false, reason: "没有要改的几何属性（left / top / width / height 至少给一个）。" };
  }
  if (box.widthPx !== undefined && !(box.widthPx > 0)) return { ok: false, reason: `宽度必须是正数 px：${box.widthPx}` };
  if (box.heightPx !== undefined && !(box.heightPx > 0)) return { ok: false, reason: `高度必须是正数 px：${box.heightPx}` };
  const located = await locateShape(file, target);
  if (!located.ok) return located;
  if (located.shape.name === "p:grpSp") return { ok: false, reason: "组合形状的几何暂不支持（子偏移语义不同）。" };
  const spPr = getOrCreateSpPr(located.shape);
  const xfrm = getOrCreateChild(spPr, "a:xfrm", true); // CT_ShapeProperties：xfrm 是第一个子元素
  if (box.leftPx !== undefined || box.topPx !== undefined) {
    const off = getOrCreateChild(xfrm, "a:off", false);
    if (box.leftPx !== undefined) off.attributes["x"] = String(pxToEmu(box.leftPx));
    if (box.topPx !== undefined) off.attributes["y"] = String(pxToEmu(box.topPx));
  }
  if (box.widthPx !== undefined || box.heightPx !== undefined) {
    const ext = getOrCreateChild(xfrm, "a:ext", false);
    if (box.widthPx !== undefined) ext.attributes["cx"] = String(pxToEmu(box.widthPx));
    if (box.heightPx !== undefined) ext.attributes["cy"] = String(pxToEmu(box.heightPx));
  }
  return { ok: true, op: toOp(target.partPath, located.beforeXml, located.root) };
}

/**
 * 产出一条"沿叠放次序移动指定形状"的编辑提案，**不落地**。
 *
 * 叠放次序就是 `<p:cSld><p:spTree>` 里形状子元素的文档顺序：第一个画在最底下、最后一个画在最顶上。
 * 依据（spec 阶段 4 记的一手来源）：ECMA-376 Part 1 §19.3.5 `CT_GroupShape` 规定 spTree 的子元素序列
 * 为 `p:nvGrpSpPr`、`p:grpSpPr`，其后是 `p:sp`/`p:grpSp`/`p:graphicFrame`/`p:cxnSp` 的 choice；
 * 遮挡语义由 python-pptx 的 `src/pptx/shapes/shapetree.py` 写明——
 * "The first shape in the sequence is the backmost in z-order and the last shape is topmost."
 * 所以这里**只重排子元素、不写任何属性**，那两个非形状子元素与节点间的空白留在原位。
 *
 * @param file 已打开的读写器。
 * @param target 形状定位坐标（部件路径 + 形状号）。
 * @param direction `up`/`down` 挪一层，`top`/`bottom` 直接到两端。
 * @returns `{ ok, op }`：这一页没有形状树、找不到形状，或**已经在那一端**（再挪是一次空改动，不值得落盘）时 `{ ok:false, reason }`。
 */
export async function moveShapeLayer(file: PptxFile, target: ShapeGeometryTarget, direction: SlideLayerDirection): Promise<SlideEditResult> {
  const located = await locateShape(file, target);
  if (!located.ok) return located;
  const spTree = shapeTreeOf(located.root);
  if (!spTree) return { ok: false, reason: `这一页没有形状树（<p:spTree>）：${target.partPath}` };
  const hosts = children(spTree).filter((el) => SHAPE_HOST_TAGS.includes(el.name));
  const at = hosts.indexOf(located.shape);
  const order = [...hosts];
  if (direction === "up") {
    if (at === hosts.length - 1) return { ok: false, reason: "这个形状已经在最顶层。" };
    order.splice(at, 1);
    order.splice(at + 1, 0, located.shape);
  } else if (direction === "down") {
    if (at === 0) return { ok: false, reason: "这个形状已经在最底层。" };
    order.splice(at, 1);
    order.splice(at - 1, 0, located.shape);
  } else if (direction === "top") {
    if (at === hosts.length - 1) return { ok: false, reason: "这个形状已经在最顶层。" };
    order.splice(at, 1);
    order.push(located.shape);
  } else {
    if (at === 0) return { ok: false, reason: "这个形状已经在最底层。" };
    order.splice(at, 1);
    order.unshift(located.shape);
  }
  rewriteShapeOrder(spTree, order);
  return { ok: true, op: toOp(target.partPath, located.beforeXml, located.root) };
}

/**
 * 落地一条编辑提案：把后像写回部件。
 *
 * @param file 读写器。
 * @param op 由 `setShapeRunText`/`setRunFont`/`setRunColor`/`setShapeGeometry`/`moveShapeLayer` 产出的 op。
 */
export function applySlideEdit(file: PptxFile, op: SlideEditOp): void {
  file.writePartText(op.partPath, op.afterXml);
}

/**
 * 撤销一条编辑提案：把前像写回部件（内容寻址，不受其它部件后续编辑牵连）。
 *
 * @param file 读写器。
 * @param op 之前已 `apply` 的 op。
 */
export function revertSlideEdit(file: PptxFile, op: SlideEditOp): void {
  file.writePartText(op.partPath, op.beforeXml);
}

/** 一条逆操作没能落地的记录，UI 据此报"部分撤销（x/y）"（与 sheet 侧 `SheetSkippedOp` 同形）。 */
export interface SlideSkippedRevert {
  op: SlideEditOp;
  reason: string;
}

/** 一批撤销的结果。`applied` 为空表示磁盘上不该动（要么全对不上，要么批是空的）。 */
export interface SlideRevertOutcome {
  /** 真正写回前像的那些 op，顺序与传入的一致（调用方按逆序传入）。 */
  applied: SlideEditOp[];
  skipped: SlideSkippedRevert[];
}

/**
 * 一批撤销的现场核对 + 落地：只把"这一页现在正是改完后的那个样子"的那些 op 的前像写回。
 *
 * 内容寻址在这里落到**逐页**判定（`AGENTS.md` §2.5）：一页后来又被手工改过，那一条就跳过并给理由，
 * 绝不能拿旧前像覆盖用户后来的编辑。整批 abort 只发生在"一条都落不了"时——那时磁盘一个字都不动。
 *
 * @param file 读写器（对着磁盘当前那份打开的）。
 * @param ops 之前落地时带回来的那批 op，**按逆序**传入（最后改的先回）。
 * @returns `{ applied, skipped }`：部件已不在包里、或现场与后像逐字不一致的条目进 `skipped`。
 */
export async function revertSlideEdits(file: PptxFile, ops: SlideEditOp[]): Promise<SlideRevertOutcome> {
  const applied: SlideEditOp[] = [];
  const skipped: SlideSkippedRevert[] = [];
  for (const op of ops) {
    const current = await file.partText(op.partPath);
    if (current === undefined) {
      skipped.push({ op, reason: `幻灯片 ${op.partPath} 在文件里已经找不到了。` });
      continue;
    }
    if (current !== op.afterXml) {
      skipped.push({ op, reason: `这一页（${op.partPath}）后来又被改过，回滚会牵连那些编辑。` });
      continue;
    }
    revertSlideEdit(file, op);
    applied.push(op);
  }
  return { applied, skipped };
}

/** `locateShape` 的成功结果：部件原文、解析好的根与目标形状元素。 */
type LocatedShape = { ok: true; beforeXml: string; root: Element; shape: Element } | { ok: false; reason: string };

/** `locateRun` 的成功结果：在形状之上再给出目标 run 元素。 */
type LocatedRun = { ok: true; beforeXml: string; root: Element; run: Element } | { ok: false; reason: string };

/**
 * 读部件、解析并按 `<p:cNvPr @id>` 定位形状元素——文字/字体/颜色/几何所有 op 共用的第一层（§3.2）。
 *
 * @param file 读写器。
 * @param target 部件路径 + 形状号。
 * @returns 成功带 `{ beforeXml, root, shape }`；部件缺失、解析失败或找不到形状带 `{ ok:false, reason }`。
 */
async function locateShape(file: PptxFile, target: ShapeGeometryTarget): Promise<LocatedShape> {
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
  return { ok: true, beforeXml, root, shape };
}

/**
 * 在 `locateShape` 之上再定位到具体 run：形状→文本体→段→run 逐层深入。
 *
 * @param file 读写器。
 * @param target run 定位坐标。
 * @returns 成功带 `{ beforeXml, root, run }`；任一层缺失或越界带 `{ ok:false, reason }`（图片/图形框无文本体也走这里）。
 */
async function locateRun(file: PptxFile, target: ShapeRunTarget): Promise<LocatedRun> {
  const located = await locateShape(file, { partPath: target.partPath, shapeId: target.shapeId });
  if (!located.ok) return located;
  const txBody = firstChild(located.shape, "p:txBody");
  if (!txBody) return { ok: false, reason: `形状 id=${target.shapeId} 没有文本体（图片/图形框不可改文字）` };
  const paras = findChildren(txBody, "a:p");
  const para = paras[target.paraIndex];
  if (!para) return { ok: false, reason: `段落序号越界：${target.paraIndex}（共 ${paras.length} 段）` };
  const runs = findChildren(para, "a:r");
  const run = runs[target.runIndex];
  if (!run) return { ok: false, reason: `run 序号越界：${target.runIndex}（该段有 ${runs.length} 个 run）` };
  return { ok: true, beforeXml: located.beforeXml, root: located.root, run };
}

/**
 * 把改过的根元素序列化回带声明的完整部件文本，装成 op。
 *
 * @param partPath 幻灯片部件路径。
 * @param beforeXml 定位时读到的原文（前像）。
 * @param root 已在内存里改过的根元素。
 * @returns `{ partPath, beforeXml, afterXml }`。
 */
function toOp(partPath: string, beforeXml: string, root: Element): SlideEditOp {
  return { partPath, beforeXml, afterXml: `${XML_DECLARATION}${serialize(root)}` };
}

/**
 * 取该 run 的 `<a:rPr>`，没有就新建一个插到 `<a:r>` 首位（OOXML：`<a:rPr>` 必须是第一个子元素，在 `<a:t>` 前）。
 *
 * 字体与颜色两条 op 都要这一步，收成一份（§3.2）。
 *
 * @param run 已定位的 `<a:r>` 元素。
 * @returns 该 run 现有的或新建的 `<a:rPr>` 元素。
 */
function getOrCreateRunPr(run: Element): Element {
  const existing = firstChild(run, "a:rPr");
  if (existing) return existing;
  const rPr = parseFragmentElement("<root><a:rPr/></root>");
  (run.children ??= []).unshift(rPr);
  return rPr;
}

/**
 * 造一个 `<a:solidFill><a:srgbClr val="RRGGBB"/></a:solidFill>` 元素节点。
 *
 * @param val 六位十六进制（不含 `#`）。
 * @returns 解析好的 `<a:solidFill>` 元素。
 */
function newSolidFillElement(val: string): Element {
  return parseFragmentElement(`<root><a:solidFill><a:srgbClr val="${val}"/></a:solidFill></root>`);
}

/**
 * 取父元素下指定名字的子元素，没有就新建一个。
 *
 * @param parent 父元素。
 * @param name 带前缀的标签名原值（如 `a:xfrm`、`a:off`）。
 * @param asFirst true 插到子列表首位，false 追加到末尾（`a:off` 先于 `a:ext` 创建，顺序天然满足 OOXML）。
 * @returns 现有或新建的子元素。
 */
function getOrCreateChild(parent: Element, name: string, asFirst: boolean): Element {
  const existing = firstChild(parent, name);
  if (existing) return existing;
  const el = parseFragmentElement(`<root><${name}/></root>`);
  const kids = (parent.children ??= []);
  if (asFirst) kids.unshift(el);
  else kids.push(el);
  return el;
}

/**
 * 取形状的 `<p:spPr>`，没有就新建并插到 OOXML 规定的位置（紧跟 `<p:nv*Pr>` 之后）。
 *
 * @param shape 形状元素（`p:sp`/`p:pic`/`p:graphicFrame`）。
 * @returns 现有或新建的 `<p:spPr>`。
 */
function getOrCreateSpPr(shape: Element): Element {
  const existing = firstChild(shape, "p:spPr");
  if (existing) return existing;
  const spPr = parseFragmentElement("<root><p:spPr/></root>");
  const kids = (shape.children ??= []);
  const hostIndex = kids.findIndex((c) => isElement(c) && c.name.startsWith("p:nv"));
  kids.splice(hostIndex + 1, 0, spPr);
  return spPr;
}

/**
 * 用 opc 的解析器从一小段包装 XML 里取出第一个子元素，避免自己 `new` parse-xml 的类（§3 只有一份 XML 基建）。
 *
 * @param wrappedXML 形如 `<root>…</root>` 的片段。
 * @returns `<root>` 的第一个子元素。
 */
function parseFragmentElement(wrappedXML: string): Element {
  return children(parseDocument(wrappedXML))[0]!;
}

/**
 * 校验并归一化编辑器里的 `#rrggbb`（§2.2 颜色一律 hex）。
 *
 * @param hex 用户/模型给的颜色，允许带或不带前导 `#`。
 * @returns 大写六位十六进制（可直接写进 `srgbClr @val`）；不是合法 hex 时返回 undefined。
 */
function normalizeHexColor(hex: string): string | undefined {
  const bare = hex.startsWith("#") ? hex.slice(1) : hex;
  return /^[0-9a-fA-F]{6}$/.test(bare) ? bare.toUpperCase() : undefined;
}

/**
 * 取这一页的形状树 `<p:cSld><p:spTree>`：叠放次序 op 要在它身上重排子元素。
 *
 * @param root 幻灯片根（`<p:sld>`）。
 * @returns 形状树元素；这一页没有 cSld/spTree 时 undefined。
 */
function shapeTreeOf(root: Element): Element | undefined {
  const cSld = firstChild(root, "p:cSld");
  return cSld ? firstChild(cSld, "p:spTree") : undefined;
}

/**
 * 在幻灯片根元素里按 `<p:cNvPr @id>` 找形状元素。
 *
 * @param root 幻灯片根（`<p:sld>`）。
 * @param shapeId 形状号。
 * @returns 形状元素；找不到返回 undefined。
 */
function findShapeById(root: Element, shapeId: number): Element | undefined {
  const spTree = shapeTreeOf(root);
  if (!spTree) return undefined;
  for (const el of children(spTree)) {
    if (!SHAPE_HOST_TAGS.includes(el.name)) continue;
    const host = firstChild(el, "p:nvSpPr") ?? firstChild(el, "p:nvPicPr") ?? firstChild(el, "p:nvGraphicFramePr") ?? firstChild(el, "p:nvGrpSpPr");
    const cNvPr = host ? firstChild(host, "p:cNvPr") : undefined;
    if (cNvPr && Number(cNvPr.attributes["id"]) === shapeId) return el;
  }
  return undefined;
}

/**
 * 把 `<p:spTree>` 里的形状子元素换成给定的顺序。
 *
 * 只填形状那些槽位：`p:nvGrpSpPr`、`p:grpSpPr` 与元素之间的空白 / 注释节点原地不动，
 * 这样重序列化后除叠放顺序外一个字都不差（§2.3 不许误伤别的内容）。
 *
 * @param spTree 形状树元素。
 * @param order 新的形状顺序，长度必须等于现有形状子元素数。
 */
function rewriteShapeOrder(spTree: Element, order: Element[]): void {
  const nodes = spTree.children;
  if (!nodes) return;
  const slots: number[] = [];
  nodes.forEach((node, index) => {
    if (isElement(node) && SHAPE_HOST_TAGS.includes(node.name)) slots.push(index);
  });
  slots.forEach((slot, k) => {
    nodes[slot] = order[k]!;
  });
}
