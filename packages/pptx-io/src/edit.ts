/**
 * `.pptx` 结构化编辑 op：改某个形状某段某 run 的文字或字体，产出可撤销的前/后像。
 *
 * 与 `sheet-io` 的 `runSheetOps`/`revertSheetOps` 同一套契约（`AGENTS.md` §2.8、§3 不许第二套撤销栈）：
 * 引擎只**产出提案**（带 `beforeXml`/`afterXml` 的 op），落地与撤销由上层在用户确认后调用 `apply`/`revert`。
 * op 存的是**整张幻灯片部件的前像与后像**——撤销按内容寻址（写回前像），不受后来对其他幻灯片的编辑牵连（§2.5）。
 *
 * 为什么只改到 run 粒度（§2.3 不许拍平局部格式）：一段里可能有多个带不同 `rPr`（粗体、字号、颜色）的 run，
 * 整段替换会把这些格式压平。这里只碰目标 run 自己的 `<a:t>` 文本或 `<a:rPr>` 属性，其它 run 一个字节都不动。
 * 实测：`serialize` 对本机真实幻灯片的往返**逐字一致**（仅声明后换行差异），所以重序列化不会误伤同段其它格式。
 */
import { XML_DECLARATION, children, firstChild, findChildren, isText, parseDocument, serialize, type Element } from "@office-copilot/opc";
import { PptxFile } from "./pptx-file.js";

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
  /** 字号（CSS pt，非半磅）；写进 `sz` 时换算成 pt×100（§2.1，这里是唯一换算点）。 */
  sizePt?: number;
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
  let rPr = firstChild(located.run, "a:rPr");
  if (!rPr) {
    // OOXML：`<a:rPr>` 必须是 `<a:r>` 的第一个子元素，放在 `<a:t>` 之前。
    rPr = newRunPrElement();
    (located.run.children ??= []).unshift(rPr);
  }
  if (format.bold !== undefined) rPr.attributes["b"] = format.bold ? "1" : "0";
  if (format.sizePt !== undefined) rPr.attributes["sz"] = String(Math.round(format.sizePt * 100)); // pt → 百分磅，唯一换算点（§2.1）
  return { ok: true, op: toOp(target.partPath, located.beforeXml, located.root) };
}

/**
 * 落地一条编辑提案：把后像写回部件。
 *
 * @param file 读写器。
 * @param op 由 `setShapeRunText`/`setRunFont` 产出的 op。
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

/** `locateRun` 的成功结果：部件原文、解析好的根与目标 run 元素。 */
type LocatedRun = { ok: true; beforeXml: string; root: Element; run: Element } | { ok: false; reason: string };

/**
 * 把文字与字体 op 共用的定位逻辑收成一份：读部件、解析、按 形状→文本体→段→run 逐层定位。
 *
 * @param file 读写器。
 * @param target run 定位坐标。
 * @returns 成功带 `{ beforeXml, root, run }`；任一层缺失或越界带 `{ ok:false, reason }`（图片/图形框无文本体也走这里）。
 */
async function locateRun(file: PptxFile, target: ShapeRunTarget): Promise<LocatedRun> {
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
  return { ok: true, beforeXml, root, run };
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
 * 造一个空的 `<a:rPr/>` 元素节点，供没有 rPr 的 run 补上。
 *
 * 复用 opc 的解析器（`@rgrove/parse-xml` 的节点形状），而不是自己 `new` 一个类——保持"只有一份 XML 基建"（§3）。
 *
 * @returns 名字为 `a:rPr`、属性与子元素皆空的元素节点。
 */
function newRunPrElement(): Element {
  const root = parseDocument("<root><a:rPr/></root>");
  return children(root)[0]!;
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
