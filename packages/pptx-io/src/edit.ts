/**
 * `.pptx` 结构化编辑 op：改某个形状某段某 run 的文字、字体、颜色，改形状的位置/尺寸与叠放次序，
 * 在这一页新建文本框、删掉一个形状，把某一页的演讲者备注整篇换掉，全部产出可撤销的前/后像。
 *
 * 与 `sheet-io` 的 `runSheetOps`/`revertSheetOps` 同一套契约（`AGENTS.md` §2.8、§3 不许第二套撤销栈）：
 * 引擎只**产出提案**（带 `beforeXml`/`afterXml` 的 op），落地与撤销由上层在用户确认后调用 `apply`/`revert`。
 * op 存的是**整张幻灯片部件的前像与后像**——撤销按内容寻址（写回前像），不受后来对其他幻灯片的编辑牵连（§2.5）。
 *
 * 为什么只改到 run 粒度（§2.3 不许拍平局部格式）：一段里可能有多个带不同 `rPr`（粗体、字号、颜色）的 run，
 * 整段替换会把这些格式压平。这里只碰目标 run 自己的 `<a:t>` 文本或 `<a:rPr>` 属性，其它 run 一个字节都不动。
 * 实测：`serialize` 对本机真实幻灯片的往返**逐字一致**（仅声明后换行差异），所以重序列化不会误伤同段其它格式。
 */
import { XML_DECLARATION, children, escapeText, firstChild, findChildren, isElement, isText, parseDocument, serialize, type Element, type XmlNode } from "@office-copilot/opc";
import { findNotesBodyShape, shapeIdOf } from "./read.js";
import { PptxFile } from "./pptx-file.js";
import { ptToHalfPoint, pxToEmu, type ShapeBoxPx, type ShapeGeometryPatch } from "./units.js";

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

const SHAPE_HOST_TAGS = ["p:sp", "p:pic", "p:graphicFrame", "p:grpSp", "p:cxnSp"];

/** 目标形状的定位坐标（几何与叠放两条 op 只到形状粒度，不需要段/run）。 */
export interface ShapeGeometryTarget {
  /** 幻灯片部件路径。 */
  partPath: string;
  /** 形状 `<p:cNvPr @id>`。 */
  shapeId: number;
}

/** 叠放方向：`up`/`down` 相邻挪一层，`top`/`bottom` 直接到两端（PowerPoint 功能区的那四个按钮）。 */
export type SlideLayerDirection = "up" | "down" | "top" | "bottom";

/** 整页粒度的坐标：新建形状不属于任何一个已有形状，只属于某一页。 */
export interface SlidePartTarget {
  /** 幻灯片部件路径。 */
  partPath: string;
}

/**
 * 备注的落点坐标：备注部件自己的包内路径（读模型 `Slide.notes.partPath` 给的就是它）。
 *
 * 为什么不用"页号 + 形状号"：备注正文那一格在**另一个部件**里，`SlideEditOp` 是部件粒度的前/后像，
 * 拿备注部件路径当坐标，落地与撤销就完全走现成那一条通道，不需要第二种 op 形状（§3.4 扩展现有接口）。
 */
export interface SlideNotesTarget {
  /** 备注部件路径，如 `ppt/notesSlides/notesSlide1.xml`。 */
  notesPartPath: string;
}

/**
 * 一条要落地的幻灯片改动：`kind` 与本文件那八个 op 生产者一一对应，坐标原样交给它们。
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
  | { kind: "layer"; target: ShapeGeometryTarget; direction: SlideLayerDirection }
  | { kind: "add_text_box"; target: SlidePartTarget; box: ShapeBoxPx; text: string }
  | { kind: "remove_shape"; target: ShapeGeometryTarget }
  | { kind: "set_notes"; target: SlideNotesTarget; paragraphTexts: string[] };

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
 * @returns `{ ok, op }`：无属性、尺寸非法、定位失败、形状是组合（子偏移语义不同，暂不支持）、
 *   或补丁会让 `<a:off>` / `<a:ext>` 缺一半必填属性时 `{ ok:false, reason }`。
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
  // 组里的形状：它 `<a:xfrm>` 里的数是**组内子坐标系**的，要换算回页面位置得连着组的 chOff/chExt 比例与旋转一起算，
  // 这一档不做那个映射，所以也不接受把页面 px 写进去（那等于改了它相对组的位置）。读侧同样不给它的几何数字。
  if (located.container !== shapeTreeOf(located.root)) {
    return { ok: false, reason: "这个形状在一个组合里，几何那一条这一档不改（它写的坐标是组内子坐标系，换算成页面位置要连着组的缩放与旋转一起算）。" };
  }
  const spPr = firstChild(located.shape, "p:spPr");
  const xfrm = spPr ? firstChild(spPr, "a:xfrm") : undefined;
  const missing = [
    ...missingPairValues(xfrm ? firstChild(xfrm, "a:off") : undefined, [
      { attribute: "x", field: "leftPx", patch: box.leftPx },
      { attribute: "y", field: "topPx", patch: box.topPx },
    ]),
    ...missingPairValues(xfrm ? firstChild(xfrm, "a:ext") : undefined, [
      { attribute: "cx", field: "widthPx", patch: box.widthPx },
      { attribute: "cy", field: "heightPx", patch: box.heightPx },
    ]),
  ];
  if (missing.length > 0) {
    return { ok: false, reason: `这个形状没写全自己的几何，而 <a:off> 与 <a:ext> 的必填属性得一次给齐，只写一半会被 PowerPoint 判损坏：还缺 ${missing.join("、")}。` };
  }
  const mutableSpPr = getOrCreateSpPr(located.shape);
  const mutableXfrm = getOrCreateChild(mutableSpPr, "a:xfrm", true); // CT_ShapeProperties：xfrm 是第一个子元素
  if (box.leftPx !== undefined || box.topPx !== undefined) {
    const off = getOrCreateChild(mutableXfrm, "a:off", false);
    if (box.leftPx !== undefined) off.attributes["x"] = String(pxToEmu(box.leftPx));
    if (box.topPx !== undefined) off.attributes["y"] = String(pxToEmu(box.topPx));
  }
  if (box.widthPx !== undefined || box.heightPx !== undefined) {
    const ext = getOrCreateChild(mutableXfrm, "a:ext", false);
    if (box.widthPx !== undefined) ext.attributes["cx"] = String(pxToEmu(box.widthPx));
    if (box.heightPx !== undefined) ext.attributes["cy"] = String(pxToEmu(box.heightPx));
  }
  return { ok: true, op: toOp(target.partPath, located.beforeXml, located.root) };
}

/**
 * 数一数 `<a:off>` / `<a:ext>` 这一对必填属性，补丁给全了没有。
 * 依据：ECMA-376 §20.1.1.15 `CT_Point2D` 与 §20.1.1.14 `CT_PositiveSize2D` 把 x、y（cx、cy）都定成必填，
 * python-pptx 读到只有 `y` 的 `<a:off>` 直接抛 `InvalidXmlError`（本机实测）。不许拿"另一轴从版式继承"去猜，一猜就把用户没要求的那一轴钉死了。
 *
 * @param pair 现有的 `<a:off>` 或 `<a:ext>`，形状没写时是 undefined。
 * @param requirements 该元素的必填属性名、对应字段名与补丁值（没给的是 undefined）三者配对。
 * @returns 现有元素与补丁**都**没覆盖到的字段名；齐了则是空数组。
 */
function missingPairValues(
  pair: Element | undefined,
  requirements: Array<{ attribute: string; field: string; patch: number | undefined }>,
): string[] {
  return requirements
    .filter((entry) => pair?.attributes[entry.attribute] === undefined && entry.patch === undefined)
    .map((entry) => entry.field);
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
 * @returns `{ ok, op }`：找不到形状，或**已经在要去的尽头**（再挪是一次空改动，不值得落盘）时 `{ ok:false, reason }`。
 *   组里的形状只在它自己那一层（那个组）里换序——叠放次序本来就是"同一层内"的性质，PowerPoint 也是这么画的。
 */
export async function moveShapeLayer(file: PptxFile, target: ShapeGeometryTarget, direction: SlideLayerDirection): Promise<SlideEditResult> {
  const located = await locateShape(file, target);
  if (!located.ok) return located;
  const hosts = children(located.container).filter((el) => SHAPE_HOST_TAGS.includes(el.name));
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
  rewriteShapeOrder(located.container, order);
  return { ok: true, op: toOp(target.partPath, located.beforeXml, located.root) };
}

/**
 * 产出一条"在这一页新建一个文本框"的编辑提案，**不落地**。
 *
 * 新建的 `<p:sp>` 逐字照 python-pptx 的 `CT_Shape._textbox_sp_tmpl()`
 * （`tmp/pylibs/pptx/oxml/shapes/autoshape.py:388`，PowerPoint 自己插入文本框时也是这一段：
 * `p:nvSpPr > cNvPr/cNvSpPr txBox="1"/nvPr`，`p:spPr > a:xfrm + a:prstGeom prst="rect" + a:noFill`，
 * `p:txBody > a:bodyPr + a:lstStyle + a:p`），只是把那个空的 `<a:p/>` 换成带一个 run 的段落。
 * 形状号取**这一页部件内**已有 `@id` 的最大值 +1（一手依据：`oxml/shapes/groupshape.py` 的 `max_shape_id`
 * 用的就是 spTree 上的 `xpath("//@id")`，注释写着 "XML id-values have document scope"——这份 document 就是这一页；
 * 实测 `tmp/pptx-edit/probe.pptx` 两页的 id 分别是 `[1,2,3]` 与 `[1,2]`，跨页本来就重复）。
 * 新形状追加在形状子序列末尾，也就是最顶层（叠放语义见 `moveShapeLayer` 的注释）。
 *
 * @param file 已打开的读写器。
 * @param target 整页坐标（部件路径）。
 * @param box 位置/尺寸（CSS px @96DPI，四个字段都要给）；px→EMU 只在 `units.ts` 发生（§2.1）。
 * @param text 文本框里的文字；空串就是空白文本框。文字来自用户/模型输入，写进 XML 前一律转义（§3.6：边界校验只在系统边界做）。
 * @returns `{ ok, op }`：部件缺失 / 解析失败 / 没有形状树 / 宽高非正时 `{ ok:false, reason }`。
 */
export async function addTextBox(file: PptxFile, target: SlidePartTarget, box: ShapeBoxPx, text: string): Promise<SlideEditResult> {
  if (!(box.widthPx > 0)) return { ok: false, reason: `新文本框的宽度必须是正数 px：${box.widthPx}` };
  if (!(box.heightPx > 0)) return { ok: false, reason: `新文本框的高度必须是正数 px：${box.heightPx}` };
  const loaded = await loadSlidePart(file, target.partPath);
  if (!loaded.ok) return loaded;
  const spTree = shapeTreeOf(loaded.root);
  if (!spTree) return { ok: false, reason: `这一页没有形状树（<p:spTree>）：${target.partPath}` };
  const shapeId = maxShapeIdInPart(loaded.root) + 1;
  appendShapeToTree(spTree, parseFragmentElement(textBoxXml(shapeId, box, text)));
  return { ok: true, op: toOp(target.partPath, loaded.beforeXml, loaded.root) };
}

/**
 * 产出一条"删掉指定形状"的编辑提案，**不落地**。
 *
 * 结构上就是把那一元素从 `<p:spTree>` 里摘掉：形状与它自己的关系（图片的 `r:embed`、图表的 `r:id`）
 * 留在关系表里不影响合规性——OPC 只要求**被引用**的关系存在，不要求每条关系都有人引用，
 * 而 `auditPptxPackage` 那五条判定查的是"部件里写到的 `r:id` 必须在关系表里"，删元素只会少引用、不会多引用。
 * 因此这里不动 `[Content_Types].xml`、不动 `.rels`，也就没有 §2.10 那"三处一起登记"要回滚的东西。
 *
 * @param file 已打开的读写器。
 * @param target 形状定位坐标（部件路径 + 形状号）。
 * @returns `{ ok, op }`：定位失败时 `{ ok:false, reason }`。删掉最后一个形状是合法操作（页面变空白），照做；
 *   组里的子形状也真能从组里摘掉（对着一层摘，不对着页面 `spTree` 摘）。
 */
export async function removeShape(file: PptxFile, target: ShapeGeometryTarget): Promise<SlideEditResult> {
  const located = await locateShape(file, target);
  if (!located.ok) return located;
  removeShapeFromTree(located.container, located.shape);
  return { ok: true, op: toOp(target.partPath, located.beforeXml, located.root) };
}

/**
 * 整篇覆盖备注会不会把现有格式拍平——只吃两个事实（每段几截 run、其中几截带字符格式）。
 *
 * 判定只存这一份（§3.2）：`setSlideNotes` 落地时按现场 XML 数出这两个数，`slideAgent` 在提案阶段按读模型数出来，
 * 两边念同一句理由。用户点了确认才发现"整批不写"（§2.4）是更坏的结果，所以装配层必须能提前判。
 *
 * @param paragraphRunCounts 备注逐段的 run 截数。
 * @param styledRunCount 带 `<a:rPr>`（颜色、加粗、字号）的 run 总截数。
 * @returns 一句中文理由；不会拍平时 null。
 */
export function notesFlattenNote(paragraphRunCounts: number[], styledRunCount: number): string | null {
  const multiRun = paragraphRunCounts.filter((count) => count > 1).length;
  if (multiRun > 0) return `备注里有 ${multiRun} 段带多截不同格式的文字，整篇覆盖会把它们压平，这一档不改。`;
  if (styledRunCount > 0) return `备注里有 ${styledRunCount} 截带字符格式的文字，整篇覆盖会把它们压平，这一档不改。`;
  return null;
}

/**
 * 产出一条"把某页的演讲者备注整篇换成给定的几段"的编辑提案，**不落地**。
 *
 * 只改**文件里已经存在**的备注部件。新建一份 `notesSlide` 按 §2.10 要三处一起登记（部件本体 + 幻灯片侧 `<Relationship>`
 * + `[Content_Types].xml` 的 Override，另加备注页自己指向幻灯片与 notesMaster 的两份关系），而本机 WPS 的 8 份真模板里
 * 只有 4 份带 notesMaster 可指（取证见 spec 阶段 2）——没有备注母版时宁可回一句理由，也不凭空造一份母版。
 *
 * @param file 已打开的读写器。
 * @param target 备注部件路径（读模型的 `Slide.notes.partPath`）。
 * @param paragraphTexts 新的备注逐段文字，一段一个字符串（空串 = 那一行没字）。
 * @returns `{ ok, op }`：部件缺失、没有 `type="body"` 那一格、现有备注带多种格式（整篇覆盖会拍平，§2.3）时 `{ ok:false, reason }`。
 *   成功时 op 的前/后像是**备注部件**的，幻灯片部件一个字节都不动。
 */
export async function setSlideNotes(file: PptxFile, target: SlideNotesTarget, paragraphTexts: string[]): Promise<SlideEditResult> {
  if (paragraphTexts.length === 0) return { ok: false, reason: "没有要写的备注段落。" };
  const loaded = await loadSlidePart(file, target.notesPartPath, "备注部件");
  if (!loaded.ok) return loaded;
  const bodyShape = findNotesBodyShape(loaded.root);
  const txBody = bodyShape ? firstChild(bodyShape, "p:txBody") : undefined;
  if (!txBody) return { ok: false, reason: `这份备注里没有正文那一格（<p:ph type="body">）：${target.notesPartPath}` };
  const paragraphs = findChildren(txBody, "a:p");
  // 新写的 run 只有 `<a:t>`：现有备注里那些多截格式与字符格式带不过来就是拍平（§2.3），判定与装配层共用那一份。
  const refusal = notesFlattenNote(
    paragraphs.map((p) => findChildren(p, "a:r").length),
    paragraphs.flatMap((p) => findChildren(p, "a:r")).filter((run) => firstChild(run, "a:rPr")).length,
  );
  if (refusal) return { ok: false, reason: refusal };
  // 段落自己的 `<a:pPr>`（项目符号、层级）与段末 `<a:endParaRPr>` 从第一段带过来：只换文字，不动这一页已有的版式意图（§2.7）。
  const template = paragraphs[0];
  const pPrXml = template ? firstChild(template, "a:pPr") : undefined;
  const endParaXml = template ? children(template).find((el) => el.name === "a:endParaRPr") : undefined;
  const heads = children(txBody).filter((el) => el.name !== "a:p");
  const built = paragraphTexts.map((line) =>
    parseFragmentElement(
      `<root><a:p>${pPrXml ? serialize(pPrXml) : ""}<a:r><a:t>${escapeText(line)}</a:t></a:r>${endParaXml ? serialize(endParaXml) : ""}</a:p></root>`,
    ),
  );
  txBody.children = [...heads, ...built];
  return { ok: true, op: toOp(target.notesPartPath, loaded.beforeXml, loaded.root) };
}

/**
 * 落地一条编辑提案：把后像写回部件。
 *
 * @param file 读写器。
 * @param op 由本文件那八个 op 生产者（`setShapeRunText` … `addTextBox`/`removeShape`/`setSlideNotes`）产出的 op。
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

/** `loadSlidePart` 的成功结果：部件原文与解析好的根元素。 */
type LoadedSlidePart = { ok: true; beforeXml: string; root: Element } | { ok: false; reason: string };

/** `locateShape` 的成功结果：部件原文、解析好的根与目标形状元素。 */
type LocatedShape = { ok: true; beforeXml: string; root: Element; shape: Element; container: Element } | { ok: false; reason: string };

/** `locateRun` 的成功结果：在形状之上再给出目标 run 元素。 */
type LocatedRun = { ok: true; beforeXml: string; root: Element; run: Element } | { ok: false; reason: string };

/**
 * 读部件并解析成根元素——所有 op 共用的第一层（§3.2：`locateShape`、页粒度的 `addTextBox` 与备注那一档都要这一步）。
 *
 * @param file 读写器。
 * @param partPath 部件路径（幻灯片或备注部件）。
 * @param partLabel 理由里怎么称呼这个部件：默认"幻灯片部件"，备注那一档传"备注部件"（同一份代码，两种说法，§5.7）。
 * @returns 成功带 `{ beforeXml, root }`（`beforeXml` 就是 op 的前像）；部件缺失或解析失败带 `{ ok:false, reason }`。
 */
async function loadSlidePart(file: PptxFile, partPath: string, partLabel = "幻灯片部件"): Promise<LoadedSlidePart> {
  const beforeXml = await file.partText(partPath);
  if (beforeXml === undefined) return { ok: false, reason: `${partLabel}不存在：${partPath}` };
  try {
    return { ok: true, beforeXml, root: parseDocument(beforeXml) };
  } catch {
    return { ok: false, reason: `${partLabel}解析失败。` };
  }
}

/**
 * 读部件、解析并按 `<p:cNvPr @id>` 定位形状元素——文字/字体/颜色/几何/删形状所有 op 共用的第二层（§3.2）。
 *
 * @param file 读写器。
 * @param target 部件路径 + 形状号。
 * @returns 成功带 `{ beforeXml, root, shape }`；部件缺失、解析失败或找不到形状带 `{ ok:false, reason }`。
 */
async function locateShape(file: PptxFile, target: ShapeGeometryTarget): Promise<LocatedShape> {
  const loaded = await loadSlidePart(file, target.partPath);
  if (!loaded.ok) return loaded;
  const found = findShapeById(loaded.root, target.shapeId);
  if (!found) return { ok: false, reason: `找不到形状 id=${target.shapeId}` };
  return { ...loaded, shape: found.shape, container: found.container };
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
 * 在幻灯片根元素里按 `<p:cNvPr @id>` 找形状元素，**并走进组合**。
 *
 * 为什么要递归：`<p:grpSp>` 里的形状在页面 XML 里真实存在、有自己的 `shapeId`、带着用户写的文字，
 * 只扫顶层就是"那些文字不存在"（实测见 spec 阶段 2）。同时把**直接装着它的那一层**（页面 `spTree` 或某个组）
 * 一起交出去——删元素与重排叠放都必须对着一层来做，对着页面 `spTree` 摘一个组里的形状会静默摘不掉
 * 却报"成功"，那是最坏的一种失败。
 *
 * @param root 幻灯片根（`<p:sld>`）。
 * @param shapeId 形状号。
 * @returns `{ shape, container }`（`container` 是直接容纳它的那一层）；找不到返回 undefined。
 */
function findShapeById(root: Element, shapeId: number): { shape: Element; container: Element } | undefined {
  const spTree = shapeTreeOf(root);
  if (!spTree) return undefined;
  const walk = (container: Element): { shape: Element; container: Element } | undefined => {
    for (const el of children(container)) {
      if (!SHAPE_HOST_TAGS.includes(el.name)) continue;
      if (shapeIdOf(el) === shapeId) return { shape: el, container };
      if (el.name === "p:grpSp") {
        const found = walk(el);
        if (found) return found;
      }
    }
    return undefined;
  };
  return walk(spTree);
}

/**
 * 重排**一层**里的形状子元素（页面的 `<p:spTree>`，或某个 `<p:grpSp>` 内部）。
 *
 * 只填形状那些槽位：`p:nvGrpSpPr`、`p:grpSpPr` 与元素之间的空白 / 注释节点原地不动，
 * 这样重序列化后除叠放顺序外一个字都不差（§2.3 不许误伤别的内容）。
 *
 * @param container 要重排的那一层元素（页面形状树或组合）。
 * @param order 新的形状顺序，长度必须等于该层现有形状子元素数。
 */
function rewriteShapeOrder(container: Element, order: Element[]): void {
  const nodes = container.children;
  if (!nodes) return;
  const slots: number[] = [];
  nodes.forEach((node, index) => {
    if (isElement(node) && SHAPE_HOST_TAGS.includes(node.name)) slots.push(index);
  });
  slots.forEach((slot, k) => {
    nodes[slot] = order[k]!;
  });
}

/**
 * 这一页部件里已用的最大 `<p:cNvPr @id>`（新建形状时要在它上面 +1）。
 *
 * 单位/范围都照 python-pptx 的 `CT_GroupShape.max_shape_id`：扫**这一页 XML 里所有**数字 `@id`
 * （不只形状那些——分区的 `p:section/@id` 是 GUID，非数字，正好被过滤掉），取最大值。
 *
 * @param root 幻灯片根（`<p:sld>`）。
 * @returns 最大数字 id；一个都没有时 0（真实文件里 spTree 自己的 `cNvPr id="1"` 一定在）。
 */
function maxShapeIdInPart(root: Element): number {
  let max = 0;
  const walk = (node: XmlNode): void => {
    if (!isElement(node)) return;
    const raw = node.attributes["id"];
    if (raw !== undefined && /^\d+$/.test(raw)) max = Math.max(max, Number(raw));
    (node.children ?? []).forEach(walk);
  };
  walk(root);
  return max;
}

/**
 * 拼出一个 PowerPoint 插入文本框时那种 `<p:sp>` 片段（结构逐字取自 python-pptx 的 `_textbox_sp_tmpl()`）。
 *
 * 片段不带 `xmlns` 声明：解析器按前缀原样读写，命名空间由幻灯片根上那三条声明覆盖，
 * 带上反而会在重序列化时多出重复属性。
 *
 * @param shapeId 分配给新形状的 `p:cNvPr @id`。
 * @param box 位置/尺寸（CSS px @96DPI），写进 `a:xfrm` 时换成 EMU（§2.1 唯一换算点）。
 * @param text 文本框文字；空串时保留模板里那个空的 `<a:p/>`，不造空 run。
 * @returns 可直接交给 `parseFragmentElement` 的包装片段。
 */
function textBoxXml(shapeId: number, box: ShapeBoxPx, text: string): string {
  const para = text === "" ? "<a:p/>" : `<a:p><a:r><a:t>${escapeText(text)}</a:t></a:r></a:p>`;
  return (
    "<root>" +
    `<p:sp><p:nvSpPr><p:cNvPr id="${shapeId}" name="TextBox ${shapeId - 1}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="${pxToEmu(box.leftPx)}" y="${pxToEmu(box.topPx)}"/>` +
    `<a:ext cx="${pxToEmu(box.widthPx)}" cy="${pxToEmu(box.heightPx)}"/></a:xfrm>` +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>' +
    `<p:txBody><a:bodyPr wrap="none"><a:spAutoFit/></a:bodyPr><a:lstStyle/>${para}</p:txBody></p:sp>` +
    "</root>"
  );
}

/**
 * 把新建的形状元素挂到 `<p:spTree>` 末尾（= 最顶层，叠放语义见 `moveShapeLayer`）。
 *
 * 只碰形状那些槽位：`p:nvGrpSpPr`/`p:grpSpPr` 与节点间空白原地不动；这一页还没有形状时挂到子列表末尾。
 *
 * @param spTree 形状树元素。
 * @param shape 新建的形状元素。
 */
function appendShapeToTree(spTree: Element, shape: Element): void {
  const nodes = (spTree.children ??= []);
  let lastShapeIndex = -1;
  nodes.forEach((node, index) => {
    if (isElement(node) && SHAPE_HOST_TAGS.includes(node.name)) lastShapeIndex = index;
  });
  nodes.splice(lastShapeIndex + 1, 0, shape);
}

/**
 * 把一个形状元素从**直接装着它的那一层**里摘掉（页面 `<p:spTree>` 或某个 `<p:grpSp>`；节点间的空白留着，多一个空行不影响合法性）。
 *
 * @param container 由 `locateShape` 交出的那一层。对着页面 `spTree` 摘一个组里的形状会静默摘不掉却报"成功"，所以必须用它。
 * @param shape 目标形状元素（同一棵树里的同一个对象，按引用比较）。
 */
function removeShapeFromTree(container: Element, shape: Element): void {
  const nodes = container.children ?? [];
  const at = nodes.indexOf(shape);
  if (at >= 0) nodes.splice(at, 1);
}
