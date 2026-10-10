import { escapeAttr, escapeText } from "./xml.js";

/**
 * 页面水印。Word 的"版式水印"不是节属性，而是**页眉部件里的一个 VML WordArt 形状**
 * （ECMA-376 §9.5.1 `w:pict` + VML `v:shape`，形状族 `_x0000_t136` 即 textshape），
 * 靠绝对定位 + 负 z-index 浮在正文后面。所以它的读写都在页眉 XML 上做，不进 `w:sectPr`。
 *
 * 单位跟着 VML 的原生量纲留在模型里：`sizePt` 是磅（`v:textpath/@style` 的 `font-size`
 * 本来就是 pt），`color` 是 `#rrggbb`（AGENTS §2.1 §2.2）。
 */
export interface SectionWatermark {
  /** 水印文字，`v:textpath/@string`。 */
  text: string;
  /** 字体族名，`v:textpath/@style` 里的 `font-family`。 */
  font: string;
  /** 字号，pt；`null` 是"自动"，也就是 Word 的 `fitshape="t"`（文字铺满形状框）。 */
  sizePt: number | null;
  /** 颜色，`#rrggbb`。 */
  color: string;
  /** 半透明，对应 Word 水印对话框里那个勾；写出去是 `v:fill/@opacity`。 */
  semiTransparent: boolean;
  /** 版式：斜排（对角）还是横排。 */
  layout: WatermarkLayout;
}

/** 水印版式，Word 对话框的两个单选钮。 */
export type WatermarkLayout = "diagonal" | "horizontal";

/** Word 自己给文字艺术形状用的 shapetype id，判定"这一段是水印"就靠它。 */
const SHAPE_TYPE_ID = "_x0000_t136";

/** 斜排的角度（VML `rotation` 是度数，315 = 左下到右上）。 */
const DIAGONAL_ROTATION = 315;

/** 半透明时写死的透明度；Word 的对话框只有"半透明"一个勾，没有滑杆。 */
const SEMI_TRANSPARENT_OPACITY = "50%";

/** 形状框：居中于页边距，尺寸取 Word 在 A4 上给水印的那个框。 */
const SHAPE_STYLE =
  "position:absolute;margin-left:0;margin-top:0;width:468pt;height:117pt;" +
  "z-index:-251654144;mso-wrap-edited:f;" +
  "mso-position-horizontal:center;mso-position-horizontal-relative:margin;" +
  "mso-position-vertical:center;mso-position-vertical-relative:margin";

const NS_V = "urn:schemas-microsoft-com:vml";
const NS_O = "urn:schemas-microsoft-com:office:office";
const NS_W10 = "urn:schemas-microsoft-com:office:word";

/** `_x0000_t136` 的形状定义。`path="m@7,l@8,…"` 引用的是这些 formula 的结果，缺一个都画不出来。 */
const SHAPE_TYPE_XML =
  `<v:shapetype id="${SHAPE_TYPE_ID}" coordsize="21600,21600" o:spt="136" adj="10800" path="m@7,l@8,m@5,21600l@6,21600e">` +
  `<v:formulas>` +
  `<v:f eqn="sum #0 0 10800"/><v:f eqn="prod #0 2 1"/><v:f eqn="sum 21600 0 @1"/>` +
  `<v:f eqn="sum 0 0 @2"/><v:f eqn="sum 21600 0 @3"/><v:f eqn="if @0 @3 0"/>` +
  `<v:f eqn="if @0 21600 @1"/><v:f eqn="if @0 0 @2"/><v:f eqn="if @0 @4 21600"/>` +
  `<v:f eqn="mid @5 @6"/><v:f eqn="mid @8 @5"/><v:f eqn="mid @7 @8"/><v:f eqn="mid @6 @7"/>` +
  `<v:f eqn="sum @6 0 @5"/>` +
  `</v:formulas>` +
  `<v:path textpathok="t" o:connecttype="custom" ` +
  `o:connectlocs="@9,0;@10,10800;@11,21600;@12,10800" o:connectangles="270,180,90,0"/>` +
  `<v:textpath on="t" fitshape="t"/>` +
  `<v:handles><v:h position="#0,bottomRight" xrange="6629,14971"/></v:handles>` +
  `<o:lock v:ext="edit" text="t" shapetype="t"/>` +
  `</v:shapetype>`;

/** 一个属性：从 `key="value"` 里取值，取不到给 undefined。 */
function attr(tag: string, key: string): string | undefined {
  const m = tag.match(new RegExp(`\\s${key}="([^"]*)"`));
  return m ? m[1] : undefined;
}

/** 把属性值里的实体还原成文字（页眉 XML 是字符串级处理，没走我们的解析器）。 */
function decodeAttr(raw: string): string {
  return raw
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&");
}

/**
 * 找出页眉 XML 里那个水印形状（带 `v:textpath/@string` 的 `_x0000_t136`）。
 *
 * 只认带文字的：`v:shapetype` 里也有一条 `<v:textpath on="t" fitshape="t"/>`，但它没有
 * `string`，靠这一点把定义块和实例分开。
 *
 * @param headerXml 页眉部件的完整 XML。
 * @returns 形状的标签属性串与内部 XML；没有水印形状时 undefined。
 */
function findWatermarkShape(
  headerXml: string,
): { attrs: string; inner: string } | undefined {
  const re = /<v:shape\b([^>]*)>([\s\S]*?)<\/v:shape>/g;
  for (let m = re.exec(headerXml); m; m = re.exec(headerXml)) {
    if (attr(m[1]!, "type") !== `#${SHAPE_TYPE_ID}`) continue;
    const tp = m[2]!.match(/<v:textpath\b[^>]*/)?.[0];
    if (!tp || attr(tp, "string") === undefined) continue;
    return { attrs: m[1]!, inner: m[2]! };
  }
  return undefined;
}

/**
 * 读页眉 XML 里的水印设置。
 *
 * 读不懂的一律当"这份文档没建模过水印"返回 undefined，让调用方原样保留页眉——
 * 例如 `fillcolor` 是颜色名而不是十六进制，我们宁可不动也不把读不懂的值写回去
 * （AGENTS §2.2）。
 *
 * @param headerXml 页眉部件的原始 XML；没有页眉时传 undefined。
 * @returns 水印模型，或 undefined。
 */
export function readWatermark(
  headerXml: string | undefined,
): SectionWatermark | undefined {
  if (!headerXml) return undefined;
  const shape = findWatermarkShape(headerXml);
  if (!shape) return undefined;
  const textpath = shape.inner.match(/<v:textpath\b[^>]*/)?.[0];
  if (!textpath) return undefined;
  const text = attr(textpath, "string");
  const color = attr(shape.attrs, "fillcolor");
  if (text === undefined || !color || !/^#[0-9a-fA-F]{6}$/.test(color)) {
    return undefined;
  }
  const style = attr(textpath, "style") ?? "";
  const fontFamily = style.match(/font-family:\s*'([^']*)'/)?.[1] ??
    style.match(/font-family:\s*([^;]*)/)?.[1]?.trim();
  if (!fontFamily) return undefined;

  const fitShape = /^(t|true|1)$/i.test(attr(textpath, "fitshape") ?? "");
  const sizePt = style.match(/font-size:\s*([\d.]+)pt/)?.[1];
  const opacity = attr(shape.inner, "opacity");
  const rawRotation =
    attr(shape.attrs, "rotation") ??
    attr(shape.attrs, "style")?.match(/rotation:\s*(-?[\d.]+)/)?.[1] ??
    "0";
  return {
    text: decodeAttr(text),
    font: fontFamily,
    // Word 的"自动"字号就是 fitshape + 一个 1pt 的占位值。
    sizePt: fitShape ? null : sizePt ? Number(sizePt) : null,
    color: color.toLowerCase(),
    semiTransparent: opacityToPercent(opacity) !== undefined,
    layout: Number(rawRotation) % 360 === 0 ? "horizontal" : "diagonal",
  };
}

/**
 * 读 VML 的 ST_Percentage：`50%` 与定点写法 `50000f` 都合法，Word 用后者。
 *
 * @param raw `v:fill/@opacity` 的字面值。
 * @returns 百分比数值（0–100）；只有真的是半透明（小于 100%）才返回，
 *   否则 undefined，含义是"没开半透明"。
 */
function opacityToPercent(raw: string | undefined): number | undefined {
  if (!raw) return undefined;
  const n = raw.endsWith("%")
    ? Number(raw.slice(0, -1))
    : raw.endsWith("f")
      ? Number(raw.slice(0, -1)) / 1000
      : Number.NaN;
  if (!Number.isFinite(n) || n <= 0 || n >= 100) return undefined;
  return n;
}

/**
 * 两份水印设置是否等价，`null`（用户选了"无水印"）与 `undefined`（没建模过）都算得进来。
 *
 * 页面设置面板保存时会把打开时读到的水印原样带回来（和节属性同一个路径），
 * 值没变就不该重写页眉——重写会把 Word 写的形状换成我们那一份，字节级别不一致。
 */
export function sameWatermark(
  a: SectionWatermark | null | undefined,
  b: SectionWatermark | null | undefined,
): boolean {
  if (!a || !b) return a === b;
  return (
    a.text === b.text &&
    a.font === b.font &&
    a.sizePt === b.sizePt &&
    a.color === b.color &&
    a.semiTransparent === b.semiTransparent &&
    a.layout === b.layout
  );
}

/**
 * 属性值转义，但**保留单引号**。
 *
 * `v:textpath/@style` 里的 CSS 写法是 `font-family:'黑体'`，双引号包着的属性值中单引号
 * 本来就不必转义；`escapeAttr` 会把它变成 `&apos;`，Word 与 WPS 就不再认得那个字体名。
 */
function escapeStyle(s: string): string {
  return escapeText(s).replace(/"/g, "&quot;");
}

/**
 * 造水印段落。
 *
 * @param watermark 水印设置。
 * @param withShapeType 页眉里还没有 `_x0000_t136` 定义时才带上定义块，重复 id 会让 Word 报错。
 * @returns 一个完整的 `<w:p>`，形状放在 `w:pict` 里。
 */
function watermarkParagraph(
  watermark: SectionWatermark,
  withShapeType: boolean,
): string {
  const auto = watermark.sizePt === null;
  const textpathStyle = `font-family:'${watermark.font}';font-size:${auto ? 1 : watermark.sizePt}pt`;
  const rotation =
    watermark.layout === "diagonal" ? ` rotation="${DIAGONAL_ROTATION}"` : "";
  const fill = watermark.semiTransparent
    ? `<v:fill on="t" opacity="${SEMI_TRANSPARENT_OPACITY}"/>`
    : `<v:fill on="t"/>`;
  const fitShape = auto ? ' fitshape="t"' : "";
  return (
    `<w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr>` +
    `<w:r><w:rPr><w:noProof/></w:rPr><w:pict>` +
    (withShapeType ? SHAPE_TYPE_XML : "") +
    `<v:shape id="OfficeCopilotWatermark" o:spid="_x0000_s1025" type="#${SHAPE_TYPE_ID}" ` +
    `style="${SHAPE_STYLE}"${rotation} fillcolor="${escapeStyle(watermark.color)}" stroked="f">` +
    fill +
    `<v:stroke on="f"/>` +
    `<v:textpath style="${escapeStyle(textpathStyle)}"${fitShape} string="${escapeAttr(watermark.text)}"/>` +
    `<o:lock v:ext="edit" shapetype="t"/>` +
    `</v:shape></w:pict></w:r></w:p>`
  );
}

/**
 * 这一段是不是水印段落。
 *
 * 判定收得很紧：必须是 `_x0000_t136` 形状 + 带 `string` 的 `v:textpath`，也就是 Word
 * 水印对话框产出的那个形状。用户自己插的艺术字、页眉里的文本框一律不算，不能删
 * （AGENTS §2.7：不擅自改动用户已有的版式意图）。
 */
function isWatermarkParagraph(para: string): boolean {
  if (para.includes("w:txbxContent")) return false;
  const shape = findWatermarkShape(para);
  return shape !== undefined;
}

/**
 * 把 `<w:hdr>` 根元素上缺的 VML 命名空间补上。
 *
 * 我们重建页眉时只声明了 `w:` 和 `r:`，水印形状用的 `v:`/`o:` 前缀没有声明就是非法 XML；
 * Word 自己的页眉一直带着这三个，所以只在缺的时候才加。
 */
function ensureVmlNamespaces(headerXml: string): string {
  const open = headerXml.match(/<w:hdr\b[^>]*>/)?.[0];
  if (!open || open.includes("xmlns:v=")) return headerXml;
  const declared =
    ` xmlns:v="${NS_V}" xmlns:o="${NS_O}" xmlns:w10="${NS_W10}"`;
  return headerXml.replace(open, `${open.slice(0, -1)}${declared}>`);
}

/**
 * 把水印写进（或从）页眉 XML。
 *
 * 先删掉已有的水印段落，再把新的插在 `<w:hdr>` 开头——形状是绝对定位的，段落顺序不影响
 * 它出现的位置。`null` 就是只删不插（用户在页面设置里选了"无水印"）。
 * 页眉里其余内容（文字、页码域、图片）一个字节都不动。
 *
 * @param headerXml 页眉部件的完整 XML（已经序列化好的那一份）。
 * @param watermark 要写的水印；null 表示清掉。
 * @returns 新的页眉 XML；没有水印可清且本来就没有时原样返回。
 */
export function applyWatermark(
  headerXml: string,
  watermark: SectionWatermark | null,
): string {
  const stripped = headerXml.replace(
    /<w:p\b[^>]*\/>|<w:p\b[^>]*>[\s\S]*?<\/w:p>/g,
    (para) => (isWatermarkParagraph(para) ? "" : para),
  );
  if (!watermark) return stripped;
  const withNs = ensureVmlNamespaces(stripped);
  const open = withNs.match(/<w:hdr\b[^>]*>/)?.[0];
  if (!open) return withNs;
  const para = watermarkParagraph(
    watermark,
    !withNs.includes(`id="${SHAPE_TYPE_ID}"`),
  );
  return withNs.replace(open, `${open}${para}`);
}
