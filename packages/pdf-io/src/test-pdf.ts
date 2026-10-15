/**
 * 测试夹具：手工搭出 `.pdf` 的**最小合法结构**，只给 `@office-copilot/pdf-io/testing` 子路径用。
 *
 * 为什么手搭而不用 pdf-lib 的 `drawText`：本引擎要钉的是"两种串形态 + 两类字体"这四格组合，
 * `drawText` 只会给 hex + WinAnsi 那一种；而 CID/Identity-H 加一份手写的 `/ToUnicode` CMap
 * 才是中文产物的真实形状（spec D8 的 P11 实测：2 个嵌入字体只有 1 个有 ToUnicode）。
 * 手搭还能把内容流写成**不压缩**，让测试读得懂失败到底在哪一层。
 */
import { deflateSync } from "node:zlib";
import { PDFDict, PDFDocument, PDFName, PDFObject, PDFRawStream, PDFString } from "@cantoo/pdf-lib";

/** 一个字体 + 一段内容流的组合。 */
export interface PdfFixtureText {
  /** 要显示的文字（进内容流时按 `text` 原样写，已由调用方保证与字体编码匹配）。 */
  content: string;
  /** 资源字典里的字体键，如 `F1`。 */
  fontKey: string;
}

/**
 * 一份两页、简单字体（Helvetica + `WinAnsiEncoding`）的 PDF：第一页 literal 串、第二页 hex 串，
 * 第一页另挂一层 Form XObject，里面也有一段字（D8 要求"必须递归进 Form XObject"）。
 *
 * @returns 完整的 `.pdf` 字节。
 */
export async function simpleWinAnsiFixturePdf(): Promise<Uint8Array> {
  const doc = await PDFDocument.create();
  const pageOne = doc.addPage([595, 842]);
  const pageTwo = doc.addPage([595, 842]);
  const fontRef = doc.context.register(
    doc.context.obj({ Type: PDFName.of("Font"), Subtype: PDFName.of("Type1"), BaseFont: PDFName.of("Helvetica"), Encoding: PDFName.of("WinAnsiEncoding") }) as PDFDict,
  );

  const inner = "BT /F1 11 Tf 120 300 Tm (Inner Form Text) Tj ET\n";
  const innerRef = doc.context.register(
    PDFRawStream.of(
      doc.context.obj({
        Type: PDFName.of("XObject"),
        Subtype: PDFName.of("Form"),
        BBox: [0, 0, 595, 842],
        Resources: doc.context.obj({ Font: doc.context.obj({ F1: fontRef }) }) as PDFDict,
        Length: Buffer.byteLength(inner, "latin1"),
      }) as PDFDict,
      new Uint8Array(Buffer.from(inner, "latin1")),
    ),
  );

  const first = "BT /F1 18 Tf 60 760 Tm (Quarterly Report 2026) Tj ET\nq /Fr1 Do Q\n";
  const second = "BT /F1 18 Tf 60 760 Tm <417070656E646978204120737461797320707574> Tj ET\n";
  await addContent(doc, pageOne.node, first, { F1: fontRef, Fr1: innerRef });
  await addContent(doc, pageTwo.node, second, { F1: fontRef });
  return await doc.save();
}

/**
 * 一份 CID/Identity-H 字体 + 手写 `/ToUnicode` CMap 的 PDF，显示「中文标题」四个 CID。
 *
 * CMap 只登记 `1→中`、`2→文`、`3→标`、`4→题`，所以夹具天然带"子集缺字"这一格可测。
 *
 * @returns 完整的 `.pdf` 字节。
 */
export async function cidIdentityHFixturePdf(): Promise<Uint8Array> {
  return cidFixture(true);
}

/**
 * 同一份 CID 夹具，但**没有 `/ToUnicode`**——这是 P11 实测的真实常态（2 个嵌入字体只有 1 个带）。
 *
 * 钉两件事：读侧这条 run 的文字必须是 `null`（不知道原文就不许猜一个字上来），
 * 写侧对它的改写必须整批拒绝（§2.4），因为换回去的字节我们无法验证念作什么。
 *
 * @returns 完整的 `.pdf` 字节。
 */
export async function cidNoToUnicodeFixturePdf(): Promise<Uint8Array> {
  return cidFixture(false);
}

/**
 * WPS 那种**一个字形一条 `Tj`** 的 CID 夹具：「中文标题」拆成四条 run（每字后跟一条 `Td` 推进 8pt），仍挂 `/ToUnicode`。
 *
 * 为什么单开一份：四个 CID 写成一条 `Tj` 时拼行只有一行一条 run，引擎"按该行每条 run 的原字数切片"（spec D8 的
 * W2 第三刀）没有对象可切，整行改写能不能落盘、前缀字节会不会被撑开，就只有真形状的文件能回答。
 *
 * @returns 完整的 `.pdf` 字节。
 */
export async function cidPerGlyphFixturePdf(): Promise<Uint8Array> {
  return cidFixture(true, "BT /F2 16 Tf 60 760 Tm <0001> Tj 8 0 Td <0002> Tj 8 0 Td <0003> Tj 8 0 Td <0004> Tj ET\n");
}

/**
 * CID/Identity-H 夹具的共用构造：`withToUnicode` 决定字体字典里挂不挂 `/ToUnicode` CMap。
 *
 * @param withToUnicode 是否登记那四条 `1→中`、`2→文`、`3→标`、`4→题`。
 * @param content 内容流文本；默认把四个 CID 写成一条 `Tj`（Word / Google 那种一行一条 run 的形态）。
 */
async function cidFixture(withToUnicode: boolean, content = "BT /F2 16 Tf 60 760 Tm <0001000200030004> Tj ET\n"): Promise<Uint8Array> {
  const doc = await PDFDocument.create();
  const page = doc.addPage([595, 842]);
  const cmap = [
    "/CIDInit /ProcSet findresource begin",
    "12 dict begin",
    "begincmap",
    "/CMapName /Adobe-Identity-UCS def",
    "/CMapType 2 def",
    "1 begincodespacerange",
    "<0000> <FFFF>",
    "endcodespacerange",
    "4 beginbfchar",
    "<0001> <4E2D>",
    "<0002> <6587>",
    "<0003> <6807>",
    "<0004> <9898>",
    "endbfchar",
    "endcmap",
    "CMapName currentdict /CMap defineresource pop",
    "end",
    "end",
  ].join("\n");
  const cmapRef = doc.context.register(
    PDFRawStream.of(doc.context.obj({ Length: Buffer.byteLength(cmap, "latin1") }) as PDFDict, new Uint8Array(Buffer.from(cmap, "latin1"))),
  );
  const descriptorRef = doc.context.register(
    doc.context.obj({
      Type: PDFName.of("FontDescriptor"),
      FontName: PDFName.of("SourceHanSansCN-Regular"),
      Flags: 4,
      FontBBox: [-100, -100, 1000, 1000],
      ItalicAngle: 0,
      Ascent: 880,
      Descent: -120,
      CapHeight: 700,
      CIDSet: PDFName.of("Identity"),
    }) as PDFDict,
  );
  const descendantRef = doc.context.register(
    doc.context.obj({
      Type: PDFName.of("Font"),
      Subtype: PDFName.of("CIDFontType2"),
      BaseFont: PDFName.of("SourceHanSansCN-Regular"),
      CIDSystemInfo: doc.context.obj({ Registry: PDFString.of("Adobe"), Ordering: PDFString.of("GB1"), Supplement: 6 }) as PDFDict,
      FontDescriptor: descriptorRef,
      W: doc.context.obj([1, 4, doc.context.obj([500, 500, 500, 500])]) as unknown as PDFObject,
    }) as PDFDict,
  );
  const fontRef = doc.context.register(
    doc.context.obj({
      Type: PDFName.of("Font"),
      Subtype: PDFName.of("Type0"),
      BaseFont: PDFName.of("SourceHanSansCN-Regular"),
      Encoding: PDFName.of("Identity-H"),
      DescendantFonts: [descendantRef],
      ...(withToUnicode ? { ToUnicode: cmapRef } : {}),
    }) as PDFDict,
  );
  await addContent(doc, page.node, content, { F2: fontRef });
  return await doc.save();
}

/** 字符串对象小工具：`CIDSystemInfo` 那三个值是 literal string 而不是 name。 */

/**
 * 一份内容流**压缩**（`/FlateDecode`）的 PDF：真实产物几乎都是这一档。
 *
 * 这条夹具是回归网——本机第一版把 FlateDecode 当 raw deflate 解，两份真实文件读出 0 条 run，
 * 而所有未压缩夹具照样全绿；只有压缩夹具能抓住这种"真实文件全瞎"的缺陷。
 *
 * @returns 完整的 `.pdf` 字节。
 */
export async function flateWinAnsiFixturePdf(): Promise<Uint8Array> {
  const doc = await PDFDocument.create();
  const page = doc.addPage([595, 842]);
  const fontRef = doc.context.register(
    doc.context.obj({ Type: PDFName.of("Font"), Subtype: PDFName.of("Type1"), BaseFont: PDFName.of("Helvetica"), Encoding: PDFName.of("WinAnsiEncoding") }) as PDFDict,
  );
  const content = "BT /F1 18 Tf 60 760 Tm (Quarterly Report 2026) Tj ET\nBT /F1 12 Tf 60 700 Tm <466c617465204865782054657874> Tj ET\n";
  await addContent(doc, page.node, content, { F1: fontRef }, { flate: true });
  return await doc.save();
}

/**
 * 给一页挂上内容流，并把资源字典搭好。
 *
 * @param doc 文档。
 * @param pageNode 页面对象。
 * @param content 内容流文本（latin1）。
 * @param resourceEntries 资源键 → 对象引用（以 `Fr` 开头的进 `XObject`，其余进 `Font`）。
 * @param options `flate` 为真时把内容流压成 `/FlateDecode`（真实产物的常态）。
 */
async function addContent(doc: PDFDocument, pageNode: PDFDict, content: string, resourceEntries: Record<string, PDFObject>, options: { flate?: boolean } = {}): Promise<void> {
  const fonts: Record<string, PDFObject> = {};
  const xobjects: Record<string, PDFObject> = {};
  for (const [key, value] of Object.entries(resourceEntries)) {
    if (key.startsWith("Fr")) xobjects[key] = value;
    else fonts[key] = value;
  }
  const resources = doc.context.obj({ Font: doc.context.obj(fonts) as PDFDict, XObject: doc.context.obj(xobjects) as PDFDict }) as PDFDict;
  const raw = Buffer.from(content, "latin1");
  const body = options.flate ? deflateSync(raw) : raw;
  const dict = (options.flate
    ? doc.context.obj({ Filter: PDFName.of("FlateDecode"), Length: body.length })
    : doc.context.obj({ Length: body.length })) as PDFDict;
  const ref = doc.context.register(PDFRawStream.of(dict, new Uint8Array(body)));
  pageNode.set(PDFName.of("Contents"), ref);
  pageNode.set(PDFName.of("Resources"), resources);
}
