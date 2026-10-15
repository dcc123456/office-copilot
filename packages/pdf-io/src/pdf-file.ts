/**
 * `.pdf` 的读模型与 **W1 档正文改写**（spec D8：同框、同字体、同长度替换，落地走增量更新）。
 *
 * 保真的答案是 ISO 32000-1 §7.5.6 增量更新——"改动必须追加到文件末尾，原内容保持不动"：
 * 本机实测（`tmp/pdf-w1` 的探针）自造两页文件改一处文字，**前缀 1 437 字节逐字不变**，只追加 396 字节 delta，
 * Apple PDFKit 读回新文字、第二页原样。对照全量 `save()` 会重写整份文件，那是"改一个字动整份"，不可接受。
 *
 * 三条实测出来的硬约束（不是推测）：
 * - 文字在内容流里以 **hex 或 literal 两种串形态**出现：我们自产的 PDF 走 hex，Word / Chrome 的走 literal，定位必须两种都认；
 * - 内容流普遍是 `/FlateDecode`，不解压就搜不到字（探针第一版就是这么扑空的）；
 * - 改完的串**必须装回原来的字节数**：PDF 没有回流，长度不一律拒绝（`length-mismatch`），不做任何"顺手补空格"的把戏。
 *
 * 依赖方向（`AGENTS.md` §5.1）：本包只依赖 `@cantoo/pdf-lib`（MIT）与 Node 自带 zlib，不碰 `opc`/`docx-io`/`sheet-io`/`pptx-io`。
 */
import { deflateSync, inflateSync } from "node:zlib";
import { PDFArray, PDFDict, PDFDocument, PDFName, PDFRawStream, type PDFObject, type PDFRef, type PDFStream } from "@cantoo/pdf-lib";
import { decodeOperandText, encodeOperandToken, parseToUnicodeCMap, standardEncodingMap, type CodeMap } from "./code-map.js";
import { scanContentStream, type TextRun } from "./content-stream.js";

/** 内容流解压上限：超过这个字节数就不解（挡住把整份大文件读进字符串的失控情况）。 */
const MAX_STREAM_BYTES = 8_388_608;

/** 一条显示文字的算子在模型里的样子。 */
export interface PdfTextRun {
  /** 页号（0 基）。 */
  pageIndex: number;
  /** 所在内容流的对象引用（如 `6 0 R`）；内联（非间接对象）内容流为 null，那种改不了。 */
  streamRef: string | null;
  /** 从页面内容流进第几层 Form XObject 的资源名链（如 `Fm0/Fm1`）；主内容流为 null。 */
  withinForm: string | null;
  /** 算子名。 */
  operator: TextRun["operator"];
  /** 串的书写形态。 */
  form: "literal" | "hex";
  /** 操作数在**解压后**内容流里的起始下标（字节）。与 `streamRef`、`withinForm` 合起来才是一条 run 的身份。 */
  start: number;
  /** 原串的字节数。 */
  encodedBytes: number;
  /** 原 token 的字符数（含定界符），改回去必须等长。 */
  tokenLength: number;
  /** 生效字体的 `/BaseFont`；资源链里查不到这个字体时 null。 */
  baseFont: string | null;
  /** 是不是 Type0/CID 复合字体（中文产物基本都是）。 */
  composite: boolean;
  /** 解出来的文字；字体认不出或子集缺码时为 null。 */
  text: string | null;
}

/** `.pdf` 的读模型。 */
export interface PdfReadModel {
  /** 页数。 */
  pageCount: number;
  /** 全部文字 run，按页、按内容流、按下标排序。 */
  runs: PdfTextRun[];
}

/** 一条改写的定位键（取自 {@link PdfTextRun} 的同名字段）。 */
export interface PdfTextTarget {
  /** 页号（0 基）。 */
  pageIndex: number;
  /** 内容流对象引用。 */
  streamRef: string;
  /** Form XObject 资源名链；主内容流给 null。 */
  withinForm: string | null;
  /** 操作数起始下标。 */
  start: number;
}

/** 一处待落地的文字改写。 */
export interface PdfTextEdit {
  /** 改哪一条 run。 */
  target: PdfTextTarget;
  /** 新文字（Unicode）。 */
  text: string;
}

/** 一处改不了的原因（结构化，文案由调用方组装，`AGENTS.md` §5.7）。 */
export interface PdfEditFailure {
  /** 失败的那一条定位键。 */
  target: PdfTextTarget;
  /** 原因码：找不到这条 run / 字体认不出 / 子集里没有这个字 / 新串装不回原字节数。 */
  reason: "run-not-found" | "no-encoding" | "char-outside-subset" | "length-mismatch";
}

/** 改写结果：要么给出新的完整字节，要么给出全部失败项且不产出任何改动（§2.4 整批语义）。 */
export type PdfEditResult = { ok: true; bytes: Uint8Array; deltaBytes: number } | { ok: false; failures: PdfEditFailure[] };

/** 内部：一条 run 附带它的码表与内容流对象，改写时要再编码一次，故不丢。 */
interface LocatedRun extends PdfTextRun {
  /** 生效字体的码表；认不出为 null。 */
  codes: CodeMap | null;
  /** 所在内容流对象（仅间接对象才有；内联流为 null）。 */
  stream: PDFRawStream | null;
  /** 内容流的引用对象，增量保存时要显式标记它。 */
  ref: PDFRef | null;
}

/**
 * 读一份 `.pdf` 的页与文字模型。
 *
 * @param bytes 磁盘上的 `.pdf` 原始字节。
 * @returns 页数与全部文字 run；不是 PDF 时由底层解析抛错，调用方按读文件失败处理。
 */
export async function readPdfModel(bytes: Uint8Array): Promise<PdfReadModel> {
  const doc = await PDFDocument.load(bytes);
  const runs = collectRuns(doc);
  return {
    pageCount: doc.getPageCount(),
    runs: runs.map(({ codes: _codes, stream: _stream, ref: _ref, ...rest }) => rest).sort(byPosition),
  };
}

/**
 * 整批改写文字并做**增量保存**。
 *
 * @param bytes 原始 `.pdf` 字节。
 * @param edits 待落地改动，定位键取自 {@link readPdfModel}。
 * @returns 任一条失败时 `{ ok: false, failures }` 且不写任何东西；全成时 `{ ok: true, bytes }`，
 *   其中 bytes 是"原文件逐字节 + 追加的 delta"，原前缀一字未动（§7.5.6）。
 */
export async function applyPdfTextEdits(bytes: Uint8Array, edits: PdfTextEdit[]): Promise<PdfEditResult> {
  const doc = await PDFDocument.load(bytes, { forIncrementalUpdate: true });
  const snapshot = doc.takeSnapshot();
  const runs = collectRuns(doc);
  const failures: PdfEditFailure[] = [];
  /** 每条待改内容流的解压文本与逐条替换。 */
  const patches = new Map<PDFRawStream, { decoded: string; flate: boolean; splices: Splice[] }>();

  for (const edit of edits) {
    const run = runs.find(
      (candidate) =>
        candidate.pageIndex === edit.target.pageIndex &&
        candidate.streamRef === edit.target.streamRef &&
        candidate.withinForm === edit.target.withinForm &&
        candidate.start === edit.target.start,
    );
    if (!run || !run.stream || !run.ref) {
      failures.push({ target: edit.target, reason: "run-not-found" });
      continue;
    }
    if (!run.codes) {
      failures.push({ target: edit.target, reason: "no-encoding" });
      continue;
    }
    const entry = patches.get(run.stream) ?? openStream(run.stream);
    if (!entry) {
      failures.push({ target: edit.target, reason: "run-not-found" });
      continue;
    }
    const encoded = encodeOperandToken(edit.text, { encodedBytes: run.encodedBytes, tokenLength: run.tokenLength }, run.codes);
    if ("failure" in encoded) {
      failures.push({ target: edit.target, reason: encoded.failure });
      continue;
    }
    patches.set(run.stream, entry);
    entry.splices.push({ start: run.start, end: run.start + run.tokenLength, token: encoded.token });
  }

  if (failures.length > 0) return { ok: false, failures };

  for (const [stream, entry] of patches) {
    const next = Buffer.from(splice(entry.decoded, entry.splices), "latin1");
    stream.updateContents(entry.flate ? new Uint8Array(deflateSync(next)) : new Uint8Array(next));
    const located = runs.find((candidate) => candidate.stream === stream);
    if (located?.ref) snapshot.markRefForSave(located.ref);
  }
  const delta = await doc.saveIncremental(snapshot);
  return { ok: true, bytes: new Uint8Array(Buffer.concat([Buffer.from(bytes), Buffer.from(delta)])), deltaBytes: delta.length };
}

/** 一处替换：内容流文本里的 `[start, end)` 换成 `token`。 */
interface Splice {
  /** 起始下标（字节）。 */
  start: number;
  /** 结束下标（不含）。 */
  end: number;
  /** 写回去的操作数。 */
  token: string;
}

/**
 * 遍历每页的内容流（含 Form XObject 递归），扫出全部文字 run 并解出文字。
 *
 * @param doc 已载入的文档。
 * @returns 未排序的 run；`streamRef` 为 null 的是内联内容流，改不了但照样报数。
 */
function collectRuns(doc: PDFDocument): LocatedRun[] {
  const runs: LocatedRun[] = [];
  doc.getPages().forEach((page, pageIndex) => {
    const pageDict = page.node;
    const resources = pageDict.lookupMaybe(PDFName.of("Resources"), PDFDict);
    for (const contents of contentKeys(doc, pageDict.get(PDFName.of("Contents")))) {
      walkStream(doc, contents, pageIndex, null, resources, new Set<string>(), runs);
    }
  });
  return runs;
}

/**
 * 一个内容流：扫 run，再递归它 `Do` 出来的 Form XObject。
 *
 * @param doc 文档。
 * @param key 内容流键（引用、数组摊开后的单项，或直接对象）。
 * @param pageIndex 页号。
 * @param withinForm 已走过的 Form 资源名链；主内容流是 null。
 * @param inheritedResources 页面资源字典（Form 自己没带 Resources 时沿用它）。
 * @param visited 本页已访问过的流，防环。
 * @param runs 输出数组，就地追加。
 */
function walkStream(doc: PDFDocument, key: unknown, pageIndex: number, withinForm: string | null, inheritedResources: PDFDict | undefined, visited: Set<string>, runs: LocatedRun[]): void {
  const stream = streamOf(doc, key);
  if (!stream) return;
  const ref = refOf(doc, stream);
  const identity = `${ref?.toString() ?? "inline"}|${withinForm ?? ""}`;
  if (visited.has(identity)) return;
  visited.add(identity);
  const decoded = decodeStream(stream);
  if (!decoded) return;
  const resources = stream.dict.lookupMaybe(PDFName.of("Resources"), PDFDict) ?? inheritedResources;
  const scan = scanContentStream(decoded.text);
  const fontDict = resources?.lookupMaybe(PDFName.of("Font"), PDFDict);
  for (const run of scan.runs) {
    const font = run.fontName && fontDict ? fontDict.lookupMaybe(PDFName.of(run.fontName), PDFDict) : undefined;
    const codes = font ? codeMapOfFont(doc, font) : null;
    runs.push({
      pageIndex,
      streamRef: ref ? ref.toString() : null,
      withinForm,
      operator: run.operator,
      form: run.operand.form,
      start: run.operand.start,
      encodedBytes: run.operand.bytes.length,
      tokenLength: run.operand.end - run.operand.start,
      baseFont: font ? plainName(font.get(PDFName.of("BaseFont"))) : null,
      composite: font ? plainName(font.get(PDFName.of("Subtype"))) === "Type0" : false,
      codes,
      text: decodeOperandText(run.operand, codes),
      stream: ref ? stream : null,
      ref,
    });
  }
  const xobjectDict = resources?.lookupMaybe(PDFName.of("XObject"), PDFDict);
  for (const call of scan.xobjects) {
    const form = xobjectDict?.get(PDFName.of(call.name));
    const formStream = streamOf(doc, form);
    if (!formStream || plainName(formStream.dict.get(PDFName.of("Subtype"))) !== "Form") continue;
    walkStream(doc, formStream, pageIndex, withinForm ? `${withinForm}/${call.name}` : call.name, resources ?? inheritedResources, visited, runs);
  }
}

/**
 * 取一个字体的码表：`/ToUnicode` 优先（§9.10.3），没有再退到标准编码表。
 *
 * @param doc 文档。
 * @param font 字体字典。
 * @returns 码表，或 null（`Identity-V`、只给 `/Differences` 字形名、无 ToUnicode 的 CID 字体等一律不猜）。
 */
function codeMapOfFont(doc: PDFDocument, font: PDFDict): CodeMap | null {
  const toUnicode = toUnicodeMap(doc, font);
  if (plainName(font.get(PDFName.of("Subtype"))) === "Type0") {
    return plainName(font.get(PDFName.of("Encoding"))) === "Identity-H" ? toUnicode : null;
  }
  return toUnicode ?? standardEncodingMap(plainName(font.get(PDFName.of("Encoding"))) ?? "");
}

/** 找一份字体（含其后代字体）上的 `/ToUnicode` CMap 并解出来。 */
function toUnicodeMap(doc: PDFDocument, font: PDFDict): CodeMap | null {
  const dicts: PDFDict[] = [font];
  const descendants = font.lookupMaybe(PDFName.of("DescendantFonts"), PDFArray);
  const firstDescendant = descendants ? dictOf(doc, descendants.get(0)) : undefined;
  if (firstDescendant) dicts.push(firstDescendant);
  for (const dict of dicts) {
    const cmap = streamOf(doc, dict.get(PDFName.of("ToUnicode")));
    const text = cmap ? decodeStream(cmap)?.text : null;
    if (!text) continue;
    const map = parseToUnicodeCMap(text);
    if (map) return map;
  }
  return null;
}

/**
 * 解压一个流并给出可扫描的文本。
 *
 * @param stream 流对象。
 * @returns `{ text, flate }`；压缩比不正常或超过 {@link MAX_STREAM_BYTES} 时 null（跳过，不抛）。
 */
function decodeStream(stream: PDFRawStream): { text: string; flate: boolean } | null {
  const raw = Buffer.from(stream.getContents());
  const flate = filtersOf(stream).includes("FlateDecode");
  if (raw.length > MAX_STREAM_BYTES) return null;
  try {
    const bytes = flate ? inflateSync(raw) : raw;
    if (bytes.length > MAX_STREAM_BYTES) return null;
    return { text: bytes.toString("latin1"), flate };
  } catch {
    return null;
  }
}

/** 一个流的 `/Filter` 名字列表（可能是单名或数组）。 */
function filtersOf(stream: PDFRawStream): string[] {
  const filter = stream.dict.get(PDFName.of("Filter"));
  if (filter instanceof PDFArray) return Array.from({ length: filter.size() }, (_, index) => plainName(filter.get(index)) ?? "");
  const single = plainName(filter);
  return single ? [single] : [];
}

/** 打开一条已定位内容流，准备收集替换位。 */
function openStream(stream: PDFRawStream): { decoded: string; flate: boolean; splices: Splice[] } | null {
  const decoded = decodeStream(stream);
  return decoded ? { decoded: decoded.text, flate: decoded.flate, splices: [] } : null;
}

/** 一个键（引用 / 直接对象 / null）对应的流对象。 */
function streamOf(doc: PDFDocument, key: unknown): PDFRawStream | null {
  const value = key instanceof PDFRawStream ? key : key ? doc.context.lookup(key as PDFObject) : undefined;
  return value instanceof PDFRawStream ? value : null;
}

/** 一个键对应的字典对象。 */
function dictOf(doc: PDFDocument, key: PDFObject): PDFDict | undefined {
  const value = key instanceof PDFDict ? key : doc.context.lookup(key);
  return value instanceof PDFDict ? value : undefined;
}

/** 一个对象在上下文里的间接引用；它不是间接对象时 null。 */
function refOf(doc: PDFDocument, object: PDFObject): PDFRef | null {
  for (const [ref, value] of doc.context.enumerateIndirectObjects()) if (value === object) return ref;
  return null;
}

/** 把 `/Contents` 的值（单引用、数组或直接流）摊成键列表。 */
function contentKeys(doc: PDFDocument, value: PDFObject | undefined): PDFObject[] {
  if (!value) return [];
  if (value instanceof PDFArray) return Array.from({ length: value.size() }, (_, index) => value.get(index));
  return [value];
}

/** 按倒序把 token 拼回内容流文本（同长替换，所以前面的下标不受影响）。 */
function splice(decoded: string, splices: Splice[]): string {
  let out = decoded;
  for (const item of [...splices].sort((left, right) => right.start - left.start)) out = out.slice(0, item.start) + item.token + out.slice(item.end);
  return out;
}

/** 读模型与编辑都要同一套顺序：按页、主内容流先于它递归进去的 Form XObject、再按流内下标。 */
function byPosition(left: PdfTextRun, right: PdfTextRun): number {
  return left.pageIndex - right.pageIndex || formDepth(left.withinForm) - formDepth(right.withinForm) || left.start - right.start;
}

/** Form XObject 的嵌套深度：主内容流是 0，`Fr1` 是 1，`Fr1/Fr2` 是 2。 */
function formDepth(withinForm: string | null): number {
  return withinForm === null ? 0 : withinForm.split("/").length;
}

/** PDFName → 去掉前导 `/` 的文本；不是名字时 null。 */
function plainName(value: unknown): string | null {
  return value instanceof PDFName ? value.decodeText() : null;
}
