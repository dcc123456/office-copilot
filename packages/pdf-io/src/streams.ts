/**
 * 内容流层的共用件：一个键（引用 / 数组摊开 / 直接对象）拿到流、按 `/Filter` 决定怎么解、解成可扫描的文本。
 *
 * 从 `pdf-file.ts` 抽出来是因为**读侧与体检侧都要同一份判定**（spec D8 的 `oc pdf check`）：
 * "这条内容流解不出来"在两处必须给出同一个答案，写成两份就会有一处悄悄放宽（`AGENTS.md` §3.2）。
 * 本模块不进包的公开入口（`index.ts` 不导出它），只给包内用。
 */
import { inflateSync } from "node:zlib";
import { PDFArray, PDFName, PDFObject, PDFRawStream, type PDFDocument } from "@cantoo/pdf-lib";

/** 内容流解压上限：超过这个字节数就不解（挡住把整份大文件读进字符串的失控情况）。 */
const MAX_STREAM_BYTES = 8_388_608;

/**
 * 解压一个流并给出可扫描的文本。
 *
 * @param stream 流对象。
 * @returns `{ text, flate }`；压缩比不正常或超过 {@link MAX_STREAM_BYTES} 时 null（跳过，不抛）。
 */
export function decodeStream(stream: PDFRawStream): { text: string; flate: boolean } | null {
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
export function filtersOf(stream: PDFRawStream): string[] {
  const filter = stream.dict.get(PDFName.of("Filter"));
  if (filter instanceof PDFArray) return Array.from({ length: filter.size() }, (_, index) => plainName(filter.get(index)) ?? "");
  const single = plainName(filter);
  return single ? [single] : [];
}

/** 一个键（引用 / 直接对象 / null）对应的流对象。 */
export function streamOf(doc: PDFDocument, key: unknown): PDFRawStream | null {
  const value = key instanceof PDFRawStream ? key : key ? doc.context.lookup(key as PDFObject) : undefined;
  return value instanceof PDFRawStream ? value : null;
}

/** 把 `/Contents` 的值（单引用、数组或直接流）摊成键列表。 */
export function contentKeys(value: PDFObject | undefined): PDFObject[] {
  if (!value) return [];
  if (value instanceof PDFArray) return Array.from({ length: value.size() }, (_, index) => value.get(index));
  return [value];
}

/** PDFName → 去掉前导 `/` 的文本；不是名字时 null。 */
export function plainName(value: unknown): string | null {
  return value instanceof PDFName ? value.decodeText() : null;
}
