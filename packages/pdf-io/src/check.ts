/**
 * `.pdf` 的结构体检（spec D8 的 `oc pdf check`）：这份文件**办公软件会不会判损坏**、以及**我们这一档能不能改**。
 *
 * 与 `auditPptxPackage` 同一条纪律：判定住在引擎里，只出问题码 + 数字，中文文案在装配层（`AGENTS.md` §5.7），
 * CLI 与桌面 app 共用 `@office-copilot/shared` 的 `describePdfAuditIssue` 那一份（§3.2，同一逻辑不许各长一套）。
 * 这里的每条硬约束都对着 ISO 32000-1 的**显式**要求，不掺风格偏好：
 * - §7.5.2：文件第一行必须是 `%PDF-1.x`，每条交叉引用段以 `%%EOF` 结束——缺了就是截断或不是 PDF；
 * - §7.7.3.4：`/Pages` 的 `/Count` 必须等于 `/Kids` 的实际长度——不等时阅读器按自己那套猜，页数会错；
 * - §7.3.5：流字典声明 `/Filter` 却解不出字节，这一页的内容显示不出来。
 *
 * 刻意**没有**查的两类（如实排队，见 spec D8 的 W4）：加密（`/Encrypt`）与数字签名（`/AcroForm` 的 `/Sig`、`/Perms`）。
 * 本机没有能产出这两类文件的工具链，写一条测不到的分支不如不写。
 */
import { PDFArray, PDFDict, PDFDocument, PDFName, PDFNumber } from "@cantoo/pdf-lib";
import type { PdfAuditIssue, PdfAuditSeverity } from "@office-copilot/shared";
import { targetOfRun, type PdfTextRun } from "./pdf-model.js";
import { readPdfModel } from "./pdf-file.js";
import { contentKeys, decodeStream, filtersOf, streamOf } from "./streams.js";

// 三个类型住在 `@office-copilot/shared`（装配层与引擎都要用，§5.1 允许 pdf-io 依赖 shared），
// 这里按公开面原样再导出一次，`oc pdf check` 与桌面的 import 路径不用改。
export type { PdfAuditCode, PdfAuditIssue, PdfAuditSeverity } from "@office-copilot/shared";

/** 头尾探测的窗口：真文件的 `%PDF-` 一定在第 1 行，`%%EOF` 在最后几字节，1 KB 足够覆盖带 BOM 与注释行的写法。 */
const EDGE_WINDOW_BYTES = 1024;

/**
 * 体检一份 `.pdf`。
 *
 * @param bytes 磁盘上的原始字节。
 * @returns 全部结论，空数组表示"没发现会让办公软件判损坏的结构问题，且有可读可改的文字"；本函数不抛异常。
 */
export async function auditPdfStructure(bytes: Uint8Array): Promise<PdfAuditIssue[]> {
  const issues: PdfAuditIssue[] = [];
  const head = latin1(bytes, 0, EDGE_WINDOW_BYTES);
  const tail = latin1(bytes, Math.max(0, bytes.length - EDGE_WINDOW_BYTES), bytes.length);
  if (!head.includes("%PDF-")) issues.push({ code: "header-missing", severity: "broken", pageNumber: null, detail: `开头 8 字节读到的是 ${JSON.stringify(head.slice(0, 8))}` });
  if (!tail.includes("%%EOF")) issues.push({ code: "eof-missing", severity: "broken", pageNumber: null, detail: `结尾 16 字节读到的是 ${JSON.stringify(tail.slice(-16))}` });

  let doc: PDFDocument;
  try {
    doc = await PDFDocument.load(bytes);
  } catch (error) {
    issues.push({ code: "parse-failed", severity: "broken", pageNumber: null, detail: error instanceof Error ? error.message : String(error) });
    return issues;
  }

  // 实测：截断到只剩一节的文件能"载入成功"，但 trailer 的 /Root 指不到对象，`doc.catalog` 是 undefined。
  const catalog = doc.catalog;
  if (!catalog) {
    issues.push({ code: "catalog-missing", severity: "broken", pageNumber: null, detail: "trailer 的 /Root（目录对象）在文件里查不到，页树无从定位" });
    return issues;
  }
  const pagesTree = catalog.lookupMaybe(PDFName.of("Pages"), PDFDict);
  const count = pagesTree?.lookupMaybe(PDFName.of("Count"), PDFNumber)?.asNumber();
  const kids = pagesTree?.lookup(PDFName.of("Kids"));
  const kidCount = kids instanceof PDFArray ? kids.size() : null;
  if (!pagesTree) {
    issues.push({ code: "catalog-missing", severity: "broken", pageNumber: null, detail: "目录（catalog）里没有 /Pages 页树，页数无从确定" });
  } else if (count === undefined || kidCount === null || count !== kidCount || count !== doc.getPageCount()) {
    issues.push({ code: "page-count-mismatch", severity: "broken", pageNumber: null, detail: `/Pages 的 /Count = ${count ?? "没写"}，/Kids 实际 ${kidCount ?? "读不出"} 项，解析出的页树 ${doc.getPageCount()} 页——三者必须相等` });
  }

  doc.getPages().forEach((page, index) => {
    const keys = contentKeys(page.node.get(PDFName.of("Contents")));
    for (const key of keys) {
      const stream = streamOf(doc, key);
      if (!stream) {
        // 没有 /Contents 是合法的空白页；有键却查不到对象才是断引用。
        if (key) issues.push({ code: "content-stream-undecodable", severity: "broken", pageNumber: index + 1, detail: "这一页的 /Contents 指向的对象在文件里查不到" });
        continue;
      }
      if (!decodeStream(stream)) {
        issues.push({ code: "content-stream-undecodable", severity: "broken", pageNumber: index + 1, detail: `这一页的内容流解不出来（/Filter = ${filtersOf(stream).join(", ") || "没写"}）` });
      }
    }
  });

  const unreadable = await countUnreadableText(bytes);
  if (unreadable) issues.push({ code: "text-not-editable", severity: "limited", pageNumber: null, detail: unreadable });

  return issues;
}

/**
 * 数出"读得出位置但改不了"的文字，组一句给人看的数字。
 *
 * @param bytes 原始字节。
 * @returns 一句中文数字（全是可改的或压根没有文字时给 null，表示不必报这条）。
 */
async function countUnreadableText(bytes: Uint8Array): Promise<string | null> {
  const { runs } = await readPdfModel(bytes);
  const undecoded = runs.filter((run: PdfTextRun) => run.text === null).length;
  const inline = runs.filter((run) => targetOfRun(run) === null).length;
  if (runs.length === 0) return "整份文件一条显示文字都没读到（多半是扫描件或纯图片页）";
  if (undecoded === 0 && inline === 0) return null;
  const parts: string[] = [];
  if (undecoded) parts.push(`${undecoded} 条认不出内容（字体没带 /ToUnicode）`);
  if (inline) parts.push(`${inline} 条写在内联内容流里、没有可标记对象`);
  return `${runs.length} 条显示文字里 ${parts.join("、")}——这些读不出也改不了`;
}

/** 按 latin1 取一段字节（PDF 的头尾标记都是 ASCII，别的字节怎么翻不影响判存在）。 */
function latin1(bytes: Uint8Array, from: number, to: number): string {
  return Buffer.from(bytes.subarray(from, to)).toString("latin1");
}
