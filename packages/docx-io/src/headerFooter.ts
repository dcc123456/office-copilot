import type JSZip from "jszip";
import { children, firstChild, parseDocument, wordPartPath } from "./xml.js";
import { ooxmlFragmentToPm, pmToOoxmlFragment } from "./ooxml-pm.js";
import { NumberingMap } from "./numbering.js";
import { emptyPmDoc, type PMDoc } from "./pm-schema.js";
import { StyleMap } from "./styles.js";

export type HeaderFooterType =
  | "header"
  | "firstHeader"
  | "evenHeader"
  | "footer"
  | "firstFooter"
  | "evenFooter";

export interface HeaderFooterEntry {
  type: HeaderFooterType;
  rId: string;
  /** Target path inside the zip, e.g. word/header1.xml. */
  path: string;
  pm: PMDoc;
  /** Raw original XML so we can write it back if the user doesn't edit it. */
  rawXml: string;
  /** True when the user edited this header/footer and we must re-serialize pm. */
  dirty?: boolean;
}

export interface HeaderFooterBundle {
  /** type → entry. We only model default header/footer for now. */
  entries: Record<string, HeaderFooterEntry>;
}

const TYPE_TAG: Record<HeaderFooterType, string> = {
  header: "w:headerReference",
  firstHeader: "w:headerReference",
  evenHeader: "w:headerReference",
  footer: "w:footerReference",
  firstFooter: "w:footerReference",
  evenFooter: "w:footerReference",
};

const TYPE_ATTR: Record<HeaderFooterType, string> = {
  header: "default",
  firstHeader: "first",
  evenHeader: "even",
  footer: "default",
  firstFooter: "first",
  evenFooter: "even",
};

/**
 * Parse all <w:headerReference>/<w:footerReference> elements from a
 * document's sectPr, then read and parse every referenced part.
 *
 * We return a map keyed by type (header/footer); even/first variants
 * are tracked but the renderer currently uses the default everywhere.
 */
export async function loadHeaderFooters(
  zip: JSZip,
  documentXml: string,
  relsXml: string,
  numbering: NumberingMap,
  styles: StyleMap,
): Promise<HeaderFooterBundle> {
  const bundle: HeaderFooterBundle = { entries: {} };
  const docRoot = parseDocument(documentXml);
  const body = firstChild(docRoot, "w:body");
  const sectPr = body && firstChild(body, "w:sectPr");
  if (!sectPr) return bundle;

  // Build rId → target map from document.xml.rels. Tests and minimal
  // docs may ship without a rels part; treat that as "no references".
  const relMap = new Map<string, string>();
  if (relsXml && /<Relationships\b/.test(relsXml)) {
    try {
      const relsRoot = parseDocument(relsXml);
      for (const rel of children(relsRoot).filter(
        (c) => c.name === "Relationship",
      )) {
        const id = rel.attributes?.Id as string | undefined;
        const target = rel.attributes?.Target as string | undefined;
        if (id && target) relMap.set(id, target);
      }
    } catch {
      // Malformed rels — ignore; we just won't find header parts.
    }
  }

  const types: HeaderFooterType[] = [
    "header",
    "firstHeader",
    "evenHeader",
    "footer",
    "firstFooter",
    "evenFooter",
  ];
  for (const type of types) {
    const ref = children(sectPr).find(
      (c) =>
        c.name === TYPE_TAG[type] &&
        (c.attributes?.["w:type"] as string | undefined) === TYPE_ATTR[type],
    );
    if (!ref) continue;
    const rId = ref.attributes?.["r:id"] as string | undefined;
    if (!rId) continue;
    const target = relMap.get(rId);
    if (!target) continue;
    // Targets in document.xml.rels are relative to word/.
    const path = wordPartPath(target);
    const f = zip.file(path);
    if (!f) continue;
    const xml = await f.async("string");
    const pm = parseHeaderFooterXml(xml, numbering, styles);
    bundle.entries[type] = { type, rId, path, pm, rawXml: xml };
  }
  return bundle;
}

function parseHeaderFooterXml(
  xml: string,
  numbering: NumberingMap,
  styles: StyleMap,
): PMDoc {
  return ooxmlFragmentToPm(parseDocument(xml), numbering, styles);
}

/**
 * Re-serialize any dirty header/footer parts back to their full XML
 * document. The caller writes the resulting strings to the zip.
 *
 * Returns a map of path → xml for every part that should be updated.
 */
export function serializeHeaderFooters(
  bundle: HeaderFooterBundle,
  numbering: NumberingMap,
): Map<string, string> {
  const out = new Map<string, string>();
  for (const entry of Object.values(bundle.entries)) {
    if (!entry.dirty) {
      out.set(entry.path, entry.rawXml);
      continue;
    }
    // 页眉页脚不带自己的 sectPr，也不能留着 <w:document><w:body> 那层壳：
    // 段落要直接落在 <w:hdr>/<w:ftr> 下面。
    const bodyContent = pmToOoxmlFragment(entry.pm, { numbering });
    const tag = entry.type.startsWith("footer") ? "w:ftr" : "w:hdr";
    const xml =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<${tag} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
      `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
      bodyContent +
      `</${tag}>`;
    out.set(entry.path, xml);
  }
  return out;
}

/**
 * 取某一类页眉/页脚部件，文档里本来没有就当场建一个空的。
 *
 * rId 按类型给死（页眉 `rIdHf1`、页脚 `rIdHf2`）：两者撞号时
 * `writeHeaderFooters` 的 `ensureRel` 会以为关系已经登记过，页脚引用于是指到页眉部件上。
 *
 * @param bundle 当前文档的页眉页脚集合。
 * @param kind 建哪一类，只建这一类——建页眉时不该顺手塞一个用户没要过的页脚。
 * @returns 条目本体，调用方直接改它的 `pm` 与 `dirty`。
 */
export function ensureHeaderFooterPart(
  bundle: HeaderFooterBundle,
  kind: "header" | "footer",
): HeaderFooterEntry {
  if (!bundle.entries[kind]) {
    bundle.entries[kind] = {
      type: kind,
      rId: kind === "header" ? "rIdHf1" : "rIdHf2",
      path: `word/${kind}1.xml`,
      pm: emptyPmDoc(),
      rawXml: "",
      dirty: true,
    };
  }
  return bundle.entries[kind]!;
}
