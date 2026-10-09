import { children, firstChild, parseDocument } from "./xml.js";
import type { Element } from "./xml.js";

function val(el: Element | undefined): string | undefined {
  if (!el) return undefined;
  return (el.attributes?.["w:val"] as string | undefined) ?? undefined;
}

/** Twentieths of a point in one CSS px, the inverse of docx.ts's TWIP_TO_PX. */
const TWIPS_PER_PX = 15;

/**
 * Parsed lookups from word/styles.xml. For now we only need a mapping from
 * styleId → heading level, because real-world .docx files use arbitrary
 * numeric styleIds (e.g. WPS writes heading 1 with styleId="2").
 */
export class StyleMap {
  /** styleId → heading level (1..6). Missing key means not a heading. */
  private headingById = new Map<string, number>();
  /** Every styleId styles.xml actually declares, so unknown ids can be guessed. */
  private declared = new Set<string>();
  /** styleId → true if this is a list-related style (ListBullet/ListNumber). */
  listById = new Set<string>();
  /**
   * Space below a paragraph the document leaves unspecified, in CSS px. Word
   * takes it from the default paragraph style, and a template whose Normal
   * carries no `w:spacing` puts body text edge to edge — the flow height that
   * floating boxes anchored to those paragraphs are measured against.
   */
  defaultSpacingAfterPx = 0;

  static fromXml(xml: string | undefined | null): StyleMap {
    const map = new StyleMap();
    if (!xml) return map;
    try {
      const root = parseDocument(xml);
      for (const s of children(root).filter((c) => c.name === "w:style")) {
        const id = s.attributes?.["w:styleId"] as string | undefined;
        if (!id) continue;
        map.declared.add(id);
        const name = val(firstChild(s, "w:name"))?.toLowerCase() ?? "";
        const m = name.match(/^heading\s+([1-6])/);
        if (m) {
          map.headingById.set(id, Number(m[1]));
          continue;
        }
        if (name.includes("list bullet") || name.includes("list number")) {
          map.listById.add(id);
        }
        if (
          s.attributes?.["w:type"] === "paragraph" &&
          s.attributes?.["w:default"] === "1"
        ) {
          const pPr = firstChild(s, "w:pPr");
          const spacing = pPr && firstChild(pPr, "w:spacing");
          // `w:after` counts twentieths of a point; 15 of them make a CSS px.
          map.defaultSpacingAfterPx =
            Number(spacing?.attributes?.["w:after"] ?? 0) / TWIPS_PER_PX;
        }
      }
    } catch {
      // Ignore malformed styles.
    }
    return map;
  }

  headingLevel(styleId: string | undefined): number | undefined {
    if (!styleId) return undefined;
    return this.headingById.get(styleId);
  }

  /** True when styles.xml names this style, so its role is known, not guessed. */
  isDeclared(styleId: string | undefined): boolean {
    return styleId !== undefined && this.declared.has(styleId);
  }
}
