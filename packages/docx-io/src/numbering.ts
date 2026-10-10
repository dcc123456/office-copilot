import { children, firstChild, parseDocument, val } from "./xml.js";
import type { Element } from "./xml.js";

/**
 * Parse word/numbering.xml so we can tell whether a <w:numPr> paragraph is
 * ordered (decimal/lowerLetter/...) or bulleted, and carry the numbering
 * through a round-trip.
 *
 * We do not need to render the numbers/bullets ourselves; we just need to
 * know which abstract level is bullet vs. ordered so the ProseMirror
 * conversion picks bulletList vs orderedList.
 *
 * For round-tripping we keep the original numbering.xml part untouched and
 * reference its numId from paragraphs on the way back out.
 */

export type NumberingKind = "bullet" | "ordered";

export interface NumberingLevel {
  kind: NumberingKind;
  /** numFmt value, e.g. "decimal", "lowerLetter", "bullet". */
  numFmt: string;
  ilvl: number;
  abstractNumId: string;
}

export class NumberingMap {
  /** numId -> abstractNumId */
  private numToAbstract = new Map<string, string>();
  /** "abstractNumId|ilvl" -> level info */
  private levels = new Map<string, NumberingLevel>();
  /** Abstract IDs we created on save (so we can emit their definitions). */
  private emittedAbstractIds = new Set<string>();

  static fromXml(xml: string | undefined): NumberingMap {
    const map = new NumberingMap();
    if (!xml) return map;
    const root = parseDocument(xml);
    for (const num of children(root).filter((c) => c.name === "w:num")) {
      const id = num.attributes?.["w:numId"] as string | undefined;
      const abstract = firstChild(num, "w:abstractNumId");
      const abstractId = abstract ? val(abstract) : undefined;
      if (id && abstractId) map.numToAbstract.set(id, abstractId);
    }
    for (const abs of children(root).filter((c) => c.name === "w:abstractNum")) {
      const absId = abs.attributes?.["w:abstractNumId"] as string | undefined;
      if (!absId) continue;
      for (const lvl of children(abs).filter((c) => c.name === "w:lvl")) {
        const ilvl = Number(lvl.attributes?.["w:ilvl"] ?? 0);
        const numFmtEl = firstChild(lvl, "w:numFmt");
        const numFmt = val(numFmtEl) ?? "decimal";
        const kind: NumberingKind = numFmt === "bullet" ? "bullet" : "ordered";
        map.levels.set(`${absId}|${ilvl}`, {
          kind,
          numFmt,
          ilvl,
          abstractNumId: absId,
        });
      }
    }
    return map;
  }

  level(numId: string | undefined, ilvl: number): NumberingLevel | undefined {
    if (!numId) return undefined;
    const absId = this.numToAbstract.get(numId);
    if (!absId) return undefined;
    return this.levels.get(`${absId}|${ilvl}`);
  }

  /**
   * Register a new (or reused) numbering definition when saving. We prefer
   * to reuse the document's existing bullet/decimal abstractNumId when
   * present so that Word/WPS render with the same glyph set.
   *
   * Returns the numId to reference; caller must ensure a corresponding
   * <w:abstractNum> + <w:num> exists in the emitted numbering.xml.
   */
  ensureFor(kind: NumberingKind): { numId: string; abstractNumId: string } | null {
    for (const [, lvl] of this.levels) {
      if (lvl.kind === kind) {
        const numIdEntry = [...this.numToAbstract.entries()].find(
          ([, abs]) => abs === lvl.abstractNumId,
        );
        if (numIdEntry) {
          this.emittedAbstractIds.add(lvl.abstractNumId);
          return { numId: numIdEntry[0], abstractNumId: lvl.abstractNumId };
        }
      }
    }
    // Nothing from the source — fall back to the synthetic defaults.
    const abstractNumId = kind === "ordered" ? "1001" : "1000";
    const numId = kind === "ordered" ? "1001" : "1000";
    if (!this.numToAbstract.has(numId)) {
      this.numToAbstract.set(numId, abstractNumId);
      for (let ilvl = 0; ilvl <= 8; ilvl++) {
        this.levels.set(`${abstractNumId}|${ilvl}`, {
          kind,
          numFmt: kind === "ordered" ? "decimal" : "bullet",
          ilvl,
          abstractNumId,
        });
      }
    }
    this.emittedAbstractIds.add(abstractNumId);
    return { numId, abstractNumId };
  }

  markEmitted(numId: string | undefined): void {
    if (!numId) return;
    const abs = this.numToAbstract.get(numId);
    if (abs) this.emittedAbstractIds.add(abs);
  }

  getEmittedAbstractIds(): string[] {
    return [...this.emittedAbstractIds];
  }

  hasAny(): boolean {
    return this.numToAbstract.size > 0;
  }

  /**
   * Build the complete <w:numbering> XML for the saved document. We emit
   * every abstractNum that any saved paragraph referenced plus, if the
   * source had no numbering at all, two default definitions so newly
   * inserted lists render sensibly in Word/WPS:
   *   - abstractNum 1000: bullet (• at level 0, ◦ at 1, ▪ at 2+)
   *   - abstractNum 1001: multi-level Chinese decimal:
   *       L0 1.   L1 1.1   L2 1.1.1   L3 (1)   L4 ①   L5 a)   L6 一、
   *
   * Numbers 1000+ are picked to avoid colliding with real document IDs.
   */
  toXml(existingNumberingXml: string | undefined): string {
    const existing = new Set(this.emittedAbstractIds);
    const extras: string[] = [];
    if (existing.size === 0) {
      extras.push(DEFAULT_BULLET_ABSTRACT, DEFAULT_DECIMAL_ABSTRACT);
      // Need num entries for them too.
      extras.push(`<w:num w:numId="1000"><w:abstractNumId w:val="1000"/></w:num>`);
      extras.push(`<w:num w:numId="1001"><w:abstractNumId w:val="1001"/></w:num>`);
    }

    // If the source already had numbering, preserve its body so we don't
    // lose abstractNum definitions we didn't model. We strip its closing
    // tag and append our extras.
    if (existingNumberingXml) {
      const body = existingNumberingXml
        .replace(/<\?xml[^?]*\?>/g, "")
        .replace(/<w:numbering\b[^>]*>/g, "")
        .replace(/<\/w:numbering>\s*$/g, "");
      return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${body}${extras.join("")}</w:numbering>`;
    }
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${extras.join("")}</w:numbering>`;
  }
}

const DEFAULT_BULLET_ABSTRACT = `
<w:abstractNum w:abstractNumId="1000">
  ${[0, 1, 2, 3, 4, 5, 6, 7, 8]
    .map((i) => {
      const glyph = i === 0 ? "•" : i === 1 ? "◦" : "▪";
      const font = "Symbol";
      const indent = 720 + i * 360;
      return `<w:lvl w:ilvl="${i}">
        <w:start w:val="1"/>
        <w:numFmt w:val="bullet"/>
        <w:lvlText w:val="${glyph}"/>
        <w:lvlJc w:val="left"/>
        <w:pPr><w:ind w:left="${indent}" w:hanging="360"/></w:pPr>
        <w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:hint="default"/></w:rPr>
      </w:lvl>`;
    })
    .join("")}
</w:abstractNum>`;

const DEFAULT_DECIMAL_ABSTRACT = `
<w:abstractNum w:abstractNumId="1001">
  ${[
    { fmt: "decimal", text: "%1.", jc: "left" },
    { fmt: "decimal", text: "%1.%2", jc: "left" },
    { fmt: "decimal", text: "%1.%2.%3", jc: "left" },
    { fmt: "decimalEnclosedParen", text: "(%4)", jc: "left" },
    { fmt: "decimalEnclosedCircle", text: "%5", jc: "left" },
    { fmt: "lowerLetter", text: "%6)", jc: "left" },
    { fmt: "chineseCounting", text: "%7、", jc: "left" },
    { fmt: "chineseCounting", text: "%7.%8", jc: "left" },
    { fmt: "decimal", text: "%9.", jc: "left" },
  ]
    .map((l, i) => {
      const indent = 720 + i * 360;
      return `<w:lvl w:ilvl="${i}">
        <w:start w:val="1"/>
        <w:numFmt w:val="${l.fmt}"/>
        <w:lvlText w:val="${l.text}"/>
        <w:lvlJc w:val="${l.jc}"/>
        <w:pPr><w:ind w:left="${indent}" w:hanging="360"/></w:pPr>
      </w:lvl>`;
    })
    .join("")}
</w:abstractNum>`;

export function readListProps(p: Element): {
  numId: string | undefined;
  ilvl: number;
} {
  const pPr = firstChild(p, "w:pPr");
  const numPr = pPr && firstChild(pPr, "w:numPr");
  if (!numPr) return { numId: undefined, ilvl: 0 };
  const numId = val(firstChild(numPr, "w:numId"));
  const ilvlRaw = val(firstChild(numPr, "w:ilvl"));
  return { numId, ilvl: ilvlRaw ? Number(ilvlRaw) : 0 };
}
