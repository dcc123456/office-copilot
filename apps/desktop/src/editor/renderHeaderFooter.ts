import type { PMDoc, PMMark, PMNode } from "@office-copilot/docx-io";

/**
 * Render a header/footer PMDoc as static HTML for a specific page,
 * resolving PAGE/NUMPAGES/DATE fields to concrete values. This is a
 * read-only preview shown on every page; the full TipTap editor is
 * only mounted while the user is editing the header or footer.
 */
export function renderHeaderFooterHtml(
  doc: PMDoc | { content: PMNode[] } | null | undefined,
  page: number,
  total: number,
): string {
  if (!doc || !doc.content) return "";
  const today = new Date().toLocaleDateString();
  const renderNode = (n: PMNode): string => {
    if (n.type === "paragraph") {
      const align = n.attrs?.textAlign as string | undefined;
      const style = align ? ` style="text-align:${align}"` : "";
      return `<p${style}>${(n.content ?? []).map(runHtml).join("")}</p>`;
    }
    return "";
  };
  const runHtml = (n: PMNode): string => {
    if (n.type === "text") {
      const t = escapeHtml((n as { text?: string }).text ?? "");
      const marks: PMMark[] = n.marks ?? [];
      const style = markStyle(marks);
      const tags = marks.map(inlineTagOf).filter((tag): tag is string => tag !== null);
      const open = tags.map((tag) => `<${tag}>`).join("");
      const close = tags
        .slice()
        .reverse()
        .map((tag) => `</${tag}>`)
        .join("");
      return `<span style="${style}">${open}${t}${close}</span>`;
    }
    if (n.type === "field") {
      const k = String(n.attrs.kind).toUpperCase();
      const val =
        k === "PAGE"
          ? String(page)
          : k === "NUMPAGES"
            ? String(total)
            : k === "DATE"
              ? today
              : n.attrs.display ?? k;
      return `<span class="pm-field-static">${escapeHtml(String(val))}</span>`;
    }
    if (n.type === "hardBreak") return "<br/>";
    return "";
  };
  return doc.content.map(renderNode).join("");
}

/**
 * 行内标记 → 语义标签名（不带尖括号），认不出的标记返回 null 表示不产标签。
 *
 * @param mark docx-io 的行内标记。
 * @returns 标签名；`vertAlign` 的上下两档分别给 `sup` / `sub`，因为它们是同一个标记的两个值。
 */
function inlineTagOf(mark: PMMark): string | null {
  switch (mark.type) {
    case "bold":
      return "b";
    case "italic":
      return "i";
    case "underline":
      return "u";
    case "strike":
      return "s";
    case "vertAlign":
      return mark.attrs.val === "subscript" ? "sub" : "sup";
    default:
      return null;
  }
}

function markStyle(marks: PMMark[]): string {
  const ts = marks.find(
    (m): m is Extract<PMMark, { type: "textStyle" }> => m.type === "textStyle",
  );
  if (!ts?.attrs) return "";
  const parts: string[] = [];
  const a = ts.attrs as Record<string, unknown>;
  if (a.fontFamily) parts.push(`font-family:${a.fontFamily}`);
  if (a.fontSize) {
    const s = String(a.fontSize);
    parts.push(`font-size:${/^\d+$/.test(s) ? s + "px" : s}`);
  }
  if (a.color) parts.push(`color:${a.color}`);
  if (a.backgroundColor) parts.push(`background:${a.backgroundColor}`);
  return parts.join(";");
}

function escapeHtml(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}
