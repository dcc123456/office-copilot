import { ptToCssFontSize } from "@office-copilot/docx-io";
import type { DocTemplate } from "./templates.js";

/**
 * Header/footer documents a template asks for, in the same TipTap JSON shape
 * HeaderFooterEditor.tsx reads and writes: paragraph → text runs with a
 * `textStyle` mark, plus an inline `field` node for the page number
 * (attrs per extensions/Field.ts). The pagination layer resolves PAGE live.
 *
 * These are part-level content, not block styles, so they go through the same
 * channel 页面设置 uses (App state + dirty flag) rather than the undoable op log.
 */

type Json = Record<string, unknown>;

const SONG = "宋体";

function run(text: string, pt: number): Json {
  return {
    type: "text",
    text,
    marks: [{ type: "textStyle", attrs: { fontFamily: SONG, fontSize: ptToCssFontSize(pt) } }],
  };
}

function pageField(): Json {
  return { type: "field", attrs: { kind: "page", instr: "PAGE", display: "1" } };
}

function paragraph(children: Json[], align: "center" | "right"): Json {
  return { type: "paragraph", attrs: { textAlign: align }, content: children };
}

function doc(block: Json): Json {
  return { type: "doc", content: [block] };
}

/** 公文页码是「— 1 —」四号宋体，其余模板用居中小五页码。 */
export function buildFooterDoc(t: DocTemplate): Json | null {
  switch (t.chrome?.footer) {
    case "dash":
      return doc(paragraph([run("— ", 14), pageField(), run(" —", 14)], "center"));
    case "center":
      return doc(paragraph([pageField()], "center"));
    default:
      return null;
  }
}

/** 页眉写文档标题（第一行标题的文字），没有标题就不写。 */
export function buildHeaderDoc(t: DocTemplate, docTitle: string): Json | null {
  if (t.chrome?.header !== "title") return null;
  const label = docTitle.trim();
  if (!label) return null;
  return doc(paragraph([run(label, 9)], "center"));
}
