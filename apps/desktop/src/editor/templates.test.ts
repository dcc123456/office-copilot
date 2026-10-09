import { describe, expect, it } from "vitest";
import {
  cssFontSizeToPt,
  cssLetterSpacingToPt,
  ptToCssLetterSpacing,
} from "@office-copilot/docx-io";
import { planEdit, revert, type Block } from "../ai/docEdit.js";
import { templateProposal } from "./formatDocument.js";
import {
  TEMPLATES,
  completeRoles,
  mmToPx,
  templateById,
  type TemplateRole,
} from "./templates.js";
import { buildFooterDoc, buildHeaderDoc } from "./templateChrome.js";
import { TextStyle } from "./extensions/TextStyle.js";

type Json = Record<string, unknown>;

const heading = (index: number, text: string, level: number): Block => ({
  index,
  type: "heading",
  text,
  style: { headingLevel: level },
});
const para = (index: number, text: string, style: Partial<Block["style"]> = {}, mixed?: Block["mixed"]): Block => ({
  index,
  type: "paragraph",
  text,
  style,
  ...(mixed ? { mixed } : {}),
});

/** Title + one of each role + the blocks a template must not touch. */
function fixture(): Block[] {
  return [
    heading(0, "关于开展年度检查的通知", 1),
    para(1, "各有关单位："),
    heading(2, "一、检查范围", 1),
    para(3, "正文内容说明。"),
    heading(4, "（一）重点领域", 2),
    para(5, "手改过字号的正文。", { fontSize: "20px" }),
    para(6, "段内字号不一致的正文。", {}, ["fontSize"]),
    heading(7, "1. 细项", 3),
    heading(8, "1.1 更细项", 4),
    heading(9, "1.1.1 细节", 5),
    heading(10, "备注栏", 6),
    { index: 11, type: "table", text: "a | b", rows: [["a", "b"]], style: {} },
    { index: 12, type: "image", text: "配图", src: "x", style: {} },
    para(13, "   "),
  ];
}

const NON_TEXTUAL = [11, 12];
const ALL_ROLES: TemplateRole[] = ["title", "h1", "h2", "h3", "h4", "h5", "h6", "body"];

describe("templateProposal", () => {
  it("every template is idempotent", () => {
    for (const t of TEMPLATES) {
      const blocks = fixture();
      const first = templateProposal(t, blocks);
      expect(first.ops, `${t.id}: 应当产生调整`).not.toHaveLength(0);
      const applied = planEdit(blocks, first);
      expect(applied.ok, `${t.id}: 提案应可执行`).toBe(true);
      expect(templateProposal(t, applied.after).ops, `${t.id}: 二次应用应为空`).toEqual([]);
    }
  });

  it("only emits style ops, and never for tables, images or empty paragraphs", () => {
    for (const t of TEMPLATES) {
      const blocks = fixture();
      const p = templateProposal(t, blocks);
      for (const op of p.ops) {
        expect(op.op).toBe("setStyle");
        if ("index" in op) expect(NON_TEXTUAL).not.toContain(op.index);
      }
      expect(p.ops.map((o) => "index" in o ? o.index : -1)).not.toContain(13);
    }
  });

  it("prunes the run-level keys a block disagrees on instead of losing the block", () => {
    for (const t of TEMPLATES) {
      const blocks = fixture();
      const p = templateProposal(t, blocks);
      const mixed = blocks[6]!;
      const op = p.ops.find((o) => o.op === "setStyle" && o.index === 6);
      if (mixed.mixed?.length && op && op.op === "setStyle") {
        for (const key of mixed.mixed) expect(op.style[key]).toBeUndefined();
      }
      const plan = planEdit(blocks, p);
      expect(plan.skipped, t.id).toEqual([]);
    }
  });

  it("undoes as one modification and never rewrites text", () => {
    for (const t of TEMPLATES) {
      const blocks = fixture();
      const plan = planEdit(blocks, templateProposal(t, blocks));
      const back = revert(plan.after, plan.inverses);
      expect(back.ok, t.id).toBe(true);
      expect(back.after.map((b) => b.style)).toEqual(blocks.map((b) => b.style));
      expect(back.after.map((b) => b.text)).toEqual(blocks.map((b) => b.text));
    }
  });

  it("keeps the document's own heading levels", () => {
    for (const t of TEMPLATES) {
      const blocks = fixture();
      const plan = planEdit(blocks, templateProposal(t, blocks));
      expect(plan.after.map((b) => b.style.headingLevel)).toEqual(
        blocks.map((b) => b.style.headingLevel),
      );
    }
  });

  it("writes 字间距 in a value OOXML's twentieths read back unchanged", () => {
    // A 字间距 written as 0.5 磅 lands on 10 twips; reading it back must give
    // the very string the template stores, or every save changes the document.
    for (const t of TEMPLATES) {
      for (const role of ALL_ROLES) {
        const value = completeRoles(t)[role].letterSpacing;
        if (!value) continue;
        const pt = cssLetterSpacingToPt(value);
        expect(pt, `${t.id}/${role}`).not.toBeNull();
        expect(ptToCssLetterSpacing(pt as number), `${t.id}/${role}`).toBe(value);
      }
    }
  });

  it("gives every role of a 分字体模板 its 西文字体", () => {
    for (const t of TEMPLATES) {
      for (const role of ALL_ROLES) {
        const patch = completeRoles(t)[role];
        if (!t.latin) expect(patch.fontFamilyWest, `${t.id}/${role}`).toBeUndefined();
        else expect(patch.fontFamilyWest, `${t.id}/${role}`).toBe(t.latin);
      }
    }
  });

  it("reaches every level 1 to 6, the written ones and the derived ones", () => {
    const LEVELS: TemplateRole[] = ["h1", "h2", "h3", "h4", "h5", "h6"];
    for (const t of TEMPLATES) {
      const roles = completeRoles(t);
      const pts = LEVELS.map(
        (r) => cssFontSizeToPt(roles[r].fontSize ?? "") ?? 0,
      );
      pts.forEach(
        (p, i) => expect(p, `${t.id}: ${LEVELS[i]}`).toBeGreaterThan(0),
      );
      // Deepening never grows the type.
      for (let i = 1; i < pts.length; i++) {
        expect(
          pts[i]! <= pts[i - 1]!,
          `${t.id}: ${LEVELS[i]} 比 ${LEVELS[i - 1]} 大`,
        ).toBe(true);
      }
      const blocks = fixture();
      const proposal = templateProposal(t, blocks);
      for (const [i, index] of [8, 9, 10].entries()) {
        expect(
          proposal.ops.some((o) => o.op === "setStyle" && o.index === index),
          `${t.id}: H${i + 4} 应有样式可套`,
        ).toBe(true);
      }
    }
  });
});

describe("模板单位换算", () => {
  it("turns millimetre margins into 96dpi px", () => {
    expect(mmToPx(37)).toBeCloseTo(139.8, 1);
    expect(mmToPx(25.4)).toBeCloseTo(96, 1);
  });
});

describe("页眉页脚构造", () => {
  it("公文页脚是「— 页码 —」，无页码要求的模板不产出页脚", () => {
    const dash = buildFooterDoc(templateById("gongwen")!) as Json;
    expect(dash.type).toBe("doc");
    const p = (dash.content as Json[])[0]!;
    expect((p.attrs as Json).textAlign).toBe("center");
    const runs = (p.content as Json[]).filter((n) => n.type === "text");
    expect(runs.map((n) => n.text).join("")).toBe("—  —");
    const field = (p.content as Json[]).find((n) => n.type === "field")!;
    expect((field.attrs as Json).instr).toBe("PAGE");
    expect(buildFooterDoc(templateById("wechat")!)).toBeNull();
  });

  it("页眉只在模板声明且文档有标题时写入标题", () => {
    const minutes = templateById("minutes")!;
    expect(buildHeaderDoc(minutes, "   ")).toBeNull();
    const h = buildHeaderDoc(minutes, "周会纪要") as Json;
    expect(((h.content as Json[])[0]!.content as Json[])[0]!.text).toBe("周会纪要");
    expect(buildHeaderDoc(templateById("gongwen")!, "x")).toBeNull();
  });
});

/**
 * The colour idempotency lock. vitest runs without a DOM here, so the element
 * the browser would hand back is modelled directly: `data-*` survive verbatim,
 * while an inline `color: #c00000` comes out of CSSOM as `rgb(192, 0, 0)`. That
 * rewrite is exactly what made "already this colour?" answer wrongly and let the
 * same template keep producing no-op edits.
 */
describe("模板色的 DOM 往返", () => {
  type AttrSpec = {
    parseHTML: (el: HTMLElement) => unknown;
    renderHTML: (attrs: Record<string, unknown>) => Record<string, string>;
  };
  const attrSpecs = (
    TextStyle.config.addAttributes as unknown as () => Record<string, AttrSpec>
  )();

  const ROLES: TemplateRole[] = ["title", "h1", "h2", "h3", "body"];

  /** What TextStyle would put on the span for these mark attrs. */
  function htmlFor(attrs: Record<string, string>): Record<string, string> {
    let rendered: Record<string, string> = {};
    for (const [key, value] of Object.entries(attrs)) {
      rendered = { ...rendered, ...(attrSpecs[key]?.renderHTML({ [key]: value }) ?? {}) };
    }
    const spec = (TextStyle.config.renderHTML as unknown as (
      this: unknown,
      arg: unknown,
    ) => [string, Record<string, string>]).call(
      { options: { HTMLAttributes: {} } },
      { mark: { attrs }, HTMLAttributes: rendered },
    );
    const style = spec[1].style;
    return style ? { ...rendered, style } : rendered;
  }

  /** A browser-shaped element: attributes verbatim, colours as CSSOM reports them. */
  function domLike(html: Record<string, string>): HTMLElement {
    const style: Record<string, string> = {};
    for (const decl of (html.style ?? "").split(";")) {
      const i = decl.indexOf(":");
      if (i < 0) continue;
      const prop = decl.slice(0, i).trim();
      const value = decl.slice(i + 1).trim();
      const camel = prop.replace(/-([a-z])/g, (_, c: string) => c.toUpperCase());
      const hex = /^#([0-9a-f]{6})$/i.exec(value)?.[1];
      style[camel] = hex
        ? `rgb(${parseInt(hex.slice(0, 2), 16)}, ${parseInt(hex.slice(2, 4), 16)}, ${parseInt(hex.slice(4, 6), 16)})`
        : value;
    }
    return {
      getAttribute: (name: string) => html[name] ?? null,
      style,
    } as unknown as HTMLElement;
  }

  function reparse(key: string, html: Record<string, string>): string | null {
    return attrSpecs[key]!.parseHTML(domLike(html)) as string | null;
  }

  it("renders each colour through data-* as well as CSS", () => {
    const html = htmlFor({ color: "#c00000", backgroundColor: "#ffff00" });
    expect(html["data-color"]).toBe("#c00000");
    expect(html["data-background-color"]).toBe("#ffff00");
    expect(html.style).toContain("color: #c00000");
  });

  it("reads its own HTML back unchanged, rgb() rewrite and all", () => {
    for (const t of TEMPLATES) {
      for (const role of ROLES) {
        const color = t.roles[role]?.color;
        if (!color) continue;
        expect(reparse("color", htmlFor({ color })), `${t.id}/${role}`).toBe(color);
        expect(reparse("color", htmlFor({ color: color.toUpperCase() })), t.id).toBe(color);
      }
    }
  });

  it("canonicalises pasted HTML that carries no data-*", () => {
    expect(reparse("color", { style: "color: rgb(192, 0, 0)" })).toBe("#c00000");
    expect(reparse("backgroundColor", { style: "background-color: rgb(255, 255, 0)" })).toBe(
      "#ffff00",
    );
  });

  it("a document re-read through the DOM still takes no further edits", () => {
    for (const t of TEMPLATES) {
      const blocks = fixture();
      const plan = planEdit(blocks, templateProposal(t, blocks));
      const asDomRead = plan.after.map((b) =>
        b.style.color
          ? { ...b, style: { ...b.style, color: reparse("color", htmlFor({ color: b.style.color })) } }
          : b,
      );
      expect(templateProposal(t, asDomRead).ops, `${t.id}: 往返后二次应用应为空`).toEqual([]);
    }
  });
});
