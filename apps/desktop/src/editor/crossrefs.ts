import type { Editor } from "@tiptap/react";
import type { Node as PMNode } from "@tiptap/pm/model";
import { addressOf, runOps, type CoreOp, type JsonNode } from "../ai/docEdit.js";
import { commitOps, type CommitResult } from "../ai/agent.js";
import { blocksOf, locateAll } from "./docBlocks.js";
import { pageNumberOf } from "./extensions/PagedLayout.js";

/**
 * 交叉引用 = 一条指向书签的 `REF`/`PAGEREF` 域。开关以官方为准（微软「字段代码：Ref 域」：
 * `{ REF 书签名 [开关] }`，该文列出的开关只有 `\d \f \h \n \p \r \t \w`）：不带开关取
 * "书签所代表的那段文字"，`\p` 给"above/below"，`\h` 表示更新域时做成超链接；
 * 页码走 `PAGEREF`（目录条目那一条 ③.2 已经在写它）。
 *
 * 目标就是本文档里已有的书签：目录条目的 `_Toc…` 包住标题文字，题注的 `_Ref…` 只包住
 * `SEQ` 域（⑤.1），所以"书签文字"对题注来说就是编号——这与 Word 的列法一致。
 * 域与书签的结构化读写早在 ②/③.1 打通，本文件只负责挑目标、算缓存结果、发 op。
 */

/** Word「交叉引用」对话框里"引用类型"那一列，按书签名前缀分档。 */
export type RefTargetKind = "caption" | "toc" | "bookmark";

/** 引用内容：`text` → `{ REF X }`，`page` → `{ PAGEREF X }`，`position` → `{ REF X \p }`。 */
export type RefWhat = "text" | "page" | "position";

export interface CrossRefOptions {
  /** 目标书签名，例如 `_Ref2`。 */
  target: string;
  what: RefWhat;
  /** 「插入为超链接」：只往指令里加 `\h`，不手写 `w:hyperlink`（同 ③.2 的取舍）。 */
  asHyperlink: boolean;
}

/** 一个可被引用的书签。 */
export interface RefTarget {
  name: string;
  kind: RefTargetKind;
  /** 书签实际包住的文字：题注是编号，目录条目是标题文字；域按缓存结果计入。 */
  covered: string;
  /** 所在段的纯文本，只给对话框当列表项。 */
  blockText: string;
  /** 所在段在文档里的顶层坐标，算页码与"以上/以下"都跟它比。 */
  blockPos: number;
}

/** 引擎只给码，中文文案在对话框里拼（AGENTS §5.7）。 */
export type CrossRefReject = "no-targets" | "bad-target" | "no-cursor" | "ops-failed";

export interface CrossRefPlan {
  ok: boolean;
  reason?: CrossRefReject;
  ops: CoreOp[];
  before: ReturnType<typeof blocksOf>;
  after: ReturnType<typeof blocksOf>;
  inverses: CoreOp[];
  /** 写进域指令的那串字，成功与否都有，供界面与修改历史显示。 */
  instr: string;
  /** 域的缓存结果（Word 打开后按 F9 会重算）。 */
  display: string;
  skipped: string[];
}

/**
 * 书签名前缀 → Word 对话框里的引用类型档位。
 *
 * @param name 书签名。
 * @returns `caption`（题注编号）/ `toc`（目录条目）/ `bookmark`（用户自己起的书签）。
 */
function refKindOf(name: string): RefTargetKind {
  if (name.startsWith("_Ref")) return "caption";
  if (name.startsWith("_Toc")) return "toc";
  return "bookmark";
}

/**
 * 行内节点为"书签所代表的文字"贡献多少。
 *
 * 原子节点没有 `textContent`，但域在 OOXML 里被书签包住的那段就是它的**缓存结果**，
 * 所以这里按 `display` 计——题注的 `_Ref…` 包住的是 `SEQ` 域，取到的是编号而不是空串。
 *
 * @param node 段落里的一个行内节点。
 * @returns 该节点在书签内的字面文字。
 */
function inlineCoveredText(node: PMNode): string {
  if (node.isText) return node.text ?? "";
  if (node.type.name === "field") return String(node.attrs.display ?? "");
  return "";
}

/**
 * 列出文档里所有可作为交叉引用目标的**完整**书签（`bookmarkStart` 与 `bookmarkEnd` 同段成对）。
 *
 * @param doc 当前文档。
 * @returns 按文档顺序；落单的半边直接跳过，Word 的对话框也不列没有终点的书签。
 */
export function collectRefTargets(doc: PMNode): RefTarget[] {
  const out: RefTarget[] = [];
  doc.forEach((block, blockPos) => {
    const open = new Map<string, { name: string; parts: string[] }>();
    // 书签不跨段：① 的读侧把跨段的那种折成段边界，这里只认同一行内层里的成对锚点。
    block.forEach((child) => {
      const id = String(child.attrs.id ?? "");
      if (child.type.name === "bookmarkStart") {
        open.set(id, { name: String(child.attrs.name ?? ""), parts: [] });
        return;
      }
      if (child.type.name === "bookmarkEnd") {
        const entry = open.get(id);
        if (!entry) return;
        open.delete(id);
        if (!entry.name) return;
        out.push({
          name: entry.name,
          kind: refKindOf(entry.name),
          covered: entry.parts.join(""),
          blockText: block.textContent,
          blockPos,
        });
        return;
      }
      for (const entry of open.values()) entry.parts.push(inlineCoveredText(child));
    });
  });
  return out;
}

/**
 * 生成域指令。
 *
 * @param name 目标书签名。
 * @param what 引用内容档位。
 * @param asHyperlink 是否加 `\h`（Word 更新域时据此生成超链接）。
 * @returns 写进 `w:instrText` 的指令文本，例如 `PAGEREF _Ref2 \h`。
 */
export function refInstr(name: string, what: RefWhat, asHyperlink: boolean): string {
  const base =
    what === "page" ? `PAGEREF ${name}` : what === "position" ? `REF ${name} \\p` : `REF ${name}`;
  return asHyperlink ? `${base} \\h` : base;
}

/**
 * `\p` 的缓存结果：官方语义是"域在书签之前算 below、在之后算 above"。
 *
 * @param fieldPos 插入点（文档坐标）。
 * @param targetPos 目标书签所在段的顶层坐标。
 * @returns 中文界面下的字面结果；Word 更新域时按它自己的界面语言重算。
 */
export function aboveOrBelow(fieldPos: number, targetPos: number): string {
  return fieldPos < targetPos ? "以下" : "以上";
}

/**
 * 交叉引用的域节点 JSON。
 *
 * @param instr {@link refInstr} 生成的指令文本。
 * @param display 域的缓存结果；空串表示本地还算不出来（例如分页没测到页码）。
 * @returns `field` 行内原子节点的 JSON。
 */
export function crossRefJson(instr: string, display: string): JsonNode {
  return {
    type: "field",
    attrs: { kind: instr.startsWith("PAGEREF") ? "PAGEREF" : "REF", instr, display },
  };
}

/**
 * 生成一次交叉引用插入的 op，并在块模型上模拟——与 `planCaptionInsertion` 同一套语义：
 * 模拟与实际不一致就整批放弃（AGENTS §2.4）。
 *
 * @param editor 桌面编辑器，插入点是当前选区。
 * @param opts   对话框选项。
 */
export function planCrossReference(editor: Editor, opts: CrossRefOptions): CrossRefPlan {
  // 页码取自本地分页（`PagedLayout`），与目录那条路径同一个来源，不再另起一处测量。
  const pageOf = (pos: number) => pageNumberOf(editor.state, pos);
  const state = editor.state;
  const before = blocksOf(state.doc);
  const empty: CrossRefPlan = {
    ok: false,
    ops: [],
    before,
    after: before,
    inverses: [],
    instr: "",
    display: "",
    skipped: [],
  };

  const targets = collectRefTargets(state.doc);
  if (!targets.length) return { ...empty, reason: "no-targets" };
  const target = targets.find((t) => t.name === opts.target);
  if (!target) return { ...empty, reason: "bad-target" };

  const cursor = state.selection.from;
  const host = locateAll(state.doc).find((l) => l.from <= cursor && cursor < l.to);
  const inner = host?.inner;
  if (!host || !inner || !host.textNode) return { ...empty, reason: "no-cursor" };

  const instr = refInstr(target.name, opts.what, opts.asHyperlink);
  const display =
    opts.what === "text"
      ? target.covered
      : opts.what === "page"
        ? String(pageOf(target.blockPos) || "")
        : aboveOrBelow(cursor, target.blockPos);

  // 插入点换成块模型的文字偏移：原子节点不计数，所以偏移不会落在域芯片中间。
  const clamped = Math.max(inner.from, Math.min(cursor, inner.to));
  const at = state.doc.textBetween(inner.from, clamped, undefined, "").length;
  const addr = addressOf(before, host.block.index);
  if (!addr) return { ...empty, reason: "no-cursor" };

  const ops: CoreOp[] = [{ kind: "insertInline", addr, at, json: crossRefJson(instr, display) }];
  const simulated = runOps(before, ops);
  if (simulated.skipped.length) {
    return { ...empty, reason: "ops-failed", instr, display, skipped: simulated.skipped };
  }
  return {
    ok: true,
    ops: simulated.applied,
    before,
    after: simulated.after,
    inverses: simulated.inverses,
    instr,
    display,
    skipped: [],
  };
}

/**
 * 落地一次交叉引用插入，并写进修改历史（逆操作按指令文本认域，§2.5）。
 *
 * @param editor 桌面编辑器。
 * @param docKey 修改历史的文档键，即文件路径。
 * @param opts   对话框选项。
 * @returns 成功时带 modlog 记录与 `instr`；失败时 `reason` 是拒绝码，文案由界面组装。
 */
export async function insertCrossRef(
  editor: Editor,
  docKey: string,
  opts: CrossRefOptions,
): Promise<CommitResult & { instr: string; reason?: CrossRefReject }> {
  const plan = planCrossReference(editor, opts);
  if (!plan.ok) return { ok: false, instr: plan.instr, reason: plan.reason };
  const r = await commitOps(editor, {
    ops: plan.ops,
    before: plan.before,
    after: plan.after,
    inverses: plan.inverses,
    label: "插入交叉引用",
    summary: `插入交叉引用（${plan.instr}）`,
    docKey,
  });
  return { ...r, instr: plan.instr };
}
