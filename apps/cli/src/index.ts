#!/usr/bin/env node
import { Command } from "commander";
import * as p from "@clack/prompts";
import pc from "picocolors";
import fs from "node:fs";
import path from "node:path";
import { loadDocx, saveDoc, log, warn, fail } from "./fsutil.js";
import { runAgent } from "./agent.js";
import type { ProposedEdit } from "@office-copilot/shared";

const program = new Command();
program
  .name("oc")
  .description("Office 智能助手 CLI —— 直接修改 .docx，Word 和 WPS 都能打开。")
  .version("0.1.0");

program
  .command("list <file>")
  .description("列出文档段落（index: style: text）")
  .action(async (file: string) => {
    const doc = await loadDocx(file);
    for (const para of doc.file.listParagraphs()) {
      const head = para.text.length > 120 ? para.text.slice(0, 120) + "…" : para.text;
      log(
        `${pc.dim(String(para.index).padStart(4, " "))} ${pc.yellow(
          para.style ?? "Normal",
        )}  ${head}`,
      );
    }
  });

program
  .command("replace <file>")
  .description("在文档里替换文字（字面量或正则）")
  .requiredOption("-f, --find <text>", "要查找的文字")
  .requiredOption("-r, --replace <text>", "替换成的文字")
  .option("--regex", "把 --find 当作 JavaScript 正则表达式")
  .option("-o, --output <path>", "另存为新文件（默认原地修改并生成 .bak）")
  .option("--no-backup", "原地修改时不生成 .bak（危险）")
  .action(async (file: string, opts: { find: string; replace: string; regex?: boolean; output?: string; backup?: boolean }) => {
    const doc = await loadDocx(file);
    const res = doc.file.replaceText({ find: opts.find, replace: opts.replace, regex: opts.regex });
    if (res.count === 0) {
      warn(`未找到：${opts.find}`);
      process.exit(2);
    }
    for (const [before, after] of res.previews) {
      log(`${pc.red("- " + before)}\n${pc.green("+ " + after)}`);
    }
    log(pc.dim(`共 ${res.count} 处替换。`));
    const out = await saveDoc(doc, opts.output, { backup: opts.backup });
    log(pc.green(`✓ 已保存：${out}`));
  });

program
  .command("insert-paragraph <file>")
  .description("插入一个段落")
  .requiredOption("-t, --text <text>", "段落文字（用 \\n 表示换行）")
  .option("-p, --position <pos>", "start|end|before|after", "end")
  .option("-a, --anchor <index>", "before/after 时的段落下标（0 起）", (v) => parseInt(v, 10))
  .option("-s, --style <name>", "Word 样式名，例如 Heading 1")
  .option("-o, --output <path>", "另存为新文件")
  .option("--no-backup", "原地修改时不生成 .bak")
  .action(async (file: string, opts: { text: string; position: "start"|"end"|"before"|"after"; anchor?: number; style?: string; output?: string; backup?: boolean }) => {
    if ((opts.position === "before" || opts.position === "after") && opts.anchor == null) {
      fail(`position=${opts.position} 需要 --anchor <index>`);
    }
    const doc = await loadDocx(file);
    const text = opts.text.replace(/\\n/g, "\n");
    doc.file.insertParagraph({ text, position: opts.position, anchor: opts.anchor, style: opts.style });
    const out = await saveDoc(doc, opts.output, { backup: opts.backup });
    log(pc.green(`✓ 已插入段落并保存：${out}`));
  });

program
  .command("insert-table <file>")
  .description("插入一个二维表格（用逗号或 JSON 传入）")
  .requiredOption(
    "-d, --data <json>",
    "二维 JSON 数组，例如 '[\"A,B\",\"1,2\"]' 或直接给 JSON",
  )
  .option("-p, --position <pos>", "start|end", "end")
  .option("--no-header", "首行不加粗")
  .option("-o, --output <path>", "另存为新文件")
  .option("--no-backup", "原地修改时不生成 .bak")
  .action(async (file: string, opts: { data: string; position: "start"|"end"; header?: boolean; output?: string; backup?: boolean }) => {
    const rows = parseRows(opts.data);
    const doc = await loadDocx(file);
    doc.file.insertTable({ rows, headerRow: opts.header ?? true, position: opts.position });
    const out = await saveDoc(doc, opts.output, { backup: opts.backup });
    log(pc.green(`✓ 已插入 ${rows.length} 行表格并保存：${out}`));
  });

program
  .command("insert-image <file> <image>")
  .description("在文档末尾插入一张图片（PNG/JPEG）")
  .option("-w, --width <px>", "宽度（像素，96 DPI 换算）", (v) => parseInt(v, 10))
  .option("-h, --height <px>", "高度（像素）", (v) => parseInt(v, 10))
  .option("-o, --output <path>", "另存为新文件")
  .option("--no-backup", "原地修改时不生成 .bak")
  .action(async (file: string, image: string, opts: { width?: number; height?: number; output?: string; backup?: boolean }) => {
    const doc = await loadDocx(file);
    const buf = await fs.promises.readFile(image);
    const ext = path.extname(image).toLowerCase();
    const contentType = ext === ".png" ? "image/png" : ext === ".jpg" || ext === ".jpeg" ? "image/jpeg" : null;
    if (!contentType) fail("只支持 PNG/JPEG。");
    await doc.file.insertImage({
      bytes: new Uint8Array(buf),
      contentType,
      filename: path.basename(image),
      widthPx: opts.width,
      heightPx: opts.height,
    });
    const out = await saveDoc(doc, opts.output, { backup: opts.backup });
    log(pc.green(`✓ 已插入图片并保存：${out}`));
  });

program
  .command("ai <file> [instruction]")
  .description("用自然语言修改文档（需要 LLM_API_KEY）")
  .option("-y, --yes", "跳过确认，直接应用提案")
  .option("-o, --output <path>", "另存为新文件")
  .option("--no-backup", "原地修改时不生成 .bak")
  .action(async (file: string, instruction: string | undefined, opts: { yes?: boolean; output?: string; backup?: boolean }) => {
    let prompt = instruction;
    if (!prompt) {
      const input = await p.text({
        message: "告诉助手怎么改这份文档：",
        placeholder: "例如：把所有'甲方'改成'客户'，并在末尾加一个签字栏表格",
        validate: (v) => (v.trim().length > 0 ? undefined : "请输入指令"),
      });
      if (p.isCancel(input)) process.exit(130);
      prompt = input as string;
    }

    const doc = await loadDocx(file);
    const paragraphs = doc.file.listParagraphs();

    const spin = p.spinner();
    spin.start("请求模型生成提案…");
    let proposal: import("@office-copilot/shared").Proposal | null = null;
    try {
      proposal = await runAgent({
        instruction: prompt!,
        paragraphs,
        onStatus: (m) => {
          spin.message(m);
        },
      });
    } finally {
      spin.stop(proposal ? "提案就绪" : "未生成提案");
    }

    if (!proposal || proposal.edits.length === 0) {
      warn(proposal?.summary || "模型没有生成可执行的修改。");
      process.exit(3);
    }

    p.log.message(pc.bold("修改提案：") + proposal.summary);
    printEdits(proposal.edits);

    const apply =
      opts.yes === true
        ? true
        : await p.confirm({ message: "接受并写入文档？", initialValue: true });
    if (!apply || p.isCancel(apply)) {
      log("已取消。");
      process.exit(0);
    }

    for (const e of proposal.edits) {
      applyEdit(doc.file, e);
    }
    const out = await saveDoc(doc, opts.output, { backup: opts.backup });
    log(pc.green(`✓ 已写入：${out}`));
  });

program.parseAsync(process.argv).catch((e: unknown) => {
  fail(e instanceof Error ? e.message : String(e));
});

function parseRows(raw: string): string[][] {
  const trimmed = raw.trim();
  if (trimmed.startsWith("[")) {
    const data = JSON.parse(trimmed);
    if (!Array.isArray(data) || !data.every((r) => Array.isArray(r))) {
      fail("--data 必须是二维字符串数组。");
    }
    return (data as unknown[][]).map((r) => r.map(String));
  }
  return trimmed
    .split(/\n|;/)
    .map((line) => line.trim())
    .filter(Boolean)
    .map((line) => line.split(",").map((c) => c.trim()));
}

function printEdits(edits: ProposedEdit[]): void {
  for (const e of edits) {
    switch (e.type) {
      case "replace_text":
        p.log.message(`  ${pc.red("- " + e.find)}\n  ${pc.green("+ " + e.replace)}`);
        break;
      case "insert_paragraph":
        p.log.message(
          `  ${pc.green("+ 段落(" + e.position + (e.anchor != null ? "@" + e.anchor : "") + ")")}: ${truncate(e.text, 100)}`,
        );
        break;
      case "insert_table":
        p.log.message(`  ${pc.green(`+ 表格 ${e.rows.length}×${e.rows[0]?.length ?? 0}`)}`);
        for (const row of e.rows) p.log.message("    " + row.join(" | "));
        break;
      case "insert_image":
        p.log.message(`  ${pc.green("+ 图片")}`);
        break;
    }
  }
}

function truncate(s: string, n: number): string {
  return s.length > n ? s.slice(0, n) + "…" : s;
}

function applyEdit(doc: import("@office-copilot/docx-io").DocxFile, edit: ProposedEdit): void {
  switch (edit.type) {
    case "replace_text":
      doc.replaceText({ find: edit.find, replace: edit.replace, regex: edit.regex });
      break;
    case "insert_paragraph":
      doc.insertParagraph({
        text: edit.text,
        position: edit.position,
        anchor: edit.anchor,
        style: edit.style,
      });
      break;
    case "insert_table":
      doc.insertTable({ rows: edit.rows, headerRow: edit.headerRow, position: "end" });
      break;
    case "insert_image":
      // CLI AI path doesn't emit insert_image edits; the standalone
      // insert-image command handles it directly.
      warn("AI 提案暂不支持插入图片，请使用 `oc insert-image` 命令。");
      break;
  }
}
