/**
 * `oc pdf` —— CLI 端的 `.pdf` 结构体检（`docs/office-parity-spec.md` D8 的 `oc pdf check`）。
 *
 * 与 `oc ppt check` 同一条纪律：只做装配，不写第二套判定——结论来自 `@office-copilot/pdf-io` 的
 * `auditPdfStructure`，与桌面 app 走同一份引擎（`AGENTS.md` §3.3）。
 * 写侧刻意不在 CLI 落地：撤销 / modlog / 原子替换都在桌面主进程那条路上（§2.4/§2.5）。
 */
import { Command } from "commander";
import pc from "picocolors";
import fs from "node:fs";
import path from "node:path";
import { auditPdfStructure, type PdfAuditIssue } from "@office-copilot/pdf-io";
import { describePdfAuditIssue, summarizePdfAudit } from "@office-copilot/shared";
import { emit, fail, log, warn } from "./fsutil.js";

export function registerPdfCommands(program: Command): void {
  const pdf = program.command("pdf").description("读 .pdf：结构体检（与桌面同一份引擎，只读）");

  pdf
    .command("check <file>")
    .description("兼容性体检：这份 PDF 会不会被阅读器判损坏，以及有多少文字改不了")
    .option("--json", "输出 JSON")
    .action(async (file: string, opts: { json?: boolean }) => {
      const abs = path.resolve(file);
      if (!fs.existsSync(abs)) fail(`文件不存在：${abs}`);
      let issues: PdfAuditIssue[];
      try {
        issues = await auditPdfStructure(new Uint8Array(await fs.promises.readFile(abs)));
      } catch (error) {
        fail(error instanceof Error ? error.message : String(error));
      }
      emit(opts.json, { fileName: path.basename(abs), ok: issues.length === 0, issues }, () => {
        for (const issue of issues) log(`  ${issue.severity === "broken" ? pc.red(describePdfAuditIssue(issue)) : pc.yellow(describePdfAuditIssue(issue))}`);
        if (issues.length === 0) log(pc.green("没有发现会让阅读器判损坏的结构问题，且全部显示文字都可改。"));
        else warn(`${summarizePdfAudit(issues)}（红色 = 可能打不开 / 丢内容，黄色 = 打得开但有一部分改不了。）`);
      });
      // 体检要能当闸门用：有问题时给非零退出码，脚本里 `oc pdf check && ...` 才拦得住。
      if (issues.length > 0) process.exitCode = 1;
    });
}
