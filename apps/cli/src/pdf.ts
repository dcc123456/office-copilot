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
import { emit, fail, log, warn } from "./fsutil.js";

/** 每个问题码一句文案（引擎只出码，§5.7）；键类型是 `PdfAuditIssue["code"]`，新增码忘了配文案会编译不过。 */
const PDF_ISSUE_TEXT: Record<PdfAuditIssue["code"], string> = {
  "parse-failed": "解析器读不动这份文件（交叉引用表或对象流坏了），办公软件多半也打不开",
  "header-missing": "文件头没有 %PDF- 版本标记，阅读器可能直接拒开",
  "eof-missing": "文件尾没有 %%EOF，说明这份是被截断的，末尾的页与交叉引用可能不完整",
  "catalog-missing": "目录（/Root 或 /Pages）在文件里查不到，阅读器根本定位不到页",
  "page-count-mismatch": "页树的数量对不上，阅读器会按自己那套猜页数",
  "content-stream-undecodable": "这一页的内容流解不出来，打开后这一页会是空白",
  "text-not-editable": "有一部分文字我们读不出内容，因而也改不了",
};

/**
 * 引擎的问题码翻成一句中文。
 *
 * @param issue `auditPdfStructure` 报出的一条结论。
 * @returns 「第 N 页 · 」前缀（整份文件级别的问题没有）+ 文案 + detail 里的数字。
 */
export function describePdfIssue(issue: PdfAuditIssue): string {
  const where = issue.pageNumber === null ? "" : `第 ${issue.pageNumber} 页 · `;
  return where + PDF_ISSUE_TEXT[issue.code] + (issue.detail ? `：${issue.detail}` : "");
}

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
        for (const issue of issues) log(`  ${issue.severity === "broken" ? pc.red(describePdfIssue(issue)) : pc.yellow(describePdfIssue(issue))}`);
        if (issues.length === 0) log(pc.green("没有发现会让阅读器判损坏的结构问题，且全部显示文字都可改。"));
        else warn(`发现 ${issues.length} 处风险（红色 = 可能打不开 / 丢内容，黄色 = 打得开但有一部分改不了）。`);
      });
      // 体检要能当闸门用：有问题时给非零退出码，脚本里 `oc pdf check && ...` 才拦得住。
      if (issues.length > 0) process.exitCode = 1;
    });
}
