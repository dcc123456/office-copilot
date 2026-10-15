/**
 * `oc ppt` —— CLI 端的幻灯片读侧与兼容性体检（`docs/office-parity-spec.md` 阶段 8）。
 *
 * 与 `oc sheet` 同一条纪律：只做装配，不写第二套实现——结构模型来自 `@office-copilot/pptx-io` 的
 * `readPresentation`，判定来自 `auditPptxPackage`，与桌面 app 走同一份引擎。
 * 写侧刻意不在 CLI 落地：撤销 / modlog / 原子替换都在桌面主进程那条路上（`AGENTS.md` §2.4/§2.5）。
 */
import { Command } from "commander";
import pc from "picocolors";
import fs from "node:fs";
import path from "node:path";
import { auditPptxPackage, readPresentation, type PresentationModel, type PptxAuditIssue, type SlideShape } from "@office-copilot/pptx-io";
import { describeSlideReadFailure } from "@office-copilot/shared";
import { clip, emit, fail, log, warn } from "./fsutil.js";

/**
 * 引擎的问题码翻成一句中文（`AGENTS.md` §5.7：引擎只出码，文案在装配层）。
 *
 * @param issue `auditPptxPackage` 报出的一条结构问题。
 * @returns 说明这条问题会让办公软件怎么做（打不开 / 忽略部件 / 丢内容），不含部件名。
 */
export function describePptIssue(issue: PptxAuditIssue): string {
  return PPT_ISSUE_TEXT[issue.code] + (issue.detail ? `：${issue.detail}` : "");
}

/** 每个问题码一句文案；键类型是 `PptxAuditIssue["code"]`，新增码忘了配文案会编译不过。 */
const PPT_ISSUE_TEXT: Record<PptxAuditIssue["code"], string> = {
  "part-not-declared": "部件没在 [Content_Types].xml 里登记，办公软件会直接忽略它",
  "declared-part-missing": "内容类型表登记了一个包里不存在的部件，打开时报损坏",
  "relationship-target-missing": "关系指向的部件不在包里，顺着这条引用会断",
  "relationship-id-duplicate": "同一张关系表里关系编号重复，引用会指错地方",
  "unresolved-relationship-id": "部件引用的关系编号在自己的关系表里查不到，这块内容会丢",
  "slide-missing-shape-tree": "幻灯片缺 <p:cSld> 里的 <p:spTree>，这是必选元素，打开时报损坏",
  "slide-missing-layout-relationship": "幻灯片没有指向版式（slideLayout）的关系，PowerPoint 会修复或拒绝打开",
  "slide-incomplete-transform": "几何变换只写了一半（<a:off> 的 x、y 与 <a:ext> 的 cx、cy 得成对，<a:xfrm> 里这两块也都要在），打开时报损坏",
};

/** 读侧不许抛异常：引擎给的"为什么不行"原样说给人听。 */
function reportFailure(error: unknown): never {
  fail(error instanceof Error ? error.message : String(error));
}

/**
 * 把一个形状的文字念成 CLI 上那一行。
 *
 * @param shape 读模型里的一个形状快照。
 * @returns 普通形状是各段用 ` / ` 隔开；图形框里是**表**时逐行用 ` ⏎ ` 隔开、行内格子用 ` | ` 隔开
 *   （表里的字不写在形状的文本体里，只念 `paragraphTexts` 会念成"没文字"）；两者都没有时是空串。
 */
function shapeText(shape: SlideShape): string {
  if (!shape.table) return shape.paragraphTexts.join(" / ");
  return shape.table.rows.map((row) => row.map((cell) => cell.paragraphTexts.join(" / ")).join(" | ")).join(" ⏎ ");
}

export function registerPptCommands(program: Command): void {
  const ppt = program
    .command("ppt")
    .description("读 .pptx：幻灯片与形状清单 / 兼容性体检（与桌面同一份引擎，只读）");

  ppt
    .command("info <file>")
    .description("演示文稿概览：幻灯片数、页面尺寸（EMU 原样）、每张表的形状与文字")
    .option("--json", "输出 JSON")
    .option("--max-text <chars>", "每个形状的文字最多显示多少字符", (v) => parseInt(v, 10), 60)
    .action(async (file: string, opts: { json?: boolean; maxText: number }) => {
      const abs = path.resolve(file);
      if (!fs.existsSync(abs)) fail(`文件不存在：${abs}`);
      let model: PresentationModel;
      try {
        model = (await readPresentation(new Uint8Array(await fs.promises.readFile(abs)))).model;
      } catch (error) {
        reportFailure(error);
      }
      emit(opts.json, model, () => {
        log(`${pc.bold(path.basename(abs))}：${model.slides.length} 张幻灯片，页面 ${model.slideSizeEmu.cx}×${model.slideSizeEmu.cy} EMU`);
        for (const slide of model.slides) {
          log(pc.dim(`  #${slide.order + 1} ${slide.partPath}`));
          if (slide.readFailure) warn(`    ${describeSlideReadFailure(slide.readFailure)}`);
          for (const shape of slide.shapes) {
            const tags = [shape.kind, shape.isPlaceholder ? `占位符:${shape.placeholderRef ?? "?"}` : null]
              .filter(Boolean)
              .join(" ");
            const text = clip(shapeText(shape), opts.maxText);
            log(`    ${pc.yellow(`id=${shape.shapeId}`)} ${shape.name}（${tags}）${text ? `  ${pc.dim(text)}` : ""}`);
          }
          // 没有备注页时这行不打（跟"形状没文字就不打"同一个口径）；有备注页时把"没写字"如实说出来，
          // 免得读的人以为备注页是读不动。
          if (slide.notes) {
            const notes = slide.notes.paragraphTexts.join(" / ").trim();
            log(pc.dim(`    备注（${slide.notes.partPath}）：${notes === "" ? "有备注页但没写字" : clip(notes, opts.maxText)}`));
          }
        }
      });
    });

  ppt
    .command("check <file>")
    .description("兼容性体检：按 OOXML 的硬约束查这份文件会不会被 PowerPoint / WPS 判成损坏")
    .option("--json", "输出 JSON")
    .action(async (file: string, opts: { json?: boolean }) => {
      const abs = path.resolve(file);
      if (!fs.existsSync(abs)) fail(`文件不存在：${abs}`);
      let issues: PptxAuditIssue[];
      try {
        issues = await auditPptxPackage(new Uint8Array(await fs.promises.readFile(abs)));
      } catch (error) {
        reportFailure(error);
      }
      emit(opts.json, { fileName: path.basename(abs), ok: issues.length === 0, issues }, () => {
        for (const issue of issues) log(`  ${pc.red(issue.part)}  ${describePptIssue(issue)}`);
        if (issues.length === 0) log(pc.green("没有发现会让办公软件判损坏的结构问题。"));
        else warn(`发现 ${issues.length} 处结构风险：PowerPoint / WPS 可能打不开，或打开后丢内容。`);
      });
      // 体检要能当闸门用：有问题时给非零退出码，脚本里 `oc ppt check && ...` 才拦得住。
      if (issues.length > 0) process.exitCode = 1;
    });
}
