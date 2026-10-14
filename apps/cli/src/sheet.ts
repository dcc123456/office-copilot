/**
 * `oc sheet` —— CLI 端的表格读侧（spec 阶段 6 / 提交 16）。
 *
 * 只做装配，不做第二套实现：解析、单位换算、列类型推断、聚合与筛选全部走
 * `@office-copilot/sheet-io` 的公开入口，与桌面 app、AI 工具层跑的是同一份引擎，
 * 所以"三端对同一份 `.xlsx` 行为一致"这条完成判定是结构性成立的，不靠对拍。
 *
 * 写侧刻意不在 CLI 落地：撤销/modlog/原子替换（`.tmp` → rename）都在主进程那条路上
 * （`apps/desktop/electron/sheetHandlers.ts`），CLI 再写一份保存等于开第二个写入口。
 */
import { Command } from "commander";
import pc from "picocolors";
import fs from "node:fs";
import path from "node:path";
import {
  SheetFile,
  aggregateSheet,
  auditXlsxPackage,
  describeSheet,
  duplicateGroups,
  findRows,
  readOnlyReason,
  readRange,
  sheetPayloadOf,
  usedCellCount,
  usedRangeOf,
  type AggregateFunction,
  type CompatIssue,
  type SheetPayload,
  type WorkbookPayload,
} from "@office-copilot/sheet-io";
import { emit, fail, log, warn } from "./fsutil.js";

/** 打开一份 `.xlsx` 后 CLI 要用的东西：载荷、按 `--sheet` 选中的表、只读理由。 */
interface LoadedWorkbook {
  workbook: WorkbookPayload;
  /** `--sheet` 选中的表；`info` 一条命令不选表，故可为 null。 */
  sheet: SheetPayload | null;
  /** 只读打开的理由；可编辑时为 null（大表闸门与桌面同一口径，`sheet-io` 的 `EDITABLE_CELL_LIMIT`）。 */
  readonlyReason: string | null;
}

/**
 * 读文件并交给引擎解析。
 *
 * @param filePath 用户给的 `.xlsx` 路径（相对 cwd 也行）。
 * @param sheetRef `--sheet` 的值：工作表名，或 `info` 里显示的 1 基序号；省略就取第一张表。
 * @param needSheet 选中的表缺失时是否直接失败（`info` 不需要单张表）。
 * @returns 载荷 + 只读理由；文件不存在、宏工作簿、认不出表名都走 `fail` 给一句中文。
 */
async function loadWorkbook(filePath: string, sheetRef?: string, needSheet = true): Promise<LoadedWorkbook> {
  const abs = path.resolve(filePath);
  if (!fs.existsSync(abs)) fail(`文件不存在：${abs}`);
  let file: SheetFile;
  try {
    const bytes = await fs.promises.readFile(abs);
    file = await SheetFile.fromBytes(new Uint8Array(bytes), { fileName: path.basename(abs) });
  } catch (error) {
    fail(error instanceof Error ? error.message : String(error));
  }
  const model = await file.toWorkbook();
  const workbook: WorkbookPayload = {
    fileName: model.fileName,
    sheets: model.sheets.map(sheetPayloadOf),
    warnings: model.warnings,
    unsupported: model.unsupported,
  };
  const sheet = pickSheet(workbook.sheets, sheetRef);
  if (needSheet && !sheet) {
    fail(
      sheetRef
        ? `这份工作簿里没有「${sheetRef}」这张表。可选：${workbook.sheets.map((item, index) => `${index + 1}=${item.name}`).join("、")}`
        : "这份工作簿里没有任何工作表。",
    );
  }
  return { workbook, sheet: sheet ?? null, readonlyReason: readOnlyReason(workbook) };
}

/**
 * 取"必须有一张表"的那张：把可空的地方收在一处，省掉满屏 `!`。
 *
 * @param loaded 已打开的工作簿。
 * @returns 选中的表载荷；理论上不会缺（`loadWorkbook` 已按 `needSheet` 挡过）。
 */
function requireSheet(loaded: LoadedWorkbook): SheetPayload {
  if (!loaded.sheet) fail("没选中工作表。");
  return loaded.sheet;
}

/**
 * 按名字或 1 基序号挑一张表。
 *
 * @param sheets 全部表载荷。
 * @param ref `--sheet` 的原始值；undefined 表示用第一张表。
 * @returns 选中的表；认不出时 undefined，由调用方决定是失败还是继续。
 */
export function pickSheet(sheets: SheetPayload[], ref?: string): SheetPayload | undefined {
  if (!ref) return sheets[0];
  const byName = sheets.find((item) => item.name === ref);
  if (byName) return byName;
  const index = Number(ref);
  return Number.isInteger(index) && index >= 1 && index <= sheets.length ? sheets[index - 1] : undefined;
}

/**
 * 把 `--metric fn:列` 翻成引擎的聚合规格。
 *
 * @param raw 用户写的度量串，例如 `sum:销量`、`avg 销量`（分隔符 `:` 或空格都行）。
 * @returns 规格；函数名不在白名单或列引用为空时给一句中文并退出。
 */
export function parseMetric(raw: string): { fn: AggregateFunction; column: string } {
  const [fnText = "", ...rest] = raw.split(/[:\s]/);
  const column = rest.join(" ").trim();
  const fn = fnText === "avg" ? "average" : fnText;
  if (!["count", "sum", "average", "min", "max"].includes(fn) || !column) {
    fail(`--metric 要写成 函数:列，函数只能是 count/sum/average(avg)/min/max。收到的是「${raw}」。`);
  }
  return { fn: fn as AggregateFunction, column };
}

/** 逗号分隔的列引用串（`--by`、`--keys`）→ 数组。 */
export function splitColumns(raw: string): string[] {
  return raw
    .split(",")
    .map((text) => text.trim())
    .filter(Boolean);
}

/** 让 `--metric` 可重复：commander 默认后者覆盖前者，聚合要的是列表。 */
function collect(value: string, previous: string[]): string[] {
  return [...previous, value];
}

/** 读侧不许抛异常：引擎给的"为什么不行"原样说给人听。 */
function reportReason(reason: string | undefined): never {
  fail(reason ?? "引擎没给出原因。");
}

/** 把打开时引擎记下的警告与未建模部件念出来（spec §零.7：不许静默丢弃）。 */
function noteDiagnostics(workbook: WorkbookPayload, readonlyReason: string | null): void {
  for (const text of workbook.warnings) warn(text);
  for (const sheet of workbook.sheets) for (const text of sheet.warnings) warn(`${sheet.name}：${text}`);
  if (workbook.unsupported.length > 0) {
    log(
      pc.dim(
        `未建模但已原样保留的结构 ${workbook.unsupported.length} 处：${[...new Set(workbook.unsupported.map((item) => item.reason))].join("、")}`,
      ),
    );
  }
  if (readonlyReason) warn(readonlyReason);
}

/** 首行是否表头：`--no-header` 时 commander 给 false，默认 true。 */
function hasHeaderOf(opts: { header?: boolean }): boolean {
  return opts.header ?? true;
}

/**
 * 引擎的问题码翻成一句中文（`AGENTS.md` §5.7：引擎只出码，文案在装配层）。
 *
 * @param issue `auditXlsxPackage` 报出的一条结构问题。
 * @returns 说明这条问题会让办公软件怎么做（打不开 / 忽略部件 / 丢内容），不含部件名。
 */
export function describeCompatIssue(issue: CompatIssue): string {
  return COMPAT_ISSUE_TEXT[issue.code] + (issue.detail ? `：${issue.detail}` : "");
}

/** 每个问题码一句文案；键类型是 `CompatIssue["code"]`，新增码忘了配文案会编译不过。 */
const COMPAT_ISSUE_TEXT: Record<CompatIssue["code"], string> = {
  "part-not-declared": "部件没在 [Content_Types].xml 里登记，办公软件会直接忽略它",
  "declared-part-missing": "内容类型表登记了一个包里不存在的部件，打开时报损坏",
  "relationship-target-missing": "关系指向的部件不在包里，顺着这条引用会断",
  "relationship-id-duplicate": "同一张关系表里关系编号重复，引用会指错地方",
  "unresolved-relationship-id": "部件引用的关系编号在自己的关系表里查不到，这块内容会丢",
  "worksheet-child-order": "表部件子元素顺序违反 §18.3.1.100 的序列，Excel 会拒绝解析",
  "worksheet-missing-sheet-data": "表部件缺 <sheetData>，这是唯一的必选元素",
  "chart-missing-chart": "图表部件是空壳，没有 <c:chart>，打开时报损坏",
};

export function registerSheetCommands(program: Command): void {
  const sheet = program
    .command("sheet")
    .description("读 .xlsx：概览 / 列画像 / 取区域 / 筛选 / 分组聚合 / 重复行预览 / 兼容性体检（与桌面同一份引擎，只读）");

  sheet
    .command("info <file>")
    .description("工作簿概览：每张表的已用区域、格数、警告与未建模结构")
    .option("--json", "输出 JSON")
    .action(async (file: string, opts: { json?: boolean }) => {
      const loaded = await loadWorkbook(file, undefined, false);
      const rows = loaded.workbook.sheets.map((item, index) => ({
        order: index + 1,
        name: item.name,
        usedRange: usedRangeOf(item) ?? "空表",
        cellCount: Object.keys(item.cells).length,
        merges: item.merges.length,
        columnWidthRecords: item.columns.length,
      }));
      emit(
        opts.json,
        {
          fileName: loaded.workbook.fileName,
          totalCells: usedCellCount(loaded.workbook),
          sheets: rows,
          warnings: loaded.workbook.warnings,
          unsupported: loaded.workbook.unsupported,
        },
        () => {
          log(pc.bold(loaded.workbook.fileName) + pc.dim(`  共 ${usedCellCount(loaded.workbook)} 个已用格`));
          for (const row of rows) {
            log(
              `  ${pc.dim(String(row.order).padStart(2, " "))} ${pc.yellow(row.name)}  ${row.usedRange}  已用格 ${row.cellCount}  合并 ${row.merges}  列宽记录 ${row.columnWidthRecords}`,
            );
          }
          noteDiagnostics(loaded.workbook, loaded.readonlyReason);
        },
      );
    });

  sheet
    .command("stats <file>")
    .description("列画像：类型推断、缺失数、众数、数值统计量与离群值")
    .option("-s, --sheet <nameOrIndex>", "工作表名或 1 基序号（默认第一张）")
    .option("--no-header", "首行不是表头，把它一起计入统计")
    .option("--json", "输出 JSON")
    .action(async (file: string, opts: { sheet?: string; header?: boolean; json?: boolean }) => {
      const loaded = await loadWorkbook(file, opts.sheet);
      const profile = describeSheet(requireSheet(loaded), { hasHeader: hasHeaderOf(opts) });
      emit(opts.json, profile, () => {
        log(
          pc.bold(`${profile.sheetName}  ${profile.usedRange ?? "空表"}`) +
            pc.dim(`  ${profile.rowCount} 行数据 × ${profile.columnCount} 列`),
        );
        for (const text of profile.headerWarnings) warn(text);
        for (const column of profile.columns) {
          const numeric = column.numericText
            ? pc.cyan(
                `  Σ ${column.numericText.sum}  均值 ${column.numericText.average}  中位 ${column.numericText.median}  ${column.numericText.min}~${column.numericText.max}`,
              )
            : "";
          const extras = [
            column.missing > 0 ? pc.yellow(`缺失 ${column.missing}`) : "",
            column.pendingRecalc > 0 ? pc.yellow(`待 Excel 重算 ${column.pendingRecalc}`) : "",
            column.outlierCount > 0 ? pc.magenta(`离群 ${column.outlierCount}`) : "",
          ].filter(Boolean);
          log(
            `  ${pc.dim(column.letter.padStart(2, " "))} ${column.header}  ${pc.dim(column.type)}  非空 ${column.populated} 互异 ${column.distinct}${numeric}${extras.length > 0 ? "  " + extras.join(" ") : ""}`,
          );
          if (column.topValues.length > 0) {
            log(pc.dim(`      众数：${column.topValues.map((item) => `${item.text}×${item.count}`).join("、")}`));
          }
        }
        if (profile.columnsTruncated) warn(`列太多，只画像了前 ${profile.columns.length} 列（共 ${profile.columnCount} 列）。`);
        noteDiagnostics(loaded.workbook, loaded.readonlyReason);
      });
    });

  sheet
    .command("read <file>")
    .description("按 A1 区域取数（读的是显示文本，日期/百分比已经换算好）")
    .requiredOption("-r, --range <ref>", "区域写法，例如 A1:D20；单格 C2 也收")
    .option("-s, --sheet <nameOrIndex>", "工作表名或 1 基序号（默认第一张）")
    .option("--max-rows <n>", "最多读多少行", (v) => Number.parseInt(v, 10))
    .option("--json", "输出 JSON")
    .action(async (file: string, opts: { range: string; sheet?: string; maxRows?: number; json?: boolean }) => {
      const loaded = await loadWorkbook(file, opts.sheet);
      const read = readRange(requireSheet(loaded), {
        range: opts.range,
        ...(opts.maxRows === undefined ? {} : { maxRows: opts.maxRows }),
      });
      if (!read.ok) reportReason(read.reason);
      emit(opts.json, read, () => {
        log(pc.bold(read.ref) + pc.dim(`  数据区里这块共 ${read.rowCount} 行 × ${read.colCount} 列`));
        for (const row of read.rows) log(`  ${pc.dim(String(row.rowNumber).padStart(4, " "))}  ${row.cells.join("\t")}`);
        if (read.truncated) warn(`只读了 ${read.rows.length} 行 × ${read.rows[0]?.cells.length ?? 0} 列，剩下的没上送（行可用 --max-rows 放宽，列一律最多读 50 列）。`);
      });
    });

  sheet
    .command("find <file>")
    .description("按条件筛行，条件写法与 COUNTIF 一致（`>100`、`<=2`、`<>北京`、`北京`）")
    .requiredOption("-c, --column <letterOrHeader>", "比较哪一列：列字母或表头原文")
    .requiredOption("--criteria <text>", "COUNTIF 风格的条件")
    .option("-s, --sheet <nameOrIndex>", "工作表名或 1 基序号（默认第一张）")
    .option("--no-header", "首行不是表头")
    .option("--json", "输出 JSON")
    .action(async (file: string, opts: { column: string; criteria: string; sheet?: string; header?: boolean; json?: boolean }) => {
      const loaded = await loadWorkbook(file, opts.sheet);
      const found = findRows(requireSheet(loaded), {
        column: opts.column,
        criteria: opts.criteria,
        hasHeader: hasHeaderOf(opts),
      });
      if (!found.ok) reportReason(found.reason);
      emit(opts.json, found, () => {
        log(pc.dim(`命中 ${found.matchedCount} 行：`));
        for (const hit of found.hits) log(`  ${pc.yellow(hit.address)}  ${hit.display}`);
        if (found.truncated) warn(`命中太多，只列了前 ${found.hits.length} 条。`);
      });
    });

  sheet
    .command("aggregate <file>")
    .description("分组聚合：数值由引擎从现场表读，不在命令行里手填")
    .requiredOption("--by <columns>", "分组列，逗号分隔（1~3 个），例如 `地区` 或 `月份,地区`")
    .requiredOption("-m, --metric <fn:column>", "聚合，形如 `sum:销量`；可重复", collect, [])
    .option("-s, --sheet <nameOrIndex>", "工作表名或 1 基序号（默认第一张）")
    .option("--no-header", "首行不是表头")
    .option("--json", "输出 JSON")
    .action(async (file: string, opts: { by: string; metric: string[]; sheet?: string; header?: boolean; json?: boolean }) => {
      const loaded = await loadWorkbook(file, opts.sheet);
      const result = aggregateSheet(requireSheet(loaded), {
        groupBy: splitColumns(opts.by),
        aggregates: opts.metric.map(parseMetric),
        hasHeader: hasHeaderOf(opts),
      });
      if (!result.ok) reportReason(result.reason);
      emit(opts.json, result, () => {
        for (const group of result.groups) {
          const values = Object.entries(group.valueTexts)
            .map(([label, text]) => `${label}=${text}`)
            .join("  ");
          log(`  ${pc.yellow(group.key)}  ${pc.dim(`${group.rowCount} 行`)}${values ? `  ${pc.cyan(values)}` : ""}`);
        }
        log(pc.dim(`共 ${result.groupCount} 组`));
        if (result.truncated) warn(`组太多，只列了前 ${result.groups.length} 组。`);
      });
    });

  sheet
    .command("dedupe <file>")
    .description("重复行预览：按这些键能安全删掉几行（删行要回桌面做，撤销与 .bak 才齐）")
    .requiredOption("--keys <columns>", "判定重复的列，逗号分隔，例如 `月份,地区`")
    .option("-s, --sheet <nameOrIndex>", "工作表名或 1 基序号（默认第一张）")
    .option("--no-header", "首行不是表头")
    .option("--json", "输出 JSON")
    .action(async (file: string, opts: { keys: string; sheet?: string; header?: boolean; json?: boolean }) => {
      const loaded = await loadWorkbook(file, opts.sheet);
      const result = duplicateGroups(requireSheet(loaded), { keys: splitColumns(opts.keys), hasHeader: hasHeaderOf(opts) });
      if (!result.ok) reportReason(result.reason);
      emit(opts.json, result, () => {
        for (const group of result.groups) {
          log(`  ${pc.yellow(group.key)}  重复 ${group.count} 次：${pc.dim(`第 ${group.rowNumbers.join("、")} 行`)}`);
        }
        log(pc.dim(`按这些键去重能安全删掉 ${result.redundantRowCount} 行`));
        if (result.redundantRowCount > 0) {
          warn("删行请回桌面 app：那里才有撤销与同名 .bak（`AGENTS.md` §2.4/§2.5）。");
        }
      });
    });

  sheet
    .command("check <file>")
    .description("兼容性体检：按 OOXML 的硬约束查这份文件会不会被 Excel / WPS 判成损坏")
    .option("--json", "输出 JSON")
    .action(async (file: string, opts: { json?: boolean }) => {
      const abs = path.resolve(file);
      if (!fs.existsSync(abs)) fail(`文件不存在：${abs}`);
      let issues: CompatIssue[];
      try {
        issues = await auditXlsxPackage(new Uint8Array(await fs.promises.readFile(abs)));
      } catch (error) {
        fail(error instanceof Error ? error.message : String(error));
      }
      emit(opts.json, { fileName: path.basename(abs), ok: issues.length === 0, issues }, () => {
        for (const issue of issues) log(`  ${pc.red(issue.part)}  ${describeCompatIssue(issue)}`);
        if (issues.length === 0) log(pc.green("没有发现会让办公软件判损坏的结构问题。"));
        else warn(`发现 ${issues.length} 处结构风险：Excel / WPS 可能打不开，或打开后丢内容。`);
      });
      // 体检要能当闸门用：有问题时给非零退出码，脚本里 `oc sheet check && ...` 才拦得住。
      if (issues.length > 0) process.exitCode = 1;
    });
}
