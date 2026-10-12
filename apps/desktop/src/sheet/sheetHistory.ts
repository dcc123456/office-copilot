/**
 * 表格改动与修改历史（modlog）之间的那一层（spec 阶段 3 · C15/C17）。
 *
 * 存储复用文档那条已有的四个通道（`electron/modlog.ts`），不新建第二份历史（`AGENTS.md` §3
 * 禁止第二套同类基础设施）；这一层只做表格专属的事：把一次 op 落地结果写成一条记录、把记录里的
 * 逆操作还原回现场的表。真正的定位与现场核对在 `@office-copilot/sheet-io`，撤销语义只有一份。
 */
import { revertSheetOps, type SheetCoreOp, type SheetPayload, type SheetOpsResult } from "@office-copilot/sheet-io";

/**
 * 一次表格改动在历史里的样子。
 *
 * 只存 op 与逆操作，不存整表快照：历史上限 60 条，而一张大表的载荷能有几 MB，存快照会让
 * `userData/modlog/<文档名>.json` 直接失控。
 */
export interface SheetEditRecord {
  id: string;
  ts: number;
  label: string;
  summary: string;
  /** 改动落在哪张工作表上（包内部件路径）；撤销时按它找表。 */
  partPath: string;
  ops: SheetCoreOp[];
  inverses: SheetCoreOp[];
  /** 已经撤销过：一条记录不许撤销两次，否则会把用户后来的编辑再倒回去一遍。 */
  undone?: boolean;
}

/** 撤销一条记录的结果。`failed` 是没能落地的逆操作的理由，UI 据此报"部分撤销"。 */
export interface SheetUndoOutcome {
  after: SheetPayload;
  /** 真正落地的逆操作条数。 */
  appliedCount: number;
  failed: string[];
  /** 一条都没落地时给个整体理由。 */
  reason: string | null;
}

/** 记进历史的默认标签，与文档那条"AI 修改"平级——手工改表格也留痕。 */
const SHEET_EDIT_LABEL = "表格改动";

/**
 * 判断一条历史记录是不是表格的。
 *
 * 历史文件里文档记录与表格记录混在同一个 docKey 下（同名 `.xlsx` 不会有 docx 记录，但
 * `listModifications` 返回的是 `unknown[]`，不校验就用会拿到 undefined 字段）。
 *
 * @param value 从主进程读回来的一条记录。
 * @returns 有 `partPath` 与 `inverses` 数组才算表格记录。
 */
export function isSheetEditRecord(value: unknown): value is SheetEditRecord {
  if (typeof value !== "object" || value === null) return false;
  const record = value as Partial<SheetEditRecord>;
  return typeof record.partPath === "string" && Array.isArray(record.inverses) && Array.isArray(record.ops);
}

/**
 * 从一堆记录里挑出最近一次还没撤销的表格改动。
 *
 * @param records `listModifications` 的原始返回（顺序不保证，这里自己按 ts 排）。
 * @returns 最新的表格记录；没有可撤销的改动时返回 null。
 */
export function pickLatestSheetEdit(records: unknown[]): SheetEditRecord | null {
  const candidates = records.filter(isSheetEditRecord).filter((record) => !record.undone);
  if (!candidates.length) return null;
  return candidates.reduce((latest, record) => (record.ts > latest.ts ? record : latest), candidates[0]!);
}

/**
 * 把一次 op 落地结果收成一条历史记录。
 *
 * @param result `runSheetOps` 的返回；`skipped` 里的条目不进历史（它们没有生效，撤销它们等于凭空造改动）。
 * @param summary 一句话说明这次改了什么，历史面板直接显示。
 * @returns 可直接交给 `appendModification` 的记录。
 */
export function sheetEditRecordOf(result: SheetOpsResult, summary: string): SheetEditRecord {
  return {
    id: crypto.randomUUID(),
    ts: Date.now(),
    label: SHEET_EDIT_LABEL,
    summary,
    partPath: result.after.partPath,
    ops: result.ops,
    inverses: result.inverses,
  };
}

/**
 * 撤销一条表格改动记录：把逆操作倒序应用回现场。
 *
 * 纯函数，不碰编辑器也不碰磁盘——调用方拿到 `after` 再决定怎么装回工作簿。
 * 现场与记录对不上的条目一律跳过并给出理由，绝不猜位置（`AGENTS.md` §2.5）。
 *
 * @param sheet 现场的表载荷。
 * @param record 要撤销的记录。
 * @returns `SheetUndoOutcome`；`partPath` 对不上时 `appliedCount` 为 0 并给理由。
 */
export function undoSheetEdit(sheet: SheetPayload, record: SheetEditRecord): SheetUndoOutcome {
  if (record.partPath !== sheet.partPath) {
    return { after: sheet, appliedCount: 0, failed: [], reason: "这条记录属于另一张表" };
  }
  const result = revertSheetOps(sheet, record.inverses);
  return {
    after: result.after,
    appliedCount: result.ops.length,
    failed: result.skipped.map((item) => item.reason),
    reason: result.ops.length ? null : `这次改动的内容后来被改过，已不再安全回滚${result.skipped.length ? `：${result.skipped.map((item) => item.reason).join("；")}` : ""}`,
  };
}
