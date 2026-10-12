/**
 * 主进程 ↔ 渲染层（以及 CLI、AI 工具层）之间传工作簿用的形状。
 *
 * 为什么单独一层：`SheetModel.sourceXml` 是整张表的 XML 原文，可能上百 KB，
 * 而它**只属于持有文件的那一侧**（`AGENTS.md` §2.4：改写必须在打开时那份原文上做定点替换）。
 * 渲染层要的是格子内容；保存时主进程重新从磁盘打开、把内容按 `applySheetPayload` 写回，
 * 于是原文永远不过 IPC，也不会被渲染层的副本污染。
 */
import type { SheetModel, UnsupportedFeature } from "./model.js";

/** 一张表去掉 XML 原文后的可结构化克隆形状（其余字段全是 JSON 安全值）。 */
export type SheetPayload = Omit<SheetModel, "sourceXml">;

/** 一个工作簿的跨界形状。 */
export interface WorkbookPayload {
  fileName: string;
  sheets: SheetPayload[];
  /** 读取时的警告，必须一路带到 UI，不许在跨界处丢掉（spec §零.7）。 */
  warnings: string[];
  /** 未建模但已原样保留的部件清单，UI 据此提示"这些结构我们不编辑，但没丢"。 */
  unsupported: UnsupportedFeature[];
}

/**
 * 把表模型压成跨界载荷。
 *
 * @param sheet reader 产出的表模型。
 * @returns 新对象；`cells` 等引用与入参共享，跨界后调用方不许再就地改原 `sheet`。
 */
export function sheetPayloadOf(sheet: SheetModel): SheetPayload {
  return {
    name: sheet.name,
    sheetId: sheet.sheetId,
    partPath: sheet.partPath,
    lastColIndex: sheet.lastColIndex,
    lastRowIndex: sheet.lastRowIndex,
    cells: sheet.cells,
    merges: sheet.merges,
    columns: sheet.columns,
    rows: sheet.rows,
    freeze: sheet.freeze,
    warnings: sheet.warnings,
  };
}

/**
 * 把渲染层带回的内容写进刚打开的表模型（保存路径的唯一入口）。
 *
 * @param target 主进程现开的表模型，持有磁盘上的 `sourceXml`。
 * @param payload 渲染层编辑后的同一张表；`partPath`/`sheetId` 以 `target` 为准，
 *   调用方须先按 `partPath` 配对，传错位不许静默写进别的表。
 */
export function applySheetPayload(target: SheetModel, payload: SheetPayload): void {
  target.name = payload.name;
  target.cells = payload.cells;
  target.merges = payload.merges;
  target.columns = payload.columns;
  target.rows = payload.rows;
  target.freeze = payload.freeze;
  target.lastColIndex = payload.lastColIndex;
  target.lastRowIndex = payload.lastRowIndex;
}

/**
 * 打开工作簿后收成跨界载荷。
 *
 * @param fileName 展示用的文件名。
 * @param sheets 已解析的全部表。
 * @param warnings 包级警告。
 * @param unsupported 未建模部件清单。
 * @returns `WorkbookPayload`；`sheets` 为空表示这个包里没有任何可读的表。
 */
export function workbookPayload(
  fileName: string,
  sheets: SheetModel[],
  warnings: string[],
  unsupported: UnsupportedFeature[],
): WorkbookPayload {
  return { fileName, sheets: sheets.map(sheetPayloadOf), warnings, unsupported };
}
