/**
 * 排序 op 的比较口径与置换计算（spec 决策 D7）。
 *
 * 这里**只算行序**，不碰载荷：置换的落地与撤销在 `sheet-edit.ts`（那里才有内容寻址与现场核对）。
 * 口径是 D7 声明的自证规则，不是从 Excel 抄来的——Microsoft 的排序帮助页拿不到逐条次序，
 * 本机也没有 Microsoft Excel 可比对（`AGENTS.md` §8），所以跨类型分秩与空格恒最后都在单测里钉死，
 * 不写"与 Excel 一致"。
 */
import type { SheetCell } from "./model.js";

/** 类型秩（D7）：数字 < 文本 < 布尔 < 错误值 < 空格。秩不受升/降影响，空格恒最后。 */
const RANK_NUMBER = 0;
const RANK_TEXT = 1;
const RANK_BOOLEAN = 2;
const RANK_ERROR = 3;
const RANK_BLANK = 4;

/** 中文文本比较器：`numeric` 让 `第2章 < 第10章`，`base` 让大小写不参与定序。 */
const TEXT_COLLATOR = new Intl.Collator("zh-Hans-CN", { numeric: true, sensitivity: "base" });

/**
 * 一行参与排序的键值。
 *
 * 三个字段按秩取用：`RANK_NUMBER` 看 `numberValue`，`RANK_TEXT`/`RANK_ERROR` 看 `textValue`，
 * `RANK_BOOLEAN` 看 `numberValue`（FALSE=0/TRUE=1），`RANK_BLANK` 谁都别看。
 */
export interface SortKey {
  rank: number;
  numberValue: number;
  textValue: string;
}

/**
 * 取一格作为排序键。
 *
 * @param cell 键列里的那一格（`undefined` 表示空位；带公式的格在造 op 时就被 D7 拒了，不会走到这里）。
 * @returns 排序键；`kind:"blank"`、`value` 为 `null`/`undefined` 都算空格（秩 `RANK_BLANK`）。
 */
export function sortKeyOf(cell: SheetCell | undefined): SortKey {
  const value = cell?.value ?? null;
  if (!cell || cell.kind === "blank" || value === null || value === undefined) {
    return { rank: RANK_BLANK, numberValue: 0, textValue: "" };
  }
  if (typeof value === "number") return { rank: RANK_NUMBER, numberValue: value, textValue: "" };
  if (typeof value === "boolean") return { rank: RANK_BOOLEAN, numberValue: value ? 1 : 0, textValue: "" };
  if (cell.kind === "error") return { rank: RANK_ERROR, numberValue: 0, textValue: String(value) };
  return { rank: RANK_TEXT, numberValue: 0, textValue: String(value) };
}

/**
 * 两个排序键的升序比较（不含升降与空格规则，调用方负责）。
 *
 * @param left 左键。
 * @param right 右键。
 * @returns 负/零/正；秩不同先按秩，秩相同才按该类型的自然序。
 */
function compareKeysAscending(left: SortKey, right: SortKey): number {
  if (left.rank !== right.rank) return left.rank - right.rank;
  if (left.rank === RANK_TEXT || left.rank === RANK_ERROR) return TEXT_COLLATOR.compare(left.textValue, right.textValue);
  return left.numberValue - right.numberValue;
}

/**
 * 算出"排完之后第 `k` 行来自带内原第几行"。
 *
 * 稳定排序：同键的行保持原有相对顺序，显式带原下标当末位比较，不依赖运行时的 `sort` 是否稳定。
 * 空格无论升降都落在带尾（D7）。
 *
 * @param keys 带内逐行的排序键，顺序就是现场行序。
 * @param direction `asc` 升序（默认）、`desc` 降序；只反转非空格之间的次序。
 * @returns 长度与 `keys` 相同的置换；置换为恒等（已经是目标次序）时调用方应报"无改动可做"。
 */
export function sortPermutation(keys: SortKey[], direction: "asc" | "desc" = "asc"): number[] {
  const sign = direction === "desc" ? -1 : 1;
  const ranked = keys.map((key, index) => ({ key, index }));
  ranked.sort((left, right) => {
    const leftBlank = left.key.rank === RANK_BLANK;
    const rightBlank = right.key.rank === RANK_BLANK;
    if (leftBlank || rightBlank) return leftBlank && rightBlank ? left.index - right.index : leftBlank ? 1 : -1;
    const compared = compareKeysAscending(left.key, right.key);
    return compared !== 0 ? sign * compared : left.index - right.index;
  });
  return ranked.map((item) => item.index);
}

/**
 * 置换是否等于"不用排"。
 *
 * @param permutation `sortPermutation` 的产物。
 * @returns 每一项都指向自己时为 true（调用方据此拒绝产生一条什么都不改的 op）。
 */
export function isIdentityPermutation(permutation: number[]): boolean {
  return permutation.every((source, position) => source === position);
}
