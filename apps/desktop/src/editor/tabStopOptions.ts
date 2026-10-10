import { ptToCm, type TabStop, type TabStopAlignment, type TabStopLeader } from "@office-copilot/docx-io";

/**
 * 制表位的候选值：段落面板的「制表位…」与「插入目录」对话框共用这一份，
 * 避免出现第二套前导符表（AGENTS §3.2）。档位取值来自 OOXML 的 `w:tab`
 * （ECMA-376 §17.3.1.80：`w:val` 与 `w:leader`），位置在编辑器里是 pt。
 */

export interface TabStopOption<T extends string> {
  value: T;
  label: string;
}

/** Word 前导符下拉框里的七项，`w:leader` 的取值。 */
export const TAB_LEADERS: TabStopOption<TabStopLeader>[] = [
  { value: "none", label: "（无）" },
  { value: "dot", label: "点线 …" },
  { value: "hyphen", label: "短划线 - - -" },
  { value: "dash", label: "虚线 — — —" },
  { value: "line", label: "实线 ———" },
  { value: "heavy", label: "重实线 ———" },
  { value: "middleDot", label: "中间点 · · ·" },
];

/**
 * 对齐方式：`w:val` 七档的中文名。只有 Word 对话框里点得出的五档进候选表，
 * `list`（编号生成的列表制表位）与 `clear`（取消样式继承来的那一格）能显示、
 * 不能新建——造一个 Word 不会用，抹一个又会改掉编号的排版本。
 */
const ALIGNMENT_LABELS: Record<TabStopAlignment, string> = {
  left: "左对齐",
  center: "居中",
  right: "右对齐",
  decimal: "小数点对齐",
  bar: "竖线",
  list: "列表制表位",
  clear: "取消的制表位",
};

const DIALOG_ALIGNMENTS: TabStopAlignment[] = [
  "left",
  "center",
  "right",
  "decimal",
  "bar",
];

/** 界面上可选的五档对齐，对应 Word 制表位对话框的一列单选。 */
export const TAB_ALIGNMENTS: TabStopOption<TabStopAlignment>[] = DIALOG_ALIGNMENTS.map(
  (value) => ({ value, label: ALIGNMENT_LABELS[value] }),
);

/**
 * 制表位对齐的中文名。
 *
 * @param alignment `w:val` 的取值；`clear`/`list` 这类界面给不出的值也认。
 * @returns 中文标签，表外值退回原值（导入文档的写法读不懂时不至于显示空白）。
 */
export function alignmentLabelOf(alignment: TabStopAlignment): string {
  return ALIGNMENT_LABELS[alignment] ?? alignment;
}

/**
 * 前导符的中文名。
 *
 * @param leader `w:leader` 的取值；没写这个属性（undefined）就是 Word 的"无"。
 * @returns 候选表里的标签，表外值退回原词。
 */
export function leaderLabelOf(leader: TabStopLeader | undefined): string {
  if (!leader) return "（无）";
  return TAB_LEADERS.find((option) => option.value === leader)?.label ?? leader;
}

/**
 * 清单里一行的文案，位置按 Word 的口径显示成厘米。
 *
 * @param stop 一个制表位，位置是 pt。
 * @returns 形如 `14.65 厘米 右对齐 点线 …`。
 */
export function stopLabelOf(stop: TabStop): string {
  return `${ptToCm(stop.positionPt).toFixed(2)} 厘米 ${alignmentLabelOf(
    stop.alignment,
  )} ${leaderLabelOf(stop.leader)}`;
}
