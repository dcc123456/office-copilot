/**
 * 字体候选值：工具栏、右键菜单的「字体(F)…」对话框共用这一份，
 * 避免出现第二套字号/颜色表（字号是 CSS px，非 Word 的半磅，见 AGENTS §2.1）。
 */

export interface FontOption {
  label: string;
  value: string;
}

export const FONT_FAMILIES: FontOption[] = [
  { label: "默认", value: "" },
  { label: "宋体", value: "SimSun, 宋体, serif" },
  { label: "黑体", value: "SimHei, 黑体, sans-serif" },
  { label: "微软雅黑", value: '"Microsoft YaHei", 微软雅黑, sans-serif' },
  { label: "苹方", value: '"PingFang SC", sans-serif' },
  { label: "楷体", value: "KaiTi, 楷体, serif" },
  { label: "仿宋", value: "FangSong, 仿宋, serif" },
  { label: "Arial", value: "Arial, sans-serif" },
  { label: "Times New Roman", value: '"Times New Roman", serif' },
  { label: "Consolas", value: "Consolas, monospace" },
];

export const FONT_SIZES: FontOption[] = [
  { label: "默认", value: "" },
  { label: "10", value: "10px" },
  { label: "12", value: "12px" },
  { label: "14 (五号)", value: "14px" },
  { label: "16 (小四)", value: "16px" },
  { label: "18 (三号)", value: "18px" },
  { label: "20", value: "20px" },
  { label: "24", value: "24px" },
  { label: "28", value: "28px" },
  { label: "32", value: "32px" },
];

/** 文字底色候选，空串代表清除。 */
export const HIGHLIGHT_COLORS = [
  "", "#ffff00", "#00ff00", "#00ffff", "#ff00ff", "#ff8c00",
  "#ff6b6b", "#a8e6cf", "#dcedc1", "#ffd3b6",
];

/** 字体颜色候选，空串代表清除。 */
export const FONT_COLORS = [
  "", "#000000", "#666666", "#a4262c", "#d13438", "#107c10", "#0078d4",
  "#5c2d91", "#ca5010", "#038387",
];
