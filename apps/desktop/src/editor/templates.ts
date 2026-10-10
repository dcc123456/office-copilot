import {
  cssFontSizeToPt,
  ptToCssFontSize,
  ptToCssLetterSpacing,
  type LinkStyle,
} from "@office-copilot/docx-io";
import type { StylePatch } from "../ai/docEdit.js";

/**
 * Document style templates for 美化排版, researched from published typography
 * specs (GB/T 9704-2012 for 公文, 学位论文 format requirements, and the usual
 * 企业/商务 conventions). Each template is a per-role `StylePatch`, so applying
 * one goes through the same proposal → diff → undo pipeline as any other edit.
 *
 * Units are the editor's, not Word's — see docEdit.StylePatch:
 *   fontSize     CSS px, produced by ptToCssFontSize() so a saved/reloaded
 *                document reads back the identical string (idempotency)
 *   letterSpacing CSS px from ptToCssLetterSpacing(); OOXML stores twentieths of
 *                a point, and only two decimals survive that round trip
 *   italic/underline/highlight live on the runs like bold does, so a block whose
 *                runs disagree keeps the value it has and leaves the rest alone
 *   lineHeight   "1.5" multiplier, "pt:29" for 固定值, "atLeast:20" for 最小值
 *   color        canonical lowercase #rrggbb (docEdit compares colours through
 *                canonColor, and so does the DOM once the mark round-trips)
 *   spacing*     points
 *   indentFirst  points; 2 字符 = 2 × 字号磅值
 *   page         millimetres (converted to px where the section is applied)
 *
 * Font families are a *single* name because docx-io writes the whole value
 * into w:rFonts' ascii/hAnsi/eastAsia slots — a CSS stack would be unresolvable
 * in Word, and reading back yields one name, which would break idempotency.
 */

export type TemplateRole = "title" | "h1" | "h2" | "h3" | "h4" | "h5" | "h6" | "body";

export interface DocTemplate {
  id: string;
  label: string;
  /** Provenance note shown under the label, e.g. the standard it follows. */
  hint: string;
  swatch: string;
  /** One-line parameter summary for the picker card. */
  spec: string;
  roles: Partial<Record<TemplateRole, StylePatch>>;
  /**
   * 西文字体（OOXML 的 `w:ascii`/`w:hAnsi`），整套模板统一：中文文档的阿拉伯
   * 数字与拉丁字母按惯例走 Times New Roman，汉字仍用各角色自己的 `fontFamily`。
   * 缺省 = 中西文同一个字体。
   */
  latin?: string | null;
  /** Page margins in mm; omitted = leave the document's own page setup alone. */
  page?: { top: number; right: number; bottom: number; left: number };
  chrome?: { footer?: "dash" | "center"; header?: "title" };
  /**
   * How the document paints hyperlinks — a document-level setting like 页边距,
   * because a link is not a paragraph role. Omitted means Word's own blue and
   * underline. It never rewrites body text, so no paragraph can disagree.
   */
  link?: LinkStyle;
  /** Three static lines the picker renders in the template's own style. */
  preview: { title: string; h1: string; body: string };
}

/* Chinese 字号 as points — the specs are written in 号, the editor in px. */
const WU_HAO = 10.5;
const XIAO_SI = 12;
const SI_HAO = 14;
const XIAO_SAN = 15;
const SAN_HAO = 16;
const XIAO_ER = 18;
const ER_HAO = 22;

const px = ptToCssFontSize;
/** 字间距，同 `px`：写模板数据时用磅，存进文档的是 CSS px。 */
const ls = ptToCssLetterSpacing;
/** 首行缩进 2 字符, in points, for a body of `pt`. */
const twoChars = (pt: number) => Math.round(pt * 2 * 10) / 10;

const SONG = "宋体";
const HEI = "黑体";
const FANGSONG = "仿宋";
const KAI = "楷体";
const YAHEI = "微软雅黑";
const TNR = "Times New Roman";

/* Ink. 公文/论文/合同 must print black; screen-first templates soften it. */
const INK = "#000000";
const BODY_INK = "#262626";
const NAVY = "#1f3864";
const BLUE = "#0b5394";
const TEAL = "#00695c";
const SLATE = "#37474f";
const SCREEN_INK = "#3f3f3f";

/**
 * CSS stack for the picker preview only. The document attr keeps the single
 * canonical name (that is what Word needs); on a Mac the browser falls back
 * through these instead of rendering 宋体 as default sans.
 */
export const FONT_FALLBACK: Record<string, string> = {
  [SONG]: '"Songti SC", "宋体", SimSun, serif',
  [HEI]: '"Heiti SC", "黑体", SimHei, sans-serif',
  [FANGSONG]: '"STFangsong", "仿宋", FangSong, serif',
  [KAI]: '"Kaiti SC", "楷体", KaiTi, serif',
  [YAHEI]: '"Microsoft YaHei", "PingFang SC", 微软雅黑, sans-serif',
};

export function fontCss(name: string | null | undefined): string | undefined {
  if (!name) return undefined;
  return FONT_FALLBACK[name] ?? name;
}

export const TEMPLATES: readonly DocTemplate[] = [
  {
    id: "gongwen",
    label: "公文",
    hint: "GB/T 9704-2012 · 正文三号仿宋、固定 29 磅行距、首行缩进 2 字符",
    swatch: "#b3261e",
    spec: "仿宋 三号 · 固定 29 磅 · 黑字 · 上37/下35/左28/右26",
    roles: {
      // 文件标题用方正小标宋，本机通常没有，黑体是最通用的替代。
      // 公文一律不斜、不加下划线，所以把这两项写成 false 而不是留空。
      title: {
        fontFamily: HEI,
        fontSize: px(ER_HAO),
        letterSpacing: ls(0.5),
        color: INK,
        bold: false,
        italic: false,
        underline: false,
        align: "center",
        lineHeight: `pt:${ER_HAO + 7}`,
        indentFirst: 0,
        spacingBefore: 0,
        spacingAfter: 0,
      },
      h1: {
        fontFamily: HEI,
        fontSize: px(SAN_HAO),
        color: INK,
        bold: false,
        italic: false,
        underline: false,
        align: null,
        lineHeight: "pt:29",
        indentFirst: twoChars(SAN_HAO),
        spacingBefore: 0,
        spacingAfter: 0,
      },
      h2: {
        fontFamily: KAI,
        fontSize: px(SAN_HAO),
        color: INK,
        bold: false,
        italic: false,
        underline: false,
        align: null,
        lineHeight: "pt:29",
        indentFirst: twoChars(SAN_HAO),
        spacingBefore: 0,
        spacingAfter: 0,
      },
      h3: {
        fontFamily: FANGSONG,
        fontSize: px(SAN_HAO),
        color: INK,
        bold: true,
        italic: false,
        underline: false,
        align: null,
        lineHeight: "pt:29",
        indentFirst: twoChars(SAN_HAO),
        spacingBefore: 0,
        spacingAfter: 0,
      },
      body: {
        fontFamily: FANGSONG,
        fontSize: px(SAN_HAO),
        color: INK,
        italic: false,
        underline: false,
        align: "justify",
        lineHeight: "pt:29",
        indentFirst: twoChars(SAN_HAO),
        spacingBefore: 0,
        spacingAfter: 0,
      },
    },
    latin: TNR,
    page: { top: 37, right: 26, bottom: 35, left: 28 },
    chrome: { footer: "dash" },
    link: { color: INK, underline: true },
    preview: {
      title: "关于开展年度检查的通知",
      h1: "一、检查范围",
      body: "各有关单位要按照职责分工，认真组织做好自查工作，并于本年度年底前将自查情况书面报送。",
    },
  },
  {
    id: "thesis",
    label: "学位论文",
    hint: "本科/硕士论文通行规范 · 宋体小四、行距不低于 20 磅",
    swatch: "#1f4e79",
    spec: "宋体 小四 · 最小值 20 磅 · 黑字 · 上30/下25/左30/右25",
    roles: {
      title: {
        fontFamily: HEI,
        fontSize: px(XIAO_ER),
        color: INK,
        bold: true,
        align: "center",
        lineHeight: "2",
        indentFirst: 0,
        spacingBefore: 12,
        spacingAfter: 12,
      },
      h1: {
        fontFamily: HEI,
        fontSize: px(SAN_HAO),
        color: INK,
        bold: true,
        align: "left",
        lineHeight: "1.5",
        indentFirst: 0,
        spacingBefore: 12,
        spacingAfter: 6,
      },
      h2: {
        fontFamily: HEI,
        fontSize: px(SI_HAO),
        color: INK,
        bold: true,
        align: "left",
        lineHeight: "1.5",
        indentFirst: 0,
        spacingBefore: 6,
        spacingAfter: 6,
      },
      h3: {
        fontFamily: HEI,
        fontSize: px(XIAO_SI),
        color: INK,
        bold: true,
        align: "left",
        lineHeight: "1.5",
        indentFirst: 0,
        spacingBefore: 6,
        spacingAfter: 0,
      },
      body: {
        fontFamily: SONG,
        fontSize: px(XIAO_SI),
        color: INK,
        align: "justify",
        lineHeight: "atLeast:20",
        indentFirst: twoChars(XIAO_SI),
        spacingBefore: 0,
        spacingAfter: 0,
      },
    },
    latin: TNR,
    page: { top: 30, right: 25, bottom: 25, left: 30 },
    chrome: { footer: "center", header: "title" },
    link: { color: INK, underline: true },
    preview: {
      title: "基于深度学习的图像分割研究",
      h1: "1 相关工作",
      body: "本章首先对现有方法进行综述，并从实验角度比较各类方法在公开数据集上的表现与不足。",
    },
  },
  {
    id: "business",
    label: "商务报告",
    hint: "企业汇报版式 · 五号雅黑、适度留白",
    swatch: "#2e7d32",
    spec: "微软雅黑 五号 · 1.4 倍 · 深蓝标题 · 四边 25.4",
    roles: {
      title: {
        fontFamily: YAHEI,
        fontSize: px(XIAO_ER),
        color: NAVY,
        bold: true,
        align: "center",
        lineHeight: "1.4",
        indentFirst: 0,
        spacingBefore: 0,
        spacingAfter: 12,
      },
      h1: {
        fontFamily: YAHEI,
        fontSize: px(SI_HAO),
        color: NAVY,
        bold: true,
        align: "left",
        lineHeight: "1.4",
        indentFirst: 0,
        spacingBefore: 12,
        spacingAfter: 6,
      },
      h2: {
        fontFamily: YAHEI,
        fontSize: px(XIAO_SI),
        color: NAVY,
        bold: true,
        align: "left",
        lineHeight: "1.4",
        indentFirst: 0,
        spacingBefore: 8,
        spacingAfter: 4,
      },
      h3: {
        fontFamily: YAHEI,
        fontSize: px(WU_HAO),
        color: SLATE,
        bold: true,
        align: "left",
        lineHeight: "1.4",
        indentFirst: 0,
        spacingBefore: 6,
        spacingAfter: 3,
      },
      body: {
        fontFamily: YAHEI,
        fontSize: px(WU_HAO),
        color: BODY_INK,
        align: "left",
        lineHeight: "1.4",
        indentFirst: 0,
        spacingBefore: 0,
        spacingAfter: 6,
      },
    },
    page: { top: 25.4, right: 25.4, bottom: 25.4, left: 25.4 },
    chrome: { footer: "center", header: "title" },
    link: { color: NAVY, underline: true },
    preview: {
      title: "2026 年第三季度经营分析",
      h1: "一、收入概览",
      body: "本季度实现收入 1.2 亿元，同比增长 18%，其中订阅业务占比首次过半，毛利率维持在 64%。",
    },
  },
  {
    id: "regulation",
    label: "企业制度",
    hint: "规章制度/管理办法 · 仿宋小四、固定 26 磅",
    swatch: "#5d4037",
    spec: "仿宋 小四 · 固定 26 磅 · 黑字 · 上37/下35/左28/右26",
    roles: {
      title: {
        fontFamily: HEI,
        fontSize: px(ER_HAO),
        color: INK,
        bold: false,
        align: "center",
        lineHeight: "pt:33",
        indentFirst: 0,
        spacingBefore: 0,
        spacingAfter: 12,
      },
      h1: {
        fontFamily: HEI,
        fontSize: px(SAN_HAO),
        color: INK,
        bold: false,
        align: null,
        lineHeight: "pt:26",
        indentFirst: twoChars(XIAO_SI),
        spacingBefore: 6,
        spacingAfter: 0,
      },
      h2: {
        fontFamily: HEI,
        fontSize: px(XIAO_SI),
        color: INK,
        bold: false,
        align: null,
        lineHeight: "pt:26",
        indentFirst: twoChars(XIAO_SI),
        spacingBefore: 0,
        spacingAfter: 0,
      },
      h3: {
        fontFamily: HEI,
        fontSize: px(XIAO_SI),
        color: INK,
        bold: true,
        align: null,
        lineHeight: "pt:26",
        indentFirst: twoChars(XIAO_SI),
        spacingBefore: 0,
        spacingAfter: 0,
      },
      body: {
        fontFamily: FANGSONG,
        fontSize: px(XIAO_SI),
        color: INK,
        align: "justify",
        lineHeight: "pt:26",
        indentFirst: twoChars(XIAO_SI),
        spacingBefore: 0,
        spacingAfter: 0,
      },
    },
    latin: TNR,
    page: { top: 37, right: 26, bottom: 35, left: 28 },
    link: { color: INK, underline: true },
    preview: {
      title: "印章使用管理办法",
      h1: "第一章 总则",
      body: "第一条 为规范公司印章的刻制、保管和使用，防范经营风险，根据国家有关规定，结合公司实际，制定本办法。",
    },
  },
  {
    id: "contract",
    label: "合同协议",
    hint: "合同/协议 · 宋体小四、固定 26 磅、条款缩进",
    swatch: "#37474f",
    spec: "宋体 小四 · 固定 26 磅 · 黑字 · 上30/下28/左30/右28",
    roles: {
      title: {
        fontFamily: SONG,
        fontSize: px(ER_HAO),
        color: INK,
        bold: true,
        align: "center",
        lineHeight: "pt:33",
        indentFirst: 0,
        spacingBefore: 0,
        spacingAfter: 12,
      },
      h1: {
        fontFamily: SONG,
        fontSize: px(XIAO_SI),
        color: INK,
        bold: true,
        align: null,
        lineHeight: "pt:26",
        indentFirst: twoChars(XIAO_SI),
        spacingBefore: 6,
        spacingAfter: 0,
      },
      h2: {
        fontFamily: SONG,
        fontSize: px(XIAO_SI),
        color: INK,
        bold: true,
        align: null,
        lineHeight: "pt:26",
        indentFirst: twoChars(XIAO_SI),
        spacingBefore: 0,
        spacingAfter: 0,
      },
      h3: {
        fontFamily: SONG,
        fontSize: px(XIAO_SI),
        color: INK,
        bold: false,
        align: null,
        lineHeight: "pt:26",
        indentFirst: twoChars(XIAO_SI),
        spacingBefore: 0,
        spacingAfter: 0,
      },
      body: {
        fontFamily: SONG,
        fontSize: px(XIAO_SI),
        color: INK,
        align: "justify",
        lineHeight: "pt:26",
        indentFirst: twoChars(XIAO_SI),
        spacingBefore: 0,
        spacingAfter: 6,
      },
    },
    latin: TNR,
    page: { top: 30, right: 28, bottom: 30, left: 28 },
    chrome: { footer: "center" },
    link: { color: INK, underline: true },
    preview: {
      title: "技术服务合同",
      h1: "第一条 服务内容",
      body: "甲方委托乙方提供系统开发与维护服务，具体服务内容、交付标准及验收方式见附件一《服务清单》。",
    },
  },
  {
    id: "resume",
    label: "简历单页",
    hint: "紧凑排版 · 无首行缩进、小段距",
    swatch: "#00695c",
    spec: "微软雅黑 五号 · 1.35 倍 · 青绿标题 · 上下12.7/左右15",
    roles: {
      title: {
        fontFamily: YAHEI,
        fontSize: px(XIAO_SAN),
        color: TEAL,
        bold: true,
        align: "center",
        lineHeight: "1.35",
        indentFirst: 0,
        spacingBefore: 0,
        spacingAfter: 6,
      },
      h1: {
        fontFamily: YAHEI,
        fontSize: px(XIAO_SI),
        color: TEAL,
        bold: true,
        align: "left",
        lineHeight: "1.35",
        indentFirst: 0,
        spacingBefore: 8,
        spacingAfter: 4,
      },
      h2: {
        fontFamily: YAHEI,
        fontSize: px(WU_HAO),
        color: SLATE,
        bold: true,
        align: "left",
        lineHeight: "1.35",
        indentFirst: 0,
        spacingBefore: 4,
        spacingAfter: 2,
      },
      h3: {
        fontFamily: YAHEI,
        fontSize: px(WU_HAO),
        color: SLATE,
        bold: true,
        align: "left",
        lineHeight: "1.35",
        indentFirst: 0,
        spacingBefore: 2,
        spacingAfter: 2,
      },
      body: {
        fontFamily: YAHEI,
        fontSize: px(WU_HAO),
        color: BODY_INK,
        align: "left",
        lineHeight: "1.35",
        indentFirst: 0,
        spacingBefore: 0,
        spacingAfter: 3,
      },
    },
    page: { top: 12.7, right: 15, bottom: 12.7, left: 15 },
    link: { color: TEAL, underline: false },
    preview: {
      title: "张晓黎 · 高级前端工程师",
      h1: "工作经历",
      body: "主导编辑器内核重构，将插件加载改为惰性初始化，冷启动从 1.8s 降至 420ms。",
    },
  },
  {
    id: "techdoc",
    label: "技术文档",
    hint: "设计/接口文档 · 分级清晰、段后留白",
    swatch: "#0b5394",
    spec: "微软雅黑 五号 · 1.6 倍 · 蓝标题 · 上下25/左右20",
    roles: {
      title: {
        fontFamily: YAHEI,
        fontSize: px(XIAO_ER),
        color: BLUE,
        bold: true,
        align: "center",
        lineHeight: "1.4",
        indentFirst: 0,
        spacingBefore: 0,
        spacingAfter: 12,
      },
      h1: {
        fontFamily: YAHEI,
        fontSize: px(XIAO_SAN),
        color: BLUE,
        bold: true,
        align: "left",
        lineHeight: "1.6",
        indentFirst: 0,
        spacingBefore: 14,
        spacingAfter: 6,
      },
      h2: {
        fontFamily: YAHEI,
        fontSize: px(SI_HAO),
        color: BLUE,
        bold: true,
        align: "left",
        lineHeight: "1.6",
        indentFirst: 0,
        spacingBefore: 10,
        spacingAfter: 4,
      },
      h3: {
        fontFamily: YAHEI,
        fontSize: px(XIAO_SI),
        color: SLATE,
        bold: true,
        align: "left",
        lineHeight: "1.6",
        indentFirst: 0,
        spacingBefore: 8,
        spacingAfter: 4,
      },
      body: {
        fontFamily: YAHEI,
        fontSize: px(WU_HAO),
        color: BODY_INK,
        align: "left",
        lineHeight: "1.6",
        indentFirst: 0,
        spacingBefore: 0,
        spacingAfter: 6,
      },
    },
    page: { top: 25, right: 20, bottom: 25, left: 20 },
    chrome: { footer: "center", header: "title" },
    link: { color: BLUE, underline: true },
    preview: {
      title: "排版引擎设计说明",
      h1: "1. 数据流",
      body: "提案由 blocksOf() 生成，经 simulateProposal() 校验后交给 commitOps()，每步都带可逆操作。",
    },
  },
  {
    id: "wechat",
    label: "公众号推文",
    hint: "移动端阅读 · 12pt 大字、宽行距、不缩进",
    swatch: "#e8710a",
    spec: "微软雅黑 小四 · 1.8 倍 · 深灰正文 · 上下20/左右16",
    roles: {
      title: {
        fontFamily: YAHEI,
        fontSize: px(SAN_HAO),
        color: INK,
        bold: true,
        align: "center",
        lineHeight: "1.8",
        indentFirst: 0,
        spacingBefore: 0,
        spacingAfter: 12,
      },
      h1: {
        fontFamily: YAHEI,
        fontSize: px(SI_HAO),
        color: INK,
        bold: true,
        align: "center",
        lineHeight: "1.8",
        indentFirst: 0,
        spacingBefore: 12,
        spacingAfter: 12,
      },
      h2: {
        fontFamily: YAHEI,
        fontSize: px(XIAO_SI),
        color: SLATE,
        bold: true,
        align: "left",
        lineHeight: "1.8",
        indentFirst: 0,
        spacingBefore: 10,
        spacingAfter: 8,
      },
      h3: {
        fontFamily: YAHEI,
        fontSize: px(XIAO_SI),
        color: SLATE,
        bold: true,
        align: "left",
        lineHeight: "1.8",
        indentFirst: 0,
        spacingBefore: 8,
        spacingAfter: 6,
      },
      body: {
        fontFamily: YAHEI,
        fontSize: px(XIAO_SI),
        color: SCREEN_INK,
        align: "left",
        lineHeight: "1.8",
        indentFirst: 0,
        spacingBefore: 0,
        spacingAfter: 12,
      },
    },
    page: { top: 20, right: 16, bottom: 20, left: 16 },
    link: { color: SCREEN_INK, underline: false },
    preview: {
      title: "我们重写了一遍排版引擎",
      h1: "为什么要做这件事",
      body: "先说结论：把「模型直接改文档」换成「提案 + 差异 + 确认」之后，绝大多数误改都在提交前被拦住了。",
    },
  },
  {
    id: "book",
    label: "书籍长文",
    hint: "出版/长文阅读 · 五号宋体、行距最小 18 磅、两端对齐",
    swatch: "#6d4c41",
    spec: "宋体 五号 · 最小值 18 磅 · 黑字 · 上下25/左右20",
    roles: {
      title: {
        fontFamily: SONG,
        fontSize: px(ER_HAO),
        color: INK,
        bold: false,
        align: "center",
        lineHeight: "1.5",
        indentFirst: 0,
        spacingBefore: 18,
        spacingAfter: 18,
      },
      h1: {
        fontFamily: SONG,
        fontSize: px(SAN_HAO),
        color: INK,
        bold: true,
        align: "center",
        lineHeight: "1.5",
        indentFirst: 0,
        spacingBefore: 14,
        spacingAfter: 10,
      },
      h2: {
        fontFamily: SONG,
        fontSize: px(SI_HAO),
        color: INK,
        bold: true,
        align: "left",
        lineHeight: "1.5",
        indentFirst: 0,
        spacingBefore: 8,
        spacingAfter: 4,
      },
      h3: {
        fontFamily: SONG,
        fontSize: px(XIAO_SI),
        color: INK,
        bold: true,
        align: "left",
        lineHeight: "1.5",
        indentFirst: 0,
        spacingBefore: 6,
        spacingAfter: 2,
      },
      body: {
        fontFamily: SONG,
        fontSize: px(WU_HAO),
        color: INK,
        align: "justify",
        lineHeight: "atLeast:18",
        indentFirst: twoChars(WU_HAO),
        spacingBefore: 0,
        spacingAfter: 0,
      },
    },
    latin: TNR,
    page: { top: 25, right: 20, bottom: 25, left: 20 },
    chrome: { footer: "center" },
    link: { color: INK, underline: false },
    preview: {
      title: "第一章 汉字的格子",
      h1: "版心的由来",
      body: "印刷厂的老师傅讲过，版面要先算版心：四周留白不是为了好看，而是为了让手指有地方捏住书页。",
    },
  },
  {
    id: "minutes",
    label: "会议纪要",
    hint: "纪要/周报 · 标题居中、紧凑段距",
    swatch: "#455a64",
    spec: "微软雅黑 五号 · 1.5 倍 · 灰蓝标题 · 四边 20",
    roles: {
      title: {
        fontFamily: YAHEI,
        fontSize: px(XIAO_SAN),
        color: SLATE,
        bold: true,
        align: "center",
        lineHeight: "1.5",
        indentFirst: 0,
        spacingBefore: 0,
        spacingAfter: 8,
      },
      h1: {
        fontFamily: YAHEI,
        fontSize: px(XIAO_SI),
        color: SLATE,
        bold: true,
        align: "left",
        lineHeight: "1.5",
        indentFirst: 0,
        spacingBefore: 10,
        spacingAfter: 4,
      },
      h2: {
        fontFamily: YAHEI,
        fontSize: px(WU_HAO),
        color: SLATE,
        bold: true,
        align: "left",
        lineHeight: "1.5",
        indentFirst: 0,
        spacingBefore: 6,
        spacingAfter: 2,
      },
      h3: {
        fontFamily: YAHEI,
        fontSize: px(WU_HAO),
        color: BODY_INK,
        bold: true,
        align: "left",
        lineHeight: "1.5",
        indentFirst: 0,
        spacingBefore: 4,
        spacingAfter: 2,
      },
      body: {
        fontFamily: YAHEI,
        fontSize: px(WU_HAO),
        color: BODY_INK,
        align: "left",
        lineHeight: "1.5",
        indentFirst: 0,
        spacingBefore: 0,
        spacingAfter: 4,
      },
    },
    page: { top: 20, right: 20, bottom: 20, left: 20 },
    chrome: { header: "title" },
    link: { color: SLATE, underline: true },
    preview: {
      title: "9 月产品例会纪要",
      h1: "一、结论",
      body: "排版模板先行上线桌面端；Word 侧保持现状，待模板数据稳定后再评估是否复用同一份规格。",
    },
  },
];

export const ROLE_LABELS: Record<TemplateRole, string> = {
  title: "标题",
  h1: "H1",
  h2: "H2",
  h3: "H3",
  h4: "H4",
  h5: "H5",
  h6: "H6",
  body: "正文",
};

/** Heading roles, shallowest first — the order `completeRoles` fills in. */
const HEADING_ROLES = ["h1", "h2", "h3", "h4", "h5", "h6"] as const;

/** 中文字号 by descending point size; the derivation steps down this ladder. */
const HAO_PT = [22, 18, 16, 15, 14, 12, 10.5, 9];

function nextSmallerPt(pt: number): number {
  const smaller = HAO_PT.filter((p) => p < pt);
  return smaller.length ? (smaller[0] as number) : pt;
}

/**
 * Every role with a patch, including the levels the spec never defines.
 *
 * A 公文 or 论文 standard names three or four heading levels; a real document
 * goes deeper, and those paragraphs still have to land somewhere. The derivation
 * copies the template's own deepest written heading — so line spacing, indent and
 * alignment keep the document's grid — and steps the 字号 down one 号 per level,
 * with the tail (h6) dropping bold so the levels still separate themselves.
 */
export function completeRoles(t: DocTemplate): Record<TemplateRole, StylePatch> {
  const out: Partial<Record<TemplateRole, StylePatch>> = { ...t.roles };
  let deepest: StylePatch | undefined;
  for (const role of HEADING_ROLES) {
    if (out[role]) deepest = out[role];
  }
  const body = out.body ?? {};
  const donor: StylePatch = deepest ?? { ...body, align: "left", indentFirst: 0, bold: true };
  let pt = cssFontSizeToPt(donor.fontSize ?? "") ?? 12;
  for (const role of ["h4", "h5", "h6"] as const) {
    if (out[role]) {
      pt = cssFontSizeToPt(out[role]!.fontSize ?? "") ?? pt;
      continue;
    }
    if (role !== "h4") pt = nextSmallerPt(pt);
    out[role] = {
      ...donor,
      fontSize: px(pt),
      bold: role === "h6" ? false : donor.bold ?? true,
      // 第五、六级标题在规范里几乎不存在，斜体是让它们最省事的区分方式。
      ...(role === "h6" ? { italic: true } : null),
    };
  }
  for (const role of HEADING_ROLES) {
    if (!out[role]) out[role] = { ...donor };
  }
  if (t.latin) {
    // 西文字体是整套模板的统一设定，不必十个角色各写一遍。
    for (const role of Object.keys(out) as TemplateRole[]) {
      const patch = out[role];
      if (patch && patch.fontFamilyWest === undefined) {
        out[role] = { ...patch, fontFamilyWest: t.latin };
      }
    }
  }
  return out as Record<TemplateRole, StylePatch>;
}

export function templateById(id: string): DocTemplate | undefined {
  return TEMPLATES.find((t) => t.id === id);
}

/** Page margins in millimetres → CSS px at 96 DPI (the section's unit). */
export function mmToPx(mm: number): number {
  return Math.round((mm / 25.4) * 96 * 10) / 10;
}
