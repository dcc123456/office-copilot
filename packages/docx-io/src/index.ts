export { DocxFile } from "./docx.js";
export type {
  DocxParagraph,
  ReplaceResult,
  Snapshot,
  SectionGeometry,
  SectionColumns,
  SectionPageBorders,
  SectionLineNumbers,
  PageBorderSide,
  DocStyle,
  PmLoadResult,
} from "./docx.js";
export * as xml from "./xml.js";
export {
  ooxmlToPm,
  pmToOoxml,
  ptToCssFontSize,
  cssFontSizeToPt,
  ptToCssLetterSpacing,
  cssLetterSpacingToPt,
  cssPxToPt,
  cmToPt,
  ptToCm,
  cssColorToOoxmlHex,
  canonColor,
  lineHeightSpacingAttrs,
} from "./ooxml-pm.js";
export type { LinkStyle, RunStyleProps } from "./ooxml-pm.js";
export { NumberingMap } from "./numbering.js";
export {
  StyleMap,
  fitStyleId,
  headingLookupFromPaneStyles,
} from "./styles.js";
export { mergeStylesXml, nextStyleId } from "./style-edit.js";
export type { StyleEdit } from "./style-edit.js";
export type { ResolvedStyle, ResolvedStyleFormat } from "./style-format.js";
export { nextFootnoteId } from "./footnotes.js";
export type {
  FootnoteBodies,
  FootnoteBody,
  FootnoteEntry,
  FootnoteKind,
  FootnoteStore,
} from "./footnotes.js";
// 批注侧栏要的：分 id（与书签同一个编号空间）＋ 跨 preload 用的条目形状。
// 以前只有 docx.test.ts 深 import 得到，渲染层按 §3 禁止深 import，所以从这里出去。
export { FALLBACK_COMMENT_AUTHOR, nextCommentId } from "./comments.js";
export type { CommentBody, CommentEntry, CommentStore } from "./comments.js";
// 脚注与批注窗格新建条目时要的空正文，实现早就在 pm-schema 里，只是没从包口出去。
export { emptyPmDoc } from "./pm-schema.js";
export type {
  StyleDefinition,
  StyleType,
  PaneStyle,
  HeadingStyleLookup,
} from "./styles.js";
export type {
  PPrAttrs,
  PMDoc,
  PMNode,
  PMMark,
  VerticalAlign,
  TabStop,
  TabStopAlignment,
  TabStopLeader,
  DropCap,
  DropCapPosition,
} from "./pm-schema.js";
