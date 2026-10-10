export { DocxFile } from "./docx.js";
export type {
  DocxParagraph,
  ReplaceResult,
  Snapshot,
  SectionGeometry,
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
export type { FootnoteEntry, FootnoteKind, FootnoteStore } from "./footnotes.js";
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
  TabStop,
  TabStopAlignment,
  TabStopLeader,
} from "./pm-schema.js";
