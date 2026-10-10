export { DocxFile } from "./docx.js";
export type {
  DocxParagraph,
  ReplaceResult,
  Snapshot,
  SectionGeometry,
  SectionColumns,
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
} from "./pm-schema.js";
