import type { Extensions } from "@tiptap/core";
import StarterKit from "@tiptap/starter-kit";
import Underline from "@tiptap/extension-underline";
import TextAlign from "@tiptap/extension-text-align";
import Link from "@tiptap/extension-link";
import Table from "@tiptap/extension-table";
import TableRow from "@tiptap/extension-table-row";
import TableCell from "@tiptap/extension-table-cell";
import TableHeader from "@tiptap/extension-table-header";
import Placeholder from "@tiptap/extension-placeholder";
import CharacterCount from "@tiptap/extension-character-count";
import { TextStyle, TextStyleCommands } from "./TextStyle.js";
import { VertAlign } from "./VertAlign.js";
import { Revision } from "./Revision.js";
import { Drawing, PageBreak } from "./Drawing.js";
import { PagedLayout } from "./PagedLayout.js";
import { FloatingImageDrag } from "./FloatingImageDrag.js";
import { FloatBox, FloatGroup, FloatingTextLayout } from "./FloatingTextBox.js";
import { Field } from "./Field.js";
import { FieldBegin, FieldEnd } from "./FieldBoundary.js";
import { CaptionNumbering } from "./CaptionNumbering.js";
import { BookmarkEnd, BookmarkStart } from "./Bookmark.js";
import { FootnoteRefs } from "./FootnoteRef.js";
import { CommentAnchors, CommentCommands } from "./Comment.js";
import { ParagraphFormatting } from "./ParagraphFormatting.js";
import { TableColumnResize } from "./TableColumnResize.js";
import { FormatPainter } from "./FormatPainter.js";

/**
 * The editor configuration, in one place so `App.tsx` builds the document with
 * exactly the schema the tests assert against. A test schema missing the image or
 * table nodes hides the bugs that wrecked real documents.
 */
export const EDITOR_EXTENSIONS: Extensions = [
  StarterKit.configure({ heading: { levels: [1, 2, 3, 4, 5, 6] } }),
  Underline,
  VertAlign,
  Revision,
  TextAlign.configure({ types: ["heading", "paragraph"] }),
  Link.configure({ openOnClick: false, autolink: true }),
  Drawing,
  PageBreak,
  PagedLayout,
  FloatingImageDrag,
  FloatGroup,
  FloatBox,
  FloatingTextLayout,
  Field,
  FieldBegin,
  FieldEnd,
  CaptionNumbering,
  BookmarkStart,
  BookmarkEnd,
  ...FootnoteRefs,
  ...CommentAnchors,
  CommentCommands,
  ParagraphFormatting,
  TableColumnResize,
  FormatPainter,
  Table.configure({ resizable: true }),
  TableRow,
  TableCell,
  TableHeader,
  Placeholder.configure({
    placeholder: "开始输入… 或按 ⌥⌘L 打开 AI 助手",
  }),
  CharacterCount,
  TextStyle,
  TextStyleCommands,
];
