import StarterKit from "@tiptap/starter-kit";
import Underline from "@tiptap/extension-underline";
import TextAlign from "@tiptap/extension-text-align";
import Link from "@tiptap/extension-link";
import { TextStyle, TextStyleCommands } from "./TextStyle.js";
import { Field } from "./Field.js";

/**
 * 嵌套小编辑器（页眉页脚、脚注窗格）共用的 schema。
 *
 * 正文档能用的行内能力都在这里，但没有表格、图片、分页与拖拽——Word 在这些位置
 * 同样不让放，所以这不是省事，是同一套限制。两处共用一份，扩展能力时只会漏一次。
 */
export const NESTED_EXTENSIONS = [
  StarterKit.configure({ heading: { levels: [1, 2, 3, 4, 5, 6] } }),
  Underline,
  TextAlign.configure({ types: ["heading", "paragraph"] }),
  TextStyle,
  TextStyleCommands,
  Field,
  Link.configure({ openOnClick: false }),
];
