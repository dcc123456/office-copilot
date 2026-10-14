/**
 * `.pptx` 部件路径、关系类型与内容类型常量（ECMA-376 Part 1 §18.3、OPC 关系 URI）。
 *
 * 只放格式相关的字符串：包/关系的读与登记原语都在 `@office-copilot/opc`，这里不再重抄一份（§3.1）。
 */

/** 主部件：幻灯片清单与页面尺寸都在这里。 */
export const PRESENTATION_PART = "ppt/presentation.xml";

/** 一条关系 `@Type` 是不是"指向幻灯片"。用 URI 结尾判断，不比对整个命名空间串。 */
export function isSlideRelationship(type: string): boolean {
  return type.endsWith("/relationships/slide");
}

/** 关系类型 URI 的公共前缀，新建关系时拼上具体角色名。 */
export const PPTX_REL_TYPE_PREFIX = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/** 各部件的内容类型 Override 值（§2.10 新建部件要登记到 `[Content_Types].xml` 的就是这些串）。 */
export const PPTX_CONTENT_TYPES = {
  presentation: "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
  slide: "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
  slideLayout: "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml",
  slideMaster: "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml",
  notesSlide: "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml",
  theme: "application/vnd.openxmlformats-officedocument.theme+xml",
  image: "image/png",
} as const;
