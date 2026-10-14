/**
 * `@office-copilot/pptx-io` 公开入口：`.pptx` 的读模型 + 保真读写器 + 结构化编辑 op + 结构审计 + 部件常量。
 *
 * 依赖方向（`AGENTS.md` §5.1）：只依赖 `@office-copilot/opc`，与 `docx-io`/`sheet-io` 平级、互不依赖。
 */
export * from "./parts.js";
export * from "./read.js";
export * from "./pptx-file.js";
export * from "./edit.js";
export * from "./audit.js";
