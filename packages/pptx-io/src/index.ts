/**
 * `@office-copilot/pptx-io` 公开入口：`.pptx` 的读模型 + 保真读写器 + 结构化编辑 op + 结构审计 + 部件常量。
 *
 * 测试夹具（最小 deck 构造器）走 `@office-copilot/pptx-io/testing` 那个子路径，不从这里出：
 * 主进程打包要是不慎带上一份假幻灯片，没人会注意到，而它是纯死重。
 *
 * 依赖方向（`AGENTS.md` §5.1）：只依赖 `@office-copilot/opc`，与 `docx-io`/`sheet-io` 平级、互不依赖。
 */
export * from "./parts.js";
export * from "./units.js";
export * from "./read.js";
export * from "./analysis.js";
export * from "./pptx-file.js";
export * from "./edit.js";
export * from "./audit.js";
