/**
 * `@office-copilot/pdf-io` 公开入口：`.pdf` 的文字读模型 + W1 档正文改写（增量更新保真）。
 *
 * 测试夹具（自造 PDF 的构造器）走 `@office-copilot/pdf-io/testing` 那个子路径，不从这里出（同 pptx-io 的理由）。
 *
 * 依赖方向（`AGENTS.md` §5.1）：只依赖 `@cantoo/pdf-lib`、Node 自带 zlib 与 `@office-copilot/shared`（截断上送文字那一份），
 * 与其它 IO 包平级、互不依赖。
 */
export * from "./code-map.js";
export * from "./content-stream.js";
export * from "./strings.js";
export * from "./pdf-file.js";
export * from "./analysis.js";
