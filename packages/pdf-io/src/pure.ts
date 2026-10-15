/**
 * `@office-copilot/pdf-io/pure`：渲染层能用的那一半公开入口——只含**不含 Node 能力**的模块。
 *
 * 分这一层只为一条红线（`AGENTS.md` §5.2）：渲染层一 import 包总入口，就把 `pdf-file.ts` 的
 * `node:zlib` 与 `@cantoo/pdf-lib` 一起打进浏览器包，运行时 `require is not defined` 整页白屏。
 * 这里转出的东西都是纯计算：数据类型与判定（`pdf-model.ts`）、AI 读面（`analysis.ts`）。
 * 读文件、改内容流、增量写盘仍旧只在 `@office-copilot/pdf-io` 那一个入口，由主进程调用。
 */
export * from "./pdf-model.js";
export * from "./analysis.js";
