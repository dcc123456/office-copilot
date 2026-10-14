/**
 * `@office-copilot/opc` 的公开入口：通用 XML 层 + OPC 包读取层。
 *
 * 分工：`xml.ts` 管"一段 OOXML 怎么解析/序列化"，`opc.ts` 管"一个 zip 包怎么列部件、查关系、读内容类型"，
 * `write.ts` 管"新建部件时 §2.10 的关系与内容类型怎么一起登记"。
 * 三者都与具体格式无关，`.docx`/`.xlsx`/`.pptx` 共用同一份（`AGENTS.md` §3.2：不许长第二套）。
 */
export * as xml from "./xml.js";
export * from "./xml.js";
export * from "./opc.js";
export * from "./write.js";
