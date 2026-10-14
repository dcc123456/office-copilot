import "dotenv/config";
import fs from "node:fs";
import path from "node:path";
import pc from "picocolors";
import type { DocxFile } from "@office-copilot/docx-io";

/**
 * Every CLI command that edits a document uses these helpers:
 *  - open a file as a DocxFile
 *  - run the mutation
 *  - save to --output or back to the input, after making a .bak copy
 */

export interface LoadedDoc {
  file: DocxFile;
  source: string;
  bytes: Uint8Array;
}

export async function loadDocx(inputPath: string): Promise<LoadedDoc> {
  const abs = path.resolve(inputPath);
  if (!fs.existsSync(abs)) {
    throw new Error(`文件不存在：${abs}`);
  }
  const bytes = await fs.promises.readFile(abs);
  const { DocxFile } = await import("@office-copilot/docx-io");
  const file = await DocxFile.fromBytes(new Uint8Array(bytes));
  return { file, source: abs, bytes: new Uint8Array(bytes) };
}

export async function saveDoc(
  doc: LoadedDoc,
  output: string | undefined,
  opts: { backup?: boolean } = {},
): Promise<string> {
  const out = path.resolve(output ?? doc.source);
  if (opts.backup !== false && out === doc.source) {
    const bak = `${out}.bak`;
    await fs.promises.writeFile(bak, doc.bytes);
    log(`已备份原文件 → ${pc.dim(bak)}`);
  }
  const outBytes = await doc.file.toBytes();
  await fs.promises.writeFile(out, outBytes);
  return out;
}

export function log(msg: string): void {
  // eslint-disable-next-line no-console
  console.log(msg);
}

export function warn(msg: string): void {
  // eslint-disable-next-line no-console
  console.warn(pc.yellow("⚠ ") + msg);
}

export function fail(msg: string, code = 1): never {
  // eslint-disable-next-line no-console
  console.error(pc.red("✖ ") + msg);
  process.exit(code);
}

/**
 * `--json` 开关的统一出口：给了 JSON 就打结构化对象，否则走人读渲染。
 *
 * @param json commander 的 `--json`（未给时是 undefined，所以参数可为假）。
 * @param structured JSON 分支要输出的对象，调用方自己保证它含 `ok` 之类的判定字段。
 * @param render 人读分支的渲染函数。
 */
export function emit(json: boolean | undefined, structured: unknown, render: () => void): void {
  if (json) log(JSON.stringify(structured, null, 2));
  else render();
}

/**
 * 把一段文字裁到 `max` 个字符，超出用省略号收尾（终端一次刷一整页正文没人会读）。
 *
 * @param text 原文。
 * @param max 保留的最大字符数。
 * @returns 截断后的串；不超过 `max` 时原样返回。
 */
export function clip(text: string, max: number): string {
  return text.length > max ? text.slice(0, max) + "…" : text;
}
