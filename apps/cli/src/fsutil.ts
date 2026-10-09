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
