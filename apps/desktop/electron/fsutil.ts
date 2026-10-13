/**
 * 主进程里几个 handler 共用的文件系统小工具（同一逻辑出现第二处才抽，`AGENTS.md` §3.2）。
 */
import fs from "node:fs/promises";

/**
 * 探测路径是否存在。
 *
 * @param p 绝对路径。
 * @returns 存在为 true；因权限之外的原因读不到（含不存在）为 false，调用方按"没有这个文件"处理。
 */
export async function fileExists(p: string): Promise<boolean> {
  try {
    await fs.access(p);
    return true;
  } catch {
    return false;
  }
}

/**
 * 原子写回一个文件：先把整份字节写进 `filePath + ".tmp"`，再 `rename` 换上去。
 *
 * 用户正在打开的那份文件不能出现"半个文件"：`rename` 在同一文件系统里是原子的，
 * 读侧要么看到旧内容要么看到新内容（`AGENTS.md` §2.4）。失败时清掉自己刚写的那个
 * `.tmp`，不把半成品留在用户目录里，再把**原来那个错误**抛给调用方去组中文消息。
 *
 * @param filePath 目标绝对路径（已存在的文件会被整体替换）。
 * @param bytes 要写出的完整字节；这里不做任何编码或换行处理。
 * @throws 写或改名失败时抛出底层错误（权限、目录不存在、目标是目录等），调用方负责转成给人读的一句话。
 */
export async function writeFileAtomically(filePath: string, bytes: Uint8Array): Promise<void> {
  const tmpPath = `${filePath}.tmp`;
  let isTmpWritten = false;
  try {
    await fs.writeFile(tmpPath, bytes);
    isTmpWritten = true;
    await fs.rename(tmpPath, filePath);
  } catch (err) {
    // `force` 只吞 ENOENT（写一半没落成）：别的错误还得让调用方看见。
    if (isTmpWritten) await fs.rm(tmpPath, { force: true });
    throw err;
  }
}
