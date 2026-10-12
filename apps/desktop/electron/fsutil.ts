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
