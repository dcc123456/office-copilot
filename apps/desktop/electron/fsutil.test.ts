/**
 * `writeFileAtomically` 的验收（`AGENTS.md` §2.4：写盘必须是"写 `.tmp` → rename"）。
 *
 * 用**真文件系统**（`os.tmpdir()` 下开临时目录）而不是 mock `fs`：这条要证的正是
 * "用户正在打开的那份文件不会出现半个文件"，只有落到磁盘上才算实测。
 */
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { writeFileAtomically } from "./fsutil.js";

let dir = "";

beforeEach(async () => {
  dir = await fs.mkdtemp(path.join(os.tmpdir(), "oc-fsutil-"));
});

afterEach(async () => {
  vi.restoreAllMocks();
  await fs.rm(dir, { recursive: true, force: true });
});

/** 目录里剩下的文件名，排序后回给断言用（"有没有留半个文件"就看这个）。 */
async function entries(): Promise<string[]> {
  return (await fs.readdir(dir)).sort();
}

describe("原子写回：写 .tmp 再 rename", () => {
  it("写成功时目标拿到完整字节，且不留下 .tmp", async () => {
    const target = path.join(dir, "out.xlsx");
    await writeFileAtomically(target, new TextEncoder().encode("完整一份"));
    expect(await fs.readFile(target, "utf8")).toBe("完整一份");
    expect(await entries()).toEqual(["out.xlsx"]);
  });

  it("覆盖已有文件：rename 那一刻目标仍是整份旧内容，换上去的也是整份新内容", async () => {
    const target = path.join(dir, "book.docx");
    await fs.writeFile(target, "旧的一份");
    const realRename = fs.rename;
    let contentAtRename: string | null = null;
    vi.spyOn(fs, "rename").mockImplementation(async (from, to) => {
      contentAtRename = await fs.readFile(String(to), "utf8");
      return realRename(from, to);
    });
    await writeFileAtomically(target, new TextEncoder().encode("新的一份"));
    // 可见状态只有"旧"和"新"两种：被截断的半个文件永远不会出现在目标路径上。
    expect(contentAtRename).toBe("旧的一份");
    expect(await fs.readFile(target, "utf8")).toBe("新的一份");
    expect(await entries()).toEqual(["book.docx"]);
  });

  it("rename 被打断（目标位置是个目录）：抛原来的错、.tmp 不留、目标里的文件没被动过", async () => {
    const target = path.join(dir, "locked");
    await fs.mkdir(target);
    const inside = path.join(target, "keep.txt");
    await fs.writeFile(inside, "别碰我");
    await expect(writeFileAtomically(target, new TextEncoder().encode("半个文件"))).rejects.toThrow();
    expect(await fs.readdir(dir)).toEqual(["locked"]);
    expect(await fs.readFile(inside, "utf8")).toBe("别碰我");
  });

  it("目标目录不存在时在写 .tmp 那一步就失败，磁盘上不留任何残留", async () => {
    const target = path.join(dir, "missing", "out.xlsx");
    await expect(writeFileAtomically(target, new TextEncoder().encode("x"))).rejects.toThrow();
    expect(await entries()).toEqual([]);
  });
});
