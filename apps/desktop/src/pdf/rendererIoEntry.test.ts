/**
 * 渲染层引用 PDF 引擎的入口那条红线（`AGENTS.md` §5.2）：只能走 `@office-copilot/pdf-io/pure`。
 *
 * 这条不是洁癖：包总入口会把 `pdf-file.ts` 一起拖进浏览器包，那里有 `node:zlib` 与 `@cantoo/pdf-lib`，
 * 实测结果是 `require is not defined`、整页白屏（PDF 视图连挂都挂不上）。Vite 构建不报，只有真跑 app 才看得见，
 * 所以按源码钉一条：`src/**` 里出现主入口的 import 就是回归。
 */
import { readdirSync, readFileSync, statSync } from "node:fs";
import path from "node:path";
import { describe, expect, it } from "vitest";

/** 递归列出目录下的源码文件（只看到会 import 的那几种扩展名）。 */
function sourceFiles(dir: string): string[] {
  const found: string[] = [];
  for (const name of readdirSync(dir)) {
    const full = path.join(dir, name);
    if (statSync(full).isDirectory()) found.push(...sourceFiles(full));
    else if (/\.(ts|tsx)$/.test(name)) found.push(full);
  }
  return found;
}

describe("渲染层不许 import 带 Node 能力的 pdf-io 总入口", () => {
  it("src 下每一处 pdf-io 的 import 都指向 /pure 子路径", () => {
    const offenders = sourceFiles(path.resolve(process.cwd(), "src"))
      .filter((file) => /from "@office-copilot\/pdf-io"/.test(readFileSync(file, "utf8")))
      .map((file) => path.relative(process.cwd(), file));
    expect(offenders).toEqual([]);
  });
});
