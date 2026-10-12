import { afterEach, describe, expect, it, vi } from "vitest";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";

/**
 * 会话存储的用例：真实临时目录（AGENTS §7.1 不 mock 文件系统），
 * 只把 electron 的 `app.getPath` 换成临时目录——主进程那套原子写与剪枝要按文件核对。
 */

const dir = await fs.mkdtemp(path.join(os.tmpdir(), "oc-sessions-"));
vi.mock("electron", () => ({
  app: { getPath: () => dir },
}));

const { normaliseRecord, registerSessionStoreHandlers } = await import("./sessionStore.js");

type Handler = (e: unknown, arg?: unknown) => Promise<unknown>;
const handlers: Record<string, Handler> = {};
registerSessionStoreHandlers({
  handle: (channel: string, fn: Handler) => {
    handlers[channel] = fn;
  },
} as never);

afterEach(async () => {
  await fs.rm(path.join(dir, "sessions"), { recursive: true, force: true });
});

describe("normaliseRecord", () => {
  it("只留合法角色与非空文字，id 以文件名为准", () => {
    const rec = normaliseRecord(
      {
        id: "别人塞进来的",
        title: "  改简历  ",
        createdAt: 5,
        updatedAt: 9,
        turns: [
          { role: "user", text: "帮我看看" },
          { role: "system", text: "不该有的角色" },
          { role: "assistant", text: "   " },
          "不是对象",
          { role: "assistant", text: "第一段建议" },
        ],
      },
      "s-1",
    );
    expect(rec).not.toBeNull();
    expect(rec!.id).toBe("s-1");
    expect(rec!.title).toBe("改简历");
    expect(rec!.turns).toEqual([
      { role: "user", text: "帮我看看" },
      { role: "assistant", text: "第一段建议" },
    ]);
  });

  it("文件完全不像会话时返回 null，列表跳过它而不是整批失败", () => {
    expect(normaliseRecord(null, "s-1")).toBeNull();
    expect(normaliseRecord({ turns: "不是数组" }, "s-1")).toBeNull();
  });
});

describe("session:* handlers", () => {
  const id = "11111111-2222-4333-8444-555555555555";

  it("保存→列表→读回→删除，走的是 userData/sessions 下的真实文件", async () => {
    const saved = await handlers["session:save"]!(null, {
      id,
      title: "改简历",
      turns: [
        { role: "user", text: "帮我改简历" },
        { role: "assistant", text: "先把项目经历写短" },
      ],
    });
    expect(saved).toEqual({ ok: true });
    await expect(fs.readFile(path.join(dir, "sessions", `${id}.json`), "utf8")).resolves.toContain(
      "先把项目经历写短",
    );

    const list = (await handlers["session:list"]!(null)) as Array<{
      id: string;
      turnCount: number;
    }>;
    expect(list).toEqual([{ id, title: "改简历", updatedAt: expect.any(Number), turnCount: 2 }]);

    const read = (await handlers["session:read"]!(null, id)) as { turns: unknown[] };
    expect(read.turns).toHaveLength(2);

    expect(await handlers["session:remove"]!(null, id)).toEqual({ ok: true });
    await expect(fs.readFile(path.join(dir, "sessions", `${id}.json`), "utf8")).rejects.toThrow();
  });

  it("空轮次不保存，越界 id 一律拒绝（防路径穿越）", async () => {
    expect(await handlers["session:save"]!(null, { id, turns: [] })).toEqual({
      ok: false,
      message: expect.stringContaining("空的"),
    });
    const hostile = "../../etc/passwd";
    expect(await handlers["session:save"]!(null, { id: hostile, turns: [{ role: "user", text: "x" }] })).toEqual({
      ok: false,
      message: expect.stringContaining("不合法"),
    });
    expect(await handlers["session:read"]!(null, hostile)).toBeNull();
    expect(await handlers["session:remove"]!(null, hostile)).toEqual({ ok: false });
  });

  it("同一个 id 再存一次是覆盖，不是并排出两份", async () => {
    await handlers["session:save"]!(null, { id, turns: [{ role: "user", text: "第一版" }] });
    await handlers["session:save"]!(null, {
      id,
      turns: [
        { role: "user", text: "第一版" },
        { role: "assistant", text: "第二版补了一句" },
      ],
    });
    const read = (await handlers["session:read"]!(null, id)) as { turns: unknown[] };
    expect(read.turns).toHaveLength(2);
    const list = (await handlers["session:list"]!(null)) as unknown[];
    expect(list).toHaveLength(1);
  });

  it("会话多过 50 个就剪掉最旧的，目录不会无限长", async () => {
    for (let i = 0; i < 51; i++) {
      await handlers["session:save"]!(null, {
        id: `sess-${i}`,
        turns: [{ role: "user", text: `第 ${i} 个会话` }],
      });
    }
    const list = (await handlers["session:list"]!(null)) as Array<{ id: string }>;
    expect(list).toHaveLength(50);
    const onDisk = await fs.readdir(path.join(dir, "sessions"));
    expect(onDisk.filter((n) => n.endsWith(".json"))).toHaveLength(50);
  });
});
