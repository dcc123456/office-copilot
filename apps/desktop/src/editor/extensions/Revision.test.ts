import { describe, expect, it } from "vitest";
import { Revision } from "./Revision.js";

/**
 * 修订标记的呈现（④ 的"插入下划线 / 删除删除线"那一半）。
 *
 * 本仓没有 DOM 测试环境，所以直接调 `renderHTML` 回调看它交出来的 DOM 描述；
 * 真正的视觉效果由 `pnpm dev:desktop` 起的 app 实测（AGENTS §7.2）。
 */

/**
 * 渲染一个修订标记的 DOM 描述。
 *
 * @param attrs 标记属性（`kind`/`id`/`author`/`date`）。
 * @returns `[标签, 属性对象, 内容位置]`，即 ProseMirror 的 DOMParser 描述。
 */
function render(attrs: Record<string, unknown>): [string, Record<string, string>, number] {
  const renderHTML = Revision.config.renderHTML as unknown as (
    this: unknown,
    props: { mark: { attrs: Record<string, unknown> }; HTMLAttributes: Record<string, unknown> },
  ) => [string, Record<string, string>, number];
  return renderHTML.call(
    { options: { HTMLAttributes: {} } },
    { mark: { attrs }, HTMLAttributes: {} },
  );
}

describe("修订标记的呈现", () => {
  it("删除画删除线并署上作者，插入画下划线", () => {
    const deleted = render({ kind: "del", id: "2", author: "李四", date: null });
    expect(deleted[0]).toBe("span");
    expect(deleted[1]["data-revision"]).toBe("del");
    expect(deleted[1].style).toContain("line-through");
    expect(deleted[1].title).toBe("删除：李四");

    const inserted = render({ kind: "ins", id: "1", author: "张三", date: "2026-01-02T03:04:05Z" });
    expect(inserted[1].style).toContain("underline");
    expect(inserted[1].title).toBe("插入：张三（2026-01-02T03:04:05Z）");
  });

  it("作者缺失时不编一个名字出来", () => {
    expect(render({ kind: "ins", id: "1", author: null, date: null })[1].title).toBe("插入");
  });

  it("不从 HTML 认修订：粘进来的内容不该带着别人的作者与编号", () => {
    const parseHTML = Revision.config.parseHTML as unknown as (this: unknown) => unknown[];
    expect(parseHTML.call({})).toEqual([]);
  });
});
