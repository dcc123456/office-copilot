import { describe, expect, it } from "vitest";
import { nextAnnotationId } from "./annotation-id.js";

/**
 * `w:id` 分配的公共那一半：批注、脚注、书签、修订各自筛好参与计数的 id 之后都走这里。
 * 各部件自己的筛选规则由 `nextCommentId` / `nextFootnoteId` / `nextBookmarkId` 的用例覆盖。
 */

describe("注解 id 分配", () => {
  it("取最大数字 id 加一，认不出数字的不参与", () => {
    expect(nextAnnotationId(["0", "12", "3"])).toBe("13");
    expect(nextAnnotationId(["", "abc", "7"])).toBe("8");
  });

  it("一条都没有时从各自的起点起：批注与书签 0 起、脚注与修订 1 起", () => {
    expect(nextAnnotationId([])).toBe("0");
    expect(nextAnnotationId([], 1)).toBe("1");
  });
});
