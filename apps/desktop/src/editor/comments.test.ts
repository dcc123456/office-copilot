import { describe, expect, it } from "vitest";
import type { Editor } from "@tiptap/react";
import {
  FALLBACK_COMMENT_AUTHOR,
  emptyPmDoc,
  type CommentBody,
} from "@office-copilot/docx-io";
import { commentStamp, deleteCommentById, insertCommentOnSelection } from "./comments.js";

/**
 * 侧栏那层只测"记账"：id 怎么分、署名与时间怎么写、条目怎么跟着锚点走。
 * 锚点的位计算在真实 schema 那边测（`docBlocks.live.test.ts` 的批注那组），这里不重复。
 */

/** 记下令的假编辑器：选区是否空的，以及两条批注命令各自收到的 id。 */
function fakeEditor(selectionEmpty: boolean): {
  editor: Editor;
  inserted: string[];
  deleted: string[];
} {
  const inserted: string[] = [];
  const deleted: string[] = [];
  const editor = {
    state: { selection: { empty: selectionEmpty } },
    commands: {
      insertCommentRange: (id: string) => {
        inserted.push(id);
        return true;
      },
      deleteComment: (id: string) => {
        deleted.push(id);
        return true;
      },
    },
  } as unknown as Editor;
  return { editor, inserted, deleted };
}

const body = (id: string): CommentBody => ({ id, doc: emptyPmDoc(), author: "前任作者" });

describe("批注条目的新建与删除", () => {
  it("新 id 跟着现有条目的最大值走，署名与到秒的时间一起写进条目", () => {
    const { editor, inserted } = fakeEditor(false);
    const r = insertCommentOnSelection(editor, [body("0"), body("2")], "张三");
    expect(r.ok).toBe(true);
    if (!r.ok) return;
    expect(inserted).toEqual(["3"]);
    expect(r.id).toBe("3");
    expect(r.bodies).toHaveLength(3);
    const added = r.bodies[2]!;
    expect(added).toMatchObject({ id: "3", author: "张三" });
    expect(added.date).toMatch(/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z$/);
    expect(added.doc).toEqual(emptyPmDoc());
    // 已有条目一个字都不动：改的是表，不是别人的批注。
    expect(r.bodies.slice(0, 2)).toEqual([body("0"), body("2")]);
  });

  it("署名留空（或全是空格）回落占位作者名——w:author 是必填属性", () => {
    const { editor } = fakeEditor(false);
    for (const authorName of [undefined, "", "   "]) {
      const r = insertCommentOnSelection(editor, [], authorName);
      expect(r.ok && r.bodies[0]?.author).toBe(FALLBACK_COMMENT_AUTHOR);
    }
  });

  it("光标没拉选区时不动正文，也不产生条目", () => {
    const { editor, inserted } = fakeEditor(true);
    const r = insertCommentOnSelection(editor, [body("0")], "张三");
    expect(r).toEqual({ ok: false, reason: "no-selection" });
    expect(inserted).toEqual([]);
  });

  it("删除把条目剔掉；锚点已经被用户撤掉时条目照样走", () => {
    const { editor, deleted } = fakeEditor(false);
    const next = deleteCommentById(editor, [body("0"), body("1")], "0");
    expect(deleted).toEqual(["0"]);
    expect(next.map((entry) => entry.id)).toEqual(["1"]);
  });
});

describe("批注时间戳", () => {
  it("到秒、带 Z，不带小数秒——Word 自己就是这个写法", () => {
    expect(commentStamp(new Date("2026-01-02T10:30:00.456Z"))).toBe("2026-01-02T10:30:00Z");
  });
});
