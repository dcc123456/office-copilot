import type { Editor } from "@tiptap/react";
import {
  FALLBACK_COMMENT_AUTHOR,
  emptyPmDoc,
  nextCommentId,
  type CommentBody,
} from "@office-copilot/docx-io";

/**
 * 批注侧栏的引擎层（排名 4 ⑤）。
 *
 * 正文只装三个锚点，批注的文字住在 `word/comments.xml` 部件里，所以"新建一条批注"永远是
 * 两件事：往正文插锚点（`extensions/Comment.ts` 的命令）＋ 往条目表里加一份空正文。表由
 * 调用方（App 的 state）拿着，保存时随 payload 进主进程；这一层不写文件，也不自己解析部件。
 */

/** 失败码；中文文案由侧栏与右键菜单组装（AGENTS §5.7）。 */
export type CommentReject = "no-selection";

/** 新建成功时的结果：新 id 与加上这条之后的表。 */
export interface CommentInsert {
  ok: true;
  id: string;
  bodies: CommentBody[];
}

/**
 * `w:date` 的写法：ISO 8601 到秒再补 `Z`。Word 自己就不写小数秒，多出来那段它的界面会忽略。
 *
 * @param now 时间，默认当前时刻。
 * @returns 形如 `2026-01-02T10:30:00Z` 的字符串。
 */
export function commentStamp(now: Date = new Date()): string {
  return `${now.toISOString().slice(0, 19)}Z`;
}

/**
 * 给当前选区新建一条批注。
 *
 * @param editor 主编辑器，选区里必须已经装着要批注的文字。
 * @param bodies 现有条目，下一个 id 从它的最大 `w:id` 加一（与正文锚点同一个编号空间）。
 * @param authorName 设置里的署名；留空回落 `FALLBACK_COMMENT_AUTHOR`。
 * @returns 新 id 与加好之后的条目表；没拉选区（或 schema 不认锚点）时是失败码。
 */
export function insertCommentOnSelection(
  editor: Editor,
  bodies: CommentBody[],
  authorName: string | undefined,
): CommentInsert | { ok: false; reason: CommentReject } {
  if (editor.state.selection.empty) return { ok: false, reason: "no-selection" };
  const id = nextCommentId(bodies);
  if (!editor.commands.insertCommentRange(id)) return { ok: false, reason: "no-selection" };
  return {
    ok: true,
    id,
    bodies: [
      ...bodies,
      {
        id,
        doc: emptyPmDoc(),
        author: authorName?.trim() || FALLBACK_COMMENT_AUTHOR,
        date: commentStamp(),
      },
    ],
  };
}

/**
 * 删掉一条批注：正文里的三个锚点与条目一起走。
 *
 * 锚点已经不在正文里（用户先 Ctrl+Z 撤掉了区间）也要把条目删掉，否则保存时它会带着
 * 没人引用的正文留在部件里。
 *
 * @param editor 主编辑器。
 * @param bodies 当前条目表。
 * @param id 要删的 `w:id`。
 * @returns 去掉这条之后的表。
 */
export function deleteCommentById(
  editor: Editor,
  bodies: CommentBody[],
  id: string,
): CommentBody[] {
  editor.commands.deleteComment(id);
  return bodies.filter((body) => body.id !== id);
}
