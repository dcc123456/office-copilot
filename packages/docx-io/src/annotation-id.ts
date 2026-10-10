/**
 * OOXML 注解 id 的分配。
 *
 * `w:id` 是 ST_DecimalNumber，批注、脚注/尾注、书签、修订都在自己的那一串里取最大值加一
 * （各部件的编号是否必须互不重复**没有官方说明**，所以照 Word 实际写出的文件一样各算各的）。
 * "哪些条目参与计数、一条都没有时从几起"每个部件不一样，取最大值这一半算法是同一份，抽在这里。
 */

/**
 * 现有 id 里的最大值加一。
 *
 * @param ids 参与计数的 `w:id`（字符串原样带入，这里只读数字，不做任何单位或格式换算——AGENTS §2.1）
 * @param firstId 一个可用 id 都没有时的起点：批注与书签从 `0`（规范样例第一条批注是 `w:id="0"`），
 *                脚注/尾注与修订从 `1`（0 被版式件占了，Word 自己也从 1 起）
 * @returns 下一个可用的 id（字符串）；认不出数字的 id 不参与计数
 */
export function nextAnnotationId(ids: readonly string[], firstId = 0): string {
  let highest = firstId - 1;
  for (const id of ids) {
    const value = Number(id);
    if (Number.isFinite(value) && value > highest) highest = value;
  }
  return String(highest + 1);
}
