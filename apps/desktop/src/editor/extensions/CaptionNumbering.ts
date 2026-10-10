import { Extension } from "@tiptap/core";
import { Plugin, PluginKey } from "@tiptap/pm/state";
import { renumberSequences } from "../captions.js";

/**
 * 题注编号跟着正文走：每次文档改动之后，把所有 `SEQ` 域的缓存结果重排成当前顺序下的号数。
 *
 * Word 自己是"更新域"（F9、打印、导出）时才重算，编辑器里等一次 F9 太反直觉——删掉一张图，
 * 后面的题注应该当场就变成 3、4，而不是带着一个旧号写进文件。所以按 ProseMirror 的事务重排，
 * 形态与 `FootnoteRef.ts` 的编号插件一致：派生值不进撤销栈，且只在真有人改号时派发。
 */
export const CaptionNumbering = Extension.create({
  name: "captionNumbering",

  addProseMirrorPlugins() {
    return [
      new Plugin({
        key: new PluginKey("captionNumbering"),
        appendTransaction: (transactions, _oldState, state) => {
          // 没有正文改动就不重排：选区移动、焦点变化都走不到这条路上。
          if (!transactions.some((tr) => tr.docChanged)) return null;
          return renumberSequences(state);
        },
      }),
    ];
  },
});
