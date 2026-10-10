import { Extension, Node } from "@tiptap/core";
import { Plugin, PluginKey } from "@tiptap/pm/state";
import type { Node as PMNode } from "@tiptap/pm/model";

/**
 * 脚注/尾注的引用点节点，对应 OOXML 的 `<w:footnoteReference w:id/>` 与
 * `<w:endnoteReference w:id/>`（ECMA-376 §17.11.18）。
 *
 * 行内 atom、不带内容：正文里只有这一个记号，脚注的**文字**住在 `word/footnotes.xml`
 * 部件里，靠同一个 id 对应（`docx-io` 的 `FootnoteStore`）。编辑器不认这个节点就得
 * 丢掉引用点，一保存整篇脚注都没了。
 *
 * 芯片上的数字由 {@link FootnoteNumbering} 现算后写回 `number` 属性：`renderHTML`
 * 只看得到单个节点，拿不到"这是本部件里第几条"，所以位置信息必须由插件按正文顺序补。
 */

/** 一条引用点在版面上的信息。 */
export interface FootnoteRefInfo {
  /** Word 显示的数字，按正文出现顺序从 1 起；重复引用同一个 id 得到同一个号。 */
  number: number;
  /** 这个 id 第一次出现的位置，窗格里点条目要跳到这里。 */
  pos: number;
}

/** `kind:id` → 版面信息；正文没有引用点时是空表。 */
export type FootnoteRefIndex = Record<string, FootnoteRefInfo>;

export const FOOTNOTE_REF_KEY = new PluginKey<FootnoteRefIndex>("footnoteRefs");

/**
 * 拼索引键。
 *
 * @param kind 脚注还是尾注——两种部件的编号各数各的。
 * @param id OOXML 的 `w:id` 字符串。
 * @returns 索引里的键。
 */
export const footnoteRefKey = (kind: "footnote" | "endnote", id: string): string =>
  `${kind}:${id}`;

/** 节点名 → 部件种类；不是引用点就返回 null。 */
const kindOf = (node: PMNode): "footnote" | "endnote" | null => {
  if (node.type.name === "footnoteReference") return "footnote";
  if (node.type.name === "endnoteReference") return "endnote";
  return null;
};

/**
 * 按正文顺序给所有引用点编号，并记下每个 id 第一次出现的位置。
 *
 * @param doc 编辑器当前正文（ProseMirror 节点，不是 JSON）。
 * @returns 索引表；id 是空串的引用点（读不认识的写法）不进表，芯片回落到"注/尾"字样。
 */
export function indexFootnoteRefs(doc: PMNode): FootnoteRefIndex {
  const out: FootnoteRefIndex = {};
  const counted: Record<"footnote" | "endnote", number> = { footnote: 0, endnote: 0 };
  doc.descendants((node, pos) => {
    const kind = kindOf(node);
    if (!kind) return;
    const id = String(node.attrs.id ?? "");
    if (!id) return;
    const key = footnoteRefKey(kind, id);
    if (out[key]) return;
    counted[kind] += 1;
    out[key] = { number: counted[kind], pos };
  });
  return out;
}

/**
 * 把索引里的编号写回芯片的 `number` 属性。
 *
 * 只在正文真变过的时候重算，且只改号数变了的节点——每次按键都遍历一遍整篇会把
 * 输入延迟放大，而无谓的 transaction 会让脏标记和用户的手工编辑混在一起。
 */
const FootnoteNumbering = Extension.create({
  name: "footnoteNumbering",

  addProseMirrorPlugins() {
    return [
      new Plugin<FootnoteRefIndex>({
        key: FOOTNOTE_REF_KEY,
        state: {
          init: (_config, state) => indexFootnoteRefs(state.doc),
          apply: (tr, value) => (tr.docChanged ? indexFootnoteRefs(tr.doc) : value),
        },
        appendTransaction: (transactions, _oldState, state) => {
          if (!transactions.some((tr) => tr.docChanged)) return null;
          const index = FOOTNOTE_REF_KEY.getState(state) ?? {};
          const tr = state.tr;
          state.doc.descendants((node, pos) => {
            const kind = kindOf(node);
            if (!kind) return;
            const id = String(node.attrs.id ?? "");
            const want = index[footnoteRefKey(kind, id)]?.number;
            if (want === undefined || want === node.attrs.number) return;
            tr.setNodeMarkup(pos, undefined, { ...node.attrs, number: want });
          });
          if (!tr.docChanged) return null;
          // 编号是派生显示值，不进撤销栈：Ctrl+Z 要撤的是用户的编辑，不是页码重排。
          return tr.setMeta("addToHistory", false);
        },
      }),
    ];
  },
});

function footnoteRefNode(kind: "footnote" | "endnote") {
  return Node.create({
    name: kind === "footnote" ? "footnoteReference" : "endnoteReference",
    group: "inline",
    inline: true,
    atom: true,

    addAttributes() {
      return {
        /** OOXML 的 `w:id`：ST_DecimalNumber 原样存字符串，可以是负数。 */
        id: {
          default: "",
          parseHTML: (el) => el.getAttribute("data-footnote-id") ?? "",
        },
        /**
         * 版面上的数字，由编号插件写入、从不从 DOM 读回（它是派生值，
         * 粘贴进来的那份数字属于另一篇文档）。没有它就是新建后还没算过号。
         */
        number: { default: undefined },
      };
    },

    parseHTML() {
      return [{ tag: `span[data-footnote-ref="${kind}"]` }];
    },

    // 只写 data-* 属性：把 HTMLAttributes 原样并进来会连 `id` 一起写到 DOM 上，
    // 那是 ProseMirror 里最不该碰的保留属性（会造出重复的 DOM id）。
    renderHTML({ node }) {
      const id = String(node.attrs.id ?? "");
      const number = node.attrs.number as number | undefined;
      return [
        "span",
        {
          "data-footnote-ref": kind,
          "data-footnote-id": id,
          title: `${kind === "footnote" ? "脚注" : "尾注"} ${number ?? id}`,
          class: `pm-footnote-ref pm-footnote-${kind}`,
          style: "vertical-align:super;font-size:0.75em;padding:0 2px;color:#0078d4;",
        },
        // 还没算出号时退回部件字样，芯片至少说得清自己是脚注还是尾注。
        String(number ?? (kind === "footnote" ? "注" : "尾")),
      ];
    },
  });
}

export const FootnoteReference = footnoteRefNode("footnote");

export const EndnoteReference = footnoteRefNode("endnote");

/** 编辑器要注册的完整一组：两种引用点 + 编号插件。 */
export const FootnoteRefs = [FootnoteReference, EndnoteReference, FootnoteNumbering];
