import { Node } from "@tiptap/core";

/**
 * 跨段落复合域（Word 生成的目录）的两个边界节点，对应
 * `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` 与 `…="end"`。
 *
 * 单个 `field` 节点只能表示同段内配好对的域；目录的 begin 在首段、end 在末段，
 * 中间夹着条目段落，只能用"起点 + 终点"两个原子节点把区间标出来，中间的内容
 * 才能继续按结构存（条目里的 PAGEREF 仍是芯片、制表符仍是 run）。
 * 与书签同理，这里不做配对校验：半截域由写侧原样保留，Word 打开后会自行重算。
 *
 * @param name 节点名，与 `docx-io` 的 PM 节点名一致。
 * @param isStart 起点带指令（悬停能看到 `TOC \o "1-3"`），终点只是结束记号。
 * @returns 注册进 `EDITOR_EXTENSIONS` 的 Tiptap 节点。
 */
function fieldBoundaryNode(name: "fieldBegin" | "fieldEnd", isStart: boolean) {
  return Node.create({
    name,
    group: "inline",
    inline: true,
    atom: true,

    addAttributes() {
      return {
        /** 由指令首词推出来的类别（`TOC` / `SEQ` …），只用于样式与提示。 */
        kind: {
          default: "TOC",
          parseHTML: (el) => el.getAttribute("data-field-kind") ?? undefined,
        },
        /** OOXML 的 `w:instrText` 内容，原样存字符串（换算与补空格由 docx-io 负责）。 */
        instr: {
          default: "",
          parseHTML: (el) => el.getAttribute("data-instr") ?? undefined,
        },
        /**
         * begin 后面有没有 `separate` 记号：false 表示 Word 标了 dirty 还没算过结果。
         * 终点没有这个属性，缺省 `undefined` 而不是 false，免得写回时凭空造记号。
         */
        isSeparated: {
          default: undefined,
          parseHTML: (el) =>
            el.hasAttribute("data-separated") ? true : el.hasAttribute("data-dirty") ? false : undefined,
        },
      };
    },

    parseHTML() {
      return [{ tag: `span[data-field-boundary="${isStart ? "begin" : "end"}"]` }];
    },

    // 只写 data-*：`instr` 是属性名里最容易被当成保留字的一个，裸写进 DOM 会漏进
    // 复制粘贴的 HTML 里；指令放 title 供悬停查看。
    renderHTML({ node }) {
      const instr = String(node.attrs.instr ?? "");
      const separated = node.attrs.isSeparated !== false;
      return [
        "span",
        {
          "data-field-boundary": isStart ? "begin" : "end",
          ...(isStart
            ? {
                "data-field-kind": String(node.attrs.kind ?? "TOC"),
                ...(instr ? { "data-instr": instr } : {}),
                [separated ? "data-separated" : "data-dirty"]: "",
              }
            : {}),
          ...(instr ? { title: instr } : {}),
          class: `pm-field-boundary pm-field-boundary-${isStart ? "begin" : "end"}`,
        },
      ];
    },
  });
}

export const FieldBegin = fieldBoundaryNode("fieldBegin", true);

export const FieldEnd = fieldBoundaryNode("fieldEnd", false);
