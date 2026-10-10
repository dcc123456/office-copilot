import { describe, expect, it } from "vitest";
import { EndnoteReference, FootnoteReference } from "./FootnoteRef.js";

/**
 * 引用点节点的 DOM 读写规则。复制粘贴走 PM 的 DOMSerializer / DOMParser，只认
 * `renderHTML` 写出的 data-* 与 `parseHTML` 的取值：写成裸 `id` 属性会撞上
 * ProseMirror 的保留属性，粘出去的两个脚注就成了同一个 id。
 * 这里直接调这两个回调，本仓没有 DOM 测试环境。
 */

type NodeRule = {
  parseHTML: (el: { getAttribute: (name: string) => string | null }) => string;
};

const idRules = (kind: "footnote" | "endnote"): NodeRule => {
  const node = kind === "footnote" ? FootnoteReference : EndnoteReference;
  const attrs = (node.config.addAttributes as unknown as () => { id: NodeRule })();
  return attrs.id;
};

const render = (
  node: typeof FootnoteReference,
  attrs: Record<string, unknown>,
): [string, Record<string, unknown>, string] =>
  (node.config.renderHTML as unknown as (arg: { node: { attrs: Record<string, unknown> } }) => [
    string,
    Record<string, unknown>,
    string,
  ])({ node: { attrs } });

/** 只实现 getAttribute 的元素替身：parseHTML 规则只用得到它。 */
const elementWith = (values: Record<string, string>) => ({
  getAttribute: (name: string) => values[name] ?? null,
});

describe("脚注引用点的 DOM 读写", () => {
  it("id 写成 data-* 属性，读回来还是同一个 id（负数也在）", () => {
    const [, html] = render(FootnoteReference, { id: "-1" });
    expect(html).toMatchObject({ "data-footnote-ref": "footnote", "data-footnote-id": "-1" });
    expect(Object.keys(html)).not.toContain("id");
    expect(idRules("footnote").parseHTML(elementWith({ "data-footnote-id": "-1" }))).toBe("-1");
  });

  it("脚注与尾注是两种节点，芯片与标记属性都各认各的", () => {
    const [, fnHtml, fnText] = render(FootnoteReference, { id: "1" });
    const [, enHtml, enText] = render(EndnoteReference, { id: "1" });
    expect(fnHtml).toMatchObject({ "data-footnote-ref": "footnote" });
    expect(enHtml).toMatchObject({ "data-footnote-ref": "endnote" });
    expect([fnText, enText]).toEqual(["注", "尾"]);
    // 读的时候按标记属性分流，粘回编辑器不会把尾注变成脚注。
    const tags = (node: typeof FootnoteReference) =>
      (node.config.parseHTML as unknown as () => { tag: string }[])().map((r) => r.tag);
    expect(tags(FootnoteReference)).toEqual(['span[data-footnote-ref="footnote"]']);
    expect(tags(EndnoteReference)).toEqual(['span[data-footnote-ref="endnote"]']);
  });
});
