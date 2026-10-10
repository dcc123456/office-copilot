import { describe, expect, it } from "vitest";
import {
  buildContextMenu,
  menuRect,
  submenuRect,
  type CtxFacts,
  type CtxItem,
  type CtxRow,
} from "./contextMenu.js";

/** 一份"选区落在正文文字上、剪贴板有内容"的基线事实，各用例只改自己关心的字段。 */
function facts(patch: Partial<CtxFacts> = {}): CtxFacts {
  return {
    target: "text",
    hasSelection: true,
    isLink: false,
    linkHref: null,
    isListItem: false,
    isBulletList: false,
    isOrderedList: false,
    headingLevel: null,
    imageWrap: null,
    painterActive: false,
    hasRevision: false,
    clipboard: { text: "甲", html: "<p>甲</p>" },
    ...patch,
  };
}

function rows(f: CtxFacts): CtxRow[] {
  const built = buildContextMenu(f);
  if (!built) throw new Error("表格目标不该由这张菜单处理");
  return built;
}

function find(rows: CtxRow[], id: string): CtxItem {
  const hit = rows.find((r): r is CtxItem => "id" in r && r.id === id);
  if (!hit) throw new Error(`菜单里没有 ${id}`);
  return hit;
}

describe("buildContextMenu", () => {
  it("表格内的右击交回 TableContextMenu，本模块不出菜单", () => {
    expect(buildContextMenu(facts({ target: "table" }))).toBeNull();
  });

  it("剪贴板为空时粘贴置灰，有内容时可用", () => {
    expect(find(rows(facts()), "paste").disabled).toBe(false);
    expect(find(rows(facts({ clipboard: { text: "", html: "" } })), "paste").disabled).toBe(
      true,
    );
  });

  it("没有选区时剪切/复制/清除格式/AI 全部置灰，与 Word 一致", () => {
    const bare = rows(facts({ hasSelection: false }));
    expect(find(bare, "cut").disabled).toBe(true);
    expect(find(bare, "copy").disabled).toBe(true);
    expect(find(bare, "clear-format").disabled).toBe(true);
    expect(find(bare, "ai:polish").disabled).toBe(true);
  });

  it("粘贴子菜单给 Word 的三种粘贴方式，剪贴板无纯文本时禁用「只保留文本」", () => {
    const paste = find(rows(facts()), "paste");
    expect(paste.children?.map((c) => c.id)).toEqual([
      "paste-source",
      "paste-merge",
      "paste-text",
    ]);
    const htmlOnly = find(
      rows(facts({ clipboard: { text: "", html: "<p>甲</p>" } })),
      "paste",
    );
    expect(htmlOnly.children?.find((c) => c.id === "paste-text")?.disabled).toBe(true);
  });

  it("缩进命令只在列表项里可用", () => {
    expect(find(rows(facts()), "indent-more").disabled).toBe(true);
    const inList = rows(facts({ isListItem: true, isBulletList: true }));
    expect(find(inList, "indent-more").disabled).toBe(false);
    expect(find(inList, "bullet-list").checked).toBe(true);
  });

  it("段落样式子菜单按当前块勾选，标题 2 时勾在标题 2", () => {
    const styles = find(rows(facts({ headingLevel: 2 })), "block");
    const checked = styles.children?.filter((c) => c.checked).map((c) => c.id);
    expect(checked).toEqual(["block:h2"]);
  });

  it("右击超链接换成打开/编辑/删除三项，普通文字只给插入", () => {
    const link = rows(facts({ isLink: true, linkHref: "https://a.test" }));
    expect(link.some((r) => "id" in r && r.id === "link-open")).toBe(true);
    expect(link.some((r) => "id" in r && r.id === "link-remove")).toBe(true);
    const plain = rows(facts());
    expect(plain.some((r) => "id" in r && r.id === "link-open")).toBe(false);
  });

  it("接受 / 拒绝修订两条按落点置灰：没落在修订里也看得见，只是点不动", () => {
    const plain = rows(facts());
    expect(find(plain, "revision-accept").disabled).toBe(true);
    expect(find(plain, "revision-reject").disabled).toBe(true);
    const onRevision = rows(facts({ hasRevision: true }));
    expect(find(onRevision, "revision-accept").disabled).toBe(false);
    expect(find(onRevision, "revision-reject").disabled).toBe(false);
  });

  it("新建批注按选区置灰：光标没拉出要批的文字就点不动", () => {
    expect(find(rows(facts()), "comment-add").disabled).toBe(false);
    expect(find(rows(facts({ hasSelection: false })), "comment-add").disabled).toBe(true);
  });

  it("图片菜单给剪切/复制/文字环绕/删除，文字环绕勾当前值", () => {
    const image = rows(facts({ target: "image", hasSelection: true, imageWrap: "tight" }));
    expect(image.some((r) => "id" in r && r.id === "font-dialog")).toBe(false);
    expect(image.some((r) => "id" in r && r.id === "delete-object")).toBe(true);
    const wrap = find(image, "wrap");
    expect(wrap.children?.find((c) => c.checked)?.id).toBe("wrap:tight");
  });

  it("字体与段落对话框始终成对出现，位置在缩进之后", () => {
    const labels = rows(facts())
      .filter((r): r is CtxItem => "id" in r)
      .map((r) => r.id);
    expect(labels.indexOf("font-dialog")).toBeGreaterThan(labels.indexOf("indent-less"));
    expect(labels).toContain("paragraph-dialog");
  });
});

describe("menuRect", () => {
  it("默认贴在鼠标点右下方", () => {
    expect(menuRect({ x: 200, y: 300 }, { width: 200, height: 320 }, { width: 1200, height: 800 })).toEqual({
      left: 200,
      top: 300,
    });
  });

  it("贴右下边缘时上翻并左移，整块留在窗口内", () => {
    const r = menuRect(
      { x: 1150, y: 780 },
      { width: 200, height: 320 },
      { width: 1200, height: 800 },
    );
    expect(r.left).toBe(950);
    expect(r.top).toBeLessThanOrEqual(800 - 320 - 4);
  });

  it("窗口很窄时不会推到负坐标", () => {
    const r = menuRect({ x: 0, y: 0 }, { width: 200, height: 320 }, { width: 120, height: 100 });
    expect(r.left).toBeGreaterThanOrEqual(4);
    expect(r.top).toBeGreaterThanOrEqual(4);
  });
});

describe("submenuRect", () => {
  it("父项右侧空间不足时翻到左边", () => {
    const r = submenuRect(
      { left: 1000, top: 100, width: 200 },
      40,
      { width: 180, height: 120 },
      { width: 1200, height: 800 },
    );
    expect(r.left).toBe(820);
  });

  it("正常情况贴在父菜单右缘并对齐被悬停项", () => {
    const r = submenuRect(
      { left: 100, top: 100, width: 200 },
      40,
      { width: 180, height: 120 },
      { width: 1200, height: 800 },
    );
    expect(r).toEqual({ left: 300, top: 140 });
  });
});
