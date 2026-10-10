import type { Editor } from "@tiptap/react";
import { promptDialog } from "./ui/Dialog.js";

/**
 * 询问链接地址并写入选区；与工具栏的「插入链接」按钮同一份实现。
 *
 * @param editor 当前编辑器；有选区时把链接套在选区上，没有选区则不产生改动。
 * @returns 始终 resolve（用户取消就什么都不做），因为这是纯交互流程。
 */
export async function promptSetLink(editor: Editor): Promise<void> {
  const previousUrl = editor.getAttributes("link").href as string | undefined;
  const url = await promptDialog({
    title: "插入链接",
    placeholder: "https://…",
    initialValue: previousUrl ?? "https://",
  });
  if (url === null) return;
  if (url === "") {
    editor.chain().focus().extendMarkRange("link").unsetLink().run();
    return;
  }
  editor.chain().focus().extendMarkRange("link").setLink({ href: url }).run();
}

/**
 * 只有能被浏览器/系统安全打开的协议才允许交给 shell 打开，
 * 避免文档里藏的 javascript: / file: 之类的链接被右键菜单触发。
 *
 * @param href 链接地址，可能来自 .docx 内容，视为不可信输入。
 * @returns 归一化后的可打开地址；协议不允许时返回 null。
 */
export function safeHref(href: string): string | null {
  try {
    const url = new URL(href);
    return url.protocol === "http:" || url.protocol === "https:" || url.protocol === "mailto:"
      ? url.href
      : null;
  } catch {
    return null;
  }
}
