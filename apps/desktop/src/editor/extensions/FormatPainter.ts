import { Extension } from "@tiptap/core";
import type { Editor } from "@tiptap/core";

/**
 * Format painter: click to "pick up" the marks at the current selection,
 * then select any other range to apply them. Clicking the button a second
 * time turns it off. We support all inline marks that StarterKit + our
 * own extensions expose: bold/italic/underline/strike/code/link, plus
 * textStyle with color/highlight and fontFamily/fontSize.
 *
 * The extension itself just stores a snapshot; the actual application
 * happens on selection updates while the mode is active. We keep it as
 * a TipTap extension so it lives with the other editor state.
 */

interface StoredStyle {
  markTypes: string[];
  markAttrs: Record<string, Record<string, unknown>>;
  paragraphAlign?: string;
  paragraphIndent?: number;
  paragraphIndentFirst?: number;
  paragraphLineHeight?: string;
}

declare module "@tiptap/core" {
  interface Commands<ReturnType> {
    formatPainter: {
      /** Snapshot current selection's marks for later painting. */
      pickFormat: () => ReturnType;
      /** Apply the snapshot to current selection and clear it. */
      paintFormat: () => ReturnType;
      /** Cancel without painting. */
      cancelFormatPaint: () => ReturnType;
    };
  }
}

interface State {
  stored: StoredStyle | null;
}

export const FormatPainter = Extension.create<Record<string, never>, State>({
  name: "formatPainter",

  addStorage() {
    return { stored: null };
  },

  addCommands() {
    const self = this;
    return {
      pickFormat:
        () =>
        ({ editor }: { editor: Editor }) => {
          const { from, to, empty } = editor.state.selection;
          if (empty) {
            self.storage.stored = snapshotAt(editor, from);
          } else {
            self.storage.stored = snapshotRange(editor, from, to);
          }
          return true;
        },
      paintFormat:
        () =>
        ({ editor }: { editor: Editor }) => {
          const s = self.storage.stored;
          if (!s) return false;
          applySnapshot(editor, s);
          self.storage.stored = null;
          return true;
        },
      cancelFormatPaint:
        () =>
        () => {
          self.storage.stored = null;
          return true;
        },
    } as never;
  },

  onCreate() {
    // When the user changes the selection after "picking" a style, and
    // the new selection is non-empty, auto-paint it (Word behavior).
    const self = this;
    this.editor.on("selectionUpdate", ({ editor }: { editor: Editor }) => {
      if (!self.storage.stored) return;
      const { from, to, empty } = editor.state.selection;
      if (empty || from === to) return;
      const s = self.storage.stored;
      self.storage.stored = null;
      if (s) applySnapshot(editor, s);
    });
  },
});

function snapshotAt(editor: Editor, pos: number): StoredStyle {
  const $pos = editor.state.doc.resolve(pos);
  const parent = $pos.parent;
  const stored: StoredStyle = { markTypes: [], markAttrs: {} };
  // Inspect marks at the position just before (preferred) or after.
  const before = pos > 0 ? editor.state.doc.resolve(pos - 1).marks() : [];
  for (const m of before) {
    stored.markTypes.push(m.type.name);
    stored.markAttrs[m.type.name] = { ...m.attrs };
  }
  if (parent?.type.name === "paragraph" || parent?.type.name === "heading") {
    stored.paragraphAlign = parent.attrs.textAlign;
    stored.paragraphIndent = parent.attrs.indent;
    stored.paragraphIndentFirst = parent.attrs.indentFirst;
    stored.paragraphLineHeight = parent.attrs.lineHeight;
  }
  return stored;
}

function snapshotRange(editor: Editor, from: number, to: number): StoredStyle {
  // Use the first text mark in range; Word does the same — style is
  // copied from the first character.
  let pos = from;
  editor.state.doc.nodesBetween(from, to, (node: any, p: number) => {
    if (node.isText && (node.marks?.length ?? 0) > 0 && pos === from) {
      pos = p;
    }
    return !node.isText;
  });
  return snapshotAt(editor, pos);
}

function applySnapshot(editor: Editor, s: StoredStyle): void {
  const { from, to } = editor.state.selection;
  if (from === to) return;
  const chain = editor.chain().focus();
  // Reset marks first so we don't carry over stale bold/italic.
  chain.unsetAllMarks();
  for (const name of s.markTypes) {
    const attrs = s.markAttrs[name] ?? {};
    switch (name) {
      case "bold":
        chain.setBold();
        break;
      case "italic":
        chain.setItalic();
        break;
      case "underline":
        chain.setUnderline();
        break;
      case "strike":
        chain.setStrike();
        break;
      case "code":
        chain.setCode();
        break;
      case "textStyle":
        if (attrs.color) {
          // TextStyleCommands exposes setColor; cast through unknown for
          // type-safety (it exists at runtime because TextStyle + the
          // color command are both registered).
          (chain as unknown as { setColor: (c: string) => unknown }).setColor(
            attrs.color as string,
          );
        }
        if (attrs.fontFamily) {
          (
            chain as unknown as { setFontFamily: (f: string) => unknown }
          ).setFontFamily?.(attrs.fontFamily as string);
        }
        if (attrs.fontSize) {
          (
            chain as unknown as { setFontSize: (s: string) => unknown }
          ).setFontSize?.(attrs.fontSize as string);
        }
        break;
      case "highlight":
        (
          chain as unknown as {
            toggleHighlight: (a: Record<string, unknown>) => unknown;
          }
        ).toggleHighlight?.(attrs);
        break;
      case "link":
        if (attrs.href)
          chain.extendMarkRange("link").setLink({ href: attrs.href as string });
        break;
    }
  }
  // Paragraph-level attrs.
  const patch: Record<string, unknown> = {};
  if (s.paragraphAlign) patch.textAlign = s.paragraphAlign;
  if (s.paragraphIndent !== undefined) patch.indent = s.paragraphIndent;
  if (s.paragraphIndentFirst !== undefined)
    patch.indentFirst = s.paragraphIndentFirst;
  if (s.paragraphLineHeight) patch.lineHeight = s.paragraphLineHeight;
  if (Object.keys(patch).length) {
    chain.setParagraphFormat(patch);
  }
  chain.run();
}
