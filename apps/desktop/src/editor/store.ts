import { create } from "zustand";
import type { Editor } from "@tiptap/react";

interface EditorState {
  editor: Editor | null;
  setModel: (editor: Editor | null) => void;
  applyModel: () => Editor | null;
}

export const useEditorStore = create<EditorState>((set, get) => ({
  editor: null,
  setModel: (editor) => set({ editor }),
  applyModel: () => get().editor,
}));
