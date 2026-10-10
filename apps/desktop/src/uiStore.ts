import { create } from "zustand";
import { persist } from "zustand/middleware";

export type ThemeId =
  | "light"
  | "dark"
  | "spring"
  | "summer"
  | "autumn"
  | "winter";

export const THEMES: { id: ThemeId; label: string; swatch: string }[] = [
  { id: "light", label: "素白", swatch: "#f7f5f2" },
  { id: "dark", label: "夜墨", swatch: "#232221" },
  { id: "spring", label: "春·青野", swatch: "#8fbf9f" },
  { id: "summer", label: "夏·陶阳", swatch: "#e08b5a" },
  { id: "autumn", label: "秋·琥珀", swatch: "#c08a5a" },
  { id: "winter", label: "冬·霜蓝", swatch: "#9db4cf" },
];

const KEY = "oc-ui";

export interface UiPrefs {
  theme: ThemeId;
  outlineWidth: number;
  aiWidth: number;
  commandUsage: Record<string, number>;
  zoom: number;
  lastTemplateId?: string;
}

interface UiState extends UiPrefs {
  setTheme: (t: ThemeId) => void;
  setOutlineWidth: (w: number) => void;
  setAiWidth: (w: number) => void;
  setZoom: (z: number) => void;
  bumpCommand: (id: string) => void;
  setLastTemplate: (id?: string) => void;
}

export function applyTheme(t: ThemeId) {
  document.documentElement.dataset.theme = t;
}

/** Synchronous read used by main.tsx before React mounts (no theme flash). */
export function readStoredTheme(): ThemeId {
  try {
    const raw = localStorage.getItem(KEY);
    if (raw) {
      const t = JSON.parse(raw)?.state?.theme as ThemeId | undefined;
      if (t && THEMES.some((x) => x.id === t)) return t;
    }
  } catch {
    /* corrupted storage falls back to light */
  }
  return "light";
}

export const useUiStore = create<UiState>()(
  persist(
    (set) => ({
      theme: "light",
      outlineWidth: 220,
      aiWidth: 340,
      commandUsage: {},
      zoom: 1,
      setTheme: (theme) => {
        applyTheme(theme);
        set({ theme });
      },
      setOutlineWidth: (outlineWidth) => set({ outlineWidth }),
      setAiWidth: (aiWidth) => set({ aiWidth }),
      setZoom: (zoom) => set({ zoom: Math.min(2, Math.max(0.5, zoom)) }),
      bumpCommand: (id) =>
        set((s) => ({
          commandUsage: { ...s.commandUsage, [id]: (s.commandUsage[id] ?? 0) + 1 },
        })),
      setLastTemplate: (lastTemplateId) => set({ lastTemplateId }),
    }),
    { name: KEY },
  ),
);
