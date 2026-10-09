import { app, type IpcMain } from "electron";
import path from "node:path";
import fs from "node:fs/promises";
import {
  DEFAULT_SETTINGS,
  type AppSettings,
  type ProviderConfig,
} from "./settings.js";

export type { AppSettings, ProviderConfig };

let cached: AppSettings | null = null;
let saving: Promise<void> = Promise.resolve();

function configPath(): string {
  return path.join(app.getPath("userData"), "config.json");
}

export async function loadSettings(): Promise<AppSettings> {
  if (cached) return cached;
  try {
    const raw = await fs.readFile(configPath(), "utf8");
    const parsed = JSON.parse(raw) as Partial<AppSettings>;
    cached = { ...DEFAULT_SETTINGS, ...parsed };
  } catch {
    cached = { ...DEFAULT_SETTINGS };
  }
  return cached;
}

async function saveSettings(next: AppSettings): Promise<void> {
  cached = next;
  saving = saving.then(async () => {
    const tmp = configPath() + ".tmp";
    await fs.writeFile(tmp, JSON.stringify(next, null, 2), "utf8");
    await fs.rename(tmp, configPath());
  });
  await saving;
}

export function getActiveProviderSync(): ProviderConfig | null {
  if (!cached) return null;
  const id = cached.defaultProviderId;
  if (!id) return cached.providers.find((p) => p.enabled) ?? null;
  return cached.providers.find((p) => p.id === id && p.enabled) ?? null;
}

export function registerSettingsHandlers(ipc: IpcMain): void {
  ipc.handle("settings:get", async () => loadSettings());

  ipc.handle("settings:save", async (_e, next: AppSettings) => {
    await saveSettings(next);
    return { ok: true as const };
  });

  ipc.handle("settings:test-provider", async (_e, provider: ProviderConfig) => {
    try {
      if (!provider.baseUrl) return { ok: false, message: "Base URL 不能为空。" };
      if (!provider.model) return { ok: false, message: "模型名不能为空。" };
      const ctrl = new AbortController();
      const t = setTimeout(() => ctrl.abort(), 15000);
      const res = await fetch(
        `${provider.baseUrl.replace(/\/$/, "")}/chat/completions`,
        {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
            Authorization: `Bearer ${provider.apiKey ?? ""}`,
          },
          body: JSON.stringify({
            model: provider.model,
            messages: [
              { role: "system", content: "ping" },
              { role: "user", content: "回复一个字：pong" },
            ],
            temperature: 0,
            max_tokens: 8,
            ...(provider.extraBody ?? {}),
          }),
          signal: ctrl.signal,
        },
      );
      clearTimeout(t);
      if (!res.ok) {
        const text = await res.text().catch(() => "");
        return { ok: false, message: `HTTP ${res.status}: ${text.slice(0, 300)}` };
      }
      return { ok: true };
    } catch (err) {
      return {
        ok: false,
        message: err instanceof Error ? err.message : String(err),
      };
    }
  });
}
