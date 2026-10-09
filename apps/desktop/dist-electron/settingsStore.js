"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.loadSettings = loadSettings;
exports.getActiveProviderSync = getActiveProviderSync;
exports.registerSettingsHandlers = registerSettingsHandlers;
const electron_1 = require("electron");
const node_path_1 = __importDefault(require("node:path"));
const promises_1 = __importDefault(require("node:fs/promises"));
const settings_js_1 = require("./settings.js");
let cached = null;
let saving = Promise.resolve();
function configPath() {
    return node_path_1.default.join(electron_1.app.getPath("userData"), "config.json");
}
async function loadSettings() {
    if (cached)
        return cached;
    try {
        const raw = await promises_1.default.readFile(configPath(), "utf8");
        const parsed = JSON.parse(raw);
        cached = { ...settings_js_1.DEFAULT_SETTINGS, ...parsed };
    }
    catch {
        cached = { ...settings_js_1.DEFAULT_SETTINGS };
    }
    return cached;
}
async function saveSettings(next) {
    cached = next;
    saving = saving.then(async () => {
        const tmp = configPath() + ".tmp";
        await promises_1.default.writeFile(tmp, JSON.stringify(next, null, 2), "utf8");
        await promises_1.default.rename(tmp, configPath());
    });
    await saving;
}
function getActiveProviderSync() {
    if (!cached)
        return null;
    const id = cached.defaultProviderId;
    if (!id)
        return cached.providers.find((p) => p.enabled) ?? null;
    return cached.providers.find((p) => p.id === id && p.enabled) ?? null;
}
function registerSettingsHandlers(ipc) {
    ipc.handle("settings:get", async () => loadSettings());
    ipc.handle("settings:save", async (_e, next) => {
        await saveSettings(next);
        return { ok: true };
    });
    ipc.handle("settings:test-provider", async (_e, provider) => {
        try {
            if (!provider.baseUrl)
                return { ok: false, message: "Base URL 不能为空。" };
            if (!provider.model)
                return { ok: false, message: "模型名不能为空。" };
            const ctrl = new AbortController();
            const t = setTimeout(() => ctrl.abort(), 15000);
            const res = await fetch(`${provider.baseUrl.replace(/\/$/, "")}/chat/completions`, {
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
            });
            clearTimeout(t);
            if (!res.ok) {
                const text = await res.text().catch(() => "");
                return { ok: false, message: `HTTP ${res.status}: ${text.slice(0, 300)}` };
            }
            return { ok: true };
        }
        catch (err) {
            return {
                ok: false,
                message: err instanceof Error ? err.message : String(err),
            };
        }
    });
}
//# sourceMappingURL=settingsStore.js.map