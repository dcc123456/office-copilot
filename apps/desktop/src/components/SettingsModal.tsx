import { useEffect, useState } from "react";
import { X, Plus, Trash2, Check, AlertCircle } from "lucide-react";
import type { AppSettings, ProviderConfig } from "../global.js";
import { THEMES, useUiStore, type ThemeId } from "../uiStore.js";

interface Preset {
  name: string;
  baseUrl: string;
  model: string;
  extraBody?: Record<string, unknown>;
  apiKey?: string;
}

const PRESETS: Preset[] = [
  { name: "DeepSeek", baseUrl: "https://api.deepseek.com/v1", model: "deepseek-chat" },
  { name: "DeepSeek Reasoner", baseUrl: "https://api.deepseek.com/v1", model: "deepseek-reasoner" },
  {
    name: "火山方舟（豆包）",
    baseUrl: "https://ark.cn-beijing.volces.com/api/v3",
    model: "",
    extraBody: { thinking: { type: "disabled" } },
  },
  { name: "智谱 GLM", baseUrl: "https://open.bigmodel.cn/api/paas/v4", model: "glm-4-plus" },
  { name: "Moonshot Kimi", baseUrl: "https://api.moonshot.cn/v1", model: "moonshot-v1-32k" },
  { name: "阿里通义千问", baseUrl: "https://dashscope.aliyuncs.com/compatible-mode/v1", model: "qwen-plus" },
  { name: "OpenAI", baseUrl: "https://api.openai.com/v1", model: "gpt-4o-mini" },
  { name: "本地 Ollama", baseUrl: "http://127.0.0.1:11434/v1", model: "qwen2.5:7b", apiKey: "ollama" },
  { name: "自定义（OpenAI 兼容）", baseUrl: "", model: "" },
];

export function SettingsModal({
  open,
  onClose,
  onSaved,
}: {
  open: boolean;
  onClose: () => void;
  onSaved?: () => void;
}) {
  const [settings, setSettings] = useState<AppSettings | null>(null);
  const [testingId, setTestingId] = useState<string | null>(null);
  const [resultFor, setResultFor] = useState<Record<string, { ok: boolean; msg: string }>>({});
  const theme = useUiStore((s) => s.theme);
  const setTheme = useUiStore((s) => s.setTheme);

  useEffect(() => {
    if (open) {
      window.officeCopilot.getSettings().then((s) => setSettings(s));
    }
  }, [open]);

  useEffect(() => {
    if (!open) return;
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Escape") onClose();
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [open, onClose]);

  if (!open || !settings) return null;

  const updateProvider = (id: string, patch: Partial<ProviderConfig>) => {
    setSettings({
      ...settings,
      providers: settings.providers.map((p) => (p.id === id ? { ...p, ...patch } : p)),
    });
  };

  const removeProvider = (id: string) => {
    setSettings({
      ...settings,
      providers: settings.providers.filter((p) => p.id !== id),
      defaultProviderId: settings.defaultProviderId === id ? null : settings.defaultProviderId,
    });
  };

  const addFromPreset = (preset: Preset) => {
    const id = "p_" + Math.random().toString(36).slice(2, 10);
    const provider: ProviderConfig = {
      id,
      name: preset.name,
      baseUrl: preset.baseUrl,
      apiKey: preset.apiKey ?? "",
      model: preset.model,
      extraBody: preset.extraBody,
      enabled: true,
    };
    setSettings({
      ...settings,
      providers: [...settings.providers, provider],
      defaultProviderId: settings.defaultProviderId ?? id,
    });
  };

  const save = async () => {
    await window.officeCopilot.saveSettings(settings);
    onSaved?.();
    onClose();
  };

  const testProvider = async (p: ProviderConfig) => {
    setTestingId(p.id);
    const r = await window.officeCopilot.testProvider(p);
    setResultFor((m) => ({ ...m, [p.id]: r.ok ? { ok: true, msg: "连接成功" } : { ok: false, msg: r.message } }));
    setTestingId(null);
  };

  return (
    <div className="settings-backdrop" onClick={onClose}>
      <div
        className="settings-modal"
        role="dialog"
        aria-modal="true"
        aria-label="设置"
        onClick={(e) => e.stopPropagation()}
      >
        <div className="settings-header">
          <strong>设置</strong>
          <button onClick={onClose} style={{ marginLeft: "auto" }} className="icon" aria-label="关闭">
            <X size={16} />
          </button>
        </div>

        <div className="settings-body">
          <div className="settings-section-title">外观 · 主题</div>
          <div className="theme-picker">
            {THEMES.map((t) => (
              <button
                key={t.id}
                className={"theme-swatch" + (theme === t.id ? " active" : "")}
                onClick={() => setTheme(t.id as ThemeId)}
              >
                <span className="theme-swatch-dot" style={{ background: t.swatch }} />
                {t.label}
              </button>
            ))}
          </div>

          <div className="settings-section-title">模型</div>
          <div style={{ marginBottom: 16 }}>
            <div className="settings-hint">
              从预设添加（支持任何 OpenAI 兼容接口）：
            </div>
            <div className="preset-chips">
              {PRESETS.map((p) => (
                <button key={p.name} className="chip" onClick={() => addFromPreset(p)}>
                  <Plus size={10} /> {p.name}
                </button>
              ))}
            </div>
          </div>

          {settings.providers.length === 0 ? (
            <div className="muted-note">还没有配置模型。点上方预设添加一个。</div>
          ) : (
            settings.providers.map((p) => (
              <div
                key={p.id}
                className={"provider-card" + (settings.defaultProviderId === p.id ? " is-default" : "")}
              >
                <div className="provider-row">
                  <input
                    type="radio"
                    checked={settings.defaultProviderId === p.id}
                    onChange={() => setSettings({ ...settings, defaultProviderId: p.id })}
                  />
                  <input
                    className="field-input"
                    value={p.name}
                    onChange={(e) => updateProvider(p.id, { name: e.target.value })}
                    style={{ flex: 1, fontWeight: 600 }}
                    placeholder="名称"
                  />
                  <label style={{ fontSize: 12 }}>
                    <input
                      type="checkbox"
                      checked={p.enabled}
                      onChange={(e) => updateProvider(p.id, { enabled: e.target.checked })}
                    />{" "}
                    启用
                  </label>
                  <button
                    onClick={() => testProvider(p)}
                    disabled={testingId === p.id}
                    title="测试连接"
                  >
                    {testingId === p.id ? "测试中…" : "测试"}
                  </button>
                  <button onClick={() => removeProvider(p.id)} className="icon" title="删除">
                    <Trash2 size={14} />
                  </button>
                </div>
                <div className="provider-grid">
                  <input
                    className="field-input"
                    value={p.baseUrl}
                    onChange={(e) => updateProvider(p.id, { baseUrl: e.target.value })}
                    placeholder="Base URL，如 https://api.deepseek.com/v1"
                  />
                  <input
                    className="field-input"
                    value={p.model}
                    onChange={(e) => updateProvider(p.id, { model: e.target.value })}
                    placeholder="模型名，如 deepseek-chat"
                  />
                </div>
                <input
                  className="field-input"
                  type="password"
                  value={p.apiKey}
                  onChange={(e) => updateProvider(p.id, { apiKey: e.target.value })}
                  placeholder="API Key"
                  style={{ width: "100%" }}
                />
                {resultFor[p.id] && (() => {
                  const r = resultFor[p.id]!;
                  return (
                    <div className={"result-note " + (r.ok ? "ok" : "err")}>
                      {r.ok ? <Check size={12} /> : <AlertCircle size={12} />}
                      {r.msg}
                    </div>
                  );
                })()}
                {p.id === settings.defaultProviderId && (
                  <div className="default-note">默认模型</div>
                )}
              </div>
            ))
          )}
          <details className="settings-details" style={{ marginTop: 12 }}>
            <summary>火山方舟 / 豆包 注意事项</summary>
            <div>
              火山方舟的 <code>model</code> 字段不是模型名，而是控制台创建的
              <strong>接入点 ID</strong>（形如 <code>ep-2024xxxx-xxxxx</code>）。
              在 https://console.volcengine.com/ark 创建在线推理接入点后把 ID 填入即可。
              Base URL 已自动设为 <code>https://ark.cn-beijing.volces.com/api/v3</code>。
            </div>
          </details>
        </div>

        <div className="settings-footer">
          <button onClick={onClose}>取消</button>
          <button className="primary" onClick={save}>
            保存
          </button>
        </div>
      </div>
    </div>
  );
}
