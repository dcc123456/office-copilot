/**
 * Provider/model configuration. All providers use the OpenAI-compatible
 * /chat/completions protocol; the only differences are base URL, API key,
 * and model id. Users can add arbitrary providers themselves.
 */

export interface ProviderConfig {
  id: string;
  name: string;
  baseUrl: string;
  apiKey: string;
  model: string;
  /** Optional override for "Authorization: Bearer ..." if a vendor needs it. */
  authHeader?: string;
  /** Extra JSON body fields, e.g. Ark requires `thinking: {type: "disabled"}`. */
  extraBody?: Record<string, unknown>;
  temperature?: number;
  enabled: boolean;
}

export interface AppSettings {
  version: 1;
  providers: ProviderConfig[];
  defaultProviderId: string | null;
  /** Optional proxy URL for fetch. */
  proxy?: string;
}

export const DEFAULT_SETTINGS: AppSettings = {
  version: 1,
  providers: [],
  defaultProviderId: null,
};

export const PRESETS: Array<Partial<Omit<ProviderConfig, "id" | "enabled">> &
  Pick<ProviderConfig, "name" | "baseUrl" | "model">> = [
  {
    name: "DeepSeek",
    baseUrl: "https://api.deepseek.com/v1",
    model: "deepseek-chat",
  },
  {
    name: "DeepSeek Reasoner",
    baseUrl: "https://api.deepseek.com/v1",
    model: "deepseek-reasoner",
  },
  {
    name: "火山方舟（豆包）",
    baseUrl: "https://ark.cn-beijing.volces.com/api/v3",
    model: "",
    extraBody: { thinking: { type: "disabled" } },
  },
  {
    name: "智谱 GLM",
    baseUrl: "https://open.bigmodel.cn/api/paas/v4",
    model: "glm-4-plus",
  },
  {
    name: "Moonshot Kimi",
    baseUrl: "https://api.moonshot.cn/v1",
    model: "moonshot-v1-32k",
  },
  {
    name: "阿里通义千问",
    baseUrl: "https://dashscope.aliyuncs.com/compatible-mode/v1",
    model: "qwen-plus",
  },
  {
    name: "OpenAI",
    baseUrl: "https://api.openai.com/v1",
    model: "gpt-4o-mini",
  },
  {
    name: "本地 Ollama",
    baseUrl: "http://127.0.0.1:11434/v1",
    model: "qwen2.5:7b",
    apiKey: "ollama",
  },
  {
    name: "自定义（OpenAI 兼容）",
    baseUrl: "",
    model: "",
  },
];

export function newProviderId(): string {
  return "p_" + Math.random().toString(36).slice(2, 10);
}
