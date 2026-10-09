"use strict";
/**
 * Provider/model configuration. All providers use the OpenAI-compatible
 * /chat/completions protocol; the only differences are base URL, API key,
 * and model id. Users can add arbitrary providers themselves.
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.PRESETS = exports.DEFAULT_SETTINGS = void 0;
exports.newProviderId = newProviderId;
exports.DEFAULT_SETTINGS = {
    version: 1,
    providers: [],
    defaultProviderId: null,
};
exports.PRESETS = [
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
function newProviderId() {
    return "p_" + Math.random().toString(36).slice(2, 10);
}
//# sourceMappingURL=settings.js.map