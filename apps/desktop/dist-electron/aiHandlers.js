"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.registerAIHandlers = registerAIHandlers;
exports.callLLM = callLLM;
const settingsStore_js_1 = require("./settingsStore.js");
const SYSTEM_BY_ACTION = {
    扩写: "你是一个中文写作助手。把用户给的文字扩写得更详细、更完整，保留原意，只输出改写后的正文，不要解释。",
    缩写: "你是一个中文写作助手。把用户给的文字缩写到一半长度以内，保留关键信息，只输出缩写后的正文，不要解释。",
    润色: "你是一个中文写作助手。润色用户给的文字，让它更通顺、更专业，但不改变意思。只输出润色后的正文。",
    "换语气:正式商务": "你是一个商务文书编辑。把用户给的文字改写为正式、严谨的商务语气，保留原意。只输出改写后的正文。",
    "换语气:口语化": "你是一个中文写作助手。把用户给的文字改写为自然、口语化但仍然专业的表达，保留原意。只输出改写后的正文。",
    "翻译:英文": "You are a professional translator. Translate the user's Chinese text into natural English. Output only the translation, no explanation.",
    "翻译:中文": "你是专业译者。把用户给的英文翻译成自然流畅的简体中文。只输出译文，不要解释。",
    "排版:规范标题": "你是文档排版助手。把用户给的标题改写为更规范、专业的中文标题，不超过20字，只输出标题。",
    chat: "你是 Office Copilot 里的写作助手。回答简洁、中文，必要时分点。可以直接给出可粘贴到文档的成稿。",
};
function buildMessages(req) {
    const system = req.system ?? SYSTEM_BY_ACTION[req.action] ?? "你是一个中文写作助手。按用户要求改写文字，只输出结果。";
    const userContent = req.action === "chat"
        ? req.text
        : req.instruction
            ? `${req.instruction}\n\n原文：\n${req.text}`
            : req.text;
    return [
        { role: "system", content: system },
        { role: "user", content: userContent },
    ];
}
function chatUrl(provider) {
    return `${provider.baseUrl.replace(/\/$/, "")}/chat/completions`;
}
async function resolveProvider(providerId) {
    const settings = await (0, settingsStore_js_1.loadSettings)();
    return providerId
        ? (settings.providers.find((p) => p.id === providerId && p.enabled) ?? null)
        : (0, settingsStore_js_1.getActiveProviderSync)();
}
const NO_PROVIDER_MESSAGE = "还没有配置模型。点右上角⚙️添加一个模型服务商（DeepSeek、火山方舟、GLM、Kimi 等都支持）。";
function registerAIHandlers(ipc) {
    ipc.handle("ai:transform", async (_e, req) => {
        try {
            const provider = await resolveProvider(req.providerId);
            if (!provider) {
                return {
                    ok: false,
                    needsSetup: true,
                    message: NO_PROVIDER_MESSAGE,
                };
            }
            return callLLM(provider, req);
        }
        catch (err) {
            return { ok: false, message: err instanceof Error ? err.message : String(err) };
        }
    });
    const inflight = new Map();
    ipc.handle("ai:stream-start", async (e, payload) => {
        const { requestId, ...req } = payload;
        const wc = e.sender;
        const send = (ev) => {
            if (!wc.isDestroyed())
                wc.send("ai:stream", ev);
        };
        let provider;
        try {
            provider = await resolveProvider(req.providerId);
        }
        catch (err) {
            return {
                ok: false,
                message: err instanceof Error ? err.message : String(err),
            };
        }
        if (!provider) {
            return { ok: false, needsSetup: true, message: NO_PROVIDER_MESSAGE };
        }
        const ctrl = new AbortController();
        const entry = { ctrl, canceled: false };
        inflight.set(requestId, entry);
        const t = setTimeout(() => ctrl.abort(), 180_000);
        // Resolve immediately; progress arrives via the "ai:stream" events.
        void runStream(provider, req, ctrl.signal, (text) => send({ requestId, kind: "delta", text }))
            .then(() => send({ requestId, kind: "done" }))
            .catch((err) => {
            if (entry.canceled) {
                send({ requestId, kind: "done" });
                return;
            }
            const aborted = err instanceof Error && err.name === "AbortError";
            const message = aborted
                ? "请求超时（180 秒未完成）。"
                : err instanceof Error
                    ? err.message
                    : String(err);
            send({ requestId, kind: "error", message });
        })
            .finally(() => {
            clearTimeout(t);
            inflight.delete(requestId);
        });
        return { ok: true };
    });
    ipc.handle("ai:stream-cancel", (_e, requestId) => {
        const entry = inflight.get(requestId);
        if (entry) {
            entry.canceled = true;
            entry.ctrl.abort();
        }
        return { ok: true };
    });
}
async function runStream(provider, req, signal, onDelta) {
    const res = await fetch(chatUrl(provider), {
        method: "POST",
        headers: {
            "Content-Type": "application/json",
            Authorization: `Bearer ${provider.apiKey ?? ""}`,
        },
        body: JSON.stringify({
            model: provider.model,
            temperature: provider.temperature ?? 0.3,
            messages: buildMessages(req),
            stream: true,
            ...(provider.extraBody ?? {}),
        }),
    });
    if (!res.ok) {
        const text = await res.text().catch(() => "");
        throw new Error(`LLM ${res.status}: ${text.slice(0, 500)}`);
    }
    const ctype = res.headers.get("content-type") ?? "";
    // Gateways that ignore `stream:true` answer with a single JSON body.
    if (ctype.includes("application/json")) {
        const data = (await res.json());
        const text = data.choices?.[0]?.message?.content?.trim() ?? "";
        if (!text)
            throw new Error("模型返回为空。");
        onDelta(text);
        return;
    }
    if (!res.body)
        throw new Error("模型返回为空。");
    const reader = res.body.getReader();
    const decoder = new TextDecoder();
    let buffer = "";
    let received = "";
    let doneByServer = false;
    for (;;) {
        const { value, done } = await reader.read();
        if (done)
            break;
        buffer += decoder.decode(value, { stream: true });
        buffer = buffer.replace(/\r\n/g, "\n");
        // SSE frames are separated by a blank line.
        let sep;
        while ((sep = buffer.indexOf("\n\n")) !== -1) {
            const frame = buffer.slice(0, sep);
            buffer = buffer.slice(sep + 2);
            for (const line of frame.split("\n")) {
                const trimmed = line.trim();
                if (!trimmed.startsWith("data:"))
                    continue;
                const payload = trimmed.slice(5).trim();
                if (payload === "[DONE]") {
                    doneByServer = true;
                    continue;
                }
                let json;
                try {
                    json = JSON.parse(payload);
                }
                catch {
                    continue; // keepalives / malformed frames
                }
                const chunk = json.choices?.[0]?.delta?.content;
                if (typeof chunk === "string" && chunk) {
                    received += chunk;
                    onDelta(chunk);
                }
            }
        }
        if (doneByServer)
            break;
    }
    if (!received.trim())
        throw new Error("模型返回为空。");
}
async function callLLM(provider, req) {
    const ctrl = new AbortController();
    const t = setTimeout(() => ctrl.abort(), 120_000);
    try {
        const res = await fetch(chatUrl(provider), {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
                Authorization: `Bearer ${provider.apiKey ?? ""}`,
            },
            body: JSON.stringify({
                model: provider.model,
                temperature: provider.temperature ?? 0.3,
                messages: buildMessages(req),
                ...(provider.extraBody ?? {}),
            }),
            signal: ctrl.signal,
        });
        if (!res.ok) {
            const text = await res.text().catch(() => "");
            return { ok: false, message: `LLM ${res.status}: ${text.slice(0, 500)}` };
        }
        const data = (await res.json());
        const text = data.choices?.[0]?.message?.content?.trim() ?? "";
        if (!text)
            return { ok: false, message: "模型返回为空。" };
        return { ok: true, text };
    }
    catch (err) {
        return { ok: false, message: err instanceof Error ? err.message : String(err) };
    }
    finally {
        clearTimeout(t);
    }
}
//# sourceMappingURL=aiHandlers.js.map