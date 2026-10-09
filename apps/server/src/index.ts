import express from "express";
import cors from "cors";
import http from "node:http";
import https from "node:https";
import { z } from "zod";
import { config } from "./config.js";
import { runWordAgent } from "./agent.js";
import { recordHistory, listHistory } from "./history.js";
import type { CommandEvent, CommandRequest } from "@office-copilot/shared";

const app = express();
app.use(express.json({ limit: "4mb" }));
app.use(
  cors({
    origin: (origin, cb) => {
      // Allow same-origin / no-origin (curl, health checks) and configured origins.
      if (!origin || config.corsOrigins.includes(origin)) return cb(null, true);
      return cb(new Error(`CORS blocked: ${origin}`));
    },
  }),
);

app.get("/health", (_req, res) => {
  res.json({ ok: true, model: config.llm.model, baseUrl: config.llm.baseUrl });
});

const CommandSchema = z.object({
  sessionId: z.string().min(1),
  instruction: z.string().min(1),
  selection: z.string().optional(),
  outline: z
    .array(
      z.object({
        index: z.number(),
        text: z.string(),
        style: z.string().optional(),
      }),
    )
    .optional(),
});

function sse(res: express.Response, event: CommandEvent) {
  res.write(`data: ${JSON.stringify(event)}\n\n`);
}

app.post("/api/word/command", async (req, res) => {
  const parsed = CommandSchema.safeParse(req.body);
  if (!parsed.success) {
    res.status(400).json({ error: parsed.error.flatten() });
    return;
  }
  const request: CommandRequest = parsed.data;

  res.setHeader("Content-Type", "text/event-stream; charset=utf-8");
  res.setHeader("Cache-Control", "no-cache, no-transform");
  res.setHeader("Connection", "keep-alive");
  res.flushHeaders?.();

  const controller = new AbortController();
  req.on("close", () => controller.abort());

  try {
    sse(res, { event: "status", message: "已接收指令…" });
    let proposalSummary = "";
    const proposal = await runWordAgent(
      request,
      {
        onStatus: (m) => sse(res, { event: "status", message: m }),
        onToolCall: (name, args) => sse(res, { event: "tool_call", name, args }),
        onProposal: (p) => {
          proposalSummary = p.summary;
          sse(res, { event: "proposal", proposal: p });
        },
      },
      controller.signal,
    );

    if (proposal) {
      recordHistory({
        sessionId: request.sessionId,
        instruction: request.instruction,
        summary: proposal.summary,
      });
    } else {
      sse(res, {
        event: "error",
        message: "模型未生成修改提案，可以换一种说法再试。",
      });
    }
    sse(res, { event: "done" });
  } catch (e) {
    const message = e instanceof Error ? e.message : String(e);
    sse(res, { event: "error", message });
    sse(res, { event: "done" });
  } finally {
    res.end();
  }
});

app.get("/api/history/:sessionId", (req, res) => {
  res.json({ items: listHistory(req.params.sessionId) });
});

const server = config.tls
  ? https.createServer(config.tls, app)
  : http.createServer(app);

server.listen(config.port, () => {
  // eslint-disable-next-line no-console
  console.log(
    `[server] listening on ${config.protocol}://localhost:${config.port} (model=${config.llm.model})`,
  );
});
