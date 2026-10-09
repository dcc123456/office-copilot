import "dotenv/config";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// The task pane is served over HTTPS (Vite), so browsers forbid calling an
// http:// API. We boot the server on HTTPS using the dev certificate that
// `office-addin-dev-certs install` generates. If the cert is missing we fall
// back to plain HTTP so local curl testing still works.
const certDir = path.join(os.homedir(), ".office-addin-dev-certs");
const certFile = path.join(certDir, "localhost.crt");
const keyFile = path.join(certDir, "localhost.key");
const hasDevCerts = fs.existsSync(certFile) && fs.existsSync(keyFile);

export const config = {
  port: Number(process.env.PORT ?? 3001),
  protocol: hasDevCerts ? ("https" as const) : ("http" as const),
  corsOrigins: (process.env.CORS_ORIGINS ?? "https://localhost:5173,https://127.0.0.1:5173")
    .split(",")
    .map((s) => s.trim())
    .filter(Boolean),
  llm: {
    baseUrl: process.env.LLM_BASE_URL ?? "https://api.deepseek.com/v1",
    apiKey: process.env.LLM_API_KEY ?? "",
    model: process.env.LLM_MODEL ?? "deepseek-chat",
  },
  dataDir: path.resolve(__dirname, "..", "data"),
  maxToolRounds: 6,
  tls: hasDevCerts
    ? {
        cert: fs.readFileSync(certFile),
        key: fs.readFileSync(keyFile),
      }
    : null,
};
