import { defineConfig } from "vitest/config";

// The app's vite.config.ts wires up vite-plugin-electron, which must not run
// under vitest, so the test run gets its own (renderer-pure) config.
export default defineConfig({
  css: { postcss: { plugins: [] } },
  test: {
    environment: "node",
    include: ["src/**/*.test.{ts,tsx}"],
  },
});
