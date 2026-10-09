import { defineConfig } from "vitest/config";

// The host machine has a ~/postcss.config.js that pulls in Tailwind. We don't
// use Tailwind in this monorepo, and Vite would otherwise try to load it and
// fail. Provide an explicit (empty) PostCSS config and disable CSS.
export default defineConfig({
  css: {
    postcss: { plugins: [] },
  },
  test: {
    environment: "node",
    passWithNoTests: true,
  },
});
