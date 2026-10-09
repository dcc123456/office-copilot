import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import electron from "vite-plugin-electron/simple";
import path from "node:path";

// The host machine has a ~/postcss.config.js referencing Tailwind; we don't
// use it in this app, so override to empty.
export default defineConfig({
  plugins: [
    react(),
    electron({
      main: {
        entry: "electron/main.ts",
        vite: {
          build: { outDir: "dist-electron" },
          css: { postcss: { plugins: [] } },
        },
      },
      preload: {
        input: path.join(__dirname, "electron/preload.ts"),
        vite: { build: { outDir: "dist-electron" }, css: { postcss: { plugins: [] } } },
      },
      renderer: {},
    }),
  ],
  css: { postcss: { plugins: [] } },
  server: {
    port: 5174,
    strictPort: true,
  },
  build: {
    outDir: "dist",
    sourcemap: true,
  },
});
