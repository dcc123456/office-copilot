import { defineConfig } from "vitest/config";

// The app's vite.config.ts wires up vite-plugin-electron, which must not run
// under vitest, so the test run gets its own (renderer-pure) config.
export default defineConfig({
  css: { postcss: { plugins: [] } },
  test: {
    environment: "node",
    include: ["src/**/*.test.{ts,tsx}"],
    // 打印 HTML 用 `?raw` 引 src/styles/dropCap.css（画布与纸上是同一条 ::first-letter
    // 规则）。Vitest 缺省把 CSS 换成空桩，单测里那份就会是空串——用例过的等于没测。
    css: true,
  },
});
