/**
 * 图表预览（spec 决策 D5 的 app 内那一半）：把引擎解析好的 `ResolvedChart` 画进一块 canvas。
 *
 * 只画，不改：数据一律是 `resolveChartSpec` 从现场表读好的引用与数值（`AGENTS.md` §2.11），
 * 写进 `.xlsx` 的那份仍是原生 chart part，这张图只是给人先看一眼（§七 的"目视没有通道"由它补一半）。
 */
import { useEffect, useRef } from "react";
import {
  ArcElement,
  BarController,
  BarElement,
  CategoryScale,
  Chart,
  Legend,
  LinearScale,
  LineController,
  LineElement,
  PieController,
  PointElement,
  Title,
  Tooltip,
  type ChartConfiguration,
} from "chart.js";
import type { ResolvedChart } from "@office-copilot/sheet-io";
import { chartPreviewConfig } from "./chartPreviewModel.js";

// 按需注册而不是 `chart.js/auto`：后者会把二十种图型与全部比例尺都打进桌面产物。
Chart.register(BarController, LineController, PieController, BarElement, LineElement, PointElement, ArcElement, CategoryScale, LinearScale, Legend, Title, Tooltip);

/**
 * 一张图表的预览画布。
 *
 * @param chart 引擎解析过的图表；引用与数值都已核对。
 * @returns 带 canvas 的容器，尺寸由 `.ai-chart-preview` 的 CSS 决定（不写内联样式，§5.4）。
 */
export function ChartPreview({ chart }: { chart: ResolvedChart }) {
  const canvasRef = useRef<HTMLCanvasElement>(null);

  useEffect(() => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    // 配置里只填了我们认识的字段，Chart.js 的泛型类型比它宽，这里按它的入口形状交进去。
    const instance = new Chart(canvas, chartPreviewConfig(chart) as unknown as ChartConfiguration);
    return () => instance.destroy();
  }, [chart]);

  return (
    <div className="ai-chart-preview">
      <canvas ref={canvasRef} aria-label={`图表预览：${chart.title || chart.kind}`} />
    </div>
  );
}
