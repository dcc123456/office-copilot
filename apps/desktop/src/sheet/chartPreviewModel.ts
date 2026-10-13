/**
 * 把引擎解析好的图表（`ResolvedChart`）翻成 Chart.js 的配置。
 *
 * 放在独立文件里而不是组件内，是因为桌面这套测试跑在 node 环境（没有 canvas），
 * 组件本身只做装配：数值一律已经在 `sheet-io` 里从现场表读过（`AGENTS.md` §2.11），
 * 这里不再算任何东西，只是换个形状交给画图库。
 */
import type { ResolvedChart } from "@office-copilot/sheet-io";

/** Chart.js 的图型名。我们的三种图型逐一对应（`bar` / `line` / `pie`）。 */
export type ChartPreviewType = "bar" | "line" | "pie";

/** 喂给 Chart.js 的一个数据集。 */
export interface ChartPreviewDataset {
  label: string;
  data: number[];
  /** 柱/线的系列色；饼图不填（它按扇区上色）。 */
  borderColor?: string;
  /** 饼图按扇区上色，柱图当填充色。 */
  backgroundColor?: string[] | string;
}

/** 一份完整的 Chart.js 配置（只声明我们真填的那几项）。 */
export interface ChartPreviewConfig {
  type: ChartPreviewType;
  data: { labels: string[]; datasets: ChartPreviewDataset[] };
  options: {
    responsive: boolean;
    maintainAspectRatio: boolean;
    plugins: {
      legend: { display: boolean };
      title: { display: boolean; text: string };
    };
  };
}

/**
 * 系列与扇区的配色。
 *
 * 画布上的颜色 CSS 够不到，只能写在这里（§5.4 管的是 DOM 样式）。不填的话 Chart.js 默认给的是
 * 浅灰，实跑截图里折线几乎看不见；顺序固定，同一张图每次画出来颜色一致。
 */
const SERIES_COLORS = ["#4e79a7", "#f28e2b", "#e15759", "#76b7b2", "#59a14f", "#edc948", "#b07aa1", "#ff9da7"];

/**
 * 生成预览配置。
 *
 * @param chart `resolveChartSpec` 的产物；类别是显示串、数值已由引擎从现场表读好。
 * @returns Chart.js 的配置；`widthPx`/`heightPx` 不在这里出现，画布尺寸由外层容器的 CSS 决定（§5.4）。
 */
export function chartPreviewConfig(chart: ResolvedChart): ChartPreviewConfig {
  const datasets: ChartPreviewDataset[] = chart.series.map((series, index) => {
    const color = SERIES_COLORS[index % SERIES_COLORS.length]!;
    return {
      label: series.name,
      data: [...series.values],
      ...(chart.kind === "pie" ? { backgroundColor: [...SERIES_COLORS] } : { borderColor: color, backgroundColor: color }),
    };
  });
  return {
    type: chart.kind as ChartPreviewType,
    data: { labels: [...chart.categories], datasets },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        // 饼图的图例就是类别名，单系列柱/线只有一条，图例没有信息量。
        legend: { display: chart.kind === "pie" || chart.series.length > 1 },
        title: { display: chart.title.length > 0, text: chart.title },
      },
    },
  };
}
