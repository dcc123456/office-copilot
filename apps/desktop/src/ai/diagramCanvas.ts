import { layoutDiagram, type DiagramLayoutResult, type LaidOutNode } from "./diagramSpec.js";
import type { DiagramSpec } from "./diagramSpec.js";

/**
 * Turns a declarative diagram into pixels on an HTML canvas, then into a PNG
 * data URL. The data URL goes into the existing image node, which docx-io already
 * knows how to embed — so an AI illustration survives a save to .docx.
 */

const SCALE = 2;

export interface RenderedImage {
  src: string;
  width: number;
  height: number;
  layout: DiagramLayoutResult;
}

const PALETTE = {
  ink: "#1f2328",
  muted: "#5b6572",
  line: "#8d97a5",
  fill: "#ffffff",
  accent: "#eef3fb",
  border: "#9db4d0",
};

export function drawDiagram(ctx: CanvasRenderingContext2D, layout: DiagramLayoutResult): void {
  const { width, height } = layout;
  ctx.save();
  ctx.scale(SCALE, SCALE);
  ctx.clearRect(0, 0, width, height);
  ctx.fillStyle = PALETTE.fill;
  ctx.fillRect(0, 0, width, height);
  ctx.textBaseline = "middle";
  ctx.textAlign = "center";

  if (layout.title && layout.titleAt) {
    ctx.fillStyle = PALETTE.ink;
    ctx.font = `600 16px "Microsoft YaHei", "PingFang SC", sans-serif`;
    ctx.fillText(layout.title, layout.titleAt.x, layout.titleAt.y);
  }

  // Edges first so boxes sit on top of the connector ends.
  for (const e of layout.edges) {
    ctx.strokeStyle = PALETTE.line;
    ctx.fillStyle = PALETTE.line;
    ctx.lineWidth = 1.4;
    ctx.beginPath();
    e.points.forEach((p, i) => (i === 0 ? ctx.moveTo(p.x, p.y) : ctx.lineTo(p.x, p.y)));
    ctx.stroke();

    const ah = 7;
    ctx.beginPath();
    ctx.moveTo(e.head.x, e.head.y);
    ctx.lineTo(
      e.head.x - ah * Math.cos(e.angle - Math.PI / 7),
      e.head.y - ah * Math.sin(e.angle - Math.PI / 7),
    );
    ctx.lineTo(
      e.head.x - ah * Math.cos(e.angle + Math.PI / 7),
      e.head.y - ah * Math.sin(e.angle + Math.PI / 7),
    );
    ctx.closePath();
    ctx.fill();

    if (e.label) {
      ctx.font = `12px "Microsoft YaHei", "PingFang SC", sans-serif`;
      const w = ctx.measureText(e.label).width + 8;
      ctx.fillStyle = PALETTE.fill;
      ctx.fillRect(e.labelAt.x - w / 2, e.labelAt.y - 8, w, 16);
      ctx.fillStyle = PALETTE.muted;
      ctx.fillText(e.label, e.labelAt.x, e.labelAt.y);
    }
  }

  for (const n of layout.nodes) drawBox(ctx, n);
  ctx.restore();
}

function drawBox(ctx: CanvasRenderingContext2D, n: LaidOutNode): void {
  ctx.lineWidth = 1.4;
  ctx.strokeStyle = PALETTE.border;
  ctx.fillStyle = PALETTE.accent;
  ctx.beginPath();
  if (n.shape === "round") {
    const r = Math.min(12, n.h / 2);
    roundedRect(ctx, n.x, n.y, n.w, n.h, r);
  } else if (n.shape === "diamond") {
    const cx = n.x + n.w / 2;
    const cy = n.y + n.h / 2;
    ctx.moveTo(cx, n.y);
    ctx.lineTo(n.x + n.w, cy);
    ctx.lineTo(cx, n.y + n.h);
    ctx.lineTo(n.x, cy);
    ctx.closePath();
  } else {
    ctx.rect(n.x, n.y, n.w, n.h);
  }
  ctx.fill();
  ctx.stroke();

  ctx.fillStyle = PALETTE.ink;
  ctx.font = `14px "Microsoft YaHei", "PingFang SC", sans-serif`;
  const lineH = 19;
  const startY = n.shape === "diamond" ? n.y + n.h / 2 : n.y + 10 + lineH / 2;
  n.lines.forEach((line, i) => {
    const y =
      n.shape === "diamond"
        ? startY + (i - (n.lines.length - 1) / 2) * lineH
        : startY + i * lineH;
    ctx.fillText(line, n.x + n.w / 2, y);
  });
}

function roundedRect(
  ctx: CanvasRenderingContext2D,
  x: number,
  y: number,
  w: number,
  h: number,
  r: number,
): void {
  ctx.moveTo(x + r, y);
  ctx.lineTo(x + w - r, y);
  ctx.quadraticCurveTo(x + w, y, x + w, y + r);
  ctx.lineTo(x + w, y + h - r);
  ctx.quadraticCurveTo(x + w, y + h, x + w - r, y + h);
  ctx.lineTo(x + r, y + h);
  ctx.quadraticCurveTo(x, y + h, x, y + h - r);
  ctx.lineTo(x, y + r);
  ctx.quadraticCurveTo(x, y, x + r, y);
  ctx.closePath();
}

export function renderDiagram(spec: DiagramSpec): RenderedImage {
  const layout = layoutDiagram(spec);
  const canvas = document.createElement("canvas");
  canvas.width = Math.max(1, Math.round(layout.width * SCALE));
  canvas.height = Math.max(1, Math.round(layout.height * SCALE));
  const ctx = canvas.getContext("2d");
  if (!ctx) throw new Error("无法创建画布上下文");
  drawDiagram(ctx, layout);
  return {
    src: canvas.toDataURL("image/png"),
    width: layout.width,
    height: layout.height,
    layout,
  };
}

/** DiagramResolver for the pure edit engine. */
export function diagramResolverFor(spec: DiagramSpec, _alt: string) {
  const img = renderDiagram(spec);
  return { src: img.src, width: img.width, height: img.height };
}
