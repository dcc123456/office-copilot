/**
 * Declarative diagram the AI can ask for ("add a flow chart explaining this
 * process"), plus a pure layout pass. Keeping layout here means it is testable
 * without a DOM; `diagramCanvas.ts` only turns the geometry into pixels.
 */

export type DiagramShape = "box" | "round" | "diamond";
export type DiagramLayout = "flow-v" | "flow-h" | "grid";

export interface DiagramNodeSpec {
  id: string;
  label: string;
  shape: DiagramShape;
}

export interface DiagramEdgeSpec {
  from: string;
  to: string;
  label?: string;
}

export interface DiagramSpec {
  layout: DiagramLayout;
  nodes: DiagramNodeSpec[];
  edges: DiagramEdgeSpec[];
  title?: string;
  width?: number;
}

export interface Pt {
  x: number;
  y: number;
}

export interface LaidOutNode {
  id: string;
  label: string;
  shape: DiagramShape;
  x: number;
  y: number;
  w: number;
  h: number;
  lines: string[];
  layer: number;
}

export interface LaidOutEdge {
  points: Pt[];
  label?: string;
  labelAt: Pt;
  head: Pt;
  angle: number;
}

export interface DiagramLayoutResult {
  width: number;
  height: number;
  nodes: LaidOutNode[];
  edges: LaidOutEdge[];
  title?: string;
  titleAt?: Pt;
}

const FONT_PX = 14;
const LINE_H = 19;
const PAD_X = 14;
const PAD_Y = 10;
const GAP_MAIN = 30;
const GAP_LANE = 46;
const MARGIN = 16;
const MAX_BOX_W = 220;
const MIN_BOX_W = 84;
const TITLE_H = 26;

/** Approximate advance width, so layout is possible before we own a canvas. */
export function approxTextWidth(s: string): number {
  let w = 0;
  for (const ch of s) {
    const c = ch.codePointAt(0) ?? 0;
    if (c > 0x2e80) w += FONT_PX; // CJK & full-width
    else if (/[A-Z0-9]/.test(ch)) w += FONT_PX * 0.68;
    else w += FONT_PX * 0.52;
  }
  return w;
}

export function wrapLabel(label: string, maxW: number): string[] {
  const limit = Math.max(MIN_BOX_W, maxW) - PAD_X * 2;
  const out: string[] = [];
  let line = "";
  for (const ch of label) {
    if (ch === "\n") {
      if (line) out.push(line);
      line = "";
      continue;
    }
    const next = line + ch;
    if (line && approxTextWidth(next) > limit) {
      out.push(line);
      line = ch;
    } else {
      line = next;
    }
  }
  if (line) out.push(line);
  return out.length ? out : [""];
}

/**
 * Layer assignment: longest path from the roots. Specs with cycles fall back to
 * declaration order so every node still gets a sane slot.
 */
export function assignLayers(spec: DiagramSpec): number[] {
  const n = spec.nodes.length;
  const incoming: number[][] = spec.nodes.map(() => []);
  const indexOf = new Map(spec.nodes.map((nd, i) => [nd.id, i]));
  for (const e of spec.edges) {
    const a = indexOf.get(e.from);
    const b = indexOf.get(e.to);
    if (a === undefined || b === undefined || a === b) continue;
    incoming[b]!.push(a);
  }

  const layer = new Array<number>(n).fill(0);
  const done = new Array<boolean>(n).fill(false);
  let progressed = true;
  while (progressed) {
    progressed = false;
    for (let i = 0; i < n; i++) {
      if (done[i]) continue;
      if (incoming[i]!.some((p) => !done[p])) continue;
      layer[i] = incoming[i]!.reduce((m, p) => Math.max(m, layer[p]! + 1), 0);
      done[i] = true;
      progressed = true;
    }
  }
  for (let i = 0; i < n; i++) {
    if (done[i]) continue;
    const preds = incoming[i]!.map((p) => layer[p]! + 1);
    layer[i] = preds.length ? Math.max(...preds) : i;
  }
  return layer;
}

export function layoutDiagram(spec: DiagramSpec): DiagramLayoutResult {
  const nodes = spec.nodes.length ? spec.nodes : [];
  let laneOf = assignLayers({ ...spec, nodes });

  if (spec.layout === "grid" && nodes.length) {
    const cols = Math.max(1, Math.ceil(Math.sqrt(nodes.length)));
    laneOf = nodes.map((_, i) => Math.floor(i / cols));
  }

  const laneCount = laneOf.reduce((m, l) => Math.max(m, l), 0) + 1;
  const lanes: number[][] = Array.from({ length: laneCount }, () => []);
  laneOf.forEach((l, i) => {
    lanes[l]!.push(i);
  });

  // flow-v / grid: lanes stack vertically, nodes run left-to-right inside a lane.
  // flow-h: lanes stack horizontally, nodes run top-to-bottom inside a lane.
  const sideways = spec.layout === "flow-h";
  const mainSize = (b: { w: number; h: number }) => (sideways ? b.h : b.w);
  const crossSize = (b: { w: number; h: number }) => (sideways ? b.w : b.h);

  const boxMaxW = Math.max(140, Math.min(MAX_BOX_W, (spec.width ?? MAX_BOX_W) - MARGIN * 2));
  const boxes: LaidOutNode[] = nodes.map((nd, i) => {
    const lines = wrapLabel(nd.label, boxMaxW);
    const widest = lines.reduce((m, s) => Math.max(m, approxTextWidth(s)), 0);
    const w = Math.min(boxMaxW, Math.max(MIN_BOX_W, Math.ceil(widest) + PAD_X * 2));
    const extra = nd.shape === "diamond" ? LINE_H : 0;
    return {
      id: nd.id,
      label: nd.label,
      shape: nd.shape,
      x: 0,
      y: 0,
      w,
      h: lines.length * LINE_H + PAD_Y * 2 + extra,
      lines,
      layer: laneOf[i] ?? 0,
    };
  });

  const titleOffset = spec.title ? TITLE_H + 8 : 0;
  let laneCursor = MARGIN + titleOffset;
  lanes.forEach((lane) => {
    const members = lane.map((i) => boxes[i]!);
    const extent = members.reduce((m, b) => Math.max(m, crossSize(b)), 0);
    let cursor = MARGIN;
    members.forEach((b) => {
      if (sideways) {
        b.y = cursor;
        b.x = laneCursor + (extent - b.w) / 2;
      } else {
        b.x = cursor;
        b.y = laneCursor + (extent - b.h) / 2;
      }
      cursor += mainSize(b) + GAP_MAIN;
    });
    laneCursor += extent + GAP_LANE;
  });

  const width = boxes.reduce((m, b) => Math.max(m, b.x + b.w + MARGIN), MARGIN * 2);
  const height = boxes.reduce((m, b) => Math.max(m, b.y + b.h + MARGIN), MARGIN * 2 + titleOffset);

  const byId = new Map(boxes.map((b) => [b.id, b]));
  const edges: LaidOutEdge[] = [];
  for (const e of spec.edges) {
    const a = byId.get(e.from);
    const b = byId.get(e.to);
    if (!a || !b || a === b) continue;
    const routed = routeEdge(a, b);
    if (!routed) continue;
    edges.push(e.label ? { ...routed, label: e.label } : routed);
  }

  return {
    width: Math.round(width),
    height: Math.round(height),
    nodes: boxes,
    edges,
    ...(spec.title ? { title: spec.title, titleAt: { x: width / 2, y: MARGIN + 16 } } : {}),
  };
}

function center(b: LaidOutNode): Pt {
  return { x: b.x + b.w / 2, y: b.y + b.h / 2 };
}

function anchor(b: LaidOutNode, side: "top" | "bottom" | "left" | "right"): Pt {
  switch (side) {
    case "top":
      return { x: b.x + b.w / 2, y: b.y };
    case "bottom":
      return { x: b.x + b.w / 2, y: b.y + b.h };
    case "left":
      return { x: b.x, y: b.y + b.h / 2 };
    case "right":
      return { x: b.x + b.w, y: b.y + b.h / 2 };
  }
}

/** Connect the facing sides of two boxes with an orthogonal polyline. */
function routeEdge(a: LaidOutNode, b: LaidOutNode): Omit<LaidOutEdge, "label"> | null {
  const ca = center(a);
  const cb = center(b);
  const dx = cb.x - ca.x;
  const dy = cb.y - ca.y;
  if (dx === 0 && dy === 0) return null;

  let from: Pt;
  let to: Pt;
  if (Math.abs(dy) >= Math.abs(dx)) {
    from = anchor(a, dy > 0 ? "bottom" : "top");
    to = anchor(b, dy > 0 ? "top" : "bottom");
  } else {
    from = anchor(a, dx > 0 ? "right" : "left");
    to = anchor(b, dx > 0 ? "left" : "right");
  }

  const points: Pt[] =
    Math.abs(from.x - to.x) < 2 || Math.abs(from.y - to.y) < 2
      ? [from, to]
      : Math.abs(dy) >= Math.abs(dx)
        ? [from, { x: from.x, y: (from.y + to.y) / 2 }, { x: to.x, y: (from.y + to.y) / 2 }, to]
        : [from, { x: (from.x + to.x) / 2, y: from.y }, { x: (from.x + to.x) / 2, y: to.y }, to];

  const head = points[points.length - 1]!;
  const prev = points[points.length - 2] ?? head;
  return {
    points,
    head,
    angle: Math.atan2(head.y - prev.y, head.x - prev.x),
    labelAt: { x: (from.x + to.x) / 2, y: (from.y + to.y) / 2 },
  };
}
