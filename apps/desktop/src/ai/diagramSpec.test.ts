import { describe, expect, it } from "vitest";
import { approxTextWidth, assignLayers, layoutDiagram, wrapLabel, type DiagramSpec } from "./diagramSpec.js";

const flow: DiagramSpec = {
  layout: "flow-v",
  title: "审批流程",
  nodes: [
    { id: "a", label: "提交申请", shape: "round" },
    { id: "b", label: "主管审核", shape: "box" },
    { id: "c", label: "是否通过？", shape: "diamond" },
    { id: "d", label: "归档", shape: "box" },
  ],
  edges: [
    { from: "a", to: "b" },
    { from: "b", to: "c" },
    { from: "c", to: "d", label: "是" },
    { from: "c", to: "b", label: "否" },
  ],
};

const overlaps = (l: ReturnType<typeof layoutDiagram>) => {
  for (let i = 0; i < l.nodes.length; i++) {
    for (let j = i + 1; j < l.nodes.length; j++) {
      const x = l.nodes[i]!,
        y = l.nodes[j]!;
      if (x.x < y.x + y.w && y.x < x.x + x.w && x.y < y.y + y.h && y.y < x.y + x.h) return true;
    }
  }
  return false;
};

describe("approxTextWidth / wrapLabel", () => {
  it("treats a Chinese glyph as roughly one em", () => {
    expect(approxTextWidth("字")).toBeGreaterThan(approxTextWidth("i"));
  });

  it("wraps long labels into several lines", () => {
    const lines = wrapLabel("这是一个非常非常长的流程节点说明文字", 80);
    expect(lines.length).toBeGreaterThan(1);
    expect(lines.join("")).toBe("这是一个非常非常长的流程节点说明文字");
  });
});

describe("assignLayers", () => {
  it("orders nodes along the flow", () => {
    const layers = assignLayers(flow);
    const at = (id: string) => layers[flow.nodes.findIndex((n) => n.id === id)]!;
    expect(at("a")).toBeLessThan(at("b"));
    expect(at("b")).toBeLessThan(at("c"));
    expect(at("c")).toBeLessThan(at("d"));
  });

  it("still terminates on a cyclic graph", () => {
    const loop: DiagramSpec = {
      layout: "flow-v",
      nodes: [
        { id: "x", label: "X", shape: "box" },
        { id: "y", label: "Y", shape: "box" },
      ],
      edges: [
        { from: "x", to: "y" },
        { from: "y", to: "x" },
      ],
    };
    expect(assignLayers(loop)).toHaveLength(2);
  });
});

describe("layoutDiagram", () => {
  it("keeps every box inside the canvas and off each other", () => {
    const l = layoutDiagram(flow);
    expect(l.width).toBeGreaterThan(0);
    expect(l.height).toBeGreaterThan(0);
    for (const n of l.nodes) {
      expect(n.x).toBeGreaterThanOrEqual(0);
      expect(n.y).toBeGreaterThanOrEqual(0);
      expect(n.x + n.w).toBeLessThanOrEqual(l.width);
      expect(n.y + n.h).toBeLessThanOrEqual(l.height);
      expect(n.w).toBeGreaterThan(0);
      expect(n.h).toBeGreaterThan(0);
    }
    expect(overlaps(l)).toBe(false);
  });

  it("stacks a vertical flow downwards", () => {
    const l = layoutDiagram(flow);
    const y = (id: string) => l.nodes.find((n) => n.id === id)!.y;
    expect(y("a")).toBeLessThan(y("b"));
    expect(y("b")).toBeLessThan(y("c"));
  });

  it("lays a horizontal flow out sideways", () => {
    const l = layoutDiagram({ ...flow, layout: "flow-h" });
    const x = (id: string) => l.nodes.find((n) => n.id === id)!.x;
    expect(x("a")).toBeLessThan(x("b"));
    expect(overlaps(l)).toBe(false);
  });

  it("routes each edge and draws a title band", () => {
    const l = layoutDiagram(flow);
    expect(l.edges).toHaveLength(flow.edges.length);
    for (const e of l.edges) {
      expect(e.points.length).toBeGreaterThanOrEqual(2);
    }
    expect(l.title).toBe("审批流程");
    expect(l.titleAt?.y).toBeLessThan(Math.min(...l.nodes.map((n) => n.y)));
  });

  it("handles a single node without edges", () => {
    const l = layoutDiagram({
      layout: "grid",
      nodes: [{ id: "only", label: "只有一个节点", shape: "box" }],
      edges: [],
    });
    expect(l.nodes).toHaveLength(1);
    expect(l.edges).toHaveLength(0);
    expect(l.width).toBeGreaterThan(0);
  });
});
