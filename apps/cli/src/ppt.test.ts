import { describe, expect, it } from "vitest";
import { describePptIssue } from "./ppt.js";

describe("oc ppt 的装配面", () => {
  it("结构问题码：念出来的是一句能照着改的中文，detail 拼在冒号后", () => {
    expect(
      describePptIssue({ code: "slide-missing-shape-tree", part: "ppt/slides/slide1.xml", detail: "" }),
    ).toBe("幻灯片缺 <p:cSld> 里的 <p:spTree>，这是必选元素，打开时报损坏");
    expect(
      describePptIssue({
        code: "relationship-target-missing",
        part: "ppt/slides/slide1.xml",
        detail: "rId1 -> ppt/slideLayouts/slideLayout9.xml",
      }),
    ).toBe("关系指向的部件不在包里，顺着这条引用会断：rId1 -> ppt/slideLayouts/slideLayout9.xml");
    expect(
      describePptIssue({ code: "slide-incomplete-transform", part: "ppt/slides/slide1.xml", detail: "a:off 缺 x" }),
    ).toBe("几何元素只写了一半必填属性（x 与 y、cx 与 cy 得成对），打开时报损坏：a:off 缺 x");
  });
});
