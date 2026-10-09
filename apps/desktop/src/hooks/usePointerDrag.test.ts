import { describe, expect, it } from "vitest";
import { startPointerDrag } from "./usePointerDrag.js";
import { nextWidth } from "./usePanelResize.js";

/** Minimal stand-in: only the listener + pointer-capture surface is modelled. */
class FakeEl {
  private listeners = new Map<string, ((ev: PointerEvent) => void)[]>();
  captured: number[] = [];
  released: number[] = [];

  addEventListener(type: string, cb: (ev: PointerEvent) => void) {
    const list = this.listeners.get(type) ?? [];
    list.push(cb);
    this.listeners.set(type, list);
  }

  removeEventListener(type: string, cb: (ev: PointerEvent) => void) {
    this.listeners.set(type, (this.listeners.get(type) ?? []).filter((f) => f !== cb));
  }

  count(type: string) {
    return (this.listeners.get(type) ?? []).length;
  }

  emit(type: string, ev: Partial<PointerEvent>) {
    for (const cb of [...(this.listeners.get(type) ?? [])]) cb(ev as PointerEvent);
  }

  setPointerCapture(id: number) {
    this.captured.push(id);
  }

  hasPointerCapture(id: number) {
    return this.captured.includes(id) && !this.released.includes(id);
  }

  releasePointerCapture(id: number) {
    this.released.push(id);
  }
}

describe("startPointerDrag", () => {
  it("captures the pointer so a release outside the window still lands", () => {
    const el = new FakeEl();
    startPointerDrag(el as unknown as HTMLElement, 1, { move: () => {} });
    expect(el.captured).toEqual([1]);
  });

  it("forwards moves and ends on pointerup", () => {
    const el = new FakeEl();
    const moved: number[] = [];
    let ended = 0;
    startPointerDrag(el as unknown as HTMLElement, 1, {
      move: (ev) => moved.push(ev.clientX ?? -1),
      end: () => (ended += 1),
    });
    el.emit("pointermove", { pointerId: 1, clientX: 120 });
    el.emit("pointerup", { pointerId: 1 });
    expect(moved).toEqual([120]);
    expect(ended).toBe(1);
    expect(el.released).toEqual([1]);
  });

  it("detaches every listener, so hovering afterwards resizes nothing", () => {
    // The reported bug: mouseup never arrived, the handlers stayed bound, and the
    // next pointer move kept changing the width with no button held.
    const el = new FakeEl();
    let moves = 0;
    startPointerDrag(el as unknown as HTMLElement, 1, { move: () => (moves += 1), end: () => {} });
    el.emit("pointerup", { pointerId: 1 });
    el.emit("pointermove", { pointerId: 1, clientX: 400 });
    expect(moves).toBe(0);
    expect(el.count("pointermove")).toBe(0);
    expect(el.count("pointerup")).toBe(0);
    expect(el.count("pointercancel")).toBe(0);
  });

  it("ends on pointercancel too", () => {
    const el = new FakeEl();
    let ended = 0;
    startPointerDrag(el as unknown as HTMLElement, 7, { move: () => {}, end: () => (ended += 1) });
    el.emit("pointercancel", { pointerId: 7 });
    expect(ended).toBe(1);
    expect(el.count("pointermove")).toBe(0);
  });

  it("ignores a second pointer", () => {
    const el = new FakeEl();
    const moved: number[] = [];
    let ended = 0;
    startPointerDrag(el as unknown as HTMLElement, 1, {
      move: (ev) => moved.push(ev.clientX ?? -1),
      end: () => (ended += 1),
    });
    el.emit("pointermove", { pointerId: 2, clientX: 999 });
    el.emit("pointerup", { pointerId: 2 });
    expect(moved).toEqual([]);
    expect(ended).toBe(0);
    el.emit("pointerup", { pointerId: 1 });
    expect(ended).toBe(1);
  });
});

describe("nextWidth", () => {
  it("grows a left panel as the pointer moves right", () => {
    expect(nextWidth("left", 220, 40, 180, 480)).toBe(260);
  });

  it("grows a right panel as the pointer moves left", () => {
    expect(nextWidth("right", 340, -40, 280, 640)).toBe(380);
  });

  it("clamps to the panel bounds", () => {
    expect(nextWidth("left", 220, 5000, 180, 480)).toBe(480);
    expect(nextWidth("left", 220, -5000, 180, 480)).toBe(180);
    expect(nextWidth("right", 340, 5000, 280, 640)).toBe(280);
  });
});
