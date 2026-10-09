import { describe, it, expect, beforeEach } from "vitest";
import { wordClient } from "./index.js";

/**
 * Minimal in-memory mock of the Word JavaScript API surface we use. This is
 * not a faithful reimplementation — it just lets us assert that the wrapper
 * wires calls through correctly. The real Word host is exercised during
 * end-to-end sideloading.
 */

interface MockParagraph {
  text: string;
  style?: string;
}

interface MockDoc {
  paragraphs: MockParagraph[];
  selectionText: string;
  ooxml: string;
}

function installWordMock(doc: MockDoc) {
  class SyncTracker {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    load(): any {
      return this;
    }
  }

  class FakeRange extends SyncTracker {
    text: string;
    constructor(
      private readonly owner: { doc: MockDoc },
      initialText?: string,
    ) {
      super();
      this.text = initialText ?? owner.doc.paragraphs.map((p) => p.text).join("\n");
    }
    search(find: string) {
      const ranges: FakeRange[] = [];
      this.owner.doc.paragraphs.forEach((p) => {
        let from = 0;
        while (true) {
          const idx = p.text.indexOf(find, from);
          if (idx < 0) break;
          const r = new FakeRange(this.owner, find);
          ranges.push(r);
          from = idx + find.length;
        }
      });
      const collection = Object.assign(ranges, {
        load: () => collection,
        items: ranges,
      });
      return collection;
    }
    insertText(text: string, loc: string) {
      if (loc === "Replace") {
        // First-match replacement across paragraphs, repeated per item.
        for (const p of this.owner.doc.paragraphs) {
          const idx = p.text.indexOf(this.text);
          if (idx >= 0) {
            p.text = p.text.slice(0, idx) + text + p.text.slice(idx + this.text.length);
            break;
          }
        }
        this.text = text;
      }
      return this;
    }
    insertParagraph(text: string, loc: string) {
      const l = loc.toLowerCase();
      if (l === "end" || l === "after") this.owner.doc.paragraphs.push({ text });
      else if (l === "start" || l === "before") this.owner.doc.paragraphs.unshift({ text });
      return new FakeParagraph(this.owner, text);
    }
    insertTable(rows: number, _cols: number, _loc: string, values: string[][]) {
      void rows;
      values.forEach((row) => this.owner.doc.paragraphs.push({ text: row.join(" | ") }));
      return new FakeTable();
    }
    insertInlinePictureFromBase64(_b64: string, _loc: string) {
      this.owner.doc.paragraphs.push({ text: "[IMAGE]" });
      return { width: 0, height: 0 };
    }
    getOoxml() {
      return { value: this.owner.doc.ooxml };
    }
    insertOoxml(ooxml: string, _loc: string) {
      this.owner.doc.ooxml = ooxml;
      return new FakeRange(this.owner);
    }
  }

  class FakeParagraph extends SyncTracker {
    text: string;
    style?: string;
    constructor(owner: { doc: MockDoc }, text: string) {
      super();
      this.text = text;
      void owner;
    }
  }

  class FakeTable extends SyncTracker {
    rows: { getFirst: () => { font: { bold: boolean }; load: () => unknown } };
    constructor() {
      super();
      const firstRow = { font: { bold: false }, load: () => firstRow };
      this.rows = { getFirst: () => firstRow };
    }
  }

  class FakeBody extends SyncTracker {
    paragraphs: { items: FakeParagraph[]; load: () => unknown };
    constructor(private readonly owner: { doc: MockDoc }) {
      super();
      this.paragraphs = {
        items: [],
        load: () => {
          this.paragraphs.items = owner.doc.paragraphs.map((p) => {
            const fp = new FakeParagraph(owner, p.text);
            fp.style = p.style;
            return fp;
          });
          return this.paragraphs;
        },
      };
    }
    getRange(loc?: string) {
      void loc;
      return new FakeRange(this.owner);
    }
    insertParagraph(text: string, loc: string) {
      return new FakeRange(this.owner).insertParagraph(text, loc);
    }
    insertTable(rows: number, cols: number, loc: string, values?: string[][]) {
      const range = new FakeRange(this.owner);
      return range.insertTable(rows, cols, loc, values ?? []);
    }
    insertInlinePictureFromBase64(b64: string, loc: string) {
      return new FakeRange(this.owner).insertInlinePictureFromBase64(b64, loc);
    }
  }

  class FakeDocument {
    body: FakeBody;
    constructor(private readonly owner: { doc: MockDoc }) {
      this.body = new FakeBody(owner);
    }
    getSelection() {
      return new FakeRange(this.owner, this.owner.doc.selectionText);
    }
  }

  class FakeRequestContext {
    document: FakeDocument;
    constructor(owner: { doc: MockDoc }) {
      this.document = new FakeDocument(owner);
    }
    sync() {
      return Promise.resolve();
    }
  }

  const owner = { doc };
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  (globalThis as any).Word = {
    InsertLocation: { before: "before", after: "after", start: "start", end: "end", replace: "replace" },
    BorderLocation: { top: "top" },
    run: async <T>(fn: (ctx: FakeRequestContext) => Promise<T>): Promise<T> => {
      const ctx = new FakeRequestContext(owner);
      return await fn(ctx);
    },
  };
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  (globalThis as any).Office = {
    context: { host: "Word" },
    HostType: { Word: "Word" },
    onReady: async () => ({ host: "Word" as const }),
  };
}

describe("word-client", () => {
  let doc: MockDoc;

  beforeEach(() => {
    doc = {
      paragraphs: [
        { text: "甲方应在约定期限内交付产品。", style: "Normal" },
        { text: "甲方对产品质量负责。" },
        { text: "本合同一式两份。" },
      ],
      selectionText: "",
      ooxml: "<document>snapshot-v1</document>",
    };
    installWordMock(doc);
  });

  it("lists paragraphs", async () => {
    const paras = await wordClient.listParagraphs();
    expect(paras).toHaveLength(3);
    expect(paras[0]!.text).toContain("甲方");
    expect(paras[0]!.style).toBe("Normal");
  });

  it("replaces literal text across the body", async () => {
    // Replace both "甲方" occurrences one per returned match (simulates loop).
    const result = await wordClient.replaceText({ find: "甲方", replace: "客户" });
    expect(result.count).toBe(2);
    expect(doc.paragraphs[0]!.text).toContain("客户");
    expect(doc.paragraphs[1]!.text).toContain("客户");
  });

  it("returns zero matches without throwing", async () => {
    const result = await wordClient.replaceText({ find: "不存在的词", replace: "x" });
    expect(result.count).toBe(0);
  });

  it("rejects regex replacement for now", async () => {
    await expect(
      wordClient.replaceText({ find: "甲.", replace: "客户", regex: true }),
    ).rejects.toThrow(/regex/);
  });

  it("inserts a paragraph at the end", async () => {
    await wordClient.insertParagraph({ text: "附加条款", position: "end" });
    expect(doc.paragraphs.at(-1)!.text).toBe("附加条款");
  });

  it("inserts a table (flattened in mock) and bolds header", async () => {
    await wordClient.insertTable({ rows: [["A", "B"], ["1", "2"]], headerRow: true });
    expect(doc.paragraphs.some((p) => p.text.includes("A | B"))).toBe(true);
  });

  it("captures and restores snapshots", async () => {
    const snap = await wordClient.getSnapshot();
    expect(snap.ooxml).toContain("snapshot-v1");
    doc.ooxml = "<document>changed</document>";
    await wordClient.restoreSnapshot(snap);
    expect(doc.ooxml).toBe("<document>snapshot-v1</document>");
  });

  it("applyEdits runs a mixed batch and returns a snapshot", async () => {
    const snap = await wordClient.applyEdits([
      { type: "replace_text", find: "甲方", replace: "客户", scope: "body", previews: [] },
      { type: "insert_paragraph", text: "结尾段落", position: "end" },
    ]);
    expect(snap.ooxml).toContain("snapshot-v1");
    expect(doc.paragraphs[0]!.text).toContain("客户");
    expect(doc.paragraphs.at(-1)!.text).toBe("结尾段落");
  });
});
