import { Extension } from "@tiptap/core";
import { Plugin, PluginKey } from "@tiptap/pm/state";
import type { Editor } from "@tiptap/core";

/**
 * Column-resize support for tables.
 *
 * We render invisible drag handles on column borders (see CSS for
 * `.col-resize-handle`); on mousedown this plugin tracks the drag,
 * updates live DOM widths, and on mouseup commits the new widths
 * back into PM by calling `setCellAttribute` with a `colwidth`
 * array — the attribute TipTap's table extension already recognises
 * and serialises to HTML.
 *
 * This is intentionally a pure DOM plugin: it works with any table
 * the editor produces without additional PM state.
 */

declare module "@tiptap/core" {
  interface Commands<ReturnType> {
    tableColumnResize: {
      commitColumnWidths: (widthsByCell: {
        from: number;
        to: number;
        width: number;
      }[]) => ReturnType;
    };
  }
}

export const TableColumnResize = Extension.create({
  name: "tableColumnResize",

  addCommands() {
    return {
      commitColumnWidths:
        (cells: { from: number; to: number; width: number }[]) =>
        ({ tr }: { tr: { doc: unknown; setNodeMarkup: (pos: number, type: null, attrs: Record<string, unknown>) => void; nodeAt: (pos: number) => { attrs?: Record<string, unknown> } | null } }) => {
          for (const c of cells) {
            tr.setNodeMarkup(c.from, null, {
              ...(tr.nodeAt(c.from)?.attrs ?? {}),
              colwidth: [c.width],
            });
          }
          return true;
        },
    } as never;
  },

  addProseMirrorPlugins() {
    const editor = this.editor;
    let handles: HTMLElement[] = [];

    const buildHandles = (tableEl: HTMLElement) => {
      clearHandles();
      const firstRow = tableEl.querySelector("tr");
      if (!firstRow) return;
      const cells = Array.from(firstRow.children) as HTMLElement[];
      let left = 0;
      cells.forEach((cell, i) => {
        const w = cell.getBoundingClientRect().width;
        left += w;
        if (i === cells.length - 1) return;
        const handle = document.createElement("div");
        handle.className = "col-resize-handle";
        handle.dataset.col = String(i + 1);
        handle.style.left = `${left - tableEl.getBoundingClientRect().left}px`;
        tableEl.appendChild(handle);
        handles.push(handle);
      });
    };

    const clearHandles = () => {
      handles.forEach((h) => h.remove());
      handles = [];
    };

    return [
      new Plugin({
        key: new PluginKey("tableColumnResize"),
        view() {
          return {
            update(view: { dom: HTMLElement }) {
              // Recompute handles when the DOM changes (throttled by rAF).
              const table = view.dom.querySelector("table");
              if (table) {
                requestAnimationFrame(() => {
                  if (!table.isConnected) return;
                  buildHandles(table);
                });
              } else {
                clearHandles();
              }
            },
            destroy() {
              clearHandles();
            },
          };
        },
        props: {
          handleDOMEvents: {
            mousedown(view, e) {
              const target = e.target as HTMLElement;
              if (!target.classList.contains("col-resize-handle"))
                return false;
              const table = target.closest("table");
              if (!table) return false;

              const colIndex = Number(target.dataset.col ?? "0");
              const rows = Array.from(table.querySelectorAll("tr"));
              type Cell = { el: HTMLElement; pos?: number; width: number };
              const collectCells = (col: number): Cell[] => {
                const out: Cell[] = [];
                for (const row of rows) {
                  const cell = row.children[col] as HTMLElement | undefined;
                  if (cell) {
                    out.push({
                      el: cell,
                      width: cell.getBoundingClientRect().width,
                    });
                  }
                }
                return out;
              };
              const leftCells = collectCells(colIndex - 1);
              const rightCells = collectCells(colIndex);
              if (!leftCells.length || !rightCells.length) return false;

              e.preventDefault();
              const startX = e.clientX;
              const startLeftW = leftCells[0]!.el.getBoundingClientRect().width;
              const startRightW = rightCells[0]!.el.getBoundingClientRect().width;
              view.dom.classList.add("resizing");

              const onMove = (ev: MouseEvent) => {
                const dx = ev.clientX - startX;
                const newLeft = Math.max(24, Math.round(startLeftW + dx));
                const newRight = Math.max(24, Math.round(startRightW - dx));
                for (const c of leftCells) c.el.style.width = `${newLeft}px`;
                for (const c of rightCells) c.el.style.width = `${newRight}px`;
              };
              const onUp = () => {
                document.removeEventListener("mousemove", onMove);
                document.removeEventListener("mouseup", onUp);
                view.dom.classList.remove("resizing");
                // Map live DOM cells back to PM positions and commit colwidth.
                const updates: { from: number; to: number; width: number }[] =
                  [];
                const probe = (cellEl: HTMLElement, width: number) => {
                  const pos = view.posAtDOM(cellEl, 0);
                  if (pos >= 0) {
                    const node = view.state.doc.nodeAt(pos);
                    if (node && /cell|header/.test(node.type.name)) {
                      updates.push({
                        from: pos,
                        to: pos + node.nodeSize,
                        width,
                      });
                    }
                  }
                };                const leftW = leftCells[0]!.el.getBoundingClientRect().width;
                const rightW = rightCells[0]!.el.getBoundingClientRect().width;
                for (const c of leftCells) probe(c.el, Math.round(leftW));
                for (const c of rightCells) probe(c.el, Math.round(rightW));
                if (updates.length) {
                  editor.commands.commitColumnWidths(updates);
                }
              };
              document.addEventListener("mousemove", onMove);
              document.addEventListener("mouseup", onUp);
              return true;
            },
          },
        },
      }),
    ];
  },
});
