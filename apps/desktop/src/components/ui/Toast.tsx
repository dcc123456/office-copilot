import { useEffect, useState } from "react";
import { CheckCircle2, Info, XCircle } from "lucide-react";

export type ToastKind = "info" | "success" | "error";

interface ToastItem {
  id: number;
  kind: ToastKind;
  text: string;
}

type Listener = (items: ToastItem[]) => void;

const listeners = new Set<Listener>();
let items: ToastItem[] = [];
let nextId = 1;

function publish() {
  listeners.forEach((l) => l(items));
}

export function toast(text: string, kind: ToastKind = "info") {
  const id = nextId++;
  items = [...items, { id, kind, text }].slice(-4);
  publish();
  window.setTimeout(() => {
    items = items.filter((t) => t.id !== id);
    publish();
  }, 6000);
}

export function dismissToast(id: number) {
  items = items.filter((t) => t.id !== id);
  publish();
}

const ICONS = { info: Info, success: CheckCircle2, error: XCircle } as const;

export function ToastHost() {
  const [list, setList] = useState<ToastItem[]>(items);
  useEffect(() => {
    listeners.add(setList);
    setList(items);
    return () => {
      listeners.delete(setList);
    };
  }, []);

  if (list.length === 0) return null;

  return (
    <div className="toast-stack" aria-live="polite">
      {list.map((t) => {
        const Icon = ICONS[t.kind];
        return (
          <div key={t.id} className={`toast toast-${t.kind}`} onClick={() => dismissToast(t.id)}>
            <Icon size={14} className="toast-icon" />
            <span className="toast-text">{t.text}</span>
          </div>
        );
      })}
    </div>
  );
}
