import { useEffect, useRef, useState } from "react";

type DialogReq =
  | {
      kind: "confirm";
      title: string;
      message: string;
      okLabel: string;
      cancelLabel: string;
      danger: boolean;
      resolve: (v: boolean) => void;
    }
  | {
      kind: "prompt";
      title: string;
      message: string;
      placeholder: string;
      initialValue: string;
      resolve: (v: string | null) => void;
    };

let emit: ((r: DialogReq) => void) | null = null;

export function confirmDialog(opts: {
  title: string;
  message: string;
  okLabel?: string;
  cancelLabel?: string;
  danger?: boolean;
}): Promise<boolean> {
  return new Promise((resolve) => {
    const req: DialogReq = {
      kind: "confirm",
      title: opts.title,
      message: opts.message,
      okLabel: opts.okLabel ?? "确定",
      cancelLabel: opts.cancelLabel ?? "取消",
      danger: opts.danger ?? false,
      resolve,
    };
    if (emit) emit(req);
    else resolve(false);
  });
}

export function promptDialog(opts: {
  title: string;
  message?: string;
  placeholder?: string;
  initialValue?: string;
}): Promise<string | null> {
  return new Promise((resolve) => {
    const req: DialogReq = {
      kind: "prompt",
      title: opts.title,
      message: opts.message ?? "",
      placeholder: opts.placeholder ?? "",
      initialValue: opts.initialValue ?? "",
      resolve,
    };
    if (emit) emit(req);
    else resolve(null);
  });
}

export function DialogHost() {
  const [req, setReq] = useState<DialogReq | null>(null);
  const inputRef = useRef<HTMLInputElement>(null);
  const okRef = useRef<HTMLButtonElement>(null);
  const resolveRef = useRef<DialogReq["resolve"] | null>(null);

  useEffect(() => {
    emit = setReq;
    return () => {
      emit = null;
    };
  }, []);

  const settle = (value: boolean | string | null) => {
    resolveRef.current?.(value as never);
    resolveRef.current = null;
    setReq(null);
  };

  useEffect(() => {
    if (!req) return;
    resolveRef.current = req.resolve;
    const t = window.setTimeout(() => {
      if (req.kind === "prompt") inputRef.current?.focus();
      else okRef.current?.focus();
    }, 30);
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Escape") {
        e.stopPropagation();
        settle(req.kind === "confirm" ? false : null);
      } else if (e.key === "Enter") {
        e.preventDefault();
        if (req.kind === "prompt") settle(inputRef.current?.value ?? null);
        else settle(true);
      }
    };
    // Capture so inner inputs' Enter doesn't double-fire elsewhere.
    window.addEventListener("keydown", onKey, true);
    return () => {
      window.clearTimeout(t);
      window.removeEventListener("keydown", onKey, true);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [req]);

  if (!req) return null;

  return (
    <div className="modal-backdrop dialog-backdrop" onClick={() => settle(req.kind === "confirm" ? false : null)}>
      <div
        className="modal dialog"
        role="dialog"
        aria-modal="true"
        aria-label={req.title}
        onClick={(e) => e.stopPropagation()}
      >
        <div className="dialog-title">{req.title}</div>
        {req.kind === "confirm" ? (
          <div className="dialog-message">{req.message}</div>
        ) : (
          <input
            ref={inputRef}
            className="field-input dialog-input"
            placeholder={req.placeholder}
            defaultValue={req.initialValue}
            spellCheck={false}
          />
        )}
        <div className="dialog-actions">
          <button onClick={() => settle(req.kind === "confirm" ? false : null)}>
            {req.kind === "confirm" ? req.cancelLabel : "取消"}
          </button>
          <button
            ref={okRef}
            className={"primary" + (req.kind === "confirm" && req.danger ? " danger" : "")}
            onClick={() =>
              settle(req.kind === "confirm" ? true : inputRef.current?.value ?? null)
            }
          >
            {req.kind === "confirm" ? req.okLabel : "确定"}
          </button>
        </div>
      </div>
    </div>
  );
}
