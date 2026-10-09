import { useEffect, useMemo, useRef, useState } from "react";
import {
  COMMANDS,
  rankCommands,
  type Command,
  type CommandContext,
} from "../commands/registry.js";
import { useUiStore } from "../uiStore.js";

export function CommandPalette({
  open,
  onClose,
  ctx,
}: {
  open: boolean;
  onClose: () => void;
  ctx: CommandContext;
}) {
  const [query, setQuery] = useState("");
  const [selected, setSelected] = useState(0);
  const inputRef = useRef<HTMLInputElement>(null);
  const listRef = useRef<HTMLDivElement>(null);
  const usage = useUiStore((s) => s.commandUsage);
  const bumpCommand = useUiStore((s) => s.bumpCommand);

  const commands = useMemo(
    () => rankCommands(COMMANDS, query.trim(), usage),
    [query, usage],
  );

  useEffect(() => {
    if (open) {
      setQuery("");
      setSelected(0);
      inputRef.current?.focus();
    }
  }, [open]);

  useEffect(() => {
    setSelected((v) => Math.min(v, Math.max(0, commands.length - 1)));
  }, [commands.length]);

  useEffect(() => {
    listRef.current
      ?.querySelector(".palette-item.selected")
      ?.scrollIntoView({ block: "nearest" });
  }, [selected]);

  if (!open) return null;

  const run = (cmd: Command) => {
    bumpCommand(cmd.id);
    onClose();
    cmd.run(ctx);
  };

  const onKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === "Escape") {
      e.preventDefault();
      onClose();
    } else if (e.key === "ArrowDown") {
      e.preventDefault();
      setSelected((v) => (v + 1) % commands.length);
    } else if (e.key === "ArrowUp") {
      e.preventDefault();
      setSelected((v) => (v - 1 + commands.length) % commands.length);
    } else if (e.key === "Enter") {
      e.preventDefault();
      const cmd = commands[selected];
      if (cmd) run(cmd);
    }
  };

  let lastGroup = "";
  return (
    <div
      className="palette-backdrop"
      onMouseDown={onClose}
      onKeyDown={onKeyDown}
    >
      <div
        className="palette"
        role="dialog"
        aria-label="命令面板"
        onMouseDown={(e) => e.stopPropagation()}
      >
        <input
          ref={inputRef}
          className="palette-input"
          value={query}
          onChange={(e) => setQuery(e.target.value)}
          placeholder="搜索命令…（如：保存、主题、AI 润色）"
        />
        <div className="palette-list" ref={listRef}>
          {commands.length === 0 && (
            <div className="palette-empty">没有匹配的命令</div>
          )}
          {commands.map((cmd, i) => {
            const showGroup = cmd.group !== lastGroup;
            lastGroup = cmd.group;
            const Icon = cmd.icon;
            return (
              <div key={cmd.id}>
                {showGroup && <div className="palette-group">{cmd.group}</div>}
                <div
                  className={`palette-item${i === selected ? " selected" : ""}`}
                  onMouseEnter={() => setSelected(i)}
                  onClick={() => run(cmd)}
                >
                  {Icon ? (
                    <Icon size={14} className="palette-item-icon" />
                  ) : (
                    <span className="palette-item-icon" />
                  )}
                  <span className="palette-item-title">{cmd.title}</span>
                  {cmd.shortcut && (
                    <span className="palette-item-shortcut">{cmd.shortcut}</span>
                  )}
                </div>
              </div>
            );
          })}
        </div>
        <div className="palette-footer">
          <span>↑↓ 选择</span>
          <span>Enter 执行</span>
          <span>Esc 关闭</span>
        </div>
      </div>
    </div>
  );
}
