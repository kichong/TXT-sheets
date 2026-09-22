import { useLayoutEffect, useRef, useState } from 'react';

export interface SheetMenuTarget {
  id: string;
  x: number;
  y: number;
  trigger: HTMLButtonElement;
  rename?: boolean;
}

export function SheetTabMenu({ target, name, canMoveLeft, canMoveRight, canDelete, onRename, onMove, onAdd, onDelete, onClose }: {
  target: SheetMenuTarget;
  name: string;
  canMoveLeft: boolean;
  canMoveRight: boolean;
  canDelete: boolean;
  onRename: (name: string) => string | null;
  onMove: (direction: -1 | 1) => void;
  onAdd: () => void;
  onDelete: () => void;
  onClose: () => void;
}) {
  const [renaming, setRenaming] = useState(Boolean(target.rename));
  const [value, setValue] = useState(name);
  const [error, setError] = useState<string | null>(null);
  const menuRef = useRef<HTMLDivElement>(null);
  const dialogRef = useRef<HTMLDialogElement>(null);
  const inputRef = useRef<HTMLInputElement>(null);
  const close = () => { onClose(); target.trigger.focus(); };

  useLayoutEffect(() => {
    if (renaming) {
      dialogRef.current?.showModal();
      inputRef.current?.select();
      return;
    }
    const menu = menuRef.current!;
    menu.style.left = `${Math.max(8, Math.min(target.x, window.innerWidth - menu.offsetWidth - 8))}px`;
    menu.style.top = `${Math.max(8, Math.min(target.y, window.innerHeight - menu.offsetHeight - 8))}px`;
    menu.querySelector<HTMLButtonElement>('button:not(:disabled)')?.focus();
  }, [renaming, target]);

  if (renaming) return <dialog ref={dialogRef} className="sheet-rename-dialog" aria-labelledby="sheet-rename-title" onCancel={(event) => { event.preventDefault(); close(); }}>
    <form onSubmit={(event) => {
      event.preventDefault();
      const problem = onRename(value.trim());
      if (problem) setError(problem);
      else close();
    }}>
      <h2 id="sheet-rename-title">Rename sheet</h2>
      <label htmlFor="sheet-name">Sheet name</label>
      <input ref={inputRef} id="sheet-name" value={value} onChange={(event) => { setValue(event.target.value); setError(null); }} aria-invalid={Boolean(error)} aria-describedby={error ? 'sheet-name-error' : undefined} />
      {error && <p id="sheet-name-error" role="alert">{error}</p>}
      <div className="sheet-rename-actions"><button type="button" onClick={close}>Cancel</button><button type="submit">Rename</button></div>
    </form>
  </dialog>;

  const run = (action: () => void) => { close(); action(); };
  return <div className="sheet-menu-backdrop" onPointerDown={close} onContextMenu={(event) => { event.preventDefault(); close(); }}>
    <div ref={menuRef} className="sheet-tab-menu" role="menu" aria-label={`Sheet options for ${name}`} onPointerDown={(event) => event.stopPropagation()} onKeyDown={(event) => {
      const buttons = [...menuRef.current!.querySelectorAll<HTMLButtonElement>('button:not(:disabled)')];
      const index = buttons.indexOf(document.activeElement as HTMLButtonElement);
      if (event.key === 'Escape' || event.key === 'Tab') { event.preventDefault(); close(); }
      else if (['ArrowDown', 'ArrowUp', 'Home', 'End'].includes(event.key)) {
        event.preventDefault();
        const next = event.key === 'Home' ? 0 : event.key === 'End' ? buttons.length - 1 : (index + (event.key === 'ArrowDown' ? 1 : -1) + buttons.length) % buttons.length;
        buttons[next]?.focus();
      }
    }}>
      <button role="menuitem" onClick={() => setRenaming(true)}>Rename…</button>
      <button role="menuitem" disabled={!canMoveLeft} onClick={() => run(() => onMove(-1))}>Move left</button>
      <button role="menuitem" disabled={!canMoveRight} onClick={() => run(() => onMove(1))}>Move right</button>
      <button role="menuitem" onClick={() => run(onAdd)}>Add sheet</button>
      <div role="separator" />
      <button role="menuitem" disabled={!canDelete} onClick={() => run(onDelete)}>Delete sheet…</button>
    </div>
  </div>;
}
