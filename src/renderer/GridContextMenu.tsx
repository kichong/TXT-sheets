import { useLayoutEffect, useRef } from 'react';
import type { Selection } from './workbook-model';

export interface GridMenuTarget { x: number; y: number; selection: Selection }
export interface GridMenuItem { label: string; shortcut?: string; disabled?: boolean; action: () => void }

export function GridContextMenu({ target, items, onClose }: {
  target: GridMenuTarget; items: Array<GridMenuItem | null>; onClose: () => void;
}) {
  const menuRef = useRef<HTMLDivElement>(null);
  useLayoutEffect(() => {
    const menu = menuRef.current!;
    menu.style.left = `${Math.max(8, Math.min(target.x, window.innerWidth - menu.offsetWidth - 8))}px`;
    menu.style.top = `${Math.max(8, Math.min(target.y, window.innerHeight - menu.offsetHeight - 8))}px`;
    menu.querySelector<HTMLButtonElement>('button:not(:disabled)')?.focus();
  }, [target]);
  return <div className="sheet-menu-backdrop" onPointerDown={onClose} onContextMenu={(event) => { event.preventDefault(); onClose(); }}>
    <div ref={menuRef} className="sheet-tab-menu grid-context-menu" role="menu" aria-label="Selection actions"
      onPointerDown={(event) => event.stopPropagation()} onKeyDown={(event) => {
        const buttons = [...menuRef.current!.querySelectorAll<HTMLButtonElement>('button:not(:disabled)')];
        const index = buttons.indexOf(document.activeElement as HTMLButtonElement);
        if (event.key === 'Escape' || event.key === 'Tab') { event.preventDefault(); onClose(); }
        else if (['ArrowDown', 'ArrowUp', 'Home', 'End'].includes(event.key)) {
          event.preventDefault();
          const next = event.key === 'Home' ? 0 : event.key === 'End' ? buttons.length - 1 : (index + (event.key === 'ArrowDown' ? 1 : -1) + buttons.length) % buttons.length;
          buttons[next]?.focus();
        }
      }}>
      {items.map((item, index) => item ? <button role="menuitem" aria-label={item.label} aria-keyshortcuts={item.shortcut?.replace('Ctrl', 'Control')} key={item.label} disabled={item.disabled} onClick={() => { onClose(); item.action(); }}>
        <span>{item.label}</span>{item.shortcut && <small aria-hidden="true">{item.shortcut}</small>}
      </button> : <div key={index} role="separator" />)}
    </div>
  </div>;
}
