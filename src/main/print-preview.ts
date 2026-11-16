import { BrowserWindow, ipcMain } from 'electron';
import { mkdtemp, writeFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

/** Preview the actual PDF pages; keep the print surface alive for native printing. */
export async function showPrintPreview(source: BrowserWindow, pdf: Uint8Array, owner: BrowserWindow | null, print: () => Promise<void>): Promise<void> {
  const directory = await mkdtemp(join(tmpdir(), 'txt-print-preview-'));
  let preview: BrowserWindow | undefined;
  let actionHandler: ((event: Electron.IpcMainEvent, action: string) => void) | undefined;
  try {
    await writeFile(join(directory, 'preview.pdf'), pdf);
    await writeFile(join(directory, 'index.html'), `<!doctype html><html><head><meta charset="utf-8">
      <meta http-equiv="Content-Security-Policy" content="default-src 'none'; frame-src 'self'; style-src 'unsafe-inline'">
      <title>Print preview</title><style>
      *{box-sizing:border-box}body{margin:0;overflow:hidden;font:14px system-ui;background:#e8e8e8;color:#222}
      header{height:56px;display:flex;align-items:center;gap:12px;padding:0 20px;background:#fafafa;border-bottom:1px solid #ccc}
      strong{flex:1}button{color:inherit;text-decoration:none;padding:7px 18px;border:1px solid #aaa;border-radius:5px;background:white}
      button:focus-visible{outline:3px solid #2563eb;outline-offset:2px}iframe{display:block;border:0;width:100%;height:calc(100vh - 56px)}
      @media(prefers-color-scheme:dark){body{background:#252525;color:#eee}header{background:#202020;border-color:#555}button{background:#333;border-color:#777}}
      </style></head><body><header><strong>Print preview</strong><button data-action="print">Print…</button><button data-action="close">Close</button></header>
      <iframe title="Printed pages" src="preview.pdf"></iframe></body></html>`);
    await writeFile(join(directory, 'preload.cjs'), `const { ipcRenderer } = require('electron');
      window.addEventListener('DOMContentLoaded', () => {
        document.addEventListener('click', event => { const button = event.target.closest('button'); if (button) ipcRenderer.send('txt-preview-action', button.dataset.action); });
        document.addEventListener('keydown', event => { if (event.key === 'Escape') ipcRenderer.send('txt-preview-action', 'close'); });
      });`);
    preview = new BrowserWindow({ title: 'Print preview', width: 1050, height: 850, parent: owner ?? undefined, autoHideMenuBar: true, show: false,
      webPreferences: { nodeIntegration: false, contextIsolation: true, sandbox: true, plugins: true, preload: join(directory, 'preload.cjs') } });
    preview.webContents.setWindowOpenHandler(() => ({ action: 'deny' }));
    await new Promise<void>((resolve, reject) => {
      let printing = false;
      preview!.once('closed', resolve);
      preview!.webContents.on('will-navigate', event => event.preventDefault());
      actionHandler = (event, action) => {
        if (!preview || preview.isDestroyed() || event.sender !== preview.webContents) return;
        if (action === 'close' && !printing) preview.close();
        if (action === 'print' && !printing && !source.isDestroyed()) {
          printing = true;
          void print().catch(reject).finally(() => { printing = false; });
        }
      };
      ipcMain.on('txt-preview-action', actionHandler);
      preview!.loadFile(join(directory, 'index.html')).then(() => preview?.show(), reject);
    });
  } finally {
    if (actionHandler) ipcMain.removeListener('txt-preview-action', actionHandler);
    if (preview && !preview.isDestroyed()) preview.destroy();
    await rm(directory, { recursive: true, force: true });
  }
}
