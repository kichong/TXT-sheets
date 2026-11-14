import { BrowserWindow } from 'electron';
import type { WorkbookPrintRequest, WorkbookPrintResult } from '../shared/types';
import { buildWorkbookPrintHtml } from './workbook-print';

async function createPrintWindow(request: WorkbookPrintRequest, owner: BrowserWindow): Promise<BrowserWindow> {
  const html = buildWorkbookPrintHtml(request);
  const window = new BrowserWindow({ show: false, parent: owner, backgroundColor: '#ffffff', webPreferences: { nodeIntegration: false, contextIsolation: true, sandbox: true } });
  try {
    await window.loadURL(`data:text/html;base64,${Buffer.from(html, 'utf8').toString('base64')}`);
    await window.webContents.executeJavaScript('document.fonts.ready.then(() => true)');
    return window;
  } catch (error) { window.destroy(); throw error; }
}

export async function renderWorkbookPdf(request: WorkbookPrintRequest, owner: BrowserWindow): Promise<Uint8Array> {
  const window = await createPrintWindow(request, owner);
  try {
    return await window.webContents.printToPDF({ pageSize: 'Letter', landscape: request.landscape, printBackground: true, displayHeaderFooter: false, preferCSSPageSize: true, generateTaggedPDF: true });
  } finally { window.destroy(); }
}

export async function printWorkbook(request: WorkbookPrintRequest, owner: BrowserWindow): Promise<WorkbookPrintResult> {
  const window = await createPrintWindow(request, owner);
  try {
    return await new Promise<WorkbookPrintResult>((resolve, reject) => {
      window.webContents.print({ silent: false, printBackground: true, landscape: request.landscape, pageSize: 'Letter', margins: { marginType: 'custom', top: 38, bottom: 38, left: 38, right: 38 } }, (success, error) => {
        if (success) resolve({ status: 'printed' });
        else if (/cancel/iu.test(error)) resolve({ status: 'canceled' });
        else reject(new Error(error || 'Printing failed.'));
      });
    });
  } finally { window.destroy(); }
}
