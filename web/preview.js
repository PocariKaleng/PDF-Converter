import { getDocument, GlobalWorkerOptions } from './vendor/pdfjs/pdf.mjs';
GlobalWorkerOptions.workerSrc = '/vendor/pdfjs/pdf.worker.mjs';

const canvas = document.querySelector('#pdf-canvas');
const label = document.querySelector('#page-label');
const message = document.querySelector('#preview-message');
const previous = document.querySelector('#previous-page');
const next = document.querySelector('#next-page');
let pdfDocument = null;
let loadingTask = null;
let renderTask = null;
let pageNumber = 1;
let generation = 0;

export function clearPreview() {
  generation++;
  if (renderTask) renderTask.cancel();
  if (loadingTask) loadingTask.destroy().catch(() => {});
  loadingTask = null;
  pdfDocument = null;
  canvas.width = canvas.height = 0;
  previous.disabled = next.disabled = true;
}

async function renderPage() {
  const current = generation;
  previous.disabled = next.disabled = true;
  try {
    const page = await pdfDocument.getPage(pageNumber);
    if (current !== generation) return;
    const width = Math.min(window.document.querySelector('#pdf-preview').clientWidth - 48, 1200);
    const scale = Math.max(width, 200) / page.getViewport({ scale: 1 }).width;
    const viewport = page.getViewport({ scale: scale * Math.min(window.devicePixelRatio || 1, 2) });
    canvas.width = Math.floor(viewport.width);
    canvas.height = Math.floor(viewport.height);
    canvas.setAttribute('aria-label', `Pratinjau halaman ${pageNumber} dari ${pdfDocument.numPages}. Unduh PDF untuk membaca teks lengkap.`);
    renderTask = page.render({ canvasContext: canvas.getContext('2d'), viewport });
    await renderTask.promise;
    if (current !== generation) return;
    label.textContent = `Halaman ${pageNumber} / ${pdfDocument.numPages}`;
    message.textContent = 'Pratinjau lokal. Unduh PDF untuk memilih teks atau mencetak dokumen.';
    previous.disabled = pageNumber <= 1;
    next.disabled = pageNumber >= pdfDocument.numPages;
  } catch (error) {
    if (error.name !== 'RenderingCancelledException' && current === generation) {
      message.textContent = 'Pratinjau tidak dapat ditampilkan. PDF tetap dapat diunduh.';
    }
  }
}

export async function showPreview(blob) {
  clearPreview();
  const current = generation;
  message.textContent = 'Menyiapkan pratinjau…';
  try {
    const data = new Uint8Array(await blob.arrayBuffer());
    if (current !== generation) return;
    loadingTask = getDocument({ data, useSystemFonts: false, isEvalSupported: false });
    const loaded = await loadingTask.promise;
    if (current !== generation) return;
    pdfDocument = loaded;
    pageNumber = 1;
    await renderPage();
  } catch {
    if (current === generation) message.textContent = 'Pratinjau tidak dapat ditampilkan. PDF tetap dapat diunduh.';
  }
}

previous.addEventListener('click', async () => { if (pdfDocument && pageNumber > 1) { pageNumber--; await renderPage(); } });
next.addEventListener('click', async () => { if (pdfDocument && pageNumber < pdfDocument.numPages) { pageNumber++; await renderPage(); } });
