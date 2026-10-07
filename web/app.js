'use strict';
import { showPreview, clearPreview } from './preview.js';
const $ = (selector) => document.querySelector(selector);
const input = $('#file-input');
const dropzone = $('#dropzone');
const fontSelect = $('#font-select');
const button = $('#convert-button');
const message = $('#message');
let selectedFile = null;
let service = null;
let busy = false;
let pdfUrl = null;

function announce(text, kind = 'error') {
  message.textContent = text;
  message.className = `message ${kind}`;
  message.hidden = !text;
}

function updateButton() {
  const available = fontSelect.value === 'original' || service?.fonts.some((font) => font.name === fontSelect.value && font.available);
  button.disabled = busy || !selectedFile || !service?.ready || !available;
}

function clearResult() {
  $('#result').hidden = true;
  clearPreview();
  $('#download-link').removeAttribute('href');
  if (pdfUrl) URL.revokeObjectURL(pdfUrl);
  pdfUrl = null;
}

function chooseFile(file) {
  if (busy) return;
  clearResult();
  selectedFile = null;
  $('#selected-file').hidden = true;
  if (file && (!file.name.toLowerCase().endsWith('.docx') || !file.size || file.size > 30 * 1024 * 1024)) {
    announce('Pilih dokumen .docx dengan ukuran maksimal 30 MB.');
    input.value = '';
  } else {
    selectedFile = file || null;
    announce('');
    if (file) {
      $('#selected-file').hidden = false;
      $('#file-name').textContent = file.name;
      $('#file-size').textContent = `${(file.size / 1024).toLocaleString('id-ID', { maximumFractionDigits: 1 })} KB · Siap dikonversi`;
    }
  }
  updateButton();
}

$('#pick-file').addEventListener('click', () => input.click());
input.addEventListener('change', () => chooseFile(input.files[0]));
$('#remove-file').addEventListener('click', () => { input.value = ''; chooseFile(null); $('#pick-file').focus(); });
for (const event of ['dragenter', 'dragover']) {
  dropzone.addEventListener(event, (e) => { e.preventDefault(); if (!busy) dropzone.classList.add('dragover'); });
}
for (const event of ['dragleave', 'drop']) {
  dropzone.addEventListener(event, (e) => { e.preventDefault(); dropzone.classList.remove('dragover'); });
}
dropzone.addEventListener('drop', (e) => {
  if (e.dataTransfer.files.length !== 1) return announce('Pilih satu dokumen untuk setiap konversi.');
  input.value = '';
  chooseFile(e.dataTransfer.files[0]);
});
document.addEventListener('dragover', (e) => e.preventDefault());
document.addEventListener('drop', (e) => e.preventDefault());
fontSelect.addEventListener('change', () => {
  clearResult();
  $('#font-hint').textContent = fontSelect.value === 'original'
    ? 'Font dan ukuran mengikuti dokumen. Bold, italic, dan bold italic tetap dipertahankan.'
    : `Teks memakai ${fontSelect.value}. Ukuran, bold, italic, dan bold italic mengikuti dokumen; perubahan font dapat menggeser tata letak.`;
  updateButton();
});

async function loadStatus() {
  try {
    const response = await fetch('/api/status');
    if (!response.ok) throw new Error('Status converter tidak dapat dimuat. Muat ulang halaman.');
    service = await response.json();
    for (const font of service.fonts) {
      const prefix = font.name === 'Bell MT' ? 'bell' : 'lm';
      const status = $(`#${prefix}-status`);
      status.textContent = font.available ? '✓ Tersedia' : 'Belum tersedia';
      status.className = `font-status ${font.available ? 'ready' : 'missing'}`;
      const names = { regular: 'Regular', bold: 'Bold', italic: 'Italic', bolditalic: 'Bold italic' };
      $(`#${prefix}-variants`).textContent = font.variants.map((v) => names[v]).join(' · ') || 'Tambahkan font, lalu restart aplikasi';
      const option = [...fontSelect.options].find((o) => o.value === font.name);
      option.disabled = !font.available;
    }
    $('#engine-label').textContent = service.ready ? 'Mesin Microsoft Word tersedia' : 'Microsoft Word belum tersedia';
    if (!service.ready) announce('Pasang dan aktifkan Microsoft Word desktop di Windows untuk menjalankan konversi.');
    updateButton();
  } catch (error) {
    announce(error.message);
    $('#engine-label').textContent = 'Koneksi converter terputus';
  }
}

$('#convert-form').addEventListener('submit', async (event) => {
  event.preventDefault();
  if (busy || button.disabled) return;
  busy = true;
  clearResult();
  updateButton();
  for (const element of [fontSelect, $('#pick-file'), $('#remove-file')]) element.disabled = true;
  button.setAttribute('aria-busy', 'true');
  $('#convert-label').textContent = 'Mengonversi dokumen…';
  $('#convert-arrow').hidden = true;
  $('#spinner').hidden = false;
  announce('Word sedang menyiapkan PDF dan memeriksa font. Dokumen besar dapat memerlukan hingga 2 menit.', 'processing');
  try {
    const response = await fetch('/api/convert', {
      method: 'POST',
      headers: { 'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        'X-Converter-Token': service.token, 'X-Filename': encodeURIComponent(selectedFile.name), 'X-Font': encodeURIComponent(fontSelect.value) },
      body: selectedFile,
    });
    if (!response.ok) {
      const error = await response.json();
      throw new Error(error.error || 'Konversi gagal. Coba lagi.');
    }
    const report = JSON.parse(decodeURIComponent(response.headers.get('X-Conversion-Report')));
    const blob = await response.blob();
    pdfUrl = URL.createObjectURL(blob);
    $('#download-link').href = response.headers.get('X-Download-Url') || pdfUrl;
    $('#download-link').download = selectedFile.name.replace(/\.docx$/i, '.pdf');
    const verified = report.verified_families.length ? `${report.verified_families.join(' dan ')} terverifikasi tertanam.` : 'Lihat detail font pada hasil PDF.';
    $('#result-summary').textContent = `${report.pages} halaman · ${(blob.size / 1024).toLocaleString('id-ID', { maximumFractionDigits: 1 })} KB · ${verified}`;
    const list = $('#font-report-list');
    list.replaceChildren();
    for (const font of report.fonts) {
      const item = document.createElement('li');
      const name = document.createElement('span');
      const status = document.createElement('span');
      name.textContent = font.name + (font.version ? ` · v${font.version}` : '');
      status.textContent = font.embedded ? '✓ Tertanam' : 'Tidak tertanam';
      if (!font.embedded) status.className = 'unembedded';
      item.append(name, status);
      list.append(item);
    }
    $('#result').hidden = false;
    announce('');
    $('#result-title').focus();
    await showPreview(blob);
  } catch (error) {
    announce(error.message === 'Failed to fetch' ? 'Koneksi aplikasi terputus. Pastikan server masih berjalan, lalu coba lagi.' : error.message);
  } finally {
    busy = false;
    for (const element of [fontSelect, $('#pick-file'), $('#remove-file')]) element.disabled = false;
    button.removeAttribute('aria-busy');
    $('#convert-label').textContent = 'Konversi ke PDF';
    $('#convert-arrow').hidden = false;
    $('#spinner').hidden = true;
    updateButton();
  }
});
window.addEventListener('beforeunload', () => { if (pdfUrl) URL.revokeObjectURL(pdfUrl); });
loadStatus();
