const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycby6jTm5xQmcQAUjdaubsOWOn7Xws4UWjV9uWbOExlEQArCSN6hubMt3U128QjmlWZP0Ow/exec';
const ROOT_FOLDER_ID = '1NZfDp_9SU50OVDJXLuTcGZNj5JvSdpdX';
const pdfjsLib = globalThis.pdfjsLib;
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.5.136/pdf.worker.min.mjs';

const state = {
  periods: [],
  selectedPeriod: '',
  records: [],
  filteredRecords: [],
  selectedRecord: null,
  pdfDoc: null,
  pdfBytes: null,
  previewBytes: null,
  pagesMeta: [],
  zoom: 1.35,
  overlays: [],
  selectedOverlayId: '',
  autosaveTimer: null,
  isDirty: false,
  isSaving: false,
};

const el = {
  periodSelect: document.getElementById('periodSelect'),
  searchInput: document.getElementById('searchInput'),
  employeeSelect: document.getElementById('employeeSelect'),
  refreshBtn: document.getElementById('refreshBtn'),
  syncBtn: document.getElementById('syncBtn'),
  addAreaBtn: document.getElementById('addAreaBtn'),
  updateAreaBtn: document.getElementById('updateAreaBtn'),
  deleteAreaBtn: document.getElementById('deleteAreaBtn'),
  saveBtn: document.getElementById('saveBtn'),
  printBtn: document.getElementById('printBtn'),
  downloadBtn: document.getElementById('downloadBtn'),
  zoomOutBtn: document.getElementById('zoomOutBtn'),
  zoomInBtn: document.getElementById('zoomInBtn'),
  pdfContainer: document.getElementById('pdfContainer'),
  areaList: document.getElementById('areaList'),
  areaPage: document.getElementById('areaPage'),
  areaFontSize: document.getElementById('areaFontSize'),
  areaX: document.getElementById('areaX'),
  areaY: document.getElementById('areaY'),
  areaWidth: document.getElementById('areaWidth'),
  areaHeight: document.getElementById('areaHeight'),
  areaLabel: document.getElementById('areaLabel'),
  areaText: document.getElementById('areaText'),
  statusText: document.getElementById('statusText'),
  fileNameText: document.getElementById('fileNameText'),
  employeeIdText: document.getElementById('employeeIdText'),
  employeeNameText: document.getElementById('employeeNameText'),
  pageInfo: document.getElementById('pageInfo'),
  areaItemTemplate: document.getElementById('areaItemTemplate'),
};

function setStatus(text) {
  el.statusText.textContent = text;
}

function apiUrl(action, params = {}) {
  const url = new URL(APPS_SCRIPT_URL);
  url.searchParams.set('action', action);
  url.searchParams.set('rootFolderId', ROOT_FOLDER_ID);
  Object.entries(params).forEach(([key, value]) => {
    if (value !== undefined && value !== null && value !== '') url.searchParams.set(key, value);
  });
  return url.toString();
}

async function apiGet(action, params = {}) {
  const res = await fetch(apiUrl(action, params));
  const text = await res.text();
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  const json = JSON.parse(text);
  if (json.ok === false) throw new Error(json.message || 'Request gagal');
  return json;
}

async function apiPost(action, payload = {}) {
  const res = await fetch(apiUrl(action), {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify({ action, rootFolderId: ROOT_FOLDER_ID, ...payload })
  });
  const text = await res.text();
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  const json = JSON.parse(text);
  if (json.ok === false) throw new Error(json.message || 'Request gagal');
  return json;
}

function normalize(value) {
  return String(value || '').toLowerCase().trim();
}

function humanFileLabel(record) {
  return `${record.employeeId || '-'} - ${record.employeeName || '-'} (${record.periodLabel || '-'})`;
}

function parseBase64(base64) {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes;
}

function bytesToBase64(bytes) {
  let binary = '';
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) {
    const sub = bytes.subarray(i, i + chunk);
    binary += String.fromCharCode(...sub);
  }
  return btoa(binary);
}

function renderPeriods() {
  el.periodSelect.innerHTML = state.periods.map(period => `<option value="${period}">${period}</option>`).join('');
  if (state.periods.length && !state.selectedPeriod) state.selectedPeriod = state.periods[0];
  el.periodSelect.value = state.selectedPeriod;
}

function renderResults() {
  const rows = state.filteredRecords;
  el.employeeSelect.innerHTML = rows.map((record, index) => `<option value="${index}">${humanFileLabel(record)}</option>`).join('');
  if (!rows.length) {
    el.employeeSelect.innerHTML = '';
  }
}

function filterRecords() {
  const q = normalize(el.searchInput.value);
  state.filteredRecords = state.records.filter(record => {
    if (state.selectedPeriod && record.periodLabel !== state.selectedPeriod) return false;
    if (!q) return true;
    return normalize(record.employeeId).includes(q) || normalize(record.employeeName).includes(q) || normalize(record.fileName).includes(q);
  });
  renderResults();
}

function syncMetaPanel() {
  const r = state.selectedRecord;
  el.fileNameText.textContent = r?.fileName || '-';
  el.employeeIdText.textContent = r?.employeeId || '-';
  el.employeeNameText.textContent = r?.employeeName || '-';
}

function getSelectedOverlay() {
  return state.overlays.find(item => item.id === state.selectedOverlayId) || null;
}

function syncAreaForm() {
  const area = getSelectedOverlay();
  if (!area) {
    el.areaPage.value = 1;
    el.areaFontSize.value = 12;
    el.areaX.value = 30;
    el.areaY.value = 50;
    el.areaWidth.value = 140;
    el.areaHeight.value = 14;
    el.areaLabel.value = '';
    el.areaText.value = '';
    return;
  }
  el.areaPage.value = area.page;
  el.areaFontSize.value = area.fontSize;
  el.areaX.value = area.x;
  el.areaY.value = area.y;
  el.areaWidth.value = area.width;
  el.areaHeight.value = area.height;
  el.areaLabel.value = area.label;
  el.areaText.value = area.text;
}

function renderAreaList() {
  if (!state.overlays.length) {
    el.areaList.className = 'area-list empty';
    el.areaList.textContent = 'Belum ada area edit.';
    return;
  }
  el.areaList.className = 'area-list';
  el.areaList.innerHTML = '';
  state.overlays.forEach(area => {
    const node = el.areaItemTemplate.content.firstElementChild.cloneNode(true);
    node.dataset.id = area.id;
    node.classList.toggle('active', area.id === state.selectedOverlayId);
    node.querySelector('.area-item-title').textContent = area.label || 'Tanpa Nama';
    node.querySelector('.area-item-subtitle').textContent = `Page ${area.page} • x:${Math.round(area.x)} y:${Math.round(area.y)} • ${area.text || '-'}`;
    node.addEventListener('click', () => {
      state.selectedOverlayId = area.id;
      syncAreaForm();
      renderAreaList();
      renderPdfPages();
    });
    el.areaList.appendChild(node);
  });
}

function areaFromForm() {
  return {
    page: Math.max(1, Number(el.areaPage.value || 1)),
    fontSize: Math.max(6, Number(el.areaFontSize.value || 12)),
    x: Math.max(0, Number(el.areaX.value || 0)),
    y: Math.max(0, Number(el.areaY.value || 0)),
    width: Math.max(10, Number(el.areaWidth.value || 20)),
    height: Math.max(8, Number(el.areaHeight.value || 10)),
    label: el.areaLabel.value.trim(),
    text: el.areaText.value,
  };
}

function markDirty() {
  state.isDirty = true;
  setStatus('Ada perubahan belum tersimpan');
  if (state.autosaveTimer) clearTimeout(state.autosaveTimer);
  state.autosaveTimer = setTimeout(() => savePdf(true), 1800);
}

function createArea() {
  const area = areaFromForm();
  const id = `area_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
  state.overlays.push({ id, ...area });
  state.selectedOverlayId = id;
  renderAreaList();
  renderPdfPages();
  markDirty();
}

function updateArea() {
  const area = getSelectedOverlay();
  if (!area) return;
  Object.assign(area, areaFromForm());
  renderAreaList();
  renderPdfPages();
  markDirty();
}

function deleteArea() {
  if (!state.selectedOverlayId) return;
  state.overlays = state.overlays.filter(item => item.id !== state.selectedOverlayId);
  state.selectedOverlayId = state.overlays[0]?.id || '';
  syncAreaForm();
  renderAreaList();
  renderPdfPages();
  markDirty();
}

function clamp(v, min, max) {
  return Math.max(min, Math.min(max, v));
}

function makeOverlayNode(area, pageMeta) {
  const scale = pageMeta.scale;
  const overlay = document.createElement('div');
  overlay.className = 'overlay-box';
  overlay.dataset.id = area.id;
  overlay.style.left = `${area.x * scale}px`;
  overlay.style.top = `${area.y * scale}px`;
  overlay.style.width = `${area.width * scale}px`;
  overlay.style.height = `${area.height * scale}px`;
  overlay.classList.toggle('active', area.id === state.selectedOverlayId);

  const label = document.createElement('div');
  label.className = 'overlay-label';
  label.textContent = area.label || 'Area';
  overlay.appendChild(label);

  const handle = document.createElement('div');
  handle.className = 'resize-handle';
  overlay.appendChild(handle);

  overlay.addEventListener('mousedown', (event) => {
    event.preventDefault();
    state.selectedOverlayId = area.id;
    syncAreaForm();
    renderAreaList();
    renderPdfPages();

    const rect = pageMeta.overlayNode.getBoundingClientRect();
    const startX = event.clientX;
    const startY = event.clientY;
    const startArea = { ...area };
    const resizing = event.target.classList.contains('resize-handle');

    function onMove(ev) {
      const dx = (ev.clientX - startX) / scale;
      const dy = (ev.clientY - startY) / scale;
      if (resizing) {
        area.width = clamp(startArea.width + dx, 10, pageMeta.widthPdf - area.x);
        area.height = clamp(startArea.height + dy, 8, pageMeta.heightPdf - area.y);
      } else {
        area.x = clamp(startArea.x + dx, 0, pageMeta.widthPdf - area.width);
        area.y = clamp(startArea.y + dy, 0, pageMeta.heightPdf - area.height);
      }
      syncAreaForm();
      renderAreaList();
      renderPdfPages();
    }

    function onUp() {
      window.removeEventListener('mousemove', onMove);
      window.removeEventListener('mouseup', onUp);
      markDirty();
    }

    window.addEventListener('mousemove', onMove);
    window.addEventListener('mouseup', onUp);
  });

  return overlay;
}

async function renderPdfPages() {
  if (!state.pdfDoc) {
    el.pdfContainer.innerHTML = '';
    el.pageInfo.textContent = 'Page 0 / 0';
    return;
  }
  el.pdfContainer.innerHTML = '';
  state.pagesMeta = [];
  for (let pageNumber = 1; pageNumber <= state.pdfDoc.numPages; pageNumber++) {
    const page = await state.pdfDoc.getPage(pageNumber);
    const viewport = page.getViewport({ scale: state.zoom });
    const viewportPdf = page.getViewport({ scale: 1 });
    const wrap = document.createElement('div');
    wrap.className = 'page-wrap';
    wrap.style.width = `${viewport.width}px`;
    wrap.style.height = `${viewport.height}px`;

    const canvas = document.createElement('canvas');
    canvas.className = 'page-canvas';
    canvas.width = viewport.width;
    canvas.height = viewport.height;
    wrap.appendChild(canvas);

    const overlayNode = document.createElement('div');
    overlayNode.className = 'page-overlay';
    wrap.appendChild(overlayNode);

    await page.render({ canvasContext: canvas.getContext('2d'), viewport }).promise;

    const pageMeta = {
      pageNumber,
      scale: viewport.width / viewportPdf.width,
      widthPdf: viewportPdf.width,
      heightPdf: viewportPdf.height,
      overlayNode,
    };
    state.pagesMeta.push(pageMeta);

    state.overlays
      .filter(area => Number(area.page) === pageNumber)
      .forEach(area => overlayNode.appendChild(makeOverlayNode(area, pageMeta)));

    el.pdfContainer.appendChild(wrap);
  }
  el.pageInfo.textContent = `Page 1 / ${state.pdfDoc.numPages}`;
}

async function loadPdfBytes(bytes) {
  state.previewBytes = bytes;
  const loadingTask = pdfjsLib.getDocument({ data: bytes });
  state.pdfDoc = await loadingTask.promise;
  await renderPdfPages();
}

async function loadRecord(record) {
  state.selectedRecord = record;
  syncMetaPanel();
  setStatus('Mengambil PDF asli...');
  const json = await apiGet('file', { fileId: record.fileId });
  state.pdfBytes = parseBase64(json.base64);
  state.previewBytes = state.pdfBytes.slice();
  state.overlays = Array.isArray(record.overlays) ? structuredClone(record.overlays) : [];
  state.selectedOverlayId = state.overlays[0]?.id || '';
  syncAreaForm();
  renderAreaList();
  await loadPdfBytes(state.previewBytes);
  setStatus('Preview PDF asli siap');
}

async function refreshDirectory() {
  setStatus('Memuat daftar file dari folder Google Drive...');
  const json = await apiGet('list');
  state.periods = json.periods || [];
  state.records = json.files || [];
  state.selectedPeriod = state.periods[0] || '';
  renderPeriods();
  filterRecords();
  setStatus('Daftar file siap');
}

function downloadBytes(bytes, fileName) {
  const blob = new Blob([bytes], { type: 'application/pdf' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName || 'edited.pdf';
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 5000);
}

async function buildEditedPdfBytes() {
  const { PDFDocument, StandardFonts, rgb } = PDFLib;
  const pdfDoc = await PDFDocument.load(state.pdfBytes);
  const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
  const pages = pdfDoc.getPages();

  state.overlays.forEach(area => {
    const page = pages[area.page - 1];
    if (!page) return;
    const size = page.getSize();
    const drawY = size.height - area.y - area.height;
    page.drawRectangle({ x: area.x, y: drawY, width: area.width, height: area.height, color: rgb(1, 1, 1) });

    const lines = String(area.text || '').split(/\n/);
    const maxChars = Math.max(1, Math.floor(area.width / (area.fontSize * 0.55)));
    const wrapped = [];
    lines.forEach(line => {
      if (line.length <= maxChars) wrapped.push(line);
      else {
        let cursor = line;
        while (cursor.length > maxChars) {
          wrapped.push(cursor.slice(0, maxChars));
          cursor = cursor.slice(maxChars);
        }
        if (cursor) wrapped.push(cursor);
      }
    });

    wrapped.slice(0, Math.max(1, Math.floor(area.height / (area.fontSize + 2)))).forEach((line, idx) => {
      page.drawText(line, {
        x: area.x + 1,
        y: drawY + area.height - area.fontSize - 1 - (idx * (area.fontSize + 1)),
        font,
        size: area.fontSize,
        color: rgb(0, 0, 0),
        maxWidth: area.width - 2,
      });
    });
  });

  return await pdfDoc.save();
}

async function savePdf(isAuto = false) {
  if (!state.selectedRecord || state.isSaving || !state.isDirty) return;
  state.isSaving = true;
  setStatus(isAuto ? 'Autosave: menyimpan dan menimpa PDF asli...' : 'Menyimpan dan menimpa PDF asli...');
  try {
    const edited = await buildEditedPdfBytes();
    const payload = {
      fileId: state.selectedRecord.fileId,
      fileName: state.selectedRecord.fileName,
      mimeType: 'application/pdf',
      base64: bytesToBase64(edited),
      overlays: state.overlays,
      record: {
        ...state.selectedRecord,
        overlays: state.overlays,
      },
    };
    await apiPost('overwrite', payload);
    state.pdfBytes = new Uint8Array(edited);
    state.previewBytes = state.pdfBytes.slice();
    await loadPdfBytes(state.previewBytes);
    state.isDirty = false;
    setStatus(isAuto ? 'Autosave selesai, PDF asli sudah ditimpa' : 'PDF asli berhasil ditimpa');
  } catch (error) {
    console.error(error);
    setStatus(`Gagal simpan: ${error.message}`);
  } finally {
    state.isSaving = false;
  }
}

function printCurrent() {
  if (!state.previewBytes) return;
  const blob = new Blob([state.previewBytes], { type: 'application/pdf' });
  const url = URL.createObjectURL(blob);
  const win = window.open(url, '_blank');
  if (win) {
    win.addEventListener('load', () => {
      win.focus();
      win.print();
    });
  }
}

function bindEvents() {
  el.periodSelect.addEventListener('change', () => {
    state.selectedPeriod = el.periodSelect.value;
    filterRecords();
  });

  el.searchInput.addEventListener('input', filterRecords);

  el.employeeSelect.addEventListener('change', async () => {
    const record = state.filteredRecords[Number(el.employeeSelect.value || 0)];
    if (record) await loadRecord(record);
  });

  el.refreshBtn.addEventListener('click', refreshDirectory);
  el.syncBtn.addEventListener('click', async () => {
    if (state.selectedRecord) await loadRecord(state.selectedRecord);
  });

  el.addAreaBtn.addEventListener('click', createArea);
  el.updateAreaBtn.addEventListener('click', updateArea);
  el.deleteAreaBtn.addEventListener('click', deleteArea);
  el.saveBtn.addEventListener('click', () => savePdf(false));
  el.printBtn.addEventListener('click', printCurrent);
  el.downloadBtn.addEventListener('click', () => {
    if (state.previewBytes) downloadBytes(state.previewBytes, state.selectedRecord?.fileName || 'payslip.pdf');
  });
  el.zoomInBtn.addEventListener('click', async () => {
    state.zoom = Math.min(2.5, state.zoom + 0.1);
    await renderPdfPages();
  });
  el.zoomOutBtn.addEventListener('click', async () => {
    state.zoom = Math.max(0.7, state.zoom - 0.1);
    await renderPdfPages();
  });

  ['input', 'change'].forEach(evt => {
    [el.areaPage, el.areaFontSize, el.areaX, el.areaY, el.areaWidth, el.areaHeight, el.areaLabel, el.areaText].forEach(node => {
      node.addEventListener(evt, () => {
        const area = getSelectedOverlay();
        if (!area) return;
        Object.assign(area, areaFromForm());
        renderAreaList();
        renderPdfPages();
        markDirty();
      });
    });
  });
}

async function init() {
  bindEvents();
  try {
    if (APPS_SCRIPT_URL.includes('PASTE_YOUR_APPS_SCRIPT')) {
      setStatus('Isi dulu APPS_SCRIPT_URL di app.js');
      return;
    }
    await refreshDirectory();
    if (state.filteredRecords.length) {
      el.employeeSelect.value = '0';
      await loadRecord(state.filteredRecords[0]);
    }
  } catch (error) {
    console.error(error);
    setStatus(`Error: ${error.message}`);
  }
}

init();
