const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycby6jTm5xQmcQAUjdaubsOWOn7Xws4UWjV9uWbOExlEQArCSN6hubMt3U128QjmlWZP0Ow/exec';
const ROOT_FOLDER_ID = '1NZfDp_9SU50OVDJXLuTcGZNj5JvSdpdX';
const CACHE_KEY = 'payslip_bip_list_cache_v2';
const SETTINGS_KEY = 'payslip_bip_ui_settings_v1';
const CACHE_TTL_MS = 1000 * 60 * 10;
const pdfjsLib = globalThis.pdfjsLib;
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.5.136/pdf.worker.min.mjs';

const state = {
  periods: [],
  records: [],
  filteredRecords: [],
  selectedPeriod: '',
  selectedRecord: null,
  pdfBytes: null,
  previewBytes: null,
  pdfDoc: null,
  zoom: 1.28,
  overlays: [],
  selectedOverlayId: '',
  autosaveTimer: null,
  isDirty: false,
  isSaving: false,
  pagesMeta: [],
  searchIndex: -1,
  cacheLoaded: false,
  lastListAt: 0,
};

const el = {
  periodSelect: document.getElementById('periodSelect'),
  searchInput: document.getElementById('searchInput'),
  searchResults: document.getElementById('searchResults'),
  refreshBtn: document.getElementById('refreshBtn'),
  syncBtn: document.getElementById('syncBtn'),
  printBtn: document.getElementById('printBtn'),
  zoomOutBtn: document.getElementById('zoomOutBtn'),
  zoomInBtn: document.getElementById('zoomInBtn'),
  pdfContainer: document.getElementById('pdfContainer'),
  statusText: document.getElementById('statusText'),
  fileNameText: document.getElementById('fileNameText'),
  employeeIdText: document.getElementById('employeeIdText'),
  employeeNameText: document.getElementById('employeeNameText'),
  pageInfo: document.getElementById('pageInfo'),
  dbStatusWrap: document.getElementById('dbStatusWrap'),
  dbStatusText: document.getElementById('dbStatusText'),
  brandLogo: document.getElementById('brandLogo'),
  settingsBtn: document.getElementById('settingsBtn'),
  settingsModal: document.getElementById('settingsModal'),
  titleInput: document.getElementById('titleInput'),
  logoInput: document.getElementById('logoInput'),
  saveSettingsBtn: document.getElementById('saveSettingsBtn'),
  resetLogoBtn: document.getElementById('resetLogoBtn'),
};

function setStatus(text) { el.statusText.textContent = text; }
function setDbState(stateName, text) {
  el.dbStatusWrap.dataset.state = stateName;
  el.dbStatusText.textContent = text;
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
  const res = await fetch(apiUrl(action, params), { cache: 'no-store' });
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
function normalize(value) { return String(value || '').toLowerCase().trim(); }
function escapeHtml(value) {
  return String(value ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}
function parseBase64(base64) {
  const binary = atob(base64); const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes;
}
function bytesToBase64(bytes) {
  let binary = ''; const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) binary += String.fromCharCode(...bytes.subarray(i, i + chunk));
  return btoa(binary);
}
function saveCache() {
  try {
    localStorage.setItem(CACHE_KEY, JSON.stringify({ at: Date.now(), periods: state.periods, records: state.records }));
  } catch {}
}
function loadCache() {
  try {
    const raw = localStorage.getItem(CACHE_KEY);
    if (!raw) return false;
    const parsed = JSON.parse(raw);
    const age = Date.now() - Number(parsed.at || 0);
    state.periods = parsed.periods || [];
    state.records = parsed.records || [];
    state.lastListAt = parsed.at || 0;
    state.selectedPeriod = state.selectedPeriod || state.periods[0] || '';
    state.cacheLoaded = true;
    renderPeriods();
    filterRecords();
    setDbState('connected', age <= CACHE_TTL_MS ? 'Terhubung (cache cepat)' : 'Terhubung (cache lama)');
    setStatus(age <= CACHE_TTL_MS ? 'Daftar file dimuat instan dari cache lokal' : 'Cache lokal dimuat, sedang cek update');
    return !!state.records.length;
  } catch {
    return false;
  }
}
function loadUiSettings() {
  try {
    const raw = localStorage.getItem(SETTINGS_KEY);
    const settings = raw ? JSON.parse(raw) : {};
    const title = settings.title || 'Payslip BIP';
    document.title = title;
    document.querySelector('.brand-left h1').textContent = title;
    el.titleInput.value = title;
    if (settings.logoDataUrl) el.brandLogo.src = settings.logoDataUrl;
  } catch {}
}
function saveUiSettings() {
  const readerTask = () => {
    const current = { title: el.titleInput.value.trim() || 'Payslip BIP', logoDataUrl: el.brandLogo.src };
    localStorage.setItem(SETTINGS_KEY, JSON.stringify(current));
    document.title = current.title;
    document.querySelector('.brand-left h1').textContent = current.title;
    closeSettings();
  };
  const file = el.logoInput.files?.[0];
  if (!file) return readerTask();
  const fr = new FileReader();
  fr.onload = () => { el.brandLogo.src = fr.result; readerTask(); };
  fr.readAsDataURL(file);
}
function openSettings() { el.settingsModal.classList.remove('hidden'); }
function closeSettings() { el.settingsModal.classList.add('hidden'); }
function renderPeriods() {
  el.periodSelect.innerHTML = state.periods.map(period => `<option value="${escapeHtml(period)}">${escapeHtml(period)}</option>`).join('');
  if (state.periods.length && !state.selectedPeriod) state.selectedPeriod = state.periods[0];
  el.periodSelect.value = state.selectedPeriod;
}
function searchMatches(query) {
  const q = normalize(query);
  return state.records.filter(record => {
    if (state.selectedPeriod && record.periodLabel !== state.selectedPeriod) return false;
    if (!q) return false;
    return normalize(record.employeeId).includes(q) || normalize(record.employeeName).includes(q) || normalize(record.fileName).includes(q);
  }).slice(0, 12);
}
function renderSearchResults() {
  const q = el.searchInput.value.trim();
  const rows = q ? searchMatches(q) : [];
  state.filteredRecords = rows;
  state.searchIndex = rows.length ? Math.min(Math.max(state.searchIndex, 0), rows.length - 1) : -1;
  if (!rows.length) {
    el.searchResults.innerHTML = '';
    el.searchResults.classList.add('hidden');
    return;
  }
  el.searchResults.innerHTML = rows.map((record, index) => `
    <div class="search-item ${index === state.searchIndex ? 'active' : ''}" data-index="${index}">
      <div class="search-item-title">${escapeHtml(record.employeeId || '-')} - ${escapeHtml(record.employeeName || '-')}</div>
      <div class="search-item-sub">${escapeHtml(record.periodLabel || '-')} • ${escapeHtml(record.fileName || '-')}</div>
    </div>`).join('');
  el.searchResults.classList.remove('hidden');
}
function hideSearchResults() { el.searchResults.classList.add('hidden'); }
function filterRecords() {
  const q = normalize(el.searchInput.value);
  state.filteredRecords = state.records.filter(record => {
    if (state.selectedPeriod && record.periodLabel !== state.selectedPeriod) return false;
    if (!q) return true;
    return normalize(record.employeeId).includes(q) || normalize(record.employeeName).includes(q) || normalize(record.fileName).includes(q);
  });
  renderSearchResults();
}
function syncMetaPanel() {
  const r = state.selectedRecord;
  el.fileNameText.textContent = r?.fileName || '-';
  el.employeeIdText.textContent = r?.employeeId || '-';
  el.employeeNameText.textContent = r?.employeeName || '-';
}
function getSelectedOverlay() { return state.overlays.find(item => item.id === state.selectedOverlayId) || null; }
function markDirty() {
  state.isDirty = true;
  setStatus('Perubahan terdeteksi, autosave berjalan...');
  if (state.autosaveTimer) clearTimeout(state.autosaveTimer);
  state.autosaveTimer = setTimeout(() => savePdf(true), 700);
}
function clamp(v, min, max) { return Math.max(min, Math.min(max, v)); }
function pickRecord(index) {
  const record = state.filteredRecords[index];
  if (!record) return;
  el.searchInput.value = `${record.employeeId || ''} - ${record.employeeName || ''}`.trim();
  hideSearchResults();
  loadRecord(record);
}
function fitTextAreaSize(area, editor, scale) {
  editor.style.width = `${Math.max(12, area.width * scale)}px`;
  editor.style.height = `${Math.max(12, area.height * scale)}px`;
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

  const chip = document.createElement('div');
  chip.className = 'overlay-chip';
  chip.textContent = area.label || 'Text';
  overlay.appendChild(chip);

  const editor = document.createElement('textarea');
  editor.className = 'overlay-editor';
  editor.spellcheck = false;
  editor.value = area.text || '';
  editor.style.fontSize = `${Math.max(8, area.fontSize * scale)}px`;
  fitTextAreaSize(area, editor, scale);

  editor.addEventListener('mousedown', ev => {
    ev.stopPropagation();
    state.selectedOverlayId = area.id;
    overlay.classList.add('active');
  });
  editor.addEventListener('focus', () => {
    state.selectedOverlayId = area.id;
    renderPdfPages();
  });
  editor.addEventListener('input', () => {
    area.text = editor.value;
    const minHeightPdf = 12;
    const pxHeight = Math.max(editor.scrollHeight + 2, minHeightPdf * scale);
    area.height = Math.max(minHeightPdf, pxHeight / scale);
    fitTextAreaSize(area, editor, scale);
    overlay.style.height = `${area.height * scale}px`;
    markDirty();
  });
  overlay.appendChild(editor);

  const handle = document.createElement('div');
  handle.className = 'resize-handle';
  overlay.appendChild(handle);

  overlay.addEventListener('mousedown', (event) => {
    state.selectedOverlayId = area.id;
    const startX = event.clientX;
    const startY = event.clientY;
    const startArea = { ...area };
    const resizing = event.target.classList.contains('resize-handle');

    function onMove(ev) {
      const dx = (ev.clientX - startX) / scale;
      const dy = (ev.clientY - startY) / scale;
      if (resizing) {
        area.width = clamp(startArea.width + dx, 18, pageMeta.widthPdf - area.x);
        area.height = clamp(startArea.height + dy, 12, pageMeta.heightPdf - area.y);
      } else if (event.target !== editor) {
        area.x = clamp(startArea.x + dx, 0, pageMeta.widthPdf - area.width);
        area.y = clamp(startArea.y + dy, 0, pageMeta.heightPdf - area.height);
      }
      renderPdfPages();
    }
    function onUp() {
      window.removeEventListener('mousemove', onMove);
      window.removeEventListener('mouseup', onUp);
      markDirty();
    }
    if (event.target !== editor) {
      window.addEventListener('mousemove', onMove);
      window.addEventListener('mouseup', onUp);
    }
  });
  return overlay;
}
async function renderPdfPages() {
  if (!state.pdfDoc) {
    el.pdfContainer.innerHTML = '';
    el.pageInfo.textContent = 'Page 0 / 0';
    return;
  }
  const activeId = state.selectedOverlayId;
  el.pdfContainer.innerHTML = '';
  state.pagesMeta = [];
  const outputScale = Math.min(2.2, Math.max(1.5, window.devicePixelRatio || 1.5));
  for (let pageNumber = 1; pageNumber <= state.pdfDoc.numPages; pageNumber++) {
    const page = await state.pdfDoc.getPage(pageNumber);
    const cssViewport = page.getViewport({ scale: state.zoom });
    const renderViewport = page.getViewport({ scale: state.zoom * outputScale });
    const viewportPdf = page.getViewport({ scale: 1 });

    const wrap = document.createElement('div');
    wrap.className = 'page-wrap';
    wrap.style.width = `${cssViewport.width}px`;
    wrap.style.height = `${cssViewport.height}px`;

    const canvas = document.createElement('canvas');
    canvas.className = 'page-canvas';
    canvas.width = Math.floor(renderViewport.width);
    canvas.height = Math.floor(renderViewport.height);
    canvas.style.width = `${cssViewport.width}px`;
    canvas.style.height = `${cssViewport.height}px`;
    const ctx = canvas.getContext('2d', { alpha: false });
    wrap.appendChild(canvas);

    const overlayNode = document.createElement('div');
    overlayNode.className = 'page-overlay';
    wrap.appendChild(overlayNode);

    await page.render({
      canvasContext: ctx,
      viewport: renderViewport,
      transform: null
    }).promise;

    const pageMeta = {
      pageNumber,
      scale: cssViewport.width / viewportPdf.width,
      widthPdf: viewportPdf.width,
      heightPdf: viewportPdf.height,
      overlayNode
    };
    state.pagesMeta.push(pageMeta);

    overlayNode.addEventListener('dblclick', (event) => {
      const rect = overlayNode.getBoundingClientRect();
      const x = (event.clientX - rect.left) / pageMeta.scale;
      const y = (event.clientY - rect.top) / pageMeta.scale;
      const id = `area_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`;
      state.overlays.push({
        id,
        page: pageNumber,
        label: 'Text Baru',
        text: '',
        x: clamp(x, 0, pageMeta.widthPdf - 140),
        y: clamp(y, 0, pageMeta.heightPdf - 22),
        width: 140,
        height: 20,
        fontSize: 12
      });
      state.selectedOverlayId = id;
      renderPdfPages().then(() => {
        overlayNode.querySelector(`.overlay-box[data-id="${id}"] .overlay-editor`)?.focus();
      });
      markDirty();
    });

    state.overlays
      .filter(area => Number(area.page) === pageNumber)
      .forEach(area => overlayNode.appendChild(makeOverlayNode(area, pageMeta)));

    el.pdfContainer.appendChild(wrap);
  }
  state.selectedOverlayId = activeId || state.selectedOverlayId;
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
  await loadPdfBytes(state.previewBytes);
  setStatus('Preview PDF asli siap');
}
async function refreshDirectory(force = false) {
  const shouldUseCacheFirst = !force && state.records.length;
  if (!shouldUseCacheFirst) {
    setDbState('loading', 'Menghubungkan database...');
    setStatus('Memuat daftar file dari folder Google Drive...');
  } else {
    setDbState('connected', 'Terhubung (cache cepat)');
    setStatus('Daftar file dari cache tampil, sedang cek update terbaru...');
  }
  const startedAt = performance.now();
  try {
    const json = await apiGet('list');
    state.periods = json.periods || [];
    state.records = json.files || [];
    state.selectedPeriod = state.selectedPeriod && state.periods.includes(state.selectedPeriod) ? state.selectedPeriod : (state.periods[0] || '');
    renderPeriods();
    filterRecords();
    saveCache();
    const ms = Math.round(performance.now() - startedAt);
    setDbState('connected', `Terhubung • ${ms} ms`);
    setStatus('Daftar file siap');
  } catch (error) {
    setDbState('error', 'Gagal terhubung');
    if (!force && loadCache()) return;
    throw error;
  }
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
    page.drawRectangle({ x: area.x, y: drawY, width: area.width, height: area.height, color: rgb(1,1,1), opacity: 0.985 });
    const lines = String(area.text || '').split(/\n/);
    const maxChars = Math.max(1, Math.floor(area.width / Math.max(1, area.fontSize * 0.52)));
    const wrapped = [];
    lines.forEach(line => {
      if (line.length <= maxChars) wrapped.push(line);
      else {
        let cursor = line;
        while (cursor.length > maxChars) { wrapped.push(cursor.slice(0, maxChars)); cursor = cursor.slice(maxChars); }
        if (cursor) wrapped.push(cursor);
      }
    });
    wrapped.slice(0, Math.max(1, Math.floor(area.height / (area.fontSize + 1.5)))).forEach((line, idx) => {
      page.drawText(line, {
        x: area.x + 1,
        y: drawY + area.height - area.fontSize - 1 - (idx * (area.fontSize + 1)),
        font,
        size: area.fontSize,
        color: rgb(0,0,0),
        maxWidth: area.width - 2
      });
    });
  });
  return await pdfDoc.save();
}
async function savePdf(isAuto = false) {
  if (!state.selectedRecord || state.isSaving || !state.isDirty) return;
  state.isSaving = true;
  setStatus(isAuto ? 'Autosave: menyimpan PDF...' : 'Menyimpan PDF...');
  try {
    const edited = await buildEditedPdfBytes();
    const overlaysCopy = structuredClone(state.overlays);
    await apiPost('overwrite', {
      fileId: state.selectedRecord.fileId,
      fileName: state.selectedRecord.fileName,
      mimeType: 'application/pdf',
      base64: bytesToBase64(edited),
      overlays: overlaysCopy,
      record: { ...state.selectedRecord, overlays: overlaysCopy },
    });
    state.pdfBytes = new Uint8Array(edited);
    state.previewBytes = state.pdfBytes.slice();
    state.selectedRecord.overlays = overlaysCopy;
    const recordIndex = state.records.findIndex(item => item.fileId === state.selectedRecord.fileId);
    if (recordIndex >= 0) state.records[recordIndex] = { ...state.records[recordIndex], overlays: overlaysCopy };
    saveCache();
    state.isDirty = false;
    setStatus(isAuto ? 'Autosave selesai' : 'PDF berhasil disimpan');
    await loadPdfBytes(state.previewBytes);
  } catch (error) {
    console.error(error);
    setStatus(`Gagal simpan: ${error.message}`);
  } finally {
    state.isSaving = false;
  }
}
async function printCurrent() {
  const bytes = state.isDirty ? await buildEditedPdfBytes() : state.previewBytes;
  if (!bytes) return;
  const blob = new Blob([bytes], { type: 'application/pdf' });
  const url = URL.createObjectURL(blob);
  const win = window.open(url, '_blank');
  if (win) {
    const tryPrint = () => {
      try {
        win.focus();
        win.print();
      } catch {}
    };
    setTimeout(tryPrint, 500);
    setTimeout(tryPrint, 1200);
  }
}
function bindEvents() {
  el.periodSelect.addEventListener('change', () => {
    state.selectedPeriod = el.periodSelect.value;
    if (el.searchInput.value.trim()) renderSearchResults();
  });
  el.searchInput.addEventListener('input', () => {
    state.searchIndex = -1;
    renderSearchResults();
  });
  el.searchInput.addEventListener('focus', renderSearchResults);
  el.searchInput.addEventListener('keydown', (event) => {
    if (el.searchResults.classList.contains('hidden')) return;
    const max = state.filteredRecords.length - 1;
    if (event.key === 'ArrowDown') {
      event.preventDefault();
      state.searchIndex = Math.min(max, state.searchIndex + 1);
      renderSearchResults();
    } else if (event.key === 'ArrowUp') {
      event.preventDefault();
      state.searchIndex = Math.max(0, state.searchIndex - 1);
      renderSearchResults();
    } else if (event.key === 'Enter') {
      event.preventDefault();
      if (state.searchIndex >= 0) pickRecord(state.searchIndex);
    } else if (event.key === 'Escape') {
      hideSearchResults();
    }
  });
  el.searchResults.addEventListener('mousedown', (event) => {
    const item = event.target.closest('.search-item');
    if (!item) return;
    pickRecord(Number(item.dataset.index));
  });
  document.addEventListener('click', (event) => {
    if (!event.target.closest('.search-field')) hideSearchResults();
  });
  document.addEventListener('keydown', (event) => {
    const tag = document.activeElement?.tagName || '';
    if ((event.key === 'Delete' || event.key === 'Backspace') && state.selectedOverlayId && tag !== 'TEXTAREA' && tag !== 'INPUT') {
      state.overlays = state.overlays.filter(item => item.id !== state.selectedOverlayId);
      state.selectedOverlayId = state.overlays[0]?.id || '';
      renderPdfPages();
      markDirty();
    }
  });
  el.refreshBtn.addEventListener('click', () => refreshDirectory(true));
  el.syncBtn.addEventListener('click', async () => { if (state.selectedRecord) await loadRecord(state.selectedRecord); });
  el.printBtn.addEventListener('click', printCurrent);
  el.zoomInBtn.addEventListener('click', async () => { state.zoom = Math.min(2.8, state.zoom + 0.1); await renderPdfPages(); });
  el.zoomOutBtn.addEventListener('click', async () => { state.zoom = Math.max(0.7, state.zoom - 0.1); await renderPdfPages(); });
  el.settingsBtn.addEventListener('click', openSettings);
  document.querySelectorAll('[data-close-settings]').forEach(node => node.addEventListener('click', closeSettings));
  el.saveSettingsBtn.addEventListener('click', saveUiSettings);
  el.resetLogoBtn.addEventListener('click', () => {
    el.brandLogo.src = './hoplun.jpg';
    el.logoInput.value = '';
    saveUiSettings();
  });
}
async function init() {
  bindEvents();
  loadUiSettings();
  try {
    if (!APPS_SCRIPT_URL) {
      setStatus('Isi dulu APPS_SCRIPT_URL di app.js');
      return;
    }
    const hadCache = loadCache();
    if (hadCache) {
      const firstCached = state.records.find(r => !state.selectedPeriod || r.periodLabel === state.selectedPeriod) || state.records[0];
      if (firstCached) loadRecord(firstCached).catch(console.error);
      refreshDirectory(false).catch(console.error);
      return;
    }
    await refreshDirectory();
    const first = state.records.find(r => !state.selectedPeriod || r.periodLabel === state.selectedPeriod) || state.records[0];
    if (first) await loadRecord(first);
  } catch (error) {
    console.error(error);
    setDbState('error', 'Error');
    setStatus(`Error: ${error.message}`);
  }
}
init();
