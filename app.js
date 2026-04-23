
const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycby6jTm5xQmcQAUjdaubsOWOn7Xws4UWjV9uWbOExlEQArCSN6hubMt3U128QjmlWZP0Ow/exec';
const ROOT_FOLDER_ID = '1NZfDp_9SU50OVDJXLuTcGZNj5JvSdpdX';
const CACHE_KEY = 'payslip_bip_list_cache_v7';
const SETTINGS_KEY = 'payslip_bip_ui_settings_v1';
const pdfjsLib = globalThis.pdfjsLib || window.pdfjsLib;
if (pdfjsLib) pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

const state = {
  years: [],
  periods: [],
  records: [],
  filteredRecords: [],
  selectedYear: '',
  selectedPeriod: '',
  selectedRecord: null,
  previewBytes: null,
  pdfDoc: null,
  zoom: 1.2,
  searchIndex: -1,
  printFrame: null,
};

const el = {
  yearSelect: document.getElementById('yearSelect'),
  periodSelect: document.getElementById('periodSelect'),
  searchInput: document.getElementById('searchInput'),
  searchResults: document.getElementById('searchResults'),
  refreshBtn: document.getElementById('refreshBtn'),
  syncBtn: document.getElementById('syncBtn'),
  printBtn: document.getElementById('printBtn'),
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

function setStatus(text){ if (el.statusText) el.statusText.textContent = text; }
function setDbState(stateName, text){ if (el.dbStatusWrap) el.dbStatusWrap.dataset.state = stateName; if (el.dbStatusText) el.dbStatusText.textContent = text; }
function normalize(v){ return String(v || '').toLowerCase().trim(); }
function escapeHtml(v){ return String(v ?? '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/\"/g,'&quot;').replace(/'/g,'&#39;'); }

function apiUrl(action, params = {}) {
  const url = new URL(APPS_SCRIPT_URL);
  url.searchParams.set('action', action);
  url.searchParams.set('rootFolderId', ROOT_FOLDER_ID);
  Object.entries(params).forEach(([k,v]) => { if (v !== undefined && v !== null && v !== '') url.searchParams.set(k, v); });
  return url.toString();
}
async function apiGet(action, params = {}) {
  const res = await fetch(apiUrl(action, params), { cache: 'no-store', redirect: 'follow' });
  const text = await res.text();
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  let json; try { json = JSON.parse(text); } catch { throw new Error('Respons Apps Script bukan JSON.'); }
  if (json.ok === false) throw new Error(json.message || 'Request gagal');
  return json;
}
function parseBase64(base64){
  const b = atob(base64);
  const bytes = new Uint8Array(b.length);
  for(let i=0;i<b.length;i++) bytes[i]=b.charCodeAt(i);
  return bytes;
}
function saveCache(){
  try {
    localStorage.setItem(CACHE_KEY, JSON.stringify({
      at: Date.now(),
      years: state.years,
      selectedYear: state.selectedYear,
      periods: state.periods,
      records: state.records
    }));
  } catch {}
}
function loadCache(){
  try {
    const raw = localStorage.getItem(CACHE_KEY);
    if (!raw) return false;
    const parsed = JSON.parse(raw);
    state.years = parsed.years || [];
    state.selectedYear = parsed.selectedYear || parsed.years?.[0] || '';
    state.periods = parsed.periods || [];
    state.records = parsed.records || [];
    state.selectedPeriod = state.selectedPeriod || state.periods[0] || '';
    renderYearOptions();
    renderPeriodOptions();
    setDbState('connected', 'Terhubung (cache cepat)');
    setStatus('Daftar file dimuat instan dari index cache');
    return !!state.records.length;
  } catch { return false; }
}
function extractSrId(record){
  const candidates = [record?.employeeId, record?.employeeCode, record?.srId, record?.fileName].filter(Boolean);
  for (const value of candidates) {
    const match = String(value).toUpperCase().match(/\bSR[0-9]{5,}\b/);
    if (match) return match[0];
  }
  return '';
}
function extractDisplayName(record){
  const explicitName = String(record?.employeeName || record?.name || '').trim();
  if (explicitName) return explicitName;
  const fileName = String(record?.fileName || '');
  const match = fileName.match(/^[^-]+-(SR[0-9]{5,})-(.+?)-ID[0-9]{8}-Payslip\.pdf$/i);
  if (match && match[2]) return match[2].replace(/[-_]+/g, ' ').trim();
  return '';
}
function getDisplayId(record){ return extractSrId(record) || String(record?.employeeId || '').trim(); }
function getDisplayName(record){ return extractDisplayName(record) || String(record?.employeeName || '').trim(); }

function loadUiSettings() {
  try {
    const raw = localStorage.getItem(SETTINGS_KEY);
    const settings = raw ? JSON.parse(raw) : {};
    const title = settings.title || 'Payslip BIP';
    document.title = title;
    const titleNode = document.querySelector('.brand-left h1');
    if (titleNode) titleNode.textContent = title;
    if (el.titleInput) el.titleInput.value = title;
    if (settings.logoDataUrl && el.brandLogo) el.brandLogo.src = settings.logoDataUrl;
  } catch {}
}
function saveUiSettings() {
  const persist = () => {
    const current = { title: (el.titleInput?.value || 'Payslip BIP').trim() || 'Payslip BIP', logoDataUrl: el.brandLogo?.src || './hoplun.jpg' };
    localStorage.setItem(SETTINGS_KEY, JSON.stringify(current));
    document.title = current.title;
    const titleNode = document.querySelector('.brand-left h1');
    if (titleNode) titleNode.textContent = current.title;
    closeSettings();
  };
  const file = el.logoInput?.files?.[0];
  if (!file) return persist();
  const fr = new FileReader();
  fr.onload = () => { if (el.brandLogo) el.brandLogo.src = fr.result; persist(); };
  fr.readAsDataURL(file);
}
function openSettings() { el.settingsModal?.classList.remove('hidden'); }
function closeSettings() { el.settingsModal?.classList.add('hidden'); }

function renderYearOptions() {
  if (!el.yearSelect) return;
  el.yearSelect.innerHTML = state.years.map(year => `<option value="${escapeHtml(year)}">${escapeHtml(year)}</option>`).join('');
  if (state.years.length && !state.selectedYear) state.selectedYear = state.years[state.years.length - 1];
  el.yearSelect.value = state.selectedYear;
}
function renderPeriodOptions() {
  if (!el.periodSelect) return;
  el.periodSelect.innerHTML = state.periods.map(period => `<option value="${escapeHtml(period)}">${escapeHtml(period)}</option>`).join('');
  if (state.periods.length && !state.selectedPeriod) state.selectedPeriod = state.periods[0];
  el.periodSelect.value = state.selectedPeriod;
}
function getSearchRows(){
  const q = normalize(el.searchInput?.value || '');
  return state.records.filter(record => {
    if (state.selectedPeriod && record.periodLabel !== state.selectedPeriod) return false;
    if (!q) return false;
    return normalize(getDisplayId(record)).includes(q) || normalize(getDisplayName(record)).includes(q);
  }).slice(0, 12);
}
function renderSearchResults(){
  if (!el.searchResults) return;
  const rows = getSearchRows();
  state.filteredRecords = rows;
  if (!rows.length) {
    el.searchResults.innerHTML = '';
    el.searchResults.classList.add('hidden');
    return;
  }
  if (state.searchIndex < 0 || state.searchIndex > rows.length - 1) state.searchIndex = 0;
  el.searchResults.innerHTML = rows.map((record, index) => `
    <div class="search-item ${index === state.searchIndex ? 'active' : ''}" data-index="${index}">
      <div class="search-item-title">${escapeHtml(getDisplayId(record) || '-')} - ${escapeHtml(getDisplayName(record) || '-')}</div>
      <div class="search-item-sub">${escapeHtml(record.periodLabel || '-')}</div>
    </div>
  `).join('');
  el.searchResults.classList.remove('hidden');
}
function hideSearchResults(){ el.searchResults?.classList.add('hidden'); }
function syncMetaPanel(){
  const r = state.selectedRecord;
  if (el.fileNameText) el.fileNameText.textContent = r?.fileName || '-';
  if (el.employeeIdText) el.employeeIdText.textContent = getDisplayId(r) || '-';
  if (el.employeeNameText) el.employeeNameText.textContent = getDisplayName(r) || '-';
}
function hideSubtitleAndPremiumButtons(){
  const subtitle = document.querySelector('.toolbar-subtitle');
  if (subtitle) subtitle.style.display = 'none';
  if (el.refreshBtn) {
    el.refreshBtn.textContent = '✦ Refresh';
    Object.assign(el.refreshBtn.style, { background:'linear-gradient(135deg,#ffffff,#eef4ff)', border:'1px solid #cfe0ff', color:'#0f172a', boxShadow:'0 10px 22px rgba(37,99,235,.10)', fontWeight:'800' });
  }
  if (el.syncBtn) {
    el.syncBtn.textContent = '⟳ Reload';
    Object.assign(el.syncBtn.style, { background:'linear-gradient(135deg,#0f172a,#334155)', border:'1px solid #1e293b', color:'#fff', boxShadow:'0 12px 24px rgba(15,23,42,.18)', fontWeight:'800' });
  }
}
async function loadPdfBytes(bytes){
  state.previewBytes = bytes;
  const loadingTask = pdfjsLib.getDocument({ data: bytes });
  state.pdfDoc = await loadingTask.promise;
  const page = await state.pdfDoc.getPage(1);
  const scaleRender = Math.min(2.2, Math.max(1.5, window.devicePixelRatio || 1.5));
  const cssViewport = page.getViewport({ scale: state.zoom });
  const renderViewport = page.getViewport({ scale: state.zoom * scaleRender });
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
  const ctx = canvas.getContext('2d', { alpha:false });
  await page.render({ canvasContext: ctx, viewport: renderViewport }).promise;
  if (el.pdfContainer) {
    el.pdfContainer.innerHTML = '';
    wrap.appendChild(canvas);
    el.pdfContainer.appendChild(wrap);
  }
  if (el.pageInfo) el.pageInfo.textContent = `Page 1 / ${state.pdfDoc.numPages}`;
}
async function loadRecord(record){
  try {
    state.selectedRecord = record;
    syncMetaPanel();
    setStatus('Mengambil PDF asli...');
    const json = await apiGet('file', { fileId: record.fileId });
    state.previewBytes = parseBase64(json.base64);
    await loadPdfBytes(state.previewBytes);
    setStatus('Preview PDF asli siap');
  } catch (error) {
    console.error(error);
    setStatus(`Gagal memuat PDF: ${error.message}`);
  }
}
async function fetchYearData(year, force = false){
  setDbState('loading', 'Menghubungkan database...');
  setStatus(force ? `Menyegarkan index ${year}...` : `Memuat data ${year}...`);
  const json = await apiGet('list', { year: year, forceRefresh: force ? '1' : '' });
  state.years = json.years || [];
  state.selectedYear = json.selectedYear || year;
  state.periods = json.periods || [];
  state.records = json.files || [];
  if (!state.periods.includes(state.selectedPeriod)) state.selectedPeriod = state.periods[0] || '';
  renderYearOptions();
  renderPeriodOptions();
  saveCache();
  setDbState('connected', force ? 'Terhubung (index baru)' : 'Terhubung');
  setStatus(`Daftar ${state.selectedYear} siap • ${state.records.length} file`);
}
async function ensurePdfReady() {
  if (state.previewBytes && state.previewBytes.length) return state.previewBytes;
  if (state.selectedRecord?.fileId) {
    const json = await apiGet('file', { fileId: state.selectedRecord.fileId });
    state.previewBytes = parseBase64(json.base64);
    return state.previewBytes;
  }
  return null;
}
function cleanupPrintFrame(){
  if (state.printFrame && state.printFrame.parentNode) state.printFrame.parentNode.removeChild(state.printFrame);
  state.printFrame = null;
}
async function printCurrent(){
  try {
    setStatus('Menyiapkan cetak...');
    const bytes = await ensurePdfReady();
    if (!bytes || !bytes.length) { setStatus('PDF belum siap untuk dicetak'); return; }
    cleanupPrintFrame();
    const blob = new Blob([bytes], { type:'application/pdf' });
    const blobUrl = URL.createObjectURL(blob);
    const iframe = document.createElement('iframe');
    Object.assign(iframe.style, { position:'fixed', right:'0', bottom:'0', width:'1px', height:'1px', border:'0', opacity:'0' });
    state.printFrame = iframe;
    iframe.onload = () => {
      const win = iframe.contentWindow;
      if (!win) return;
      setTimeout(() => { try { win.focus(); win.print(); setStatus('Dialog print dibuka'); } catch(e) { setStatus('Gagal membuka dialog print'); } }, 700);
      setTimeout(() => { URL.revokeObjectURL(blobUrl); cleanupPrintFrame(); }, 15000);
    };
    document.body.appendChild(iframe);
    iframe.src = blobUrl;
  } catch (error) {
    console.error(error);
    setStatus(`Gagal print: ${error.message}`);
  }
}
async function pickRecord(index){
  const record = state.filteredRecords[index];
  if (!record) return;
  if (el.searchInput) el.searchInput.value = `${getDisplayId(record) || ''} - ${getDisplayName(record) || ''}`.trim();
  hideSearchResults();
  await loadRecord(record);
}
function bindEvents(){
  el.yearSelect?.addEventListener('change', async () => {
    state.selectedYear = el.yearSelect.value;
    state.selectedPeriod = '';
    state.searchIndex = -1;
    hideSearchResults();
    await fetchYearData(state.selectedYear, false);
  });
  el.periodSelect?.addEventListener('change', () => {
    state.selectedPeriod = el.periodSelect.value;
    state.searchIndex = -1;
    renderSearchResults();
  });
  el.searchInput?.addEventListener('input', () => { state.searchIndex = -1; renderSearchResults(); });
  el.searchInput?.addEventListener('focus', renderSearchResults);
  el.searchInput?.addEventListener('keydown', async (event) => {
    if (el.searchResults?.classList.contains('hidden')) return;
    const max = state.filteredRecords.length - 1;
    if (event.key === 'ArrowDown') { event.preventDefault(); state.searchIndex = Math.min(max, state.searchIndex + 1); renderSearchResults(); }
    else if (event.key === 'ArrowUp') { event.preventDefault(); state.searchIndex = Math.max(0, state.searchIndex - 1); renderSearchResults(); }
    else if (event.key === 'Enter') { event.preventDefault(); if (state.searchIndex >= 0) await pickRecord(state.searchIndex); }
    else if (event.key === 'Escape') hideSearchResults();
  });
  el.searchResults?.addEventListener('mousedown', async (event) => {
    const item = event.target.closest('.search-item');
    if (!item) return;
    await pickRecord(Number(item.dataset.index));
  });
  document.addEventListener('click', (event) => { if (!event.target.closest('.search-field')) hideSearchResults(); });
  el.refreshBtn?.addEventListener('click', async () => { await fetchYearData(state.selectedYear || '', true); });
  el.syncBtn?.addEventListener('click', async () => { if (state.selectedRecord) await loadRecord(state.selectedRecord); });
  el.printBtn?.addEventListener('click', printCurrent);
  el.settingsBtn?.addEventListener('click', openSettings);
  document.querySelectorAll('[data-close-settings]').forEach(node => node.addEventListener('click', closeSettings));
  el.saveSettingsBtn?.addEventListener('click', saveUiSettings);
  el.resetLogoBtn?.addEventListener('click', () => { if (el.brandLogo) el.brandLogo.src = './hoplun.jpg'; if (el.logoInput) el.logoInput.value = ''; saveUiSettings(); });
}
async function init(){
  bindEvents();
  loadUiSettings();
  hideSubtitleAndPremiumButtons();
  if (!pdfjsLib) {
    setDbState('error','PDF.js gagal dimuat');
    setStatus('Library PDF.js gagal dimuat.');
    return;
  }
  try {
    loadCache();
    const firstYear = state.selectedYear || '';
    await fetchYearData(firstYear, false);
    const first = state.records.find(r => !state.selectedPeriod || r.periodLabel === state.selectedPeriod) || state.records[0];
    if (first) await loadRecord(first);
  } catch (error) {
    console.error(error);
  }
}
init();
