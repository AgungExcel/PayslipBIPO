const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycby6jTm5xQmcQAUjdaubsOWOn7Xws4UWjV9uWbOExlEQArCSN6hubMt3U128QjmlWZP0Ow/exec';
const ROOT_FOLDER_ID = '1NZfDp_9SU50OVDJXLuTcGZNj5JvSdpdX';
const CACHE_KEY = 'payslip_bip_list_cache_v2';
const SETTINGS_KEY = 'payslip_bip_ui_settings_v1';
const CACHE_TTL_MS = 1000 * 60 * 10;

const pdfjsLib = globalThis.pdfjsLib || window.pdfjsLib;
if (pdfjsLib) {
  pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
}

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
function escapeHtml(value) {
  return String(value ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}
function normalize(value) { return String(value || '').toLowerCase().trim(); }

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
  const url = apiUrl(action, params);
  const res = await fetch(url, { cache: 'no-store', redirect: 'follow' });
  const text = await res.text();
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  let json;
  try {
    json = JSON.parse(text);
  } catch {
    throw new Error('Respons Apps Script bukan JSON. Pastikan deployment Web App aktif dan aksesnya Anyone.');
  }
  if (json.ok === false) throw new Error(json.message || 'Request gagal');
  return json;
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
    state.periods = parsed.periods || [];
    state.records = parsed.records || [];
    state.selectedPeriod = state.selectedPeriod || state.periods[0] || '';
    renderPeriods();
    renderSearchResults();
    setDbState('connected', 'Terhubung (cache cepat)');
    setStatus('Daftar file dimuat dari cache lokal');
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

async function refreshDirectory(force = false) {
  setDbState('loading', 'Menghubungkan database...');
  setStatus('Memuat daftar file dari folder Google Drive...');
  try {
    const json = await apiGet('list');
    state.periods = json.periods || [];
    state.records = json.files || [];
    state.selectedPeriod = state.selectedPeriod && state.periods.includes(state.selectedPeriod) ? state.selectedPeriod : (state.periods[0] || '');
    renderPeriods();
    saveCache();
    setDbState('connected', 'Terhubung');
    setStatus('Daftar file siap');
  } catch (error) {
    if (!force && loadCache()) return;
    setDbState('error', 'Gagal terhubung');
    setStatus(`Error: ${error.message}`);
    throw error;
  }
}

function bindEvents() {
  el.periodSelect.addEventListener('change', () => {
    state.selectedPeriod = el.periodSelect.value;
    renderSearchResults();
  });
  el.searchInput.addEventListener('input', () => {
    state.searchIndex = -1;
    renderSearchResults();
  });
  el.searchInput.addEventListener('focus', renderSearchResults);
  el.refreshBtn.addEventListener('click', () => refreshDirectory(true).catch(() => {}));
  el.syncBtn.addEventListener('click', () => refreshDirectory(true).catch(() => {}));
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
  if (!pdfjsLib) {
    setDbState('error', 'PDF.js gagal dimuat');
    setStatus('Library PDF.js gagal dimuat. Ganti file index.html dan app.js dengan versi fix ini.');
    return;
  }
  try {
    const hadCache = loadCache();
    await refreshDirectory(!hadCache);
  } catch (error) {
    console.error(error);
  }
}
init();
