
const APPS_SCRIPT_URL='https://script.google.com/macros/s/AKfycby6jTm5xQmcQAUjdaubsOWOn7Xws4UWjV9uWbOExlEQArCSN6hubMt3U128QjmlWZP0Ow/exec';
const ROOT_FOLDER_ID='1NZfDp_9SU50OVDJXLuTcGZNj5JvSdpdX';
const CACHE_KEY='payslip_bip_list_cache_v16';
const SETTINGS_KEY='payslip_bip_ui_settings_v1';
const pdfjsLib=globalThis.pdfjsLib||window.pdfjsLib;
if(pdfjsLib) pdfjsLib.GlobalWorkerOptions.workerSrc='https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

const COLOR_BY_MONTH={'01':'#2563eb','02':'#9333ea','03':'#0f766e','04':'#ea580c','05':'#dc2626','06':'#0891b2','07':'#7c3aed','08':'#15803d','09':'#b45309','10':'#be185d','11':'#1d4ed8','12':'#047857'};

const state={
  years:[], periods:[], records:[], filteredRecords:[],
  selectedYear:'', selectedPeriod:'', selectedRecord:null,
  previewBytes:null, pdfDoc:null, zoom:1.2, searchIndex:-1,
  isRefreshing:false, printFrame:null, syncPoller:null, syncRunning:false
};

const el={
  yearSelect:document.getElementById('yearSelect'),
  periodSelect:document.getElementById('periodSelect'),
  searchInput:document.getElementById('searchInput'),
  searchResults:document.getElementById('searchResults'),
  refreshBtn:document.getElementById('refreshBtn'),
  syncBtn:document.getElementById('syncBtn'),
  printBtn:document.getElementById('printBtn'),
  pdfContainer:document.getElementById('pdfContainer'),
  statusText:document.getElementById('statusText'),
  fileNameText:document.getElementById('fileNameText'),
  employeeIdText:document.getElementById('employeeIdText'),
  employeeNameText:document.getElementById('employeeNameText'),
  pageInfo:document.getElementById('pageInfo'),
  dbStatusWrap:document.getElementById('dbStatusWrap'),
  dbStatusText:document.getElementById('dbStatusText'),
  brandLogo:document.getElementById('brandLogo'),
  settingsBtn:document.getElementById('settingsBtn'),
  settingsModal:document.getElementById('settingsModal'),
  titleInput:document.getElementById('titleInput'),
  logoInput:document.getElementById('logoInput'),
  saveSettingsBtn:document.getElementById('saveSettingsBtn'),
  resetLogoBtn:document.getElementById('resetLogoBtn')
};

function setStatus(t){ if(el.statusText) el.statusText.textContent=t; }
function setDbState(s,t){ if(el.dbStatusWrap) el.dbStatusWrap.dataset.state=s; if(el.dbStatusText) el.dbStatusText.textContent=t; }
function normalize(v){ return String(v||'').toLowerCase().trim(); }
function digitsOnly(v){ return String(v||'').replace(/\D/g,''); }
function escapeHtml(v){ return String(v??'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/\"/g,'&quot;').replace(/'/g,'&#39;'); }

function apiUrl(action,params={}){
  const url=new URL(APPS_SCRIPT_URL);
  url.searchParams.set('action',action);
  url.searchParams.set('rootFolderId',ROOT_FOLDER_ID);
  Object.entries(params).forEach(([k,v])=>{ if(v!==undefined&&v!==null&&v!=='') url.searchParams.set(k,v); });
  return url.toString();
}
async function apiGet(action,params={},timeoutMs=10000){
  const controller=new AbortController();
  const timer=setTimeout(()=>controller.abort(),timeoutMs);
  try{
    const res=await fetch(apiUrl(action,params),{cache:'no-store',redirect:'follow',signal:controller.signal});
    const text=await res.text();
    if(!res.ok) throw new Error(`HTTP ${res.status}`);
    let json;
    try{ json=JSON.parse(text); }catch{ throw new Error('Respons Apps Script bukan JSON.'); }
    if(json.ok===false) throw new Error(json.message||'Request gagal');
    return json;
  }catch(error){
    if(error.name==='AbortError') throw new Error('Request timeout. Backend terlalu lama merespons.');
    throw error;
  }finally{ clearTimeout(timer); }
}
function parseBase64(base64){
  const b=atob(base64); const bytes=new Uint8Array(b.length);
  for(let i=0;i<b.length;i++) bytes[i]=b.charCodeAt(i);
  return bytes;
}

function ensureProgressUi(){
  let wrap=document.getElementById('syncProgressWrap');
  if(wrap) return wrap;
  const searchCard=document.querySelector('.search-card');
  if(!searchCard) return null;

  wrap=document.createElement('div');
  wrap.id='syncProgressWrap';
  wrap.style.display='none';
  wrap.style.marginTop='12px';
  wrap.innerHTML=`
    <div id="syncProgressLabel" style="font-size:12px;font-weight:800;color:#334155;margin-bottom:6px;">Sinkronisasi data</div>
    <div style="height:10px;background:#e2e8f0;border-radius:999px;overflow:hidden;border:1px solid #cbd5e1;">
      <div id="syncProgressBar" style="width:0%;height:100%;background:linear-gradient(90deg,#22c55e,#16a34a);transition:width .25s ease;"></div>
    </div>
    <div style="display:flex;justify-content:space-between;gap:12px;margin-top:6px;font-size:12px;color:#475569;">
      <span id="syncProgressText">0%</span>
      <span id="syncProgressMeta">Menunggu...</span>
    </div>
  `;
  searchCard.appendChild(wrap);
  return wrap;
}
function updateProgressUi(sync){
  const wrap=ensureProgressUi();
  if(!wrap) return;
  const label=document.getElementById('syncProgressLabel');
  const bar=document.getElementById('syncProgressBar');
  const text=document.getElementById('syncProgressText');
  const meta=document.getElementById('syncProgressMeta');

  if(!sync || (!sync.running && !sync.percent && !sync.message && !sync.lastSuccessAt)){
    wrap.style.display='none';
    return;
  }

  wrap.style.display='block';
  const pct=Math.max(0,Math.min(100,Number(sync.percent||0)));
  if(bar) bar.style.width=`${pct}%`;
  if(text) text.textContent=`${pct}%`;
  const detail=[sync.currentYear||'', sync.currentPeriod||''].filter(Boolean).join(' • ');
  const counter=(sync.total||sync.processed) ? `${sync.processed||0}/${sync.total||0}` : '';
  if(meta) meta.textContent=[detail,counter].filter(Boolean).join(' • ') || (sync.message||'Sinkronisasi data...');
  if(label) label.textContent=sync.message || (sync.running ? 'Sinkronisasi data...' : 'Sinkronisasi selesai');
}
function startProgressPolling(){
  stopProgressPolling();
  state.syncRunning=true;
  state.syncPoller=setInterval(async()=>{
    try{
      const status=await apiGet('status',{},6000);
      updateProgressUi(status);
      if(!status.running){
        stopProgressPolling();
        await fetchYearData(state.selectedYear||'', false, true);
        if(el.searchInput?.value) localSearch(el.searchInput.value);
      }
    }catch(error){
      console.error(error);
    }
  }, 1200);
}
function stopProgressPolling(){
  state.syncRunning=false;
  if(state.syncPoller){
    clearInterval(state.syncPoller);
    state.syncPoller=null;
  }
}
async function startSyncFlow(forceYear=''){
  setDbState('loading','Sinkronisasi database...');
  setStatus('Memulai sinkronisasi data...');
  updateProgressUi({running:true,percent:1,message:'Menyiapkan sinkronisasi...',currentYear:forceYear||'',currentPeriod:'',processed:0,total:0});
  const json=await apiGet('start_sync',{year:forceYear},300000);
  updateProgressUi(json.sync||{running:true,percent:1,message:'Sinkronisasi dimulai'});
  startProgressPolling();
}

function extractYear(text){ const m=String(text||'').match(/\b(20\d{2})\b/); return m?m[1]:''; }
function extractSrId(record){
  const candidates=[record?.employeeId,record?.employeeCode,record?.srId,record?.fileName].filter(Boolean);
  for(const value of candidates){
    const m=String(value).toUpperCase().match(/\bSR[0-9]{5,}\b/);
    if(m) return m[0];
  }
  return '';
}
function extractDisplayName(record){
  const explicit=String(record?.employeeName||record?.name||'').trim();
  if(explicit) return explicit;
  const fileName=String(record?.fileName||'');
  const m=fileName.match(/^[^-]+-(SR[0-9]{5,})-(.+?)-ID[0-9]{8}-Payslip\.pdf$/i);
  if(m&&m[2]) return m[2].replace(/[-_]+/g,' ').trim();
  return '';
}
function getDisplayId(r){ return extractSrId(r)||String(r?.employeeId||'').trim(); }
function getDisplayName(r){ return extractDisplayName(r)||String(r?.employeeName||'').trim(); }

function saveCache(){
  try{
    localStorage.setItem(CACHE_KEY,JSON.stringify({
      at:Date.now(), years:state.years, selectedYear:state.selectedYear,
      periods:state.periods, records:state.records
    }));
  }catch{}
}
function loadCache(){
  try{
    const raw=localStorage.getItem(CACHE_KEY);
    if(!raw) return false;
    const parsed=JSON.parse(raw);
    state.years=parsed.years||[];
    state.selectedYear=parsed.selectedYear||'';
    state.periods=parsed.periods||[];
    state.records=parsed.records||[];
    normalizeState();
    renderSelectors();
    applyPeriodColors();
    setDbState('connected','Terhubung (cache cepat)');
    setStatus('Daftar file dimuat instan dari cache');
    return !!state.records.length;
  }catch{ return false; }
}
function normalizeState(){
  state.records=(state.records||[]).map(r=>({...r,year:String(r.year||extractYear(r.periodLabel)||extractYear(r.fileName)||'')}));
  if(!state.years.length){
    state.years=[...new Set(state.records.map(r=>r.year).filter(Boolean))].sort();
  }
  if(!state.selectedYear||!state.years.includes(state.selectedYear)){
    state.selectedYear=state.years[state.years.length-1]||'';
  }
  state.periods=[...new Set(
    state.records
      .filter(r=>!state.selectedYear||String(r.year||'')===String(state.selectedYear))
      .map(r=>r.periodLabel)
      .filter(Boolean)
  )].sort();
  if(!state.periods.includes(state.selectedPeriod)){
    state.selectedPeriod=state.periods[0]||'';
  }
}
function renderSelectors(){
  if(el.yearSelect){
    el.yearSelect.innerHTML=state.years.map(y=>`<option value="${escapeHtml(y)}">${escapeHtml(y)}</option>`).join('');
    el.yearSelect.value=state.selectedYear||'';
  }
  if(el.periodSelect){
    el.periodSelect.innerHTML=state.periods.map(p=>`<option value="${escapeHtml(p)}">${escapeHtml(p)}</option>`).join('');
    el.periodSelect.value=state.selectedPeriod||'';
  }
}
function getPeriodColor(){
  const m=String(state.selectedPeriod||'').match(/Payslip-(\d{2})-\d{4}/i);
  const month=m?m[1]:'';
  return COLOR_BY_MONTH[month]||'#2563eb';
}
function applyPeriodColors(){
  const color=getPeriodColor();
  if(el.periodSelect){
    el.periodSelect.style.color=color;
    el.periodSelect.style.fontWeight='800';
    el.periodSelect.style.borderColor=color+'55';
    el.periodSelect.style.boxShadow=`0 0 0 3px ${color}12`;
    el.periodSelect.style.fontFamily='"Inter", sans-serif';
  }
  if(el.yearSelect){
    el.yearSelect.style.color=color;
    el.yearSelect.style.fontWeight='800';
    el.yearSelect.style.borderColor=color+'55';
    el.yearSelect.style.boxShadow=`0 0 0 3px ${color}12`;
    el.yearSelect.style.fontFamily='"Inter", sans-serif';
  }
  const yearLabel=document.querySelector('label[for="yearSelect"]');
  const periodLabel=document.querySelector('label[for="periodSelect"]');
  if(yearLabel){ yearLabel.style.color=color; yearLabel.style.fontWeight='800'; }
  if(periodLabel){ periodLabel.style.color=color; periodLabel.style.fontWeight='800'; }
}
function hideSubtitleAndButtons(){
  const subtitle=document.querySelector('.toolbar-subtitle');
  if(subtitle) subtitle.style.display='none';
  if(el.refreshBtn){
    Object.assign(el.refreshBtn.style,{background:'linear-gradient(135deg,#ffffff,#eef4ff)',border:'1px solid #cfe0ff',color:'#0f172a',boxShadow:'0 10px 22px rgba(37,99,235,.10)',fontWeight:'800'});
    el.refreshBtn.textContent='✦ Refresh';
  }
  if(el.syncBtn) el.syncBtn.style.display='none';
}
function setRefreshBusy(isBusy){
  state.isRefreshing=isBusy;
  if(!el.refreshBtn) return;
  el.refreshBtn.disabled=isBusy;
  el.refreshBtn.style.opacity=isBusy?'0.7':'1';
  el.refreshBtn.textContent=isBusy?'⏳ Sync...' : '✦ Refresh';
}

function recordMatchesQuery(record,rawQuery){
  const q=normalize(rawQuery);
  const qDigits=digitsOnly(rawQuery);
  if(!q) return false;
  const displayId=getDisplayId(record);
  const displayName=getDisplayName(record);
  const employeeId=String(record?.employeeId||'');
  const fileName=String(record?.fileName||'');
  const textHaystacks=[displayId,displayName,employeeId,fileName].map(normalize);
  if(textHaystacks.some(v=>v.includes(q))) return true;
  if(qDigits){
    const numericHaystacks=[displayId,employeeId,fileName].map(digitsOnly).filter(Boolean);
    if(numericHaystacks.some(v=>v.includes(qDigits))) return true;
  }
  return false;
}
function localSearch(rawQuery){
  const q=String(rawQuery||'').trim();
  if(!q){
    state.filteredRecords=[];
    renderSearchResults();
    return;
  }
  const rows=state.records.filter(r=>{
    if(state.selectedYear&&String(r.year||'')!==String(state.selectedYear)) return false;
    if(state.selectedPeriod&&String(r.periodLabel||'')!==String(state.selectedPeriod)) return false;
    return recordMatchesQuery(r,q);
  }).slice(0,20);
  state.filteredRecords=rows;
  state.searchIndex=0;
  renderSearchResults();
}
function renderSearchResults(){
  if(!el.searchResults) return;
  const rows=state.filteredRecords||[];
  if(!rows.length){
    el.searchResults.innerHTML='';
    el.searchResults.classList.add('hidden');
    return;
  }
  if(state.searchIndex<0||state.searchIndex>rows.length-1) state.searchIndex=0;
  el.searchResults.innerHTML=rows.map((r,i)=>`
    <div class="search-item ${i===state.searchIndex?'active':''}" data-index="${i}">
      <div class="search-item-title">${escapeHtml(getDisplayId(r)||'-')} - ${escapeHtml(getDisplayName(r)||'-')}</div>
      <div class="search-item-sub">${escapeHtml(r.periodLabel||'-')} • ${escapeHtml(r.year||'-')}</div>
    </div>
  `).join('');
  el.searchResults.classList.remove('hidden');
}
function hideSearch(){ el.searchResults?.classList.add('hidden'); }

function syncMeta(){
  const r=state.selectedRecord;
  if(el.fileNameText) el.fileNameText.textContent=r?.fileName||'-';
  if(el.employeeIdText) el.employeeIdText.textContent=getDisplayId(r)||'-';
  if(el.employeeNameText) el.employeeNameText.textContent=getDisplayName(r)||'-';
}

async function loadPdfBytes(bytes){
  state.previewBytes=bytes;
  const doc=await pdfjsLib.getDocument({data:bytes}).promise;
  state.pdfDoc=doc;
  const page=await doc.getPage(1);
  const cssViewport=page.getViewport({scale:state.zoom});
  const renderViewport=page.getViewport({scale:state.zoom*Math.min(2.2,Math.max(1.5,window.devicePixelRatio||1.5))});
  const wrap=document.createElement('div');
  wrap.className='page-wrap';
  wrap.style.width=`${cssViewport.width}px`;
  wrap.style.height=`${cssViewport.height}px`;
  const canvas=document.createElement('canvas');
  canvas.className='page-canvas';
  canvas.width=Math.floor(renderViewport.width);
  canvas.height=Math.floor(renderViewport.height);
  canvas.style.width=`${cssViewport.width}px`;
  canvas.style.height=`${cssViewport.height}px`;
  const ctx=canvas.getContext('2d',{alpha:false});
  await page.render({canvasContext:ctx,viewport:renderViewport}).promise;
  if(el.pdfContainer){
    el.pdfContainer.innerHTML='';
    wrap.appendChild(canvas);
    el.pdfContainer.appendChild(wrap);
  }
  if(el.pageInfo) el.pageInfo.textContent=`Page 1 / ${doc.numPages}`;
}
async function loadRecord(record){
  try{
    state.selectedRecord=record;
    syncMeta();
    setStatus('Mengambil PDF asli...');
    const json=await apiGet('file',{fileId:record.fileId},12000);
    await loadPdfBytes(parseBase64(json.base64));
    setStatus('Preview PDF asli siap');
  }catch(error){
    console.error(error);
    setStatus(`Gagal memuat PDF: ${error.message}`);
  }
}
async function fetchYearData(year,force=false,quiet=false){
  if(!quiet){
    setDbState('loading','Menghubungkan database...');
    setStatus(force?`Menyegarkan ${year||'semua tahun'}...`:`Memuat data ${year||'terbaru'}...`);
  }
  if(force) setRefreshBusy(true);
  try{
    const json=await apiGet('list',{year:year,forceRefresh:force?'1':''},10000);
    if(json.needsSync){
      if(!state.syncRunning){
        await startSyncFlow(year||'');
      }
      return;
    }
    state.years=json.years||[];
    state.selectedYear=json.selectedYear||year||state.selectedYear;
    state.records=json.files||[];
    normalizeState();
    renderSelectors();
    applyPeriodColors();
    saveCache();
    setDbState('connected',force?'Terhubung (index baru)':'Terhubung');
    setStatus(`Daftar ${state.selectedYear||'-'} siap • ${state.records.length} file`);
    if(json.sync) updateProgressUi(json.sync);
  }catch(error){
    console.error(error);
    setStatus(`Error: ${error.message}`);
    throw error;
  }finally{
    if(force) setRefreshBusy(false);
  }
}
async function ensurePdfReady(){
  if(state.previewBytes&&state.previewBytes.length) return state.previewBytes;
  if(state.selectedRecord?.fileId){
    const json=await apiGet('file',{fileId:state.selectedRecord.fileId},12000);
    state.previewBytes=parseBase64(json.base64);
    return state.previewBytes;
  }
  return null;
}
function cleanupPrintFrame(){
  if(state.printFrame&&state.printFrame.parentNode) state.printFrame.parentNode.removeChild(state.printFrame);
  state.printFrame=null;
}
async function printCurrent(){
  try{
    setStatus('Menyiapkan cetak...');
    const bytes=await ensurePdfReady();
    if(!bytes||!bytes.length){
      setStatus('PDF belum siap untuk dicetak');
      return;
    }
    cleanupPrintFrame();
    const blob=new Blob([bytes],{type:'application/pdf'});
    const blobUrl=URL.createObjectURL(blob);
    const iframe=document.createElement('iframe');
    Object.assign(iframe.style,{position:'fixed',right:'0',bottom:'0',width:'1px',height:'1px',border:'0',opacity:'0'});
    state.printFrame=iframe;
    iframe.onload=()=>{
      const win=iframe.contentWindow;
      if(win){ setTimeout(()=>{ try{ win.focus(); win.print(); setStatus('Dialog print dibuka'); }catch{} },700); }
      setTimeout(()=>{ URL.revokeObjectURL(blobUrl); cleanupPrintFrame(); },15000);
    };
    document.body.appendChild(iframe);
    iframe.src=blobUrl;
  }catch(error){
    console.error(error);
    setStatus(`Gagal print: ${error.message}`);
  }
}
async function pickRecord(index){
  const r=state.filteredRecords[index];
  if(!r) return;
  state.selectedPeriod=r.periodLabel||state.selectedPeriod;
  renderSelectors();
  applyPeriodColors();
  if(el.searchInput) el.searchInput.value=`${getDisplayId(r)} - ${getDisplayName(r)}`;
  hideSearch();
  await loadRecord(r);
}

function bind(){
  el.yearSelect?.addEventListener('change',async()=>{
    state.selectedYear=el.yearSelect.value;
    state.selectedPeriod='';
    hideSearch();
    await fetchYearData(state.selectedYear,false);
    localSearch(el.searchInput?.value||'');
  });
  el.periodSelect?.addEventListener('change',()=>{
    state.selectedPeriod=el.periodSelect.value;
    applyPeriodColors();
    localSearch(el.searchInput?.value||'');
  });
  let searchTimer=null;
  el.searchInput?.addEventListener('input',()=>{
    clearTimeout(searchTimer);
    searchTimer=setTimeout(()=>localSearch(el.searchInput.value),80);
  });
  el.searchInput?.addEventListener('focus',()=>localSearch(el.searchInput.value));
  el.searchInput?.addEventListener('keydown',async(e)=>{
    if(el.searchResults?.classList.contains('hidden')) return;
    const max=state.filteredRecords.length-1;
    if(e.key==='ArrowDown'){ e.preventDefault(); state.searchIndex=Math.min(max,state.searchIndex+1); renderSearchResults(); }
    else if(e.key==='ArrowUp'){ e.preventDefault(); state.searchIndex=Math.max(0,state.searchIndex-1); renderSearchResults(); }
    else if(e.key==='Enter'){ e.preventDefault(); if(state.searchIndex>=0) await pickRecord(state.searchIndex); }
    else if(e.key==='Escape'){ hideSearch(); }
  });
  el.searchResults?.addEventListener('mousedown',async(e)=>{
    const item=e.target.closest('.search-item');
    if(!item) return;
    await pickRecord(Number(item.dataset.index));
  });
  document.addEventListener('click',(e)=>{ if(!e.target.closest('.search-field')) hideSearch(); });
  el.refreshBtn?.addEventListener('click',async()=>{
    if(state.isRefreshing||state.syncRunning) return;
    try{
      setRefreshBusy(true);
    }catch{}
    await startSyncFlow(state.selectedYear||'');
  });
  el.printBtn?.addEventListener('click',printCurrent);
  el.settingsBtn?.addEventListener('click',()=>el.settingsModal?.classList.remove('hidden'));
  document.querySelectorAll('[data-close-settings]').forEach(node=>node.addEventListener('click',()=>el.settingsModal?.classList.add('hidden')));
  el.saveSettingsBtn?.addEventListener('click',()=>{
    const persist=()=>{
      const current={title:(el.titleInput?.value||'Payslip BIP').trim()||'Payslip BIP',logoDataUrl:el.brandLogo?.src||'./hoplun.jpg'};
      localStorage.setItem(SETTINGS_KEY,JSON.stringify(current));
      document.title=current.title;
      const titleNode=document.querySelector('.brand-left h1');
      if(titleNode) titleNode.textContent=current.title;
      el.settingsModal?.classList.add('hidden');
    };
    const file=el.logoInput?.files?.[0];
    if(!file) return persist();
    const fr=new FileReader();
    fr.onload=()=>{ if(el.brandLogo) el.brandLogo.src=fr.result; persist(); };
    fr.readAsDataURL(file);
  });
  el.resetLogoBtn?.addEventListener('click',()=>{
    if(el.brandLogo) el.brandLogo.src='./hoplun.jpg';
    if(el.logoInput) el.logoInput.value='';
  });
}

async function init(){
  bind();
  hideSubtitleAndButtons();
  ensureProgressUi();
  loadCache();
  const rawSettings=localStorage.getItem(SETTINGS_KEY);
  if(rawSettings){
    try{
      const settings=JSON.parse(rawSettings);
      const title=settings.title||'Payslip BIP';
      document.title=title;
      const titleNode=document.querySelector('.brand-left h1');
      if(titleNode) titleNode.textContent=title;
      if(el.titleInput) el.titleInput.value=title;
      if(settings.logoDataUrl&&el.brandLogo) el.brandLogo.src=settings.logoDataUrl;
    }catch{}
  }
  try{
    const status=await apiGet('status',{},6000);
    updateProgressUi(status);
    if(status.running){
      setDbState('loading','Sinkronisasi database...');
      setStatus('Sinkronisasi sedang berjalan...');
      startProgressPolling();
      return;
    }
  }catch(error){
    console.error(error);
  }

  try{
    await fetchYearData(state.selectedYear||'',false);
    const first=state.records.find(r=>!state.selectedPeriod||r.periodLabel===state.selectedPeriod)||state.records[0];
    if(first) await loadRecord(first);
    if(!state.records.length){
      await startSyncFlow(state.selectedYear||'');
    }
  }catch(error){
    console.error(error);
    await startSyncFlow(state.selectedYear||'');
  }
}
init();
