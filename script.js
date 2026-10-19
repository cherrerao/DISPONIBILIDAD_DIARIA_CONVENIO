// ═══════════════════════════════════════════════════════════════
// CONFIG — mismo patrón que "Productos No Atendidos"
// ═══════════════════════════════════════════════════════════════
const FILE_ID  = '1MbUNufl4XwqNa5AS7RhNYTypZw-4giwV'; // ID del archivo/Sheet en Drive
const TABS = {
  detalle:   'DISPO_NEW_JUN_2026',              // detalle fila a fila (red, EESS, medicamento)
  red:       'DISPO_RESUMEN_RED_JUN_2026',       // disponibilidad ya calculada por red
  estab:     'DISPO_RESUMEN_ESTABLECIMIENTO_J',  // disponibilidad ya calculada por EESS
  promedio:  'DISPO_PROMEDIO_ESTABLECIMIENTO_'   // promedio global ya calculado
};

// ═══════════════════════════════════════════════════════════════
// FETCH — CSV export + fallback gviz/proxy, y respaldo binario xlsx (SheetJS)
// ═══════════════════════════════════════════════════════════════
function parseCSVText(csv) {
  const lines = csv.split('\n').filter(l => l.trim());
  if (lines.length < 2) throw new Error('La hoja está vacía o sin datos.');
  function parseLine(line) {
    const result = []; let current = '', inQ = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (ch === '"') { inQ = !inQ; continue; }
      if (ch === ',' && !inQ) { result.push(current.trim()); current = ''; continue; }
      current += ch;
    }
    result.push(current.trim()); return result;
  }
  const h = parseLine(lines[0]).map(x => x.replace(/^"|"$/g,'').trim());
  const idx = {};
  const COL_MAP = {
    red:    ['RED','red'],
    cod:    ['codigo_pre','COD_PRE','cod_pre'],
    est:    ['ESTABLEC','estab','establecimiento'],
    codmed: ['codigo_med','COD_MED','cod_med'],
    med:    ['nombre_med','MEDICAMENTO','producto'],
    stk:    ['stk_sismed','STOCK','stock'],
    con:    ['con_sismed','CONSUMO','consumo'],
    cpa:    ['CPA','cpa'],
    disp:   ['disp','DISP','meses_disponibilidad'],
    ind:    ['indicador_nuevo','INDICADOR','indicador']
  };
  for (const [key, candidates] of Object.entries(COL_MAP)) {
    for (const c of candidates) { const i = h.findIndex(x => x.toLowerCase() === c.toLowerCase()); if (i !== -1) { idx[key] = i; break; } }
  }
  const rows = [];
  for (let i = 1; i < lines.length; i++) {
    const cols = parseLine(lines[i]);
    if (!cols.length) continue;
    const get = key => idx[key] !== undefined ? (cols[idx[key]] || '').replace(/^"|"$/g,'') : '';
    const getNum = key => parseFloat(get(key).replace(/,/g,'.')) || 0;
    const red = get('red'), est = get('est'), med = get('med'), ind = (get('ind') || 'SIN INFORMACION').toUpperCase();
    if (!red || !est || !med) continue;
    rows.push({ red, cod:get('cod'), est, codmed:get('codmed'), med, stk:getNum('stk'), con:getNum('con'), cpa:getNum('cpa'), disp:getNum('disp'), ind });
  }
  return rows;
}

// Parser genérico: devuelve filas como objetos {ENCABEZADO: valor}
function parseGenericCSV(csv) {
  const lines = csv.split('\n').filter(l => l.trim());
  if (lines.length < 2) throw new Error('La hoja está vacía o sin datos.');
  function parseLine(line) {
    const result = []; let current = '', inQ = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (ch === '"') { inQ = !inQ; continue; }
      if (ch === ',' && !inQ) { result.push(current.trim()); current = ''; continue; }
      current += ch;
    }
    result.push(current.trim()); return result;
  }
  const headers = parseLine(lines[0]).map(x => x.replace(/^"|"$/g,'').trim());
  const rows = [];
  for (let i = 1; i < lines.length; i++) {
    const cols = parseLine(lines[i]);
    if (!cols.length || cols.every(c => c === '')) continue;
    const obj = {};
    headers.forEach((h, j) => { obj[h] = (cols[j] || '').replace(/^"|"$/g,''); });
    rows.push(obj);
  }
  return rows;
}
function numFrom(obj, key){ const v = obj[key]; if(v===undefined) return 0; return parseFloat(String(v).replace(/,/g,'.')) || 0; }
function strFrom(obj, key){ return obj[key] !== undefined ? String(obj[key]).trim() : ''; }

// Google no siempre filtra por nombre de pestaña en /export (requiere gid numérico) —
// esto valida que la hoja recibida sea realmente la que se pidió, por sus encabezados.
function headerSignature(csv){ return (csv.split('\n')[0] || '').toUpperCase(); }
const TAB_VALIDATORS = {
  detalle:  sig => sig.includes('NOMBRE_MED') || sig.includes('INDICADOR'),
  red:      sig => sig.includes('DISPONIBILIDAD') && sig.includes('NORMOSTOCK') && !sig.includes('ESTABLEC'),
  estab:    sig => sig.includes('DISPONIBILIDAD') && sig.includes('ESTABLEC'),
  promedio: sig => sig.includes('PROMEDIO'),
};
function tabKeyFor(tabName){ return Object.keys(TABS).find(k => TABS[k] === tabName); }
function isValidForTab(tabName, csv){
  const key = tabKeyFor(tabName);
  const validator = key && TAB_VALIDATORS[key];
  if (!validator || !csv) return !!csv;
  return validator(headerSignature(csv));
}

// ── Conexión robusta: timeouts, descarga única del libro como respaldo y caché local ──
const FETCH_TIMEOUT_MS = 9000;
const CACHE_KEY = 'dispo_convenio_cache_v1';
const sleep = ms => new Promise(r => setTimeout(r, ms));
let _csvBroken = false; // si una hoja falla por CSV, las demás van directo al libro completo

async function fetchWithTimeout(url, opts = {}, ms = FETCH_TIMEOUT_MS) {
  const ctrl = new AbortController();
  const t = setTimeout(() => ctrl.abort(), ms);
  try { return await fetch(url, { ...opts, signal: ctrl.signal, cache: 'no-store' }); }
  finally { clearTimeout(t); }
}

async function fetchSheetCSV(tab) {
  const enc = encodeURIComponent(tab);
  const gvizUrl   = `https://docs.google.com/spreadsheets/d/${FILE_ID}/gviz/tq?tqx=out:csv&sheet=${enc}`;
  const directUrl = `https://docs.google.com/spreadsheets/d/${FILE_ID}/export?format=csv&sheet=${enc}`;
  const sources = [
    gvizUrl,
    directUrl,
    `https://corsproxy.io/?${encodeURIComponent(gvizUrl)}`,
    `https://api.allorigins.win/raw?url=${encodeURIComponent(gvizUrl)}`,
  ];
  for (const url of sources) {
    try {
      const res = await fetchWithTimeout(url, { redirect: 'follow' });
      if (!res.ok) continue;
      const csv = await res.text();
      if (csv && csv.length > 5 && !csv.includes('<html') && isValidForTab(tab, csv)) return csv;
    } catch (e) { /* siguiente fuente */ }
  }
  return null; // señal para usar el respaldo del libro completo
}

// Un xlsx real empieza con "PK". Drive a veces devuelve una página HTML de confirmación
// (más de 2 KB) que antes se aceptaba como si fuera el archivo → fallos intermitentes.
function looksLikeXlsx(buf) {
  if (!buf || buf.byteLength < 2000) return false;
  const b = new Uint8Array(buf, 0, 2);
  return b[0] === 0x50 && b[1] === 0x4B;
}

async function fetchXlsxBinary(fileId) {
  const urls = [
    `https://docs.google.com/spreadsheets/d/${fileId}/export?format=xlsx`,
    `https://drive.google.com/uc?export=download&confirm=t&id=${fileId}`,
    `https://drive.usercontent.google.com/download?id=${fileId}&export=download&confirm=t`,
  ];
  const sources = [...urls];
  urls.forEach(u => {
    sources.push(`https://corsproxy.io/?${encodeURIComponent(u)}`);
    sources.push(`https://api.allorigins.win/raw?url=${encodeURIComponent(u)}`);
  });
  for (const url of sources) {
    try {
      const res = await fetchWithTimeout(url, { redirect: 'follow' }, 25000);
      if (!res.ok) continue;
      const buf = await res.arrayBuffer();
      if (looksLikeXlsx(buf)) return buf;
    } catch (e) { /* siguiente fuente */ }
  }
  throw new Error('No se pudo conectar con Google Drive. Verifica que el archivo esté compartido como "Cualquiera con el enlace puede ver".');
}

// Promesa compartida: las 4 hojas esperan UNA sola descarga (antes se descargaba 4 veces en paralelo)
let _wbPromise = null;
function getWorkbookCached() {
  if (!_wbPromise) {
    _wbPromise = fetchXlsxBinary(FILE_ID)
      .then(buf => XLSX.read(buf, { type: 'array' }))
      .catch(e => { _wbPromise = null; throw e; });
  }
  return _wbPromise;
}

async function fetchTabCSV(tab) {
  if (!_csvBroken) {
    const csv = await fetchSheetCSV(tab);
    if (csv) return csv;
    _csvBroken = true;
  }
  const wb = await getWorkbookCached();
  const wsName = wb.SheetNames.find(n => n.toLowerCase() === tab.toLowerCase())
    || wb.SheetNames.find(n => n.toLowerCase().startsWith(tab.toLowerCase()))
    || null;
  if (!wsName) throw new Error(`No se encontró la hoja "${tab}" en el archivo.`);
  const csv2 = XLSX.utils.sheet_to_csv(wb.Sheets[wsName]);
  if (!isValidForTab(tab, csv2)) throw new Error(`La hoja "${tab}" no tiene el formato esperado.`);
  return csv2;
}

function saveCache(csvs) {
  try { localStorage.setItem(CACHE_KEY, JSON.stringify({ ts: Date.now(), csvs })); } catch (e) {}
}
function loadCache() {
  try { const j = localStorage.getItem(CACHE_KEY); return j ? JSON.parse(j) : null; } catch (e) { return null; }
}

async function fetchAllCsv() {
  const [detalleCsv, redCsv, estabCsv, promedioCsv] = await Promise.all([
    fetchTabCSV(TABS.detalle),
    fetchTabCSV(TABS.red),
    fetchTabCSV(TABS.estab),
    fetchTabCSV(TABS.promedio),
  ]);
  return { detalleCsv, redCsv, estabCsv, promedioCsv };
}

function parseAllCsv({ detalleCsv, redCsv, estabCsv, promedioCsv }) {
  const rawRows = parseCSVText(detalleCsv);

  const redRows = parseGenericCSV(redCsv).map(o => ({
    red: strFrom(o,'RED'),
    NORMOSTOCK: numFrom(o,'NORMOSTOCK'), SOBRESTOCK: numFrom(o,'SOBRESTOCK'),
    SIN_ROTACION: numFrom(o,'SIN_ROTACION'), SUBSTOCK: numFrom(o,'SUBSTOCK'),
    DESABASTECIDO: numFrom(o,'DESABASTECIDO'), DISPONIBLES: numFrom(o,'DISPONIBLES'),
    TOTAL: numFrom(o,'TOTAL'), DISPONIBILIDAD: numFrom(o,'DISPONIBILIDAD')
  })).filter(r => r.red);

  const estRows = parseGenericCSV(estabCsv).map(o => ({
    red: strFrom(o,'RED'), cod: strFrom(o,'codigo_pre') || strFrom(o,'COD_PRE'), est: strFrom(o,'ESTABLEC'),
    NORMOSTOCK: numFrom(o,'NORMOSTOCK'), SOBRESTOCK: numFrom(o,'SOBRESTOCK'),
    SIN_ROTACION: numFrom(o,'SIN_ROTACION'), SUBSTOCK: numFrom(o,'SUBSTOCK'),
    DESABASTECIDO: numFrom(o,'DESABASTECIDO'), DISPONIBLES: numFrom(o,'DISPONIBLES'),
    TOTAL: numFrom(o,'TOTAL'), DISPONIBILIDAD: numFrom(o,'DISPONIBILIDAD')
  })).filter(r => r.red && r.est);

  const promedioRows = parseGenericCSV(promedioCsv);
  const promedioKey = Object.keys(promedioRows[0] || {}).find(k => k.toUpperCase().includes('PROMEDIO')) || 'PROMEDIO_DISPONIBILIDAD';
  const promedioDisponibilidad = promedioRows.length ? numFrom(promedioRows[0], promedioKey) : null;

  return { rawRows, redRows, estRows, promedioDisponibilidad };
}

async function fetchAllData() {
  const csvs = await fetchAllCsv();
  const parsed = parseAllCsv(csvs);
  saveCache(csvs); // solo se guarda si todo se descargó y se interpretó bien
  return parsed;
}

// ═══════════════════════════════════════════════════════════════
// AGREGACIÓN — construye RED / EST / PROD / DETAIL_RAW desde el detalle crudo
// ═══════════════════════════════════════════════════════════════
function buildAggregates(rawRows, redRows, estRows) {
  const IND_KEYS = ['NORMOSTOCK','SOBRESTOCK','SIN ROTACION','SUBSTOCK','DESABASTECIDO','SIN INFORMACION'];
  const prodMap = {};
  const reds = [], ests = [], meds = [], inds = [];
  const redIdx = {}, estIdx = {}, medIdx = {}, indIdx = {};
  const rows = [];

  function emptyCounts(){ return {NORMOSTOCK:0,SOBRESTOCK:0,'SIN ROTACION':0,SUBSTOCK:0,DESABASTECIDO:0,'SIN INFORMACION':0}; }

  rawRows.forEach(r => {
    const ind = IND_KEYS.includes(r.ind) ? r.ind : 'SIN INFORMACION';

    if (!(r.med in prodMap)) prodMap[r.med] = emptyCounts();
    prodMap[r.med][ind]++;

    const estKey = r.red + '||' + r.est;
    if (!(r.red in redIdx)) { redIdx[r.red]=reds.length; reds.push(r.red); }
    if (!(estKey in estIdx)) { estIdx[estKey]=ests.length; ests.push(r.est); }
    if (!(r.med in medIdx)) { medIdx[r.med]=meds.length; meds.push(r.med); }
    if (!(ind in indIdx)) { indIdx[ind]=inds.length; inds.push(ind); }
    rows.push([redIdx[r.red], r.cod, estIdx[estKey], r.codmed, medIdx[r.med], r.stk, r.con, r.cpa, r.disp, indIdx[ind]]);
  });

  const PROD = Object.entries(prodMap).map(([med, counts]) => {
    const total = Object.values(counts).reduce((a,b)=>a+b,0);
    return {med, NORMOSTOCK:counts.NORMOSTOCK, SOBRESTOCK:counts.SOBRESTOCK, 'SIN ROTACION':counts['SIN ROTACION'], SUBSTOCK:counts.SUBSTOCK, DESABASTECIDO:counts.DESABASTECIDO, 'SIN INFORMACION':counts['SIN INFORMACION'], total};
  });

  // RED y EST vienen ya calculados de sus propias pestañas — se usan tal cual
  const RED = redRows;
  const EST = estRows;

  const DETAIL_RAW = {reds, ests, meds, inds, rows};
  return {RED, EST, PROD, DETAIL_RAW};
}

// ═══════════════════════════════════════════════════════════════
// CARGA PRINCIPAL
// ═══════════════════════════════════════════════════════════════
let RED=[], EST=[], PROD=[], DETAIL_RAW={reds:[],ests:[],meds:[],inds:[],rows:[]};
let globalTotals={}, grandTotal=0, grandDisponibles=0, globalDisponibilidad=0;

function computeGlobals(promedioDisponibilidad){
  globalTotals = {NORMOSTOCK:0,SOBRESTOCK:0,'SIN ROTACION':0,SUBSTOCK:0,DESABASTECIDO:0,'SIN INFORMACION':0};
  RED.forEach(r=>{
    globalTotals.NORMOSTOCK += r.NORMOSTOCK;
    globalTotals.SOBRESTOCK += r.SOBRESTOCK;
    globalTotals['SIN ROTACION'] += r.SIN_ROTACION;
    globalTotals.SUBSTOCK += r.SUBSTOCK;
    globalTotals.DESABASTECIDO += r.DESABASTECIDO;
  });
  DETAIL_RAW.rows.forEach(r=>{ if(DETAIL_RAW.inds[r[9]]==='SIN INFORMACION') globalTotals['SIN INFORMACION']++; });
  grandTotal = Object.values(globalTotals).reduce((a,b)=>a+b,0);
  grandDisponibles = globalTotals.NORMOSTOCK + globalTotals.SOBRESTOCK;
  // Disponibilidad global oficial = promedio de DISPO_PROMEDIO_ESTABLECIMIENTO_ (ya calculado en el Excel)
  globalDisponibilidad = (promedioDisponibilidad !== null && promedioDisponibilidad !== undefined)
    ? promedioDisponibilidad
    : (grandTotal ? (grandDisponibles/grandTotal*100) : 0);
}

function setLiveBadge(state, text){
  const b = document.getElementById('liveBadge');
  const colors = {ok:['rgba(34,163,90,.15)','#5fd996'], err:['rgba(224,72,63,.18)','#f28b85'], loading:['rgba(224,151,42,.18)','#f2c185']};
  const [bg, fg] = colors[state] || colors.loading;
  b.style.background = bg; b.style.color = fg;
  b.querySelector('.foot-dot').style.background = fg;
  b.innerHTML = `<span class="foot-dot" style="background:${fg}"></span> ${text}`;
}

function applyData({ rawRows, redRows, estRows, promedioDisponibilidad }) {
  if (!rawRows.length) throw new Error('La hoja de detalle no devolvió registros.');
  if (!redRows.length) throw new Error('La hoja de resumen por red no devolvió registros.');
  if (!estRows.length) throw new Error('La hoja de resumen por establecimiento no devolvió registros.');
  const agg = buildAggregates(rawRows, redRows, estRows);
  RAW = rawRows;
  RED = agg.RED; EST = agg.EST; PROD = agg.PROD; DETAIL_RAW = agg.DETAIL_RAW;
  computeGlobals(promedioDisponibilidad);
  document.getElementById('loadingPanel').style.display = 'none';
  document.getElementById('dashViews').style.display = 'block';
  render();
}

let _loading = false;
async function loadAndRender(){
  if (_loading) return;
  _loading = true;
  document.getElementById('loadingPanel').style.display = 'flex';
  document.getElementById('errorPanel').style.display = 'none';
  document.getElementById('dashViews').style.display = 'none';
  setLiveBadge('loading', 'CONECTANDO…');
  _wbPromise = null; _csvBroken = false; // limpia el estado de conexión en cada actualización
  try {
    let parsed = null, lastErr = null;
    for (let attempt = 0; attempt < 2 && !parsed; attempt++) {
      try { parsed = await fetchAllData(); }
      catch (e) { lastErr = e; _wbPromise = null; _csvBroken = false; if (attempt === 0) await sleep(1500); }
    }
    if (!parsed) throw lastErr;
    applyData(parsed);
    setLiveBadge('ok', 'EN VIVO');
    document.getElementById('lastUpdated').textContent = 'Actualizado ' + new Date().toLocaleTimeString('es-PE',{hour:'2-digit',minute:'2-digit'});
  } catch(err){
    // Sin conexión: se muestra la última copia buena guardada en este navegador
    const cached = loadCache();
    if (cached) {
      try {
        applyData(parseAllCsv(cached.csvs));
        setLiveBadge('err', 'DATOS EN CACHÉ');
        document.getElementById('lastUpdated').textContent = 'Sin conexión · datos del ' + new Date(cached.ts).toLocaleString('es-PE',{day:'2-digit',month:'2-digit',hour:'2-digit',minute:'2-digit'}) + ' · pulsa ↺ para reintentar';
        return;
      } catch (e2) { /* si la caché tampoco sirve, se muestra el error */ }
    }
    document.getElementById('loadingPanel').style.display = 'none';
    setLiveBadge('err', 'SIN CONEXIÓN');
    document.getElementById('errorPanel').style.display = 'block';
    document.getElementById('errorPanel').innerHTML = `
      <div class="error-panel">
        <p>No se pudieron cargar los datos</p>
        <small>${err && err.message ? err.message : 'Error de conexión'}</small>
        <small>Verifica que el archivo esté compartido como "Cualquiera con el enlace puede ver".</small>
        <button onclick="loadAndRender()">↺ Reintentar</button>
      </div>`;
  } finally {
    _loading = false;
  }
}
document.getElementById('refreshBtn').addEventListener('click', loadAndRender);

const IND_COLORS = {
  'NORMOSTOCK': '#22a35a',
  'SOBRESTOCK': '#3b7ddd',
  'SIN ROTACION': '#8b6fd6',
  'SUBSTOCK': '#e0972a',
  'DESABASTECIDO': '#e0483f',
  'SIN INFORMACION': '#8a93a6'
};
const IND_ORDER = ['NORMOSTOCK','SOBRESTOCK','SIN ROTACION','SUBSTOCK','DESABASTECIDO','SIN INFORMACION'];
const IND_LABELS = {
  'NORMOSTOCK':'Normostock','SOBRESTOCK':'Sobrestock','SIN ROTACION':'Sin rotación',
  'SUBSTOCK':'Substock','DESABASTECIDO':'Desabastecido','SIN INFORMACION':'Sin información'
};

function fmt(n, d=1){ if(n===null||n===undefined||isNaN(n)) return '–'; return Number(n).toLocaleString('es-PE',{minimumFractionDigits:d,maximumFractionDigits:d}); }
function fmtInt(n){ return Number(n).toLocaleString('es-PE'); }

function badgeClass(disp){
  if(disp<=80) return 'b-red';
  if(disp<=90) return 'b-orange';
  return 'b-green';
}
function textColorForDisp(disp){
  if(disp<=80) return '#9c2a23';
  if(disp<=90) return '#8a5b12';
  return '#166534';
}

/* ---------------- Navigation ---------------- */
const pageMeta = {
  resumen:{title:'Resumen general', desc:'Disponibilidad de medicamentos e insumos — Red del Callao'},
  red:{title:'Por Red', desc:'Comparativo de disponibilidad entre las redes de salud del Callao'},
  estab:{title:'Por Establecimiento', desc:'Situación de stock por EESS, 47 establecimientos evaluados'},
  alertas:{title:'Ranking & Alertas', desc:'Establecimientos y medicamentos que requieren atención prioritaria'},
  producto:{title:'Por Producto', desc:'Disponibilidad agregada por medicamento, todas las redes'},
  plan:{title:'Plan de abastecimiento', desc:'Productos y cantidades a distribuir por establecimiento para alcanzar la meta de disponibilidad'},
  detalle:{title:'Detalle', desc:'Registro completo de stock, consumo y disponibilidad por ítem'}
};
document.getElementById('nav').addEventListener('click', (e)=>{
  const item = e.target.closest('.nav-item');
  if(!item) return;
  document.querySelectorAll('.nav-item').forEach(n=>n.classList.remove('active'));
  item.classList.add('active');
  const v = item.dataset.view;
  document.querySelectorAll('.view').forEach(vv=>vv.classList.remove('active'));
  document.getElementById('view-'+v).classList.add('active');
  document.getElementById('pageTitle').textContent = pageMeta[v].title;
  document.getElementById('pageDesc').textContent = pageMeta[v].desc;
});

/* ============================================================
   RESUMEN VIEW
   (RED, EST, PROD, DETAIL_RAW, globalTotals, grandTotal,
    grandDisponibles, globalDisponibilidad are populated live
    by loadAndRender() → buildAggregates() / computeGlobals())
   ============================================================ */
function renderAlert(){
  const box = document.getElementById('alertBox');
  let cls='ok', ico='✅', title='Disponibilidad adecuada', msg='';
  if(globalDisponibilidad<=80){cls='crit';ico='🔴';title='Alerta crítica de disponibilidad';}
  else if(globalDisponibilidad<=90){cls='warn';ico='🟠';title='Disponibilidad en observación';}
  msg = `Disponibilidad promedio ${fmt(globalDisponibilidad,1)}% entre los ${fmtInt(EST.length)} establecimientos evaluados (según DISPO_PROMEDIO_ESTABLECIMIENTO_). ${fmtInt(globalTotals.DESABASTECIDO)} ítems desabastecidos y ${fmtInt(globalTotals.SUBSTOCK)} en substock requieren reposición.`;
  box.innerHTML = `<div class="alert ${cls}"><div class="alert-ico">${ico}</div><div><b>${title}</b><span>${msg}</span></div></div>`;
}

function renderKpis(){
  const nEst = EST.length;
  const nRed = RED.length;
  const nMed = PROD.length;
  const html = `
    <div class="kpi"><div class="label">Disponibilidad global</div><div class="value">${fmt(globalDisponibilidad,1)}%</div><div class="sub">promedio entre ${nEst} EESS</div></div>
    <div class="kpi"><div class="label">Establecimientos evaluados</div><div class="value">${nEst}</div><div class="sub">en ${nRed} redes de salud</div></div>
    <div class="kpi"><div class="label">Medicamentos monitoreados</div><div class="value">${nMed}</div><div class="sub">insumos del petitorio SISMED</div></div>
    <div class="kpi"><div class="label">Ítems desabastecidos</div><div class="value" style="color:#e0483f">${fmtInt(globalTotals.DESABASTECIDO)}</div><div class="sub">de ${fmtInt(grandTotal)} ítems evaluados</div></div>
  `;
  document.getElementById('kpiRow').innerHTML = html;
}

let chartDonut, chartRedBar, chartRedStack, chartEstabBar, chartProdBar;

function renderDonut(){
  const ctx = document.getElementById('chartDonut');
  const labels = IND_ORDER.filter(k=>globalTotals[k]>0);
  const vals = labels.map(k=>globalTotals[k]);
  chartDonut = new Chart(ctx, {
    type:'doughnut',
    data:{labels:labels.map(l=>IND_LABELS[l]), datasets:[{data:vals, backgroundColor:labels.map(l=>IND_COLORS[l]), borderWidth:2, borderColor:'#fff'}]},
    options:{plugins:{legend:{display:false}}, cutout:'62%', maintainAspectRatio:false}
  });
  const legend = labels.map(l=>{
    const pct = (globalTotals[l]/grandTotal*100).toFixed(1);
    const cls = ({'NORMOSTOCK':'normo','SOBRESTOCK':'sobre','SIN ROTACION':'sinrot','SUBSTOCK':'subst','DESABASTECIDO':'desab','SIN INFORMACION':'sininfo'})[l];
    return `<span class="chip ${cls}"><span class="dot"></span>${IND_LABELS[l]} ${pct}%</span>`;
  }).join('');
  document.getElementById('donutLegend').innerHTML = legend;
}

function renderRedBar(){
  const ctx = document.getElementById('chartRedBar');
  const sorted = [...RED].sort((a,b)=>b.DISPONIBILIDAD-a.DISPONIBILIDAD);
  chartRedBar = new Chart(ctx,{
    type:'bar',
    data:{labels:sorted.map(r=>r.red), datasets:[{
      data:sorted.map(r=>r.DISPONIBILIDAD),
      backgroundColor:sorted.map(r=>r.DISPONIBILIDAD<=80?'#e0483f':r.DISPONIBILIDAD<=90?'#e0972a':'#22a35a'),
      borderRadius:6, maxBarThickness:46
    }]},
    options:{
      indexAxis:'y', maintainAspectRatio:false,
      plugins:{legend:{display:false}, tooltip:{callbacks:{label:(c)=>c.parsed.x.toFixed(1)+'%'}}},
      scales:{x:{min:0,max:100,ticks:{callback:v=>v+'%'}}}
    }
  });
}

function renderTopCriticoResumen(){
  const sorted = [...EST].sort((a,b)=>a.DISPONIBILIDAD-b.DISPONIBILIDAD).slice(0,8);
  const html = sorted.map((e,i)=>`
    <div class="rank-item">
      <div class="rank-num">${i+1}</div>
      <div class="rank-name"><b>${e.est}</b><span>${e.red}</span></div>
      <div class="rank-val" style="color:${textColorForDisp(e.DISPONIBILIDAD)}">${fmt(e.DISPONIBILIDAD,1)}%</div>
    </div>`).join('');
  document.getElementById('topCriticoResumen').innerHTML = html;
}

function renderTopProdResumen(){
  const sorted = [...PROD].map(p=>({...p, riesgo:(p.DESABASTECIDO||0)+(p['SIN ROTACION']||0)}))
    .sort((a,b)=>b.riesgo-a.riesgo).slice(0,8);
  const html = sorted.map((p,i)=>`
    <div class="rank-item">
      <div class="rank-num">${i+1}</div>
      <div class="rank-name"><b>${p.med}</b><span>${p.DESABASTECIDO} desab. · ${p['SIN ROTACION']} sin rotación · de ${p.total} EESS</span></div>
      <div class="rank-val" style="color:#e0483f">${p.riesgo}</div>
    </div>`).join('');
  document.getElementById('topProdResumen').innerHTML = html;
}

/* ============================================================
   POR RED VIEW
   ============================================================ */
function renderRedCards(){
  const html = RED.map(r=>{
    const total = r.TOTAL;
    const segs = [
      {k:'NORMOSTOCK', v:r.NORMOSTOCK},{k:'SOBRESTOCK', v:r.SOBRESTOCK},
      {k:'SIN ROTACION', v:r.SIN_ROTACION},{k:'SUBSTOCK', v:r.SUBSTOCK},{k:'DESABASTECIDO', v:r.DESABASTECIDO}
    ];
    const stack = segs.map(s=>`<span style="width:${(s.v/total*100).toFixed(2)}%;background:${IND_COLORS[s.k]}"></span>`).join('');
    return `<div class="red-card">
      <h4>${r.red}</h4>
      <div class="disp-big" style="color:${textColorForDisp(r.DISPONIBILIDAD)}">${fmt(r.DISPONIBILIDAD,1)}%</div>
      <div class="disp-lbl">disponibilidad · ${fmtInt(total)} ítems evaluados</div>
      <div class="stack">${stack}</div>
    </div>`;
  }).join('');
  document.getElementById('redCards').innerHTML = html;
}

function renderRedStack(){
  const ctx = document.getElementById('chartRedStack');
  const labels = RED.map(r=>r.red);
  const datasets = [
    {label:'Normostock', data:RED.map(r=>r.NORMOSTOCK), backgroundColor:IND_COLORS.NORMOSTOCK},
    {label:'Sobrestock', data:RED.map(r=>r.SOBRESTOCK), backgroundColor:IND_COLORS.SOBRESTOCK},
    {label:'Sin rotación', data:RED.map(r=>r.SIN_ROTACION), backgroundColor:IND_COLORS['SIN ROTACION']},
    {label:'Substock', data:RED.map(r=>r.SUBSTOCK), backgroundColor:IND_COLORS.SUBSTOCK},
    {label:'Desabastecido', data:RED.map(r=>r.DESABASTECIDO), backgroundColor:IND_COLORS.DESABASTECIDO}
  ];
  chartRedStack = new Chart(ctx,{
    type:'bar',
    data:{labels, datasets},
    options:{
      maintainAspectRatio:false,
      plugins:{legend:{position:'bottom', labels:{boxWidth:10,font:{size:11}}}},
      scales:{x:{stacked:true}, y:{stacked:true}}
    }
  });
}

function renderTableRed(){
  const thead = `<tr><th>Red</th><th>Normostock</th><th>Sobrestock</th><th>Sin rotación</th><th>Substock</th><th>Desabastecido</th><th>Total evaluado</th><th>Disponibilidad</th></tr>`;
  const rows = [...RED].sort((a,b)=>b.DISPONIBILIDAD-a.DISPONIBILIDAD).map(r=>`
    <tr>
      <td><b>${r.red}</b></td>
      <td>${fmtInt(r.NORMOSTOCK)}</td>
      <td>${fmtInt(r.SOBRESTOCK)}</td>
      <td>${fmtInt(r.SIN_ROTACION)}</td>
      <td>${fmtInt(r.SUBSTOCK)}</td>
      <td>${fmtInt(r.DESABASTECIDO)}</td>
      <td>${fmtInt(r.TOTAL)}</td>
      <td><span class="badge ${badgeClass(r.DISPONIBILIDAD)}">${fmt(r.DISPONIBILIDAD,1)}%</span></td>
    </tr>`).join('');
  document.getElementById('tableRed').innerHTML = `<thead>${thead}</thead><tbody>${rows}</tbody>`;
}

/* ============================================================
   POR ESTABLECIMIENTO VIEW
   ============================================================ */
function renderEstabBar(){
  const ctx = document.getElementById('chartEstabBar');
  const sorted = [...EST].sort((a,b)=>a.DISPONIBILIDAD-b.DISPONIBILIDAD).slice(0,15);
  chartEstabBar = new Chart(ctx,{
    type:'bar',
    data:{labels:sorted.map(e=>e.est), datasets:[{
      data:sorted.map(e=>e.DISPONIBILIDAD),
      backgroundColor:sorted.map(e=>e.DISPONIBILIDAD<=80?'#e0483f':e.DISPONIBILIDAD<=90?'#e0972a':'#22a35a'),
      borderRadius:5, maxBarThickness:22
    }]},
    options:{
      indexAxis:'y', maintainAspectRatio:false,
      plugins:{legend:{display:false}, tooltip:{callbacks:{label:c=>c.parsed.x.toFixed(1)+'%'}}},
      scales:{x:{min:0,max:100,ticks:{callback:v=>v+'%'}}, y:{ticks:{autoSkip:false,font:{size:10.5}}}}
    }
  });
}

let estabSort = {key:'DISPONIBILIDAD', dir:1};
function renderEstabTable(){
  const search = document.getElementById('estabSearch').value.toLowerCase();
  const redFilter = document.getElementById('estabRedFilter').value;
  let rows = EST.filter(e=>
    (!redFilter || e.red===redFilter) &&
    (!search || e.est.toLowerCase().includes(search))
  );
  rows.sort((a,b)=>(a[estabSort.key]-b[estabSort.key])*estabSort.dir);
  document.getElementById('estabCount').textContent = rows.length + ' de ' + EST.length + ' EESS';

  const cols = [
    {k:'red', l:'Red'}, {k:'est', l:'Establecimiento'}, {k:'NORMOSTOCK', l:'Normostock'},
    {k:'SOBRESTOCK', l:'Sobrestock'}, {k:'SIN_ROTACION', l:'Sin rotación'}, {k:'SUBSTOCK', l:'Substock'},
    {k:'DESABASTECIDO', l:'Desabastecido'}, {k:'TOTAL', l:'Total'}, {k:'DISPONIBILIDAD', l:'Disponibilidad'}
  ];
  const thead = '<tr>'+cols.map(c=>`<th data-key="${c.k}" class="${estabSort.key===c.k?'sorted':''}">${c.l}</th>`).join('')+'</tr>';
  const tbody = rows.map(e=>`
    <tr>
      <td>${e.red}</td>
      <td><b>${e.est}</b></td>
      <td>${fmtInt(e.NORMOSTOCK)}</td>
      <td>${fmtInt(e.SOBRESTOCK)}</td>
      <td>${fmtInt(e.SIN_ROTACION)}</td>
      <td>${fmtInt(e.SUBSTOCK)}</td>
      <td>${fmtInt(e.DESABASTECIDO)}</td>
      <td>${fmtInt(e.TOTAL)}</td>
      <td><span class="badge ${badgeClass(e.DISPONIBILIDAD)}">${fmt(e.DISPONIBILIDAD,1)}%</span></td>
    </tr>`).join('');
  document.getElementById('tableEstab').innerHTML = `<thead>${thead}</thead><tbody>${tbody}</tbody>`;
  document.querySelectorAll('#tableEstab th').forEach(th=>{
    th.addEventListener('click', ()=>{
      const key = th.dataset.key;
      if(estabSort.key===key) estabSort.dir *= -1; else {estabSort.key=key; estabSort.dir = (key==='DISPONIBILIDAD')?1:-1;}
      renderEstabTable();
    });
  });
}

/* ============================================================
   ALERTAS VIEW
   ============================================================ */
function renderAlertasView(){
  const critEst = EST.filter(e=>e.DISPONIBILIDAD<=80);
  document.getElementById('alCritCount').textContent = critEst.length;
  document.getElementById('alDesab').textContent = fmtInt(globalTotals.DESABASTECIDO);
  document.getElementById('alSubst').textContent = fmtInt(globalTotals.SUBSTOCK);

  const box = document.getElementById('alertBoxAlertas');
  let cls='ok', ico='✅', title='Sin alertas críticas de red';
  if(critEst.length>0){
    cls='crit'; ico='🔴';
    title = `${critEst.length} establecimiento(s) por debajo del 80% de disponibilidad`;
  }
  const worst = [...EST].sort((a,b)=>a.DISPONIBILIDAD-b.DISPONIBILIDAD)[0];
  box.innerHTML = `<div class="alert ${cls}"><div class="alert-ico">${ico}</div><div><b>${title}</b><span>El caso más crítico es ${worst.est} (${worst.red}) con ${fmt(worst.DISPONIBILIDAD,1)}% de disponibilidad.</span></div></div>`;

  const sorted = [...EST].sort((a,b)=>a.DISPONIBILIDAD-b.DISPONIBILIDAD);
  document.getElementById('rankFull').innerHTML = sorted.map((e,i)=>`
    <div class="rank-item">
      <div class="rank-num">${i+1}</div>
      <div class="rank-name"><b>${e.est}</b><span>${e.red} · ${fmtInt(e.DESABASTECIDO)} desab. / ${fmtInt(e.SUBSTOCK)} substock</span></div>
      <div class="rank-val" style="color:${textColorForDisp(e.DISPONIBILIDAD)}">${fmt(e.DISPONIBILIDAD,1)}%</div>
    </div>`).join('');

  const prodSorted = [...PROD].map(p=>({...p, riesgo:(p.DESABASTECIDO||0)}))
    .filter(p=>p.riesgo>0).sort((a,b)=>b.riesgo-a.riesgo).slice(0,25);
  document.getElementById('rankProdCrit').innerHTML = prodSorted.map((p,i)=>`
    <div class="rank-item">
      <div class="rank-num">${i+1}</div>
      <div class="rank-name"><b>${p.med}</b><span>de ${p.total} EESS evaluados</span></div>
      <div class="rank-val" style="color:#e0483f">${p.riesgo}</div>
    </div>`).join('');
}

/* ============================================================
   POR PRODUCTO VIEW
   ============================================================ */
function computeProdRiesgo(p){ return (p.DESABASTECIDO||0)+(p['SIN ROTACION']||0)+(p.SUBSTOCK||0); }

function renderProdBar(){
  const ctx = document.getElementById('chartProdBar');
  const sorted = [...PROD].map(p=>({...p, riesgo:computeProdRiesgo(p)})).sort((a,b)=>b.riesgo-a.riesgo).slice(0,15);
  chartProdBar = new Chart(ctx,{
    type:'bar',
    data:{labels:sorted.map(p=>p.med), datasets:[
      {label:'Desabastecido', data:sorted.map(p=>p.DESABASTECIDO), backgroundColor:IND_COLORS.DESABASTECIDO},
      {label:'Sin rotación', data:sorted.map(p=>p['SIN ROTACION']), backgroundColor:IND_COLORS['SIN ROTACION']},
      {label:'Substock', data:sorted.map(p=>p.SUBSTOCK), backgroundColor:IND_COLORS.SUBSTOCK}
    ]},
    options:{
      indexAxis:'y', maintainAspectRatio:false,
      plugins:{legend:{position:'bottom',labels:{boxWidth:10,font:{size:11}}}},
      scales:{x:{stacked:true}, y:{stacked:true, ticks:{autoSkip:false,font:{size:10.5}}}}
    }
  });
}

let prodSortKey='riesgo';
function renderProdTable(){
  const search = document.getElementById('prodSearch').value.toLowerCase();
  let rows = PROD.map(p=>({...p, riesgo:computeProdRiesgo(p)})).filter(p=>!search || p.med.toLowerCase().includes(search));
  if(prodSortKey==='riesgo') rows.sort((a,b)=>b.riesgo-a.riesgo);
  else if(prodSortKey==='desab') rows.sort((a,b)=>b.DESABASTECIDO-a.DESABASTECIDO);
  else if(prodSortKey==='normo') rows.sort((a,b)=>b.NORMOSTOCK-a.NORMOSTOCK);
  else rows.sort((a,b)=>a.med.localeCompare(b.med));

  document.getElementById('prodCount').textContent = rows.length + ' de ' + PROD.length + ' medicamentos';
  const thead = `<tr><th>Medicamento</th><th>Normostock</th><th>Sobrestock</th><th>Sin rotación</th><th>Substock</th><th>Desabastecido</th><th>Total EESS</th></tr>`;
  const tbody = rows.map(p=>`
    <tr>
      <td><b>${p.med}</b></td>
      <td>${fmtInt(p.NORMOSTOCK)}</td>
      <td>${fmtInt(p.SOBRESTOCK)}</td>
      <td>${fmtInt(p['SIN ROTACION'])}</td>
      <td>${fmtInt(p.SUBSTOCK)}</td>
      <td style="${p.DESABASTECIDO>0?'color:#9c2a23;font-weight:700':''}">${fmtInt(p.DESABASTECIDO)}</td>
      <td>${fmtInt(p.total)}</td>
    </tr>`).join('');
  document.getElementById('tableProd').innerHTML = `<thead>${thead}</thead><tbody>${tbody}</tbody>`;
}

/* ============================================================
   DETALLE VIEW
   ============================================================ */
let detPage = 1;
const DET_PAGE_SIZE = 50;
function detRows(){
  const search = document.getElementById('detSearch').value.toLowerCase();
  const redFilter = document.getElementById('detRedFilter').value;
  const indFilter = document.getElementById('detIndFilter').value;
  return DETAIL_RAW.rows.filter(r=>{
    const red = DETAIL_RAW.reds[r[0]];
    const est = DETAIL_RAW.ests[r[2]];
    const med = DETAIL_RAW.meds[r[4]];
    const ind = DETAIL_RAW.inds[r[9]];
    if(redFilter && red!==redFilter) return false;
    if(indFilter && ind!==indFilter) return false;
    if(search && !(est.toLowerCase().includes(search) || med.toLowerCase().includes(search))) return false;
    return true;
  });
}
function renderDetTable(){
  const rows = detRows();
  document.getElementById('detCount').textContent = fmtInt(rows.length) + ' de ' + fmtInt(DETAIL_RAW.rows.length) + ' registros';
  const totalPages = Math.max(1, Math.ceil(rows.length/DET_PAGE_SIZE));
  if(detPage>totalPages) detPage = totalPages;
  const start = (detPage-1)*DET_PAGE_SIZE;
  const pageRows = rows.slice(start, start+DET_PAGE_SIZE);

  const thead = `<tr><th>Red</th><th>Establecimiento</th><th>Medicamento</th><th>Stock (u.)</th><th>Consumo</th><th>CPA</th><th>Meses disp.</th><th>Clasificación</th></tr>`;
  const tbody = pageRows.map(r=>{
    const red=DETAIL_RAW.reds[r[0]], est=DETAIL_RAW.ests[r[2]], med=DETAIL_RAW.meds[r[4]], ind=DETAIL_RAW.inds[r[9]];
    const chipCls = ({'NORMOSTOCK':'normo','SOBRESTOCK':'sobre','SIN ROTACION':'sinrot','SUBSTOCK':'subst','DESABASTECIDO':'desab','SIN INFORMACION':'sininfo'})[ind]||'sininfo';
    return `<tr>
      <td>${red}</td><td>${est}</td><td>${med}</td>
      <td>${fmtInt(r[5])}</td><td>${fmtInt(r[6])}</td><td>${fmt(r[7],2)}</td><td>${fmt(r[8],2)}</td>
      <td><span class="chip ${chipCls}"><span class="dot"></span>${IND_LABELS[ind]||ind}</span></td>
    </tr>`;
  }).join('');
  document.getElementById('tableDet').innerHTML = `<thead>${thead}</thead><tbody>${tbody}</tbody>`;
  document.getElementById('detPagi').innerHTML = `
    <button id="detPrev" ${detPage<=1?'disabled':''}>← Anterior</button>
    <span>Página ${detPage} de ${totalPages}</span>
    <button id="detNext" ${detPage>=totalPages?'disabled':''}>Siguiente →</button>
  `;
  const prevBtn = document.getElementById('detPrev'), nextBtn = document.getElementById('detNext');
  if(prevBtn) prevBtn.addEventListener('click', ()=>{detPage--; renderDetTable();});
  if(nextBtn) nextBtn.addEventListener('click', ()=>{detPage++; renderDetTable();});
}

/* ============================================================
   Populate filter dropdowns
   ============================================================ */
function populateFilters(){
  const reds = [...new Set(EST.map(e=>e.red))].sort();
  const redOptsHtml = reds.map(r=>`<option value="${r}">${r}</option>`).join('');
  document.getElementById('estabRedFilter').innerHTML += redOptsHtml;
  document.getElementById('detRedFilter').innerHTML += redOptsHtml;
  const indOptsHtml = IND_ORDER.filter(k=>globalTotals[k]>0).map(k=>`<option value="${k}">${IND_LABELS[k]}</option>`).join('');
  document.getElementById('detIndFilter').innerHTML += indOptsHtml;
}

/* ============================================================
   Render (llamado por loadAndRender tras cada carga de datos)
   ============================================================ */
let listenersBound = false;
function destroyCharts(){
  [chartDonut, chartRedBar, chartRedStack, chartEstabBar, chartProdBar].forEach(c=>{ if(c) c.destroy(); });
}
function render(){
  destroyCharts();
  detPage = 1;
  document.getElementById('estabRedFilter').innerHTML = '<option value="">Todas las redes</option>';
  document.getElementById('detRedFilter').innerHTML = '<option value="">Todas las redes</option>';
  document.getElementById('detIndFilter').innerHTML = '<option value="">Todas las clasificaciones</option>';
  populateFilters();

  renderAlert();
  renderKpis();
  renderDonut();
  renderRedBar();
  renderTopCriticoResumen();
  renderTopProdResumen();

  renderRedCards();
  renderRedStack();
  renderTableRed();

  renderEstabBar();
  renderEstabTable();

  renderAlertasView();

  renderProdBar();
  renderProdTable();

  renderDetTable();

  renderRiskPanel();
  populatePlanFilters();
  renderPlan();

  if(!listenersBound){
    bindPlan();
    document.getElementById('estabSearch').addEventListener('input', renderEstabTable);
    document.getElementById('estabRedFilter').addEventListener('change', renderEstabTable);
    document.getElementById('prodSearch').addEventListener('input', renderProdTable);
    document.getElementById('prodSort').addEventListener('change', (e)=>{prodSortKey=e.target.value; renderProdTable();});
    document.getElementById('detSearch').addEventListener('input', ()=>{detPage=1; renderDetTable();});
    document.getElementById('detRedFilter').addEventListener('change', ()=>{detPage=1; renderDetTable();});
    document.getElementById('detIndFilter').addEventListener('change', ()=>{detPage=1; renderDetTable();});
    listenersBound = true;
  }
}
/* ============================================================
   ALERTA DE RIESGO + PLAN DE ABASTECIMIENTO
   ============================================================ */
let RAW = [];
const PLAN_STATE = { target: 95, months: 3, red: '', prio: 'desab', scope: 'below', tab: 'eess' };
let PLAN = null;
const CHIP_CLS = {'NORMOSTOCK':'normo','SOBRESTOCK':'sobre','SIN ROTACION':'sinrot','SUBSTOCK':'subst','DESABASTECIDO':'desab','SIN INFORMACION':'sininfo'};

function goView(v){
  document.querySelectorAll('.nav-item').forEach(n => n.classList.toggle('active', n.dataset.view === v));
  document.querySelectorAll('.view').forEach(vv => vv.classList.remove('active'));
  document.getElementById('view-' + v).classList.add('active');
  document.getElementById('pageTitle').textContent = pageMeta[v].title;
  document.getElementById('pageDesc').textContent = pageMeta[v].desc;
  window.scrollTo(0, 0);
}

/* ---------- Alerta de establecimientos en riesgo (Resumen + insignia del menú) ---------- */
function renderRiskPanel(){
  const crit = EST.filter(e => e.DISPONIBILIDAD <= 80).sort((a,b) => a.DISPONIBILIDAD - b.DISPONIBILIDAD);
  const warn = EST.filter(e => e.DISPONIBILIDAD > 80 && e.DISPONIBILIDAD <= 90);
  const box = document.getElementById('riskBox');
  const badge = document.getElementById('navRiskBadge');
  badge.textContent = crit.length;
  badge.style.display = crit.length ? 'inline-block' : 'none';
  if (!crit.length && !warn.length) { box.innerHTML = ''; return; }
  const cls = crit.length ? 'crit' : 'warn';
  const names = (crit.length ? crit : warn).slice(0, 5)
    .map(e => `${e.est} (${fmt(e.DISPONIBILIDAD,1)}%)`).join(' · ');
  const more = (crit.length ? crit.length : warn.length) > 5 ? ' …' : '';
  box.innerHTML = `<div class="alert ${cls}">
    <div class="alert-ico">${crit.length ? '🚨' : '⚠️'}</div>
    <div>
      <b>${crit.length} establecimiento(s) en riesgo crítico (≤80%) y ${warn.length} en alerta (80–90%)</b>
      <span>Más bajos: ${names}${more}</span>
      <button class="alert-btn" onclick="goView('plan')">Ver plan de abastecimiento →</button>
    </div>
  </div>`;
}

/* ---------- Cálculo del plan ---------- */
function planRank(it){ return it.ind === 'DESABASTECIDO' ? 0 : 1; }

function computePlan(){
  const { target, months, red, prio, scope } = PLAN_STATE;
  const byEst = {};
  RAW.forEach(r => {
    if (r.ind !== 'DESABASTECIDO' && r.ind !== 'SUBSTOCK') return;
    const k = r.red + '||' + r.est;
    (byEst[k] = byEst[k] || []).push(r);
  });

  const eess = [], prodMap = new Map();
  let sumCur = 0, sumProj = 0, noCpaTotal = 0, unreachableCount = 0;

  EST.forEach(e => {
    const cur = e.DISPONIBILIDAD;
    const T = e.TOTAL;
    const D = e.DISPONIBLES || (e.NORMOSTOCK + e.SOBRESTOCK);
    let proj = cur;
    const inScope = T > 0 && (!red || e.red === red) && (scope === 'crit' ? (cur <= 80 && cur < target) : cur < target);

    if (inScope) {
      // ítems que faltan convertir a Normostock/Sobrestock para llegar a la meta
      const need = Math.max(0, Math.ceil(target * T / 100 - 1e-9) - D);
      const cand = []; let noCpa = 0;
      (byEst[e.red + '||' + e.est] || []).forEach(r => {
        if (!(r.cpa > 0)) { noCpa++; return; }          // sin consumo: no se puede proyectar
        const goal = r.cpa * months;                     // stock objetivo = CPA × meses de cobertura
        const qty = Math.ceil(goal - r.stk - 1e-9);      // cantidad a distribuir
        if (qty <= 0) return;
        cand.push({ codmed: r.codmed, med: r.med, ind: r.ind, stk: r.stk, cpa: r.cpa, disp: r.disp, qty, rel: qty / goal });
      });
      cand.sort((a, b) => prio === 'desab'
        ? (planRank(a) - planRank(b)) || (a.rel - b.rel) || (b.cpa - a.cpa)
        : (a.rel - b.rel) || (b.cpa - a.cpa));
      const items = cand.slice(0, need);
      const unreachable = cand.length < need;
      proj = Math.min(100, cur + items.length / T * 100);
      const units = items.reduce((s, i) => s + i.qty, 0);
      eess.push({ red: e.red, est: e.est, cod: e.cod, cur, proj, need, items, units, unreachable, short: Math.max(0, need - cand.length), noCpa });
      noCpaTotal += noCpa;
      if (unreachable) unreachableCount++;
      items.forEach(it => {
        const key = it.codmed || it.med;
        let p = prodMap.get(key);
        if (!p) { p = { codmed: it.codmed, med: it.med, units: 0, list: [] }; prodMap.set(key, p); }
        p.units += it.qty;
        p.list.push({ red: e.red, est: e.est, ind: it.ind, qty: it.qty });
      });
    }
    sumCur += cur; sumProj += proj;
  });

  eess.sort((a, b) => a.cur - b.cur);
  const prods = [...prodMap.values()].sort((a, b) => b.list.length - a.list.length || b.units - a.units);
  prods.forEach(p => p.list.sort((a, b) => b.qty - a.qty));
  const n = EST.length || 1;
  const globalProj = Math.min(100, globalDisponibilidad + (sumProj - sumCur) / n);
  const itemsTotal = eess.reduce((s, e) => s + e.items.length, 0);
  return { eess, prods, itemsTotal, globalProj, noCpaTotal, unreachableCount };
}

/* ---------- Render del plan ---------- */
function renderPlan(){
  if (!EST.length) return;
  const P = PLAN_STATE;
  PLAN = computePlan();
  const hoy = new Date().toLocaleDateString('es-PE', { day: '2-digit', month: '2-digit', year: 'numeric' });

  document.getElementById('planPrintHead').innerHTML =
    `<h2>Plan de abastecimiento · DIRESA Callao – DEMID</h2>
     <p>Meta de disponibilidad ${P.target}% · Cobertura objetivo ${P.months} meses${P.red ? ' · ' + P.red : ''} · Generado el ${hoy}<br>
     ${PLAN.eess.length} EESS a abastecer · ${fmtInt(PLAN.itemsTotal)} ítems · Disponibilidad global ${fmt(globalDisponibilidad,1)}% → ${fmt(PLAN.globalProj,1)}% proyectada</p>`;

  document.getElementById('planKpis').innerHTML = `
    <div class="kpi"><div class="label">EESS a abastecer</div><div class="value">${PLAN.eess.length}</div><div class="sub">por debajo de la meta de ${P.target}%</div></div>
    <div class="kpi"><div class="label">Ítems a abastecer</div><div class="value">${fmtInt(PLAN.itemsTotal)}</div><div class="sub">combinaciones producto × EESS</div></div>
    <div class="kpi"><div class="label">Productos distintos</div><div class="value">${fmtInt(PLAN.prods.length)}</div><div class="sub">a distribuir en total</div></div>
    <div class="kpi"><div class="label">Disponibilidad global proyectada</div><div class="value" style="color:${textColorForDisp(PLAN.globalProj)}">${fmt(PLAN.globalProj,1)}%</div><div class="sub">hoy ${fmt(globalDisponibilidad,1)}% (promedio entre ${EST.length} EESS)</div></div>`;

  let notes = `Cantidad a distribuir = CPA × ${P.months} meses − stock actual (redondeado hacia arriba). Solo se consideran ítems Desabastecidos y Substock con consumo (CPA &gt; 0); los ítems Sin rotación se resuelven redistribuyendo stock entre EESS, no con nuevo abastecimiento. No se descuenta el stock del almacén central.`;
  if (PLAN.unreachableCount) notes += ` <strong>${PLAN.unreachableCount} EESS no llegan a la meta</strong> solo con abastecimiento (faltan ítems con consumo); aparecen marcados.`;
  if (PLAN.noCpaTotal) notes += ` ${fmtInt(PLAN.noCpaTotal)} ítems sin CPA no se pueden proyectar.`;
  document.getElementById('planNotes').innerHTML = `<div class="alert warn no-print"><div class="alert-ico">ℹ️</div><div><b>Cómo se calcula el plan</b><span>${notes}</span></div></div>`;

  // ----- Por establecimiento -----
  const eessHtml = PLAN.eess.map(e => {
    const rows = e.items.length ? e.items.map((it, i) => `
      <tr><td>${i + 1}</td><td>${it.codmed}</td><td>${it.med}</td>
      <td><span class="chip ${CHIP_CLS[it.ind]}"><span class="dot"></span>${IND_LABELS[it.ind]}</span></td>
      <td>${fmtInt(it.stk)}</td><td>${fmt(it.cpa,1)}</td><td>${fmt(it.disp,1)}</td>
      <td><b>${fmtInt(it.qty)}</b></td></tr>`).join('')
      : `<tr><td colspan="8" style="color:#8a5b12">Sin ítems con consumo para abastecer en este establecimiento.</td></tr>`;
    const warn = e.unreachable
      ? `<div class="plan-warn">Con ítems abastecibles solo se alcanza ${fmt(e.proj,1)}% (faltan ${e.short} ítems para la meta${e.noCpa ? '; ' + e.noCpa + ' sin CPA' : ''}).</div>` : '';
    return `<div class="plan-eess">
      <div class="plan-eess-head">
        <div><b>${e.est}</b><span>${e.red}${e.cod ? ' · ' + e.cod : ''}</span></div>
        <div class="plan-prog"><span style="color:${textColorForDisp(e.cur)}">${fmt(e.cur,1)}%</span> → <span style="color:${textColorForDisp(e.proj)}">${fmt(e.proj,1)}%</span></div>
        <div class="plan-count">${e.items.length} ítems</div>
      </div>
      ${warn}
      <div class="table-scroll"><table>
        <thead><tr><th>#</th><th>Cód.</th><th>Medicamento</th><th>Clasificación</th><th>Stock actual</th><th>CPA</th><th>Meses disp.</th><th>Cantidad a distribuir</th></tr></thead>
        <tbody>${rows}</tbody></table></div>
    </div>`;
  }).join('');
  document.getElementById('planEessList').innerHTML = eessHtml ||
    `<div class="alert ok"><div class="alert-ico">✅</div><div><b>Ningún establecimiento por debajo de la meta</b><span>Con los filtros actuales no hay nada que abastecer.</span></div></div>`;
  document.getElementById('planEessCount').textContent = `${PLAN.eess.length} establecimientos · ${fmtInt(PLAN.itemsTotal)} ítems`;

  // ----- Por producto -----
  const prodHtml = PLAN.prods.map(p => {
    const head = `<tr class="plan-prod-head"><td colspan="3"><b>${p.med}</b> <span class="muted">${p.codmed || ''}</span></td><td colspan="2"><b>${fmtInt(p.units)} u.</b> en total · ${p.list.length} EESS</td></tr>`;
    const kids = p.list.map(x => `
      <tr><td></td><td>${x.est}</td><td>${x.red}</td>
      <td><span class="chip ${CHIP_CLS[x.ind]}"><span class="dot"></span>${IND_LABELS[x.ind]}</span></td>
      <td><b>${fmtInt(x.qty)}</b></td></tr>`).join('');
    return head + kids;
  }).join('');
  document.getElementById('tablePlanProd').innerHTML =
    `<thead><tr><th></th><th>Establecimiento</th><th>Red</th><th>Clasificación</th><th>Cantidad a distribuir</th></tr></thead><tbody>${prodHtml || '<tr><td colspan="5">Sin productos por distribuir.</td></tr>'}</tbody>`;
  document.getElementById('planProdCount').textContent = `${PLAN.prods.length} productos · ${fmtInt(PLAN.itemsTotal)} entregas`;
}

/* ---------- Exportar a Excel ---------- */
function exportPlanXlsx(){
  if (!PLAN) return;
  const wb = XLSX.utils.book_new();
  const flat = [];
  PLAN.eess.forEach(e => e.items.forEach(it => flat.push({
    'Red': e.red, 'Establecimiento': e.est,
    'Disp. actual %': +e.cur.toFixed(1), 'Disp. proyectada %': +e.proj.toFixed(1),
    'Código': it.codmed, 'Medicamento': it.med, 'Clasificación': IND_LABELS[it.ind],
    'Stock actual': it.stk, 'CPA': it.cpa, 'Meses disp.': it.disp, 'Cantidad a distribuir': it.qty
  })));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(flat), 'Por establecimiento');
  const prodRows = PLAN.prods.map(p => ({ 'Código': p.codmed, 'Medicamento': p.med, 'EESS a abastecer': p.list.length, 'Cantidad total': p.units }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(prodRows), 'Por producto');
  const resRows = PLAN.eess.map(e => ({ 'Red': e.red, 'Establecimiento': e.est, 'Disp. actual %': +e.cur.toFixed(1), 'Disp. proyectada %': +e.proj.toFixed(1), 'Ítems a abastecer': e.items.length, 'Alcanza la meta': e.unreachable ? 'No' : 'Sí' }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(resRows), 'Resumen EESS');
  const f = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(wb, `plan_abastecimiento_${PLAN_STATE.target}pct_${f}.xlsx`);
}

/* ---------- Imprimir ---------- */
function printPlan(){
  document.body.classList.add('print-' + PLAN_STATE.tab);
  window.print();
}

/* ---------- Eventos ---------- */
function setPlanTab(tab){
  PLAN_STATE.tab = tab;
  document.querySelectorAll('#planTabs button[data-ptab]').forEach(b => b.classList.toggle('active', b.dataset.ptab === tab));
  document.getElementById('planPaneEess').style.display = tab === 'eess' ? '' : 'none';
  document.getElementById('planPaneProd').style.display = tab === 'prod' ? '' : 'none';
}
function bindPlan(){
  const num = (id, key, min, max) => document.getElementById(id).addEventListener('input', e => {
    const v = parseFloat(e.target.value);
    if (isNaN(v)) return;
    PLAN_STATE[key] = Math.min(max, Math.max(min, v));
    renderPlan();
  });
  num('planTarget', 'target', 50, 100);
  num('planMonths', 'months', 2, 6);
  [['planRed', 'red'], ['planPrio', 'prio'], ['planScope', 'scope']].forEach(([id, key]) =>
    document.getElementById(id).addEventListener('change', e => { PLAN_STATE[key] = e.target.value; renderPlan(); }));
  document.querySelectorAll('#planTabs button[data-ptab]').forEach(b => b.addEventListener('click', () => setPlanTab(b.dataset.ptab)));
  document.getElementById('planBtnPrint').addEventListener('click', printPlan);
  document.getElementById('planBtnXlsx').addEventListener('click', exportPlanXlsx);
  window.addEventListener('afterprint', () => document.body.classList.remove('print-eess', 'print-prod'));
}

function populatePlanFilters(){
  const sel = document.getElementById('planRed');
  const reds = [...new Set(EST.map(e=>e.red))].sort();
  sel.innerHTML = '<option value="">Todas las redes</option>' + reds.map(r=>`<option value="${r}">${r}</option>`).join('');
  sel.value = PLAN_STATE.red;
}

loadAndRender();
