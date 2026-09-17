/* =========================================================================
   Tableau de Bord des Projets OCBA — logique front-end (100% client-side)
   ========================================================================= */

const DEFAULT_DATA_URL = 'data/projets.json';

// Mapping en-têtes Excel (FR) -> clés internes, identique au script Python
// scripts/excel_to_json.py, pour que l'import client garde le même contrat.
const COLUMN_MAP = {
  'Nom du site': 'site',
  'No de site': 'noSite',
  'No de sous-site': 'noSousSite',
  'Code bâtiment': 'codeBatiment',
  'EGID': 'egid',
  'Nom du bâtiment': 'batiment',
  'Type de bâtiment': 'typeBatiment',
  'Année de construction': 'anneeConstruction',
  'Adresse': 'adresse',
  'Commune': 'commune',
  'Code postal': 'codePostal',
  'Propriétaire': 'proprietaire',
  'Destination': 'destination',
  'Politique publique': 'politique',
  'Nom du gérant technique': 'gerantTechnique',
  'Nom du gérant administratif': 'gerantAdministratif',
  'Surface référentiel énergétique (SRE) m²': 'sre',
  'Numéro SIBAT': 'numeroSibat',
  'Lien GE-Invest  Projet': 'lienGeInvest',
  'Numéro Projet': 'numeroProjet',
  'Projet': 'projet',
  'Chef de Projet': 'chefProjet',
  'Statut': 'statut',
  'Loi': 'loi',
  'Budget': 'budget',
  'Avancement  (%)': 'avancement',
  'Date de début prévue de projet': 'dateDebutPrevue',
  'Date de fin prévue de projet': 'dateFinPrevue',
  'Durée prévue projet (Jours)': 'dureePrevueJours',
  'Date de début réelle de projet': 'dateDebutReelle',
  'Date de mise en service de projet': 'dateMiseEnService',
  "Durée jusqu'à mise en service (jours)": 'dureeMiseEnServiceJours',
};

const FILTER_DEFS = [
  { id: 'f-chef', key: 'chefProjet' },
  { id: 'f-site', key: 'site' },
  { id: 'f-statut', key: 'statut' },
  { id: 'f-commune', key: 'commune' },
  { id: 'f-destination', key: 'destination' },
  { id: 'f-politique', key: 'politique' },
  { id: 'f-proprietaire', key: 'proprietaire' },
  { id: 'f-loi', key: 'loi' },
];

const STATUT_BADGE = {
  '2-EN COURS': 'badge-en-cours',
  '3-MIS EN SERVICE': 'badge-mis-en-service',
  '4-TERMINE': 'badge-termine',
  '5-BOUCLE': 'badge-boucle',
};

const state = {
  allRecords: [],
  filteredRecords: [],
  search: '',
  selections: {},        // key -> Set of selected values
  sourceLabel: "Données d'exemple",
  isImported: false,
  table: { page: 1, pageSize: 20, sortKey: 'dateMiseEnService', sortDir: 'desc' },
};

FILTER_DEFS.forEach(f => (state.selections[f.key] = new Set()));

/* ------------------------------- Utils --------------------------------- */

function fmtMoney(value) {
  if (value == null || Number.isNaN(value)) return '—';
  if (Math.abs(value) >= 1e6) return `CHF ${(value / 1e6).toFixed(1)} M`;
  if (Math.abs(value) >= 1e3) return `CHF ${(value / 1e3).toFixed(0)} k`;
  return `CHF ${value.toFixed(0)}`;
}

function fmtNumber(value) {
  if (value == null || Number.isNaN(value)) return '—';
  return new Intl.NumberFormat('fr-CH').format(value);
}

function fmtDate(value) {
  if (!value) return '—';
  return value;
}

function uniqueSorted(records, key) {
  const values = new Set();
  records.forEach(r => {
    const v = r[key];
    if (v !== null && v !== undefined && v !== '') values.add(v);
  });
  return Array.from(values).sort((a, b) => String(a).localeCompare(String(b), 'fr'));
}

function excelSerialToISO(serial) {
  const utcDays = Math.floor(serial - 25569);
  const date = new Date(utcDays * 86400 * 1000);
  return date.toISOString().slice(0, 10);
}

function normalizeDateValue(value) {
  if (value === null || value === undefined || value === '') return null;
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value.toISOString().slice(0, 10);
  }
  if (typeof value === 'number') {
    try { return excelSerialToISO(value); } catch { return null; }
  }
  const str = String(value).trim();
  if (!str || str === '00:00:00') return null;
  // dd.mm.yyyy ou dd/mm/yyyy
  const m = str.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{4})$/);
  if (m) return `${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`;
  const parsed = new Date(str);
  if (!Number.isNaN(parsed.getTime())) return parsed.toISOString().slice(0, 10);
  return null;
}

/* --------------------------- Chargement JSON ---------------------------- */

async function loadDefaultData() {
  const res = await fetch(DEFAULT_DATA_URL, { cache: 'no-store' });
  const payload = await res.json();
  applyNewDataset(payload.records, "Données d'exemple", false, payload.generatedAt);
}

function applyNewDataset(records, label, imported, generatedAt) {
  state.allRecords = records;
  state.sourceLabel = label;
  state.isImported = imported;
  state.search = '';
  document.getElementById('f-search').value = '';
  FILTER_DEFS.forEach(f => state.selections[f.key].clear());

  populateFilterOptions();
  updateDataStatus(generatedAt);
  refresh();
}

function updateDataStatus(generatedAt) {
  const box = document.getElementById('data-status');
  const text = document.getElementById('data-status-text');
  box.classList.toggle('imported', state.isImported);
  text.textContent = state.isImported ? `Import : ${state.sourceLabel}` : "Données d'exemple";
  document.getElementById('last-update').textContent = generatedAt
    ? `mis à jour ${new Date(generatedAt).toLocaleString('fr-CH')}`
    : '';
}

/* --------------------------- Import Excel (SheetJS) ---------------------------- */

function handleFileImport(file) {
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const wb = XLSX.read(e.target.result, { type: 'array', cellDates: true });
      const sheetName = wb.SheetNames.includes('Feuil1') ? 'Feuil1' : wb.SheetNames[0];
      const sheet = wb.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(sheet, { defval: null, raw: true });
      const records = rows.map(mapRawRow).filter(r => r.projet && r.chefProjet);
      records.forEach((r, i) => (r.id = i + 1));
      applyNewDataset(records, file.name, true, new Date().toISOString());
    } catch (err) {
      console.error(err);
      alert("Impossible de lire ce fichier. Vérifiez qu'il s'agit bien d'un export Excel du même format.");
    }
  };
  reader.readAsArrayBuffer(file);
}

function mapRawRow(row) {
  const out = {};
  Object.entries(row).forEach(([rawHeader, value]) => {
    // Les en-têtes Excel multi-lignes utilisent \r\n ; ne garder que le \n comme en Python (openpyxl absorbe le \r)
    const header = String(rawHeader).replace(/\r/g, '').replace(/\n/g, ' ').trim();
    const key = COLUMN_MAP[header];
    if (!key) return;
    out[key] = value;
  });

  ['dateDebutPrevue', 'dateFinPrevue', 'dateDebutReelle', 'dateMiseEnService'].forEach(k => {
    if (k in out) out[k] = normalizeDateValue(out[k]);
  });

  if ('avancement' in out && out.avancement !== null && out.avancement !== '') {
    const v = Number(out.avancement);
    out.avancement = Number.isNaN(v) ? null : (v <= 1 ? Math.round(v * 1000) / 10 : Math.round(v * 10) / 10);
  }
  if ('loi' in out && out.loi !== null && out.loi !== '') {
    const v = Number(out.loi);
    out.loi = Number.isNaN(v) ? null : Math.round(v);
  }
  ['destination', 'politique', 'proprietaire', 'typeBatiment'].forEach(k => {
    if (out[k] === 0 || out[k] === '0') out[k] = null;
  });
  ['budget', 'sre', 'anneeConstruction', 'noSite', 'noSousSite', 'codePostal', 'egid', 'numeroSibat',
   'dureePrevueJours', 'dureeMiseEnServiceJours'].forEach(k => {
    if (k in out && out[k] !== null && out[k] !== '') {
      const v = Number(out[k]);
      out[k] = Number.isNaN(v) ? null : v;
    }
  });
  return out;
}

/* ------------------------------ Filtres UI ------------------------------ */

function populateFilterOptions() {
  FILTER_DEFS.forEach(({ id, key }) => {
    const select = document.getElementById(id);
    const values = uniqueSorted(state.allRecords, key);
    select.innerHTML = values.map(v => `<option value="${escapeHtml(String(v))}">${escapeHtml(String(v))}</option>`).join('');
  });
}

function escapeHtml(str) {
  return str.replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

function readFilterSelections() {
  FILTER_DEFS.forEach(({ id, key }) => {
    const select = document.getElementById(id);
    const values = Array.from(select.selectedOptions).map(o => o.value);
    state.selections[key] = new Set(values);
  });
  state.search = document.getElementById('f-search').value.trim().toLowerCase();
}

function applyFilters() {
  readFilterSelections();
  state.filteredRecords = state.allRecords.filter(r => {
    for (const { key } of FILTER_DEFS) {
      const sel = state.selections[key];
      if (sel.size > 0 && !sel.has(String(r[key]))) return false;
    }
    if (state.search) {
      const haystack = `${r.projet ?? ''} ${r.site ?? ''} ${r.batiment ?? ''} ${r.chefProjet ?? ''}`.toLowerCase();
      if (!haystack.includes(state.search)) return false;
    }
    return true;
  });
}

/* -------------------------------- KPIs ---------------------------------- */

function renderKPIs() {
  const data = state.filteredRecords;
  const total = data.length;
  const budgetTotal = data.reduce((s, r) => s + (r.budget || 0), 0);
  const avancements = data.map(r => r.avancement).filter(v => v !== null && v !== undefined);
  const avgAvancement = avancements.length ? avancements.reduce((a, b) => a + b, 0) / avancements.length : null;
  const enCours = data.filter(r => r.statut === '2-EN COURS').length;
  const nbSites = new Set(data.map(r => r.site).filter(Boolean)).size;
  const nbChefs = new Set(data.map(r => r.chefProjet).filter(Boolean)).size;

  const cards = [
    { label: 'Projets filtrés', value: fmtNumber(total), sub: `sur ${fmtNumber(state.allRecords.length)} au total`, cls: 'accent-blue' },
    { label: 'Budget cumulé', value: fmtMoney(budgetTotal), sub: 'toutes phases confondues', cls: 'accent-green' },
    { label: 'Avancement moyen', value: avgAvancement !== null ? `${avgAvancement.toFixed(0)}%` : '—', sub: `${enCours} projet(s) en cours`, cls: 'accent-orange' },
    { label: 'Sites / Chefs de projet', value: `${fmtNumber(nbSites)} / ${fmtNumber(nbChefs)}`, sub: 'périmètre actif', cls: 'accent-purple' },
  ];

  document.getElementById('kpi-grid').innerHTML = cards.map(c => `
    <div class="kpi-card ${c.cls}">
      <div class="kpi-label">${c.label}</div>
      <div class="kpi-value">${c.value}</div>
      <div class="kpi-sub">${c.sub}</div>
    </div>`).join('');

  document.getElementById('filter-count-info').textContent = `${fmtNumber(total)} projet${total > 1 ? 's' : ''}`;
}

/* ------------------------------- Graphiques ------------------------------ */

const PLOTLY_LAYOUT_BASE = {
  font: { family: 'Inter, sans-serif', size: 12, color: '#1a2233' },
  margin: { l: 60, r: 20, t: 10, b: 40 },
  paper_bgcolor: 'rgba(0,0,0,0)',
  plot_bgcolor: 'rgba(0,0,0,0)',
  colorway: ['#2f5dfc', '#12b76a', '#f79009', '#7c3aed', '#f04438', '#0891b2', '#db2777'],
};

function plot(id, traces, layout, config) {
  Plotly.react(id, traces, { ...PLOTLY_LAYOUT_BASE, ...layout }, { displayModeBar: false, responsive: true, ...config });
}

function countBy(records, key) {
  const map = new Map();
  records.forEach(r => {
    const v = r[key];
    if (v === null || v === undefined || v === '') return;
    map.set(v, (map.get(v) || 0) + 1);
  });
  return map;
}

function sumBy(records, key, valueKey) {
  const map = new Map();
  records.forEach(r => {
    const v = r[key];
    if (v === null || v === undefined || v === '') return;
    map.set(v, (map.get(v) || 0) + (r[valueKey] || 0));
  });
  return map;
}

function avgBy(records, key, valueKey) {
  const sums = new Map(), counts = new Map();
  records.forEach(r => {
    const v = r[key];
    const val = r[valueKey];
    if (v === null || v === undefined || v === '' || val === null || val === undefined) return;
    sums.set(v, (sums.get(v) || 0) + val);
    counts.set(v, (counts.get(v) || 0) + 1);
  });
  const out = new Map();
  sums.forEach((s, k) => out.set(k, s / counts.get(k)));
  return out;
}

function topEntries(map, n) {
  return Array.from(map.entries()).sort((a, b) => b[1] - a[1]).slice(0, n);
}

function renderChartSites() {
  const entries = topEntries(countBy(state.filteredRecords, 'site'), 10).reverse();
  plot('chart-sites', [{
    type: 'bar', orientation: 'h',
    x: entries.map(e => e[1]), y: entries.map(e => e[0]),
    marker: { color: '#2f5dfc' },
  }], { margin: { l: 190, r: 20, t: 10, b: 30 } });
}

function renderChartStatut() {
  const map = countBy(state.filteredRecords, 'statut');
  const entries = Array.from(map.entries());
  plot('chart-statut', [{
    type: 'pie', hole: 0.45,
    labels: entries.map(e => e[0]), values: entries.map(e => e[1]),
    textinfo: 'label+percent', textposition: 'outside',
  }], { margin: { l: 10, r: 10, t: 10, b: 10 }, showlegend: false });
}

function renderChartPolitique() {
  const entries = topEntries(sumBy(state.filteredRecords, 'politique', 'budget'), 12).reverse();
  plot('chart-politique', [{
    type: 'bar', orientation: 'h',
    x: entries.map(e => e[1]), y: entries.map(e => e[0]),
    marker: { color: '#12b76a' },
    hovertemplate: '%{y}<br>CHF %{x:,.0f}<extra></extra>',
  }], { margin: { l: 260, r: 20, t: 10, b: 30 } });
}

function renderChartCommune() {
  const entries = topEntries(countBy(state.filteredRecords, 'commune'), 10);
  plot('chart-commune', [{
    type: 'bar', x: entries.map(e => e[0]), y: entries.map(e => e[1]),
    marker: { color: '#f79009' },
  }], { margin: { l: 50, r: 20, t: 10, b: 90 }, xaxis: { tickangle: -35 } });
}

function renderChartDestination() {
  const map = avgBy(state.filteredRecords, 'destination', 'avancement');
  const entries = Array.from(map.entries()).sort((a, b) => b[1] - a[1]);
  plot('chart-destination', [{
    type: 'bar', orientation: 'h',
    x: entries.map(e => e[1]), y: entries.map(e => e[0]),
    marker: { color: '#7c3aed' },
    hovertemplate: '%{y}<br>%{x:.1f}%<extra></extra>',
  }], { margin: { l: 170, r: 20, t: 10, b: 30 }, xaxis: { ticksuffix: '%' } });
}

function renderChartChef() {
  const entries = topEntries(countBy(state.filteredRecords, 'chefProjet'), 20).reverse();
  plot('chart-chef', [{
    type: 'bar', orientation: 'h',
    x: entries.map(e => e[1]), y: entries.map(e => e[0]),
    marker: { color: '#2f5dfc' },
  }], { margin: { l: 190, r: 20, t: 10, b: 30 } });
}

function renderChartChefBudget() {
  const entries = topEntries(sumBy(state.filteredRecords, 'chefProjet', 'budget'), 10).reverse();
  plot('chart-chef-budget', [{
    type: 'bar', orientation: 'h',
    x: entries.map(e => e[1]), y: entries.map(e => e[0]),
    marker: { color: '#12b76a' },
    hovertemplate: '%{y}<br>CHF %{x:,.0f}<extra></extra>',
  }], { margin: { l: 190, r: 20, t: 10, b: 30 } });
}

function renderChartChefStatut() {
  const topChefs = topEntries(countBy(state.filteredRecords, 'chefProjet'), 8).map(e => e[0]);
  const statuts = Array.from(new Set(state.filteredRecords.map(r => r.statut).filter(Boolean)));
  const traces = statuts.map(statut => ({
    type: 'bar', name: statut,
    x: topChefs,
    y: topChefs.map(chef => state.filteredRecords.filter(r => r.chefProjet === chef && r.statut === statut).length),
  }));
  plot('chart-chef-statut', traces, { barmode: 'stack', xaxis: { tickangle: -30 }, legend: { orientation: 'h', y: -0.3 } });
}

function renderChartTimeline() {
  const counts = new Map();
  state.filteredRecords.forEach(r => {
    if (!r.dateMiseEnService) return;
    const year = r.dateMiseEnService.slice(0, 4);
    const y = Number(year);
    if (y < 1990 || y > 2035) return; // écarte les dates aberrantes (ex. 2121)
    counts.set(year, (counts.get(year) || 0) + 1);
  });
  const years = Array.from(counts.keys()).sort();
  plot('chart-timeline', [{
    type: 'bar', x: years, y: years.map(y => counts.get(y)),
    marker: { color: '#2f5dfc' },
  }], { margin: { l: 50, r: 20, t: 10, b: 40 } });
}

function renderChartGantt() {
  const sample = state.filteredRecords
    .filter(r => r.dateDebutPrevue || r.dateDebutReelle)
    .slice(0, 30);

  const labels = sample.map(r => r.projet.length > 40 ? r.projet.slice(0, 40) + '…' : r.projet);
  const prevX = [], prevY = [], realX = [], realY = [];

  sample.forEach((r, i) => {
    if (r.dateDebutPrevue && r.dateFinPrevue) {
      prevX.push(r.dateDebutPrevue, r.dateFinPrevue, null);
      prevY.push(labels[i], labels[i], null);
    }
    if (r.dateDebutReelle && r.dateMiseEnService) {
      realX.push(r.dateDebutReelle, r.dateMiseEnService, null);
      realY.push(labels[i], labels[i], null);
    }
  });

  plot('chart-gantt', [
    { type: 'scatter', mode: 'lines', name: 'Prévu', x: prevX, y: prevY, line: { color: '#2f5dfc', width: 8 } },
    { type: 'scatter', mode: 'lines', name: 'Réel', x: realX, y: realY, line: { color: '#12b76a', width: 8 } },
  ], { margin: { l: 260, r: 20, t: 10, b: 40 }, legend: { orientation: 'h', y: -0.08 } });
}

function renderAllCharts() {
  renderChartSites();
  renderChartStatut();
  renderChartPolitique();
  renderChartCommune();
  renderChartDestination();
  renderChartChef();
  renderChartChefBudget();
  renderChartChefStatut();
  renderChartTimeline();
  renderChartGantt();
}

/* --------------------------------- Table --------------------------------- */

function renderTable() {
  const { sortKey, sortDir, page, pageSize } = state.table;
  const rows = [...state.filteredRecords].sort((a, b) => {
    const va = a[sortKey], vb = b[sortKey];
    if (va === null || va === undefined) return 1;
    if (vb === null || vb === undefined) return -1;
    if (va < vb) return sortDir === 'asc' ? -1 : 1;
    if (va > vb) return sortDir === 'asc' ? 1 : -1;
    return 0;
  });

  const totalPages = Math.max(1, Math.ceil(rows.length / pageSize));
  state.table.page = Math.min(page, totalPages);
  const start = (state.table.page - 1) * pageSize;
  const pageRows = rows.slice(start, start + pageSize);

  document.getElementById('data-table-body').innerHTML = pageRows.map(r => `
    <tr>
      <td>${escapeHtml(r.projet ?? '—')}</td>
      <td>${escapeHtml(r.site ?? '—')}</td>
      <td>${escapeHtml(r.chefProjet ?? '—')}</td>
      <td><span class="badge ${STATUT_BADGE[r.statut] || 'badge-boucle'}">${escapeHtml(r.statut ?? '—')}</span></td>
      <td>${fmtMoney(r.budget)}</td>
      <td>
        <div class="progress-bar"><span style="width:${Math.max(0, Math.min(100, r.avancement ?? 0))}%"></span></div>
        <small>${r.avancement != null ? r.avancement + '%' : '—'}</small>
      </td>
      <td>${fmtDate(r.dateMiseEnService)}</td>
    </tr>`).join('') || `<tr><td colspan="7" style="text-align:center;color:var(--text-muted);padding:24px;">Aucun projet ne correspond aux filtres.</td></tr>`;

  document.getElementById('table-count-info').textContent = `${fmtNumber(rows.length)} projet(s)`;
  renderPagination(totalPages);
}

function renderPagination(totalPages) {
  const { page } = state.table;
  const el = document.getElementById('pagination');
  let html = `<button ${page <= 1 ? 'disabled' : ''} data-page="${page - 1}">‹ Précédent</button>`;
  const maxButtons = 7;
  let start = Math.max(1, page - 3);
  let end = Math.min(totalPages, start + maxButtons - 1);
  start = Math.max(1, end - maxButtons + 1);
  for (let p = start; p <= end; p++) {
    html += `<button class="${p === page ? 'active' : ''}" data-page="${p}">${p}</button>`;
  }
  html += `<button ${page >= totalPages ? 'disabled' : ''} data-page="${page + 1}">Suivant ›</button>`;
  el.innerHTML = html;
  el.querySelectorAll('button[data-page]').forEach(btn => {
    btn.addEventListener('click', () => {
      state.table.page = Number(btn.dataset.page);
      renderTable();
    });
  });
}

document.querySelectorAll('#data-table th[data-key]').forEach(th => {
  th.addEventListener('click', () => {
    const key = th.dataset.key;
    if (state.table.sortKey === key) {
      state.table.sortDir = state.table.sortDir === 'asc' ? 'desc' : 'asc';
    } else {
      state.table.sortKey = key;
      state.table.sortDir = 'asc';
    }
    renderTable();
  });
});

/* ------------------------------- Orchestration ---------------------------- */

function refresh() {
  applyFilters();
  renderKPIs();
  renderAllCharts();
  renderTable();
}

function setupTabs() {
  document.querySelectorAll('.tab').forEach(tab => {
    tab.addEventListener('click', () => {
      document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
      document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
      tab.classList.add('active');
      document.querySelector(`.tab-panel[data-panel="${tab.dataset.tab}"]`).classList.add('active');
      window.dispatchEvent(new Event('resize')); // force Plotly à se redimensionner
    });
  });
}

function setupFilterEvents() {
  FILTER_DEFS.forEach(({ id }) => document.getElementById(id).addEventListener('change', refresh));
  document.getElementById('f-search').addEventListener('input', debounce(refresh, 200));
  document.getElementById('btn-clear-filters').addEventListener('click', () => {
    FILTER_DEFS.forEach(({ id }) => {
      document.getElementById(id).querySelectorAll('option').forEach(o => (o.selected = false));
    });
    document.getElementById('f-search').value = '';
    refresh();
  });
}

function setupImportEvents() {
  document.getElementById('file-import').addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (file) handleFileImport(file);
    e.target.value = '';
  });
  document.getElementById('btn-reset-data').addEventListener('click', () => loadDefaultData());
}

function debounce(fn, delay) {
  let t;
  return (...args) => { clearTimeout(t); t = setTimeout(() => fn(...args), delay); };
}

setupTabs();
setupFilterEvents();
setupImportEvents();
loadDefaultData().catch(err => {
  console.error(err);
  document.getElementById('kpi-grid').innerHTML = `<div class="kpi-card"><div class="kpi-label">Erreur</div><div class="kpi-value">Chargement impossible</div></div>`;
});
