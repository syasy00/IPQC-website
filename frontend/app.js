/* ============================================================
   IPQC Tracker — vanilla JS frontend
   No build step. No React. Talks to the Express+Excel backend
   defined in /backend (see config.js for the API_BASE URL).
   ============================================================ */

// Constants 
const DEPARTMENTS = [
  'Production Team',
  'Test Team',
  'IE Team',
  'Quality Team',
  'Calibration Team',
  'PE Team'
];

const CATEGORIES = [
  'Compliance_6S',
  'Calibration_PM',
  'Documentation_And_Process_Adherence',
  'ESD_Control',
  'Material_Control_And_Chemical_Management',
  'Safety_Concern',
  'Tooling_Labeling',
  'Training_Certification'
];

const FINDING_DETAILS = [
  'Visual Standard Expired',
  'Assembly process conducted without glove usage',
  'Cable wire damage',
  'Calibration Label damage, Torn on Tools / Equipment',
  'Calibration Overdue ESD Monitor',
  'Calibration Overdue Torque Drive',
  'Chemical / Material Overdue',
  'Dust on workstation/rack/ect',
  'Dustbin located at non-kanban area',
  'Equipment without Calibration / PM Label',
  'ESD Monitoring not function',
  'Improper storage of Kit / Bulk Material',
  'Improper storage of Tool/Equipment',
  'Ionizer turn off',
  'IPA without Expiry Date Label',
  'Missing Label Expiry Date',
  'Mix material inside the material bin',
  'No ESD grounding points',
  'No Insulative Mat',
  'No Set-Up Checklist displayed',
  'Not Wear Safety Glass',
  'Preventive Maintenance Overdue',
  'Setup check list not updated',
  'Torque number is smear',
  'Unnecessary item/material found on the workstation'
];

const PLATFORMS = [
  'Apex',
  'Ascent',
  'Cesar',
  'Cumulus',
  'Evos',
  'Ewave',
  'HASS & Burn In',
  'HV',
  'HV (MV)',
  'HV (OL)',
  'Insource (Potting)',
  'Maxstream',
  'Navi I/AZX/LM/LFM/RFG',
  'Navi II',
  'OBA & PACKING',
  'Packing',
  'Paramount',
  'PDX',
  'Pinnacle III',
  'Scorpius',
  'Solvix',
  'VHF'
];

const CATEGORY_GROUP_MAPPING = {
  Compliance_6S: 'Method',
  Calibration_PM: 'Machine',
  Documentation_And_Process_Adherence: 'Method',
  ESD_Control: 'Machine',
  Material_Control_And_Chemical_Management: 'Material',
  Safety_Concern: 'Man',
  Tooling_Labeling: 'Material',
  Training_Certification: 'Man'
};

const PLATFORM_MQE_MAPPING = {
  Apex: 'Siti Naimah',
  PDX: 'Larry',
  'Navi I/AZX/LM/LFM/RFG': 'Farid',
  'Navi II': 'Farid'
};

const STATUSES = [
  'Open',
  'In Progress',
  'Closed'
];

const SHIFTS = [
  'A',
  'B',
  'C'
];

const ICONS = {
  dashboard: 'layout-dashboard',
  ipqc: 'clipboard-check',
  history: 'history',
  add: 'plus',
  settings: 'settings',
  search: 'search',
  edit: 'pencil',
  trash: 'trash-2',
  image: 'image',
  x: 'x',
  download: 'download',
  filter: 'filter',
  calendar: 'calendar-days',
  users: 'users',
  mapPin: 'map-pin',
  tag: 'tag',
  camera: 'camera',
  save: 'save',
  wifiOff: 'wifi-off',
  clipboardList: 'clipboard-list',
  circleAlert: 'circle-alert',
  circleCheck: 'circle-check',
  clock: 'clock',
  inbox: 'inbox',
  chevronRight: 'chevron-right',
};


const SVG_ICON = {
  calendar: '<path d="M8 2v4"/><path d="M16 2v4"/><rect width="18" height="18" x="3" y="4" rx="2"/><path d="M3 10h18"/>',
  mapPin: '<path d="M20 10c0 4.993-5.539 10.193-7.399 11.799a1 1 0 0 1-1.202 0C9.539 20.193 4 14.993 4 10a8 8 0 0 1 16 0"/><circle cx="12" cy="10" r="3"/>',
  tag: '<path d="M12.586 2.586A2 2 0 0 0 11.172 2H4a2 2 0 0 0-2 2v7.172a2 2 0 0 0 .586 1.414l8.704 8.704a2.426 2.426 0 0 0 3.42 0l6.58-6.58a2.426 2.426 0 0 0 0-3.42z"/><circle cx="7.5" cy="7.5" r="1" fill="currentColor"/>',
  circleCheck: '<path d="M20 6 9 17l-5-5"/>',
  camera: '<path d="M14.5 4h-5L7 7H4a2 2 0 0 0-2 2v9a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2V9a2 2 0 0 0-2-2h-3l-2.5-3z"/><circle cx="12" cy="13" r="3"/>',
  save: '<path d="M20 6 9 17l-5-5"/>',
  x: '<path d="M18 6 6 18"/><path d="m6 6 12 12"/>',
};
function svgIcon(name, cls = '') {
  return `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" class="${cls}">${SVG_ICON[name] || ''}</svg>`;
}
const state = {
  view: 'dashboard',
  records: [],
  loading: true,
  error: null,
  filters: { search: '', department: '', category: '', status: '' },
  editingRecord: null, // record object when editing, null when adding
  previewImage: null,
  highlightId: null, // id of a just-added/edited record, briefly flashed in the table
};

// Utilities 
function calculateWW(dateStr) {
  if (!dateStr) return '';
  const date = new Date(dateStr);
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  const weekNo = Math.ceil((((d.getTime() - yearStart.getTime()) / 86400000) + 1) / 7);
  return String(weekNo);
}

function esc(str) {
  return String(str ?? '').replace(/[&<>"']/g, c => ({
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#39;'
  }[c]));
}

function statusBadgeClass(status) {
  switch (status) {
    case 'Open': return 'bg-rose-50 text-rose-600 border-rose-100';
    case 'In Progress': return 'bg-amber-50 text-amber-600 border-amber-100';
    case 'Closed': return 'bg-emerald-50 text-emerald-600 border-emerald-100';
    default: return 'bg-slate-50 text-slate-600 border-slate-100';
  }
}

function toast(message, isError = false) {
  const el = document.getElementById('toast');
  el.textContent = message;
  el.className = `fixed bottom-6 right-6 z-[100] px-4 py-3 rounded-xl shadow-lg text-xs font-bold text-white slide-up ${isError ? 'bg-rose-600' : 'bg-slate-900'}`;
  el.classList.remove('hidden');
  clearTimeout(toast._t);
  toast._t = setTimeout(() => el.classList.add('hidden'), 2800);
}

function icons() { try { if (window.lucide) lucide.createIcons(); } catch (e) { console.warn('Icon render issue:', e); } }

// API layer 
const api = {
  async list() {
    const res = await fetch(`${API_BASE}/api/records`);
    if (!res.ok) throw new Error('Failed to load records');
    return res.json();
  },
  async create(record) {
    const res = await fetch(`${API_BASE}/api/records`, {
      method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(record),
    });
    if (!res.ok) throw new Error((await res.json()).error || 'Failed to create record');
    return res.json();
  },
  async update(id, patch) {
    const res = await fetch(`${API_BASE}/api/records/${id}`, {
      method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(patch),
    });
    if (!res.ok) throw new Error((await res.json()).error || 'Failed to update record');
    return res.json();
  },
  async remove(id) {
    const res = await fetch(`${API_BASE}/api/records/${id}`, { method: 'DELETE' });
    if (!res.ok) throw new Error((await res.json()).error || 'Failed to delete record');
    return res.json();
  },
  async uploadImage(file) {
    const fd = new FormData();
    fd.append('image', file);
    const res = await fetch(`${API_BASE}/api/upload`, { method: 'POST', body: fd });
    if (!res.ok) throw new Error('Image upload failed');
    return res.json();
  },
   async importExcel(file) {
  const fd = new FormData();
  fd.append('file', file);
  const res = await fetch(
    `${API_BASE}/api/import`,
    {
      method: 'POST',
      body: fd
    }
  );
  if (!res.ok) {
    throw new Error(
      (await res.json()).error ||
      'Import failed'
    );
  }
  return res.json();
},
};

async function loadRecords() {
  state.loading = true; state.error = null; render();
  try {
    state.records = await api.list();
  } catch (err) {
    state.error = err.message;
  } finally {
    state.loading = false; render();
  }
}

// Root render
function render() {
  const app = document.getElementById('app');
  app.innerHTML = `
    ${renderSidebar()}
    <div class="flex-1 flex flex-col overflow-hidden">
      ${renderHeader()}
      <main class="flex-1 overflow-y-auto p-6 bg-bg-main">
        ${state.loading ? renderLoading() : state.error ? renderError() : renderCurrentView()}
      </main>
    </div>
    ${state.previewImage ? renderImageLightbox() : ''}
  `;
  icons();
  bindEvents();
  if (state.view === 'dashboard' && !state.loading && !state.error) drawCharts();
}

function renderLoading() {
  return `
  <div class="flex flex-col items-center justify-center h-full gap-3">
    <div class="w-10 h-10 rounded-full border-[3px] border-slate-200 border-t-brand-orange animate-spin"></div>
    <div class="text-text-muted text-xs font-bold uppercase tracking-widest">Loading records…</div>
  </div>`;
}
function renderError() {
  return `
  <div class="flex flex-col items-center justify-center h-full gap-5 max-w-md mx-auto text-center">
    <div class="w-14 h-14 rounded-2xl bg-rose-50 flex items-center justify-center text-rose-500">
      <i data-lucide="${ICONS.wifiOff}" class="w-7 h-7"></i>
    </div>
    <div>
      <div class="text-base font-black text-slate-800">Can't reach the backend</div>
      <div class="text-xs text-text-muted mt-1.5 leading-relaxed">${esc(state.error)}</div>
    </div>
    <div class="w-full bg-white border border-border-subtle rounded-xl p-4 text-left space-y-2.5">
      <div class="text-[9px] font-black text-slate-400 uppercase tracking-[0.15em]">Quick checklist</div>
      <label class="flex items-start gap-2.5 text-[11px] text-text-main font-semibold">
        <span class="w-4 h-4 mt-0.5 shrink-0 rounded border border-slate-300"></span>
        Backend is running — <span class="font-mono bg-slate-50 px-1 rounded">npm start</span> inside <span class="font-mono bg-slate-50 px-1 rounded">backend/</span>
      </label>
      <label class="flex items-start gap-2.5 text-[11px] text-text-main font-semibold">
        <span class="w-4 h-4 mt-0.5 shrink-0 rounded border border-slate-300"></span>
        <span class="font-mono bg-slate-50 px-1 rounded">API_BASE</span> in <span class="font-mono bg-slate-50 px-1 rounded">config.js</span> matches that server's URL
      </label>
      <label class="flex items-start gap-2.5 text-[11px] text-text-main font-semibold">
        <span class="w-4 h-4 mt-0.5 shrink-0 rounded border border-slate-300"></span>
        Currently pointing at <span class="font-mono bg-slate-50 px-1 rounded">${esc(API_BASE)}</span>
      </label>
    </div>
    <button data-action="retry" class="px-5 py-2.5 bg-slate-900 hover:bg-slate-800 text-white text-xs font-black uppercase tracking-widest rounded-lg transition-colors">Retry connection</button>
  </div>`;
}

function renderCurrentView() {
  switch (state.view) {
    case 'dashboard':
      return renderDashboard();
    case 'ipqc':
      return renderIPQCList();
    case 'add-audit':
      return renderAddAuditForm();
    case 'import':
      return renderImportPage();
    case 'settings':
      return renderSettings();
    default:
      return renderDashboard();
  }
}

// Sidebar / Header 
function renderSidebar() {
  const items = [
    { id: 'dashboard', label: 'Dashboard', icon: ICONS.dashboard },
    { id: 'add-audit', label: 'Add Finding', icon: ICONS.add },
    { id: 'ipqc', label: 'IPQC Records', icon: ICONS.ipqc },
    { id: 'import', label: 'Import Excel', icon: 'upload' },
    { id: 'settings', label: 'Settings', icon: ICONS.settings },
  ];
  return `
  <aside class="w-56 shrink-0 bg-sidebar-bg flex flex-col">
    <div class="h-16 flex items-center gap-2 px-5 border-b border-white/5">
      <div class="w-8 h-8 rounded-lg bg-brand-orange flex items-center justify-center text-white font-black text-xs">Q</div>
      <span class="text-white font-black text-sm tracking-wide">IPQC TRACKER</span>
    </div>
    <nav class="flex-1 py-4 px-2.5 space-y-0.5">
      ${items.map(it => `
        <button data-nav="${it.id}" class="w-full flex items-center gap-3 px-3.5 py-2.5 rounded-lg transition-all text-[11px] font-semibold ${state.view === it.id ? 'bg-sidebar-active text-white' : 'text-slate-400 hover:text-slate-100 hover:bg-slate-800/50'}">
          <span class="${state.view === it.id ? 'text-brand-orange' : 'text-slate-500'}"><i data-lucide="${it.icon}" class="w-4 h-4"></i></span>
          <span class="tracking-wide uppercase">${it.label}</span>
          ${state.view === it.id ? `<span class="ml-auto w-1.5 h-1.5 rounded-full bg-brand-orange"></span>` : ''}
        </button>
      `).join('')}
    </nav>
    <div class="p-4 border-t border-white/5">
      <a href="${API_BASE}/api/download" class="flex items-center gap-2 text-[10px] font-bold text-slate-400 hover:text-white transition-colors">
        <i data-lucide="${ICONS.download}"></i> Download Excel
      </a>
    </div>
  </aside>`;
}

function renderHeader() {
 const titles = {
  dashboard: 'Dashboard',
  ipqc: 'IPQC Records',
  import: 'Import Excel',
  'add-audit': 'Add Finding',
  settings: 'Settings'
};
  return `
  <header class="h-16 shrink-0 bg-white border-b border-border-subtle flex items-center justify-between px-6">
    <h1 class="text-sm font-black text-slate-800 uppercase tracking-wider">${titles[state.view] || ''}</h1>
    <div class="text-[10px] font-bold text-text-muted uppercase tracking-widest">${state.records.length} total records</div>
  </header>`;
}

//Dashboard 
function renderDashboard() {
  const recs = state.records;

  const locked =
    recs.filter(r => r.icarStatus === 'Locked').length;

  const submitted =
    recs.filter(r => r.icarStatus === 'Submitted').length;

  if (recs.length === 0) return renderEmptyDashboard();

  return `
  <div class="space-y-6 fade-in">
  <div class="grid grid-cols-2 md:grid-cols-3 gap-4">
  ${statCard('Total Records', recs.length, ICONS.clipboardList, 'slate')}
  ${statCard('Locked', locked, ICONS.clock, 'amber')}
  ${statCard('Submitted', submitted, ICONS.circleCheck, 'emerald')}
</div>
    <div class="grid grid-cols-1 lg:grid-cols-2 gap-4">
      <div class="bg-white p-5 rounded-xl border border-border-subtle shadow-sm">
        <div class="text-[10px] font-bold text-text-muted uppercase tracking-widest mb-3">Findings by Department</div>
        <canvas id="chartDept" height="220"></canvas>
      </div>
      <div class="bg-white p-5 rounded-xl border border-border-subtle shadow-sm">
        <div class="text-[10px] font-bold text-text-muted uppercase tracking-widest mb-3">Status Distribution</div>
        <canvas id="chartStatus" height="220"></canvas>
      </div>
    </div>
    <div class="bg-white p-5 rounded-xl border border-border-subtle shadow-sm">
      <div class="text-[10px] font-bold text-text-muted uppercase tracking-widest mb-3">Findings by Work Week</div>
      <canvas id="chartTrend" height="180"></canvas>
    </div>
  </div>`;
}

function renderEmptyDashboard() {
  return `
  <div class="flex flex-col items-center justify-center h-full gap-4 text-center fade-in">
    <div class="w-14 h-14 rounded-2xl bg-slate-100 flex items-center justify-center text-slate-400">
      <i data-lucide="${ICONS.inbox}" class="w-7 h-7"></i>
    </div>
    <div>
      <div class="text-sm font-black text-slate-800">No audit records yet</div>
      <div class="text-xs text-text-muted mt-1">Submit your first audit, or import your existing tracker on the backend.</div>
    </div>
    <button data-nav="add-audit" class="px-4 py-2.5 bg-brand-orange text-white text-xs font-black uppercase tracking-widest rounded-lg">Add an Audit</button>
  </div>`;
}

const STAT_ACCENTS = {
  slate: { chip: 'bg-slate-100 text-slate-500', ring: '' },
  rose: { chip: 'bg-rose-50 text-rose-500', ring: '' },
  amber: { chip: 'bg-amber-50 text-amber-500', ring: '' },
  emerald: { chip: 'bg-emerald-50 text-emerald-500', ring: '' },
};

function statCard(label, value, icon, accent = 'slate') {
  const a = STAT_ACCENTS[accent] || STAT_ACCENTS.slate;
  return `<div class="bg-white p-5 rounded-xl border border-border-subtle shadow-sm flex items-start justify-between">
    <div>
      <div class="text-[10px] text-text-muted font-bold uppercase tracking-[0.1em]">${esc(label)}</div>
      <div class="text-3xl font-black text-slate-900 mt-2 tracking-tight font-mono">${value}</div>
    </div>
    <div class="w-9 h-9 rounded-lg ${a.chip} flex items-center justify-center shrink-0">
      <i data-lucide="${icon}" class="w-4 h-4"></i>
    </div>
  </div>`;
}

let charts = {};
function drawCharts() {
  Object.values(charts).forEach(c => c && c.destroy());

  const recs = state.records;

  // Department
  const byDept = {};
  recs.forEach(r => {
    byDept[r.department || 'Unspecified'] =
      (byDept[r.department || 'Unspecified'] || 0) + 1;
  });

  const deptCtx = document.getElementById('chartDept');

  if (deptCtx) {
    charts.department = new Chart(deptCtx, {
      type: 'bar',
      data: {
        labels: Object.keys(byDept),
        datasets: [{
          data: Object.values(byDept),
          backgroundColor: '#F15D22'
        }]
      },
      options: {
        plugins: {
          legend: {
            display: false
          }
        }
      }
    });
  }

  // Status
  const byStatus = {
    Locked: 0,
    Submitted: 0
  };

  recs.forEach(r => {
    if (byStatus[r.icarStatus] !== undefined) {
      byStatus[r.icarStatus]++;
    }
  });

  const statusCtx = document.getElementById('chartStatus');

  if (statusCtx) {
    charts.status = new Chart(statusCtx, {
      type: 'doughnut',
      data: {
        labels: Object.keys(byStatus),
        datasets: [{
          data: Object.values(byStatus),
          backgroundColor: [
            '#f59e0b',
            '#10b981'
          ]
        }]
      }
    });
  }

  // Trend by WW
  const byWW = {};

  recs.forEach(r => {
    byWW[r.ww || '?'] =
      (byWW[r.ww || '?'] || 0) + 1;
  });

  const wwLabels = Object.keys(byWW)
    .sort((a, b) => Number(a) - Number(b));

  const trendCtx = document.getElementById('chartTrend');

  if (trendCtx) {
    charts.trend = new Chart(trendCtx, {
      type: 'line',
      data: {
        labels: wwLabels.map(w => `WW${w}`),
        datasets: [{
          label: 'Findings',
          data: wwLabels.map(w => byWW[w]),
          borderColor: '#F15D22',
          backgroundColor: 'rgba(241,93,34,0.1)',
          fill: true,
          tension: 0.3
        }]
      },
      options: {
        plugins: {
          legend: {
            display: false
          }
        },
        scales: {
          y: {
            beginAtZero: true,
            ticks: {
              precision: 0
            }
          }
        }
      }
    });
  }
}

//  IPQC List / History
function filteredRecords() {
  const f = state.filters;
  return state.records
    .filter(r => {
      if (f.department && r.department !== f.department) return false;
      if (f.category && r.category !== f.category) return false;
     if (
  f.status &&
  r.icarStatus !== f.status
) {
  return false;
}
      if (f.search) {
        const q = f.search.toLowerCase();
        const hay = `${r.auditors} ${r.personOnJob} ${r.detailsFindings} ${r.areaStation} ${r.icarNum}`.toLowerCase();
        if (!hay.includes(q)) return false;
      }
      return true;
    })
    .sort((a, b) => new Date(b.auditDate) - new Date(a.auditDate));
}

function renderFilterBar() {
  const f = state.filters;
  const recs = state.records;
 const counts = {
  '': recs.length,
  Locked: recs.filter(
    r => r.icarStatus === 'Locked'
  ).length,

  Submitted: recs.filter(
    r => r.icarStatus === 'Submitted'
  ).length,
};
 const chipStyle = (s) =>
  s === f.status
    ? (
        s === 'Locked'
          ? 'bg-amber-500 text-white'
          : s === 'Submitted'
          ? 'bg-emerald-500 text-white'
          : 'bg-slate-900 text-white'
      )
    : 'bg-slate-100 text-slate-500 hover:bg-slate-200';
  return `
  <div class="space-y-3">
    <div class="flex flex-wrap gap-2">
      ${['', 'Locked', 'Submitted'].map(s => `
        <button data-status-chip="${s}" class="px-3.5 py-1.5 rounded-full text-[10px] font-black uppercase tracking-wider transition-colors ${chipStyle(s)}">
          ${s === '' ? 'All' : s} <span class="opacity-70">${counts[s]}</span>
        </button>
      `).join('')}
    </div>
    <div class="bg-white p-4 rounded-xl border border-border-subtle shadow-sm flex flex-wrap gap-3 items-end">
      <div class="flex-1 min-w-[180px]">
        <label class="text-[9px] font-black text-slate-400 uppercase tracking-[0.2em]">Search</label>
        <input id="filterSearch" type="text" value="${esc(f.search)}" placeholder="Search findings, auditor, ICAR…"
          class="w-full mt-1 bg-slate-50 border border-slate-200 rounded-lg py-2 px-3 text-[11px] font-semibold focus:outline-none focus:border-brand-orange" />
      </div>
      <div>
        <label class="text-[9px] font-black text-slate-400 uppercase tracking-[0.2em]">Department</label>
        <select id="filterDepartment" class="mt-1 bg-slate-50 border border-slate-200 rounded-lg py-2 px-3 text-[11px] font-semibold">
          <option value="">All</option>
          ${DEPARTMENTS.map(d => `<option value="${esc(d)}" ${f.department === d ? 'selected' : ''}>${esc(d)}</option>`).join('')}
        </select>
      </div>
      <div>
        <label class="text-[9px] font-black text-slate-400 uppercase tracking-[0.2em]">Category</label>
        <select id="filterCategory" class="mt-1 bg-slate-50 border border-slate-200 rounded-lg py-2 px-3 text-[11px] font-semibold">
          <option value="">All</option>
          ${CATEGORIES.map(c => `<option value="${esc(c)}" ${f.category === c ? 'selected' : ''}>${esc(c)}</option>`).join('')}
        </select>
      </div>
    </div>
  </div>`;
}

function renderTable(records, { showActions = true } = {}) {
  if (records.length === 0) {
    return `
    <div class="bg-white p-12 rounded-xl border border-border-subtle text-center flex flex-col items-center gap-3">
      <div class="w-11 h-11 rounded-xl bg-slate-100 flex items-center justify-center text-slate-400">
        <i data-lucide="${ICONS.search}" class="w-5 h-5"></i>
      </div>
      <div class="text-xs font-bold text-slate-700">No records match your filters</div>
      <div class="text-[11px] text-text-muted">Try clearing a filter or adjusting your search.</div>
    </div>`;
  }
  return `
  <div class="bg-white rounded-xl border border-border-subtle shadow-sm overflow-auto max-h-[65vh]">
    <table class="w-full text-[11px]">
      <thead>
        <tr class="border-b border-border-subtle text-text-muted uppercase text-[9px] font-black tracking-widest bg-slate-50/80 sticky top-0">
          <th class="text-left px-3 py-3">No</th>
          <th class="text-left px-3 py-3">Date</th>
          <th class="text-left px-3 py-3">WW</th>
          <th class="text-left px-3 py-3">Shift</th>
          <th class="text-left px-3 py-3">Auditor</th>
          <th class="text-left px-3 py-3">Department</th>
          <th class="text-left px-3 py-3">Platform</th>
          <th class="text-left px-3 py-3">Findings</th>
          <th class="text-left px-3 py-3">ICAR</th>
          <th class="text-left px-3 py-3">ICAR Status</th>
          <th class="text-left px-3 py-3">Photo</th>
          ${showActions ? '<th class="text-left px-3 py-3">Actions</th>' : ''}
        </tr>
      </thead>
      <tbody>
        ${records.map(r => `
          <tr class="border-b border-border-subtle/60 odd:bg-white even:bg-slate-50/40 hover:bg-orange-50/50 transition-colors ${String(r.id) === String(state.highlightId) ? 'row-flash' : ''}">
            <td class="px-3 py-2 font-bold">${esc(r.no)}</td>
            <td class="px-3 py-2">${esc(r.auditDate)}</td>
            <td class="px-3 py-2">${esc(r.ww)}</td>
            <td class="px-3 py-2">${esc(r.shift)}</td>
            <td class="px-3 py-2 font-semibold">${esc(r.auditors)}</td>
            <td class="px-3 py-2">${esc(r.department)}</td>
            <td class="px-3 py-2">${esc(r.platform)}</td>
            <td class="px-3 py-2 max-w-[240px] truncate" title="${esc(r.detailsFindings)}">${esc(r.detailsFindings)}</td>

<td class="px-3 py-2">
  ${esc(r.icarNum)}
</td>

<td class="px-3 py-2">
  ${esc(r.icarStatus || 'Locked')}
</td>

<td class="px-3 py-2">
  ${r.picture
    ? `<button data-preview="${esc(r.picture)}" class="text-brand-orange">
         <i data-lucide="${ICONS.image}"></i>
       </button>`
    : '—'}
</td>
            ${showActions ? `
            <td class="px-3 py-2">
              <div class="flex gap-2">
                <button data-edit="${esc(r.id)}" class="text-slate-400 hover:text-brand-orange"><i data-lucide="${ICONS.edit}"></i></button>
                <button data-delete="${esc(r.id)}" class="text-slate-400 hover:text-rose-600"><i data-lucide="${ICONS.trash}"></i></button>
              </div>
            </td>` : ''}
          </tr>
        `).join('')}
      </tbody>
    </table>
  </div>`;
}

function renderIPQCList() {
  return `
  <div class="space-y-4 fade-in">
    <div class="text-xs text-text-muted font-semibold">Every audit record, most recent first — imported history and everything logged through this app, in one place.</div>
    ${renderFilterBar()}
    ${renderTable(filteredRecords())}
  </div>`;
}

// Add / Edit Audit Form 
function blankAuditRecord() {
  return {
    auditDate: new Date().toISOString().split('T')[0],
    shift: 'A',
    auditors: '',
    personOnJob: '',
    department: DEPARTMENTS[0],
    platform: PLATFORMS[0],
    areaStation: '',
    groupFinding: '',
    category: CATEGORIES[0],
    detailsFindings: '',
    picture: '',
    remark: '',
    icarNum: '',
    mqeEngineer: '',
    icarStatus: 'Locked'
  };
}

function auditFormFieldsHtml(r) {

  return `

    <div class="grid lg:grid-cols-4 md:grid-cols-2 grid-cols-1 gap-5">
      ${formInput('auditDate', 'Audit Date', r.auditDate, 'date')}
      ${formInput('ww', 'Work Week (WW)', calculateWW(r.auditDate), 'text')}
      ${formSelect('shift', 'Shift', r.shift, SHIFTS)}
      ${formSelect('department', 'Department', r.department, DEPARTMENTS)}
    </div>

    <div class="grid lg:grid-cols-3 md:grid-cols-2 grid-cols-1 gap-5 mt-4">
      ${formSelect('platform', 'Platform', r.platform, PLATFORMS)}
      ${formInput('areaStation', 'Area / Station', r.areaStation, 'text', true)}
      ${formSelect('category', 'Category', r.category, CATEGORIES)}
    </div>

    <div class="grid lg:grid-cols-2 md:grid-cols-2 grid-cols-1 gap-5 mt-4">
      <div>
        <label class="flex items-center gap-2 mb-3 text-[11px] tracking-[0.12em] uppercase font-black text-[#5d7697]">
          Group Finding
        </label>
        <input
          name="groupFinding"
          id="groupFinding"
          type="text"
          value="${esc(r.groupFinding)}"
          readonly
          class="w-full h-12 px-4 rounded-2xl bg-slate-100 border border-slate-300 text-base font-medium text-slate-700 focus:outline-none"
        />
      </div>
      ${formSelect('detailsFindings', 'Finding Details', r.detailsFindings, (r.detailsFindings && !FINDING_DETAILS.includes(r.detailsFindings) ? [r.detailsFindings, ...FINDING_DETAILS] : FINDING_DETAILS))}
    </div>

    <div class="grid lg:grid-cols-2 md:grid-cols-2 grid-cols-1 gap-5 mt-4">
      ${formInput('auditors', 'IPQC Auditor Name', r.auditors, 'text')}
      ${formInput('personOnJob', 'PIC Name (Finding)', r.personOnJob, 'text')}
    </div>

    <div class="grid lg:grid-cols-2 md:grid-cols-2 grid-cols-1 gap-5 mt-4">
<div class="grid grid-cols-2 gap-5">

  ${formInput(
    'icarNum',
    'ICAR#',
    r.icarNum || 'N/A'
  )}

${formSelect(
  'icarStatus',
  'ICAR Status',
  r.icarStatus || 'Locked',
  ['Locked', 'Submitted']
)}
</div>

      ${formTextarea('remark', 'Remark', r.remark, 4)}
    </div>

    <div class="mt-4">
      <label class="block mb-3 text-[11px] tracking-[0.12em] font-black uppercase text-[#5d7697]">
        Audit Evidence Picture
      </label>
      <label
        id="imageDropZone"
        for="imageInput"
        class="h-[120px] rounded-2xl border-2 border-dashed border-slate-300 bg-[#f2f5f8] flex flex-col items-center justify-center cursor-pointer hover:border-blue-300"
      >
        ${r.picture ? `<img src="${esc(r.picture)}" class="w-full h-full object-cover rounded-[10px]" />` : `<div class="text-center"><i data-lucide="image" class="w-14 h-14 text-slate-400 mx-auto"></i><div class="font-black text-slate-600 uppercase mt-4">Drag & Drop or Click to Upload</div><div class="text-slate-400 text-sm mt-2">Supports JPG, PNG, WEBP</div></div>`}
      </label>
      <input id="imageInput" type="file" accept="image/*" class="hidden" />
      <input type="hidden" name="picture" value="${esc(r.picture)}" id="pictureField" />
    </div>
  `;
}

function renderAddAuditForm() {
  const r = state.editingRecord || blankAuditRecord();
  const isEdit = !!state.editingRecord;

  return `
  <div class="fade-in max-w-6xl mx-auto">
    <div class="bg-[#f7f9fc] border border-slate-200 rounded-2xl overflow-hidden">

      <div class="px-6 py-6 border-b border-slate-200">
        <div class="flex justify-between items-start">

          <div>
            <h1 class="text-3xl font-black text-slate-900">
              ${isEdit ? 'Edit Audit Entry' : 'New Audit Entry'}
            </h1>
          </div>

          <button
            data-nav="ipqc"
            class="text-slate-500 hover:text-slate-800 text-sm font-bold uppercase"
          >
            ✕ Exit
          </button>

        </div>
      </div>

      <form id="auditForm" class="p-6">

        ${auditFormFieldsHtml(r)}

<div class="mt-6 flex gap-3">

  <button
    type="submit"
    class="px-6 py-3 bg-[#0f172a] hover:bg-slate-800 text-white rounded-xl font-bold"
  >
    ${isEdit ? 'Save Changes' : 'Submit'}
  </button>

  <button
    type="button"
    data-nav="ipqc"
    class="px-6 py-3 border border-slate-300 rounded-xl font-bold"
  >
    Cancel
  </button>

</div>

      </form>

    </div>
  </div>
  `;
}

function formSection(title, iconKey, innerHtml) {
  return `
  <div class="bg-white rounded-xl border border-border-subtle shadow-sm overflow-hidden">
    <div class="flex items-center gap-2.5 px-5 py-3.5 border-b border-border-subtle bg-slate-50/60">
      <div class="w-6 h-6 rounded-md bg-orange-50 text-brand-orange flex items-center justify-center">
        ${svgIcon(iconKey, 'w-3.5 h-3.5')}
      </div>
      <div class="text-[11px] font-black text-slate-600 uppercase tracking-[0.12em]">${esc(title)}</div>
    </div>
    <div class="p-5">${innerHtml}</div>
  </div>`;
}

function formInput(
  name,
  label,
  value,
  type = 'text',
  required = false
) {

  return `
  <div>

    <label
      class="
      flex
      items-center
      gap-2
      mb-3
      text-[11px]
      tracking-[0.12em]
      uppercase
      font-black
      text-[#5d7697]
      "
    >
      ${label}

      ${
        required
          ? `
            <span class="
              text-[9px]
              px-1.5
              py-0.5
              rounded-md
              bg-red-100
              text-red-500
            ">
              REQUIRED
            </span>
          `
          : ''
      }
    </label>

    <input
      name="${name}"
      type="${type}"
      value="${esc(value)}"
      ${required ? 'required' : ''}
      class="
      w-full
      h-12
      px-4
      rounded-2xl
      bg-[#f2f5f8]
      border
      border-slate-300
      text-base
      font-medium
      text-slate-900
      focus:outline-none
      focus:border-blue-400
      "
    />

  </div>
  `;
}
function formSelect(
  name,
  label,
  value,
  options,
  required = false
) {

  return `
  <div>

    <label
      class="
      flex
      items-center
      gap-2
      mb-3
      text-[11px]
      tracking-[0.12em]
      uppercase
      font-black
      text-[#5d7697]
      "
    >
      ${label}

      ${
        required
          ? `
            <span class="
              text-[9px]
              px-1.5
              py-0.5
              rounded-md
              bg-red-100
              text-red-500
            ">
              REQUIRED
            </span>
          `
          : ''
      }
    </label>

    <select
      name="${name}"
      ${required ? 'required' : ''}
      class="
      w-full
      h-12
      px-4
      rounded-2xl
      bg-[#f2f5f8]
      border
      border-slate-300
      text-base
      font-medium
      text-slate-900
      focus:outline-none
      focus:border-blue-400
      "
    >
      ${options.map(o => `
        <option
          value="${esc(o)}"
          ${value === o ? 'selected' : ''}
        >
          ${esc(o)}
        </option>
      `).join('')}
    </select>

  </div>
  `;
}
function formTextarea(
  name,
  label,
  value,
  rows = 4,
  required = false
) {

  return `
  <div>

    <label
      class="
      block
      mb-3
      text-[11px]
      tracking-[0.12em]
      uppercase
      font-black
      text-[#5d7697]
      "
    >
      ${label}
    </label>

    <textarea
      name="${name}"
      rows="${rows}"
      ${required ? 'required' : ''}
      class="
      w-full
      rounded-2xl
      p-4
      bg-[#f2f5f8]
      border
      border-slate-300
      text-sm
      resize-none
      focus:outline-none
      focus:border-blue-400
      "
    >${esc(value)}</textarea>

  </div>
  `;
}
function renderImportPage() {
  return `
  <div class="max-w-xl">

    <div class="bg-white p-6 rounded-xl border">

      <h2 class="text-xl font-bold mb-4">
        Import Historical Records
      </h2>

      <input
        id="excelImport"
        type="file"
        accept=".xlsx"
        class="mb-4"
      />

      <button
        id="importBtn"
        class="px-4 py-2 bg-brand-orange text-white rounded-lg"
      >
        Import Excel
      </button>

    </div>

  </div>
  `;
}

// Settings
function renderSettings() {
  return `
  <div class="max-w-xl space-y-4 fade-in">
    <div class="bg-white p-5 rounded-xl border border-border-subtle shadow-sm space-y-3">
      <div class="text-xs font-black text-slate-500 uppercase tracking-widest">Connection</div>
      <div class="text-[11px] text-text-muted">API base URL (set in <code>config.js</code>):</div>
      <div class="font-mono text-xs bg-slate-50 border border-slate-200 rounded-lg px-3 py-2">${esc(API_BASE)}</div>
      <button id="healthCheckBtn" class="px-4 py-2 bg-slate-900 text-white text-[11px] font-bold rounded-lg">Test Connection</button>
      <div id="healthResult" class="text-xs font-bold"></div>
    </div>
    <div class="bg-white p-5 rounded-xl border border-border-subtle shadow-sm space-y-2">
      <div class="text-xs font-black text-slate-500 uppercase tracking-widest">Data Source</div>
      <p class="text-[11px] text-text-muted leading-relaxed">All records — historical and newly submitted — live in a single Excel file on the backend. Power BI should point at
        <span class="font-mono">${esc(API_BASE)}/api/download</span> (Get Data → Web) with scheduled refresh so it always reflects the latest data.</p>
      <a href="${API_BASE}/api/download" class="inline-block mt-1 px-4 py-2 bg-brand-orange text-white text-[11px] font-bold rounded-lg">Download Current Excel File</a>
    </div>
  </div>`;
}

// Image Lightbox 
function renderImageLightbox() {
  return `
  <div id="lightboxOverlay" class="fixed inset-0 bg-black/80 z-[90] flex items-center justify-center p-6">
    <button data-close-preview class="absolute top-6 right-6 text-white/70 hover:text-white">${svgIcon('x', 'w-5 h-5')}</button>
    <img src="${esc(state.previewImage)}" class="max-w-[90vw] max-h-[85vh] rounded-xl shadow-2xl" />
  </div>`;
}

//Event binding 
function bindEvents() {
  document.querySelectorAll('[data-nav]').forEach(btn => {
    btn.addEventListener('click', () => {
      state.view = btn.getAttribute('data-nav');
      if (state.view === 'add-audit') state.editingRecord = null;
      render();
    });
  });

  const categoryField = document.querySelector('select[name="category"]');
  const groupField = document.querySelector('input[name="groupFinding"]');
  if (categoryField && groupField) {
    const syncGroupFinding = () => {
      groupField.value = CATEGORY_GROUP_MAPPING[categoryField.value] || '';
    };
    syncGroupFinding();
    categoryField.addEventListener('change', syncGroupFinding);
  }

  const retryBtn = document.querySelector('[data-action="retry"]');
  if (retryBtn) retryBtn.addEventListener('click', loadRecords);

  // status filter chips
  document.querySelectorAll('[data-status-chip]').forEach(b => b.addEventListener('click', () => {
    state.filters.status = b.getAttribute('data-status-chip'); render();
  }));

  // filters
  ['filterSearch', 'filterDepartment', 'filterCategory'].forEach(id => {
    const el = document.getElementById(id);
    if (!el) return;
    const key = id.replace('filter', '').replace(/^\w/, c => c.toLowerCase());
    el.addEventListener('input', () => { state.filters[key] = el.value; render(); });
    el.addEventListener('change', () => { state.filters[key] = el.value; render(); });
  });

  // table actions
  document.querySelectorAll('[data-preview]').forEach(b => b.addEventListener('click', () => {
    state.previewImage = b.getAttribute('data-preview'); render();
  }));
  const closePreview = document.querySelector('[data-close-preview]');
  if (closePreview) closePreview.addEventListener('click', () => { state.previewImage = null; render(); });

  document.querySelectorAll('[data-edit]').forEach(b => b.addEventListener('click', () => {
    const id = b.getAttribute('data-edit');
    state.editingRecord = state.records.find(r => String(r.id) === String(id));
    state.view = 'add-audit'; render();
  }));

  document.querySelectorAll('[data-delete]').forEach(b => b.addEventListener('click', async () => {
    const id = b.getAttribute('data-delete');
    if (!confirm('Delete this audit record? This cannot be undone.')) return;
    try {
      await api.remove(id);
      state.records = state.records.filter(r => String(r.id) !== String(id));
      toast('Record deleted');
      render();
    } catch (err) { toast(err.message, true); }
  }));

  // image upload
  const imageInput = document.getElementById('imageInput');
  if (imageInput) imageInput.addEventListener('change', async () => {
    const file = imageInput.files[0];
    if (!file) return;
    const dropZone = document.getElementById('imageDropZone');
    try {
      toast('Uploading image…');
      const { url } = await api.uploadImage(file);
      document.getElementById('pictureField').value = url;
      if (dropZone) dropZone.innerHTML = `<img src="${esc(url)}" class="w-full h-full object-cover rounded-[10px]" />`;
      const captionEl = document.querySelector('#imageDropZone')?.parentElement?.querySelector('.font-bold');
      if (captionEl) captionEl.textContent = 'Photo attached';
      toast('Image uploaded');
    } catch (err) { toast(err.message, true); }
  });

  // form submit
  const form = document.getElementById('auditForm');
 
if (form) {
  form.addEventListener('submit', async (e) => {

    e.preventDefault();

    const fd = new FormData(form);
    const payload = Object.fromEntries(fd.entries());

payload.groupFinding =
  CATEGORY_GROUP_MAPPING[payload.category] || '';

payload.ww =
  calculateWW(payload.auditDate);

// determine ICAR status
payload.icarStatus =
  payload.icarStatus || 'Locked';
    // IMPORTANT:
    // Preserve ID and No during edit
    if (state.editingRecord) {
      payload.id = state.editingRecord.id;
      payload.no = state.editingRecord.no;
    }

    if (!payload.mqeEngineer) {
      payload.mqeEngineer =
        PLATFORM_MQE_MAPPING[payload.platform] || '';
    }

try {

  let savedRecord;

  if (state.editingRecord) {

    console.log('UPDATE MODE', state.editingRecord);

    savedRecord = await api.update(
      state.editingRecord.id,
      payload
    );

    state.records = state.records.map(r =>
      String(r.id) === String(savedRecord.id)
        ? savedRecord
        : r
    );

    toast('Record updated');

  } else {

    console.log('CREATE MODE');

    savedRecord = await api.create(payload);

    state.records.push(savedRecord);

    toast('Audit submitted');
  }

      state.editingRecord = null;

      // Clear filters
      state.filters = {
        search: '',
        department: '',
        category: '',
        status: ''
      };

      state.view = 'ipqc';

      state.highlightId = savedRecord.id;

      render();

      setTimeout(() => {
        state.highlightId = null;
        render();
      }, 2000);

    } catch (err) {
      console.error(err);
      toast(err.message, true);
    }
  });

  // Import Excel
  const importBtn = document.getElementById('importBtn');
  if (importBtn) {
    importBtn.addEventListener('click', async () => {
      const fileInput = document.getElementById('excelImport');
      const file = fileInput?.files?.[0];
      if (!file) {
        toast('Please select an Excel file', true);
        return;
      }
      try {
        toast('Importing records...');
        const result = await api.importExcel(file);
        toast(`Imported ${result.imported} records`);
        await loadRecords();
        state.view = 'ipqc';
        render();
      } catch (err) {
        console.error(err);
        toast(err.message, true);
      }
    });
  }
}
  // settings
  const healthBtn = document.getElementById('healthCheckBtn');
  if (healthBtn) healthBtn.addEventListener('click', async () => {
    const resultEl = document.getElementById('healthResult');
    resultEl.textContent = 'Checking…'; resultEl.className = 'text-xs font-bold text-text-muted';
    try {
      const res = await fetch(`${API_BASE}/api/health`);
      if (!res.ok) throw new Error('Unreachable');
      resultEl.textContent = '✓ Connected'; resultEl.className = 'text-xs font-bold text-emerald-600';
    } catch {
      resultEl.textContent = '✗ Could not reach backend'; resultEl.className = 'text-xs font-bold text-rose-600';
    }
  });
}

// Init 
loadRecords();
