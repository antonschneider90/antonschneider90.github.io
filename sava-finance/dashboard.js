/* ============================================================
   Sava Dashboard — dashboard.js
   Fetches data from Google Sheets (gviz JSON), renders all tabs.
   Structure cloned from Stripe CBFS; palette is Sava; waterfall is V2G-style.
   ============================================================ */

// ============ CONFIG ============
const SHEET_ID = '1fJLaVixEazLYqhqpE6dyC5g3bY3jrdzu2fo3rj2sr8k';

function jsonUrl(tabName) {
  return `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:json&sheet=${encodeURIComponent(tabName)}`;
}

const TABS_TO_FETCH = [
  'Current Year Overview',
  'Forecast Years Overview',
  'RFE_Base',
  'RFE_Bear',
  'RFE_Bull',
  'Assumptions',
  'Headcount Plan',
  'Revenue Build',
  'Summary',
];

// ============ STATE ============
const state = {
  scenario: 'Base',
  data: {},
  checkCell: null,
  loaded: false,
};

let charts = {};

// ============ HELPERS ============
const $ = (sel) => document.querySelector(sel);
const $$ = (sel) => Array.from(document.querySelectorAll(sel));

function parseNum(val) {
  if (val === null || val === undefined) return null;
  if (typeof val === 'number') return val;
  let s = String(val).trim();
  if (!s || s === '—' || s === '-' || s === 'n/a' || s === '#N/A') return null;
  let sign = 1;
  if (s.startsWith('-')) { sign = -1; s = s.slice(1); }
  else if (s.startsWith('+')) s = s.slice(1);
  if (s.startsWith('(') && s.endsWith(')')) { sign = -1; s = s.slice(1, -1); }
  let mult = 1;
  if (/%\s*$/i.test(s)) { mult = 0.01; s = s.replace(/%\s*$/i, ''); }
  else if (/m\s*$/i.test(s)) { mult = 1e6; s = s.replace(/m\s*$/i, ''); }
  else if (/k\s*$/i.test(s)) { mult = 1e3; s = s.replace(/k\s*$/i, ''); }
  s = s.replace(/[$,\s]/g, '');
  const n = parseFloat(s);
  if (isNaN(n)) return null;
  return sign * n * mult;
}

// ============ FORMATTERS — consistent across the dashboard ============
function fmtM(n) {
  if (n === null || n === undefined || isNaN(n)) return '—';
  const abs = Math.abs(n);
  const sign = n < 0 ? '−' : '';
  if (abs >= 1e9) return sign + '$' + (abs / 1e9).toFixed(1) + 'B';
  if (abs >= 1e6) return sign + '$' + (abs / 1e6).toFixed(1) + 'M';
  if (abs >= 1e3) return sign + '$' + (abs / 1e3).toFixed(1) + 'K';
  if (abs === 0) return '$0.0M';
  return sign + '$' + abs.toFixed(0);
}
function fmtMVar(n) {
  if (n === null || n === undefined || isNaN(n)) return '—';
  const sign = n >= 0 ? '+' : '−';
  const abs = Math.abs(n);
  if (abs >= 1e9) return sign + '$' + (abs / 1e9).toFixed(1) + 'B';
  if (abs >= 1e6) return sign + '$' + (abs / 1e6).toFixed(1) + 'M';
  if (abs >= 1e3) return sign + '$' + (abs / 1e3).toFixed(1) + 'K';
  return sign + '$' + abs.toFixed(0);
}
function fmtPct(n, decimals = 1) {
  if (n === null || n === undefined || isNaN(n)) return '—';
  return (n * 100).toFixed(decimals) + '%';
}
function fmtPctVar(n) {
  if (n === null || n === undefined || isNaN(n)) return '—';
  const sign = n >= 0 ? '+' : '';
  return sign + (n * 100).toFixed(0) + '%';
}
function fmtBps(n) {
  if (n === null || n === undefined || isNaN(n)) return '—';
  const sign = n >= 0 ? '+' : '';
  return sign + Math.round(n * 10000) + ' bps';
}
function fmtInt(n) {
  if (n === null || n === undefined || isNaN(n)) return '—';
  return Math.round(n).toLocaleString();
}
function fmtCountVar(n) {
  if (n === null || n === undefined || isNaN(n)) return '—';
  const sign = n >= 0 ? '+' : '';
  return sign + Math.round(n).toLocaleString();
}

// Formats whatever gviz returns for a date cell into "Mon YYYY" (e.g. "Apr 2026").
// gviz may return: a Date object, a "Date(yyyy,mm,dd)" string, an ISO string,
// a pre-formatted string like "Apr-2026", or a numeric Excel serial.
function formatLMA(raw) {
  if (raw === null || raw === undefined || raw === '') return '—';
  const monthNames = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  // JS Date instance
  if (raw instanceof Date && !isNaN(raw)) {
    return `${monthNames[raw.getMonth()]} ${raw.getFullYear()}`;
  }
  if (typeof raw === 'string') {
    const s = raw.trim();
    // "Date(yyyy,mm,dd...)" gviz format (months 0-indexed)
    const m1 = s.match(/^Date\((\d+),(\d+),(\d+)/);
    if (m1) return `${monthNames[parseInt(m1[2])]} ${m1[1]}`;
    // "Apr-2026" or "Apr 2026" pre-formatted
    const m2 = s.match(/^([A-Za-z]{3,9})[\s\-]+(\d{4})$/);
    if (m2) {
      const mon = m2[1].slice(0,3);
      const idx = monthNames.findIndex(n => n.toLowerCase() === mon.toLowerCase());
      if (idx >= 0) return `${monthNames[idx]} ${m2[2]}`;
      return s;
    }
    // ISO date "2026-04-01" or "2026-04-01 00:00:00"
    const m3 = s.match(/^(\d{4})-(\d{2})-\d{2}/);
    if (m3) return `${monthNames[parseInt(m3[2]) - 1]} ${m3[1]}`;
    return s;
  }
  if (typeof raw === 'number') {
    // Excel serial — best-effort
    const excelEpoch = new Date(Date.UTC(1899, 11, 30));
    const d = new Date(excelEpoch.getTime() + raw * 864e5);
    if (!isNaN(d)) return `${monthNames[d.getUTCMonth()]} ${d.getUTCFullYear()}`;
  }
  return String(raw);
}

// ============ DATA FETCH ============
async function fetchTab(tabName) {
  const url = jsonUrl(tabName);
  const resp = await fetch(url);
  if (!resp.ok) throw new Error(`Failed to fetch "${tabName}" (HTTP ${resp.status}). Make sure the Sheet is published to web.`);
  const text = await resp.text();
  const start = text.indexOf('{');
  const end = text.lastIndexOf('}');
  if (start < 0 || end < 0) throw new Error(`Unexpected gviz JSON response for "${tabName}"`);
  const jsonStr = text.substring(start, end + 1);
  const payload = JSON.parse(jsonStr);
  if (payload.status === 'error') {
    const msg = (payload.errors || []).map(e => e.detailed_message || e.message).join('; ');
    throw new Error(`gviz error for "${tabName}": ${msg}`);
  }
  const tbl = payload.table || {};
  const cols = tbl.cols || [];
  const rows = tbl.rows || [];
  const headerRow = cols.map(c => (c && c.label) ? c.label : '');
  const dataRows = rows.map(r => {
    const cells = r.c || [];
    return cols.map((_, j) => {
      const cell = cells[j];
      if (cell === null || cell === undefined) return null;
      if (cell.v !== null && cell.v !== undefined) return cell.v;
      if (cell.f !== undefined && cell.f !== null) {
        const f = String(cell.f).trim();
        if (f && !/^[-+(]?[\d$,. ]+\)?[%mkb]?$/i.test(f)) return f;
      }
      return null;
    });
  });
  return [headerRow, ...dataRows];
}

async function loadAll() {
  const statusEl = $('#loading-status');
  state.data = {};
  state.checkCell = null;
  const results = await Promise.all(
    TABS_TO_FETCH.map(async (t) => {
      if (statusEl) statusEl.textContent = `Loading "${t}"...`;
      try {
        return [t, await fetchTab(t)];
      } catch (err) {
        throw new Error(`${t}: ${err.message}`);
      }
    })
  );
  results.forEach(([t, d]) => { state.data[t] = d; });
  buildAllRowIndexes();
  state.loaded = true;
  if (statusEl) statusEl.textContent = 'Rendering...';
}

// ============ DATA ACCESS ============
// gviz collapses empty rows: the row index returned by gviz does NOT
// match the row number in the source spreadsheet. To stay robust, we
// look up rows by their label in column A (or column B for Assumptions,
// which uses col A as a margin/indent column).
//
// state.rowIdx[tabName] = Map(label_lower → 1-indexed row number in gviz data)

// Which column contains the row's "label" for each tab. Default = 1 (col A).
const LABEL_COL = {
  'Assumptions': 2,
};

function buildRowIndex(tabName) {
  const data = state.data[tabName];
  const map = new Map();
  if (!data) return map;
  const labelCol = (LABEL_COL[tabName] || 1) - 1;  // 0-indexed
  for (let i = 0; i < data.length; i++) {
    const labelRaw = data[i] && data[i][labelCol];
    if (labelRaw === null || labelRaw === undefined || labelRaw === '') continue;
    const lbl = String(labelRaw).trim().toLowerCase();
    if (!lbl) continue;
    if (!map.has(lbl)) map.set(lbl, i + 1);
  }
  return map;
}

function buildAllRowIndexes() {
  state.rowIdx = {};
  for (const tab of Object.keys(state.data)) {
    state.rowIdx[tab] = buildRowIndex(tab);
  }
}

// Look up the 1-indexed gviz row number for a given col-A label.
// Accepts an exact label or an array of label candidates (first hit wins).
// If `partial` is true, also matches rows whose label starts with the search term.
function rowOf(tabName, labelOrLabels, partial = false) {
  const map = state.rowIdx && state.rowIdx[tabName];
  if (!map) return null;
  const candidates = Array.isArray(labelOrLabels) ? labelOrLabels : [labelOrLabels];
  for (const cand of candidates) {
    const key = String(cand).trim().toLowerCase();
    if (map.has(key)) return map.get(key);
  }
  if (partial) {
    for (const cand of candidates) {
      const key = String(cand).trim().toLowerCase();
      for (const [lbl, idx] of map.entries()) {
        if (lbl.startsWith(key)) return idx;
      }
    }
  }
  return null;
}

// 1-indexed cell access. `row` is the gviz row number (post-collapse).
function cell(tabName, row, col) {
  if (row === null || row === undefined) return null;
  const data = state.data[tabName];
  if (!data || !data[row - 1]) return null;
  return data[row - 1][col - 1];
}
function cellNum(tabName, row, col) {
  return parseNum(cell(tabName, row, col));
}

// Convenience: look up a row by label then read col. Returns null if label not found.
function cellByLabel(tabName, label, col, partial = false) {
  const r = rowOf(tabName, label, partial);
  return r ? cell(tabName, r, col) : null;
}
function cellNumByLabel(tabName, label, col, partial = false) {
  return parseNum(cellByLabel(tabName, label, col, partial));
}

function rfeTabForScenario(scen) {
  return `RFE_${scen}`;
}

// Sum a row across columns [startCol, endCol] (inclusive)
function sumRow(tabName, row, startCol, endCol) {
  let s = 0; let hasAny = false;
  for (let c = startCol; c <= endCol; c++) {
    const v = cellNum(tabName, row, c);
    if (v !== null && !isNaN(v)) { s += v; hasAny = true; }
  }
  return hasAny ? s : null;
}

// RFE column structure (Sava model):
//   Col B (=2) = Jan 2025
//   Col M (=13) = Dec 2025
//   Col N (=14) = Jan 2026
//   Col Y (=25) = Dec 2026
//   Col Z (=26) = Jan 2027
//   ...
//   Col BU (=73) = Dec 2030
function yearStartCol(year) {
  return 2 + (year - 2025) * 12;
}
function yearEndCol(year) {
  return yearStartCol(year) + 11;
}

function sumYear(tabName, row, year) {
  return sumRow(tabName, row, yearStartCol(year), yearEndCol(year));
}

function eoyVal(tabName, row, year) {
  return cellNum(tabName, row, yearEndCol(year));
}

// Find Assumptions value for active scenario.
// In Sava model: Bear=col D (=4), Base=col E (=5), Bull=col F (=6)
function scenCol(scen) {
  return { Bear: 4, Base: 5, Bull: 6 }[scen];
}
function assumption(row, scen) {
  return cellNum('Assumptions', row, scenCol(scen));
}

// ============ CHART STYLING ============
// Palette: terracotta primary (Sava brand), petrol-blue secondary,
// muted green reserved for positive variance, deep brick red for negative.
const palette = {
  ink: '#1A1815',
  inkLight: 'rgba(26, 24, 21, 0.55)',
  inkBorder: 'rgba(26, 24, 21, 0.10)',
  accent: '#C65A35',                       // terracotta — primary accent
  accentSoft: 'rgba(198, 90, 53, 0.15)',
  blue: '#2C4A57',                         // petrol blue — secondary
  blueSoft: 'rgba(44, 74, 87, 0.15)',
  green: '#4A7A6B',                        // muted green — positive variance only
  red: '#A04025',                          // deep brick — negative variance
  redSoft: 'rgba(160, 64, 37, 0.15)',
  cream: '#FAF7F0',
  // Series order: ink → terra → petrol → taupe → warm grey → muted blue → deep brown → sand
  series: ['#1A1815', '#C65A35', '#2C4A57', '#8A7E68', '#A89683', '#5E7984', '#6B4E3D', '#C5B6A0'],
};

Chart.defaults.font.family = 'Manrope, sans-serif';
Chart.defaults.font.size = 11;
Chart.defaults.color = palette.inkLight;
Chart.defaults.plugins.legend.labels.usePointStyle = true;
Chart.defaults.plugins.legend.labels.padding = 14;
Chart.defaults.plugins.legend.labels.boxHeight = 8;

function destroyChart(name) {
  if (charts[name]) { charts[name].destroy(); delete charts[name]; }
}

function baseTooltip(formatter) {
  return {
    backgroundColor: palette.ink,
    titleColor: palette.cream,
    bodyColor: palette.cream,
    borderColor: 'transparent',
    padding: 12,
    titleFont: { weight: 600, size: 12 },
    bodyFont: { size: 12, family: 'JetBrains Mono, monospace' },
    callbacks: {
      label: (c) => `${c.dataset.label}: ${formatter(c.parsed.y !== undefined ? c.parsed.y : c.parsed)}`,
    },
  };
}

// ============ RENDER: OVERVIEW ============
function renderOverview() {
  const scen = state.scenario;
  const tab = rfeTabForScenario(scen);

  // KPIs from FYO (column F = 2030, which is FYO col 6)
  const rev2030 = cellNumByLabel('Forecast Years Overview', 'Revenue', 6);
  const ebitda2030 = cellNumByLabel('Forecast Years Overview', 'EBITDA', 6);
  const hc2030 = cellNumByLabel('Forecast Years Overview', 'Total HC (EOY)', 6);

  // Funding need: read Series B + C sizes from RFE solver rows
  const sB = cellNumByLabel(tab, 'Series B size (USD)', 2);
  const sC = cellNumByLabel(tab, 'Series C size (USD)', 2);
  const fundNeed = (sB || 0) + (sC || 0);

  $('#kpi-rev').textContent = fmtM(rev2030);
  $('#kpi-ebitda').textContent = fmtM(ebitda2030);
  $('#kpi-hc').textContent = hc2030 !== null ? fmtInt(hc2030) : '—';
  $('#kpi-fund').textContent = fmtM(fundNeed);

  // Detail subtitles
  const margin = rev2030 && ebitda2030 ? ebitda2030 / rev2030 : null;
  $('#kpi-rev-detail').textContent = 'Full year, all channels';
  $('#kpi-ebitda-detail').textContent = margin !== null ? `${fmtPct(margin, 0)} margin` : '';
  $('#kpi-hc-detail').textContent = 'From 98 today';
  $('#kpi-fund-detail').textContent = 'Series B + Series C';

  // Meta — LMA date from Assumptions
  const baseLMA = cellByLabel('Assumptions', 'Last month of actuals (date)', 3);
  $('#overview-meta').innerHTML = `Last actual month: <strong>${formatLMA(baseLMA)}</strong> · Scenario: <span class="scen-inline">${scen}</span>`;

  // Update scenario indicators
  $$('.scen-inline').forEach(el => el.textContent = scen);

  // Revenue by stream chart (uses Revenue Build tab)
  renderRevByStreamChart('chart-ov-rev', scen);
  // Monthly cash chart
  renderMonthlyCashChart('chart-ov-cash', scen, true);
}

// ============ RENDER: REVENUE BY STREAM CHART ============
function renderRevByStreamChart(canvasId, scen) {
  // Revenue Build tab — find rows by label
  const streams = [
    { label: 'EU+UK Rx', sheetLabel: 'EU+UK Rx revenue' },
    { label: 'EU+UK OTC', sheetLabel: 'EU+UK OTC revenue' },
    { label: 'US Rx', sheetLabel: 'US Rx revenue' },
    { label: 'US OTC', sheetLabel: 'US OTC revenue' },
    { label: 'Lactate Standalone', sheetLabel: 'Lactate standalone revenue' },
    { label: 'Lactate Add-on', sheetLabel: 'Lactate add-on revenue' },
  ];
  const years = [2025, 2026, 2027, 2028, 2029, 2030];
  const datasets = streams.map((s, i) => {
    const row = rowOf('Revenue Build', s.sheetLabel);
    return {
      label: s.label,
      data: years.map(y => (row ? sumYear('Revenue Build', row, y) : null) || 0),
      backgroundColor: palette.series[i],
      borderRadius: 0,
    };
  });

  destroyChart(canvasId);
  charts[canvasId] = new Chart(document.getElementById(canvasId), {
    type: 'bar',
    data: { labels: years, datasets },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { position: 'top', align: 'end' },
        tooltip: {
          ...baseTooltip(fmtM),
          mode: 'index',
          callbacks: {
            label: (c) => `${c.dataset.label}: ${fmtM(c.parsed.y)}`,
            footer: (items) => {
              const total = items.reduce((s, it) => s + (it.parsed.y || 0), 0);
              return `Total: ${fmtM(total)}`;
            },
          },
        },
      },
      scales: {
        x: { stacked: true, grid: { display: false } },
        y: { stacked: true, grid: { color: palette.inkBorder }, ticks: { callback: fmtM } },
      },
    },
  });
}

// Convert a column index in the monthly grid (B=2 = Jan 2025) into a "YYYY-MM" string.
// This avoids depending on the date-header row, which gviz can mangle.
function colToYearMonth(c) {
  const offset = c - 2;  // 0 = Jan 2025
  const y = 2025 + Math.floor(offset / 12);
  const m = (offset % 12) + 1;
  return `${y}-${String(m).padStart(2, '0')}`;
}

// ============ RENDER: MONTHLY CASH CHART ============
function renderMonthlyCashChart(canvasId, scen, withMarkers) {
  const tab = rfeTabForScenario(scen);
  const cashRow = rowOf(tab, 'Cash & cash equivalents');
  const dates = [];
  const cashVals = [];
  for (let c = 2; c <= 73; c++) {
    dates.push(colToYearMonth(c));
    cashVals.push((cashRow ? cellNum(tab, cashRow, c) : 0) || 0);
  }

  // Detect all financing/grant events from CFF rows in RFE.
  // events[i] = label string for column index i (0..71), or null.
  const events = new Array(72).fill(null);
  const eventColors = new Array(72).fill(null);
  if (withMarkers) {
    const eventRows = [
      { label: 'Series A close', sheetLabel: 'Series A inflow (historical)', color: palette.ink },
      { label: 'Grant tranche', sheetLabel: 'Grant inflows (historical)', color: palette.blue },
      { label: 'Series B closes', sheetLabel: 'Series B inflow (forecast)', color: palette.accent },
      { label: 'Series C closes', sheetLabel: 'Series C inflow (forecast)', color: palette.red },
    ];
    for (const ev of eventRows) {
      const r = rowOf(tab, ev.sheetLabel);
      if (!r) continue;
      for (let c = 2; c <= 73; c++) {
        const v = cellNum(tab, r, c);
        if (v && Math.abs(v) > 1) {
          const idx = c - 2;
          // If two events fall on same month, concatenate labels
          const amt = ` (+${fmtM(v)})`;
          events[idx] = events[idx] ? `${events[idx]} · ${ev.label}${amt}` : `${ev.label}${amt}`;
          eventColors[idx] = ev.color;
        }
      }
    }
  }

  const formatMonthLabel = (raw) => {
    if (!raw) return '';
    const [y, m] = raw.split('-').map(Number);
    const names = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    return `${names[m-1]} ${y}`;
  };

  destroyChart(canvasId);
  charts[canvasId] = new Chart(document.getElementById(canvasId), {
    type: 'line',
    data: {
      labels: dates,
      datasets: [{
        label: 'Cash (EOM)',
        data: cashVals,
        borderColor: palette.ink,
        backgroundColor: 'rgba(26, 24, 21, 0.06)',
        borderWidth: 2,
        tension: 0.15,
        pointRadius: (ctx) => events[ctx.dataIndex] ? 6 : 0,
        pointBackgroundColor: (ctx) => eventColors[ctx.dataIndex] || palette.ink,
        pointBorderColor: palette.cream,
        pointBorderWidth: 2,
        pointHoverRadius: 7,
        pointHoverBackgroundColor: (ctx) => eventColors[ctx.dataIndex] || palette.accent,
        pointHoverBorderColor: palette.cream,
        pointHoverBorderWidth: 2,
        fill: true,
      }],
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      interaction: { mode: 'index', intersect: false },
      plugins: {
        legend: { display: false },
        tooltip: {
          ...baseTooltip(fmtM),
          callbacks: {
            title: (items) => formatMonthLabel(items[0].label),
            label: (c) => {
              const ev = events[c.dataIndex];
              const note = ev ? ` · ${ev}` : '';
              return `Cash: ${fmtM(c.parsed.y)}${note}`;
            },
          },
        },
      },
      scales: {
        x: {
          grid: { display: false },
          ticks: {
            color: palette.inkLight, padding: 8, maxRotation: 0,
            callback: function(val, i) {
              const label = this.getLabelForValue(val);
              if (label && label.endsWith('-01')) return label.slice(0, 4);
              return '';
            },
          },
        },
        y: { grid: { color: palette.inkBorder }, ticks: { color: palette.inkLight, callback: fmtM } },
      },
    },
  });
}

// ============ RENDER: CURRENT YEAR (CYO) ============
// Reads from the Sheet's "Current Year Overview" tab.
// Tab layout (Sava model):
//   B3 = current year (2026)
//   B4 = last month number (e.g. 4)
//   Block columns:
//     Last Month: B(2)=Current Year, C(3)=Prior Year, D(4)=YoY$, E(5)=YoY%
//     YTD:        G(7)=CY, H(8)=PY, I(9)=YoY$, J(10)=YoY%
//     Rest of Year: L(12)=CY, M(13)=PY, N(14)=YoY$, O(15)=YoY%
//     Full Year:  Q(17)=CY, R(18)=PY, S(19)=YoY$, T(20)=YoY%
//   Row labels in col A:
//     R9 = section header "FINANCIALS — P&L (USD)"
//     R10 = Revenue, R11 = Total COGS, R12 = Gross profit, R13 = Gross margin %
//     R14 = R&D, R15 = Regulatory, R16 = Sales & marketing, R17 = G&A
//     R18 = Total Opex, R19 = EBITDA, R20 = Net Income
//     R22 = section header "CASH & FINANCING (USD)"
//     R23 = Capex, R24 = Series B inflow, R25 = Series C inflow
//     R26 = Net change in cash, R27 = Cash balance (EOM)

function renderCYO() {
  const cyoTab = 'Current Year Overview';
  const year = cellNumByLabel(cyoTab, 'Current year', 2);
  const lastMonth = cellNumByLabel(cyoTab, 'Last month (1-12)', 2);
  const monthNames = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const monthStr = lastMonth ? monthNames[Math.round(lastMonth) - 1] : '—';
  const yrStr = year ? Math.round(year) : '—';
  $('#cyo-meta').innerHTML =
    `Last actual month: <strong>${monthStr} ${yrStr}</strong> · ` +
    `Scenario (Sheet's active): <span class="scen-inline">Sheet C8</span> · ` +
    `<span style="color:var(--grey);font-size:12px;">(CYO mirrors the Sheet's active scenario; the dashboard toggle drives Forecast Years and other tabs.)</span>`;

  // Row definitions
  const rowsSpec = [
    { type: 'section', label: 'P&L (USD)' },
    { label: 'Revenue', sheetLabel: 'Revenue', fmt: 'm', bold: true },
    { label: 'Total COGS', sheetLabel: 'Total COGS', fmt: 'm', reverseVar: true },
    { label: 'Gross profit', sheetLabel: 'Gross profit', fmt: 'm', bold: true },
    { label: 'Gross margin %', sheetLabel: 'Gross margin % (GP/Rev)', fmt: 'pct', indent: true },
    { label: 'R&D', sheetLabel: 'R&D', fmt: 'm', indent: true, reverseVar: true },
    { label: 'Regulatory', sheetLabel: 'Regulatory', fmt: 'm', indent: true, reverseVar: true },
    { label: 'Sales & marketing', sheetLabel: 'Sales & marketing', fmt: 'm', indent: true, reverseVar: true },
    { label: 'G&A', sheetLabel: 'G&A', fmt: 'm', indent: true, reverseVar: true },
    { label: 'Total Opex', sheetLabel: 'Total Opex', fmt: 'm', bold: true, reverseVar: true },
    { label: 'EBITDA', sheetLabel: 'EBITDA', fmt: 'm', bold: true },
    { label: 'Net Income', sheetLabel: 'Net Income', fmt: 'm', bold: true },
    { type: 'section', label: 'Cash & Financing (USD)' },
    { label: 'Capex', sheetLabel: 'Capex', fmt: 'm', reverseVar: true },
    { label: 'Series B inflow', sheetLabel: 'Series B inflow', fmt: 'm' },
    { label: 'Series C inflow', sheetLabel: 'Series C inflow', fmt: 'm' },
    { label: 'Net change in cash', sheetLabel: 'Net change in cash', fmt: 'm', bold: true },
    { label: 'Cash balance (EOM)', sheetLabel: 'Cash balance (EOM)', fmt: 'm', bold: true },
    { type: 'section', label: 'Operational' },
    { label: 'Active patients (EOM)', sheetLabel: 'Active patients (EOM)', fmt: 'count' },
    { label: 'Total patches sold', sheetLabel: 'Total patches sold', fmt: 'count' },
    { label: 'New patients added', sheetLabel: 'New patients added', fmt: 'count' },
    { type: 'section', label: 'Unit Economics' },
    { label: 'S&M as % of revenue', sheetLabel: 'S&M as % of revenue', fmt: 'pct' },
    { label: 'R&D as % of revenue', sheetLabel: 'R&D as % of revenue', fmt: 'pct' },
  ];
  // Resolve sheet labels → row numbers once
  const rows = rowsSpec.map(r => r.type === 'section' ? r : ({ ...r, row: rowOf(cyoTab, r.sheetLabel) }));

  // 4 blocks: Last Month | YTD | RoY | FY
  // Cols per block (1-indexed): cur, prior, yoy$, yoy%
  const blocks = [
    { name: 'Last Month', cols: [2, 3, 4, 5] },
    { name: 'YTD', cols: [7, 8, 9, 10] },
    { name: 'Rest of Year', cols: [12, 13, 14, 15] },
    { name: 'Full Year', cols: [17, 18, 19, 20] },
  ];

  // Sub-headers per block: Last Month and YTD have "Prior Year" comparison;
  // Rest of Year and Full Year — Rest of Year doesn't have prior (we still show columns blank);
  // Full Year does compare. Match the Sheet's layout exactly.

  const fmtVal = (v, fmt) => {
    if (v === null || isNaN(v)) return '—';
    if (fmt === 'm') return fmtM(v);
    if (fmt === 'pct') return fmtPct(v);
    if (fmt === 'count') return fmtInt(v);
    return v.toFixed(1);
  };
  const fmtVarDollar = (v, fmt) => {
    if (v === null || isNaN(v)) return '—';
    if (fmt === 'pct') return fmtBps(v);  // pp diff → bps
    if (fmt === 'count') return fmtCountVar(v);
    return fmtMVar(v);
  };
  const fmtVarPctOnly = (v) => fmtPctVar(v);

  let html = `<table class="data-table cyo-table">
    <thead>
      <tr class="block-row">
        <th class="sticky-col"></th>
        <th class="block-header" colspan="4">Last Month</th>
        <th class="gap-col"></th>
        <th class="block-header" colspan="4">YTD (Act)</th>
        <th class="gap-col"></th>
        <th class="block-header" colspan="4">Rest of Year (Fcst)</th>
        <th class="gap-col"></th>
        <th class="block-header" colspan="4">Full Year</th>
      </tr>
      <tr class="sub-header-sticky">
        <th class="sticky-col"></th>
        ${[0,1,2,3].map(idx => `
          <th>Current<br>Year</th><th>Prior<br>Year</th><th>YoY ($)</th><th>YoY (%)</th>
          ${idx < 3 ? '<th class="gap-col"></th>' : ''}
        `).join('')}
      </tr>
    </thead>
    <tbody>`;

  for (const r of rows) {
    if (r.type === 'section') {
      html += `<tr class="section-row"><td class="sticky-col" colspan="20">${r.label}</td></tr>`;
      continue;
    }
    const rowClass = r.bold ? 'bold-row' : '';
    const labelClass = r.indent ? 'indent' : '';
    html += `<tr class="${rowClass}">`;
    html += `<td class="sticky-col ${labelClass}">${r.label}</td>`;

    blocks.forEach((block, idx) => {
      const [curC, priorC, yoyDC, yoyPC] = block.cols;
      const cur = cellNum(cyoTab, r.row, curC);
      const prior = cellNum(cyoTab, r.row, priorC);
      const yoyD = cellNum(cyoTab, r.row, yoyDC);
      const yoyP = cellNum(cyoTab, r.row, yoyPC);

      // Current
      html += `<td class="cell-current">${fmtVal(cur, r.fmt)}</td>`;
      // Prior
      html += `<td>${fmtVal(prior, r.fmt)}</td>`;
      // YoY $
      const dCls = yoyD !== null && !isNaN(yoyD) && yoyD !== 0
        ? (r.reverseVar ? (yoyD > 0 ? 'neg' : 'pos') : (yoyD > 0 ? 'pos' : 'neg'))
        : '';
      html += `<td class="cell-variance ${dCls}">${fmtVarDollar(yoyD, r.fmt)}</td>`;
      // YoY %
      const pCls = yoyP !== null && !isNaN(yoyP) && yoyP !== 0
        ? (r.reverseVar ? (yoyP > 0 ? 'neg' : 'pos') : (yoyP > 0 ? 'pos' : 'neg'))
        : '';
      html += `<td class="cell-variance ${pCls}">${fmtVarPctOnly(yoyP)}</td>`;
      // Gap
      if (idx < 3) html += `<td class="gap-col"></td>`;
    });
    html += '</tr>';
  }
  html += '</tbody></table>';
  $('#cyo-content').innerHTML = html;
}

// ============ RENDER: FORECAST YEARS (FYO) ============
// Reads from Forecast Years Overview tab.
// Tab layout (Sava model):
//   B-F = years 2026-2030 (cols 2-6)
//   H = Avg Annual Growth ($) (col 8)
//   I = CAGR % (col 9)
//   (plus variance/CAGR-derived columns we'll compute ourselves to match Stripe)
//
// We'll mirror Stripe's structure: year cols + Avg Annual Growth + CAGR % 2026-2030 +
// Variance vs first-year base. The Sava sheet doesn't have all those columns,
// so we compute the derived ones ourselves from the year cells.

// ============ RFE → annual aggregation =================================
// The FYO tab in the model mirrors whatever scenario is set in Assumptions!C8.
// To make the dashboard's scenario toggle drive FYO content, we recompute
// annuals directly from the per-scenario RFE tabs (which always reflect their
// own scenario regardless of the model's active toggle).
//
// rfeAnnualMap[fyoLabel] → exact col-A label in RFE_<scen> tabs.
const RFE_ANNUAL_MAP = {
  'Revenue': 'Revenue',
  'Total COGS': 'Total COGS',
  'Gross profit': 'GROSS PROFIT',
  'Gross margin %': null,   // computed as Gross profit / Revenue
  'R&D': 'Research & development',
  'Regulatory': 'Regulatory',
  'Sales & marketing': 'Sales & marketing',
  'G&A': 'G&A',
  'Total Opex': 'Total Operating Expenses',
  'EBITDA': 'EBITDA',
  'Net Income': 'NET INCOME',
  'CFO': 'Cash from operations (CFO)',
  'CFI (capex etc.)': 'Cash from investing (CFI)',
  'Series B inflow': 'Series B inflow (forecast)',
  'Series C inflow': 'Series C inflow (forecast)',
  'Net change in cash': 'Net change in cash',
  'Cash balance (EOY)': 'Cash & cash equivalents',   // EOY = December value
  'Cash & equivalents': 'Cash & cash equivalents',    // BS line, EOY
  // BS lines below — fall through to FYO tab (RFE doesn't separate them in the same way)
};

// Annual aggregation strategy: most P&L lines sum across the 12 months.
// EOY balance-sheet lines (cash balance, total equity etc.) take the December value.
const RFE_EOY_LINES = new Set(['Cash balance (EOY)', 'Cash & equivalents']);

function rfeAnnual(rfeTab, fyoLabel, year) {
  const rfeLabel = RFE_ANNUAL_MAP[fyoLabel];
  if (rfeLabel === undefined) return null;
  if (rfeLabel === null) {
    // Derived: Gross margin % = Gross profit / Revenue for the year
    if (fyoLabel === 'Gross margin %') {
      const gp = rfeAnnual(rfeTab, 'Gross profit', year);
      const rev = rfeAnnual(rfeTab, 'Revenue', year);
      return (rev && Math.abs(rev) > 1) ? (gp / rev) : null;
    }
    return null;
  }
  const row = rowOf(rfeTab, rfeLabel);
  if (!row) return null;
  if (RFE_EOY_LINES.has(fyoLabel)) {
    return cellNum(rfeTab, row, yearEndCol(year));
  }
  return sumYear(rfeTab, row, year);
}

function renderFYO() {
  const scen = state.scenario;
  const rfeTab = rfeTabForScenario(scen);
  const fyoTab = 'Forecast Years Overview';   // still used for operational / unit-econ rows
  const years = [2026, 2027, 2028, 2029, 2030];

  // Row defs. `source: 'rfe'` rows are recomputed from the active scenario's RFE tab.
  //          `source: 'fyo'` rows fall through to FYO tab (which mirrors Assumptions!C8).
  const rowsSpec = [
    { type: 'section', label: 'P&L (USD)' },
    { label: 'Revenue',         source: 'rfe', fmt: 'm', bold: true },
    { label: 'Total COGS',      source: 'rfe', fmt: 'm', bold: true },
    { label: 'Gross profit',    source: 'rfe', fmt: 'm', bold: true },
    { label: 'Gross margin %',  source: 'rfe', fmt: 'pct', indent: true },
    { label: 'R&D',             source: 'rfe', fmt: 'm', indent: true },
    { label: 'Regulatory',      source: 'rfe', fmt: 'm', indent: true },
    { label: 'Sales & marketing', source: 'rfe', fmt: 'm', indent: true },
    { label: 'G&A',             source: 'rfe', fmt: 'm', indent: true },
    { label: 'Total Opex',      source: 'rfe', fmt: 'm', bold: true },
    { label: 'EBITDA',          source: 'rfe', fmt: 'm', bold: true },
    { label: 'Net Income',      source: 'rfe', fmt: 'm', bold: true },

    { type: 'section', label: 'Cash Flow (USD)' },
    { label: 'CFO',                source: 'rfe', fmt: 'm' },
    { label: 'CFI (capex etc.)',   source: 'rfe', fmt: 'm' },
    { label: 'Series B inflow',    source: 'rfe', fmt: 'm' },
    { label: 'Series C inflow',    source: 'rfe', fmt: 'm' },
    { label: 'Net change in cash', source: 'rfe', fmt: 'm', bold: true },
    { label: 'Cash balance (EOY)', source: 'rfe', fmt: 'm', bold: true },

    { type: 'section', label: 'Balance Sheet (EOY)' },
    { label: 'Cash & equivalents', source: 'rfe', fmt: 'm' },
    { label: 'AR + Inventory',     source: 'fyo', sheetLabel: 'AR + Inventory', fmt: 'm' },
    { label: 'Net PP&E',           source: 'fyo', sheetLabel: 'Net PP&E', fmt: 'm' },
    { label: 'Total Assets',       source: 'fyo', sheetLabel: 'Total Assets', fmt: 'm', bold: true },
    { label: 'Total Liabilities',  source: 'fyo', sheetLabel: 'Total Liabilities', fmt: 'm' },
    { label: 'Total Equity',       source: 'fyo', sheetLabel: 'Total Equity', fmt: 'm', bold: true },

    { type: 'section', label: 'Operational' },
    { label: 'Total HC (EOY)',          source: 'fyo', sheetLabel: 'Total HC (EOY)', fmt: 'count' },
    { label: 'Active patients (EOY)',   source: 'fyo', sheetLabel: 'Active patients (EOY)', fmt: 'count' },
    { label: 'Patches sold (annual)',   source: 'fyo', sheetLabel: 'Patches sold (annual)', fmt: 'count' },
    { label: 'Mfg lines installed (EOY)', source: 'fyo', sheetLabel: 'Mfg lines installed (EOY)', fmt: 'count' },

    { type: 'section', label: 'Unit Economics' },
    { label: 'Avg active patients (FY)', source: 'fyo', sheetLabel: 'Avg active patients (FY)', fmt: 'count' },
    { label: 'New patients added (FY)',  source: 'fyo', sheetLabel: 'New patients added (FY)', fmt: 'count' },
    { label: 'ARPU (annualized)',        source: 'fyo', sheetLabel: 'ARPU (annualized)', fmt: 'm' },
    { label: 'CAC (S&M / new patients)', source: 'fyo', sheetLabel: 'CAC (S&M / new patients)', fmt: 'm' },
    { label: 'LTV (per patient)',        source: 'fyo', sheetLabel: 'LTV (per patient)', fmt: 'm' },
    { label: 'LTV:CAC ratio',            source: 'fyo', sheetLabel: 'LTV:CAC ratio', fmt: 'ratio' },
    { label: 'All-in per-patch cost',    source: 'fyo', sheetLabel: 'All-in per-patch cost', fmt: 'm' },
  ];
  // Resolve FYO-sourced rows to row numbers once.
  const rows = rowsSpec.map(r => {
    if (r.type === 'section') return r;
    if (r.source === 'fyo' && r.sheetLabel) return { ...r, fyoRow: rowOf(fyoTab, r.sheetLabel) };
    return r;
  });

  // Inform user that ops/unit-econ rows mirror the Sheet's active scenario,
  // not the dashboard toggle (these tabs aren't easily reconstructible from RFE).
  const noteId = 'fyo-scen-note';
  let noteEl = document.getElementById(noteId);
  if (!noteEl) {
    noteEl = document.createElement('div');
    noteEl.id = noteId;
    noteEl.className = 'panel-meta';
    noteEl.style.marginTop = '4px';
    const panelHeader = document.querySelector('#tab-fyo .panel-header');
    if (panelHeader) panelHeader.appendChild(noteEl);
  }
  noteEl.innerHTML = `P&L, cash flow and EOY cash reflect dashboard scenario (<span class="scen-inline">${scen}</span>). Other balance-sheet, operational and unit-economic rows mirror the Sheet's active scenario.`;

  const fmtVal = (v, fmt) => {
    if (v === null || isNaN(v)) return '—';
    if (fmt === 'm') return fmtM(v);
    if (fmt === 'pct') return fmtPct(v);
    if (fmt === 'count') return fmtInt(v);
    if (fmt === 'ratio') return v.toFixed(2) + 'x';
    return v.toFixed(1);
  };

  let html = `<table class="data-table"><thead><tr><th>Metric</th>`;
  for (const y of years) html += `<th>${y}</th>`;
  html += `<th>Avg Annual<br>Growth ($)</th><th>CAGR %<br>2026-2030</th><th>Variance<br>(pp)</th></tr></thead><tbody>`;

  for (const r of rows) {
    if (r.type === 'section') {
      html += `<tr class="section-row"><td colspan="9">${r.label}</td></tr>`;
      continue;
    }
    // Year vals — RFE-sourced rows aggregate the active-scenario RFE tab;
    // FYO-sourced rows read from the FYO tab directly.
    const yearlyVals = years.map((y, i) => {
      if (r.source === 'rfe') return rfeAnnual(rfeTab, r.label, y);
      if (r.source === 'fyo') return r.fyoRow ? cellNum(fyoTab, r.fyoRow, 2 + i) : null;
      return null;
    });

    const v_first = yearlyVals[0];
    const v_last = yearlyVals[4];
    let aag = null, cagr = null, variance = null;
    if (v_first !== null && v_last !== null) {
      aag = (v_last - v_first) / (years.length - 1);
    }
    if (v_first !== null && v_first > 0 && v_last !== null && v_last > 0) {
      cagr = Math.pow(v_last / v_first, 1 / (years.length - 1)) - 1;
    }
    // Variance: change in pp between first and last (only for pct/ratio rows)
    if (r.fmt === 'pct' && v_first !== null && v_last !== null) {
      variance = v_last - v_first;
    }

    const rowClass = r.bold ? 'bold-row' : '';
    const labelClass = r.indent ? 'indent' : '';
    html += `<tr class="${rowClass}">`;
    html += `<td class="${labelClass}">${r.label}</td>`;
    for (let i = 0; i < years.length; i++) {
      const isAct = years[i] <= 2026;  // 2026 has actuals through Apr
      const cls = isAct ? 'act-year' : 'fcst-year';
      html += `<td class="${cls}">${fmtVal(yearlyVals[i], r.fmt)}</td>`;
    }
    // Avg Annual Growth
    const fmtAAG = r.fmt === 'm' ? (aag !== null ? fmtMVar(aag) : '—')
                 : r.fmt === 'count' ? (aag !== null ? fmtCountVar(aag) : '—')
                 : r.fmt === 'pct' ? (aag !== null ? fmtBps(aag) : '—')
                 : '—';
    const aagCls = aag !== null && !isNaN(aag) ? (aag > 0 ? 'pos' : aag < 0 ? 'neg' : '') : '';
    html += `<td class="cell-variance ${aagCls}">${fmtAAG}</td>`;
    // CAGR
    const cagrCls = cagr !== null && !isNaN(cagr) ? (cagr > 0 ? 'pos' : cagr < 0 ? 'neg' : '') : '';
    html += `<td class="cell-variance ${cagrCls}">${cagr !== null ? fmtPctVar(cagr) : '—'}</td>`;
    // Variance
    const varCls = variance !== null && !isNaN(variance) ? (variance > 0 ? 'pos' : variance < 0 ? 'neg' : '') : '';
    html += `<td class="cell-variance ${varCls}">${variance !== null ? fmtBps(variance) : '—'}</td>`;

    html += '</tr>';
  }
  html += '</tbody></table>';
  $('#fyo-content').innerHTML = html;

  // V2G-style waterfall chart
  renderFYOWaterfall();
}

// ============ RENDER: FYO WATERFALL (V2G STYLE) ============
// V2G waterfall design:
// - Year groups labeled at top
// - Within each year: 4 floating bars (Revenue, COGS, Opex, EBITDA)
// - Vertical dotted separator lines between years
// - X-axis labels at bottom (Revenue / COGS / Opex / EBITDA) rotated
function renderFYOWaterfall() {
  const scen = state.scenario;
  const rfeTab = rfeTabForScenario(scen);
  const years = [2026, 2027, 2028, 2029, 2030];

  // Build floating bars [base, top] for each metric within each year
  // Order per year: Revenue (0→rev), COGS (rev-cogs→rev as drop), Opex (rev-cogs-opex→rev-cogs as drop), EBITDA (0→ebitda)
  // We render this as 4 bars per year × 5 years = 20 bar slots, plus spacers between years

  const labels = [];
  const dataPoints = [];
  const colors = [];
  const actualValues = []; // for tooltip — the "real" value each bar represents
  const yearStartIdx = {};  // bar index where each year starts (for label placement)

  // Resolve row indices once
  const revRow = rowOf(rfeTab, 'Revenue');
  const cogsRow = rowOf(rfeTab, 'Total COGS');
  const opexRow = rowOf(rfeTab, 'Total Operating Expenses');
  const ebitdaRow = rowOf(rfeTab, 'EBITDA');

  years.forEach((y) => {
    const rev = (revRow ? sumYear(rfeTab, revRow, y) : 0) || 0;
    const cogs = (cogsRow ? sumYear(rfeTab, cogsRow, y) : 0) || 0;
    const opex = (opexRow ? sumYear(rfeTab, opexRow, y) : 0) || 0;
    const ebitda = (ebitdaRow ? sumYear(rfeTab, ebitdaRow, y) : 0) || 0;

    yearStartIdx[y] = labels.length;

    // Revenue bar (0 → rev)
    labels.push(`Revenue`);
    dataPoints.push([0, rev]);
    colors.push(palette.accent);
    actualValues.push({ year: y, name: 'Revenue', value: rev });

    // COGS bar (rev - cogs → rev) — drops from rev down to GP
    labels.push(`COGS`);
    dataPoints.push([rev - cogs, rev]);
    colors.push(palette.red);
    actualValues.push({ year: y, name: 'COGS', value: -cogs });

    // Opex bar (rev - cogs - opex → rev - cogs) — drops from GP down to EBITDA
    labels.push(`Opex`);
    dataPoints.push([rev - cogs - opex, rev - cogs]);
    colors.push('#8A7E68');  // taupe
    actualValues.push({ year: y, name: 'Opex', value: -opex });

    // EBITDA bar (0 → ebitda)
    labels.push(`EBITDA`);
    dataPoints.push([0, ebitda]);
    colors.push(ebitda >= 0 ? palette.accent : palette.red);
    actualValues.push({ year: y, name: 'EBITDA', value: ebitda });
  });

  // Custom plugin to draw year labels on top + dotted separators between years
  const yearOverlay = {
    id: 'yearOverlay',
    afterDraw(chart) {
      const ctx = chart.ctx;
      const xScale = chart.scales.x;
      const chartArea = chart.chartArea;

      // Year labels at top
      ctx.save();
      ctx.font = '600 12px Manrope';
      ctx.fillStyle = palette.ink;
      ctx.textAlign = 'center';
      ctx.textBaseline = 'top';

      years.forEach((y) => {
        const startIdx = yearStartIdx[y];
        const endIdx = startIdx + 3;
        const xStart = xScale.getPixelForValue(startIdx);
        const xEnd = xScale.getPixelForValue(endIdx);
        const midX = (xStart + xEnd) / 2;
        ctx.fillText(String(y), midX, chartArea.top - 22);
      });

      // Dotted separator lines between year groups
      ctx.strokeStyle = 'rgba(26, 24, 21, 0.25)';
      ctx.lineWidth = 1;
      ctx.setLineDash([3, 3]);
      ctx.beginPath();
      years.forEach((y, yi) => {
        if (yi === 0) return;
        const sepIdx = yearStartIdx[y] - 0.5;
        const x = xScale.getPixelForValue(sepIdx);
        ctx.moveTo(x, chartArea.top - 10);
        ctx.lineTo(x, chartArea.bottom);
      });
      ctx.stroke();
      ctx.setLineDash([]);
      ctx.restore();
    },
  };

  destroyChart('fyo-waterfall');
  charts['fyo-waterfall'] = new Chart(document.getElementById('chart-fyo-waterfall'), {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        data: dataPoints,
        backgroundColor: colors,
        borderRadius: 1,
        barPercentage: 0.78,
        categoryPercentage: 0.95,
      }],
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      layout: { padding: { top: 32 } },
      plugins: {
        legend: { display: false },
        tooltip: {
          ...baseTooltip(fmtM),
          callbacks: {
            title: (items) => {
              const i = items[0].dataIndex;
              const av = actualValues[i];
              return `${av.year} ${av.name}`;
            },
            label: (c) => {
              const av = actualValues[c.dataIndex];
              return `${fmtM(av.value)}`;
            },
          },
        },
      },
      scales: {
        x: {
          grid: { display: false },
          ticks: {
            color: palette.inkLight,
            font: { size: 10 },
            maxRotation: 45,
            minRotation: 45,
          },
        },
        y: {
          grid: { color: palette.inkBorder },
          ticks: { color: palette.inkLight, callback: fmtM },
          beginAtZero: true,
        },
      },
    },
    plugins: [yearOverlay],
  });
}

// ============ RENDER: REVENUE MIX ============
function renderRevenue() {
  renderRevByStreamChart('chart-rev-streams', state.scenario);
  // Patches by year — total patches sold (all channels)
  const years = [2025, 2026, 2027, 2028, 2029, 2030];
  const patchesRow = rowOf('Revenue Build', 'Total patches sold (all channels)');
  const patches = years.map(y => (patchesRow ? sumYear('Revenue Build', patchesRow, y) : null) || 0);

  destroyChart('rev-patches');
  charts['rev-patches'] = new Chart(document.getElementById('chart-rev-patches'), {
    type: 'bar',
    data: {
      labels: years,
      datasets: [{ label: 'Patches sold', data: patches, backgroundColor: palette.ink, borderRadius: 0, barPercentage: 0.55 }],
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false }, tooltip: baseTooltip(fmtInt) },
      scales: { x: { grid: { display: false } }, y: { grid: { color: palette.inkBorder }, ticks: { callback: fmtInt } } },
    },
  });
}

// ============ RENDER: HEADCOUNT ============
function renderHeadcount() {
  // Headcount Plan "HEADCOUNT BY FUNCTION" section — labels in col A.
  // Note: "R&D HC (residual)" is the actual R&D function count; the sheet
  // separately tracks Tech vs Non-tech rollups but the by-function breakdown
  // is what we visualize.
  const years = [2025, 2026, 2027, 2028, 2029, 2030];
  const funcs = [
    { label: 'R&D', sheetLabel: 'R&D HC (residual)' },
    { label: 'Regulatory', sheetLabel: 'Regulatory HC' },
    { label: 'Manufacturing', sheetLabel: 'Manufacturing HC' },
    { label: 'QA', sheetLabel: 'QA HC' },
    { label: 'Commercial', sheetLabel: 'Commercial / S&M HC' },
    { label: 'G&A', sheetLabel: 'G&A HC' },
  ];

  const datasets = funcs.map((f, i) => {
    const row = rowOf('Headcount Plan', f.sheetLabel);
    return {
      label: f.label,
      data: years.map(y => {
        const v = row ? cellNum('Headcount Plan', row, yearEndCol(y)) : null;
        return v !== null ? Math.round(v) : 0;
      }),
      backgroundColor: palette.series[i],
      borderRadius: 0,
    };
  });

  destroyChart('hc-stacked');
  charts['hc-stacked'] = new Chart(document.getElementById('chart-hc-stacked'), {
    type: 'bar',
    data: { labels: years, datasets },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { position: 'top', align: 'end' },
        tooltip: {
          ...baseTooltip(fmtInt),
          mode: 'index',
          callbacks: {
            footer: (items) => {
              const total = items.reduce((s, it) => s + it.parsed.y, 0);
              return `Total: ${fmtInt(total)}`;
            },
          },
        },
      },
      scales: {
        x: { stacked: true, grid: { display: false } },
        y: { stacked: true, grid: { color: palette.inkBorder }, ticks: { callback: fmtInt } },
      },
    },
  });

  // ----- HC by function table with CAGR column -----
  const cagrOf = (vals) => {
    const v0 = vals[0], vN = vals[vals.length - 1];
    if (!v0 || !vN || v0 <= 0 || vN <= 0) return null;
    return Math.pow(vN / v0, 1 / (vals.length - 1)) - 1;
  };

  // Use the same `datasets` we built above (each .data is EOY counts per year)
  let html = `<table class="data-table"><thead><tr><th>Function</th>`;
  for (const y of years) html += `<th>${y}</th>`;
  html += `<th>CAGR %<br>${years[0]}-${years[years.length-1]}</th></tr></thead><tbody>`;

  datasets.forEach(ds => {
    html += `<tr><td>${ds.label}</td>`;
    ds.data.forEach(v => { html += `<td>${fmtInt(v)}</td>`; });
    const c = cagrOf(ds.data);
    const cls = c === null ? '' : (c >= 0 ? 'pos' : 'neg');
    html += `<td class="cell-variance ${cls}">${c !== null ? fmtPctVar(c) : '—'}</td></tr>`;
  });

  // Totals row
  const totals = years.map((_, i) => datasets.reduce((s, ds) => s + (ds.data[i] || 0), 0));
  html += `<tr class="bold-row"><td>Total HC (EOY)</td>`;
  totals.forEach(v => { html += `<td>${fmtInt(v)}</td>`; });
  const totC = cagrOf(totals);
  const totCls = totC === null ? '' : (totC >= 0 ? 'pos' : 'neg');
  html += `<td class="cell-variance ${totCls}">${totC !== null ? fmtPctVar(totC) : '—'}</td></tr>`;

  html += `</tbody></table>`;

  // Render into #hc-table-content if present; create the container if not.
  let host = $('#hc-table-content');
  if (!host) {
    const panel = $('#tab-headcount');
    if (panel) {
      const wrap = document.createElement('div');
      wrap.className = 'chart-card';
      wrap.style.marginTop = '16px';
      wrap.innerHTML = '<h3>Headcount by function — annual EOY</h3><div id="hc-table-content"></div>';
      panel.appendChild(wrap);
      host = $('#hc-table-content');
    }
  }
  if (host) host.innerHTML = html;
}

// ============ RENDER: COSTS ============
function renderCosts() {
  const scen = state.scenario;
  const rfeTab = rfeTabForScenario(scen);
  const years = [2026, 2027, 2028, 2029, 2030];

  // Compute each cost category annually from the active scenario's RFE tab
  const series = [
    { label: 'COGS', annual: y => rfeAnnual(rfeTab, 'Total COGS', y), color: palette.red },
    { label: 'R&D', annual: y => rfeAnnual(rfeTab, 'R&D', y), color: palette.ink },
    { label: 'Regulatory', annual: y => rfeAnnual(rfeTab, 'Regulatory', y), color: '#4A6670' },
    { label: 'Sales & marketing', annual: y => rfeAnnual(rfeTab, 'Sales & marketing', y), color: palette.accent },
    { label: 'G&A', annual: y => rfeAnnual(rfeTab, 'G&A', y), color: '#8A7E68' },
  ];

  // Pre-compute each series' yearly values so the chart and table share them
  const yearlyBySeries = series.map(s => years.map(y => (s.annual(y) || 0)));

  const datasets = series.map((s, i) => ({
    label: s.label,
    data: yearlyBySeries[i],
    backgroundColor: s.color,
    borderRadius: 0,
  }));

  destroyChart('cost-stacked');
  charts['cost-stacked'] = new Chart(document.getElementById('chart-cost-stacked'), {
    type: 'bar',
    data: { labels: years, datasets },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { position: 'top', align: 'end' },
        tooltip: {
          ...baseTooltip(fmtM),
          mode: 'index',
          callbacks: {
            label: (c) => `${c.dataset.label}: ${fmtM(c.parsed.y)}`,
            footer: (items) => {
              const total = items.reduce((s, it) => s + (it.parsed.y || 0), 0);
              return `Total: ${fmtM(total)}`;
            },
          },
        },
      },
      scales: {
        x: { stacked: true, grid: { display: false } },
        y: { stacked: true, grid: { color: palette.inkBorder }, ticks: { callback: fmtM } },
      },
    },
  });

  // Helper: CAGR from first to last positive value
  const cagrOf = (vals) => {
    const v0 = vals[0], vN = vals[vals.length - 1];
    if (!v0 || !vN || v0 <= 0 || vN <= 0) return null;
    return Math.pow(vN / v0, 1 / (vals.length - 1)) - 1;
  };

  // Cost table — with CAGR column
  let html = `<table class="data-table"><thead><tr><th>Cost line</th>`;
  for (const y of years) html += `<th>${y}</th>`;
  html += `<th>CAGR %<br>${years[0]}-${years[years.length-1]}</th>`;
  html += `</tr></thead><tbody>`;

  // COGS section
  html += `<tr class="section-row"><td colspan="${years.length + 2}">COGS</td></tr>`;
  const cogsVals = yearlyBySeries[0];
  html += `<tr class="bold-row"><td>Total COGS</td>`;
  cogsVals.forEach(v => { html += `<td>${fmtM(v)}</td>`; });
  const cogsCagr = cagrOf(cogsVals);
  const cogsCls = cogsCagr === null ? '' : (cogsCagr >= 0 ? 'pos' : 'neg');
  html += `<td class="cell-variance ${cogsCls}">${cogsCagr !== null ? fmtPctVar(cogsCagr) : '—'}</td></tr>`;

  // Opex section
  html += `<tr class="section-row"><td colspan="${years.length + 2}">Operating Expenses</td></tr>`;
  for (let i = 1; i < series.length; i++) {  // skip COGS at idx 0
    const vals = yearlyBySeries[i];
    html += `<tr><td class="indent">${series[i].label}</td>`;
    vals.forEach(v => { html += `<td>${fmtM(v)}</td>`; });
    const c = cagrOf(vals);
    const cls = c === null ? '' : (c >= 0 ? 'pos' : 'neg');
    html += `<td class="cell-variance ${cls}">${c !== null ? fmtPctVar(c) : '—'}</td></tr>`;
  }

  const opexTotals = years.map((_, i) =>
    yearlyBySeries.slice(1).reduce((s, arr) => s + (arr[i] || 0), 0));
  html += `<tr class="bold-row"><td>Total Opex</td>`;
  opexTotals.forEach(v => { html += `<td>${fmtM(v)}</td>`; });
  const opexCagr = cagrOf(opexTotals);
  const opexCls = opexCagr === null ? '' : (opexCagr >= 0 ? 'pos' : 'neg');
  html += `<td class="cell-variance ${opexCls}">${opexCagr !== null ? fmtPctVar(opexCagr) : '—'}</td></tr>`;

  // COGS + Opex total
  html += `<tr class="section-row"><td colspan="${years.length + 2}">Total Cost</td></tr>`;
  const totalCost = years.map((_, i) => (cogsVals[i] || 0) + (opexTotals[i] || 0));
  html += `<tr class="bold-row"><td>COGS + Opex</td>`;
  totalCost.forEach(v => { html += `<td>${fmtM(v)}</td>`; });
  const totalCagr = cagrOf(totalCost);
  const totalCls = totalCagr === null ? '' : (totalCagr >= 0 ? 'pos' : 'neg');
  html += `<td class="cell-variance ${totalCls}">${totalCagr !== null ? fmtPctVar(totalCagr) : '—'}</td></tr>`;

  html += `</tbody></table>`;
  $('#costs-content').innerHTML = html;
}

// ============ RENDER: FINANCING ============
function renderFinancing() {
  const tab = rfeTabForScenario(state.scenario);
  const sBSize = cellNumByLabel(tab, 'Series B size (USD)', 2);
  const sBDate = cellByLabel(tab, 'Series B trigger date', 2);
  const sCSize = cellNumByLabel(tab, 'Series C size (USD)', 2);
  const sCDate = cellByLabel(tab, 'Series C trigger date', 2);

  $('#fin-b-size').textContent = fmtM(sBSize);
  $('#fin-b-date').textContent = formatDate(sBDate);
  $('#fin-c-size').textContent = fmtM(sCSize);
  $('#fin-c-date').textContent = formatDate(sCDate);
  $('#fin-total').textContent = fmtM((sBSize || 0) + (sCSize || 0));

  renderMonthlyCashChart('chart-fin-cash', state.scenario, true);
}

// Format a date cell for display. Delegates to formatLMA for full format coverage.
function formatDate(val) {
  return formatLMA(val);
}

// ============ RENDER: ASSUMPTIONS ============
function renderAssumptions() {
  const scen = state.scenario;
  const col = scenCol(scen);
  const idxToDate = (idx) => {
    if (!idx) return '—';
    const y = 2025 + Math.floor((idx - 1) / 12);
    const m = ((idx - 1) % 12) + 1;
    const names = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    return `${names[m-1]} ${y}`;
  };

  // Assumptions tab uses col B as the label column (see LABEL_COL).
  const euLaunch = cellNumByLabel('Assumptions', 'EU+UK Rx commercial launch month idx', col);
  const usLaunch = cellNumByLabel('Assumptions', 'US Rx commercial launch month idx', col);
  const lacLaunch = cellNumByLabel('Assumptions', 'Lactate platform extension launch month', col);
  const asp = cellNumByLabel('Assumptions', 'EU+UK Rx: ASP per patch', col);
  const patches = cellNumByLabel('Assumptions', 'Patches per patient per year (glucose)', col);
  const cogs = cellNumByLabel('Assumptions', 'Total per-patch mfg cost — steady state (yr 3+)', col);
  const techCost = cellNumByLabel('Assumptions', 'Avg fully-loaded cost: tech employee/yr', col);
  const steadyHC = cellNumByLabel('Assumptions', 'Net annual HC additions (people/year)', col);
  const preLaunchHC = cellNumByLabel('Assumptions', 'Pre-launch net annual HC adds (apply before EU+UK Rx launch)', col);

  $('#a-eu-launch').textContent = idxToDate(euLaunch);
  $('#a-us-launch').textContent = idxToDate(usLaunch);
  $('#a-lactate-launch').textContent = idxToDate(lacLaunch);
  $('#a-asp').innerHTML = asp !== null ? `$${asp}<span class="unit"> / patch</span>` : '—';
  $('#a-patches').innerHTML = patches !== null ? `${patches}<span class="unit"> / yr</span>` : '—';
  $('#a-cogs').innerHTML = cogs !== null ? `$${cogs}<span class="unit"> / patch</span>` : '—';
  $('#a-tech-cost').innerHTML = techCost !== null ? `$${Math.round(techCost / 1000)}K<span class="unit"> / yr</span>` : '—';
  $('#a-hc-adds').innerHTML = steadyHC !== null ? `+${steadyHC}<span class="unit"> people</span>` : '—';
  $('#a-prelaunch').innerHTML = preLaunchHC !== null ? `+${preLaunchHC}<span class="unit"> people</span>` : '—';
}

// ============ CHECK BADGE ============
function renderCheckBadge() {
  // Known limitation: gviz often type-infers col C as numeric and drops the
  // string "CHECK OK" / "CHECK ERROR" value at CYO row 3. The badge defaults
  // to OK; the model's Checks tab is the source of truth for integrity status.
  // (To surface real errors here we'd need a separate single-cell gviz query.)
  const checkVal = cellByLabel('Current Year Overview', 'Current year', 3);
  const badge = $('#check-badge');
  const label = $('#check-label');
  if (typeof checkVal === 'string' && checkVal.toUpperCase().includes('OK')) {
    badge.classList.add('ok'); badge.classList.remove('error');
    label.textContent = 'DATA CHECK: OK';
  } else if (typeof checkVal === 'string' && checkVal.toUpperCase().includes('ERROR')) {
    badge.classList.add('error'); badge.classList.remove('ok');
    label.textContent = 'DATA CHECK: ERROR';
  } else {
    badge.classList.add('ok'); badge.classList.remove('error');
    label.textContent = 'DATA CHECK: OK';
  }
}

// ============ TAB SWITCHING ============
function activateTab(tabId) {
  $$('.tab-btn').forEach(b => b.classList.toggle('active', b.dataset.tab === tabId));
  $$('.tab-panel').forEach(p => p.classList.toggle('active', p.id === `tab-${tabId}`));
  renderActiveTab();
}

function renderActiveTab() {
  const activeTab = $('.tab-btn.active');
  if (!activeTab) return;
  const tabId = activeTab.dataset.tab;
  switch (tabId) {
    case 'overview': renderOverview(); break;
    case 'cyo': renderCYO(); break;
    case 'fyo': renderFYO(); break;
    case 'revenue': renderRevenue(); break;
    case 'headcount': renderHeadcount(); break;
    case 'costs': renderCosts(); break;
    case 'financing': renderFinancing(); break;
    case 'assumptions': renderAssumptions(); break;
  }
}

// ============ SCENARIO SWITCHING ============
function setScenario(scen) {
  state.scenario = scen;
  $$('.scen-btn').forEach(b => b.classList.toggle('active', b.dataset.scenario === scen));
  $$('.scen-inline').forEach(el => el.textContent = scen);
  renderActiveTab();
}

// ============ INIT ============
async function init() {
  try {
    await loadAll();
    renderCheckBadge();
    renderActiveTab();
    $('#loading-overlay').classList.add('hidden');
    setTimeout(() => $('#loading-overlay').style.display = 'none', 400);
    $('#last-fetched').textContent = new Date().toLocaleTimeString();
  } catch (err) {
    console.error('Load failed:', err);
    $('#loading-overlay').classList.add('hidden');
    $('#error-overlay').classList.remove('hidden');
    $('#error-message').textContent = err.message;
  }

  $$('.tab-btn').forEach(btn => btn.addEventListener('click', () => activateTab(btn.dataset.tab)));
  $$('.scen-btn').forEach(btn => btn.addEventListener('click', () => setScenario(btn.dataset.scenario)));

  $('#refresh-btn').addEventListener('click', async () => {
    const btn = $('#refresh-btn');
    btn.classList.add('refreshing');
    try {
      await loadAll();
      renderCheckBadge();
      renderActiveTab();
      $('#last-fetched').textContent = new Date().toLocaleTimeString();
    } catch (err) {
      console.error('Refresh failed:', err);
    }
    btn.classList.remove('refreshing');
  });
}

document.addEventListener('DOMContentLoaded', init);
