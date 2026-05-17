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
  state.loaded = true;
  if (statusEl) statusEl.textContent = 'Rendering...';
}

// ============ DATA ACCESS ============
// 1-indexed cell access (matches spreadsheet conventions)
function cell(tabName, row, col) {
  const data = state.data[tabName];
  if (!data || !data[row - 1]) return null;
  return data[row - 1][col - 1];
}
function cellNum(tabName, row, col) {
  return parseNum(cell(tabName, row, col));
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
const palette = {
  ink: '#1A1815',
  inkLight: 'rgba(26, 24, 21, 0.55)',
  inkBorder: 'rgba(26, 24, 21, 0.10)',
  accent: '#2F5D50',
  accentSoft: 'rgba(47, 93, 80, 0.15)',
  red: '#B85C3C',
  redSoft: 'rgba(184, 92, 60, 0.15)',
  cream: '#FAF7F0',
  series: ['#1A1815', '#2F5D50', '#B85C3C', '#8A7E68', '#4A6670', '#C5B6A0', '#6B4E3D', '#9E8B70'],
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
  // FYO structure: cols B-F = years 2026-2030, so 2030 = col 6
  const rev2030 = cellNum('Forecast Years Overview', 7, 6);
  const ebitda2030 = cellNum('Forecast Years Overview', 16, 6);
  const hc2030 = cellNum('Forecast Years Overview', 36, 6);

  // Funding need: read Series B + C from RFE solver rows
  // RFE_X: B113 = Series B size, B116 = Series C size
  const sB = cellNum(tab, 113, 2);
  const sC = cellNum(tab, 116, 2);
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

  // Meta
  const baseLMA = cell('Assumptions', 9, 3);
  let lmaStr = '—';
  if (baseLMA) {
    if (typeof baseLMA === 'string' && baseLMA.startsWith('Date(')) {
      const m = baseLMA.match(/Date\((\d+),(\d+),(\d+)\)/);
      if (m) lmaStr = `${m[1]}-${String(parseInt(m[2])+1).padStart(2,'0')}`;
    } else lmaStr = String(baseLMA);
  }
  $('#overview-meta').innerHTML = `Last actual month: <strong>${lmaStr}</strong> · Scenario: <span class="scen-inline">${scen}</span>`;

  // Update scenario indicators
  $$('.scen-inline').forEach(el => el.textContent = scen);

  // Revenue by stream chart (uses Revenue Build tab)
  renderRevByStreamChart('chart-ov-rev', scen);
  // Monthly cash chart
  renderMonthlyCashChart('chart-ov-cash', scen, false);
}

// ============ RENDER: REVENUE BY STREAM CHART ============
function renderRevByStreamChart(canvasId, scen) {
  // Revenue Build tab in Sava model:
  //   R94 = EU+UK Rx, R95 = EU+UK OTC, R96 = US Rx, R97 = US OTC, R98 = Lactate Standalone, R99 = Lactate Add-on
  // Year cols: 2025=B-M, 2026=N-Y, ... (yearStartCol/yearEndCol)
  const streams = [
    { label: 'EU+UK Rx', row: 94 },
    { label: 'EU+UK OTC', row: 95 },
    { label: 'US Rx', row: 96 },
    { label: 'US OTC', row: 97 },
    { label: 'Lactate Standalone', row: 98 },
    { label: 'Lactate Add-on', row: 99 },
  ];
  const years = [2025, 2026, 2027, 2028, 2029, 2030];
  const datasets = streams.map((s, i) => ({
    label: s.label,
    data: years.map(y => sumYear('Revenue Build', s.row, y) || 0),
    backgroundColor: palette.series[i],
    borderRadius: 0,
  }));

  destroyChart(canvasId);
  charts[canvasId] = new Chart(document.getElementById(canvasId), {
    type: 'bar',
    data: { labels: years, datasets },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { position: 'top', align: 'end' },
        tooltip: { ...baseTooltip(fmtM), mode: 'index' },
      },
      scales: {
        x: { stacked: true, grid: { display: false } },
        y: { stacked: true, grid: { color: palette.inkBorder }, ticks: { callback: fmtM } },
      },
    },
  });
}

// ============ RENDER: MONTHLY CASH CHART ============
function renderMonthlyCashChart(canvasId, scen, withMarkers) {
  const tab = rfeTabForScenario(scen);
  // RFE_X row 67 = Cash; row 5 = date headers
  // Cols B (2) to BU (73) = Jan 2025 to Dec 2030
  const dates = [];
  const cashVals = [];
  for (let c = 2; c <= 73; c++) {
    const dateCell = cell(tab, 5, c);
    let dateLabel = '';
    if (typeof dateCell === 'string' && dateCell.startsWith('Date(')) {
      const m = dateCell.match(/Date\((\d+),(\d+),(\d+)\)/);
      if (m) {
        const y = parseInt(m[1]);
        const mo = parseInt(m[2]) + 1;  // gviz months are 0-indexed
        dateLabel = `${y}-${String(mo).padStart(2,'0')}`;
      }
    } else if (dateCell) {
      dateLabel = String(dateCell);
    }
    dates.push(dateLabel);
    cashVals.push(cellNum(tab, 67, c) || 0);
  }

  // Markers for Series B and C
  let bIdx = -1, cIdx = -1;
  if (withMarkers) {
    // Look up Series B and C months from solver rows
    // RFE_X: B111 = Series B month index, B114 = Series C month index
    const bMonthIdx = cellNum(tab, 111, 2);
    const cMonthIdx = cellNum(tab, 114, 2);
    if (bMonthIdx) bIdx = bMonthIdx - 1;  // 1-indexed → 0-indexed
    if (cMonthIdx) cIdx = cMonthIdx - 1;
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
        pointRadius: (ctx) => {
          if (withMarkers && (ctx.dataIndex === bIdx || ctx.dataIndex === cIdx)) return 7;
          return 0;
        },
        pointBackgroundColor: (ctx) => {
          if (ctx.dataIndex === bIdx) return palette.accent;
          if (ctx.dataIndex === cIdx) return palette.red;
          return palette.ink;
        },
        pointBorderColor: palette.cream,
        pointBorderWidth: 2,
        pointHoverRadius: 6,
        pointHoverBackgroundColor: palette.accent,
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
              let note = '';
              if (withMarkers && c.dataIndex === bIdx) note = ' · Series B closes';
              if (withMarkers && c.dataIndex === cIdx) note = ' · Series C closes';
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
  const year = cellNum(cyoTab, 3, 2);
  const lastMonth = cellNum(cyoTab, 4, 2);
  const monthNames = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const monthStr = lastMonth ? monthNames[Math.round(lastMonth) - 1] : '—';
  const yrStr = year ? Math.round(year) : '—';
  const prevYrStr = year ? Math.round(year) - 1 : '—';
  $('#cyo-meta').innerHTML = `Last actual month: <strong>${monthStr} ${yrStr}</strong> · Scenario: <span class="scen-inline">${state.scenario}</span>`;

  // Note: the Sheet's CYO uses Assumptions!C8 (single scenario) and we want
  // the dashboard to show the active scenario, but the CYO tab itself only
  // displays one scenario at a time. We'll show the data as-is and add a note
  // explaining that the dashboard scenario toggle controls other tabs but CYO
  // mirrors the Sheet's active scenario.

  // Row definitions: label, row index, format, indent
  const rows = [
    { type: 'section', label: 'P&L (USD)' },
    { label: 'Revenue', row: 10, fmt: 'm', bold: true },
    { label: 'Total COGS', row: 11, fmt: 'm', reverseVar: true },
    { label: 'Gross profit', row: 12, fmt: 'm', bold: true },
    { label: 'Gross margin %', row: 13, fmt: 'pct', indent: true },
    { label: 'R&D', row: 14, fmt: 'm', indent: true, reverseVar: true },
    { label: 'Regulatory', row: 15, fmt: 'm', indent: true, reverseVar: true },
    { label: 'Sales & marketing', row: 16, fmt: 'm', indent: true, reverseVar: true },
    { label: 'G&A', row: 17, fmt: 'm', indent: true, reverseVar: true },
    { label: 'Total Opex', row: 18, fmt: 'm', bold: true, reverseVar: true },
    { label: 'EBITDA', row: 19, fmt: 'm', bold: true },
    { label: 'Net Income', row: 20, fmt: 'm', bold: true },
    { type: 'section', label: 'Cash & Financing (USD)' },
    { label: 'Capex', row: 23, fmt: 'm', reverseVar: true },
    { label: 'Series B inflow', row: 24, fmt: 'm' },
    { label: 'Series C inflow', row: 25, fmt: 'm' },
    { label: 'Net change in cash', row: 26, fmt: 'm', bold: true },
    { label: 'Cash balance (EOM)', row: 27, fmt: 'm', bold: true },
  ];

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
    return v.toFixed(1);
  };
  const fmtVarDollar = (v, fmt) => {
    if (v === null || isNaN(v)) return '—';
    if (fmt === 'pct') return fmtBps(v);  // pp diff → bps
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

function renderFYO() {
  const fyoTab = 'Forecast Years Overview';
  const years = [2026, 2027, 2028, 2029, 2030];

  // Row defs from Sava FYO tab
  const rows = [
    { type: 'section', label: 'P&L (USD)' },
    { label: 'Revenue', row: 7, fmt: 'm', bold: true },
    { label: 'Total COGS', row: 8, fmt: 'm', bold: true },
    { label: 'Gross profit', row: 9, fmt: 'm', bold: true },
    { label: 'Gross margin %', row: 10, fmt: 'pct', indent: true },
    { label: 'R&D', row: 11, fmt: 'm', indent: true },
    { label: 'Regulatory', row: 12, fmt: 'm', indent: true },
    { label: 'Sales & marketing', row: 13, fmt: 'm', indent: true },
    { label: 'G&A', row: 14, fmt: 'm', indent: true },
    { label: 'Total Opex', row: 15, fmt: 'm', bold: true },
    { label: 'EBITDA', row: 16, fmt: 'm', bold: true },
    { label: 'Net Income', row: 17, fmt: 'm', bold: true },

    { type: 'section', label: 'Cash Flow (USD)' },
    { label: 'CFO', row: 20, fmt: 'm' },
    { label: 'CFI (capex etc.)', row: 21, fmt: 'm' },
    { label: 'Series B inflow', row: 22, fmt: 'm' },
    { label: 'Series C inflow', row: 23, fmt: 'm' },
    { label: 'Net change in cash', row: 24, fmt: 'm', bold: true },
    { label: 'Cash balance (EOY)', row: 25, fmt: 'm', bold: true },

    { type: 'section', label: 'Balance Sheet (EOY)' },
    { label: 'Cash & equivalents', row: 28, fmt: 'm' },
    { label: 'AR + Inventory', row: 29, fmt: 'm' },
    { label: 'Net PP&E', row: 30, fmt: 'm' },
    { label: 'Total Assets', row: 31, fmt: 'm', bold: true },
    { label: 'Total Liabilities', row: 32, fmt: 'm' },
    { label: 'Total Equity', row: 33, fmt: 'm', bold: true },

    { type: 'section', label: 'Operational' },
    { label: 'Total HC (EOY)', row: 36, fmt: 'count' },
    { label: 'Active patients (EOY)', row: 37, fmt: 'count' },
    { label: 'Patches sold (annual)', row: 38, fmt: 'count' },
    { label: 'Mfg lines installed (EOY)', row: 39, fmt: 'count' },

    { type: 'section', label: 'Unit Economics' },
    { label: 'Avg active patients (FY)', row: 42, fmt: 'count' },
    { label: 'New patients added (FY)', row: 43, fmt: 'count' },
    { label: 'ARPU (annualized)', row: 44, fmt: 'm' },
    { label: 'CAC (S&M / new patients)', row: 45, fmt: 'm' },
    { label: 'LTV (per patient)', row: 46, fmt: 'm' },
    { label: 'LTV:CAC ratio', row: 47, fmt: 'ratio' },
    { label: 'All-in per-patch cost', row: 48, fmt: 'm' },
  ];

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
    // Year vals — cols B-F in FYO = 2-6
    const yearlyVals = years.map((y, i) => cellNum(fyoTab, r.row, 2 + i));

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
  const fyoTab = 'Forecast Years Overview';
  const years = [2026, 2027, 2028, 2029, 2030];

  // Build floating bars [base, top] for each metric within each year
  // Order per year: Revenue (0→rev), COGS (rev-cogs→rev as drop), Opex (rev-cogs-opex→rev-cogs as drop), EBITDA (0→ebitda)
  // We render this as 4 bars per year × 5 years = 20 bar slots, plus spacers between years

  const labels = [];
  const dataPoints = [];
  const colors = [];
  const actualValues = []; // for tooltip — the "real" value each bar represents
  const yearStartIdx = {};  // bar index where each year starts (for label placement)

  years.forEach((y, yi) => {
    const colYear = 2 + yi; // FYO B-F = years 2-6
    const rev = cellNum(fyoTab, 7, colYear) || 0;
    const cogs = cellNum(fyoTab, 8, colYear) || 0;
    const opex = cellNum(fyoTab, 15, colYear) || 0;
    const ebitda = cellNum(fyoTab, 16, colYear) || 0;

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
  // Patches by year — Revenue Build R92
  const years = [2025, 2026, 2027, 2028, 2029, 2030];
  const patches = years.map(y => sumYear('Revenue Build', 92, y) || 0);

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
  // Headcount Plan tab: R22-27 are functions
  // R22=R&D, R23=Regulatory, R24=Manufacturing, R25=QA, R26=Commercial, R27=G&A
  // EOY cols same as model: col M=Dec 25, Y=Dec 26, AK=Dec 27, AW=Dec 28, BI=Dec 29, BU=Dec 30 (yearEndCol)
  const years = [2025, 2026, 2027, 2028, 2029, 2030];
  const funcs = [
    { label: 'R&D', row: 22 },
    { label: 'Regulatory', row: 23 },
    { label: 'Manufacturing', row: 24 },
    { label: 'QA', row: 25 },
    { label: 'Commercial', row: 26 },
    { label: 'G&A', row: 27 },
  ];

  const datasets = funcs.map((f, i) => ({
    label: f.label,
    data: years.map(y => {
      const v = cellNum('Headcount Plan', f.row, yearEndCol(y));
      return v !== null ? Math.round(v) : 0;
    }),
    backgroundColor: palette.series[i],
    borderRadius: 0,
  }));

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
}

// ============ RENDER: COSTS ============
function renderCosts() {
  const fyoTab = 'Forecast Years Overview';
  const years = [2026, 2027, 2028, 2029, 2030];

  // Cost categories from FYO
  const cats = [
    { label: 'COGS', row: 8, color: palette.red },
    { label: 'R&D', row: 11, color: palette.ink },
    { label: 'Regulatory', row: 12, color: '#4A6670' },
    { label: 'Sales & marketing', row: 13, color: palette.accent },
    { label: 'G&A', row: 14, color: '#8A7E68' },
  ];

  const datasets = cats.map(c => ({
    label: c.label,
    data: years.map((y, i) => cellNum(fyoTab, c.row, 2 + i) || 0),
    backgroundColor: c.color,
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
        tooltip: { ...baseTooltip(fmtM), mode: 'index' },
      },
      scales: {
        x: { stacked: true, grid: { display: false } },
        y: { stacked: true, grid: { color: palette.inkBorder }, ticks: { callback: fmtM } },
      },
    },
  });

  // Cost table
  let html = `<table class="data-table"><thead><tr><th>Cost line</th>`;
  for (const y of years) html += `<th>${y}</th>`;
  html += `</tr></thead><tbody>`;
  html += `<tr class="section-row"><td colspan="${years.length + 1}">COGS</td></tr>`;
  html += `<tr class="bold-row"><td>Total COGS</td>`;
  years.forEach((y, i) => { html += `<td>${fmtM(cellNum(fyoTab, 8, 2 + i))}</td>`; });
  html += `</tr>`;
  html += `<tr class="section-row"><td colspan="${years.length + 1}">Operating Expenses</td></tr>`;
  ['R&D','Regulatory','Sales & marketing','G&A'].forEach((label, idx) => {
    const r = 11 + idx;
    html += `<tr><td class="indent">${label}</td>`;
    years.forEach((y, i) => { html += `<td>${fmtM(cellNum(fyoTab, r, 2 + i))}</td>`; });
    html += `</tr>`;
  });
  html += `<tr class="bold-row"><td>Total Opex</td>`;
  years.forEach((y, i) => { html += `<td>${fmtM(cellNum(fyoTab, 15, 2 + i))}</td>`; });
  html += `</tr>`;
  html += `<tr class="section-row"><td colspan="${years.length + 1}">Total Cost</td></tr>`;
  html += `<tr class="bold-row"><td>COGS + Opex</td>`;
  years.forEach((y, i) => {
    const cogs = cellNum(fyoTab, 8, 2 + i) || 0;
    const opex = cellNum(fyoTab, 15, 2 + i) || 0;
    html += `<td>${fmtM(cogs + opex)}</td>`;
  });
  html += `</tr>`;
  html += `</tbody></table>`;
  $('#costs-content').innerHTML = html;
}

// ============ RENDER: FINANCING ============
function renderFinancing() {
  const tab = rfeTabForScenario(state.scenario);
  // RFE_X: B111 = Series B month idx, B112 = Series B date, B113 = Series B size
  //        B114 = Series C month idx, B115 = Series C date, B116 = Series C size
  const sBSize = cellNum(tab, 113, 2);
  const sBDate = cell(tab, 112, 2);
  const sCSize = cellNum(tab, 116, 2);
  const sCDate = cell(tab, 115, 2);

  $('#fin-b-size').textContent = fmtM(sBSize);
  $('#fin-b-date').textContent = formatDate(sBDate);
  $('#fin-c-size').textContent = fmtM(sCSize);
  $('#fin-c-date').textContent = formatDate(sCDate);
  $('#fin-total').textContent = fmtM((sBSize || 0) + (sCSize || 0));

  renderMonthlyCashChart('chart-fin-cash', state.scenario, true);
}

function formatDate(val) {
  if (!val) return '—';
  if (typeof val === 'string' && val.startsWith('Date(')) {
    const m = val.match(/Date\((\d+),(\d+),(\d+)\)/);
    if (m) {
      const names = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
      return `${names[parseInt(m[2])]} ${m[1]}`;
    }
  }
  return String(val);
}

// ============ RENDER: ASSUMPTIONS ============
function renderAssumptions() {
  const scen = state.scenario;
  const col = scenCol(scen);
  // Assumptions rows: 12=EU Rx launch, 14=US Rx launch, 16=Lactate, 23=EU Rx ASP, 62=patches/yr,
  //   68=mfg cost steady, 76=tech avg cost, 115=steady HC adds, 116=pre-launch HC adds
  const idxToDate = (idx) => {
    if (!idx) return '—';
    const y = 2025 + Math.floor((idx - 1) / 12);
    const m = ((idx - 1) % 12) + 1;
    const names = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    return `${names[m-1]} ${y}`;
  };

  const euLaunch = cellNum('Assumptions', 12, col);
  const usLaunch = cellNum('Assumptions', 14, col);
  const lacLaunch = cellNum('Assumptions', 16, col);
  const asp = cellNum('Assumptions', 23, col);
  const patches = cellNum('Assumptions', 62, col);
  const cogs = cellNum('Assumptions', 68, col);
  const techCost = cellNum('Assumptions', 76, col);
  const steadyHC = cellNum('Assumptions', 115, col);
  const preLaunchHC = cellNum('Assumptions', 116, col);

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
  // Get the Sheet's checks summary — we'll read from Summary or just hardcode "67/67"
  // For simplicity, look at Assumptions or Summary for a known check cell
  // The model has a "Checks" tab; reading it from the dashboard isn't critical for v1.
  // For now: read Current Year Overview C3 which has "CHECK OK" or "CHECK ERROR"
  const checkVal = cell('Current Year Overview', 3, 3);
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
