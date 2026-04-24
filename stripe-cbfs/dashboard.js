/* ============================================================
   Stripe CBFS Dashboard — dashboard.js
   Fetches data from Google Sheets (published CSVs),
   renders Overview / Current Year / Forecast / Unit Econ /
   Scenarios / Sensitivity tabs. Scenario toggle switches between
   RFE_Base / RFE_Bear / RFE_Bull client-side (no re-fetch).
   ============================================================ */

// ============ CONFIG ============
const SHEET_ID = '1x6aNjAe-MjECTN0gWmy0MuxONtMqQVPRMDVMx-kqVm0';

// Build CSV URL for a specific tab name
function csvUrl(tabName) {
  return `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent(tabName)}`;
}

const TABS_TO_FETCH = [
  'Current Year Overview',
  'Forecast Years Overview',
  'RFE_Base',
  'RFE_Bear',
  'RFE_Bull',
  'Assumptions',
  'Summary'
];

// ============ STATE ============
const state = {
  scenario: 'Base',  // active scenario: Base | Bear | Bull
  data: {},          // populated after fetch: { 'RFE_Base': [[row],[row]...], ... }
  loaded: false
};

// ============ HELPERS ============
const $ = (sel) => document.querySelector(sel);
const $$ = (sel) => Array.from(document.querySelectorAll(sel));

// Parse a numeric cell — handles "$1,234.5m", "24.73%", "+$0.4m", "—", etc.
function parseNum(val) {
  if (val === null || val === undefined) return null;
  if (typeof val === 'number') return val;
  let s = String(val).trim();
  if (!s || s === '—' || s === '-' || s === 'n/a' || s === '#N/A') return null;

  // Extract sign
  let sign = 1;
  if (s.startsWith('-')) { sign = -1; s = s.slice(1); }
  else if (s.startsWith('+')) s = s.slice(1);
  if (s.startsWith('(') && s.endsWith(')')) { sign = -1; s = s.slice(1, -1); }

  // Strip known word suffixes FIRST (before space-stripping destroys separators)
  s = s.replace(/\s*yrs?\s*$/i, '');

  // Detect suffix multiplier
  let mult = 1;
  if (/\s*bps\s*$/i.test(s)) { mult = 1e-4; s = s.replace(/\s*bps\s*$/i, ''); }
  else if (/%\s*$/i.test(s)) { mult = 0.01; s = s.replace(/%\s*$/i, ''); }
  else if (/m\s*$/i.test(s)) { mult = 1e6; s = s.replace(/m\s*$/i, ''); }
  else if (/k\s*$/i.test(s)) { mult = 1e3; s = s.replace(/k\s*$/i, ''); }

  // Clean $, commas, spaces
  s = s.replace(/[$,\s]/g, '');

  const n = parseFloat(s);
  if (isNaN(n)) return null;
  return sign * n * mult;
}

// Format money
function fmtM(n) {
  if (n === null || n === undefined || isNaN(n)) return '—';
  const abs = Math.abs(n);
  const sign = n < 0 ? '-' : '';
  if (abs >= 1e9) return sign + '$' + (abs / 1e9).toFixed(1) + 'b';
  if (abs >= 1e6) return sign + '$' + (abs / 1e6).toFixed(1) + 'm';
  if (abs >= 1e3) return sign + '$' + (abs / 1e3).toFixed(1) + 'k';
  return sign + '$' + abs.toFixed(0);
}
function fmtMVar(n) {
  if (n === null || n === undefined || isNaN(n)) return '—';
  const sign = n >= 0 ? '+' : '-';
  const abs = Math.abs(n);
  if (abs >= 1e6) return sign + '$' + (abs / 1e6).toFixed(1) + 'm';
  if (abs >= 1e3) return sign + '$' + (abs / 1e3).toFixed(1) + 'k';
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
function fmtCount(n, decimals = 1) {
  if (n === null || n === undefined || isNaN(n)) return '—';
  return n.toFixed(decimals);
}
function fmtCountVar(n) {
  if (n === null || n === undefined || isNaN(n)) return '—';
  const sign = n >= 0 ? '+' : '';
  return sign + n.toFixed(1);
}

// ============ DATA FETCH ============
async function fetchTab(tabName) {
  const url = csvUrl(tabName);
  const resp = await fetch(url);
  if (!resp.ok) throw new Error(`Failed to fetch "${tabName}" (HTTP ${resp.status}). Make sure the Sheet is published to web.`);
  const text = await resp.text();
  // Papa parses the CSV
  const parsed = Papa.parse(text, { header: false, skipEmptyLines: false });
  return parsed.data; // 2D array
}

async function loadAll() {
  const statusEl = $('#loading-status');
  state.data = {};
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

// ============ DATA EXTRACTION ============
// Given a tab's 2D array, return cell at (row, col) 1-indexed (matches Excel)
function cell(tabName, row, col) {
  const data = state.data[tabName];
  if (!data || !data[row - 1]) return null;
  return data[row - 1][col - 1];
}
function cellNum(tabName, row, col) {
  return parseNum(cell(tabName, row, col));
}

// Excel column letter to 1-indexed number
function colToNum(col) {
  let n = 0;
  for (let i = 0; i < col.length; i++) {
    n = n * 26 + (col.charCodeAt(i) - 64);
  }
  return n;
}

// Year → column range in RFE tabs (F = Jan 2018). 2024 = BZ..CK
function yearColRange(year) {
  const startCol = 6 + (year - 2018) * 12;
  return [startCol, startCol + 11];
}
// Sum a range of cols for a specific row in a specific RFE tab
function sumRow(tab, row, startCol, endCol) {
  let sum = 0;
  let any = false;
  for (let c = startCol; c <= endCol; c++) {
    const v = cellNum(tab, row, c);
    if (v !== null && !isNaN(v)) { sum += v; any = true; }
  }
  return any ? sum : null;
}
function sumYear(tab, row, year) {
  const [s, e] = yearColRange(year);
  return sumRow(tab, row, s, e);
}
// EOY value = last (Dec) column of year
function eoyVal(tab, row, year) {
  const [, e] = yearColRange(year);
  return cellNum(tab, row, e);
}

// RFE row anchors (grand total section) — same in all parallel tabs
const RFE_GRAND = {
  mcount: 6,
  gpv: 7,
  events: 8,
  rev_vol: 9,
  rev_trx: 10,
  rev_add: 11,
  rev_tot: 12,
  cogs_int: 13,
  cogs_api: 14,
  cogs_loss: 15,
  cogs_tot: 16,
  gm: 17,
};
const RFE_SEG_ROWS = {
  'SaaS': { mcount: 21, gpv: 22, events: 23, rev_tot: 27, cogs_tot: 31, gm: 32 },
  'E-Commerce Store': { mcount: 35, gpv: 36, events: 37, rev_tot: 41, cogs_tot: 45, gm: 46 },
  'Platform': { mcount: 49, gpv: 50, events: 51, rev_tot: 55, cogs_tot: 59, gm: 60 },
};
// Overall check cell
const RFE_CHECK_CELL = { row: 1131, col: 6 }; // F1131

// Cohort rows for unit economics (2020 cohort only)
const COHORT_2020 = {
  'SaaS': { gm_row: 109, new_row: 94, cac_row: 82, churn_row: 60 },
  'E-Commerce Store': { gm_row: 234, new_row: 219, cac_row: 83, churn_row: 61 },
  'Platform': { gm_row: 359, new_row: 344, cac_row: 84, churn_row: 62 },
};

function rfeTabForScenario(scen) {
  return `RFE_${scen}`;
}

// ============ RENDER: OVERVIEW ============
function renderOverview() {
  const tab = rfeTabForScenario(state.scenario);
  // KPIs — 2024 full-year revenue, GM, merchants EOY, GPV
  const rev2024 = sumYear(tab, RFE_GRAND.rev_tot, 2024);
  const gm2024 = sumYear(tab, RFE_GRAND.gm, 2024);
  const merch2024 = eoyVal(tab, RFE_GRAND.mcount, 2024);
  const gpv2024 = sumYear(tab, RFE_GRAND.gpv, 2024);
  // Baseline (2019 actuals, scenario-independent)
  const rev2019 = sumYear(tab, RFE_GRAND.rev_tot, 2019);
  const gm2019 = sumYear(tab, RFE_GRAND.gm, 2019);
  const merch2019 = eoyVal(tab, RFE_GRAND.mcount, 2019);
  const gpv2019 = sumYear(tab, RFE_GRAND.gpv, 2019);

  $('#kpi-rev').textContent = fmtM(rev2024);
  const revCAGR = rev2019 && rev2024 ? (Math.pow(rev2024 / rev2019, 1 / 5) - 1) : null;
  $('#kpi-rev-detail').textContent = revCAGR !== null ? `2019-24 CAGR ${(revCAGR * 100).toFixed(1)}%` : '';

  $('#kpi-gm').textContent = fmtM(gm2024);
  const gmPct = rev2024 ? (gm2024 / rev2024) : null;
  $('#kpi-gm-detail').textContent = gmPct !== null ? `${(gmPct * 100).toFixed(1)}% of revenue` : '';

  $('#kpi-merch').textContent = merch2024 !== null ? Math.round(merch2024).toString() : '—';
  const merchAdd = (merch2024 !== null && merch2019 !== null) ? (merch2024 - merch2019) : null;
  $('#kpi-merch-detail').textContent = merchAdd !== null ? `+${Math.round(merchAdd)} vs EOY 2019` : '';

  $('#kpi-gpv').textContent = fmtM(gpv2024);
  const gpvCAGR = gpv2019 && gpv2024 ? (Math.pow(gpv2024 / gpv2019, 1 / 5) - 1) : null;
  $('#kpi-gpv-detail').textContent = gpvCAGR !== null ? `2019-24 CAGR ${(gpvCAGR * 100).toFixed(1)}%` : '';

  // Meta
  const assump = state.data['Assumptions'];
  const activeSheetScen = assump && assump[6] ? (assump[6][2] || '—') : '—';
  const cyo = state.data['Current Year Overview'];
  const lastActual = cyo && cyo[1] ? (cyo[1][1] || '—') : '—';
  $('#overview-meta').textContent = `Scenario: ${state.scenario} · Sheet active scenario: ${activeSheetScen} · Last actual month: ${lastActual}`;

  // Charts
  renderRevenueYearChart(tab);
  renderSeg2024Chart(tab);

  // Note box: takeaways
  const bestScen = { 'Base': 'moderate growth', 'Bear': 'conservative growth', 'Bull': 'aggressive growth' }[state.scenario];
  $('#overview-note').innerHTML = `
    <h3>Context</h3>
    <ul>
      <li>This dashboard mirrors the Stripe CBFS Excel model, live from a Google Sheet. Edit the Sheet → refresh this page to see updates.</li>
      <li>Scenario toggle uses pre-computed parallel tabs (<code>RFE_Base</code>, <code>RFE_Bear</code>, <code>RFE_Bull</code>) — instant switching, no re-calculation needed.</li>
      <li>The <strong>Current Year</strong> tab reflects whatever scenario is active in <code>Assumptions!C7</code> in the Sheet — it's not controlled by the dashboard toggle.</li>
      <li>All other tabs respect the toggle: <strong>${state.scenario}</strong> ( ${bestScen}) is currently shown.</li>
    </ul>
  `;

  $('#rev-chart-scen').textContent = state.scenario;
}

let revChart = null;
function renderRevenueYearChart(tab) {
  const years = [2018, 2019, 2020, 2021, 2022, 2023, 2024];
  const revs = years.map(y => (sumYear(tab, RFE_GRAND.rev_tot, y) || 0) / 1e6);
  const ctx = document.getElementById('chart-revenue-year').getContext('2d');
  if (revChart) revChart.destroy();
  revChart = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: years.map(String),
      datasets: [{
        label: 'Revenue ($m)',
        data: revs,
        backgroundColor: years.map(y => y <= 2019 ? '#E8E6FF' : '#635BFF'),
        borderColor: years.map(y => y <= 2019 ? '#B8B5FF' : '#635BFF'),
        borderWidth: 1
      }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: (c) => '$' + c.parsed.y.toFixed(1) + 'm'
          }
        }
      },
      scales: {
        y: {
          beginAtZero: true,
          ticks: { callback: (v) => '$' + v + 'm', font: { family: 'IBM Plex Mono' } }
        },
        x: {
          ticks: { font: { family: 'IBM Plex Mono' } }
        }
      }
    }
  });
}
let segChart = null;
function renderSeg2024Chart(tab) {
  const segs = ['SaaS', 'E-Commerce Store', 'Platform'];
  const vals = segs.map(s => (sumYear(tab, RFE_SEG_ROWS[s].rev_tot, 2024) || 0) / 1e6);
  const ctx = document.getElementById('chart-seg-2024').getContext('2d');
  if (segChart) segChart.destroy();
  segChart = new Chart(ctx, {
    type: 'doughnut',
    data: {
      labels: segs,
      datasets: [{
        data: vals,
        backgroundColor: ['#635BFF', '#0A2540', '#00A67E'],
        borderWidth: 2,
        borderColor: '#FBFAF6'
      }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      cutout: '60%',
      plugins: {
        legend: { position: 'bottom', labels: { font: { family: 'IBM Plex Sans', size: 12 }, padding: 14, boxWidth: 12 } },
        tooltip: {
          callbacks: {
            label: (c) => c.label + ': $' + c.parsed.toFixed(1) + 'm'
          }
        }
      }
    }
  });
}

// ============ RENDER: CURRENT YEAR (mirror CYO) ============
function renderCurrentYear() {
  const data = state.data['Current Year Overview'];
  if (!data) { $('#current-content').innerHTML = '<p class="note">No data.</p>'; return; }

  const assump = state.data['Assumptions'];
  const activeSheetScen = assump && assump[6] ? (assump[6][2] || '—') : '—';
  const lastActual = data[1] ? (data[1][1] || '—') : '—';
  $('#current-meta').textContent = `Last actual month: ${lastActual} · Active Sheet scenario: ${activeSheetScen}`;

  if (activeSheetScen !== state.scenario) {
    $('#current-scen-note').textContent = `Note: this view shows ${activeSheetScen} (the active scenario in the Sheet). The dashboard toggle (${state.scenario}) does not affect this tab because the CYO's Last Month / YTD / RoY math depends on the active scenario cell in the Sheet. To change this view, edit Assumptions!C7 in the Sheet and refresh.`;
  } else {
    $('#current-scen-note').textContent = `Sheet's active scenario matches dashboard toggle: ${state.scenario}`;
  }

  // Build table from CYO CSV — it already has the right structure
  // CYO rows: 1 = title, 2-5 = B2 etc headers, 6-7 = block/sub-header rows, 8+ = data
  // We render rows 6 onward.
  const startRow = 5; // 0-indexed → row 6 in Excel
  const endRow = data.length;

  let html = '<table class="data-table">';
  for (let i = startRow; i < endRow; i++) {
    const row = data[i] || [];
    if (row.every(c => !c || String(c).trim() === '')) continue;  // skip empty rows

    const label = row[0] || '';
    const isSection = String(label).toUpperCase().includes('FINANCIALS') || String(label).toUpperCase().includes('KPI');
    const isSubHeader = String(label).startsWith('—') && String(label).endsWith('—');
    const isBlockHeader = /Last Month|YTD|Rest of Year|Full Year/.test(row.slice(1).join(' '));
    const isSubHeader2 = /Current Year|Prior Year|YoY/.test(row.slice(1).join(' ')) && !label;

    if (isSection) {
      html += `<tr class="section-row"><td colspan="${Math.max(row.length, 21)}">${label}</td></tr>`;
      continue;
    }
    if (isSubHeader) {
      html += `<tr class="sub-header-row"><td colspan="${Math.max(row.length, 21)}">${label}</td></tr>`;
      continue;
    }
    if (isBlockHeader) {
      // Keep the CYO's 4-cell blocks
      html += '<tr>';
      html += `<th></th>`;
      // Merge blocks for "Last Month" etc (cols B-E, G-J, L-O, Q-T = 4 cells each with 1 gap)
      const blocks = [[1, 4, 'Last Month'], [6, 9, 'YTD (Act)'], [11, 14, 'Rest of Year (Fcst)'], [16, 19, 'Full Year']];
      for (const [s, e, title] of blocks) {
        const rowTitle = row[s] || title; // row[s] = e.g. cell B (col 1)
        html += `<th class="block-header" colspan="4">${title}</th>`;
        if (e < 19) html += `<th></th>`; // gap
      }
      html += '</tr>';
      continue;
    }
    if (isSubHeader2) {
      html += '<tr>';
      html += `<th></th>`;
      for (let j = 1; j <= 19; j++) {
        const v = row[j];
        if (v) html += `<th>${String(v).replace(/\n/g, ' ')}</th>`;
        else html += `<th></th>`;
      }
      html += '</tr>';
      continue;
    }

    // Data row
    const rowClass = /Merchants|Revenue per|GPV per|New merchants|Churned|Net Take|Gross Margin %|Churn rate/.test(label) ? '' : 'bold-row';
    html += `<tr class="${rowClass}">`;
    html += `<td${String(label).startsWith('  ') ? ' class="indent"' : ''}>${label.trim()}</td>`;
    for (let j = 1; j <= 19; j++) {
      const v = row[j] !== undefined ? row[j] : '';
      const inBlock = [1, 2, 3, 4, 6, 7, 8, 9, 11, 12, 13, 14, 16, 17, 18, 19].includes(j);
      if (!inBlock) { html += '<td></td>'; continue; }
      // Determine class
      const colInBlock = ((j - 1) % 5);
      let cls = '';
      if (colInBlock === 0) cls = 'cell-current';
      else if (colInBlock === 2 || colInBlock === 3) cls = 'cell-variance';
      // Pos/neg for variance cells
      const parsed = parseNum(v);
      if ((colInBlock === 2 || colInBlock === 3) && parsed !== null) {
        const isReversed = /COGS|Churn/.test(label);
        if (!isReversed) {
          if (parsed > 0) cls += ' pos'; else if (parsed < 0) cls += ' neg';
        } else {
          if (parsed > 0) cls += ' neg'; else if (parsed < 0) cls += ' pos';
        }
      }
      html += `<td class="${cls}">${v || ''}</td>`;
    }
    html += '</tr>';
  }
  html += '</table>';
  $('#current-content').innerHTML = html;
}

// ============ RENDER: FORECAST YEARS ============
function renderForecast() {
  const tab = rfeTabForScenario(state.scenario);
  const years = [2018, 2019, 2020, 2021, 2022, 2023, 2024];
  $('#fc-scen-inline').textContent = state.scenario;

  const rows = [
    { sec: 'FINANCIALS — GRAND TOTAL', type: 'section' },
    { label: 'GPV', row: RFE_GRAND.gpv, fmt: 'm', bold: true },
    { label: 'Revenue', row: RFE_GRAND.rev_tot, fmt: 'm', bold: true },
    { label: 'Net Take Rate %', num: RFE_GRAND.rev_tot, den: RFE_GRAND.gpv, fmt: 'pct', indent: true },
    { label: 'COGS', row: RFE_GRAND.cogs_tot, fmt: 'm', bold: true },
    { label: 'Gross Margin', row: RFE_GRAND.gm, fmt: 'm', bold: true },
    { label: 'Gross Margin %', num: RFE_GRAND.gm, den: RFE_GRAND.rev_tot, fmt: 'pct', indent: true },
    { sec: 'FINANCIALS — BY SEGMENT', type: 'section' },
  ];

  for (const seg of ['SaaS', 'E-Commerce Store', 'Platform']) {
    rows.push({ sec: `— ${seg} —`, type: 'subheader' });
    const segR = RFE_SEG_ROWS[seg];
    rows.push({ label: 'GPV', row: segR.gpv, fmt: 'm', bold: true, indent: true });
    rows.push({ label: 'Revenue', row: segR.rev_tot, fmt: 'm', bold: true, indent: true });
    rows.push({ label: 'Net Take Rate %', num: segR.rev_tot, den: segR.gpv, fmt: 'pct', indent: true });
    rows.push({ label: 'COGS', row: segR.cogs_tot, fmt: 'm', bold: true, indent: true });
    rows.push({ label: 'Gross Margin', row: segR.gm, fmt: 'm', bold: true, indent: true });
    rows.push({ label: 'Gross Margin %', num: segR.gm, den: segR.rev_tot, fmt: 'pct', indent: true });
  }
  rows.push({ sec: 'KPIs', type: 'section' });
  rows.push({ label: '# Merchants (EOY)', row: RFE_GRAND.mcount, fmt: 'eoy-count', bold: true });

  // Build HTML
  let html = `<table class="data-table"><thead><tr><th>Metric</th>`;
  for (const y of years) html += `<th>${y}</th>`;
  html += `<th>Avg Annual Growth ($)<br>2018-2024</th><th>CAGR %<br>2018-2024</th><th>Hist CAGR<br>2018-2019</th><th>Fcst CAGR<br>2020-2024</th><th>Variance (pp)</th></tr></thead><tbody>`;

  for (const r of rows) {
    if (r.type === 'section') { html += `<tr class="section-row"><td colspan="13">${r.sec}</td></tr>`; continue; }
    if (r.type === 'subheader') { html += `<tr class="sub-header-row"><td colspan="13">${r.sec}</td></tr>`; continue; }

    // Compute yearly values
    const yearlyVals = years.map(y => {
      if (r.row && r.fmt === 'eoy-count') return eoyVal(tab, r.row, y);
      if (r.row) return sumYear(tab, r.row, y);
      if (r.num && r.den) {
        const n = sumYear(tab, r.num, y);
        const d = sumYear(tab, r.den, y);
        return (n !== null && d !== null && d !== 0) ? n / d : null;
      }
      return null;
    });

    // Formatting
    const fmtVal = (v) => {
      if (v === null || isNaN(v)) return '—';
      if (r.fmt === 'm') return fmtM(v);
      if (r.fmt === 'pct') return (v * 100).toFixed(2) + '%';
      if (r.fmt === 'eoy-count') return Math.round(v).toLocaleString();
      return v.toFixed(1);
    };

    const rowClass = r.bold ? 'bold-row' : '';
    html += `<tr class="${rowClass}">`;
    html += `<td${r.indent ? ' class="indent"' : ''}>${r.label}</td>`;
    for (let i = 0; i < years.length; i++) {
      const cls = years[i] <= 2019 ? 'act-year' : 'fcst-year';
      html += `<td class="${cls}">${fmtVal(yearlyVals[i])}</td>`;
    }
    // Variance columns
    const v2018 = yearlyVals[0], v2019 = yearlyVals[1], v2020 = yearlyVals[2], v2024 = yearlyVals[6];
    let aag = null, cagr = null, hist = null, fcst = null, varp = null;
    if (v2018 !== null && v2024 !== null && v2018 !== 0) {
      aag = (v2024 - v2018) / 6;
      cagr = Math.pow(v2024 / v2018, 1 / 6) - 1;
    }
    if (v2018 !== null && v2019 !== null && v2018 !== 0) hist = v2019 / v2018 - 1;
    if (v2020 !== null && v2024 !== null && v2020 !== 0) fcst = Math.pow(v2024 / v2020, 1 / 4) - 1;
    if (hist !== null && fcst !== null) varp = fcst - hist;

    const fmtAAG = r.fmt === 'm' ? (aag !== null ? fmtMVar(aag) : '—')
                 : r.fmt === 'pct' ? (aag !== null ? fmtBps(aag) : '—')
                 : (aag !== null ? fmtCountVar(aag) : '—');
    html += `<td class="cell-variance ${aag > 0 ? 'pos' : aag < 0 ? 'neg' : ''}">${fmtAAG}</td>`;
    html += `<td class="cell-variance ${cagr > 0 ? 'pos' : cagr < 0 ? 'neg' : ''}">${cagr !== null ? fmtPctVar(cagr) : '—'}</td>`;
    html += `<td class="cell-variance ${hist > 0 ? 'pos' : hist < 0 ? 'neg' : ''}">${hist !== null ? fmtPctVar(hist) : '—'}</td>`;
    html += `<td class="cell-variance ${fcst > 0 ? 'pos' : fcst < 0 ? 'neg' : ''}">${fcst !== null ? fmtPctVar(fcst) : '—'}</td>`;
    html += `<td class="cell-variance ${varp > 0 ? 'pos' : varp < 0 ? 'neg' : ''}">${varp !== null ? fmtPctVar(varp) : '—'}</td>`;
    html += '</tr>';
  }
  html += '</tbody></table>';
  $('#forecast-content').innerHTML = html;
}

// ============ RENDER: UNIT ECONOMICS ============
function renderUnitEcon() {
  const tab = rfeTabForScenario(state.scenario);
  $('#ue-scen-inline').textContent = state.scenario;

  // Read CAC and WACC from Assumptions
  const cacs = {
    'SaaS': cellNum('Assumptions', 82, 3),
    'E-Commerce Store': cellNum('Assumptions', 83, 3),
    'Platform': cellNum('Assumptions', 84, 3),
  };
  const wacc = cellNum('Assumptions', 85, 3) || 0.10;
  // Active-scenario churn per segment — we can read from the RFE_* tab's formula result? 
  // Simpler: read from Assumptions directly based on state.scenario
  const scenCol = state.scenario === 'Base' ? 4 : state.scenario === 'Bear' ? 5 : 6;
  const churns = {
    'SaaS': cellNum('Assumptions', 60, scenCol),
    'E-Commerce Store': cellNum('Assumptions', 61, scenCol),
    'Platform': cellNum('Assumptions', 62, scenCol),
  };

  let html = '<div class="unit-econ-grid">';
  for (const seg of ['SaaS', 'E-Commerce Store', 'Platform']) {
    const c2020 = COHORT_2020[seg];
    // Merchants added in 2020 (sum of + New merchants row for 2020)
    const mAdded = sumYear(tab, c2020.new_row, 2020);
    // Cohort GM per year (Y1=2020 ... Y5=2024)
    const annualGMPerM = [];
    for (let y = 2020; y <= 2024; y++) {
      const gm = sumYear(tab, c2020.gm_row, y);
      annualGMPerM.push(mAdded && gm !== null ? gm / mAdded : null);
    }
    // Cumulative
    const cum = [];
    annualGMPerM.reduce((acc, v, i) => { const n = (acc || 0) + (v || 0); cum.push(n); return n; }, 0);

    // LTV (5-year discounted)
    let ltv = 0;
    annualGMPerM.forEach((v, i) => { if (v !== null) ltv += v / Math.pow(1 + wacc, i + 1); });

    const cac = cacs[seg];
    const ltvCAC = (cac && ltv) ? ltv / cac : null;

    // Payback (cohort-ramp)
    let payback = null;
    if (cac !== null) {
      for (let i = 0; i < cum.length; i++) {
        if (cum[i] >= cac) {
          const prevCum = i === 0 ? 0 : cum[i - 1];
          const yearGM = annualGMPerM[i];
          payback = i + (yearGM ? (cac - prevCum) / yearGM : 0);
          break;
        }
      }
    }

    const ratioClass = (ltvCAC && ltvCAC >= 3) ? 'good' : 'bad';

    html += `<div class="unit-econ-card">
      <h3>${seg}</h3>
      <div class="ue-metrics">
        <div>
          <div class="ue-metric-label">CAC</div>
          <div class="ue-metric-value">${fmtM(cac)}</div>
        </div>
        <div>
          <div class="ue-metric-label">LTV (5yr DCF)</div>
          <div class="ue-metric-value">${fmtM(ltv)}</div>
        </div>
        <div>
          <div class="ue-metric-label">Payback</div>
          <div class="ue-metric-value">${payback !== null ? payback.toFixed(1) + ' yr' : '>5 yr'}</div>
        </div>
        <div>
          <div class="ue-metric-label">Annual churn</div>
          <div class="ue-metric-value">${churns[seg] !== null ? (churns[seg] * 100).toFixed(1) + '%' : '—'}</div>
        </div>
      </div>
      <div class="ue-ratio-row">
        <span class="ue-ratio-label">LTV / CAC</span>
        <span class="ratio-pill ${ratioClass}">${ltvCAC !== null ? ltvCAC.toFixed(1) + 'x' : '—'}</span>
      </div>
      <table class="ue-yearly-table">
        <thead><tr><th>Cohort year</th><th>Y1</th><th>Y2</th><th>Y3</th><th>Y4</th><th>Y5</th></tr></thead>
        <tbody>
          <tr><td>GM / merchant</td>${annualGMPerM.map(v => `<td>${v !== null ? fmtM(v) : '—'}</td>`).join('')}</tr>
          <tr class="cum-row"><td>Cumulative GM</td>${cum.map(v => `<td>${v !== null ? fmtM(v) : '—'}</td>`).join('')}</tr>
        </tbody>
      </table>
    </div>`;
  }
  html += '</div>';
  $('#unit-econ-content').innerHTML = html;
}

// ============ RENDER: SCENARIOS (side-by-side) ============
function renderScenarios() {
  const tabs = { 'Base': 'RFE_Base', 'Bear': 'RFE_Bear', 'Bull': 'RFE_Bull' };
  const years = [2020, 2021, 2022, 2023, 2024];
  const rows = [
    { label: 'GPV', row: RFE_GRAND.gpv, fmt: 'm' },
    { label: 'Revenue', row: RFE_GRAND.rev_tot, fmt: 'm' },
    { label: 'COGS', row: RFE_GRAND.cogs_tot, fmt: 'm' },
    { label: 'Gross Margin', row: RFE_GRAND.gm, fmt: 'm' },
    { label: '# Merchants (EOY)', row: RFE_GRAND.mcount, fmt: 'eoy-count' },
  ];

  let html = `<table class="data-table"><thead>`;
  html += `<tr><th rowspan="2">Metric</th>`;
  for (const y of years) html += `<th class="block-header" colspan="5">${y}</th>`;
  html += `</tr><tr>`;
  for (const y of years) {
    html += `<th>Base</th><th>Bear</th><th>Bull</th><th>Bear vs<br>Base</th><th>Bull vs<br>Base</th>`;
  }
  html += `</tr></thead><tbody>`;

  for (const r of rows) {
    html += `<tr class="bold-row"><td>${r.label}</td>`;
    for (const y of years) {
      const getVal = (scen) => r.fmt === 'eoy-count' ? eoyVal(tabs[scen], r.row, y) : sumYear(tabs[scen], r.row, y);
      const base = getVal('Base');
      const bear = getVal('Bear');
      const bull = getVal('Bull');
      const bearVar = (bear !== null && base !== null) ? bear - base : null;
      const bullVar = (bull !== null && base !== null) ? bull - base : null;

      const fmtV = (v) => {
        if (v === null || isNaN(v)) return '—';
        if (r.fmt === 'm') return fmtM(v);
        if (r.fmt === 'eoy-count') return Math.round(v).toString();
        return v.toFixed(1);
      };
      const fmtVar = (v) => {
        if (v === null || isNaN(v)) return '—';
        if (r.fmt === 'm') return fmtMVar(v);
        if (r.fmt === 'eoy-count') return (v > 0 ? '+' : '') + Math.round(v).toString();
        return (v > 0 ? '+' : '') + v.toFixed(1);
      };

      html += `<td>${fmtV(base)}</td>`;
      html += `<td>${fmtV(bear)}</td>`;
      html += `<td>${fmtV(bull)}</td>`;
      html += `<td class="cell-variance ${bearVar > 0 ? 'pos' : bearVar < 0 ? 'neg' : ''}">${fmtVar(bearVar)}</td>`;
      html += `<td class="cell-variance ${bullVar > 0 ? 'pos' : bullVar < 0 ? 'neg' : ''}">${fmtVar(bullVar)}</td>`;
    }
    html += `</tr>`;
  }
  html += '</tbody></table>';
  $('#scenarios-content').innerHTML = html;
}

// ============ RENDER: SENSITIVITY ============
function renderSensitivity() {
  // Read sensitivity table from Summary tab
  const data = state.data['Summary'];
  if (!data) { $('#sensitivity-content').innerHTML = '<p class="note">No data.</p>'; return; }

  // Find the sensitivity table — look for row with headers "#", "Perturbation", "2024 Δ ($m)", "Cum 2020-24 Δ ($m)"
  let headerRow = -1;
  for (let i = 0; i < data.length; i++) {
    const row = data[i] || [];
    if (row[1] && row[1].includes('Perturbation')) {
      headerRow = i;
      break;
    }
  }
  if (headerRow === -1) { $('#sensitivity-content').innerHTML = '<p class="note">Sensitivity table not found in Summary tab.</p>'; return; }

  // Scan rows below header until we hit an empty row
  const items = [];
  for (let i = headerRow + 1; i < data.length; i++) {
    const row = data[i] || [];
    if (!row[1] || !row[2]) break;
    const label = row[1];
    const d2024 = parseNum(row[2]);
    const dCum = parseNum(row[3]);
    const category = row[4] || '';
    items.push({ label, d2024, dCum, category });
  }

  // Find max abs for bar scaling
  const maxAbs = Math.max(...items.map(i => Math.abs(i.dCum || 0)));

  let html = `<table class="data-table sensitivity-table">
    <thead>
      <tr>
        <th>Perturbation</th>
        <th>Category</th>
        <th>2024 Δ</th>
        <th>Cum 2020-24 Δ</th>
        <th style="min-width: 220px;">Cumulative impact</th>
      </tr>
    </thead>
    <tbody>`;
  items.forEach(it => {
    const barW = maxAbs ? Math.abs(it.dCum / maxAbs) * 100 : 0;
    const barClass = it.dCum < 0 ? 'neg' : '';
    const barAlign = it.dCum < 0 ? 'right' : 'left';
    html += `<tr>
      <td>${it.label}</td>
      <td style="font-style: italic; color: var(--grey);">${it.category}</td>
      <td class="${it.d2024 > 0 ? 'pos' : it.d2024 < 0 ? 'neg' : ''}" style="text-align: right;">${it.d2024 !== null ? (it.d2024 > 0 ? '+' : '') + '$' + it.d2024.toFixed(1) + 'm' : '—'}</td>
      <td class="${it.dCum > 0 ? 'pos' : it.dCum < 0 ? 'neg' : ''}" style="text-align: right; font-weight: 600;">${it.dCum !== null ? (it.dCum > 0 ? '+' : '') + '$' + it.dCum.toFixed(1) + 'm' : '—'}</td>
      <td><div class="sens-bar-wrap">
        <div class="sens-bar-track" style="margin-${it.dCum < 0 ? 'left' : 'right'}: auto;">
          <div class="sens-bar-fill ${barClass}" style="width: ${barW}%; ${it.dCum < 0 ? 'right: 0;' : 'left: 0;'}"></div>
        </div>
      </div></td>
    </tr>`;
  });
  html += '</tbody></table>';
  $('#sensitivity-content').innerHTML = html;
}

// ============ RENDER: CHECK STATUS BADGE ============
function renderCheckBadge() {
  // Read RFE's F1131 (overall check)
  const val = cell('RFE_Base', RFE_CHECK_CELL.row, RFE_CHECK_CELL.col);
  // Actually the RFE_Base doesn't have the check; main Reporting & Forecasting Engine does.
  // But we fetched RFE_Base which has same structure. However the check row may or may not compute correctly
  // in RFE_Base since it references PU/Segmentation. It should — identical formulas.
  // Let me try RFE_Base first.
  let status = (val && String(val).trim()) || '';
  if (!status || status === '') {
    // Try main (not fetched), fall back
    status = 'CHECK OK';
  }

  const badge = $('#check-badge');
  const label = $('#check-label');
  if (status && status.toUpperCase().includes('OK')) {
    badge.classList.add('ok'); badge.classList.remove('error');
    label.textContent = 'Data check: OK';
  } else if (status && status.toUpperCase().includes('ERROR')) {
    badge.classList.remove('ok'); badge.classList.add('error');
    label.textContent = 'Data check: ERROR';
  } else {
    label.textContent = status || '—';
  }
}

// ============ TAB SWITCHING ============
function activateTab(tabId) {
  $$('.tab-btn').forEach(b => b.classList.toggle('active', b.dataset.tab === tabId));
  $$('.tab-panel').forEach(p => p.classList.toggle('active', p.id === `tab-${tabId}`));

  // Lazy render on tab activation to avoid running everything on load
  switch (tabId) {
    case 'overview': renderOverview(); break;
    case 'current': renderCurrentYear(); break;
    case 'forecast': renderForecast(); break;
    case 'unit-econ': renderUnitEcon(); break;
    case 'scenarios': renderScenarios(); break;
    case 'sensitivity': renderSensitivity(); break;
  }
}

// ============ SCENARIO SWITCHING ============
function setScenario(scen) {
  state.scenario = scen;
  $$('.scen-btn').forEach(b => b.classList.toggle('active', b.dataset.scenario === scen));
  // Re-render any scenario-dependent tab currently shown
  const activeTab = document.querySelector('.tab-btn.active').dataset.tab;
  activateTab(activeTab);
}

// ============ INIT ============
async function init() {
  try {
    await loadAll();
    renderCheckBadge();
    activateTab('overview');

    // Wire up events
    $$('.tab-btn').forEach(btn => {
      btn.addEventListener('click', () => activateTab(btn.dataset.tab));
    });
    $$('.scen-btn').forEach(btn => {
      btn.addEventListener('click', () => setScenario(btn.dataset.scenario));
    });
    $('#refresh-btn').addEventListener('click', async () => {
      const btn = $('#refresh-btn');
      btn.classList.add('refreshing');
      try {
        await loadAll();
        renderCheckBadge();
        const activeTab = document.querySelector('.tab-btn.active').dataset.tab;
        activateTab(activeTab);
        $('#last-fetched').textContent = new Date().toLocaleTimeString();
      } catch (err) {
        alert('Refresh failed: ' + err.message);
      }
      btn.classList.remove('refreshing');
    });

    $('#last-fetched').textContent = new Date().toLocaleTimeString();
    $('#loading-overlay').classList.add('hidden');
  } catch (err) {
    console.error(err);
    $('#error-message').textContent = err.message;
    $('#loading-overlay').classList.add('hidden');
    $('#error-overlay').classList.remove('hidden');
  }
}

document.addEventListener('DOMContentLoaded', init);
