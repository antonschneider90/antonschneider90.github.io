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
  // Clear RFE anchor cache — recomputed on next access
  for (const k of Object.keys(_rfeAnchors)) delete _rfeAnchors[k];
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

// Find first row in `tab` where:
//   col A (1) matches colAText AND col B (2) matches colBText (if provided) AND col E (5) matches colEText (if provided)
// Returns 1-indexed row number, or null
function findRow(tab, colAText, colBText, colEText) {
  const data = state.data[tab];
  if (!data) return null;
  for (let i = 0; i < data.length; i++) {
    const row = data[i] || [];
    const a = (row[0] || '').toString().trim();
    const b = (row[1] || '').toString().trim();
    const e = (row[4] || '').toString().trim();
    if (colAText !== null && a !== colAText) continue;
    if (colBText !== null && b !== colBText) continue;
    if (colEText !== null && e !== colEText) continue;
    return i + 1;  // 1-indexed
  }
  return null;
}

// Find the column offset for dates. Google gviz CSV may shift column positions too.
// Look at the row with col A = "Segment" (header row with dates), find the col where
// the date starts (matches 2018-01 or contains "2018")
function findDateStartCol(tab) {
  const data = state.data[tab];
  if (!data) return 6;  // default fallback: col F
  for (let i = 0; i < Math.min(10, data.length); i++) {
    const row = data[i] || [];
    for (let j = 0; j < row.length; j++) {
      const v = (row[j] || '').toString();
      if (/^2018-01/.test(v) || /^1\/1\/2018/.test(v) || v === '2018-01' || /Jan.*2018/i.test(v)) {
        return j + 1;  // 1-indexed
      }
    }
  }
  return 6;  // default fallback
}

// Cache-aware helpers for each RFE tab
const _rfeAnchors = {};
function rfeAnchors(tab) {
  if (_rfeAnchors[tab]) return _rfeAnchors[tab];
  // Find all anchor rows dynamically
  const grand = {};
  const metricMap = {
    mcount: '# Merchants total',
    gpv: 'GPV ($)',
    events: 'Events (#)',
    rev_vol: 'Revenue (vol)',
    rev_trx: 'Revenue (trx)',
    rev_add: 'Revenue (add)',
    rev_tot: 'Revenue Total',
    cogs_int: 'COGS inter/network',
    cogs_api: 'COGS API',
    cogs_loss: 'COGS loss',
    cogs_tot: 'COGS Total',
    gm: 'Gross Margin',
  };
  for (const [k, label] of Object.entries(metricMap)) {
    grand[k] = findRow(tab, 'TOTAL', 'Grand Total', label);
  }

  const segments = {};
  const segMetricMap = {
    mcount: '# Merchants total',
    gpv: 'GPV ($) total',
    events: 'Events (#) total',
    rev_tot: 'Revenue Total total',
    cogs_tot: 'COGS Total total',
    gm: 'Gross Margin',  // segment GM row label is just "Gross Margin"
  };
  for (const seg of ['SaaS', 'E-Commerce Store', 'Platform']) {
    segments[seg] = {};
    for (const [k, label] of Object.entries(segMetricMap)) {
      segments[seg][k] = findRow(tab, seg, 'SEGMENT TOTAL', label);
    }
  }

  // Cohort 2020 rows — find by col A = segment, col B = "New 2020", col E = metric
  const cohort2020 = {};
  for (const seg of ['SaaS', 'E-Commerce Store', 'Platform']) {
    cohort2020[seg] = {
      gm_row: findRow(tab, seg, 'New 2020', 'Gross Margin'),
      new_row: findRow(tab, seg, 'New 2020', '+ New merchants'),
    };
  }

  // Check cell (CHECK OK / CHECK ERROR) — find the row where col A contains "OVERALL" and grab col F
  // Actually our check row has col A = "OVERALL" and the result in col F
  let checkRow = null;
  for (let i = 0; i < (state.data[tab] || []).length; i++) {
    const row = state.data[tab][i] || [];
    const a = (row[0] || '').toString();
    if (a.includes('OVERALL')) { checkRow = i + 1; break; }
  }

  const dateStartCol = findDateStartCol(tab);

  _rfeAnchors[tab] = { grand, segments, cohort2020, checkRow, dateStartCol };
  return _rfeAnchors[tab];
}

// Excel column letter to 1-indexed number
function colToNum(col) {
  let n = 0;
  for (let i = 0; i < col.length; i++) {
    n = n * 26 + (col.charCodeAt(i) - 64);
  }
  return n;
}

// Year → column range in an RFE tab (uses dynamic date-start column)
function yearColRange(tab, year) {
  const dateStart = rfeAnchors(tab).dateStartCol;
  const startCol = dateStart + (year - 2018) * 12;
  return [startCol, startCol + 11];
}
// Sum a range of cols for a specific row
function sumRow(tab, row, startCol, endCol) {
  if (!row) return null;
  let sum = 0;
  let any = false;
  for (let c = startCol; c <= endCol; c++) {
    const v = cellNum(tab, row, c);
    if (v !== null && !isNaN(v)) { sum += v; any = true; }
  }
  return any ? sum : null;
}
function sumYear(tab, row, year) {
  if (!row) return null;
  const [s, e] = yearColRange(tab, year);
  return sumRow(tab, row, s, e);
}
// EOY value = last (Dec) column of year
function eoyVal(tab, row, year) {
  if (!row) return null;
  const [, e] = yearColRange(tab, year);
  return cellNum(tab, row, e);
}

function rfeTabForScenario(scen) {
  return `RFE_${scen}`;
}

// Find active scenario text in Assumptions tab — look for cell containing "Base", "Bear", or "Bull"
// near a row labeled "Active scenario" or similar
function findAssumptionScenario() {
  const data = state.data['Assumptions'];
  if (!data) return null;
  for (let i = 0; i < data.length; i++) {
    const row = data[i] || [];
    const a = (row[0] || '').toString().toLowerCase();
    if (a.includes('active scenario') || a.includes('scenario switch') || a.includes('scenario')) {
      // Check adjacent cells for Base/Bear/Bull
      for (let j = 1; j < row.length; j++) {
        const v = (row[j] || '').toString().trim();
        if (v === 'Base' || v === 'Bear' || v === 'Bull') return v;
      }
    }
  }
  return null;
}

// Find CAC for a segment in Assumptions tab — look for row with label like "CAC — SaaS" or "CAC SaaS"
function findAssumptionCAC(seg) {
  const data = state.data['Assumptions'];
  if (!data) return null;
  const segKey = seg.toLowerCase();
  for (let i = 0; i < data.length; i++) {
    const row = data[i] || [];
    const a = (row[0] || '').toString().toLowerCase();
    if (a.includes('cac') && (
        (segKey === 'saas' && a.includes('saas')) ||
        (segKey === 'e-commerce store' && (a.includes('e-commerce') || a.includes('ecommerce') || a.includes('e-com'))) ||
        (segKey === 'platform' && a.includes('platform'))
      )) {
      // Return the numeric value in the next few cells
      for (let j = 1; j < row.length; j++) {
        const n = parseNum(row[j]);
        if (n !== null && n > 0) return n;
      }
    }
  }
  return null;
}

// Find WACC in Assumptions
function findAssumptionWACC() {
  const data = state.data['Assumptions'];
  if (!data) return null;
  for (let i = 0; i < data.length; i++) {
    const row = data[i] || [];
    const a = (row[0] || '').toString().toLowerCase();
    if (a.includes('wacc') || a.includes('discount rate') || a.includes('cost of capital')) {
      for (let j = 1; j < row.length; j++) {
        const n = parseNum(row[j]);
        if (n !== null && n > 0 && n < 1) return n;  // WACC should be between 0 and 1
      }
    }
  }
  return 0.10;  // fallback
}

// Find segment churn for active scenario in Assumptions
function findAssumptionChurn(seg, scenario) {
  const data = state.data['Assumptions'];
  if (!data) return null;
  const segKey = seg.toLowerCase();
  // Look for "Churn" related rows per segment
  for (let i = 0; i < data.length; i++) {
    const row = data[i] || [];
    const a = (row[0] || '').toString().toLowerCase();
    if (a.includes('churn') && (
        (segKey === 'saas' && a.includes('saas')) ||
        (segKey === 'e-commerce store' && (a.includes('e-commerce') || a.includes('ecommerce'))) ||
        (segKey === 'platform' && a.includes('platform'))
      )) {
      // Find the column that corresponds to the scenario
      // Structure: col 3 = base case comment, col 4 = base, col 5 = bear, col 6 = bull (1-indexed)
      // But this may shift. Just look at first 3 numeric values as Base/Bear/Bull in order.
      const nums = [];
      for (let j = 1; j < row.length; j++) {
        const n = parseNum(row[j]);
        if (n !== null && n >= 0 && n <= 1) nums.push(n);
      }
      // Typically: base, bear, bull in that column order
      const scenIdx = { 'Base': 0, 'Bear': 1, 'Bull': 2 }[scenario] || 0;
      return nums[scenIdx] !== undefined ? nums[scenIdx] : nums[0];
    }
  }
  return null;
}

// Find Last Actual Month from CYO tab — look for the row with label "Last month of actuals"
function findCYOLastActualMonth() {
  const data = state.data['Current Year Overview'];
  if (!data) return null;
  for (let i = 0; i < data.length; i++) {
    const row = data[i] || [];
    const a = (row[0] || '').toString().toLowerCase();
    if (a.includes('last month of actuals') || a.includes('last actual')) {
      for (let j = 1; j < row.length; j++) {
        const v = (row[j] || '').toString().trim();
        if (v && v !== '') return v;
      }
    }
  }
  return null;
}

// ============ RENDER: OVERVIEW ============
function renderOverview() {
  const tab = rfeTabForScenario(state.scenario);
  const a = rfeAnchors(tab);
  // KPIs — 2024 full-year revenue, GM, merchants EOY, GPV
  const rev2024 = sumYear(tab, a.grand.rev_tot, 2024);
  const gm2024 = sumYear(tab, a.grand.gm, 2024);
  const merch2024 = eoyVal(tab, a.grand.mcount, 2024);
  const gpv2024 = sumYear(tab, a.grand.gpv, 2024);
  // Baseline (2019 actuals, scenario-independent)
  const rev2019 = sumYear(tab, a.grand.rev_tot, 2019);
  const gm2019 = sumYear(tab, a.grand.gm, 2019);
  const merch2019 = eoyVal(tab, a.grand.mcount, 2019);
  const gpv2019 = sumYear(tab, a.grand.gpv, 2019);

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
  const activeSheetScen = findAssumptionScenario() || '—';
  const cyo = state.data['Current Year Overview'];
  const lastActual = findCYOLastActualMonth() || '—';
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
  const a = rfeAnchors(tab);
  const years = [2018, 2019, 2020, 2021, 2022, 2023, 2024];
  const revs = years.map(y => (sumYear(tab, a.grand.rev_tot, y) || 0) / 1e6);
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
  const a = rfeAnchors(tab);
  const segs = ['SaaS', 'E-Commerce Store', 'Platform'];
  const vals = segs.map(s => (sumYear(tab, a.segments[s].rev_tot, 2024) || 0) / 1e6);
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

  const activeSheetScen = findAssumptionScenario() || '—';
  const lastActual = findCYOLastActualMonth() || '—';
  $('#current-meta').textContent = `Last actual month: ${lastActual} · Active Sheet scenario: ${activeSheetScen}`;

  if (activeSheetScen !== state.scenario && activeSheetScen !== '—') {
    $('#current-scen-note').textContent = `Note: this view shows ${activeSheetScen} (the active scenario in the Sheet). The dashboard toggle (${state.scenario}) does not affect this tab because the CYO's Last Month / YTD / RoY math depends on the active scenario cell in the Sheet. To change this view, edit Assumptions!C7 in the Sheet and refresh.`;
  } else if (activeSheetScen !== '—') {
    $('#current-scen-note').textContent = `Sheet's active scenario matches dashboard toggle: ${state.scenario}`;
  }

  // Build table by scanning CYO rows sequentially and rendering them as-is.
  // The CYO is already a nicely-formatted table; we just render its content.
  // Find the start: look for the first row that has meaningful content (skip title, b2, etc)
  // We render from the row that looks like a section header downwards.

  let html = '<table class="data-table">';

  // Find first meaningful row — scan for row with "FINANCIALS" in col A (uppercase)
  let startIdx = -1;
  for (let i = 0; i < data.length; i++) {
    const row = data[i] || [];
    const a = (row[0] || '').toString().toUpperCase();
    if (a.includes('FINANCIALS') || a.includes('KPI')) { startIdx = i; break; }
    // Also check for a row that has block headers like "Last Month"
    if (row.some(c => c && /^Last Month$|^YTD/.test(String(c).trim()))) {
      // This is the block header row — start 2 rows earlier to pick up any preceding labels
      startIdx = Math.max(0, i - 1);
      break;
    }
  }
  if (startIdx === -1) startIdx = 5;  // fallback

  for (let i = startIdx; i < data.length; i++) {
    const row = data[i] || [];
    if (row.every(c => !c || String(c).trim() === '')) continue;  // skip empty rows

    const label = (row[0] || '').toString().trim();
    const rowText = row.slice(1).join(' ');
    const isSection = /FINANCIALS|KPI/.test(label.toUpperCase()) && label.length > 4;
    const isSubHeader = /^—.*—$/.test(label);
    const isBlockHeader = /Last Month|YTD.*Act|Rest of Year|Full Year/.test(rowText);
    const isSubHeader2 = /Current Year|Prior Year|YoY/.test(rowText) && (!label || label === '');

    if (isSection) {
      html += `<tr class="section-row"><td colspan="25">${label}</td></tr>`;
      continue;
    }
    if (isSubHeader) {
      html += `<tr class="sub-header-row"><td colspan="25">${label}</td></tr>`;
      continue;
    }
    if (isBlockHeader) {
      html += '<tr><th></th>';
      html += '<th class="block-header" colspan="4">Last Month</th><th></th>';
      html += '<th class="block-header" colspan="4">YTD (Act)</th><th></th>';
      html += '<th class="block-header" colspan="4">Rest of Year (Fcst)</th><th></th>';
      html += '<th class="block-header" colspan="4">Full Year</th>';
      html += '</tr>';
      continue;
    }
    if (isSubHeader2) {
      html += '<tr><th></th>';
      for (let blk = 0; blk < 4; blk++) {
        html += '<th>Current<br>Year</th><th>Prior<br>Year</th><th>YoY ($)</th><th>YoY (%)</th>';
        if (blk < 3) html += '<th></th>';
      }
      html += '</tr>';
      continue;
    }

    // Data row — just emit all cells as-is
    const rowClass = /^(Merchants|Revenue per|GPV per|New merchants|Churned|Net Take|Gross Margin %|Churn rate|  )/.test(label) ? '' : 'bold-row';
    html += `<tr class="${rowClass}">`;
    const indent = label.startsWith('  ') || label.startsWith('\t');
    html += `<td${indent ? ' class="indent"' : ''}>${label.trim()}</td>`;

    // Render remaining cells. The 4 blocks of 4 columns each, separated by gap columns.
    // Layout: col 1 = label. Then cols 2-5 = Last Month block (Cur/Prior/YoY$/YoY%). Col 6 = gap. Cols 7-10 = YTD. etc.
    // We'll render 19 data cells following the label.
    for (let j = 1; j <= 19; j++) {
      const v = row[j] !== undefined ? row[j] : '';
      const inBlock = [1, 2, 3, 4, 6, 7, 8, 9, 11, 12, 13, 14, 16, 17, 18, 19].includes(j);
      if (!inBlock) { html += '<td></td>'; continue; }
      const colInBlock = ((j - 1) % 5);
      let cls = '';
      if (colInBlock === 0) cls = 'cell-current';
      else if (colInBlock === 2 || colInBlock === 3) cls = 'cell-variance';
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
  const anch = rfeAnchors(tab);
  const years = [2018, 2019, 2020, 2021, 2022, 2023, 2024];
  $('#fc-scen-inline').textContent = state.scenario;

  const rows = [
    { sec: 'FINANCIALS — GRAND TOTAL', type: 'section' },
    { label: 'GPV', row: anch.grand.gpv, fmt: 'm', bold: true },
    { label: 'Revenue', row: anch.grand.rev_tot, fmt: 'm', bold: true },
    { label: 'Net Take Rate %', num: anch.grand.rev_tot, den: anch.grand.gpv, fmt: 'pct', indent: true },
    { label: 'COGS', row: anch.grand.cogs_tot, fmt: 'm', bold: true },
    { label: 'Gross Margin', row: anch.grand.gm, fmt: 'm', bold: true },
    { label: 'Gross Margin %', num: anch.grand.gm, den: anch.grand.rev_tot, fmt: 'pct', indent: true },
    { sec: 'FINANCIALS — BY SEGMENT', type: 'section' },
  ];

  for (const seg of ['SaaS', 'E-Commerce Store', 'Platform']) {
    rows.push({ sec: `— ${seg} —`, type: 'subheader' });
    const segR = anch.segments[seg];
    rows.push({ label: 'GPV', row: segR.gpv, fmt: 'm', bold: true, indent: true });
    rows.push({ label: 'Revenue', row: segR.rev_tot, fmt: 'm', bold: true, indent: true });
    rows.push({ label: 'Net Take Rate %', num: segR.rev_tot, den: segR.gpv, fmt: 'pct', indent: true });
    rows.push({ label: 'COGS', row: segR.cogs_tot, fmt: 'm', bold: true, indent: true });
    rows.push({ label: 'Gross Margin', row: segR.gm, fmt: 'm', bold: true, indent: true });
    rows.push({ label: 'Gross Margin %', num: segR.gm, den: segR.rev_tot, fmt: 'pct', indent: true });
  }
  rows.push({ sec: 'KPIs', type: 'section' });
  rows.push({ label: '# Merchants (EOY)', row: anch.grand.mcount, fmt: 'eoy-count', bold: true });

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
  const anch = rfeAnchors(tab);
  $('#ue-scen-inline').textContent = state.scenario;

  // Read CAC and WACC from Assumptions via label-based lookups
  const cacs = {
    'SaaS': findAssumptionCAC('SaaS'),
    'E-Commerce Store': findAssumptionCAC('E-Commerce Store'),
    'Platform': findAssumptionCAC('Platform'),
  };
  const wacc = findAssumptionWACC() || 0.10;
  const churns = {
    'SaaS': findAssumptionChurn('SaaS', state.scenario),
    'E-Commerce Store': findAssumptionChurn('E-Commerce Store', state.scenario),
    'Platform': findAssumptionChurn('Platform', state.scenario),
  };

  let html = '<div class="unit-econ-grid">';
  for (const seg of ['SaaS', 'E-Commerce Store', 'Platform']) {
    const c2020 = anch.cohort2020[seg];
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
    if (cac !== null && cac !== undefined) {
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
          <div class="ue-metric-value">${churns[seg] !== null && churns[seg] !== undefined ? (churns[seg] * 100).toFixed(1) + '%' : '—'}</div>
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
  // Get anchors per scenario tab — they're structurally identical but let's be safe
  const anchors = {
    'Base': rfeAnchors('RFE_Base'),
    'Bear': rfeAnchors('RFE_Bear'),
    'Bull': rfeAnchors('RFE_Bull'),
  };
  const rows = [
    { label: 'GPV', key: 'gpv', fmt: 'm' },
    { label: 'Revenue', key: 'rev_tot', fmt: 'm' },
    { label: 'COGS', key: 'cogs_tot', fmt: 'm' },
    { label: 'Gross Margin', key: 'gm', fmt: 'm' },
    { label: '# Merchants (EOY)', key: 'mcount', fmt: 'eoy-count' },
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
      const getVal = (scen) => {
        const rowNum = anchors[scen].grand[r.key];
        return r.fmt === 'eoy-count' ? eoyVal(tabs[scen], rowNum, y) : sumYear(tabs[scen], rowNum, y);
      };
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
  // Scan entire RFE_Base for the cell that contains "CHECK OK" or "CHECK ERROR"
  // This is the most robust — no dependency on row numbers
  const data = state.data['RFE_Base'];
  let status = null;
  if (data) {
    for (let i = data.length - 1; i >= 0; i--) {  // scan from bottom, check is near the end
      const row = data[i] || [];
      for (const v of row) {
        const s = (v || '').toString().trim();
        if (s === 'CHECK OK' || s === 'CHECK ERROR') {
          status = s;
          break;
        }
      }
      if (status) break;
    }
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
    label.textContent = status || 'Data check: —';
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
