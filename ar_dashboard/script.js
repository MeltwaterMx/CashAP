// ────────────────────────────────────────────────────────────────
//  AR ANALYTICS DASHBOARD – script.js
// ────────────────────────────────────────────────────────────────

// ── GLOBAL DATA & PERSISTENCE ──────────────────────────────────
const STORAGE_KEY = 'ar_dashboard_data';
const SCHEMA_VERSION = 6; // Increment this when DASHBOARD_DATA structure changes
let currentCaFilter = 'all'; // Filter for Cash App items
let activeTab = 'overview'; // Track currently active tab globally
let currentLang = 'es';    // Language toggle state ('es' | 'en')

function saveState() {
  const toSave = JSON.parse(JSON.stringify(DATA));
  toSave.__version = SCHEMA_VERSION;
  localStorage.setItem(STORAGE_KEY, JSON.stringify(toSave));
}

function loadState() {
  const saved = localStorage.getItem(STORAGE_KEY);
  if (!saved) return;
  try {
    const parsed = JSON.parse(saved);
    // ── SCHEMA VERSION CHECK ──────────────────────────────────────
    // If saved version doesn't match current schema, discard stale data
    // to prevent crashes from structural mismatches.
    if (!parsed.__version || parsed.__version !== SCHEMA_VERSION) {
      console.warn('[AR Dashboard] Schema version mismatch. Clearing localStorage and loading defaults.');
      localStorage.removeItem(STORAGE_KEY);
      return;
    }
    delete parsed.__version;
    // Deep merge loaded state into DATA
    Object.assign(DATA, parsed);
  } catch (err) {
    console.error('[AR Dashboard] Error loading saved state, clearing:', err);
    localStorage.removeItem(STORAGE_KEY);
  }
}

function resyncData() {
  // 1. Recalculate Total AR from Aging buckets
  const ag = DATA.aging;
  DATA.totalAR = ag.current + ag.d30 + ag.d60 + ag.d90 + ag.d90p;

  // 2. Sync Executive Summary KPIs
  const kpiTotal = document.getElementById('kpi-total');
  if (kpiTotal) kpiTotal.textContent = fmt(DATA.totalAR);

  const kpiDso = document.getElementById('kpi-dso');
  if (kpiDso) kpiDso.textContent = DATA.dso.actual + ' días';

  const dsoDiff = (DATA.dso.actual - DATA.dso.target).toFixed(1);
  const dsoDelta = document.getElementById('kpi-dso-delta');
  if (dsoDelta) {
    if (+dsoDiff > 0) {
      dsoDelta.textContent = `↑ ${dsoDiff} días vs objetivo`;
      dsoDelta.className = 'kpi-delta negative';
    } else {
      dsoDelta.textContent = `↓ ${Math.abs(dsoDiff)} días vs objetivo`;
      dsoDelta.className = 'kpi-delta positive';
    }
  }

  // 3. Sync Cash App KPIs strictly from item sums
  const ca = DATA.cashapp;
  if (ca.items) {
    // Unapplied = Total bucket
    ca.kpis.unapplied = ca.items.reduce((s, i) => s + (Number(i.amount) || 0), 0);

    // Suspense = Pending or Unknown (client with ?)
    ca.kpis.suspense = ca.items
      .filter(i => i.status === 'Pendiente' || (i.client && i.client.includes('?')))
      .reduce((s, i) => s + (Number(i.amount) || 0), 0);
  }

  // 4. Update all UI elements that might be visible
  refreshAllUI();
  saveState();
}

function refreshAllUI() {
  const ca = DATA.cashapp;
  const elements = {
    'kpi-total': fmt(DATA.totalAR),
    'kpi-collected': fmt(DATA.collected || Math.round(DATA.totalAR * 0.45)),
    'kpi-risk': DATA.clients.filter(c => c.score >= 70).length,
    'ca-unapplied': fmt(ca.kpis.unapplied),
    'ca-suspense': fmt(ca.kpis.suspense),
    'ca-automatch': ca.kpis.autoMatch + '%',
    'ca-refunds': fmt(ca.refunds.total)
  };

  Object.entries(elements).forEach(([id, val]) => {
    const el = document.getElementById(id);
    if (el) el.textContent = val;
  });

  updateDsoGauges();

  // Refresh active tab charts and tables
  const activeBtn = document.querySelector('.nav-btn.active');
  if (activeBtn) initCharts(activeBtn.dataset.tab);
}

function resetData() {
  if (confirm('¿Restablecer datos originales del archivo data.js? Perderás los cambios no guardados en el archivo.')) {
    localStorage.removeItem(STORAGE_KEY);
    location.reload();
  }
}

// Inicializar con DASHBOARD_DATA (de data.js) para respetar ediciones manuales
const DATA = JSON.parse(JSON.stringify(DASHBOARD_DATA));

// ── CHART INSTANCES ────────────────────────────────────────────
const charts = {};

// ── TAB SWITCHER ────────────────────────────────────────────────
const titles = {
  overview: 'Resumen Ejecutivo', dso: 'Análisis DSO', aging: 'Aging Report',
  risk: 'Análisis de Riesgo', segmentation: 'Segmentación de Cartera', projection: 'Proyección de Recaudos',
  cashapp: 'Cash Applications', refunds: 'Refunds'
};
function switchTab(id, btn) {
  document.querySelectorAll('.tab-content').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));
  document.getElementById('tab-' + id).classList.add('active');
  btn.classList.add('active');
  activeTab = id; // Keep global activeTab in sync
  document.getElementById('page-title').textContent = titles[id];
  // lazy-init charts
  setTimeout(() => { initCharts(id); }, 50);
}

// ── HELPERS ────────────────────────────────────────────────────
const fmt = n => {
  const locale = currentLang === 'en' ? 'en-US' : 'es-CR';
  return '$' + n.toLocaleString(locale);
};
const pct = (a, t) => ((a / t) * 100).toFixed(1) + '%';
const COLORS = {
  blue: '#3b82f6',
  purple: '#8b5cf6',
  green: '#10b981',
  orange: '#f59e0b',
  red: '#ef4444',
  yellow: '#eab308',
  bg: '#090A0F',
  surface: '#11131A',
  text: '#ffffff',
  text2: '#94a3b8'
};

function chartDefaults() {
  Chart.defaults.color = COLORS.text2;
  Chart.defaults.borderColor = 'rgba(255, 255, 255, 0.03)';
  Chart.defaults.font.family = 'Inter';
  Chart.defaults.font.size = 11;
  Chart.defaults.scale.grid.color = 'rgba(255, 255, 255, 0.02)';
  Chart.defaults.scale.grid.drawBorder = false;

  // Global Tooltip Styling for better readability on high-res screens
  Chart.defaults.plugins.tooltip.backgroundColor = 'rgba(15, 23, 42, 0.98)';
  Chart.defaults.plugins.tooltip.padding = 16;
  Chart.defaults.plugins.tooltip.titleFont = { size: 16, weight: 'bold', family: 'Inter' };
  Chart.defaults.plugins.tooltip.bodyFont = { size: 15, family: 'Inter' };
  Chart.defaults.plugins.tooltip.footerFont = { size: 14 };
  Chart.defaults.plugins.tooltip.cornerRadius = 10;
  Chart.defaults.plugins.tooltip.boxPadding = 8;
  Chart.defaults.plugins.tooltip.borderColor = 'rgba(255, 255, 255, 0.1)';
  Chart.defaults.plugins.tooltip.borderWidth = 1;
  Chart.defaults.plugins.tooltip.usePointStyle = true;
  
  // Let's add drop shadow via custom plugin for all charts
  Chart.register({
    id: 'glowPlugin',
    beforeDatasetsDraw: function(chart) {
      let ctx = chart.ctx;
      ctx.save();
      if (chart.config.type === 'line') {
        ctx.shadowColor = 'rgba(176, 38, 255, 0.5)';
        ctx.shadowBlur = 12;
        ctx.shadowOffsetX = 0;
        ctx.shadowOffsetY = 4;
      } else if (chart.config.type === 'doughnut') {
        ctx.shadowColor = 'rgba(0, 0, 0, 0.6)';
        ctx.shadowBlur = 14;
        ctx.shadowOffsetX = 0;
        ctx.shadowOffsetY = 6;
      } else if (chart.config.type === 'bar') {
        ctx.shadowColor = 'rgba(0, 0, 0, 0.4)';
        ctx.shadowBlur = 8;
        ctx.shadowOffsetX = 2;
        ctx.shadowOffsetY = 4;
      }
    },
    afterDatasetsDraw: function(chart) {
      chart.ctx.restore();
    }
  });
}

// ── CIRCULAR GAUGE DRAWING ─────────────────────────────────────
// ── PREMIUM CIRCULAR GAUGE DRAWING ────────────────────────────
function drawCircularGauge(canvasId, value, targetVal, color) {
  const canvas = document.getElementById(canvasId);
  if (!canvas) return;
  const ctx = canvas.getContext('2d');
  const w = canvas.width, h = canvas.height;
  const MAX = 60;

  ctx.clearRect(0, 0, w, h);

  const paddingY = 18;
  const lineWidth = 12;

  const cx = w / 2;
  const cy = h * 0.85; 
  const maxR = Math.min(cx - lineWidth, cy - paddingY);
  const r = maxR;

  const start = Math.PI;
  const end = 2 * Math.PI;

  // Background arc (dark gray)
  ctx.beginPath();
  ctx.arc(cx, cy, r, start, end);
  ctx.lineWidth = lineWidth;
  ctx.strokeStyle = 'rgba(255, 255, 255, 0.05)';
  ctx.lineCap = 'round';
  ctx.stroke();

  // Value arc with solid neon color and glow
  const pct = Math.min(value / MAX, 1);
  const valEnd = start + pct * Math.PI;

  ctx.shadowBlur = 15;
  ctx.shadowColor = color;
  
  ctx.beginPath();
  ctx.arc(cx, cy, r, start, valEnd);
  ctx.lineWidth = lineWidth;
  ctx.strokeStyle = color;
  ctx.lineCap = 'round';
  ctx.stroke();

  // Reset shadow
  ctx.shadowBlur = 0;

  // Target marker - single stylish line
  if (targetVal) {
    const tPct = Math.min(targetVal / MAX, 1);
    const ta = start + tPct * Math.PI;
    const innerR = r - Math.max(lineWidth/2 + 4, 10);
    const outerR = r + Math.max(lineWidth/2 + 4, 10);
    const tx1 = cx + innerR * Math.cos(ta), ty1 = cy + innerR * Math.sin(ta);
    const tx2 = cx + outerR * Math.cos(ta), ty2 = cy + outerR * Math.sin(ta);

    ctx.beginPath();
    ctx.moveTo(tx1, ty1);
    ctx.lineTo(tx2, ty2);
    ctx.strokeStyle = '#f1fa8c'; // Bright yellow for target
    ctx.lineWidth = 3;
    ctx.lineCap = 'round';
    ctx.stroke();
  }
}

function updateDsoGauges() {
  const d = DATA.dso;
  drawCircularGauge('gaugeActual', d.actual, d.target, d.actual > d.target ? COLORS.orange : COLORS.blue);
  drawCircularGauge('gaugePrev', d.prev, d.target, COLORS.purple);
  drawCircularGauge('gaugeBest', d.best, d.target, COLORS.green);
  drawCircularGauge('gaugeTarget', d.target, null, COLORS.yellow);

  document.getElementById('gaugeActualVal').textContent = d.actual;
  document.getElementById('gaugePrevVal').textContent = d.prev;
  document.getElementById('gaugeBestVal').textContent = d.best;
  document.getElementById('gaugeTargetVal').textContent = d.target;
}

// ── CHART BUILDERS ─────────────────────────────────────────────
function initCharts(tab) {
  if (tab === 'overview' || tab === '__all') {
    if (!charts.overviewAging) {
      const ag = DATA.aging;
      const total = ag.current + ag.d30 + ag.d60 + ag.d90 + ag.d90p;
      charts.overviewAging = new Chart(document.getElementById('overviewAgingChart'), {
        type: 'bar',
        data: {
          labels: ['Curr', '1–30d', '31–60d', '61–90d', '+90d'],
          datasets: [{
            label: 'Saldo',
            data: [ag.current, ag.d30, ag.d60, ag.d90, ag.d90p],
            backgroundColor: [COLORS.green, COLORS.purple, COLORS.yellow, COLORS.orange, COLORS.red],
            borderRadius: 8,
            borderSkipped: false,
            barThickness: 20
          }]
        },
        options: {
          indexAxis: 'y', responsive: true, maintainAspectRatio: false,
          plugins: { legend: { display: false } },
          layout: { padding: { top: 10, bottom: 10, left: 0, right: 20 } },
          scales: {
            x: { grid: { display: false }, ticks: { display: false } },
            y: { grid: { display: false }, ticks: { font: { size: 10 }, color: COLORS.text2 } }
          }
        }
      });
    }
    if (!charts.riskDonut) {
      const hiRisk = DATA.clients.filter(c => c.score >= 70).reduce((s, c) => s + c.balance, 0);
      const medRisk = DATA.clients.filter(c => c.score >= 40 && c.score < 70).reduce((s, c) => s + c.balance, 0);
      const lowRisk = DATA.clients.filter(c => c.score < 40).reduce((s, c) => s + c.balance, 0);
      charts.riskDonut = new Chart(document.getElementById('riskDonutChart'), {
        type: 'doughnut',
        data: {
          labels: ['Riesgo Alto', 'Riesgo Medio', 'Riesgo Bajo'],
          datasets: [{ data: [hiRisk, medRisk, lowRisk], backgroundColor: [COLORS.red, COLORS.yellow, COLORS.green], hoverOffset: 4, borderWidth: 4, borderColor: 'rgba(0,0,0,0.5)', borderRadius: 4 }]
        },
        options: {
          responsive: true, maintainAspectRatio: false, cutout: '70%',
          layout: { padding: { top: 10, bottom: 10 } },
          plugins: {
            legend: { display: true, position: 'bottom', labels: { boxWidth: 12, padding: 20, font: { size: 11 } } },
            tooltip: { enabled: true }
          }
        }
      });
    }
  }

  if (tab === 'dso') {
    updateDsoGauges();
    if (!charts.dsoTrend) {
      charts.dsoTrend = new Chart(document.getElementById('dsoTrendChart'), {
        type: 'line',
        data: {
          labels: DATA.months,
          datasets: [
            {
              label: 'DSO Real',
              data: DATA.dsoHistory,
              borderColor: COLORS.green,
              backgroundColor: 'rgba(57, 255, 20, 0.15)',
              fill: true,
              tension: 0.5,
              pointRadius: 3,
              pointHoverRadius: 6,
              pointBackgroundColor: COLORS.bg,
              borderWidth: 2
            },
            { label: 'Objetivo', data: Array(6).fill(DATA.dso.target), borderColor: COLORS.yellow, borderDash: [5, 5], pointRadius: 0, fill: false, borderWidth: 2 },
            { label: 'Best DSO', data: Array(6).fill(DATA.dso.best), borderColor: COLORS.green, borderDash: [3, 3], pointRadius: 0, fill: false, borderWidth: 1 }
          ]
        },
        options: {
          responsive: true, maintainAspectRatio: false,
          interaction: { intersect: false, mode: 'index' },
          plugins: {
            legend: { display: true, position: 'top', labels: { boxWidth: 12, usePointStyle: true, font: { size: 11 } } },
            tooltip: { backgroundColor: '#1c2038', titleColor: '#fff', bodyColor: '#8892b0', borderColor: '#252a45', borderWidth: 1 }
          },
          scales: {
            y: { min: 20, max: 45, grid: { color: 'rgba(35, 40, 64, 0.5)' }, ticks: { font: { size: 10 } } },
            x: { grid: { display: false }, ticks: { font: { size: 10 } } }
          }
        }
      });
    }

    if (!charts.dsoComposition) {
      // Simulation of DSO components: Terms vs Delays
      charts.dsoComposition = new Chart(document.getElementById('dsoCompositionChart'), {
        type: 'bar',
        data: {
          labels: DATA.months,
          datasets: [
            { label: 'Términos de Crédito (Base)', data: [25, 25, 25, 25, 25, 25], backgroundColor: 'rgba(168, 85, 247, 0.6)', stack: 'stack0', borderRadius: { bottomLeft: 6, bottomRight: 6, topLeft: 0, topRight: 0 } },
            { label: 'Retraso en Cobro', data: DATA.dsoHistory.map(v => v - 25), backgroundColor: 'rgba(217, 70, 239, 0.6)', stack: 'stack0', borderRadius: { topLeft: 6, topRight: 6, bottomLeft: 0, bottomRight: 0 } }
          ]
        },
        options: {
          responsive: true, maintainAspectRatio: false,
          plugins: {
            legend: { display: true, position: 'top', labels: { boxWidth: 10, font: { size: 10 } } },
            title: { display: false }
          },
          scales: {
            x: { stacked: true, grid: { display: false } },
            y: { stacked: true, min: 0, max: 50, grid: { color: 'rgba(35, 40, 64, 0.5)' } }
          }
        }
      });
    }
  }

  if (tab === 'aging') {
    const ag = DATA.aging;
    const total = ag.current + ag.d30 + ag.d60 + ag.d90 + ag.d90p;
    document.getElementById('aging-current').textContent = fmt(ag.current);
    document.getElementById('aging-30').textContent = fmt(ag.d30);
    document.getElementById('aging-60').textContent = fmt(ag.d60);
    document.getElementById('aging-90').textContent = fmt(ag.d90);
    document.getElementById('aging-90plus').textContent = fmt(ag.d90p);
    document.getElementById('aging-current-pct').textContent = pct(ag.current, total);
    document.getElementById('aging-30-pct').textContent = pct(ag.d30, total);
    document.getElementById('aging-60-pct').textContent = pct(ag.d60, total);
    document.getElementById('aging-90-pct').textContent = pct(ag.d90, total);
    document.getElementById('aging-90plus-pct').textContent = pct(ag.d90p, total);
    if (!charts.agingBar) {
      charts.agingBar = new Chart(document.getElementById('agingBarChart'), {
        type: 'bar',
        data: {
          labels: ['Corriente', '1–30 días', '31–60 días', '61–90 días', '+90 días'],
          datasets: [{
            label: 'Saldo (USD)',
            data: [ag.current, ag.d30, ag.d60, ag.d90, ag.d90p],
            backgroundColor: [COLORS.green, COLORS.blue, COLORS.yellow, COLORS.orange, COLORS.red],
            borderRadius: 8, borderSkipped: false
          }]
        },
        options: {
          indexAxis: 'y', responsive: true, maintainAspectRatio: false,
          layout: { padding: 8 },
          plugins: { legend: { display: false } },
          scales: {
            x: { grid: { color: '#232840' }, ticks: { display: true, font: { size: 10 } } },
            y: { grid: { display: false }, ticks: { font: { size: 10 }, color: COLORS.text2 } }
          }
        }
      });
    }
    if (!charts.agingStacked) {
      // top 8 by balance
      const top8 = [...DATA.clients].sort((a, b) => b.balance - a.balance).slice(0, 8);
      // simulate aging split
      const getRand = (b, f) => Math.round(b * f + Math.random() * b * 0.04);
      const stacked = top8.map(c => {
        const base = c.balance;
        const o = c.overdue;
        if (o < 30) return [base * 0.85, base * 0.15, 0, 0, 0];
        if (o < 60) return [base * 0.4, base * 0.25, base * 0.25, base * 0.1, 0];
        if (o < 90) return [base * 0.2, base * 0.15, base * 0.2, base * 0.3, base * 0.15];
        return [base * 0.1, base * 0.1, base * 0.15, base * 0.25, base * 0.4];
      });
      const mkDs = (label, col, idx) => ({
        label, data: top8.map((_, i) => Math.round(stacked[i][idx])),
        backgroundColor: col, borderRadius: 4, borderSkipped: false
      });
      charts.agingStacked = new Chart(document.getElementById('agingStackedChart'), {
        type: 'bar',
        data: {
          labels: top8.map(c => c.name.split(' ')[0]),
          datasets: [
            mkDs('Corriente', COLORS.green, 0), mkDs('1–30d', COLORS.blue, 1),
            mkDs('31–60d', COLORS.yellow, 2), mkDs('61–90d', COLORS.orange, 3), mkDs('+90d', COLORS.red, 4)
          ]
        },
        options: {
          responsive: true, maintainAspectRatio: false,
          layout: { padding: 10 },
          plugins: { legend: { display: true, position: 'top', labels: { boxWidth: 10, font: { size: 10 } } } },
          scales: { x: { stacked: true, grid: { display: false }, ticks: { font: { size: 10 } } }, y: { stacked: true, grid: { color: 'rgba(255, 255, 255, 0.03)', borderDash: [4, 4], drawBorder: false }, ticks: { font: { size: 10 }, callback: v => '$' + (v / 1000).toFixed(0) + 'K' } } }
        }
      });
    }
  }

  if (tab === 'risk') buildRiskTable();
  if (tab === 'segmentation') buildSegmentation();
  if (tab === 'projection') {
    buildProjectionChart();
    buildProjectionTable();
  }
  if (tab === 'cashapp') {
    buildCashApp();
  }
  if (tab === 'refunds' || tab === '__all') {
    buildRefunds();
    buildRefundComparisonChart();
  }
}

// ── REFUNDS INTERANNUAL COMPARISON ─────────────────────────────
function buildRefundComparisonChart() {
  const rf = DATA.cashapp.refunds;
  const el = document.getElementById('refComparisonChart');
  if (!el || !rf) return;

  const isEn = typeof currentLang !== 'undefined' && currentLang === 'en';
  const monthNames = ['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic'];
  const monthNamesFull = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];

  // 1. Find the latest month in the data
  let latestSortKey = 0;
  if (rf.items && rf.items.length > 0) {
    rf.items.forEach(item => {
      if (!item.created) return;
      const parts = item.created.toLowerCase().split('-');
      if (parts.length >= 2) {
        const mStr = parts[1].trim();
        const yStr = parts.length >= 3 ? parts[2].trim() : '26';
        const mNum = monthNames.indexOf(mStr) + 1;
        const fullYear = yStr.length === 2 ? 2000 + parseInt(yStr) : parseInt(yStr);
        const key = fullYear * 100 + mNum;
        if (key > latestSortKey) latestSortKey = key;
      }
    });
  }

  // If no data, use current date
  if (latestSortKey === 0) {
    const now = new Date();
    latestSortKey = now.getFullYear() * 100 + (now.getMonth() + 1);
  }

  // 2. Generate a rolling 6-month window ending at latestSortKey
  const rollingLabels = [];
  const rollingKeys = []; // [202511, 202512, 202601, ...]
  
  let curYear = Math.floor(latestSortKey / 100);
  let curMonth = latestSortKey % 100;

  for (let i = 0; i < 6; i++) {
    const label = `${monthNamesFull[curMonth-1]} '${curYear.toString().slice(-2)}`;
    rollingLabels.unshift(label);
    rollingKeys.unshift(curYear * 100 + curMonth);
    
    curMonth--;
    if (curMonth === 0) {
      curMonth = 12;
      curYear--;
    }
  }

  // 3. Aggregate data for these 6 months
  const dataPoints = {}; // { 202603: { 'USD': 100, ... } }
  rollingKeys.forEach(k => dataPoints[k] = {});

  if (rf.items) {
    rf.items.forEach(item => {
      if (!item.created) return;
      const parts = item.created.toLowerCase().split('-');
      if (parts.length >= 2) {
        const mStr = parts[1].trim();
        const yStr = parts.length >= 3 ? parts[2].trim() : '26';
        const mNum = monthNames.indexOf(mStr) + 1;
        const fullYear = yStr.length === 2 ? 2000 + parseInt(yStr) : parseInt(yStr);
        const key = fullYear * 100 + mNum;
        
        if (dataPoints[key]) {
          const cur = item.currency || 'USD';
          const amt = Number(item.amount) || 0;
          dataPoints[key][cur] = (dataPoints[key][cur] || 0) + amt;
        }
      }
    });
  }

  const currenciesSet = new Set();
  rollingKeys.forEach(k => {
    Object.keys(dataPoints[k]).forEach(cur => currenciesSet.add(cur));
  });
  const currenciesFound = Array.from(currenciesSet);
  if (currenciesFound.length === 0) currenciesFound.push('USD');

  // 4. Update KPIs with FX Normalization
  const FX_RATES = { 'USD': 1.0, 'EUR': 1.08, 'MXN': 0.058, 'GBP': 1.25, 'CAD': 0.74 };
  
  const getUSD = (amt, cur) => amt * (FX_RATES[cur.toUpperCase()] || 1.0);

  const totalActualUSD = rf.items.reduce((s, i) => s + getUSD(Number(i.amount) || 0, i.currency || 'USD'), 0);
  
  // Like-for-Like simulation: only sum prevYear months that have data in current window
  const baselineMonthUSD = 25000; 
  let totalPrevUSD = 0;
  rollingKeys.forEach(k => {
     let monthTotalUSD = 0;
     currenciesFound.forEach(cur => {
        const amt = dataPoints[k][cur] || 0;
        monthTotalUSD += getUSD(amt, cur);
     });
     if (monthTotalUSD > 0) totalPrevUSD += baselineMonthUSD;
  });

  const growthPct = totalPrevUSD > 0 ? ((totalActualUSD - totalPrevUSD) / totalPrevUSD * 100).toFixed(1) : '0.0';

  const elActual = document.getElementById('refCompTotalActual');
  const elPrev = document.getElementById('refCompTotalPrev');
  const elGrowth = document.getElementById('refCompGrowth');
  const elBest = document.getElementById('refCompBestMonth');

  if (elActual) elActual.textContent = fmt(totalActualUSD) + ' (USD Eq.)';
  if (elPrev) elPrev.textContent = fmt(totalPrevUSD) + ' (USD Eq.)';
  if (elGrowth) {
    const sign = +growthPct >= 0 ? '+' : '';
    elGrowth.textContent = `${sign}${growthPct}%`;
    elGrowth.style.color = +growthPct >= 0 ? '#10b981' : '#ef4444';
  }

  // Best Month calculation (using USD normalization for comparison)
  let maxMonthUSD = 0;
  let bestMonthLabel = rollingLabels[0];
  rollingKeys.forEach((k, i) => {
     let monthTotalUSD = 0;
     currenciesFound.forEach(cur => {
        monthTotalUSD += getUSD(dataPoints[k][cur] || 0, cur);
     });
     if (monthTotalUSD > maxMonthUSD) {
       maxMonthUSD = monthTotalUSD;
       bestMonthLabel = rollingLabels[i];
     }
  });
  if (elBest) elBest.textContent = `${bestMonthLabel} · ${fmt(maxMonthUSD)} (USD Eq.)`;

  // 5. Delta Badges (MoM)
  const deltaRow = document.getElementById('refCompDeltaRow');
  if (deltaRow) {
    deltaRow.innerHTML = '';
    rollingKeys.forEach((k, i) => {
      if (i === 0) return;
      const prevK = rollingKeys[i-1];
      let currTotal = 0, prevTotal = 0;
      currenciesFound.forEach(cur => {
        currTotal += (dataPoints[k][cur] || 0);
        prevTotal += (dataPoints[prevK][cur] || 0);
      });
      if (prevTotal === 0 && currTotal === 0) return;
      
      const diff = currTotal - prevTotal;
      const pctVal = prevTotal > 0 ? ((diff / prevTotal) * 100).toFixed(1) : '100';
      const sign = diff >= 0 ? '+' : '';
      const cls = diff >= 0 ? 'positive' : 'negative';
      const badge = document.createElement('div');
      badge.className = `comp-delta-badge ${cls}`;
      badge.innerHTML = `<span class="delta-month">${rollingLabels[i]}</span><span>${sign}${pctVal}% MoM</span>`;
      deltaRow.appendChild(badge);
    });
  }

  // 6. Build Chart
  const CURRENCY_COLORS = { 'USD': '#8b5cf6', 'EUR': '#0ea5e9', 'MXN': '#10b981', 'GBP': '#f59e0b', 'CAD': '#f43f5e' };
  const datasets = currenciesFound.map((cur, i) => {
    const col = CURRENCY_COLORS[cur] || (i === 0 ? '#8b5cf6' : '#3b82f6');
    return {
      label: `Refunds (${cur})`,
      data: rollingKeys.map(k => dataPoints[k][cur] || 0),
      borderColor: col, borderWidth: 4, pointRadius: 6, pointHoverRadius: 9,
      pointBackgroundColor: col, pointBorderColor: '#fff', pointBorderWidth: 2, tension: 0.4, fill: true,
      backgroundColor: (context) => {
        const ctx = context.chart.ctx;
        const gradient = ctx.createLinearGradient(0, 0, 0, 300);
        gradient.addColorStop(0, hexToRgbA(col, 0.2));
        gradient.addColorStop(1, hexToRgbA(col, 0));
        return gradient;
      }
    };
  });

  function hexToRgbA(hex, alpha) {
    let r = 0, g = 0, b = 0;
    if (hex.length === 4) {
      r = "0x" + hex[1] + hex[1]; g = "0x" + hex[2] + hex[2]; b = "0x" + hex[3] + hex[3];
    } else if (hex.length === 7) {
      r = "0x" + hex[1] + hex[2]; g = "0x" + hex[3] + hex[4]; b = "0x" + hex[5] + hex[6];
    }
    return `rgba(${+r},${+g},${+b},${alpha})`;
  }

  if (!charts.refComparison) {
    charts.refComparison = new Chart(el, {
      type: 'line',
      data: { labels: rollingLabels, datasets: datasets },
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: {
          legend: { display: true, position: 'top', labels: { color: '#fff', font: { weight: '600' }, usePointStyle: true, padding: 20 } },
          tooltip: {
            backgroundColor: 'rgba(15, 23, 42, 0.98)', padding: 16, titleFont: { size: 18, weight: 'bold' }, bodyFont: { size: 16 },
            borderColor: 'rgba(255,255,255,0.2)', borderWidth: 1, displayColors: true, boxPadding: 8,
            callbacks: { label: (c) => `${c.dataset.label}: ${fmt(c.parsed.y)}` }
          }
        },
        scales: {
          x: { grid: { color: 'rgba(255,255,255,0.03)' }, ticks: { color: '#94a3b8', font: { size: 12 } } },
          y: { grid: { color: 'rgba(255,255,255,0.03)', drawBorder: false }, ticks: { color: '#94a3b8', font: { size: 12 }, callback: v => '$' + (v / 1000).toFixed(0) + 'K' } }
        }
      }
    });
  } else {
    charts.refComparison.data.labels = rollingLabels;
    charts.refComparison.data.datasets = datasets;
    charts.refComparison.update();
  }
}

// ── REFUNDS ANALYSIS ───────────────────────────────────────────
function buildRefunds() {
  const rf = DATA.cashapp.refunds;
  
  // Calculate KPIs and aggregate data for charts
  const totalAmtByCurrency = {};
  let maxAge = 0;
  const statusCounts = {};
  const subAmt = {};
  
  rf.items.forEach(item => {
    const amt = Number(item.amount) || 0;
    const cur = item.currency || 'USD';
    totalAmtByCurrency[cur] = (totalAmtByCurrency[cur] || 0) + amt;
    if (item.age > maxAge) maxAge = item.age;
    
    statusCounts[item.status] = (statusCounts[item.status] || 0) + 1;
    subAmt[item.subsidiary] = (subAmt[item.subsidiary] || 0) + amt;
  });
  
  // Update Top KPIs
  const totalAmtHtml = Object.entries(totalAmtByCurrency).map(([cur, amt]) => {
    return `<div style="display:flex; align-items:baseline; gap:10px; margin-bottom: 4px;">${fmt(amt)} <span style="font-size: 16px; color: var(--text2); font-weight: 600; opacity: 0.7;">${cur}</span></div>`;
  }).join('');
  
  const elTotal = document.getElementById('refTotalAmt');
  if (elTotal) elTotal.innerHTML = totalAmtHtml || '$0';
  
  const elCount = document.getElementById('refTotalCount');
  if (elCount) elCount.textContent = rf.items.length;
  
  const elAge = document.getElementById('refMaxAge');
  if (elAge) elAge.textContent = maxAge;

  // Chart 1: Subsidiary Doughnut
  const subLabels = Object.keys(subAmt);
  const subData = Object.values(subAmt);
  const subEl = document.getElementById('refSubsidiaryChart');
  if (subEl) {
    if (!charts.refSubsidiaryChart) {
      charts.refSubsidiaryChart = new Chart(subEl, {
        type: 'doughnut',
        data: {
          labels: subLabels,
          datasets: [{
            data: subData,
            backgroundColor: [COLORS.blue, COLORS.purple, COLORS.green, COLORS.yellow, COLORS.orange],
            borderWidth: 0, borderRadius: 8, spacing: 4, hoverOffset: 6
          }]
        },
        options: {
          responsive: true, maintainAspectRatio: false, cutout: '80%',
          plugins: { 
            legend: { position: 'right', labels: { color: COLORS.text2, font: { size: 11, family: 'Inter' }, usePointStyle: true, pointStyle: 'circle', boxWidth: 8 } },
            tooltip: {
              callbacks: {
                label: function(context) {
                  const total = context.dataset.data.reduce((a, b) => a + b, 0);
                  const val = context.raw;
                  const pct = ((val / total) * 100).toFixed(1) + '%';
                  return ` ${context.label}: $${val.toLocaleString()} (${pct})`;
                }
              }
            }
          }
        }
      });
    } else {
      charts.refSubsidiaryChart.data.labels = subLabels;
      charts.refSubsidiaryChart.data.datasets[0].data = subData;
      charts.refSubsidiaryChart.update();
    }
  }

  // Chart 2: Status Doughnut
  const statusLabels = Object.keys(statusCounts);
  const statusData = Object.values(statusCounts);
  const statusEl = document.getElementById('refStatusChart');
  if (statusEl) {
    if (!charts.refStatusChart) {
      charts.refStatusChart = new Chart(statusEl, {
        type: 'doughnut',
        data: {
          labels: statusLabels,
          datasets: [{
            data: statusData,
            backgroundColor: [COLORS.red, COLORS.yellow, COLORS.green],
            borderWidth: 0, borderRadius: 8, spacing: 4, hoverOffset: 6
          }]
        },
        options: {
          responsive: true, maintainAspectRatio: false, cutout: '80%',
          plugins: { 
            legend: { position: 'right', labels: { color: COLORS.text2, font: { size: 11, family: 'Inter' }, usePointStyle: true, pointStyle: 'circle', boxWidth: 8 } },
            tooltip: {
              callbacks: {
                label: function(context) {
                  const total = context.dataset.data.reduce((a, b) => a + b, 0);
                  const val = context.raw;
                  const pct = ((val / total) * 100).toFixed(1) + '%';
                  return ` ${context.label}: ${val} (${pct})`;
                }
              }
            }
          }
        }
      });
    } else {
      charts.refStatusChart.data.labels = statusLabels;
      charts.refStatusChart.data.datasets[0].data = statusData;
      charts.refStatusChart.update();
    }
  }

  // Table: Pending Refunds
  const tableBody = document.getElementById('refundTableBody');
  if (tableBody) {
    tableBody.innerHTML = '';
    rf.items.forEach((item, index) => {
      const cls = item.status === 'Pendiente' ? 'critical' : (item.status === 'Validando' ? 'high' : 'medium');
      const ageColor = item.age > 10 ? 'color: var(--red)' : '';
      
      let formattedDate = item.created;
      if (item.created) {
        const parts = item.created.toLowerCase().split('-');
        if (parts.length >= 2) {
          const d = parts[0].padStart(2, '0');
          const m = parts[1].charAt(0).toUpperCase() + parts[1].slice(1);
          const y = parts.length >= 3 ? (parts[2].length === 2 ? '20'+parts[2] : parts[2]) : '';
          formattedDate = y ? `${d}-${m}-${y}` : `${d}-${m}`;
        }
      }
      
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td style="background-color: rgba(16, 185, 129, 0.15); color: #10b981;"><strong>${item.rftNumber}</strong></td>
        <td>${formattedDate}</td>
        <td>${item.subsidiary}</td>
        <td>${Number(item.amount).toLocaleString('en-US', {minimumFractionDigits: 2})}</td>
        <td>${item.currency}</td>
        <td><span style="${ageColor}; font-weight: bold;">${item.age}</span></td>
        <td>${item.responsable}</td>
        <td><a href="${item.link}" target="_blank" style="color: var(--blue); text-decoration: underline;">${item.link}</a></td>
        <td><span class="risk-badge ${cls}">${item.status}</span></td>
        <td>
          <button onclick="editRefund(${index})" style="background:none;border:none;color:var(--yellow);cursor:pointer;margin-right:8px;" title="Editar">✏️</button>
          <button onclick="deleteRefund(${index})" style="background:none;border:none;color:var(--red);cursor:pointer;" title="Eliminar">🗑️</button>
        </td>
      `;
      tableBody.appendChild(tr);
    });
  }
}

window.deleteRefund = function(index) {
  if (confirm('¿Seguro que deseas eliminar este registro?')) {
    DATA.cashapp.refunds.items.splice(index, 1);
    saveState();
    buildRefunds();
    buildRefundComparisonChart();
  }
};

window.editRefund = function(index) {
  const tableBody = document.getElementById('refundTableBody');
  const row = tableBody.children[index];
  const item = DATA.cashapp.refunds.items[index];
  
  row.innerHTML = `
    <td><input type="text" id="edit_rft_${index}" value="${item.rftNumber}" style="width:80px;background:var(--surface);color:var(--text);border:1px solid #444;padding:4px;border-radius:4px;"></td>
    <td><input type="text" id="edit_cre_${index}" value="${item.created}" style="width:70px;background:var(--surface);color:var(--text);border:1px solid #444;padding:4px;border-radius:4px;"></td>
    <td><input type="text" id="edit_sub_${index}" value="${item.subsidiary}" style="width:110px;background:var(--surface);color:var(--text);border:1px solid #444;padding:4px;border-radius:4px;"></td>
    <td><input type="number" id="edit_amt_${index}" value="${item.amount}" style="width:90px;background:var(--surface);color:var(--text);border:1px solid #444;padding:4px;border-radius:4px;"></td>
    <td><input type="text" id="edit_cur_${index}" value="${item.currency}" style="width:50px;background:var(--surface);color:var(--text);border:1px solid #444;padding:4px;border-radius:4px;"></td>
    <td><input type="number" id="edit_age_${index}" value="${item.age}" style="width:50px;background:var(--surface);color:var(--text);border:1px solid #444;padding:4px;border-radius:4px;"></td>
    <td><input type="text" id="edit_res_${index}" value="${item.responsable}" style="width:80px;background:var(--surface);color:var(--text);border:1px solid #444;padding:4px;border-radius:4px;"></td>
    <td><input type="text" id="edit_lnk_${index}" value="${item.link}" style="width:100px;background:var(--surface);color:var(--text);border:1px solid #444;padding:4px;border-radius:4px;"></td>
    <td>
      <select id="edit_sta_${index}" style="background:var(--surface);color:var(--text);border:1px solid #444;padding:4px;border-radius:4px;">
        <option value="Pendiente" ${item.status==='Pendiente'?'selected':''}>Pendiente</option>
        <option value="Validando" ${item.status==='Validando'?'selected':''}>Validando</option>
        <option value="Completado" ${item.status==='Completado'?'selected':''}>Completado</option>
      </select>
    </td>
    <td>
      <button onclick="saveEditRefund(${index})" style="background:none;border:none;color:var(--green);cursor:pointer;margin-right:8px;" title="Guardar">✔️</button>
      <button onclick="buildRefunds()" style="background:none;border:none;color:var(--red);cursor:pointer;" title="Cancelar">❌</button>
    </td>
  `;
};

window.saveEditRefund = function(index) {
  const item = DATA.cashapp.refunds.items[index];
  item.rftNumber = document.getElementById(`edit_rft_${index}`).value;
  item.created = document.getElementById(`edit_cre_${index}`).value;
  item.subsidiary = document.getElementById(`edit_sub_${index}`).value;
  item.amount = Number(document.getElementById(`edit_amt_${index}`).value);
  item.currency = document.getElementById(`edit_cur_${index}`).value;
  item.age = Number(document.getElementById(`edit_age_${index}`).value);
  item.responsable = document.getElementById(`edit_res_${index}`).value;
  item.link = document.getElementById(`edit_lnk_${index}`).value;
  item.status = document.getElementById(`edit_sta_${index}`).value;
  
  saveState();
  buildRefunds();
  buildRefundComparisonChart();
}

// ── CASH APPLICATIONS ──────────────────────────────────────────
function filterCashApp(type) {
  // If clicking same filter, toggle to 'all'
  if (currentCaFilter === type) {
    currentCaFilter = 'all';
  } else {
    currentCaFilter = type;
  }

  // Visual update of cards
  document.querySelectorAll('.kpi-card').forEach(c => c.classList.remove('active-filter'));
  if (currentCaFilter !== 'all') {
    const card = document.getElementById(`ca-card-${currentCaFilter}`);
    if (card) card.classList.add('active-filter');
  }

  // Update Table Title
  const title = document.getElementById('ca-table-title');
  if (title) {
    if (currentCaFilter === 'unapplied') title.textContent = 'Partidas Pendientes: Unapplied Cash';
    else if (currentCaFilter === 'suspense') title.textContent = 'Partidas Pendientes: Suspense Account';
    else title.textContent = 'Top Partidas Pendientes de Aplicar';
  }

  buildCashApp();
}

function buildCashApp() {
  const ca = DATA.cashapp;

  document.getElementById('ca-unapplied').textContent = fmt(ca.kpis.unapplied);
  document.getElementById('ca-suspense').textContent = fmt(ca.kpis.suspense);
  document.getElementById('ca-automatch').textContent = ca.kpis.autoMatch + '%';
  if (document.getElementById('ca-time')) {
    document.getElementById('ca-time').textContent = ca.kpis.manTime + ' min';
  }

  // --- Generate Insights ---
  const insightsList = document.getElementById('ca-insights-list');
  if (insightsList) {
    insightsList.innerHTML = '';

    // Insight 1: Auto-match rate
    const autoInsight = document.createElement('li');
    if (ca.kpis.autoMatch >= 80) {
      autoInsight.innerHTML = `<strong>Tasa de Aplicación Automática Saludable:</strong> El sistema está emparejando automáticamente el ${ca.kpis.autoMatch}% de los ingresos, reduciendo significativamente la carga manual.`;
      autoInsight.style.marginBottom = "6px";
    } else {
      autoInsight.innerHTML = `<strong>Oportunidad de Eficiencia (${ca.kpis.autoMatch}%):</strong> Aumentar el auto-match reduciría el tiempo extra manual promedio que actualmente requiere ${ca.kpis.manTime} min por partida.`;
      autoInsight.style.marginBottom = "6px";
    }
    insightsList.appendChild(autoInsight);

    // Insight 2: Unapplied vs Suspense
    const unappInsight = document.createElement('li');
    unappInsight.innerHTML = `<strong>Flujo de Efectivo Retenido:</strong> Actualmente, existen <strong>${fmt(ca.kpis.unapplied)}</strong> pendientes de aplicar a las cuentas de los clientes. Reducir esta cantidad impactaría positivamente en el flujo de caja inmediato. De este monto total, <strong>${fmt(ca.kpis.suspense)}</strong> se encuentran etiquetados como <em>Cuenta de Suspenso</em> por estar totalmente sin identificar.`;
    unappInsight.style.marginBottom = "6px";
    insightsList.appendChild(unappInsight);

    // Insight 3: Biggest Suspense Offender
    const sus = ca.suspense;
    const maxSusVal = Math.max(sus.noRef, sus.invalidAmt, sus.noClient, sus.doublePay);
    let topReason = '';
    if (maxSusVal === sus.noRef) topReason = "Falta de Referencia";
    else if (maxSusVal === sus.invalidAmt) topReason = "Monto Inválido";
    else if (maxSusVal === sus.noClient) topReason = "Cliente No Encontrado";
    else topReason = "Doble Pago";

    const susInsight = document.createElement('li');
    susInsight.innerHTML = `<strong>Causa Principal de Descuadres:</strong> El motivo #1 de partidas sin registrar es por <strong>${topReason}</strong> (${maxSusVal}% en la muestra de suspenso). <em>Acción recomendada: Automatizar recordatorios para que los clientes adjunten esta información en sus comprobantes de pago.</em>`;
    susInsight.style.marginBottom = "6px";
    insightsList.appendChild(susInsight);

    // Insight 4: YoY Comparison Chart Analysis
    const history = ca.appliedCashHistory;
    if (history && Object.keys(history).length >= 1) {
      const yearKeys = Object.keys(history);
      const key1 = yearKeys[0];
      const key2 = yearKeys[1]; // might be undefined

      const curr = history[key1];
      const prev = key2 ? history[key2] : null;
      const months = DATA.months;

      const totalCurr = curr.reduce((s, v) => s + v, 0);
      const label1 = key1.replace('year', '').replace('currentYear', '2026');
      
      const yoyInsight = document.createElement('li');
      yoyInsight.style.marginTop = "10px";
      yoyInsight.style.paddingTop = "10px";
      yoyInsight.style.borderTop = "1px solid rgba(255,255,255,0.06)";

      if (prev) {
        const totalPrev = prev.reduce((s, v) => s + v, 0);
        const yoyPct = ((totalCurr - totalPrev) / totalPrev * 100).toFixed(1);
        const label2 = key2.replace('year', '').replace('prevYear', '2025');

        // Best and worst month by delta
        const deltas = curr.map((v, i) => ({ month: months[i], delta: v - (prev[i] || 0), pct: prev[i] ? ((v - prev[i]) / prev[i] * 100).toFixed(1) : '0' }));
        const bestMonth = deltas.reduce((a, b) => b.delta > a.delta ? b : a);
        const worstMonth = deltas.reduce((a, b) => b.delta < a.delta ? b : a);

        // Acceleration: compare avg growth of last 3 vs first 3 months
        const firstHalf = deltas.slice(0, Math.floor(deltas.length / 2));
        const secondHalf = deltas.slice(Math.floor(deltas.length / 2));
        const avgFirst = firstHalf.reduce((s, d) => s + parseFloat(d.pct), 0) / (firstHalf.length || 1);
        const avgSecond = secondHalf.reduce((s, d) => s + parseFloat(d.pct), 0) / (secondHalf.length || 1);
        const isAccelerating = avgSecond > avgFirst;
        const trendLabel = isAccelerating
          ? `<span style="color:#10b981">acelerando ▲</span> (promedio reciente +${avgSecond.toFixed(1)}% vs inicio +${avgFirst.toFixed(1)}%)`
          : `<span style="color:#f59e0b">desacelerando ▼</span> (promedio reciente +${avgSecond.toFixed(1)}% vs inicio +${avgFirst.toFixed(1)}%)`;

        const yoySign = yoyPct >= 0 ? '+' : '';
        const yoyColor = yoyPct >= 0 ? '#10b981' : '#ef4444';

        yoyInsight.innerHTML = `
          <strong>📈 Análisis Interanual (${label1} vs ${label2}):</strong>
          El efectivo aplicado acumulado en ${label1} es de <strong style="color:#3b82f6">${fmt(totalCurr)}</strong>,
          ${totalCurr > totalPrev ? 'superando' : 'por debajo de'} los <strong style="color:rgba(255,255,255,0.5)">${fmt(totalPrev)}</strong> de ${label2}
          — un ${totalCurr > totalPrev ? 'crecimiento' : 'cambio'} de <strong style="color:${yoyColor}">${yoySign}${yoyPct}%</strong>.
          El mes con mayor mejora fue <strong style="color:#10b981">${bestMonth.month} (+${bestMonth.pct}%)</strong>.
          La tendencia está ${trendLabel}.`;
      } else {
        yoyInsight.innerHTML = `
          <strong>📈 Análisis de Efectivo Aplicado (${label1}):</strong>
          El efectivo aplicado acumulado en ${label1} es de <strong style="color:#3b82f6">${fmt(totalCurr)}</strong>.
          No hay datos de años anteriores para realizar una comparativa interanual, pero el flujo se mantiene estable.`;
      }
      insightsList.appendChild(yoyInsight);
    }
  }

  // Chart 1: Auto vs Manual applying matching rate per month
  const months = DATA.months;
  const baseAuto = Number(ca.kpis.autoMatch) || 80;
  const autoData = months.map(() => Math.min(100, Math.floor(baseAuto + (Math.random() * 10 - 5))));
  const manData = autoData.map(v => 100 - v);

  if (!charts.caMatch) {
    charts.caMatch = new Chart(document.getElementById('caMatchChart'), {
      type: 'bar',
      data: {
        labels: months,
        datasets: [
          { label: 'Automático (%)', data: autoData, backgroundColor: COLORS.green, borderRadius: 4 },
          { label: 'Manual (%)', data: manData, backgroundColor: COLORS.orange, borderRadius: 4 }
        ]
      },
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: { legend: { display: true, position: 'top', labels: { boxWidth: 10, font: { size: 10 } } } },
        scales: {
          x: { stacked: true, grid: { display: false }, ticks: { font: { size: 10 } } },
          y: { stacked: true, max: 100, grid: { color: '#232840' }, ticks: { font: { size: 10 } } }
        }
      }
    });
  } else {
    charts.caMatch.data.labels = months;
    charts.caMatch.data.datasets[0].data = autoData;
    charts.caMatch.data.datasets[1].data = manData;
    charts.caMatch.update();
  }

  buildCashComparison();

  // Chart 2: Aging of Unapplied Cash (Data Driven)
  const items = ca.items || [];
  let d0_3 = 0, d4_7 = 0, d8_14 = 0, d15p = 0;
  
  items.forEach(item => {
    const days = Number(item.days) || 0;
    if (days <= 3) d0_3 += item.amount;
    else if (days <= 7) d4_7 += item.amount;
    else if (days <= 14) d8_14 += item.amount;
    else d15p += item.amount;
  });

  // If no items, fallback to dummy data or zeros
  if (items.length === 0) {
    const total = ca.kpis.unapplied || 0;
    d0_3 = total * 0.6;
    d4_7 = total * 0.25;
    d8_14 = total * 0.1;
    d15p = total * 0.05;
  }

  if (!charts.caAging) {
    charts.caAging = new Chart(document.getElementById('caAgingChart'), {
      type: 'doughnut',
      data: {
        labels: ['0-3 días', '4-7 días', '8-14 días', '+15 días'],
        datasets: [{
          data: [d0_3, d4_7, d8_14, d15p],
          backgroundColor: [COLORS.blue, COLORS.green, COLORS.yellow, COLORS.red],
          borderWidth: 0,
          hoverOffset: 8
        }]
      },
      options: {
        responsive: true, maintainAspectRatio: false, cutout: '65%',
        plugins: {
          legend: { display: true, position: 'right', labels: { boxWidth: 10, font: { size: 10 } } },
          tooltip: { enabled: true, callbacks: { label: (ctx) => ' ' + fmt(ctx.raw) } }
        }
      }
    });
  } else {
    charts.caAging.data.datasets[0].data = [d0_3, d4_7, d8_14, d15p];
    charts.caAging.update();
  }

  // Chart 3: Suspense Composition
  const suspData = [ca.suspense.noRef, ca.suspense.invalidAmt, ca.suspense.noClient, ca.suspense.doublePay];

  if (!charts.caSuspense) {
    charts.caSuspense = new Chart(document.getElementById('caSuspenseChart'), {
      type: 'pie',
      data: {
        labels: ['Falta Referencia', 'Monto Inválido', 'Cliente No Encontrado', 'Doble Pago'],
        datasets: [{
          data: suspData,
          backgroundColor: [COLORS.orange, COLORS.purple, COLORS.yellow, COLORS.red],
          borderWidth: 0,
          hoverOffset: 8
        }]
      },
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: {
          legend: { display: true, position: 'right', labels: { boxWidth: 10, font: { size: 10 } } },
          tooltip: { enabled: true }
        }
      }
    });
  } else {
    charts.caSuspense.data.datasets[0].data = suspData;
    charts.caSuspense.update();
  }

  // Table: Top Unapplied Items
  const tableBody = document.getElementById('caTableBody');
  if (tableBody) {
    tableBody.innerHTML = '';

    // Filter logic unified with resyncData
    let items = ca.items;
    if (currentCaFilter === 'unapplied') {
      items = ca.items; // Everything is unapplied
    } else if (currentCaFilter === 'suspense') {
      items = ca.items.filter(i => i.status === 'Pendiente' || (i.client && i.client.includes('?')));
    }

    items.forEach(item => {
      const cls = item.status === 'Investigando' ? 'high' : item.status === 'Contactado' ? 'medium' : 'critical';
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td><strong>${item.ref}</strong></td>
        <td>${fmt(item.amount)}</td>
        <td>${item.date}</td>
        <td><span style="color:var(--red)">${item.days}</span></td>
        <td>${item.client}</td>
        <td><span class="risk-badge ${cls}">${item.status}</span></td>
      `;
      tableBody.appendChild(tr);
    });
  }
}

// ── YEARLY CASH COMPARISON ──────────────────────────────────────
function buildCashComparison() {
  const history = DATA.cashapp.appliedCashHistory;
  if (!history) return;
  const el = document.getElementById('caComparisonChart');
  if (!el) return;

  const yearKeys = Object.keys(history);
  if (yearKeys.length === 0) return;

  const allMonths = DATA.months;
  const isEn = typeof currentLang !== 'undefined' && currentLang === 'en';

  // Filter out empty keys and sort them so currentYear/prevYear are first if they exist
  const validKeys = yearKeys.filter(k => k.trim() !== '' && history[k] && history[k].length > 0);
  if (validKeys.length === 0) return;

  const key1 = validKeys[0];
  const key2 = validKeys[1];
  const data1 = history[key1] || [];
  const data2 = history[key2] || [];
  
  // Sync labels with data length to avoid gaps on the right
  const months = allMonths.slice(0, data1.length);

  const getLabel = (k) => {
    let l = k.replace('year', '').replace('Year', '');
    const low = l.toLowerCase();
    if (low.includes('current') || low.includes('actual') || k === 'currentYear') {
      return isEn ? 'Current (2026)' : 'Actual (2026)';
    }
    if (low.includes('prev') || low.includes('anterior') || k === 'prevYear') {
      return isEn ? 'Previous (2025)' : 'Anterior (2025)';
    }
    // If it's just a number, leave it. If not, capitalize.
    if (isNaN(l)) l = l.charAt(0).toUpperCase() + l.slice(1);
    return l;
  };

  const label1 = getLabel(key1);
  const label2 = key2 ? getLabel(key2) : '';

  // ── Populate KPI summary pills ──────────────────────────────────
  const total1 = data1.reduce((s, v) => s + v, 0);
  const total2 = data2.length > 0 ? data2.reduce((s, v) => s + v, 0) : 0;
  
  const growthPct = total2 > 0 ? ((total1 - total2) / total2 * 100).toFixed(1) : '0.0';
  const bestIdx = data1.indexOf(Math.max(...data1));

  const el2026 = document.getElementById('compTotal2026');
  const el2025 = document.getElementById('compTotal2025');
  const elGrowth = document.getElementById('compGrowth');
  const elBest = document.getElementById('compBestMonth');

  if (el2026) {
    el2026.textContent = fmt(total1);
    const labelEl = el2026.previousElementSibling;
    if (labelEl) labelEl.textContent = (isEn ? 'Total Applied ' : 'Total Aplicado ') + label1;
  }
  if (el2025) {
    if (key2) {
      el2025.textContent = fmt(total2);
      const labelEl = el2025.previousElementSibling;
      if (labelEl) labelEl.textContent = (isEn ? 'Total Applied ' : 'Total Aplicado ') + label2;
      el2025.parentElement.style.display = '';
    } else {
      el2025.parentElement.style.display = 'none';
    }
  }
  if (elGrowth) {
    if (key2) {
      const sign = growthPct >= 0 ? '+' : '';
      elGrowth.textContent = `${sign}${growthPct}%`;
      elGrowth.style.color = +growthPct >= 0 ? '#10b981' : '#ef4444';
      elGrowth.parentElement.style.display = '';
    } else {
      elGrowth.parentElement.style.display = 'none';
    }
  }
  if (elBest) {
    const bestLabel = (isEn ? 'Best Month (' : 'Mejor Mes (') + label1 + ')';
    const labelEl = elBest.previousElementSibling;
    if (labelEl) labelEl.textContent = bestLabel;
    elBest.textContent = `${months[bestIdx]} · ${fmt(data1[bestIdx])}`;
  }

  // ── Populate month delta badges ─────────────────────────────────
  const deltaRow = document.getElementById('caCompDeltaRow');
  if (deltaRow) {
    deltaRow.innerHTML = ''; // Always clear to rebuild
    if (key2) {
      months.forEach((m, i) => {
        const val1 = data1[i] || 0;
        const val2 = data2[i] || 0;
        if (val2 === 0) return;
        
        const delta = val1 - val2;
        const pct = ((delta / val2) * 100).toFixed(1);
        const sign = delta >= 0 ? '+' : '';
        const cls = delta >= 0 ? 'positive' : 'negative';
        const badge = document.createElement('div');
        badge.className = `comp-delta-badge ${cls}`;
        badge.innerHTML = `<span class="delta-month">${m}</span><span>${sign}${pct}%</span>`;
        deltaRow.appendChild(badge);
      });
    }
  }

  // ── Build or update chart ───────────────────────────────────────
  if (!charts.caComparison) {
    const datasets = [];
    
    // Update the chart tag (e.g. "2026 vs 2025")
    const chartTag = document.querySelector('#tab-cashapp .chart-tag');
    if (chartTag && validKeys.length >= 2) {
      chartTag.textContent = `${label1} vs ${label2}`;
    }

    const COLORS_LIST = ['#3b82f6', 'rgba(255,255,255,0.4)', '#f59e0b', '#10b981', '#ef4444', '#8b5cf6', '#eab308'];

    validKeys.forEach((yk, idx) => {
      const label = getLabel(yk);
      const color = COLORS_LIST[idx % COLORS_LIST.length];
      const isPrimary = idx === 0;

      datasets.push({
        label: label,
        data: history[yk],
        borderColor: color,
        backgroundColor: isPrimary ? 'rgba(59,130,246,0.12)' : 'transparent',
        borderDash: isPrimary ? [] : [6, 4],
        fill: isPrimary,
        tension: 0.45,
        pointRadius: isPrimary ? 5 : 3,
        pointHoverRadius: isPrimary ? 8 : 6,
        pointBackgroundColor: color,
        borderWidth: isPrimary ? 3 : 2,
        order: idx + 1
      });
    });

    charts.caComparison = new Chart(el, {
      type: 'line',
      data: {
        labels: months,
        datasets: datasets
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        interaction: { mode: 'index', intersect: false },
        plugins: {
          legend: {
            display: true,
            position: 'top',
            labels: { 
              color: COLORS.text, 
              font: { size: 11 },
              padding: 20,
              usePointStyle: false,
              boxWidth: 30
            }
          },
          tooltip: {
            backgroundColor: 'rgba(17,19,26,0.97)',
            titleColor: '#fff',
            bodyColor: '#94a3b8',
            borderColor: 'rgba(255,255,255,0.08)',
            borderWidth: 1,
            padding: 12,
            callbacks: {
              title: (items) => months[items[0].dataIndex],
              beforeBody: () => '─────────────────',
              label: (ctx) => {
                const idx = ctx.dataIndex;
                const val = ctx.raw;
                return `  ${ctx.dataset.label}: ${fmt(val)}`;
              }
            }
          }
        },
        scales: {
          x: { grid: { display: false }, ticks: { color: COLORS.text2, font: { size: 11 }, padding: 6 } },
          y: { grid: { color: 'rgba(255,255,255,0.04)', borderDash: [4, 4] }, ticks: { color: COLORS.text2, font: { size: 11 }, padding: 8, callback: v => '$' + (v / 1000000).toFixed(1) + 'M' }, border: { display: false } }
        }
      }
    });
  } else {
    // Re-initialize if the number of years changed or labels changed
    const currentDatasetCount = charts.caComparison.data.datasets.length;
    const newDatasetCount = Object.keys(history).length;
    
    if (currentDatasetCount !== newDatasetCount) {
      charts.caComparison.destroy();
      delete charts.caComparison;
      buildCashComparison(); // Re-run to create fresh
    } else {
      // Just update data of existing datasets
      const yearKeys = Object.keys(history);
      yearKeys.forEach((yk, idx) => {
        if (charts.caComparison.data.datasets[idx]) {
          charts.caComparison.data.datasets[idx].data = history[yk];
        }
      });
      charts.caComparison.update();
    }
  }
}

// ── RISK TABLE ─────────────────────────────────────────────────
function buildRiskTable() {
  const body = document.getElementById('riskTableBody');
  if (body.children.length > 0) return;
  const risky = [...DATA.clients].filter(c => c.score >= 50).sort((a, b) => b.score - a.score);
  risky.forEach(c => {
    const cls = c.score >= 80 ? 'critical' : c.score >= 60 ? 'high' : 'medium';
    const lbl = c.score >= 80 ? 'Crítico' : c.score >= 60 ? 'Alto' : 'Medio';
    const trendHtml = c.trend === 'up' ? '<span class="trend-tag trend-up">▲ Deteriorando</span>' :
      c.trend === 'down' ? '<span class="trend-tag trend-down">▼ Mejorando</span>' :
        '<span class="trend-tag trend-stable">→ Estable</span>';
    const barColor = c.score >= 80 ? COLORS.red : c.score >= 60 ? COLORS.orange : COLORS.yellow;
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><strong>${c.name}</strong></td>
      <td>${fmt(c.balance)}</td>
      <td>${c.overdue} días</td>
      <td>${c.limitExc > 0 ? '<span style="color:var(--red)">+' + c.limitExc + '%</span>' : '<span style="color:var(--green)">No excedido</span>'}</td>
      <td>${trendHtml}</td>
      <td>
        <div class="score-bar-wrap">
          <div class="score-bar"><div class="score-fill" style="width:${c.score}%;background:${barColor}"></div></div>
          <span style="font-size:12px;font-weight:700;color:${barColor}">${c.score}</span>
        </div>
      </td>
      <td><span class="risk-badge ${cls}">${lbl}</span></td>`;
    body.appendChild(tr);
  });
}

// ── SEGMENTATION MATRIX ────────────────────────────────────────
function buildSegmentation() {
  const quads = { strategic: [], alert: [], stable: [], lowrisk: [] };
  DATA.clients.forEach(c => quads[c.seg].push(c));
  const render = (id, arr) => {
    const el = document.getElementById('seg-' + id);
    if (el.children.length > 0) return;
    arr.forEach(c => {
      const div = document.createElement('div');
      div.className = 'client-chip';
      div.innerHTML = `<span class="cn">${c.name}</span><span class="cv">${fmt(c.balance)}</span>`;
      el.appendChild(div);
    });
  };
  render('strategic', quads.strategic);
  render('alert', quads.alert);
  render('stable', quads.stable);
  render('lowrisk', quads.lowrisk);
}

// ── PROJECTION ─────────────────────────────────────────────────
function buildProjectionChart() {
  if (charts.projection) {
    charts.projection.destroy();
    delete charts.projection;
  }
  
  // ── Weekly detection logic ──────────────────────────────────────
  // Instead of hardcoded strings, we'll try to extract the day number
  const getWeekIndex = (weekStr) => {
    if (!weekStr) return -1;
    // Extract numbers from string (e.g., "Mar 04" -> 4, "2026-03-05" -> 5)
    const matches = weekStr.match(/\d+/g);
    if (!matches || matches.length === 0) return -1;
    
    // Assume the last or only number is the day if it's <= 31
    const day = parseInt(matches[matches.length - 1]);
    if (day >= 1 && day <= 7) return 0;
    if (day >= 8 && day <= 14) return 1;
    if (day >= 15 && day <= 21) return 2;
    if (day >= 22) return 3;
    return -1;
  };

  const isEn = currentLang === 'en';
  const weeks = isEn 
    ? ['Week 1 (Mar 1–7)', 'Week 2 (Mar 8–14)', 'Week 3 (Mar 15–21)', 'Week 4 (Mar 22–31)']
    : ['Semana 1 (Mar 1–7)', 'Semana 2 (Mar 8–14)', 'Semana 3 (Mar 15–21)', 'Semana 4 (Mar 22–31)'];

  const dataByWeek = [
    { high: 0, med: 0, low: 0 },
    { high: 0, med: 0, low: 0 },
    { high: 0, med: 0, low: 0 },
    { high: 0, med: 0, low: 0 }
  ];

  DATA.projection.forEach(p => {
    const idx = getWeekIndex(p.week);
    if (idx !== -1 && dataByWeek[idx]) {
      dataByWeek[idx][p.prob] += p.amount;
    }
  });

  charts.projection = new Chart(document.getElementById('projectionChart'), {
    type: 'bar',
    data: {
      labels: weeks,
      datasets: [
        { label: isEn ? 'High Probability' : 'Alta Probabilidad', data: dataByWeek.map(w => w.high), backgroundColor: COLORS.green, borderRadius: 6, stack: 's' },
        { label: isEn ? 'Medium Probability' : 'Media Probabilidad', data: dataByWeek.map(w => w.med), backgroundColor: COLORS.yellow, borderRadius: 0, stack: 's' },
        { label: isEn ? 'Low Probability' : 'Baja Probabilidad', data: dataByWeek.map(w => w.low), backgroundColor: COLORS.red, borderRadius: 0, stack: 's' },
      ]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      layout: { padding: 10 },
      plugins: { legend: { display: true, position: 'top', labels: { boxWidth: 10, font: { size: 10 } } } },
      scales: { 
        x: { stacked: true, grid: { display: false }, ticks: { font: { size: 10 } } }, 
        y: { stacked: true, grid: { color: '#232840' }, ticks: { font: { size: 10 }, callback: v => '$' + (v / 1000).toFixed(0) + 'K' } } 
      }
    }
  });
}

function buildProjectionTable() {
  const cont = document.getElementById('projectionTable');
  if (cont.children.length > 0) return;
  const table = document.createElement('table');
  table.className = 'proj-table';
  table.innerHTML = `<thead><tr><th>Cliente</th><th>Fecha Estimada</th><th>Monto Proyectado</th><th>Probabilidad</th><th>Estado</th></tr></thead><tbody></tbody>`;
  const body = table.querySelector('tbody');
  DATA.projection.forEach(p => {
    const probClass = p.prob === 'high' ? 'prob-high' : p.prob === 'med' ? 'prob-med' : 'prob-low';
    const probText = p.prob === 'high' ? 'Alta (≥75%)' : p.prob === 'med' ? 'Media (40-74%)' : 'Baja (<40%)';
    const status = p.prob === 'high' ? '✔ Comprometido' : p.prob === 'med' ? '⏳ En Negociación' : '⚠ Incierto';
    const tr = document.createElement('tr');
    tr.innerHTML = `<td><strong>${p.client}</strong></td><td>${p.week}, 2026</td><td>${fmt(p.amount)}</td><td class="${probClass}">${probText}</td><td>${status}</td>`;
    body.appendChild(tr);
  });
  cont.appendChild(table);
}

// ── DATA REGENERATION ──────────────────────────────────────────
function regenerateData() {
  const rnd = (a, b) => +(a + (Math.random() * (b - a))).toFixed(1);
  DATA.dso.actual = rnd(33, 45);
  DATA.dso.prev = rnd(31, 42);
  DATA.aging.current = Math.round(1200000 + Math.random() * 900000);
  DATA.aging.d30 = Math.round(700000 + Math.random() * 500000);
  DATA.aging.d60 = Math.round(400000 + Math.random() * 400000);
  DATA.aging.d90 = Math.round(250000 + Math.random() * 350000);
  DATA.aging.d90p = Math.round(150000 + Math.random() * 300000);
  DATA.totalAR = DATA.aging.current + DATA.aging.d30 + DATA.aging.d60 + DATA.aging.d90 + DATA.aging.d90p;

  // Randomize Refunds
  DATA.cashapp.refunds.total = Math.round(50000 + Math.random() * 100000);
  DATA.cashapp.refunds.history = DATA.months.map(() => Math.round(10000 + Math.random() * 20000));
  DATA.cashapp.refunds.items.forEach(item => {
    item.amount = Math.round(2000 + Math.random() * 15000);
  });
  // Destroy all cached charts so they rebuild
  Object.values(charts).forEach(c => c.destroy && c.destroy());
  Object.keys(charts).forEach(k => delete charts[k]);
  // Clear rendered tables/segments
  ['riskTableBody', 'seg-strategic', 'seg-alert', 'seg-stable', 'seg-lowrisk'].forEach(id => {
    const el = document.getElementById(id); if (el) el.innerHTML = '';
  });
  const projTable = document.getElementById('projectionTable');
  if (projTable) projTable.innerHTML = '';
  // Update KPIs
  document.getElementById('kpi-total').textContent = fmt(DATA.totalAR);
  document.getElementById('kpi-dso').textContent = DATA.dso.actual + ' días';
  const dsoDiff = (DATA.dso.actual - DATA.dso.target).toFixed(1);
  const dsoDelta = document.getElementById('kpi-dso-delta');
  if (dsoDiff > 0) {
    dsoDelta.textContent = `↑ ${dsoDiff} días vs objetivo`;
    dsoDelta.className = 'kpi-delta negative';
  } else {
    dsoDelta.textContent = `↓ ${Math.abs(dsoDiff)} días vs objetivo`;
    dsoDelta.className = 'kpi-delta positive';
  }
  document.getElementById('kpi-collected').textContent = fmt(Math.round(DATA.totalAR * 0.45));
  resyncData();
  initCharts(activeTab);
  showToast('🔄 Datos regenerados exitosamente');
}

// ════════════════════════════════════════════════════════════════
//  UPLOAD MODAL – Subir datos reales (CSV / JSON)
// ════════════════════════════════════════════════════════════════

let _pendingFile = null;

function openUploadModal() {
  document.getElementById('uploadModal').classList.add('open');
  resetDropZone();
}

function closeUploadModal() {
  document.getElementById('uploadModal').classList.remove('open');
  resetDropZone();
}

function handleOverlayClick(e) {
  if (e.target === document.getElementById('uploadModal')) closeUploadModal();
}

function resetDropZone() {
  _pendingFile = null;
  const dz = document.getElementById('dropZone');
  dz.classList.remove('dragover', 'has-file');
  dz.innerHTML = `
    <div class="drop-icon">📄</div>
    <p class="drop-text">Arrastra tu archivo aquí</p>
    <p class="drop-sub">o haz <strong>clic</strong> para seleccionar</p>
    <p class="drop-types">Formatos soportados: <strong>.csv</strong> · <strong>.json</strong></p>
    <input type="file" id="fileInput" accept=".csv,.json" style="display:none" onchange="handleFileSelect(event)">`;
  const status = document.getElementById('uploadStatus');
  status.style.display = 'none';
  status.className = 'upload-status';
  document.getElementById('btnLoad').disabled = true;
}

// ── Drag & Drop ─────────────────────────────────────────────────
function handleDragOver(e) {
  e.preventDefault();
  document.getElementById('dropZone').classList.add('dragover');
}
function handleDragLeave(e) {
  document.getElementById('dropZone').classList.remove('dragover');
}
function handleDrop(e) {
  e.preventDefault();
  const dz = document.getElementById('dropZone');
  dz.classList.remove('dragover');
  const file = e.dataTransfer.files[0];
  if (file) processFile(file);
}
function handleFileSelect(e) {
  const file = e.target.files[0];
  if (file) processFile(file);
}

// ── Process selected file ───────────────────────────────────────
function processFile(file) {
  const ext = file.name.split('.').pop().toLowerCase();
  if (!['csv', 'json'].includes(ext)) {
    showStatus('❌ Formato no válido. Usa .csv o .json', 'error');
    return;
  }
  _pendingFile = file;
  const dz = document.getElementById('dropZone');
  dz.classList.add('has-file');
  dz.innerHTML = `
    <div class="drop-icon">✅</div>
    <p class="drop-text">Archivo listo</p>
    <p class="drop-filename">📎 ${file.name}</p>
    <p class="drop-sub" style="margin-top:6px">Tamaño: ${(file.size / 1024).toFixed(1)} KB</p>`;
  showStatus(`✔ "${file.name}" seleccionado. Presiona "Cargar Archivo" para actualizar el dashboard.`, 'success');
  const btn = document.getElementById('btnLoad');
  btn.disabled = false;
  btn.onclick = () => loadFileData(file, ext);
}

// ── Parse & Load ────────────────────────────────────────────────
function loadFileData(file, ext) {
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      if (ext === 'json') {
        const parsed = JSON.parse(e.target.result);
        applyData(parsed);
      } else {
        parseCSVAndApply(e.target.result);
      }
    } catch (err) {
      showStatus('❌ Error al parsear el archivo: ' + err.message, 'error');
    }
  };
  reader.readAsText(file, 'UTF-8');
}

// ── CSV Parser ──────────────────────────────────────────────────
function parseCSVAndApply(csvText) {
  // Helper to safely parse numbers from strings with commas, symbols, etc.
  const n = (val) => {
    if (val === undefined || val === null || val === '') return 0;
    const clean = val.toString().replace(/[^-0-9.]/g, '');
    return parseFloat(clean) || 0;
  };

  // Try to find sections, if Excel fused them we will parse row by row grouping headers
  const lines = csvText.trim().split('\n').map(l => l.trim().replace(/\r/g, '')).filter(l => l && !l.startsWith('#'));
  const parsed = {};

  // Group rows by identifying header rows (rows without numbers, or specific known headers)
  let currentHeaders = null;
  let currentRows = [];

  const processGroup = (headers, rows) => {
    if (!headers || headers.length === 0 || rows.length === 0) return;

    // Check for empty CSV trailing lines
    if (headers.length === 1 && headers[0] === "") return;

    // Detect type by columns present using robust exact matches of common patterns
    if (headers.includes('dso_actual') || headers.includes('total_ar')) {
      const r = rows[0] || {};
      parsed.dso = {
        actual: n(r.dso_actual),
        prev: n(r.dso_prev),
        best: n(r.dso_best),
        target: n(r.dso_target)
      };
      if (r.total_ar) parsed.totalAR = n(r.total_ar);
      if (r.collected) parsed.collected = n(r.collected);
    }

    if (headers.includes('aging_bucket') && headers.includes('amount')) {
      parsed.aging = {};
      rows.forEach(r => {
        const b = (r.aging_bucket || r.bucket || '').toLowerCase();
        const v = n(r.amount || r.saldo);
        if (b.includes('corr') || b === 'current') parsed.aging.current = v;
        else if (b.includes('1') || b.includes('30')) parsed.aging.d30 = v;
        else if (b.includes('31') || b.includes('60')) parsed.aging.d60 = v;
        else if (b.includes('61') || b.includes('90')) parsed.aging.d90 = v;
        else if (b.includes('+90') || b.includes('90+')) parsed.aging.d90p = v;
      });
    }

    if (headers.includes('name') && headers.includes('balance') && headers.includes('overdue')) {
      parsed.clients = rows.map(r => ({
        name: r.name,
        balance: n(r.balance),
        overdue: n(r.overdue),
        limitExc: n(r.limitexc || r.limit_exc),
        trend: r.trend || 'stable',
        score: n(r.score || 50),
        seg: r.seg || 'stable'
      }));
    }

    if (headers.includes('client') && headers.includes('week') && headers.includes('amount')) {
      parsed.projection = rows.map(r => ({
        client: r.client,
        week: r.week,
        amount: n(r.amount),
        prob: r.prob || 'med'
      }));
    }

    if (headers.includes('month') && headers.includes('dso')) {
      parsed.months = rows.map(r => r.month);
      parsed.dsoHistory = rows.map(r => n(r.dso));
    }

    if (headers.some(h => h.includes('ca_unapplied')) || headers.some(h => h.includes('ca_automatch'))) {
      const r = rows[0] || {};
      parsed.cashapp = parsed.cashapp || { kpis: {}, suspense: {}, items: [] };
      parsed.cashapp.kpis = {
        unapplied: n(r.ca_unapplied),
        suspense: n(r.ca_suspense),
        autoMatch: n(r.ca_automatch),
        manTime: n(r.ca_mantime)
      };
    }

    if (headers.some(h => h.includes('sus_noref')) || headers.some(h => h.includes('sus_invalidamt'))) {
      const r = rows[0] || {};
      parsed.cashapp = parsed.cashapp || { kpis: {}, suspense: {}, items: [] };
      parsed.cashapp.suspense = {
        noRef: n(r.sus_noref),
        invalidAmt: n(r.sus_invalidamt),
        noClient: n(r.sus_noclient),
        doublePay: n(r.sus_doublepay)
      };
    }

    if (headers.some(h => h.includes('ca_ref')) || headers.some(h => h.includes('ca_status'))) {
      parsed.cashapp = parsed.cashapp || { kpis: {}, suspense: {}, items: [], appliedCashHistory: null };
      parsed.cashapp.items = rows.filter(r => r.ca_ref).map(r => ({
        ref: r.ca_ref,
        amount: n(r.ca_amount),
        date: r.ca_date || '',
        days: n(r.ca_days),
        client: r.ca_client || '',
        status: r.ca_status || ''
      }));
    }

    if (headers.includes('rft_number')) {
      parsed.cashapp = parsed.cashapp || { kpis: {}, suspense: {}, items: [], refunds: { items: [] } };
      parsed.cashapp.refunds = parsed.cashapp.refunds || { items: [] };
      parsed.cashapp.refunds.items = rows.filter(r => r.rft_number).map(r => ({
        rftNumber: r.rft_number,
        created: r.created || '',
        subsidiary: r.subsidiary || '',
        amount: n(r.amount),
        currency: r.currency || '',
        age: n(r.age),
        responsable: r.responsable || '',
        link: r.link || '',
        status: r.status || 'Pendiente'
      }));
    }

    // Section 9: Comparativa Interanual (YoY Cash Applied History)
    const yoyHeader = headers.find(h => h.includes('ca_history') || h === 'month');
    if (yoyHeader) {
      parsed.cashapp = parsed.cashapp || { kpis: {}, suspense: {}, items: [], appliedCashHistory: {} };
      const historyObj = {};
      
      // Identify all year columns (exclude the month column)
      const yearHeaders = headers.filter(h => h !== yoyHeader && h !== '');
      
      yearHeaders.forEach(yh => {
        // Map common headers back to internal keys if needed, or just use the header name
        let key = yh;
        if (yh === 'ca_history_curr') key = 'currentYear';
        if (yh === 'ca_history_prev') key = 'prevYear';
        
        historyObj[key] = rows.map(r => n(r[yh]));
      });
      
      parsed.cashapp.appliedCashHistory = historyObj;
      
      if (!parsed.months) {
        parsed.months = rows.map(r => r[yoyHeader]);
      }
    }
  };

  const isHeaderRow = (cols) => {
    // A row is likely a header if it contains specific known keywords
    return cols.some(c => ['dso_actual', 'aging_bucket', 'name', 'client', 'month', 'ca_unapplied', 'sus_noref', 'ca_ref', 'ca_history_month', 'rft_number'].includes(c));
  };

  for (let i = 0; i < lines.length; i++) {
    const rawCols = lines[i].split(',').map(c => c.trim().toLowerCase());

    if (isHeaderRow(rawCols)) {
      if (currentHeaders) {
        processGroup(currentHeaders, currentRows);
      }
      currentHeaders = rawCols;
      currentRows = [];
    } else if (currentHeaders) {
      const vals = lines[i].split(',').map(v => v.replace(/[\r\n]+/g, '').trim());
      // Skip totally blank rows
      if (vals.some(v => v !== '')) {
        const obj = {};
        currentHeaders.forEach((h, idx) => obj[h] = vals[idx] ? vals[idx] : '');
        currentRows.push(obj);
      }
    }
  }

  // Process last group
  if (currentHeaders && currentRows.length > 0) {
    processGroup(currentHeaders, currentRows);
  }

  applyData(parsed);
}

// ── Apply data to dashboard ─────────────────────────────────────
function applyData(parsed) {
  let changed = 0;

  if (parsed.dso) {
    Object.assign(DATA.dso, parsed.dso);
    changed++;
  }
  if (parsed.aging) {
    Object.assign(DATA.aging, parsed.aging);
    changed++;
  }
  if (parsed.totalAR) { DATA.totalAR = parsed.totalAR; changed++; }
  if (parsed.collected) { DATA.collected = parsed.collected; changed++; }
  if (parsed.clients && parsed.clients.length > 0) {
    DATA.clients = parsed.clients;
    changed++;
  }
  if (parsed.projection && parsed.projection.length > 0) {
    DATA.projection = parsed.projection;
    // Destroy the projection chart and table so they rebuild with fresh data
    if (charts.projection) { charts.projection.destroy(); delete charts.projection; }
    const pt = document.getElementById('projectionTable');
    if (pt) pt.innerHTML = '';
    changed++;
  }
  if (parsed.months && parsed.months.length > 0) {
    DATA.months = parsed.months;
    DATA.dsoHistory = parsed.dsoHistory;
    changed++;
  }
  if (parsed.cashapp) {
    if (parsed.cashapp.kpis && Object.keys(parsed.cashapp.kpis).length > 0) DATA.cashapp.kpis = parsed.cashapp.kpis;
    if (parsed.cashapp.suspense && Object.keys(parsed.cashapp.suspense).length > 0) DATA.cashapp.suspense = parsed.cashapp.suspense;
    if (parsed.cashapp.items && parsed.cashapp.items.length > 0) DATA.cashapp.items = parsed.cashapp.items;
    if (parsed.cashapp.refunds && parsed.cashapp.refunds.items && parsed.cashapp.refunds.items.length > 0) {
      DATA.cashapp.refunds.items = parsed.cashapp.refunds.items;
    }
    // Apply new YoY comparison chart data
    if (parsed.cashapp.appliedCashHistory) {
      DATA.cashapp.appliedCashHistory = parsed.cashapp.appliedCashHistory;
      // Destroy cached chart so it rebuilds with fresh data
      if (charts.caComparison) { charts.caComparison.destroy(); delete charts.caComparison; }
      // Clear delta badges so they regenerate
      const dr = document.getElementById('caCompDeltaRow');
      if (dr) dr.innerHTML = '';
    }
    changed++;
  }

  resyncData();
  closeUploadModal();
  showToast('✅ Dashboard actualizado con datos reales', 'success');
}

// ── EXPORT CSV ─────────────────────────────────────────────────
function exportCSV() {
  const d = DATA;
  let csv = "";

  // 1. Resumen DSO y Totales
  csv += "dso_actual,dso_prev,dso_best,dso_target,total_ar,collected\n";
  csv += `${d.dso.actual},${d.dso.prev},${d.dso.best},${d.dso.target},${d.totalAR},${d.collected}\n\n`;

  // 2. Reporte de Antigüedad (Aging)
  csv += "aging_bucket,amount\n";
  csv += `current,${d.aging.current}\n`;
  csv += `d30,${d.aging.d30}\n`;
  csv += `d60,${d.aging.d60}\n`;
  csv += `d90,${d.aging.d90}\n`;
  csv += `d90p,${d.aging.d90p}\n\n`;

  // 3. Clientes
  csv += "name,balance,overdue,limitExc,trend,score,seg\n";
  d.clients.forEach(c => {
    csv += `${c.name},${c.balance},${c.overdue},${c.limitExc},${c.trend},${c.score},${c.seg}\n`;
  });
  csv += "\n";

  // 4. Proyección
  csv += "client,week,amount,prob\n";
  d.projection.forEach(p => {
    csv += `${p.client},${p.week},${p.amount},${p.prob}\n`;
  });
  csv += "\n";

  // 5. Histórico DSO
  csv += "month,dso\n";
  d.months.forEach((m, i) => {
    csv += `${m},${d.dsoHistory[i] || 0}\n`;
  });
  csv += "\n";

  // 6. Cash App KPIs
  csv += "ca_unapplied,ca_suspense,ca_automatch,ca_mantime\n";
  csv += `${d.cashapp.kpis.unapplied},${d.cashapp.kpis.suspense},${d.cashapp.kpis.autoMatch},${d.cashapp.kpis.manTime}\n\n`;

  // 7. Cash App Suspense Composition
  csv += "sus_noref,sus_invalidamt,sus_noclient,sus_doublepay\n";
  csv += `${d.cashapp.suspense.noRef},${d.cashapp.suspense.invalidAmt},${d.cashapp.suspense.noClient},${d.cashapp.suspense.doublePay}\n\n`;

  // 8. Cash App Items Table
  csv += "ca_ref,ca_amount,ca_date,ca_days,ca_client,ca_status\n";
  d.cashapp.items.forEach(item => {
    csv += `${item.ref},${item.amount},${item.date},${item.days},${item.client},${item.status}\n`;
  });
  csv += "\n";

  // 9. Comparativa Interanual de Efectivo Aplicado (Gráfica YoY)
  if (d.cashapp.appliedCashHistory) {
    const h = d.cashapp.appliedCashHistory;
    const yearKeys = Object.keys(h);
    
    // Header
    csv += "ca_history_month," + yearKeys.join(",") + "\n";
    
    // Rows
    d.months.forEach((m, i) => {
      const vals = yearKeys.map(yk => h[yk][i] || 0);
      csv += `${m},` + vals.join(",") + "\n";
    });
    csv += "\n";
  }

  // 10. Refunds Table
  csv += "rft_number,created,subsidiary,amount,currency,age,responsable,link,status\n";
  d.cashapp.refunds.items.forEach(item => {
    csv += `${item.rftNumber},${item.created},${item.subsidiary},${item.amount},${item.currency},${item.age},${item.responsable},${item.link},${item.status}\n`;
  });
  csv += "\n";

  // Crear y descargar archivo
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.setAttribute("href", url);
  link.setAttribute("download", "ar_datos_exportados.csv");
  link.style.visibility = 'hidden';
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);

  showToast('✅ Archivo CSV descargado correctamente', 'success');
}

// ── Format tabs in modal ─────────────────────────────────────────
function showFmtTab(type) {
  document.getElementById('fmtPanelCSV').style.display = type === 'csv' ? '' : 'none';
  document.getElementById('fmtPanelJSON').style.display = type === 'json' ? '' : 'none';
  document.getElementById('fmtTabCSV').className = 'fmt-tab' + (type === 'csv' ? ' active' : '');
  document.getElementById('fmtTabJSON').className = 'fmt-tab' + (type === 'json' ? ' active' : '');
}

// ── Status msg in modal ──────────────────────────────────────────
function showStatus(msg, type) {
  const el = document.getElementById('uploadStatus');
  el.textContent = msg;
  el.className = 'upload-status ' + type;
  el.style.display = '';
}

// ── Toast notification ───────────────────────────────────────────
let _toastTimer = null;
function showToast(msg, type = '') {
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.className = 'toast show' + (type ? ' ' + type : '');
  clearTimeout(_toastTimer);
  _toastTimer = setTimeout(() => { t.className = 'toast'; }, 4000);
}

// ── ADD CLIENT MODAL ────────────────────────────────────────────
function openAddClientModal() {
  document.getElementById('addClientModal').classList.add('open');
}

function closeAddClientModal() {
  document.getElementById('addClientModal').classList.remove('open');
  document.getElementById('addClientForm').reset();
}

function handleAddClientOverlayClick(e) {
  if (e.target === document.getElementById('addClientModal')) {
    closeAddClientModal();
  }
}

function handleAddClientSubmit(e) {
  e.preventDefault();

  const v = Number(document.getElementById('acBalance').value);
  const o = Number(document.getElementById('acOverdue').value);

  const newClient = {
    name: document.getElementById('acName').value,
    balance: v,
    overdue: o,
    limitExc: Number(document.getElementById('acLimitExc').value),
    trend: document.getElementById('acTrend').value,
    score: Number(document.getElementById('acScore').value),
    seg: document.getElementById('acSeg').value
  };

  DATA.clients.push(newClient);

  if (o <= 0) DATA.aging.current += v;
  else if (o <= 30) DATA.aging.d30 += v;
  else if (o <= 60) DATA.aging.d60 += v;
  else if (o <= 90) DATA.aging.d90 += v;
  else DATA.aging.d90p += v;

  DATA.totalAR += v;

  resyncData();
  closeAddClientModal();
  showToast('✅ Cliente agregado exitosamente', 'success');
}



// ── SINGLE INIT POINT ──────────────────────────────────────────
// All initialization runs ONCE here, after all functions are defined.
(function initDashboard() {
  loadState();       // Merge saved state (with schema version check)
  chartDefaults();   // Configure Chart.js global defaults
  resyncData();      // Sync data → DOM → charts
})();

// ── LANGUAGE TOGGLE (I18N) ──────────────────────────────────────────
// (currentLang declared at top of file as a global)

const esToEn = {
  "Resumen Ejecutivo": "Executive Summary",
  "Cuentas por Cobrar · Análisis Estratégico": "Accounts Receivable · Strategic Analysis",
  "Resumen": "Overview",
  "Riesgo": "Risk",
  "Segmentación": "Segmentation",
  "Proyección": "Projection",
  "Agregar Cliente": "Add Client",
  "Subir Datos Reales": "Upload Real Data",
  "Exportar CSV": "Export CSV",
  "Restablecer Datos": "Reset Data",
  "Cartera Total": "Total Portfolio",
  "DSO Actual": "Current DSO",
  "Recaudado MTD": "Collected MTD",
  "Clientes en Riesgo": "At-Risk Clients",
  "vs mes anterior": "vs previous month",
  "vs objetivo": "vs target",
  "nuevos esta semana": "new this week",
  "Distribución de Cartera por Antigüedad": "Portfolio Aging Distribution",
  "Composición de Riesgo": "Risk Composition",
  "Abril": "April",
  "Objetivo:": "Target:",
  "Arriba del objetivo": "Above Target",
  "Mes Anterior": "Previous Month",
  "Deterioro": "Deterioration",
  "Meta Saludable": "Healthy Goal",
  "Objetivo Anual": "Annual Target",
  "Objetivo Fijo": "Fixed Target",
  "Tendencia DSO (6 Meses)": "DSO Trend (6 Months)",
  "Real vs Meta": "Actual vs Target",
  "Composición del DSO": "DSO Composition",
  "Base vs Retraso": "Base vs Delay",
  "Corriente": "Current",
  "días": "days",
  "dias": "days",
  "Aging Report – Desglose por Antigüedad": "Aging Report – Breakdown by Age",
  "Monto": "Amount",
  "Aging por Cliente": "Aging by Client",
  "Clientes con Mayor Riesgo de Impago": "Clients with Highest Default Risk",
  "Clasificados por score de riesgo compuesto": "Ranked by composite risk score",
  "Cliente": "Client",
  "Saldo Vencido": "Overdue Balance",
  "Días Vencido": "Days Overdue",
  "Límite Excedido": "Limit Exceeded",
  "Tendencia 6M": "6M Trend",
  "Score Riesgo": "Risk Score",
  "Clasificación": "Classification",
  "Crítico": "Critical",
  "Alto": "High",
  "Medio": "Medium",
  "Deteriorando": "Deteriorating",
  "Mejorando": "Improving",
  "Estable": "Stable",
  "No excedido": "Not exceeded",
  "Matriz de Segmentación de Cartera": "Portfolio Segmentation Matrix",
  "Valor del Cliente vs. Riesgo de Cobro": "Client Value vs. Collection Risk",
  "Clientes Estratégicos": "Strategic Clients",
  "Bajo Riesgo": "Low Risk",
  "Alto Valor": "High Value",
  "Clientes en Alerta": "Alert Clients",
  "Alto Riesgo": "High Risk",
  "Bajo Valor": "Low Value",
  "Clientes Estables": "Stable Clients",
  "Proyección 30 días": "30-Day Projection",
  "Confianza": "Confidence",
  "Recaudo Probable": "Probable Collection",
  "Recaudo Posible": "Possible Collection",
  "En Riesgo": "At Risk",
  "Prob. alta": "High prob.",
  "Prob. media": "Med prob.",
  "Proyección Semanal de Recaudos": "Weekly Collection Projection",
  "Próximos 30 Días": "Next 30 Days",
  "Calendario de Vencimientos por Cliente": "Client Maturity Calendar",
  "Fecha Estimada": "Estimated Date",
  "Monto Proyectado": "Projected Amount",
  "Probabilidad": "Probability",
  "Estado": "Status",
  "Comprometido": "Committed",
  "En Negociación": "In Negotiation",
  "Incierto": "Uncertain",
  "Pendiente de aplicar": "Pending application",
  "Partidas sin identificar": "Unidentified items",
  "Eficiencia del sistema": "System efficiency",
  "Por procesar": "To be processed",
  "Análisis Inteligente de Cash Applications": "Intelligent Cash Application Analysis",
  "Tasa de Aplicación Automática Saludable": "Healthy Automatic Application Rate",
  "Oportunidad de Eficiencia": "Efficiency Opportunity",
  "Flujo de Efectivo Retenido": "Retained Cash Flow",
  "Causa Principal de Descuadres": "Main Cause of Mismatches",
  "Aplicación Automática vs Manual": "Automatic vs Manual Application",
  "Últimos 6 Meses": "Last 6 Months",
  "Tendencia de Refunds": "Refunds Trend",
  "Antigüedad de Unapplied Cash": "Aging of Unapplied Cash",
  "Composición de Partidas Sin Identificar": "Composition of Unidentified Items",
  "Top Partidas Pendientes de Aplicar": "Top Pending Items",
  "Detalle de depósitos recibidos que no han podido ser conciliados contra facturas.": "Details of received deposits unable to be reconciled.",
  "Referencia / Banco": "Reference / Bank",
  "Fecha de Depósito": "Deposit Date",
  "Días sin aplicar": "Days unapplied",
  "Posible Origen": "Possible Origin",
  "Estatus": "Status",
  "Análisis de Refunds": "Refunds Analysis",
  "Detalle de las solicitudes de devolución pendientes de validación y pago.": "Details of refund requests pending validation and payment.",
  "Monto Refund": "Refund Amount",
  "Motivo": "Reason",
  "Fecha Solicitud": "Request Date",
  "Cartera Saludable": "Healthy Portfolio",
  "Investigando": "Investigating",
  "Contactado": "Contacted",
  "Pendiente": "Pending",
  "Desconocido": "Unknown",
  "Doble Pago": "Double Payment",
  "Error en Factura": "Invoice Error",
  "Nota de Crédito": "Credit Note",
  "Falta de Referencia": "Missing Reference",
  "Monto Inválido": "Invalid Amount",
  "Cliente No Encontrado": "Client Not Found",
  "Validando": "Validating",
  "Semana": "Week",
  "Últimos": "Last",
  "Tasa de Aplicación Automática Saludable:": "Healthy Automatic Application Rate:",
  "El sistema está emparejando automáticamente el": "The system is automatically matching",
  "de los ingresos, reduciendo significativamente la carga manual.": "of income, significantly reducing manual effort.",
  "Oportunidad de Eficiencia": "Efficiency Opportunity",
  "Aumentar el auto-match reduciría el tiempo extra manual promedio que actualmente requiere": "Increasing auto-match would reduce the average manual overtime currently required",
  "min por partida.": "min per item.",
  "Flujo de Efectivo Retenido:": "Retained Cash Flow:",
  "Actualmente, existen": "Currently, there are",
  "pendientes de aplicar a las cuentas de los clientes. Reducir esta cantidad impactaría positivamente en el flujo de caja inmediato. De este monto total,": "pending application to client accounts. Reducing this amount would positively impact immediate cash flow. Of this total amount,",
  "se encuentran etiquetados como": "are labeled as",
  "Cuenta de Suspenso": "Suspense Account",
  "por estar totalmente sin identificar.": "for being completely unidentified.",
  "Causa Principal de Descuadres:": "Main Cause of Mismatches:",
  "El motivo #1 de partidas sin registrar es por": "The #1 reason for unregistered items is",
  "en la muestra de suspenso": "in the suspense sample",
  "Acción recomendada: Automatizar recordatorios para que los clientes adjunten esta información en sus comprobantes de pago.": "Recommended action: Automate reminders for clients to attach this information to their payment receipts.",
  "Comparativa Interanual de Efectivo Aplicado": "Year-on-Year Applied Cash Comparison",
  "Año Actual (2026)": "Current Year (2026)",
  "Año Anterior (2025)": "Previous Year (2025)",
  "Análisis Interanual": "Year-on-Year Analysis",
  "El efectivo aplicado acumulado en": "The accumulated applied cash in",
  "superando los": "surpassing",
  "por debajo de los": "below",
  "un crecimiento de": "a growth of",
  "un cambio de": "a change of",
  "El mes con mayor mejora fue": "The month with the greatest improvement was",
  "La tendencia está": "The trend is",
  "acelerando": "accelerating",
  "desacelerando": "decelerating",
  "promedio reciente": "recent average",
  "vs inicio": "vs start",
  "Total Aplicado": "Total Applied",
  "Mejor Mes": "Best Month",
  "Alta Probabilidad": "High Probability",
  "Media Probabilidad": "Medium Probability",
  "Baja Probabilidad": "Low Probability",
  "Crítico": "Critical",
  "Alto": "High",
  "Medio": "Medium",
  "Estable": "Stable",
  "Mejorando": "Improving",
  "Deteriorando": "Deteriorating",
  "No excedido": "Not exceeded",
  "días": "days",
  "Actual": "Actual",
  "Anterior": "Previous",
  "Ene": "Jan", "Feb": "Feb", "Mar": "Mar", "Abr": "Apr", "May": "May", "Jun": "Jun",
  "Jul": "Jul", "Ago": "Aug", "Sep": "Sep", "Oct": "Oct", "Nov": "Nov", "Dic": "Dec"
};

const enToEs = Object.fromEntries(Object.entries(esToEn).map(([k,v]) => [v,k]));

const titlesEn = {
  overview: 'Executive Summary', dso: 'DSO Analysis', aging: 'Aging Report',
  risk: 'Risk Analysis', segmentation: 'Portfolio Segmentation', projection: 'Collections Projection',
  cashapp: 'Cash Applications', refunds: 'Refunds'
};
const titlesEs = {
  overview: 'Resumen Ejecutivo', dso: 'Análisis DSO', aging: 'Aging Report',
  risk: 'Análisis de Riesgo', segmentation: 'Segmentación de Cartera', projection: 'Proyección de Recaudos',
  cashapp: 'Cash Applications', refunds: 'Análisis de Refunds'
};

function translateDOMNode(element, dict) {
  const keys = Object.keys(dict).sort((a, b) => b.length - a.length);
  const walker = document.createTreeWalker(element, NodeFilter.SHOW_TEXT, null, false);
  let node;
  while ((node = walker.nextNode())) {
    let text = node.nodeValue;
    if (text.trim() === '') continue;
    
    let changed = false;
    for (const key of keys) {
      // Use word boundaries to avoid replacing substrings inside other words
      const escaped = key.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      const regex = new RegExp('(?<![a-zA-ZáéíóúñÁÉÍÓÚÑ])' + escaped + '(?![a-zA-ZáéíóúñÁÉÍÓÚÑ])', 'gi');
      if (regex.test(text)) {
        text = text.replace(regex, dict[key]);
        changed = true;
      }
    }
    if (changed) node.nodeValue = text;
  }
}

// Custom observer to translate dynamic table rows immediately after they are generated
const observer = new MutationObserver((mutations) => {
  if (currentLang === 'es') return;
  mutations.forEach((mutation) => {
    mutation.addedNodes.forEach((node) => {
      if (node.nodeType === 1) {
        if (node.tagName === 'TR' || node.classList.contains('client-chip') || node.tagName === 'LI') {
          translateDOMNode(node, esToEn);
        } else {
          // Fallback for dynamically appended complex blocks
          const els = node.querySelectorAll ? node.querySelectorAll('tr, .client-chip, li') : [];
          els.forEach(el => translateDOMNode(el, esToEn));
        }
      }
    });
  });
});

// Attach observer to containers
window.addEventListener('load', () => {
    const containersToWatch = ['riskTableBody', 'seg-strategic', 'seg-alert', 'seg-stable', 'seg-lowrisk', 'projectionTable', 'caTableBody', 'refundTableBody', 'ca-insights-list'];
    containersToWatch.forEach(id => {
        const el = document.getElementById(id);
        if(el) observer.observe(el, { childList: true, subtree: true, characterData: true });
    });
});

function toggleLang() {
  currentLang = currentLang === 'es' ? 'en' : 'es';
  document.getElementById('langText').textContent = currentLang.toUpperCase();
  
  const dict = currentLang === 'en' ? esToEn : enToEs;
  translateDOMNode(document.body, dict);

  // Update titles map
  if (currentLang === 'en') {
    Object.assign(titles, titlesEn);
  } else {
    Object.assign(titles, titlesEs);
  }
  const activeTabObj = document.querySelector('.nav-btn.active');
  if (activeTabObj) {
    document.getElementById('page-title').textContent = titles[activeTabObj.dataset.tab];
  }

  updateChartsLang();
}

function updateChartsLang() {
  const isEn = currentLang === 'en';
  const dict = isEn ? esToEn : enToEs;
  
  const translateLabels = (chart) => {
    if (!chart || !chart.data.labels) return;
    chart.data.labels = chart.data.labels.map(l => dict[l] || l);
  };

  if (charts.overviewAging) {
    translateLabels(charts.overviewAging);
    charts.overviewAging.data.datasets[0].label = isEn ? 'Balance' : 'Saldo';
    charts.overviewAging.update();
  }
  if (charts.riskDonut) {
    charts.riskDonut.data.labels = isEn ? ['High Risk', 'Medium Risk', 'Low Risk'] : ['Riesgo Alto', 'Riesgo Medio', 'Riesgo Bajo'];
    charts.riskDonut.update();
  }
  if (charts.dsoTrend) {
    translateLabels(charts.dsoTrend);
    charts.dsoTrend.data.datasets[0].label = isEn ? 'Actual DSO' : 'DSO Real';
    charts.dsoTrend.data.datasets[1].label = isEn ? 'Target' : 'Objetivo';
    charts.dsoTrend.update();
  }
  if (charts.dsoComposition) {
    translateLabels(charts.dsoComposition);
    charts.dsoComposition.data.datasets[0].label = isEn ? 'Credit Terms (Base)' : 'Términos de Crédito (Base)';
    charts.dsoComposition.data.datasets[1].label = isEn ? 'Collection Delay' : 'Retraso en Cobro';
    charts.dsoComposition.update();
  }
  if (charts.agingBar) {
    charts.agingBar.data.labels = isEn ? ['Current', '1-30 days', '31-60 days', '61-90 days', '+90 days'] : ['Corriente', '1–30 días', '31–60 días', '61–90 días', '+90 días'];
    charts.agingBar.data.datasets[0].label = isEn ? 'Balance (USD)' : 'Saldo (USD)';
    charts.agingBar.update();
  }
  if (charts.agingStacked) {
    translateLabels(charts.agingStacked);
    charts.agingStacked.data.datasets[0].label = isEn ? 'Current' : 'Corriente';
    charts.agingStacked.update();
  }
  if (charts.caMatch) {
    translateLabels(charts.caMatch);
    charts.caMatch.data.datasets[0].label = isEn ? 'Automatic (%)' : 'Automático (%)';
    charts.caMatch.data.datasets[1].label = isEn ? 'Manual (%)' : 'Manual (%)';
    charts.caMatch.update();
  }
  if (charts.caAging) {
    charts.caAging.data.labels = isEn ? ['0-3 days', '4-7 days', '8-14 days', '+15 days'] : ['0-3 días', '4-7 días', '8-14 días', '+15 días'];
    charts.caAging.update();
  }
  if (charts.caSuspense) {
    charts.caSuspense.data.labels = isEn ? ['Missing Ref', 'Invalid Amount', 'Client Not Found', 'Double Payment'] : ['Falta Referencia', 'Monto Inválido', 'Cliente No Encontrado', 'Doble Pago'];
    charts.caSuspense.update();
  }
  if (charts.projection) {
    charts.projection.data.labels = isEn ? ['Week 1 (Mar 1-7)', 'Week 2 (Mar 8-14)', 'Week 3 (Mar 15-21)', 'Week 4 (Mar 22-31)'] : ['Semana 1 (Mar 1–7)', 'Semana 2 (Mar 8–14)', 'Semana 3 (Mar 15–21)', 'Semana 4 (Mar 22–31)'];
    charts.projection.data.datasets[0].label = isEn ? 'High Probability' : 'Alta Probabilidad';
    charts.projection.data.datasets[1].label = isEn ? 'Medium Probability' : 'Media Probabilidad';
    charts.projection.data.datasets[2].label = isEn ? 'Low Probability' : 'Baja Probabilidad';
    charts.projection.update();
  }
  if (charts.caComparison) {
    translateLabels(charts.caComparison);
    charts.caComparison.data.datasets[0].label = isEn ? 'Current Year (2026)' : 'Año Actual (2026)';
    charts.caComparison.data.datasets[1].label = isEn ? 'Previous Year (2025)' : 'Año Anterior (2025)';
    charts.caComparison.update();
  }
}
