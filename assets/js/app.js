(() => {
  const toastEl = document.getElementById('toast');
  function showToast(message) {
    if (!toastEl) return;
    toastEl.textContent = message;
    toastEl.classList.add('show');
    setTimeout(() => toastEl.classList.remove('show'), 2400);
  }

  const elements = {
    btnFilters: document.getElementById('btn-filters'),
    btnCloseDrawer: document.getElementById('btn-close-drawer'),
    btnApply: document.getElementById('btn-apply'),
    btnReset: document.getElementById('btn-reset'),
    btnRefresh: document.getElementById('btn-refresh'),
    drawer: document.getElementById('filter-drawer'),
    selectDistrict: document.getElementById('select-district'),
    selectZone: document.getElementById('select-zone'),
    selectUnit: document.getElementById('select-unit'),
    selectPayment: document.getElementById('select-payment'),
    selectSubmitted: document.getElementById('select-submitted'),
    selectCountOperator: document.getElementById('select-count-operator'),
    inputCount: document.getElementById('input-count'),
    inputTopN: document.getElementById('input-top-n'),
    kpiTotal: document.getElementById('kpi-total'),
    kpiPaid: document.getElementById('kpi-paid'),
    kpiUnpaid: document.getElementById('kpi-unpaid'),
    kpiSubmitted: document.getElementById('kpi-submitted'),
    kpiPending: document.getElementById('kpi-pending'),
    tableBody: document.querySelector('#table-units tbody')
  };

  const charts = {
    payment: null,
    submit: null,
    byDistrict: null,
    byZone: null
  };

  const State = {
    rawRows: [],
    filteredRows: [],
    filters: {
      districts: [],
      zones: [],
      units: [],
      payment: '',
      submitted: '',
      unitCountOp: '',
      unitCountVal: ''
    },
    headerMap: {
      district: null,
      zone: null,
      unit: null,
      payment: null,
      submitted: null,
      membershipStatus: null
    }
  };

  function normalizeHeader(h) {
    return String(h || '')
      .trim()
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  function mapHeaders(headers) {
    const target = {
      district: ['district', 'dist', 'district name'],
      zone: ['zone', 'area', 'region'],
      unit: ['unit', 'branch', 'unit name'],
      payment: ['payment status', 'payment', 'paid status', 'fee status'],
      submitted: ['submitted status', 'submission status', 'submitted', 'form status'],
      membershipStatus: ['membership status', 'status']
    };
    const normalized = headers.map(normalizeHeader);
    const map = {};
    for (const key of Object.keys(target)) {
      let idx = -1;
      for (const alias of target[key]) {
        idx = normalized.indexOf(alias);
        if (idx !== -1) break;
      }
      map[key] = idx === -1 ? null : headers[idx];
    }
    return map;
  }

  function readCell(row, key) {
    const header = State.headerMap[key];
    if (!header) return '';
    const value = row[header];
    return value == null ? '' : String(value).trim();
  }

  function normalizeValue(value) {
    return String(value || '')
      .trim()
      .toLowerCase();
  }

  function loadWorkbookFromArrayBuffer(buf) {
    const wb = XLSX.read(buf, { type: 'array' });
    const firstSheetName = wb.SheetNames[0];
    const ws = wb.Sheets[firstSheetName];
    const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
    return rows;
  }

  async function fetchFirstAvailable(paths) {
    for (const p of paths) {
      try {
        const res = await fetch(p);
        if (!res.ok) continue;
        const ab = await res.arrayBuffer();
        return { path: p, buffer: ab };
      } catch (_) {
        // try next
      }
    }
    throw new Error('Excel file not found. Place it under assets/data/.');
  }

  async function loadData() {
    const paths = (window.__EXCEL_PATHS__ || []);
    const { path, buffer } = await fetchFirstAvailable(paths);
    const rows = loadWorkbookFromArrayBuffer(buffer);
    if (!rows.length) throw new Error('No rows found in the Excel sheet.');
    State.headerMap = mapHeaders(Object.keys(rows[0]));
    const cleaned = rows.filter((r) => {
      const d = readCell(r, 'district');
      const z = readCell(r, 'zone');
      const u = readCell(r, 'unit');
      if (!d && !z && !u) return false;
      const dn = normalizeValue(d);
      if (!dn || dn === 'general' || dn === 'others') return false;
      const status = normalizeValue(readCell(r, 'membershipStatus'));
      if (status === 'rejected') return false;
      return true;
    });
    State.rawRows = cleaned;
    State.filteredRows = cleaned.slice();
    populateFilters(State.rawRows);
    showToast(`Loaded ${State.rawRows.length} records from ${path}`);
  }

  function uniqueSorted(values) {
    return [...new Set(values.filter(Boolean))].sort((a, b) => a.localeCompare(b));
  }

  function populateSelect(select, values) {
    select.innerHTML = '';
    for (const v of values) {
      const opt = document.createElement('option');
      opt.value = v;
      opt.textContent = v;
      select.appendChild(opt);
    }
  }

  function populateFilters(rows) {
    const districts = uniqueSorted(rows.map(r => readCell(r, 'district')));
    const zones = uniqueSorted(rows.map(r => readCell(r, 'zone')));
    const units = uniqueSorted(rows.map(r => readCell(r, 'unit')));
    populateSelect(elements.selectDistrict, districts);
    populateSelect(elements.selectZone, zones);
    populateSelect(elements.selectUnit, units);
  }

  function getMultiSelectValues(select) {
    return Array.from(select.selectedOptions).map(o => o.value);
  }

  function applyFilters() {
    const f = State.filters;
    const rows = State.rawRows.filter((r) => {
      const d = readCell(r, 'district');
      const z = readCell(r, 'zone');
      const u = readCell(r, 'unit');
      const pay = normalizeValue(readCell(r, 'payment'));
      const sub = normalizeValue(readCell(r, 'submitted'));

      if (f.districts.length && !f.districts.includes(d)) return false;
      if (f.zones.length && !f.zones.includes(z)) return false;
      if (f.units.length && !f.units.includes(u)) return false;

      if (f.payment) {
        const isPaid = pay === 'paid' || pay === 'success' || pay === 'completed' || pay === 'yes';
        if (f.payment === 'paid' && !isPaid) return false;
        if (f.payment === 'unpaid' && isPaid) return false;
      }

      if (f.submitted) {
        const isSubmitted = sub === 'submitted' || sub === 'yes' || sub === 'complete' || sub === 'completed';
        if (f.submitted === 'submitted' && !isSubmitted) return false;
        if (f.submitted === 'pending' && isSubmitted) return false;
      }
      return true;
    });

    if (f.unitCountOp && f.unitCountVal !== '' && !isNaN(Number(f.unitCountVal))) {
      const target = Number(f.unitCountVal);
      const unitToCount = new Map();
      for (const r of rows) {
        const u = readCell(r, 'unit');
        unitToCount.set(u, (unitToCount.get(u) || 0) + 1);
      }
      const check = (cnt) => {
        switch (f.unitCountOp) {
          case '>': return cnt > target;
          case '<': return cnt < target;
          case '>=': return cnt >= target;
          case '<=': return cnt <= target;
          case '=': return cnt === target;
          default: return true;
        }
      };
      State.filteredRows = rows.filter(r => check(unitToCount.get(readCell(r, 'unit')) || 0));
    } else {
      State.filteredRows = rows;
    }
  }

  function computeKPIs(rows) {
    const total = rows.length;
    let paid = 0, unpaid = 0, submitted = 0, pending = 0;
    for (const r of rows) {
      const pay = normalizeValue(readCell(r, 'payment'));
      const sub = normalizeValue(readCell(r, 'submitted'));
      const isPaid = pay === 'paid' || pay === 'success' || pay === 'completed' || pay === 'yes';
      const isSubmitted = sub === 'submitted' || sub === 'yes' || sub === 'complete' || sub === 'completed';
      if (isPaid) paid++; else unpaid++;
      if (isSubmitted) submitted++; else pending++;
    }
    return { total, paid, unpaid, submitted, pending };
  }

  function groupCount(rows, key) {
    const m = new Map();
    for (const r of rows) {
      const k = readCell(r, key) || 'Unknown';
      m.set(k, (m.get(k) || 0) + 1);
    }
    return m;
  }

  function computeUnitStats(rows) {
    const m = new Map();
    for (const r of rows) {
      const unit = readCell(r, 'unit') || 'Unknown';
      const zone = readCell(r, 'zone') || '';
      const district = readCell(r, 'district') || '';
      const pay = normalizeValue(readCell(r, 'payment'));
      const sub = normalizeValue(readCell(r, 'submitted'));
      if (!m.has(unit)) m.set(unit, { unit, zone, district, members: 0, paid: 0, unpaid: 0, submitted: 0, pending: 0 });
      const o = m.get(unit);
      o.members++;
      const isPaid = pay === 'paid' || pay === 'success' || pay === 'completed' || pay === 'yes';
      const isSubmitted = sub === 'submitted' || sub === 'yes' || sub === 'complete' || sub === 'completed';
      if (isPaid) o.paid++; else o.unpaid++;
      if (isSubmitted) o.submitted++; else o.pending++;
    }
    return Array.from(m.values()).sort((a, b) => b.members - a.members);
  }

  function updateKPIs(kpis) {
    elements.kpiTotal.textContent = Intl.NumberFormat().format(kpis.total);
    elements.kpiPaid.textContent = Intl.NumberFormat().format(kpis.paid);
    elements.kpiUnpaid.textContent = Intl.NumberFormat().format(kpis.unpaid);
    elements.kpiSubmitted.textContent = Intl.NumberFormat().format(kpis.submitted);
    elements.kpiPending.textContent = Intl.NumberFormat().format(kpis.pending);
  }

  function chartColors(n) {
    const base = ['#4f7cff', '#3ccf91', '#ffcc66', '#ff6b6b', '#a78bfa', '#22d3ee', '#f472b6', '#f59e0b'];
    const arr = [];
    for (let i = 0; i < n; i++) arr.push(base[i % base.length]);
    return arr;
  }

  function ensureChart(ctxId, type, data, options) {
    const ctx = document.getElementById(ctxId);
    if (!ctx) return null;
    if (charts[ctxIdToKey(ctxId)]) charts[ctxIdToKey(ctxId)].destroy();
    const chart = new Chart(ctx, { type, data, options });
    charts[ctxIdToKey(ctxId)] = chart;
    return chart;
  }

  function ctxIdToKey(id) {
    if (id === 'chart-payment') return 'payment';
    if (id === 'chart-submit') return 'submit';
    if (id === 'chart-by-district') return 'byDistrict';
    if (id === 'chart-by-zone') return 'byZone';
    return id;
  }

  function renderCharts(rows) {
    const kpis = computeKPIs(rows);
    updateKPIs(kpis);

    ensureChart('chart-payment', 'doughnut', {
      labels: ['Paid', 'Unpaid'],
      datasets: [{ data: [kpis.paid, kpis.unpaid], backgroundColor: ['#3ccf91', '#ff6b6b'] }]
    }, { plugins: { legend: { labels: { color: '#e8ecf8' } } } });

    ensureChart('chart-submit', 'doughnut', {
      labels: ['Submitted', 'Pending'],
      datasets: [{ data: [kpis.submitted, kpis.pending], backgroundColor: ['#4f7cff', '#ffcc66'] }]
    }, { plugins: { legend: { labels: { color: '#e8ecf8' } } } });

    const byDistrict = groupCount(rows, 'district');
    const dLabels = Array.from(byDistrict.keys());
    const dValues = Array.from(byDistrict.values());
    ensureChart('chart-by-district', 'bar', {
      labels: dLabels,
      datasets: [{ label: 'Members', data: dValues, backgroundColor: chartColors(dLabels.length) }]
    }, {
      scales: {
        x: { ticks: { color: '#9aa3ba' } },
        y: { ticks: { color: '#9aa3ba' }, beginAtZero: true }
      },
      plugins: { legend: { labels: { color: '#e8ecf8' } } }
    });

    const byZone = groupCount(rows, 'zone');
    const zLabels = Array.from(byZone.keys());
    const zValues = Array.from(byZone.values());
    ensureChart('chart-by-zone', 'bar', {
      labels: zLabels,
      datasets: [{ label: 'Members', data: zValues, backgroundColor: chartColors(zLabels.length) }]
    }, {
      scales: {
        x: { ticks: { color: '#9aa3ba' } },
        y: { ticks: { color: '#9aa3ba' }, beginAtZero: true }
      },
      plugins: { legend: { labels: { color: '#e8ecf8' } } }
    });
  }

  function renderTable(rows) {
    const topN = Math.max(1, Number(elements.inputTopN.value) || 10);
    const stats = computeUnitStats(rows).slice(0, topN);
    const tbody = elements.tableBody;
    tbody.innerHTML = '';
    for (const s of stats) {
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td>${escapeHtml(s.unit)}</td>
        <td>${escapeHtml(s.zone)}</td>
        <td>${escapeHtml(s.district)}</td>
        <td>${s.members}</td>
        <td>${s.paid}</td>
        <td>${s.unpaid}</td>
        <td>${s.submitted}</td>
        <td>${s.pending}</td>
      `;
      tbody.appendChild(tr);
    }
  }

  function escapeHtml(str) {
    return String(str).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  }

  function openDrawer() { elements.drawer.classList.add('open'); }
  function closeDrawer() { elements.drawer.classList.remove('open'); }

  function resetFilters() {
    State.filters = { districts: [], zones: [], units: [], payment: '', submitted: '', unitCountOp: '', unitCountVal: '' };
    elements.selectDistrict.selectedIndex = -1;
    elements.selectZone.selectedIndex = -1;
    elements.selectUnit.selectedIndex = -1;
    elements.selectPayment.value = '';
    elements.selectSubmitted.value = '';
    elements.selectCountOperator.value = '';
    elements.inputCount.value = '';
  }

  function captureFiltersFromUI() {
    State.filters.districts = getMultiSelectValues(elements.selectDistrict);
    State.filters.zones = getMultiSelectValues(elements.selectZone);
    State.filters.units = getMultiSelectValues(elements.selectUnit);
    State.filters.payment = elements.selectPayment.value;
    State.filters.submitted = elements.selectSubmitted.value;
    State.filters.unitCountOp = elements.selectCountOperator.value;
    State.filters.unitCountVal = elements.inputCount.value;
  }

  function regenerate() {
    applyFilters();
    renderCharts(State.filteredRows);
    renderTable(State.filteredRows);
  }

  function bindEvents() {
    elements.btnFilters.addEventListener('click', openDrawer);
    elements.btnCloseDrawer.addEventListener('click', closeDrawer);
    elements.btnReset.addEventListener('click', () => { resetFilters(); regenerate(); showToast('Filters reset'); });
    elements.btnApply.addEventListener('click', () => { captureFiltersFromUI(); regenerate(); closeDrawer(); showToast('Filters applied'); });
    elements.btnRefresh.addEventListener('click', async () => { try { await init(true); } catch (e) { showToast(String(e.message || e)); } });
    elements.inputTopN.addEventListener('change', () => renderTable(State.filteredRows));
  }

  async function init(forceReload = false) {
    try {
      if (forceReload || State.rawRows.length === 0) {
        await loadData();
        resetFilters();
      }
      regenerate();
    } catch (e) {
      console.error(e);
      showToast(e.message || 'Failed to load data');
    }
  }

  bindEvents();
  init();
})();


