(() => {
  const isLocalFile = typeof location !== 'undefined' && location.protocol === 'file:';
  if (window.Chart && window.ChartDataLabels) {
    try { Chart.register(ChartDataLabels); } catch (_) {}
  }
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
    kpiUDistricts: document.getElementById('kpi-udistricts'),
    kpiUZones: document.getElementById('kpi-uzones'),
    kpiUUnits: document.getElementById('kpi-uunits'),
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
    sort: { key: 'members', dir: 'desc' },
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

  function loadWorkbookFromCsvText(text) {
    const wb = XLSX.read(text, { type: 'string' });
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
      const dn = normalizeValue(d);
      if (!dn || dn === 'general' || dn === 'others' || dn === 'unknown') return false;
      if (!z || normalizeValue(z) === 'unknown') return false;
      if (!u || normalizeValue(u) === 'unknown') return false;
      const status = normalizeValue(readCell(r, 'membershipStatus'));
      if (status === 'rejected') return false;
      return true;
    });
    State.rawRows = cleaned;
    State.filteredRows = cleaned.slice();
    populateFilters(State.rawRows);
    showToast(`Loaded ${State.rawRows.length} records from ${path}`);
  }

  async function loadFromLocalFile(file) {
    const ext = (file.name.split('.').pop() || '').toLowerCase();
    const rows = await new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onerror = () => reject(new Error('Failed to read file'));
      if (ext === 'csv') {
        reader.onload = () => {
          try { resolve(loadWorkbookFromCsvText(reader.result)); } catch (e) { reject(e); }
        };
        reader.readAsText(file);
      } else {
        reader.onload = () => {
          try { resolve(loadWorkbookFromArrayBuffer(reader.result)); } catch (e) { reject(e); }
        };
        reader.readAsArrayBuffer(file);
      }
    });
    if (!rows.length) throw new Error('No rows found in the Excel file.');
    State.headerMap = mapHeaders(Object.keys(rows[0]));
    const cleaned = rows.filter((r) => {
      const d = readCell(r, 'district');
      const z = readCell(r, 'zone');
      const u = readCell(r, 'unit');
      const dn = normalizeValue(d);
      if (!dn || dn === 'general' || dn === 'others' || dn === 'unknown') return false;
      if (!z || normalizeValue(z) === 'unknown') return false;
      if (!u || normalizeValue(u) === 'unknown') return false;
      const status = normalizeValue(readCell(r, 'membershipStatus'));
      if (status === 'rejected') return false;
      return true;
    });
    State.rawRows = cleaned;
    State.filteredRows = cleaned.slice();
    populateFilters(State.rawRows);
    resetFilters();
    regenerate();
    showToast(`Loaded ${State.rawRows.length} records from ${file.name}`);
  }

  function uniqueSorted(values) {
    return [...new Set(values.filter(Boolean))].sort((a, b) => a.localeCompare(b));
  }

  function populateSelect(select, values, includeAll = false) {
    select.innerHTML = '';
    if (includeAll) {
      const optAll = document.createElement('option');
      optAll.value = '';
      optAll.textContent = 'All';
      select.appendChild(optAll);
    }
    for (const v of values) {
      const opt = document.createElement('option');
      opt.value = v;
      opt.textContent = v;
      select.appendChild(opt);
    }
  }

  function populateFilters(rows) {
    const districts = uniqueSorted(rows.map(r => readCell(r, 'district')).filter(v => v && normalizeValue(v) !== 'unknown'));
    populateSelect(elements.selectDistrict, districts, true);
    // default All selected
    elements.selectDistrict.value = '';
    updateCascadingOptions();
  }

  function updateCascadingOptions() {
    const selectedDistrict = elements.selectDistrict.value;
    const rowsAfterDistrict = selectedDistrict
      ? State.rawRows.filter(r => readCell(r, 'district') === selectedDistrict)
      : State.rawRows;
    const zones = uniqueSorted(rowsAfterDistrict.map(r => readCell(r, 'zone')).filter(v => v && normalizeValue(v) !== 'unknown'));
    const prevZone = elements.selectZone.value;
    populateSelect(elements.selectZone, zones, true);
    elements.selectZone.value = prevZone && zones.includes(prevZone) ? prevZone : '';

    const selectedZone = elements.selectZone.value;
    const rowsAfterZone = selectedZone
      ? rowsAfterDistrict.filter(r => readCell(r, 'zone') === selectedZone)
      : rowsAfterDistrict;
    const units = uniqueSorted(rowsAfterZone.map(r => readCell(r, 'unit')).filter(v => v && normalizeValue(v) !== 'unknown'));
    const prevUnit = elements.selectUnit.value;
    populateSelect(elements.selectUnit, units, true);
    elements.selectUnit.value = prevUnit && units.includes(prevUnit) ? prevUnit : '';
  }

  function getSingleValue(select) { return select.value || ''; }

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
    const dset = new Set();
    const zset = new Set();
    const uset = new Set();
    for (const r of rows) {
      const pay = normalizeValue(readCell(r, 'payment'));
      const sub = normalizeValue(readCell(r, 'submitted'));
      const d = readCell(r, 'district');
      const z = readCell(r, 'zone');
      const u = readCell(r, 'unit');
      const isPaid = pay === 'paid' || pay === 'success' || pay === 'completed' || pay === 'yes';
      const isSubmitted = sub === 'submitted' || sub === 'yes' || sub === 'complete' || sub === 'completed';
      if (isPaid) paid++; else unpaid++;
      if (isSubmitted) submitted++; else pending++;
      if (d && normalizeValue(d) !== 'unknown') dset.add(d);
      if (z && normalizeValue(z) !== 'unknown') zset.add(z);
      if (u && normalizeValue(u) !== 'unknown') uset.add(u);
    }
    return { total, paid, unpaid, submitted, pending, uDistricts: dset.size, uZones: zset.size, uUnits: uset.size };
  }

  function groupCount(rows, key) {
    const m = new Map();
    for (const r of rows) {
      const k = readCell(r, key);
      if (!k || normalizeValue(k) === 'unknown') continue;
      m.set(k, (m.get(k) || 0) + 1);
    }
    return m;
  }

  function computeUnitStats(rows) {
    const m = new Map();
    for (const r of rows) {
      const unit = readCell(r, 'unit');
      const zone = readCell(r, 'zone');
      const district = readCell(r, 'district');
      if (!unit || normalizeValue(unit) === 'unknown') continue;
      if (!zone || normalizeValue(zone) === 'unknown') continue;
      if (!district || normalizeValue(district) === 'unknown') continue;
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
    if (elements.kpiUDistricts) elements.kpiUDistricts.textContent = Intl.NumberFormat().format(kpis.uDistricts || 0);
    if (elements.kpiUZones) elements.kpiUZones.textContent = Intl.NumberFormat().format(kpis.uZones || 0);
    if (elements.kpiUUnits) elements.kpiUUnits.textContent = Intl.NumberFormat().format(kpis.uUnits || 0);
  }

  function aggregateByKey(rows, key) {
    const m = new Map();
    for (const r of rows) {
      const name = readCell(r, key);
      if (!name || normalizeValue(name) === 'unknown') continue;
      const pay = normalizeValue(readCell(r, 'payment'));
      const sub = normalizeValue(readCell(r, 'submitted'));
      if (!m.has(name)) m.set(name, { name, members: 0, paid: 0, unpaid: 0, submitted: 0, pending: 0 });
      const o = m.get(name);
      o.members++;
      const isPaid = pay === 'paid' || pay === 'success' || pay === 'completed' || pay === 'yes';
      const isSubmitted = sub === 'submitted' || sub === 'yes' || sub === 'complete' || sub === 'completed';
      if (isPaid) o.paid++; else o.unpaid++;
      if (isSubmitted) o.submitted++; else o.pending++;
    }
    return Array.from(m.values()).sort((a, b) => b.members - a.members);
  }

  function renderTopDistricts(rows) {
    const topN = Math.max(1, Number((document.getElementById('input-top-districts') || {}).value) || 20);
    const stats = aggregateByKey(rows, 'district').slice(0, topN);
    const tbody = (document.querySelector('#table-districts tbody'));
    if (!tbody) return;
    tbody.innerHTML = '';
    for (const s of stats) {
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td>${escapeHtml(s.name)}</td>
        <td>${s.members}</td>
        <td>${s.paid}</td>
        <td>${s.unpaid}</td>
        <td>${s.submitted}</td>
        <td>${s.pending}</td>
      `;
      tbody.appendChild(tr);
    }
  }

  function renderTopZones(rows) {
    const topN = Math.max(1, Number((document.getElementById('input-top-zones') || {}).value) || 20);
    // aggregate by Zone with District context
    const m = new Map(); // key: `${district}||${zone}`
    for (const r of rows) {
      const district = readCell(r, 'district');
      const zone = readCell(r, 'zone');
      if (!district || normalizeValue(district) === 'unknown') continue;
      if (!zone || normalizeValue(zone) === 'unknown') continue;
      const key = district + '||' + zone;
      const pay = normalizeValue(readCell(r, 'payment'));
      const sub = normalizeValue(readCell(r, 'submitted'));
      if (!m.has(key)) m.set(key, { district, zone, members: 0, paid: 0, unpaid: 0, submitted: 0, pending: 0 });
      const o = m.get(key);
      o.members++;
      const isPaid = pay === 'paid' || pay === 'success' || pay === 'completed' || pay === 'yes';
      const isSubmitted = sub === 'submitted' || sub === 'yes' || sub === 'complete' || sub === 'completed';
      if (isPaid) o.paid++; else o.unpaid++;
      if (isSubmitted) o.submitted++; else o.pending++;
    }
    const stats = Array.from(m.values()).sort((a, b) => b.members - a.members).slice(0, topN);
    const tbody = (document.querySelector('#table-zones tbody'));
    if (!tbody) return;
    tbody.innerHTML = '';
    for (const s of stats) {
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td>${escapeHtml(s.district)}</td>
        <td>${escapeHtml(s.zone)}</td>
        <td>${s.members}</td>
        <td>${s.paid}</td>
        <td>${s.unpaid}</td>
        <td>${s.submitted}</td>
        <td>${s.pending}</td>
      `;
      tbody.appendChild(tr);
    }
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
    if (id === 'chart-age-group') return 'ageGroup';
    if (id === 'chart-profession') return 'profession';
    if (id === 'chart-qualification') return 'qualification';
    if (id === 'chart-top-zones') return 'topZones';
    if (id === 'chart-top-units') return 'topUnits';
    return id;
  }

  /* Added for adjusting bar chart height based on label count */
  function setBarChartHeight(canvasId, labelCount, min=320, perLabel=40, max=1200) {
    const canvas = document.getElementById(canvasId);
    if (!canvas) return;
    // force vertical height for horizontal bar charts
    const px = Math.max(min, Math.min(max, labelCount * perLabel));
    canvas.style.height = px + 'px';
  }
/* End of Added for adjusting bar chart height based on label count */

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

    // Age Group Pie:
  const dobKey = getColNameMatch(Object.keys(rows[0]||{}), ['date of birth','dob','d.o.b','birth date','birthdate']);
  if (dobKey) {
    const ag = parseAgeGroups(rows, dobKey);
      ensureChart('chart-age-group', 'pie', {
        labels: Object.keys(ag),
      datasets:[{data:Object.values(ag), backgroundColor: chartColors(5)}]
    }, {plugins:{legend:{labels:{color:'#e8ecf8'}}, datalabels:{color:'#e8ecf8', formatter:(v)=>v}}});
    }
    // Profession:
  const professionKey = getColNameMatch(
    Object.keys(rows[0]||{}),
    ['occupation','profession','working as','job','employment','profession/occupation']
  );
    if (professionKey) {
      const top10 = groupBarData(rows, professionKey, 10);
      ensureChart('chart-profession','bar',{
        labels: top10.map(([k])=>k),
      datasets:[{data:top10.map(([,v])=>v), backgroundColor:chartColors(top10.length), label:'Count'}]
    }, {indexAxis: 'y',plugins:{legend:{display:false}, datalabels:{align:'end', anchor:'end', color:'#e8ecf8', formatter:(v)=>v}},scales:{x:{ticks:{color:'#9aa3ba'}},y:{ticks:{color:'#9aa3ba'}}}});
    }
    // Qualification:
    const qualificationKey = getColNameMatch(Object.keys(rows[0]||{}), ['qualification','education','educational qualification','degree','quali']);
    if (qualificationKey) {
      const top10 = groupBarData(rows, qualificationKey, 10);
      ensureChart('chart-qualification','bar',{
        labels: top10.map(([k])=>k),
      datasets:[{data:top10.map(([,v])=>v), backgroundColor:chartColors(top10.length), label:'Count'}]
    }, {indexAxis: 'y',plugins:{legend:{display:false}, datalabels:{align:'end', anchor:'end', color:'#e8ecf8', formatter:(v)=>v}},scales:{x:{ticks:{color:'#9aa3ba'}},y:{ticks:{color:'#9aa3ba'}}}});
    }
    // Top 10 Zones:
    const zoneKey = getColNameMatch(Object.keys(rows[0]||{}), ['zone','area','region']);
    if (zoneKey) {
    const zoneCounts = groupBarData(rows, zoneKey, 10);
      ensureChart('chart-top-zones', 'bar', {
        labels: zoneCounts.map(([k])=>k),
        datasets: [{data: zoneCounts.map(([,v])=>v), backgroundColor: chartColors(zoneCounts.length), label:'Members'}]
      }, {
        indexAxis: 'y',
        plugins: {
        legend:{display:false},
        datalabels:{anchor:'end',align:'end', color:'#e8ecf8', formatter:(v)=>v}
        },
        scales:{x:{ticks:{color:'#9aa3ba'}}, y:{ticks:{color:'#9aa3ba'}}}
      });
    }
    // Top 10 Units:
    const unitKey = getColNameMatch(Object.keys(rows[0]||{}), ['unit','branch','unit name']);
    if (unitKey) {
    const unitCounts = groupBarData(rows, unitKey, 10);
      ensureChart('chart-top-units', 'bar', {
        labels: unitCounts.map(([k])=>k),
        datasets: [{data: unitCounts.map(([,v])=>v), backgroundColor: chartColors(unitCounts.length), label:'Members'}]
      }, {
        indexAxis: 'y',
        plugins: {
        legend:{display:false},
        datalabels:{anchor:'end',align:'end', color:'#e8ecf8', formatter:(v)=>v}
        },
        scales:{x:{ticks:{color:'#9aa3ba'}}, y:{ticks:{color:'#9aa3ba'}}}
      });
    }
  }

  function renderTable(rows) {
    const topN = Math.max(1, Number(elements.inputTopN.value) || 10);
    const statsAll = computeUnitStats(rows);
    const stats = sortStats(statsAll, State.sort).slice(0, topN);
    const tbody = elements.tableBody;
    tbody.innerHTML = '';
    for (const s of stats) {
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td>${escapeHtml(s.district)}</td>        
        <td>${escapeHtml(s.zone)}</td>
        <td>${escapeHtml(s.unit)}</td>  
        <td>${s.members}</td>
        <td>${s.paid}</td>
        <td>${s.unpaid}</td>
        <td>${s.submitted}</td>
        <td>${s.pending}</td>
      `;
      tbody.appendChild(tr);
    }
    applyHeaderSortClasses();
  }

  function sortStats(list, sort) {
    const { key, dir } = sort;
    const factor = dir === 'asc' ? 1 : -1;
    return list.slice().sort((a, b) => {
      const av = a[key];
      const bv = b[key];
      if (typeof av === 'number' && typeof bv === 'number') return (av - bv) * factor;
      return String(av).localeCompare(String(bv)) * factor;
    });
  }

  function applyHeaderSortClasses() {
    const ths = document.querySelectorAll('#table-units thead th.sortable');
    ths.forEach(th => {
      th.classList.remove('sort-asc', 'sort-desc');
      const key = th.getAttribute('data-key');
      if (key === State.sort.key) th.classList.add(State.sort.dir === 'asc' ? 'sort-asc' : 'sort-desc');
    });
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
    State.filters.districts = getSingleValue(elements.selectDistrict) ? [getSingleValue(elements.selectDistrict)] : [];
    State.filters.zones = getSingleValue(elements.selectZone) ? [getSingleValue(elements.selectZone)] : [];
    State.filters.units = getSingleValue(elements.selectUnit) ? [getSingleValue(elements.selectUnit)] : [];
    State.filters.payment = elements.selectPayment.value;
    State.filters.submitted = elements.selectSubmitted.value;
    State.filters.unitCountOp = elements.selectCountOperator.value;
    State.filters.unitCountVal = elements.inputCount.value;
  }

  function regenerate() {
    applyFilters();
    renderCharts(State.filteredRows);
    renderTable(State.filteredRows);
    renderTopDistricts(State.filteredRows);
    renderTopZones(State.filteredRows);
  }

  function bindEvents() {
    elements.btnFilters.addEventListener('click', openDrawer);
    elements.btnCloseDrawer.addEventListener('click', closeDrawer);
    elements.btnReset.addEventListener('click', () => { resetFilters(); regenerate(); showToast('Filters reset'); });
    elements.btnApply.addEventListener('click', () => { captureFiltersFromUI(); regenerate(); closeDrawer(); showToast('Filters applied'); });
    elements.btnRefresh.addEventListener('click', async () => {
      // clear in-memory data and reload Excel
      State.rawRows = [];
      State.filteredRows = [];
      location.reload(); // ensures we re-fetch from server, not stale data
    });
    elements.inputTopN.addEventListener('change', () => renderTable(State.filteredRows));
    const topD = document.getElementById('input-top-districts');
    const topZ = document.getElementById('input-top-zones');
    if (topD) topD.addEventListener('change', () => renderTopDistricts(State.filteredRows));
    if (topZ) topZ.addEventListener('change', () => renderTopZones(State.filteredRows));
    const ths = document.querySelectorAll('#table-units thead th.sortable');
    ths.forEach(th => {
      th.addEventListener('click', () => {
        const key = th.getAttribute('data-key');
        if (State.sort.key === key) {
          State.sort.dir = State.sort.dir === 'asc' ? 'desc' : 'asc';
        } else {
          State.sort.key = key;
          State.sort.dir = isNumericKey(key) ? 'desc' : 'asc';
        }
        renderTable(State.filteredRows);
      });
    });
    // cascade selects
    elements.selectDistrict.addEventListener('change', () => { updateCascadingOptions(); });
    elements.selectZone.addEventListener('change', () => { updateCascadingOptions(); });
  }

  function isNumericKey(key) {
    return ['members', 'paid', 'unpaid', 'submitted', 'pending'].includes(key);
  }

  // --- UTILS ---
function excelSerialToDate(serial) {
  const utcDays = Math.floor(serial - 25569);
  const utcValue = utcDays * 86400; // seconds
  const dateInfo = new Date(utcValue * 1000);
  return dateInfo;
}
function toDate(value) {
  if (value instanceof Date) return value;
  if (typeof value === 'number' && value > 1000 && value < 60000) {
    return excelSerialToDate(value);
  }
  if (typeof value === 'string') {
    const s = value.trim();
    // Try common dd/mm/yyyy, yyyy-mm-dd
    const m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
    if (m) {
      const d = parseInt(m[1], 10);
      const mo = parseInt(m[2], 10) - 1;
      let y = parseInt(m[3], 10);
      if (y < 100) y += 2000;
      return new Date(y, mo, d);
    }
    const dt = new Date(s);
    if (!isNaN(dt.getTime())) return dt;
  }
  return null;
}
function parseAgeGroups(rows, dobKey) {
    const ageGroups = {
      '15-22': 0,
      '22-30': 0,
      '30-40': 0,
      '40-45': 0,
      'Unknown': 0
    };
    for (const r of rows) {
    const dobVal = r[dobKey];
    const dt = toDate(dobVal);
    let age = NaN;
    if (dt) {
      const today = new Date();
      age = today.getFullYear() - dt.getFullYear();
      const m = today.getMonth() - dt.getMonth();
      if (m < 0 || (m === 0 && today.getDate() < dt.getDate())) age--;
    }
      if (isNaN(age) || age < 10) { ageGroups['Unknown']++; continue; }
      if (age >= 15 && age <= 22) ageGroups['15-22']++;
      else if (age > 22 && age <= 30) ageGroups['22-30']++;
      else if (age > 30 && age <= 40) ageGroups['30-40']++;
      else if (age > 40 && age <= 45) ageGroups['40-45']++;
      else ageGroups['Unknown']++;
    }
    return ageGroups;
  }
function getColNameMatch(headers, wanted) {
  // Returns first column matching any in wanted array (exact or partial match after normalization)
  for (let h of headers) {
    const nh = normalizeHeader(h);
    for (let w of wanted) {
      const nw = normalizeHeader(w);
      if (nh === nw || nh.includes(nw) || nw.includes(nh)) return h;
    }
  }
  return null;
}
  function groupBarData(rows, key, topN = 999) {
    const count = {};
    for (const r of rows) {
      const val = r[key];
      if (!val || normalizeValue(val) === 'unknown') continue;
      count[val] = (count[val] || 0) + 1;
    }
    const groupArr = Object.entries(count).sort((a,b) => b[1] - a[1]);
    return groupArr.slice(0, topN);
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


