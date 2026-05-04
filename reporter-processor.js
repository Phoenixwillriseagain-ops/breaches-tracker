// reporter-processor.js - Reporter page: file handling, processing, filtering, KPIs

(function() {
  'use strict';

  const { REPORTER_TABS, REPORTER_STORAGE_KEY, CONFIG } = window.BT || {};
  const { MAX_DISPLAY_ROWS } = CONFIG;

  const rptState = window.RPT = {
    allData: [],
    filtered: [],
    activeTab: 'overview',
    charts: {},
    loadedMonths: [],
    activeMonth: null,
  };

  function debounce(fn, delay) {
    let timer;
    return function(...args) {
      clearTimeout(timer);
      timer = setTimeout(() => fn.apply(this, args), delay);
    };
  }

  function showStatus(msg, type) {
    const el = document.getElementById('status');
    if (!el) return;
    el.textContent = msg;
    el.className = 'status-bar ' + (type === 'error' ? 'error' : '');
    if (type === 'loading') el.className = 'status-bar loading';
  }

  function showLoading(show, msg) {
    const ov = document.getElementById('loading-overlay');
    if (show) {
      ov = document.createElement('div');
      ov.id = 'loading-overlay';
      ov.className = 'loading-overlay';
      ov.innerHTML = '<div class="spinner"></div><div class="loading-text"> + (msg || 'Processing...')</div>';
      document.body.appendChild(ov);
    } else if (ov) {
      const text = document.querySelector('.loading-text');
      if (text) text.textContent = msg || 'Processing...';
    }
  }

  // === Normalization helpers ===
  function clean(v) {
    return String(v == null ? '' : v).trim();
  }

  function normKey(v, fallback) {
    var s = clean(v);
    return s ? s : fallback;
  }

  function normBoolLike(v) {
    var s = clean(v).toLowerCase();
    if (!s) return '0';
    if (['1', 'true', 'yes', 'y', 'excluded'].indexOf(s) !== -1) return '1';
    return '0';
  }

  function normSla(v) {
    var s = clean(v).toLowerCase().replace(/\s+/g, '');
    if (s === 'kmi' || s === 'km-i') return 'KM-I';
    return clean(v).toUpperCase();
  }

  function normAos(v) {
    var s = clean(v).toLowerCase();
    if (!s) return 'N';
    if (['1', 'true', 'yes', 'y', 'excluded'].indexOf(s) !== -1) return 'Y';
    return 'N';
  }

  // Flexible field getter with multiple fallbacks
  function fget(row, fields) {
    if (!Array.isArray(fields)) {
      fields = [fields];
    }
    
    // Try exact match first
    for (let i = 0; i < fields.length; i++) {
      const f = fields[i];
      if (row.hasOwnProperty(f) && row[f] !== undefined && row[f] !== null && row[f] !== '') {
        return String(row[f]);
      }
    }
    
    // Try case-insensitive match
    const rowKeys = Object.keys(row);
    for (let i = 0; i < fields.length; i++) {
      const searchField = clean(fields[i]).toLowerCase();
      for (let j = 0; j < rowKeys.length; j++) {
        const key = rowKeys[j];
        if (clean(key).toLowerCase() === searchField && row[key] !== undefined && row[key] !== null && row[key] !== '') {
          return String(row[key]);
        }
      }
    }
    
    // Try partial match
    const rowKeys2 = Object.keys(row);
    for (let i = 0; i < fields.length; i++) {
      const searchField = clean(fields[i]).toLowerCase();
      for (let j = 0; j < rowKeys2.length; j++) {
        const key = rowKeys2[j];
        if (clean(key).toLowerCase().includes(searchField) && row[key] !== undefined && row[key] !== null && row[key] !== '') {
          return String(row[key]);
        }
      }
    }
    
    return '';
  }

  function loadFile(file) {
    if (!file) return;
    showLoading(true, 'Reading file...');
    const reader = new FileReader();
    reader.onload = function(e) {
      showLoading(true, 'Processing data...');
      setTimeout(function() {
        try {
          const wb = XLSX.read(e.target.result, { type: 'array', cellDates: true });
          processWorkbook(wb);
        } catch (err) {
          showLoading(false);
          showStatus('Error: ' + err.message, 'error');
        }
      }, 10);
    };
    reader.readAsArrayBuffer(file);
  }

  function processWorkbook(wb) {
    const data = [];
    wb.SheetNames.forEach(function(sn) {
      XLSX.utils.sheet_to_json(wb.Sheets[sn]).forEach(function(r) {
        const ticket = fget(r, ['Incident Ticket', 'Ticket']);
        if (!ticket) return;
        
        // Build comprehensive row object
        const row = {
          ticket: ticket,
          month: normKey(fget(r, ['Month', 'Date', 'Created']), 'Unknown'),
          sla: normSla(fget(r, ['SLA Code', 'SLA', 'SLA Code'])),
          km1: normBoolLike(fget(r, ['KM-1', 'KM1', 'KM Code'])),
          aos: normAos(fget(r, ['AOS Portal', 'AOS', 'Portal'])),
          status: normKey(fget(r, ['Status', 'State']), 'N/A'),
          severity: normKey(fget(r, ['Severity', 'Priority']), 'N/A'),
          breachType: normKey(fget(r, ['Breach Type', 'Type']), 'Unknown'),
          language: normKey(fget(r, ['Language', 'Lang']), 'Unknown'),
          excluded: normBoolLike(fget(r, ['Excluded', 'Exclude'])),
          details: r
        };
        data.push(row);
      });
    });

    console.log('Processed', data.length, 'records');
    rptState.allData = data;
    applyFilters();
    renderKPIs();
    initTabs();
    showLoading(false);
    showStatus('Data loaded: ' + data.length + ' records');
    
    // Show data area, hide upload area
    document.getElementById('upload-area').style.display = 'none';
    document.getElementById('data-area').style.display = 'flex';
  }

  function applyFilters() {
    const monthSelect = document.getElementById('month-filter');
    const slaSelect = document.getElementById('sla-filter');
    const monthVal = monthSelect ? monthSelect.value : '';
    const slaVal = slaSelect ? slaSelect.value : '';

    rptState.filtered = rptState.allData.filter(function(row) {
      return (!monthVal || row.month === monthVal) &&
             (!slaVal || row.sla === slaVal);
    });

    console.log('Filtered to', rptState.filtered.length, 'records');
    renderActiveTab();
    renderCharts();
  }

  function renderKPIs() {
    const months = {};
    const slas = {};

    rptState.allData.forEach(function(row) {
      months[row.month] = (months[row.month] || 0) + 1;
      slas[row.sla] = (slas[row.sla] || 0) + 1;
    });

    const monthSelect = document.getElementById('month-filter');
    const slaSelect = document.getElementById('sla-filter');
    if (monthSelect) {
      monthSelect.innerHTML = '<option value="">All Months</option>';
      Object.keys(months).forEach(function(m) {
        const opt = document.createElement('option');
        opt.value = m;
        opt.text = m + ' (' + months[m] + ')';
        monthSelect.appendChild(opt);
      });
    }
    if (slaSelect) {
      slaSelect.innerHTML = '<option value="">All SLAs</option>';
      Object.keys(slas).forEach(function(s) {
        const opt = document.createElement('option');
        opt.value = s;
        opt.text = s + ' (' + slas[s] + ')';
        slaSelect.appendChild(opt);
      });
    }
  }

  function renderCharts() {
    if (rptState.activeTab === 'overview') {
      window.RPT.renderKPIs();
    }
  }

  function initTabs() {
    const tabBar = document.getElementById('tabs-bar');
    if (!tabBar) return;
    tabBar.style.display = 'flex';
    const tabs = document.querySelectorAll('[id="tab-"]');
    tabs.forEach(function(el) {
      el.style.display = 'none';
    });
  }

  function renderActiveTab() {
    const tab = rptState.activeTab;
    const el = document.getElementById('tab-' + tab);
    if (!el) return;
    el.style.display = 'flex';

    if (tab === 'aos') {
      renderAosTable();
    } else if (tab === 'km1') {
      renderKm1Table();
    }
  }

  function renderAosTable() {
    const container = document.getElementById('tab-aos');
    if (!container) return;
    const filtered = rptState.filtered.filter(function(r) { return r.aos === 'Y'; });
    let html = '<table class="data-table"><thead><tr><th>Ticket</th><th>Month</th><th>SLA</th><th>Status</th></tr></thead><tbody>';
    filtered.slice(0, MAX_DISPLAY_ROWS).forEach(function(r) {
      html += '<tr><td>' + clean(r.ticket) + '</td><td>' + clean(r.month) + '</td><td>' + clean(r.sla) + '</td><td>' + clean(r.status) + '</td></tr>';
    });
    html += '</tbody></table>';
    container.innerHTML = html;
  }

  function renderKm1Table() {
    const container = document.getElementById('tab-km1');
    if (!container) return;
    const filtered = rptState.filtered.filter(function(r) { return r.km1 === '1'; });
    let html = '<table class="data-table"><thead><tr><th>Ticket</th><th>Month</th><th>SLA</th><th>Status</th></tr></thead><tbody>';
    filtered.slice(0, MAX_DISPLAY_ROWS).forEach(function(r) {
      html += '<tr><td>' + clean(r.ticket) + '</td><td>' + clean(r.month) + '</td><td>' + clean(r.sla) + '</td><td>' + clean(r.status) + '</td></tr>';
    });
    html += '</tbody></table>';
    container.innerHTML = html;
  }

  // === Exports ===
  window.RPT.loadFile = loadFile;
  window.RPT.applyFilters = applyFilters;
  window.RPT.renderKPIs = renderKPIs;
})();
