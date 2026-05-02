// reporter-processor.js - Reporter page: file handling, processing, filtering, KPIs
(function() {
  'use strict';

  const { REPORTER_TABS, REPORTER_STORAGE_KEY, CONFIG } = window.BT || {};
  const { MAX_DISPLAY_ROWS } = CONFIG;

  const rptState = window.RPT = {
    allData: [], filtered: [], activeTab: 'overview',
    charts: {}, loadedMonths: [], activeMonth: null,
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
    el.className = 'status-bar' + (type === 'error' ? ' error' : '');
    el.style.display = '';
  }

  function showLoading(show, msg) {
    let ov = document.getElementById('loading-overlay');
    if (show) {
      if (!ov) {
        ov = document.createElement('div');
        ov.id = 'loading-overlay'; ov.className = 'loading-overlay';
        ov.innerHTML = '<div class="spinner"></div><div class="loading-text">' + (msg || 'Processing...') + '</div>';
        document.body.appendChild(ov);
      } else {
        ov.querySelector('.loading-text').textContent = msg || 'Processing...';
        ov.style.display = 'flex';
      }
    } else if (ov) ov.style.display = 'none';
  }

  function fget(row, fields) {
    if (Array.isArray(fields)) {
      for (var i = 0; i < fields.length; i++) {
        var f = fields[i];
        if (row.hasOwnProperty(f) && row[f] !== undefined && row[f] !== null && row[f] !== '') return String(row[f]);
      }
      return '';
    }
    return String(row[fields] || '');
  }
  

  function loadFile(file) {
    if (!file) return;
    showLoading(true, 'Reading ' + file.name + '...');
    const reader = new FileReader();
    reader.onload = function(e) {
      showLoading(true, 'Processing data...');
      setTimeout(function() {
        try {
          const wb = XLSX.read(e.target.result, { type: 'array', cellDates: true });
          processWorkbook(wb);
        } catch (err) {
          showLoading(false); showStatus('Error: ' + err.message, 'error');
        }
      }, 10);
    };
    reader.readAsArrayBuffer(file);
  }

  function processWorkbook(wb) {
    const data = [];
    wb.SheetNames.forEach(function(sn) {
      XLSX.utils.sheet_to_json(wb.Sheets[sn]).forEach(function(r) {
        const ticket = fget(r, ['Incident Ticket','Ticket']);
        if (!ticket) return;
        data.push({
          ticket: ticket, sla_code: fget(r,['SLA_Code','SLA Code']),
          excluded: fget(r,['Excluded','excluded','EXCLUDED']),
          week: fget(r,['Week','week']), reason: fget(r,['Reason','reason','Breach_Description']),
          application: fget(r,['Application','application','APP']),
          language: fget(r,['Language','ISO_Language','lang']),
          date_close: fget(r,['DATE_CLOSE','Closure Date']),
          remote: fget(r,['Remote','remote']), callback: fget(r,['Callback','callback']),
          aos: fget(r,['AOS','aos']), month: fget(r,['Month','month']),
          date_breach: fget(r,['DATE_TIME_Breach','Breach Time']),
        });
      });
    });
    rptState.allData = data;
    rptState.loadedMonths = [];
    data.forEach(function(d) { if (d.month && rptState.loadedMonths.indexOf(d.month) === -1) rptState.loadedMonths.push(d.month); });
    saveData(); showLoading(false);
    showStatus('Loaded ' + data.length + ' records');
    renderAll();
  }

  function saveData() {
    try { localStorage.setItem(REPORTER_STORAGE_KEY, JSON.stringify(rptState.allData)); }
    catch (e) { console.warn('Cannot save:', e); }
  }

  function loadSavedData() {
    try {
      const d = localStorage.getItem(REPORTER_STORAGE_KEY);
      if (d) { rptState.allData = JSON.parse(d); showStatus('Restored ' + rptState.allData.length + ' records'); renderAll(); }
    } catch (e) { console.warn('Cannot load:', e); }
  }
  

  function applyFilters() {
    var wk = document.getElementById('filter-week').value;
    var sla = document.getElementById('filter-sla').value;
    var lang = document.getElementById('filter-lang').value;
    var excl = document.getElementById('filter-excl').value;
    var month = rptState.activeMonth;
    rptState.filtered = rptState.allData.filter(function(d) {
      if (month && d.month !== month) return false;
      if (wk && d.week !== wk) return false;
      if (sla && d.sla_code !== sla) return false;
      if (lang && d.language !== lang) return false;
      if (excl && d.excluded !== excl) return false;
      return true;
    });
    renderKPIs(); renderActiveTab();
  }

  var debouncedApply = debounce(applyFilters, 200);

  function renderKPIs() {
    var c = document.getElementById('kpi-area'); if (!c) return;
    var d = rptState.filtered;
    var total = d.length;
    var excl = d.filter(function(r) { return r.excluded === '1'; }).length;
    var aos = d.filter(function(r) { return r.aos && r.aos !== '0'; }).length;
    var kpis = [
      { l: 'Total', v: total, s: 'records' },
      { l: 'Counted', v: total - excl, s: 'tickets' },
      { l: 'Excluded', v: excl, s: 'tickets' },
      { l: 'AOS Issues', v: aos, s: 'found' },
      { l: 'Agents', v: uniqueCount(d, 'agent'), s: 'unique' },
      { l: 'Weeks', v: uniqueCount(d, 'week'), s: 'span' },
    ];
    c.innerHTML = kpis.map(function(k) {
      return '<div class="kpi"><div class="kpi-label">' + k.l + '</div><div class="kpi-value">' + k.v + '</div><div class="kpi-sub">' + k.s + '</div></div>';
    }).join('');
  }

  function uniqueCount(data, key) {
    var seen = {};
    data.forEach(function(d) { if (d[key]) seen[d[key]] = true; });
    return Object.keys(seen).length;
  }

  function uniqueValues(data, key) {
    var seen = {};
    data.forEach(function(d) { if (d[key] && d[key] !== '') seen[d[key]] = true; });
    return Object.keys(seen).sort();
  }

  function fillSelect(sel, values) {
    if (!sel) return;
    sel.innerHTML = '<option value="">All</option>' + values.map(function(v) {
      return '<option value="' + v + '">' + v + '</option>';
    }).join('');
  }

  function buildTable(tableId, rows, cols) {
    var tbl = document.getElementById(tableId); if (!tbl) return;
    tbl.innerHTML = '';
    if (!rows || !rows.length) {
      tbl.innerHTML = '<tr><td colspan="' + cols.length + '" style="text-align:center;padding:3rem;color:var(--muted)">No data</td></tr>';
      return;
    }
    var thead = document.createElement('thead');
    var trh = document.createElement('tr');
    cols.forEach(function(col) {
      var th = document.createElement('th'); th.textContent = col;
      trh.appendChild(th);
    });
    thead.appendChild(trh); tbl.appendChild(thead);
    var tbody = document.createElement('tbody');
    var frag = document.createDocumentFragment();
    var limit = Math.min(rows.length, MAX_DISPLAY_ROWS);
    for (var i = 0; i < limit; i++) {
      var tr = document.createElement('tr');
      cols.forEach(function(col) {
        var td = document.createElement('td');
        td.textContent = rows[i][col] || '';
        td.title = String(rows[i][col] || '');
        tr.appendChild(td);
      });
      frag.appendChild(tr);
    }
    tbody.appendChild(frag); tbl.appendChild(tbody);
    if (rows.length > MAX_DISPLAY_ROWS) {
      var note = document.createElement('tr');
      note.innerHTML = '<td colspan="' + cols.length + '" style="text-align:center;color:var(--muted);font-size:0.75rem;">Showing ' + limit + ' of ' + rows.length + '</td>';
      tbody.appendChild(note);
    }
  }
  

  function renderMonths() {
    var mArea = document.getElementById('months-area');
    if (!mArea) return;
    var months = rptState.loadedMonths.sort();
    if (!months.length) { mArea.innerHTML = ''; return; }
    mArea.innerHTML = months.map(function(m) {
      var active = m === rptState.activeMonth ? ' active' : '';
      return '<span class="month-badge' + active + '" data-month="' + m + '">' + m + '</span>';
    }).join('');
    mArea.querySelectorAll('.month-badge').forEach(function(b) {
      b.addEventListener('click', function() {
        rptState.activeMonth = rptState.activeMonth === b.dataset.month ? null : b.dataset.month;
        mArea.querySelectorAll('.month-badge').forEach(function(b2) {
          b2.classList.toggle('active', b2 === b && rptState.activeMonth);
        });
        if (!rptState.activeMonth) b.classList.remove('active');
        applyFilters();
      });
    });
  }

  function renderFilters() {
    renderMonths();
    fillSelect(document.getElementById('filter-week'), uniqueValues(rptState.filtered, 'week'));
    fillSelect(document.getElementById('filter-sla'), uniqueValues(rptState.filtered, 'sla_code'));
    fillSelect(document.getElementById('filter-lang'), uniqueValues(rptState.filtered, 'language'));
  }

  function switchRptTab(id) {
    rptState.activeTab = id;
    if (REPORTER_TABS) {
      var bar = document.getElementById('tabs-bar');
      if (bar) {
        bar.innerHTML = REPORTER_TABS.map(function(t) {
          return '<button class="tab-btn' + (t.id === id ? ' active' : '') + '" data-tab="' + t.id + '">' + t.label + '</button>';
        }).join('');
        bar.querySelectorAll('.tab-btn').forEach(function(b) {
          b.addEventListener('click', function() { switchRptTab(b.dataset.tab); });
        });
      }
    }
    document.querySelectorAll('[id^="tab-"]').forEach(function(el) { el.style.display = 'none'; });
    var tEl = document.getElementById('tab-' + id);
    if (tEl) tEl.style.display = '';
    renderActiveTab();
  }

  function renderActiveTab() {
    var t = rptState.activeTab;
    var rc = window.RPT_CHARTS;
    if (t === 'overview' && rc) rc.renderOverview();
    if (t === 'reasons' && rc) rc.renderReasons();
    if (t === 'apps' && rc) rc.renderApps();
    if (t === 'aos' && rc) rc.renderAOS();
    if (t === 'km1' && rc) rc.renderKM1();
    if (t === 'all') renderAllRecords();
  }

  function renderAllRecords() {
    buildTable('report-table', rptState.filtered, ['ticket','sla_code','excluded','week','reason','language','date_close']);
  }

  function renderAll() {
    var has = rptState.allData.length > 0;
    document.getElementById('upload-area').style.display = has ? 'none' : '';
    document.getElementById('kpi-area').style.display = has ? '' : 'none';
    document.getElementById('months-area').style.display = has ? '' : 'none';
    document.querySelector('.toolbar').style.display = has ? '' : 'none';
    document.getElementById('tabs-bar').style.display = has ? '' : 'none';
    if (has) { renderFilters(); applyFilters(); switchRptTab('overview'); }
  }
  

  function today() { return new Date().toISOString().slice(0,10).replace(/-/g,''); }

  function exportFiltered() {
    var d = rptState.filtered;
    if (!d.length) { showStatus('No data to export.', 'error'); return; }
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(d), 'Filtered');
    XLSX.writeFile(wb, 'Breaches_Filtered_' + today() + '.xlsx');
  }

  function exportReport() {
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rptState.filtered), 'All Records');
    var rm = {};
    rptState.filtered.forEach(function(d) {
      if (!d.reason) d.reason = 'Unknown';
      rm[d.reason] = (rm[d.reason] || 0) + 1;
    });
    var rd = Object.keys(rm).map(function(k) { return { Reason: k, Count: rm[k] }; });
    rd.sort(function(a,b) { return b.Count - a.Count; });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rd), 'Reasons');
    XLSX.writeFile(wb, 'Breaches_FullReport_' + today() + '.xlsx');
  }

  function setupUpload() {
    var input = document.getElementById('reporter-file-input');
    var btn = document.querySelector('.upload-btn');
    if (input) input.addEventListener('change', function() { loadFile(input.files[0]); });
    if (btn) btn.addEventListener('click', function() { input.click(); });
    ['filter-week','filter-sla','filter-lang','filter-excl'].forEach(function(id) {
      var el = document.getElementById(id);
      if (el) el.addEventListener('change', debouncedApply);
    });
  }

  function initTheme() {
    var btn = document.querySelector('[data-theme-toggle]');
    var root = document.documentElement;
    var d = window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
    root.setAttribute('data-theme', d);
    if (btn) btn.addEventListener('click', function() {
      d = d === 'dark' ? 'light' : 'dark';
      root.setAttribute('data-theme', d);
    });
  }

  function init() {
    initTheme(); setupUpload(); loadSavedData();
    if (!rptState.allData.length) {
      document.getElementById('kpi-area').style.display = 'none';
      document.getElementById('months-area').style.display = 'none';
      document.querySelector('.toolbar').style.display = 'none';
      document.getElementById('tabs-bar').style.display = 'none';
    }
  }

  window.RPT.loadFile = loadFile;
  window.RPT.applyFilters = applyFilters;
  window.RPT.exportFiltered = exportFiltered;
  window.RPT.exportReport = exportReport;
  window.RPT.switchTab = switchRptTab;
  window.RPT.renderKPIs = renderKPIs;
  window.RPT.init = init;

})();
