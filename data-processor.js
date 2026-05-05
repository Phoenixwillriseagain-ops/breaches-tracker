// data-processor.js - File handling and data processing for Breaches Tracker
(function() {
  'use strict';

  const { TAB_DEFS, C_XLSX, C_CSV, OUT_COLS, MANUAL_COLS, CONFIG } = window.BT;
  const { MAX_DISPLAY_ROWS, EXPORT_PREFIXES } = CONFIG;

  // Shared state for index.html
  const trackerState = {
    tabs: {},
    activeTab: '',
    sortCol: null,
    sortDir: 'asc',
    isProcessing: false,
    loadedMonths: [],
  };

  window.BT.trackerState = trackerState;

  // Debounce helper
  function debounce(fn, delay) {
    let timer;
    return function(...args) {
      clearTimeout(timer);
      timer = setTimeout(() => fn.apply(this, args), delay);
    };
  }
  window.BT.debounce = debounce;

  // Format date helper
  function formatDate(val) {
    if (!val) return '';
    const d = new Date(val);
    return d.toISOString().slice(0, 19).replace('T', ' ');
  }

  // Get week number from date
  function getWeek(dateStr) {
    const d = new Date(dateStr);
    const firstDay = new Date(d.getFullYear(), 0, 1);
    const wkNum = Math.ceil((((d - firstDay) / 86400000) + firstDay.getDay() + 1) / 7);
    return 'W' + wkNum + '/' + d.getFullYear();
  }

  // Show status message
  function showStatus(msg, type = 'ok') {
    const el = document.getElementById('status');
    if (!el) return;
    el.textContent = msg;
    el.className = 'status-bar' + (type === 'error' ? ' error' : '');
    el.style.display = '';
  }
  window.BT.showStatus = showStatus;

  // Show/hide loading
  function showLoading(show, msg) {
    let overlay = document.getElementById('loading-overlay');
    if (show) {
      if (!overlay) {
        overlay = document.createElement('div');
        overlay.id = 'loading-overlay';
        overlay.className = 'loading-overlay';
        overlay.innerHTML = '<div class="spinner"></div><div class="loading-text">' + (msg || 'Processing...') + '</div>';
        document.body.appendChild(overlay);
      } else {
        overlay.querySelector('.loading-text').textContent = msg || 'Processing...';
        overlay.style.display = 'flex';
      }
    } else if (overlay) {
      overlay.style.display = 'none';
    }
  }
  window.BT.showLoading = showLoading;

  // Process CSV file
  function loadCSV(file) {
    showLoading(true, 'Reading CSV...');
    const reader = new FileReader();
    reader.onload = function(e) {
      showLoading(true, 'Parsing CSV...');
      setTimeout(function() {
        const text = e.target.result;
        const rows = text.trim().split(/\r?\n/).map(r => {
  const result = [];
  let cur = '', inQ = false;
  for (let i = 0; i < r.length; i++) {
    if (r[i] === '"') { inQ = !inQ; }
    else if (r[i] === ',' && !inQ) { result.push(cur.trim()); cur = ''; }
    else { cur += r[i]; }
  }
  result.push(cur.trim());
  return result;
});
        processRows(rows.slice(1), C_CSV, 'csv');
      }, 10);
    };
    reader.readAsText(file, 'UTF-8');
  }

  // Process XLSX file
  function loadXLSX(file) {
    showLoading(true, 'Reading XLSX...');
    const reader = new FileReader();
    reader.onload = function(e) {
      showLoading(true, 'Parsing spreadsheet...');
      setTimeout(function() {
        const wb = XLSX.read(e.target.result, { type: 'array', cellDates: true });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
        processRows(rows.slice(1), C_XLSX, 'xlsx');
      }, 10);
    };
    reader.readAsArrayBuffer(file);
  }

  // Map row to output format
  function mapRow(row, C) {
    return {
      'Incident Ticket':      String(row[C.ticket] || ''),
      'DATE_CLOSE':           formatDate(row[C.date_close] || ''),
      'Status':               String(row[C.status] || ''),
      'Queue':                String(row[C.queue] || ''),
      'Priority':             String(row[C.priority] || ''),
      'ISO_Language':         String(row[C.lang] || ''),
      'Tool':                 String(row[C.tool] || ''),
      'TOPIC':                String(row[C.topic] || ''),
      'SLA_Code':             String(row[C.sla_code] || ''),
      'SLA_N':                String(row[C.sla_n] || ''),
      'Breach_Description':   String(row[C.breach_desc] || ''),
      'DATE_TIME_Breach':     formatDate(row[C.breach_dt] || ''),
      'Agent':                '',
      'BMS ID':               '',
      'Comment if excluded':  '',
      'Additional comment':   '',
      'Excluded':             '',
      'Jira':                 '',
      'Week':                 '',
      'Unique':               '',
    };
  }

  // Core processing - distribute rows across tabs
  function processRows(rows, C, source) {
    const tabs = {};
    TAB_DEFS.forEach(t => { tabs[t.id] = []; });
    trackerState.tabs = tabs;

    rows.forEach(function(row) {
      if (!row[C.ticket]) return;
      const sla = String(row[C.sla_code] || '').trim();
      const lang = String(row[C.lang] || '').trim().toLowerCase();
      const nok = parseInt(row[C.nok]) || 0;
      const mapped = mapRow(row, C);

      TAB_DEFS.forEach(function(t) {
        if (t.code !== sla) return;
        const isDE = lang === 'de';
        if (t.lang === 'de' && !isDE) return;
        if (t.lang === '!de' && isDE) return;
        if (t.nokFilter && nok !== 1) return;
        tabs[t.id].push(mapped);
      });
    });

    trackerState.tabs = tabs;
    showLoading(false);
    showStatus('Loaded ' + rows.length + ' records from ' + source.toUpperCase());

    if (typeof window.render === 'function') {
      window.render();
    }
  }

  // Handle file upload
  function handleFile(file) {
    if (!file) return;
    const ext = file.name.split('.').pop().toLowerCase();
    if (ext === 'csv') {
      loadCSV(file);
    } else if (['xlsx', 'xls'].includes(ext)) {
      loadXLSX(file);
    } else {
      showStatus('Unsupported file type: .' + ext, 'error');
    }
  }
  window.BT.handleFile = handleFile;

  // Export single tab
  function exportTab(tabId) {
    const data = trackerState.tabs[tabId];
    if (!data || !data.length) { showStatus('No data.', 'error'); return; }
    const ws = XLSX.utils.json_to_sheet(data, { header: OUT_COLS });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, tabId);
    const today = new Date().toISOString().slice(0,10).replace(/-/g,'');
    XLSX.writeFile(wb, EXPORT_PREFIXES.tab + tabId + '_' + today + '.xlsx');
  }
  window.BT.exportTab = exportTab;

  // Export all tabs
  function exportAll() {
    const wb = XLSX.utils.book_new();
    TAB_DEFS.forEach(function(t) {
      const data = trackerState.tabs[t.id];
      const ws = XLSX.utils.json_to_sheet(data.length ? data : [Object.fromEntries(OUT_COLS.map(c => [c,'']))], { header: OUT_COLS });
      XLSX.utils.book_append_sheet(wb, ws, t.label);
    });
    const today = new Date().toISOString().slice(0,10).replace(/-/g,'');
    XLSX.writeFile(wb, EXPORT_PREFIXES.all + today + '.xlsx');
  }
  window.BT.exportAll = exportAll;

  // Setup drag-drop
  function setupDropZone() {
    const dz = document.getElementById('drop-zone');
    if (!dz) return;
    dz.addEventListener('dragover', function(e) { e.preventDefault(); dz.classList.add('drag'); });
    dz.addEventListener('dragleave', function() { dz.classList.remove('drag'); });
    dz.addEventListener('drop', function(e) {
      e.preventDefault(); dz.classList.remove('drag');
      const f = e.dataTransfer.files[0]; if (f) handleFile(f);
    });
    dz.addEventListener('click', function() {
      const input = document.getElementById('file-input');
      if (input) input.click();
    });
  }
  window.BT.setupDropZone = setupDropZone;

  // Initialize on load
  function init() {
    showLoading(false);
  }
  window.BT.init = init;

})();
