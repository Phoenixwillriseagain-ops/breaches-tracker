// data-processor.js - File handling and data processing for Breaches Tracker
(function() {
  'use strict';

  const { TAB_DEFS, C_XLSX, C_CSV, OUT_COLS, MANUAL_COLS, CONFIG } = window.BT;
  const { MAX_DISPLAY_ROWS, EXPORT_PREFIXES } = CONFIG;

  const trackerState = {
    tabs: {},
    activeTab: '',
    sortCol: null,
    sortDir: 'asc',
    isProcessing: false,
    loadedMonths: [],
    // Set of all Incident Ticket values seen — used for deduplication across imports
    _seenTickets: new Set(),
  };

  window.BT.trackerState = trackerState;

  function debounce(fn, delay) {
    let timer;
    return function(...args) {
      clearTimeout(timer);
      timer = setTimeout(() => fn.apply(this, args), delay);
    };
  }
  window.BT.debounce = debounce;

  // Format date → dd.mm.yyyy hh:mm:ss in EEST (UTC+3)
  function formatDate(val) {
    if (!val) return '';
    const d = new Date(val);
    if (isNaN(d.getTime())) return String(val);
    const TZ_OFFSET_MS = 3 * 60 * 60 * 1000; // EEST = UTC+3
    const local = new Date(d.getTime() + TZ_OFFSET_MS);
    const dd   = String(local.getUTCDate()).padStart(2, '0');
    const mm   = String(local.getUTCMonth() + 1).padStart(2, '0');
    const yyyy = local.getUTCFullYear();
    const hh   = String(local.getUTCHours()).padStart(2, '0');
    const min  = String(local.getUTCMinutes()).padStart(2, '0');
    const ss   = String(local.getUTCSeconds()).padStart(2, '0');
    return `${dd}.${mm}.${yyyy} ${hh}:${min}:${ss}`;
  }

  function showStatus(msg, type = 'ok') {
    const el = document.getElementById('status');
    if (!el) return;
    el.textContent = msg;
    el.className = 'status-bar' + (type === 'error' ? ' error' : '');
    el.style.display = '';
  }
  window.BT.showStatus = showStatus;

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

  // ── Import log helpers ──────────────────────────────────────────────────
  function showImportLog(lines) {
    const el = document.getElementById('import-log');
    if (!el) return;
    el.innerHTML = lines.map(function(l) {
      const icon = l.type === 'ok'   ? '✓' :
                   l.type === 'warn' ? '⚠' : '✕';
      const color = l.type === 'ok'   ? 'var(--success, #437a22)' :
                    l.type === 'warn' ? 'var(--warning, #964219)' : 'var(--error, #a12c7b)';
      return `<span style="color:${color};margin-right:8px;">${icon}</span>${l.msg}`;
    }).join('<br>');
    el.style.display = 'block';
  }

  function hideImportLog() {
    const el = document.getElementById('import-log');
    if (el) el.style.display = 'none';
  }

  // ── Multi-file import entry point ───────────────────────────────────────
  // Accepts an array of File objects, processes them sequentially, merges into state
  window.BT.startMultiImport = function(files) {
    hideImportLog();
    const logLines = [];
    let fileIndex = 0;

    function next() {
      if (fileIndex >= files.length) {
        // All files done — render once
        showLoading(false);
        const total = Object.values(trackerState.tabs).reduce((s, a) => s + a.length, 0);
        showStatus('Total: ' + total + ' records across ' + files.length + ' file(s)');
        showImportLog(logLines);
        if (typeof window.render === 'function') window.render();
        return;
      }
      const file = files[fileIndex++];
      showLoading(true, 'Importing ' + file.name + ' (' + fileIndex + '/' + files.length + ')...');
      setTimeout(function() {
        readFile(file, function(rows, C, err) {
          if (err) {
            logLines.push({ type: 'error', msg: file.name + ' — ' + err });
            next();
            return;
          }
          const result = mergeRows(rows, C);
          const icon = result.duplicates > 0 ? 'warn' : 'ok';
          let msg = file.name + ' — ' + result.added + ' record(s) added';
          if (result.duplicates > 0) msg += ', ' + result.duplicates + ' duplicate(s) skipped';
          logLines.push({ type: icon, msg: msg });
          next();
        });
      }, 10);
    }

    // Initialise tabs if this is the very first import
    if (Object.keys(trackerState.tabs).length === 0) {
      TAB_DEFS.forEach(t => { trackerState.tabs[t.id] = []; });
    }

    next();
  };

  // ── Clear all data ──────────────────────────────────────────────────────
  window.BT.clearAllData = function() {
    TAB_DEFS.forEach(t => { trackerState.tabs[t.id] = []; });
    trackerState._seenTickets = new Set();
    hideImportLog();
    showStatus('All data cleared.');
    if (typeof window.BT.renderEmptyState === 'function') window.BT.renderEmptyState(true);
  };

  // ── File reader — returns raw rows + column map via callback ───────────
  function readFile(file, cb) {
    const ext = file.name.split('.').pop().toLowerCase();
    const reader = new FileReader();

    if (ext === 'csv') {
      reader.onload = function(e) {
        try {
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
          cb(rows.slice(1), C_CSV, null);
        } catch(e) { cb(null, null, e.message); }
      };
      reader.onerror = function() { cb(null, null, 'Read error'); };
      reader.readAsText(file, 'UTF-8');

    } else if (['xlsx','xls'].includes(ext)) {
      reader.onload = function(e) {
        try {
          const wb = XLSX.read(e.target.result, { type: 'array', cellDates: true });
          const allRows = [];
          wb.SheetNames.forEach(function(sn) {
            if (sn === 'Instructions') return;
            const sheet = wb.Sheets[sn];
            const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
            rows.slice(1).forEach(function(r) { allRows.push(r); });
          });
          cb(allRows, C_XLSX, null);
        } catch(e) { cb(null, null, e.message); }
      };
      reader.onerror = function() { cb(null, null, 'Read error'); };
      reader.readAsArrayBuffer(file);

    } else {
      cb(null, null, 'Unsupported file type: .' + ext);
    }
  }

  // ── Merge rows into existing state (with deduplication) ─────────────────
  // Returns { added, duplicates }
  function mergeRows(rows, C) {
    let added = 0, duplicates = 0;

    rows.forEach(function(row) {
      if (!row[C.ticket]) return;
      const ticket = String(row[C.ticket]).trim();
      if (!ticket) return;

      // Deduplication: skip if this ticket was already loaded
      if (trackerState._seenTickets.has(ticket)) {
        duplicates++;
        return;
      }
      trackerState._seenTickets.add(ticket);

      const sla  = String(row[C.sla_code] || '').trim();
      const lang = String(row[C.lang]     || '').trim().toLowerCase();
      const nok  = parseInt(row[C.nok])   || 0;
      const mapped = mapRow(row, C);

      TAB_DEFS.forEach(function(t) {
        if (t.code !== sla) return;
        const isDE = lang === 'de';
        if (t.lang === 'de'  && !isDE) return;
        if (t.lang === '!de' && isDE)  return;
        if (t.nokFilter && nok !== 1)  return;
        trackerState.tabs[t.id].push(mapped);
        added++;
      });
    });

    return { added, duplicates };
  }

  // ── Map a row array to output column object ──────────────────────────────
  function mapRow(row, C) {
    return {
      'Incident Ticket':    String(row[C.ticket]      || ''),
      'DATE_CLOSE':         formatDate(row[C.date_close] || ''),
      'Status':             String(row[C.status]      || ''),
      'Queue':              String(row[C.queue]        || ''),
      'Priority':           String(row[C.priority]    || ''),
      'ISO_Language':       String(row[C.lang]         || ''),
      'Tool':               String(row[C.tool]         || ''),
      'TOPIC':              String(row[C.topic]        || ''),
      'SLA_Code':           String(row[C.sla_code]     || ''),
      'SLA_N':              String(row[C.sla_n]        || ''),
      'Breach_Description': String(row[C.breach_desc]  || ''),
      'DATE_TIME_Breach':   formatDate(row[C.breach_dt]  || ''),
      'COMPASS ID':         C.compass_id  !== undefined ? String(row[C.compass_id]  || '') : '',
      'Reason':             C.reason      !== undefined ? String(row[C.reason]       || '') : '',
      'AOS':                C.aos         !== undefined ? String(row[C.aos]          || '') : '',
      'Agent':              C.agent       !== undefined ? String(row[C.agent]        || '') : '',
      'BMS ID':             C.bms_id      !== undefined ? String(row[C.bms_id]       || '') : '',
      'Comment':            C.comment     !== undefined ? String(row[C.comment]      || '') : '',
      'AOS Issue':          C.aos_issue   !== undefined ? String(row[C.aos_issue]    || '') : '',
      'Excluded':           C.excluded    !== undefined ? String(row[C.excluded]     || '') : '',
      'Jira':               C.jira        !== undefined ? String(row[C.jira]         || '') : '',
      'Week':               C.week        !== undefined ? String(row[C.week]         || '') : '',
      'Unique':             C.unique      !== undefined ? String(row[C.unique]       || '') : '',
    };
  }

  // ── Legacy single-file entry point (kept for internal safety) ───────────
  function handleFile(file) {
    if (!file) return;
    window.BT.startMultiImport([file]);
  }
  window.BT.handleFile = handleFile;

  // ── Export helpers ───────────────────────────────────────────────────────
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

  function init() {
    showLoading(false);
  }
  window.BT.init = init;

})();
