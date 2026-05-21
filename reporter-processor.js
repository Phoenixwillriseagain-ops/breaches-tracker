// reporter-processor.js - V2 model processor for Breaches Reporter
(function() {
  'use strict';
  const { CONFIG } = window.BT || {};
  const { MAX_DISPLAY_ROWS } = CONFIG || { MAX_DISPLAY_ROWS: 100 };

  window.RPT = {
    allData: [],
    filtered: [],
    aosFiltered: [],
    charts: {},
    columnMap: {},
    uniqueValues: {},
    loadFile: function(file) { loadFile(file); },
    clearFilters: function() {
      // Reset all filter selects by their known IDs
      ['filter-week','filter-sla','filter-lang','filter-excl',
       'filter-aos','filter-tool','filter-queue','filter-sheet'].forEach(function(id) {
        var el = document.getElementById(id);
        if (el) el.value = 'All';
      });
      applyFilters();
    },
    resetToUpload: function() { location.reload(); },
    exportFiltered: function() {
      if (window.EXPORTER) window.EXPORTER.exportAsExcel(window.RPT.filtered);
      else alert('Exporter not loaded');
    },
    exportReport: function() {
      if (window.EXPORTER) window.EXPORTER.exportAsExcel(
        window.RPT.allData,
        'Breaches_FullReport_' + new Date().toISOString().slice(0,10).replace(/-/g,'') + '.xlsx'
      );
      else alert('Exporter not loaded');
    },
  };

  function clean(v) {
    return String(v == null ? '' : v).trim();
  }

  // Format date → dd.mm.yyyy hh:mm:ss in CET (UTC+2)
  function formatDate(val) {
    if (!val) return '';
    const d = new Date(val);
    if (isNaN(d.getTime())) return clean(val);
    const TZ_OFFSET_MS = 2 * 60 * 60 * 1000; // CET = UTC+2
    const local = new Date(d.getTime() + TZ_OFFSET_MS);
    const dd  = String(local.getUTCDate()).padStart(2, '0');
    const mm  = String(local.getUTCMonth() + 1).padStart(2, '0');
    const yyyy = local.getUTCFullYear();
    const hh  = String(local.getUTCHours()).padStart(2, '0');
    const min = String(local.getUTCMinutes()).padStart(2, '0');
    const ss  = String(local.getUTCSeconds()).padStart(2, '0');
    return `${dd}.${mm}.${yyyy} ${hh}:${min}:${ss}`;
  }

  function normBoolLike(v) {
    const s = clean(v).toLowerCase();
    if (!s) return 'N';
    if (['t', 'true', 'yes', 'y', '1'].indexOf(s) !== -1) return 'Y';
    return 'N';
  }

  // ─── V2 MODEL: processWorkbook ────────────────────────────────────────────
  // V2 headers (same across all sheets):
  // 0:Incident Ticket  1:DATE_CLOSE  2:Status  3:Queue  4:Priority
  // 5:ISO_Language  6:Tool  7:TOPIC  8:SLA_Code  9:SLA_N
  // 10:Breach_Description  11:DATE_TIME_Breach  12:Munich time
  // 13:COMPASS ID  14:Reason  15:AOS  16:Agent  17:BMS ID
  // 18:Comment  19:AOS Issue  20:Excluded  21:Jira  22:Week  23:Unique
  function processWorkbook(wb) {
    const data = [];

    wb.SheetNames.forEach(function(sn) {
      if (sn === 'Instructions') return;
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[sn], { defval: '' });

      rows.forEach(function(r) {
        const ticket = clean(r['Incident Ticket'] || '');
        if (!ticket) return;

        const aosFlag  = normBoolLike(r['AOS']       || '');
        const aosIssue = normBoolLike(r['AOS Issue'] || '');

        const prow = {
          ticket:         ticket,
          dateClosed:     formatDate(r['DATE_CLOSE']       || ''),
          dateTimeBreach: formatDate(r['DATE_TIME_Breach'] || ''),
          munichTime:     formatDate(r['Munich time']      || ''),
          status:         clean(r['Status']        || 'N/A'),
          queue:          clean(r['Queue']         || ''),
          priority:       clean(r['Priority']      || ''),
          language:       clean(r['ISO_Language']  || 'Unknown'),
          tool:           clean(r['Tool']          || 'Unknown'),
          topic:          clean(r['TOPIC']         || ''),
          sla:            clean(r['SLA_Code']      || 'Unknown'),
          slaN:           clean(r['SLA_N']         || ''),
          breachDesc:     clean(r['Breach_Description'] || ''),
          compassId:      clean(r['COMPASS ID']    || ''),
          reason:         clean(r['Reason']        || ''),
          aos:            aosFlag,
          agent:          clean(r['Agent']         || ''),
          bmsId:          clean(r['BMS ID']        || ''),
          comment:        clean(r['Comment']       || ''),
          aosIssue:       aosIssue,
          excluded:       normBoolLike(r['Excluded'] || ''),
          jira:           clean(r['Jira']          || ''),
          week:           clean(r['Week']          || ''),
          unique:         clean(r['Unique']        || ''),
          sheet:          sn,
          isAos:          aosFlag === 'Y' || aosIssue === 'Y',
        };
        data.push(prow);
      });
    });

    console.log('V2 Reporter loaded', data.length, 'records');
    window.RPT.allData     = data;
    window.RPT.filtered    = data.slice();
    window.RPT.aosFiltered = data.filter(r => r.isAos);

    populateUniqueValues();
    updateFilterDropdowns(); // shows #data-section, hides #upload-section

    // Apply filters updates RPT.filtered, then defer renderTables one tick
    // so that reporter.html's inline <script> (which defines window.renderTables)
    // is guaranteed to have executed before we call it.
    applyFiltersOnly();
    setTimeout(function() {
      if (typeof window.renderTables === 'function') window.renderTables();
    }, 0);
  }

  function loadFile(file) {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = function(e) {
      try {
        const wb = XLSX.read(e.target.result, { type: 'array', cellDates: true });
        processWorkbook(wb);
      } catch (err) {
        console.error('Error loading file:', err);
        alert('Error loading file: ' + err.message);
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function populateUniqueValues() {
    const uv = { weeks: [], slas: [], languages: [], tools: [], queues: [], sheets: [] };
    window.RPT.allData.forEach(function(row) {
      if (uv.weeks.indexOf(row.week)         === -1 && row.week)     uv.weeks.push(row.week);
      if (uv.slas.indexOf(row.sla)           === -1 && row.sla)      uv.slas.push(row.sla);
      if (uv.languages.indexOf(row.language) === -1 && row.language) uv.languages.push(row.language);
      if (uv.tools.indexOf(row.tool)         === -1 && row.tool)     uv.tools.push(row.tool);
      if (uv.queues.indexOf(row.queue)       === -1 && row.queue)    uv.queues.push(row.queue);
      if (uv.sheets.indexOf(row.sheet)       === -1 && row.sheet)    uv.sheets.push(row.sheet);
    });
    uv.weeks.sort(); uv.slas.sort(); uv.languages.sort(); uv.tools.sort(); uv.queues.sort();
    window.RPT.uniqueValues = uv;
  }

  function updateFilterDropdowns() {
    const uv = window.RPT.uniqueValues;
    const sel = (id, arr) => {
      const el = document.getElementById(id);
      if (!el) return;
      el.innerHTML = '<option value="All">All</option>' +
        arr.map(v => `<option value="${v}">${v}</option>`).join('');
    };
    sel('filter-week',  uv.weeks);
    sel('filter-sla',   uv.slas);
    sel('filter-lang',  uv.languages);
    sel('filter-tool',  uv.tools);
    sel('filter-queue', uv.queues);
    sel('filter-sheet', uv.sheets);

    const exclEl = document.getElementById('filter-excl');
    if (exclEl) exclEl.innerHTML =
      '<option value="All">All</option><option value="Y">Excluded</option><option value="N">Counted</option>';
    const aosEl = document.getElementById('filter-aos');
    if (aosEl) aosEl.innerHTML =
      '<option value="All">All</option><option value="Y">AOS Only</option><option value="N">Non-AOS</option>';

    // Show data area, hide upload area
    const uploadSection = document.getElementById('upload-section');
    const dataSection   = document.getElementById('data-section');
    if (uploadSection) uploadSection.style.display = 'none';
    if (dataSection)   dataSection.style.display   = '';
  }

  // Filters data only — does NOT call renderTables (avoids double render on load)
  function applyFiltersOnly() {
    const g = id => (document.getElementById(id)?.value || 'All');
    const weekF  = g('filter-week');
    const slaF   = g('filter-sla');
    const langF  = g('filter-lang');
    const exclF  = g('filter-excl');
    const aosF   = g('filter-aos');
    const toolF  = g('filter-tool');
    const queueF = g('filter-queue');
    const sheetF = g('filter-sheet');

    window.RPT.filtered = window.RPT.allData.filter(function(row) {
      if (weekF  !== 'All' && row.week     !== weekF)  return false;
      if (slaF   !== 'All' && row.sla      !== slaF)   return false;
      if (langF  !== 'All' && row.language !== langF)  return false;
      if (exclF  !== 'All' && row.excluded !== exclF)  return false;
      if (aosF   !== 'All' && (row.isAos ? 'Y' : 'N') !== aosF) return false;
      if (toolF  !== 'All' && row.tool     !== toolF)  return false;
      if (queueF !== 'All' && row.queue    !== queueF) return false;
      if (sheetF !== 'All' && row.sheet    !== sheetF) return false;
      return true;
    });

    const countEl = document.getElementById('record-count');
    if (countEl) countEl.textContent = window.RPT.filtered.length + ' records';
  }

  // Full apply: filter + re-render (used by dropdowns and clearFilters)
  function applyFilters() {
    applyFiltersOnly();
    if (typeof window.renderTables === 'function') window.renderTables();
  }

  document.addEventListener('DOMContentLoaded', function() {
    ['filter-week','filter-sla','filter-lang','filter-excl',
     'filter-aos','filter-tool','filter-queue','filter-sheet'].forEach(function(id) {
      const el = document.getElementById(id);
      if (el) el.addEventListener('change', applyFilters);
    });
  });

})();
