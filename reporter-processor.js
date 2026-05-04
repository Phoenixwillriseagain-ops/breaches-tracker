// reporter-processor.js - Simple direct column mapping from Excel

(function() {
  'use strict';

  const { CONFIG } = window.BT || {};
  const { MAX_DISPLAY_ROWS } = CONFIG;

  const rptState = window.RPT = {
    allData: [],
    filtered: [],
    activeTab: 'overview',
    charts: {},
    columnMap: {}
  };

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
    if (['1', 'true', 'yes', 'y'].indexOf(s) !== -1) return '1';
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
    if (['1', 'true', 'yes', 'y'].indexOf(s) !== -1) return 'Y';
    return 'N';
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
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function processWorkbook(wb) {
    const data = [];
    let colMap = {};
    
    wb.SheetNames.forEach(function(sn) {
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[sn]);
      
      rows.forEach(function(r, idx) {
        // On first row, detect column mappings
        if (idx === 0) {
          const keys = Object.keys(r);
          console.log('Excel columns:', keys);
          
          // Try to find columns by partial matching on keys
          keys.forEach(function(key) {
            const lower = key.toLowerCase();
            if (lower.includes('ticket') || lower.includes('incident')) colMap.ticket = key;
            if (lower.includes('week') || lower.includes('month') || lower.includes('date')) colMap.month = key;
            if (lower.includes('sla') || lower.includes('code')) colMap.sla = key;
            if (lower.includes('km')) colMap.km1 = key;
            if (lower.includes('aos') || lower.includes('portal')) colMap.aos = key;
            if (lower.includes('reason') || lower.includes('breach')) colMap.reason = key;
            if (lower.includes('lang')) colMap.lang = key;
            if (lower.includes('type') || lower.includes('app')) colMap.type = key;
            if (lower.includes('status')) colMap.status = key;
            if (lower.includes('exclude')) colMap.excluded = key;
          });
          
          // If we still don't have mappings, use first columns as defaults
          if (!colMap.ticket) colMap.ticket = keys[0];
          if (!colMap.month) colMap.month = keys[1];
          if (!colMap.sla) colMap.sla = keys[2];
          
          console.log('Detected mappings:', colMap);
          rptState.columnMap = colMap;
        }
        
        const ticket = clean(r[colMap.ticket] || '');
        if (!ticket) return;
        
        const row = {
          ticket: ticket,
          month: normKey(clean(r[colMap.month] || ''), 'Unknown'),
          sla: normSla(clean(r[colMap.sla] || '')),
          km1: normBoolLike(clean(r[colMap.km1] || '')),
          aos: normAos(clean(r[colMap.aos] || '')),
          status: normKey(clean(r[colMap.status] || ''), 'N/A'),
          breachType: normKey(clean(r[colMap.type] || ''), 'Unknown'),
          language: normKey(clean(r[colMap.lang] || ''), 'Unknown'),
          excluded: normBoolLike(clean(r[colMap.excluded] || '')),
          reason: normKey(clean(r[colMap.reason] || ''), 'Unknown'),
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

    renderActiveTab();
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
      Object.keys(months).sort().forEach(function(m) {
        const opt = document.createElement('option');
        opt.value = m;
        opt.text = m + ' (' + months[m] + ')';
        monthSelect.appendChild(opt);
      });
    }
    
    if (slaSelect) {
      slaSelect.innerHTML = '<option value="">All SLAs</option>';
      Object.keys(slas).sort().forEach(function(s) {
        const opt = document.createElement('option');
        opt.value = s;
        opt.text = s + ' (' + slas[s] + ')';
        slaSelect.appendChild(opt);
      });
    }
  }

  function initTabs() {
    const tabBar = document.getElementById('tabs-bar');
    if (!tabBar) return;
    tabBar.style.display = 'flex';
  }

  function renderActiveTab() {
    const tab = rptState.activeTab;
    const el = document.getElementById('tab-' + tab);
    if (!el) return;
    el.style.display = 'flex';
  }

  // === Exports ===
  window.RPT.loadFile = loadFile;
  window.RPT.applyFilters = applyFilters;
  window.RPT.renderKPIs = renderKPIs;
})();
