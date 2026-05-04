// reporter-processor.js - Improved with smart column detection

(function() {
  'use strict';

  const { CONFIG } = window.BT || {};
  const { MAX_DISPLAY_ROWS } = CONFIG;

  const rptState = window.RPT = {
    allData: [],
    filtered: [],
    activeTab: 'overview',
    charts: {},
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
    if (['1', 'true', 'yes', 'y'].indexOf(s) !== -1) return 'Y';
    return 'N';
  }

  // Intelligent column finder
  function findColumn(row, keywords) {
    if (!Array.isArray(keywords)) keywords = [keywords];
    
    const rowKeys = Object.keys(row);
    
    // Exact match first
    for (let keyword of keywords) {
      if (row[keyword] !== undefined && row[keyword] !== null && row[keyword] !== '') {
        return row[keyword];
      }
    }
    
    // Case-insensitive exact match
    const keywordsLower = keywords.map(k => clean(k).toLowerCase());
    for (let rowKey of rowKeys) {
      const cleanedKey = clean(rowKey).toLowerCase();
      if (keywordsLower.includes(cleanedKey)) {
        const val = row[rowKey];
        if (val !== undefined && val !== null && val !== '') return val;
      }
    }
    
    // Partial/fuzzy match
    for (let keyword of keywords) {
      const searchTerm = clean(keyword).toLowerCase();
      for (let rowKey of rowKeys) {
        const cleanedKey = clean(rowKey).toLowerCase();
        if (cleanedKey.includes(searchTerm) || searchTerm.includes(cleanedKey)) {
          const val = row[rowKey];
          if (val !== undefined && val !== null && val !== '') return val;
        }
      }
    }
    
    return '';
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
    
    wb.SheetNames.forEach(function(sn) {
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[sn]);
      
      rows.forEach(function(r) {
        const ticket = findColumn(r, ['Incident Ticket', 'Ticket', 'ticket', 'Incident ID', 'ID']);
        if (!ticket) return;
        
        const monthVal = findColumn(r, ['Month', 'Date', 'Created', 'Week', 'month', 'date', 'week']);
        const slaVal = findColumn(r, ['SLA Code', 'SLA', 'SLA_Code', 'sla', 'KSL']);
        const km1Val = findColumn(r, ['KM-1', 'KM1', 'KM Code', 'km1', 'km-1']);
        const aosVal = findColumn(r, ['AOS Portal', 'AOS', 'Portal', 'aos', 'aos_portal']);
        const statusVal = findColumn(r, ['Status', 'State', 'status', 'state']);
        const breachTypeVal = findColumn(r, ['Breach Type', 'Type', 'Application', 'breach_type', 'type', 'application']);
        const langVal = findColumn(r, ['Language', 'Lang', 'language', 'lang', 'Language Code']);
        const excludedVal = findColumn(r, ['Excluded', 'Exclude', 'excluded']);
        const reasonVal = findColumn(r, ['Reason', 'Breach Reason', 'reason', 'breach_reason']);
        
        const row = {
          ticket: clean(ticket),
          month: normKey(monthVal, 'Unknown'),
          sla: normSla(slaVal),
          km1: normBoolLike(km1Val),
          aos: normAos(aosVal),
          status: normKey(statusVal, 'N/A'),
          breachType: normKey(breachTypeVal, 'Unknown'),
          language: normKey(langVal, 'Unknown'),
          excluded: normBoolLike(excludedVal),
          reason: normKey(reasonVal, 'Unknown'),
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
