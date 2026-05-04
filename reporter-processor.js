// reporter-processor.js - Fixed column mapping with correct element IDs

(function() {
  'use strict';

  const { CONFIG } = window.BT || {};
  const { MAX_DISPLAY_ROWS } = CONFIG || { MAX_DISPLAY_ROWS: 100 };

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
    if (!s) return 'N';
    if (['t', 'true', 'yes', 'y'].indexOf(s) !== -1) return 'Y';
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
        // Build lowercase header map on first iteration
        if (idx === 0) {
          Object.keys(r).forEach(key => {
            const lower = key.toLowerCase();
            console.log('Header column:', key, '->', lower);
            
            if (lower.includes('ticket') || lower.includes('incident')) colMap.ticket = key;
            if (lower.includes('week') || lower.includes('month') || lower.includes('date')) colMap.month = key;
            if (lower.includes('sla') || lower.includes('code')) colMap.sla = key;
            if (lower.includes('km')) colMap.kml = key;
            if (lower.includes('aos') || lower.includes('portal')) colMap.aos = key;
            if (lower.includes('reason') || lower.includes('breach') || lower.includes('type')) colMap.reason = key;
            if (lower.includes('lang')) colMap.lang = key;
            if (lower.includes('app')) colMap.type = key;
            if (lower.includes('status')) colMap.status = key;
            if (lower.includes('excluded')) colMap.excluded = key;
          });
        }
        
        const ticket = clean(r[colMap.ticket] || '');
        if (!ticket) return;

        const prow = {
          ticket: ticket,
          month: normKey(clean(r[colMap.month] || ''), 'Unknown'),
          sla: normSla(clean(r[colMap.sla] || '')),
          kml: normBoolLike(clean(r[colMap.kml] || '')),
          aos: normAos(clean(r[colMap.aos] || '')),
          status: normKey(clean(r[colMap.status] || ''), 'N/A'),
          breachType: normKey(clean(r[colMap.type] || ''), 'Unknown'),
          language: normKey(clean(r[colMap.lang] || ''), 'Unknown'),
          reason: normKey(clean(r[colMap.reason] || ''), 'Unknown'),
          excluded: normBoolLike(clean(r[colMap.excluded] || ''))
        };
        data.push(prow);
      });
    });

    console.log('Detected mappings:', colMap);
    console.log('Loaded records:', data.length);
    rptState.columnMap = colMap;
    rptState.allData = data;
    rptState.filtered = data.slice();
    applyFilters();
    renderCharts();
    renderTables();
  }

  function normSla(v) {
    if (!v) return 'Unknown';
    const lower = v.toLowerCase();
    if (lower.includes('critical')) return 'Critical';
    if (lower.includes('high')) return 'High';
    if (lower.includes('medium')) return 'Medium';
    if (lower.includes('low')) return 'Low';
    return v.toUpperCase();
  }

  function normAos(v) {
    if (!v) return 'Unknown';
    const lower = v.toLowerCase();
    if (lower.includes('tick')) return 'Ticket';
    if (lower.includes('incident')) return 'Incident';
    if (lower.includes('request')) return 'Request';
    return v;
  }

  function applyFilters() {
    const weekFilter = document.getElementById('filter-week')?.value || 'All';
    const slaFilter = document.getElementById('filter-sla')?.value || 'All';
    const langFilter = document.getElementById('filter-lang')?.value || 'All';
    const excludedFilter = document.getElementById('filter-excl')?.value || 'All';

    rptState.filtered = rptState.allData.filter(row => {
      if (weekFilter !== 'All' && row.month !== weekFilter) return false;
      if (slaFilter !== 'All' && row.sla !== slaFilter) return false;
      if (langFilter !== 'All' && row.language !== langFilter) return false;
      if (excludedFilter === 'Y' && row.excluded !== 'Y') return false;
      if (excludedFilter === 'N' && row.excluded !== 'N') return false;
      return true;
    });
  }

  function renderCharts() {
    // Breaches per Week
    const bpwCtx = document.getElementById('chart-week')?.getContext('2d');
    if (bpwCtx && rptState.charts['week']) {
      rptState.charts['week'].destroy();
    }
    if (bpwCtx) {
      const weekData = {};
      rptState.filtered.forEach(row => {
        weekData[row.month] = (weekData[row.month] || 0) + 1;
      });
      rptState.charts['week'] = new Chart(bpwCtx, {
        type: 'bar',
        data: {
          labels: Object.keys(weekData),
          datasets: [{ label: 'Breaches', data: Object.values(weekData), backgroundColor: '#0099cc' }]
        }
      });
    }

    // SLA Category
    const slaCatCtx = document.getElementById('chart-sla')?.getContext('2d');
    if (slaCatCtx && rptState.charts['sla']) {
      rptState.charts['sla'].destroy();
    }
    if (slaCatCtx) {
      const slaData = {};
      rptState.filtered.forEach(row => {
        slaData[row.sla] = (slaData[row.sla] || 0) + 1;
      });
      rptState.charts['sla'] = new Chart(slaCatCtx, {
        type: 'pie',
        data: {
          labels: Object.keys(slaData),
          datasets: [{ data: Object.values(slaData), backgroundColor: ['#ff6b6b', '#ffa500', '#4ecdc4', '#45b7d1'] }]
        }
      });
    }
  }

  function renderTables() {
    // AOS Portal Issues
    const aosTable = document.getElementById('aos-table');
    if (aosTable) {
      aosTable.innerHTML = '';
      const thead = document.createElement('thead');
      thead.innerHTML = '<tr><th>Ticket</th><th>AOS</th><th>Status</th><th>Month</th></tr>';
      aosTable.appendChild(thead);
      
      const tbody = document.createElement('tbody');
      rptState.filtered.forEach((row, i) => {
        if (i >= MAX_DISPLAY_ROWS) return;
        const tr = document.createElement('tr');
        tr.innerHTML = `<td>${row.ticket}</td><td>${row.aos}</td><td>${row.status}</td><td>${row.month}</td>`;
        tbody.appendChild(tr);
      });
      aosTable.appendChild(tbody);
    }
  }

  window.addEventListener('load', function() {
    console.log('Reporter processor loaded');
    
    const fileInput = document.getElementById('reporter-file-input');
    if (fileInput) {
      fileInput.addEventListener('change', function(e) {
        console.log('File selected:', e.target.files[0]?.name);
        loadFile(e.target.files[0]);
      });
    } else {
      console.warn('File input element not found');
    }
    
    // Filter listeners
    ['filter-week', 'filter-sla', 'filter-lang', 'filter-excl'].forEach(filterId => {
      const elem = document.getElementById(filterId);
      if (elem) {
        elem.addEventListener('change', function() {
          applyFilters();
          renderCharts();
          renderTables();
        });
      }
    });
  });
})();
