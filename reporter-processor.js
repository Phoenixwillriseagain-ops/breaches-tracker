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
        const lower = {};
        Object.keys(r).forEach(key => {
          lower[key.toLowerCase()] = key;
        });
        
        // Detect columns from header (first row)
        if (idx === 0) {
          Object.keys(lower).forEach(lowerKey => {
            const origKey = lower[lowerKey];
            
            // Ticket/Incident detection
            if (lowerKey.includes('ticket') || lowerKey.includes('incident')) {
              colMap.ticket = origKey;
            }
            // Week/Month/Date detection
            if ((lowerKey.includes('week') || lowerKey.includes('month')) || lowerKey.includes('date')) {
              if (!colMap.month) colMap.month = origKey;
            }
            // SLA detection
            if (lowerKey.includes('sla') || lowerKey.includes('code')) {
              colMap.sla = origKey;
            }
            // KM detection
            if (lowerKey.includes('km')) {
              colMap.kml = origKey;
            }
            // AOS/Portal detection
            if (lowerKey.includes('aos') || lowerKey.includes('portal')) {
              colMap.aos = origKey;
            }
            // Reason detection
            if (lowerKey.includes('reason') || lowerKey.includes('breach')) {
              colMap.reason = origKey;
            }
            // Language detection
            if (lowerKey.includes('lang')) {
              colMap.lang = origKey;
            }
            // Type/App detection
            if (lowerKey.includes('type') || lowerKey.includes('app')) {
              colMap.type = origKey;
            }
            // Status detection
            if (lowerKey.includes('status')) {
              colMap.status = origKey;
            }
            // Excluded detection
            if (lowerKey.includes('excluded')) {
              colMap.excluded = origKey;
            }
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
          excluded: normBoolLike(clean(r[colMap.excluded] || ''))
        };
        data.push(prow);
      });
    });

    console.log('Detected mappings:', colMap);
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
    const cleaned = lower.replace(/\D/g, '');
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
    const weekFilter = document.getElementById('week-filter')?.value || 'All';
    const slaFilter = document.getElementById('sla-filter')?.value || 'All';
    const langFilter = document.getElementById('lang-filter')?.value || 'All';
    const excludedFilter = document.getElementById('excluded-filter')?.value || 'All';

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
    const bpwCtx = document.getElementById('chart-breaches-week')?.getContext('2d');
    const slaCatCtx = document.getElementById('chart-sla-category')?.getContext('2d');
    const langCtx = document.getElementById('chart-language')?.getContext('2d');
    const excCtx = document.getElementById('chart-excluded')?.getContext('2d');

    if (bpwCtx) {
      const weekData = {};
      rptState.filtered.forEach(row => {
        weekData[row.month] = (weekData[row.month] || 0) + 1;
      });
      new Chart(bpwCtx, {
        type: 'bar',
        data: {
          labels: Object.keys(weekData),
          datasets: [{ label: 'Breaches', data: Object.values(weekData), backgroundColor: '#0099cc' }]
        }
      });
    }

    if (slaCatCtx) {
      const slaData = {};
      rptState.filtered.forEach(row => {
        slaData[row.sla] = (slaData[row.sla] || 0) + 1;
      });
      new Chart(slaCatCtx, {
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
      const th = document.createElement('thead');
      th.innerHTML = '<tr><th>Ticket</th><th>AOS</th><th>Status</th><th>Month</th></tr>';
      aosTable.appendChild(th);
      
      const tb = document.createElement('tbody');
      rptState.filtered.forEach((row, i) => {
        if (i >= MAX_DISPLAY_ROWS) return;
        const tr = document.createElement('tr');
        tr.innerHTML = `<td>${row.ticket}</td><td>${row.aos}</td><td>${row.status}</td><td>${row.month}</td>`;
        tb.appendChild(tr);
      });
      aosTable.appendChild(tb);
    }
  }

  window.addEventListener('load', function() {
    const fileInput = document.getElementById('file-input');
    if (fileInput) {
      fileInput.addEventListener('change', function(e) {
        loadFile(e.target.files[0]);
      });
    }
    
    document.getElementById('week-filter')?.addEventListener('change', function() {
      applyFilters();
      renderCharts();
      renderTables();
    });
    document.getElementById('sla-filter')?.addEventListener('change', function() {
      applyFilters();
      renderCharts();
      renderTables();
    });
    document.getElementById('lang-filter')?.addEventListener('change', function() {
      applyFilters();
      renderCharts();
      renderTables();
    });
    document.getElementById('excluded-filter')?.addEventListener('change', function() {
      applyFilters();
      renderCharts();
      renderTables();
    });
  });
})();
