// reporter-processor.js - Fixed column mapping, normalization, and filter dropdowns

(function() {
  'use strict';

  const { CONFIG } = window.BT || {};
  const { MAX_DISPLAY_ROWS } = CONFIG || { MAX_DISPLAY_ROWS: 100 };

  const rptState = window.RPT = {
    allData: [],
    filtered: [],
    activeTab: 'overview',
    charts: {},
    columnMap: {},
    uniqueValues: {}
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

  function extractWeek(v) {
    if (!v) return 'Unknown';
    const s = String(v).toLowerCase();
    // Extract week number if present
    const weekMatch = s.match(/w(\d+)|week[\s-]*(\d+)/);
    if (weekMatch) return 'W' + (weekMatch[1] || weekMatch[2]);
    // Extract month if present
    const monthMatch = s.match(/jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec|\d{1,2}[\/\-]\d{1,2}/);
    if (monthMatch) return monthMatch[0];
    // If it looks like a timestamp, extract just the date part
    if (s.includes('summer time') || s.includes('gmt')) {
      return s.split('(')[0].trim().split(' ').slice(0, 3).join(' ');
    }
    return clean(v).substring(0, 20);
  }

  function normSla(v) {
    if (!v) return 'Unknown';
    const lower = String(v).toLowerCase();
    
    // Check for SLA numbers/codes first
    if (lower === '1' || lower.includes('1.')) return 'SLA-1';
    if (lower === '2' || lower.includes('2.')) return 'SLA-2';
    if (lower === '3' || lower.includes('3.')) return 'SLA-3';
    if (lower === 'unknown' || lower === '') return 'Unknown';
    
    // Check for named SLAs
    if (lower.includes('critical')) return 'Critical';
    if (lower.includes('high')) return 'High';
    if (lower.includes('medium')) return 'Medium';
    if (lower.includes('low')) return 'Low';
    
    // Return as-is if it's already a code
    return clean(v).toUpperCase();
  }

  function normAos(v) {
    if (!v) return 'Unknown';
    const lower = String(v).toLowerCase();
    if (lower.includes('tick')) return 'Ticket';
    if (lower.includes('incident')) return 'Incident';
    if (lower.includes('request')) return 'Request';
    return clean(v);
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
        if (idx === 0) {
          Object.keys(r).forEach(key => {
            const lower = key.toLowerCase();
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
          month: extractWeek(r[colMap.month]),
          sla: normSla(r[colMap.sla]),
          kml: normBoolLike(r[colMap.kml]),
          aos: normAos(r[colMap.aos]),
          status: normKey(clean(r[colMap.status] || ''), 'N/A'),
          breachType: normKey(clean(r[colMap.type] || ''), 'Unknown'),
          language: normKey(clean(r[colMap.lang] || ''), 'Unknown'),
          reason: normKey(clean(r[colMap.reason] || ''), 'Unknown'),
          excluded: normBoolLike(clean(r[colMap.excluded] || ''))
        };
        data.push(prow);
      });
    });

    console.log('Loaded records:', data.length);
    rptState.columnMap = colMap;
    rptState.allData = data;
    rptState.filtered = data.slice();
    
    // Extract unique values for dropdowns
    populateUniqueValues();
    updateFilterDropdowns();
    applyFilters();
    renderCharts();
    renderTables();
  }

  function populateUniqueValues() {
    rptState.uniqueValues = {
      months: [],
      slas: [],
      languages: [],
      excluded: ['Y', 'N']
    };
    
    rptState.allData.forEach(row => {
      if (rptState.uniqueValues.months.indexOf(row.month) === -1) {
        rptState.uniqueValues.months.push(row.month);
      }
      if (rptState.uniqueValues.slas.indexOf(row.sla) === -1) {
        rptState.uniqueValues.slas.push(row.sla);
      }
      if (rptState.uniqueValues.languages.indexOf(row.language) === -1) {
        rptState.uniqueValues.languages.push(row.language);
      }
    });
    
    rptState.uniqueValues.months.sort();
    rptState.uniqueValues.slas.sort();
    rptState.uniqueValues.languages.sort();
  }

  function updateFilterDropdowns() {
    const weekSelect = document.getElementById('filter-week');
    const slaSelect = document.getElementById('filter-sla');
    const langSelect = document.getElementById('filter-lang');
    const exclSelect = document.getElementById('filter-excl');
    
    if (weekSelect) {
      const html = '<option value="All">All</option>' + 
        rptState.uniqueValues.months.map(m => `<option value="${m}">${m}</option>`).join('');
      weekSelect.innerHTML = html;
    }
    
    if (slaSelect) {
      const html = '<option value="All">All</option>' + 
        rptState.uniqueValues.slas.map(s => `<option value="${s}">${s}</option>`).join('');
      slaSelect.innerHTML = html;
    }
    
    if (langSelect) {
      const html = '<option value="All">All</option>' + 
        rptState.uniqueValues.languages.map(l => `<option value="${l}">${l}</option>`).join('');
      langSelect.innerHTML = html;
    }
    
    if (exclSelect) {
      exclSelect.innerHTML = '<option value="All">All</option><option value="Y">Yes</option><option value="N">No</option>';
    }
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
    if (bpwCtx && rptState.filtered.length > 0) {
      const weekData = {};
      rptState.filtered.forEach(row => {
        weekData[row.month] = (weekData[row.month] || 0) + 1;
      });
      const labels = Object.keys(weekData).sort();
      rptState.charts['week'] = new Chart(bpwCtx, {
        type: 'bar',
        data: {
          labels: labels,
          datasets: [{ label: 'Breaches', data: labels.map(l => weekData[l]), backgroundColor: '#0099cc' }]
        },
        options: {
          responsive: true,
          plugins: { legend: { display: true } }
        }
      });
    }

    // SLA Category
    const slaCatCtx = document.getElementById('chart-sla')?.getContext('2d');
    if (slaCatCtx && rptState.charts['sla']) {
      rptState.charts['sla'].destroy();
    }
    if (slaCatCtx && rptState.filtered.length > 0) {
      const slaData = {};
      rptState.filtered.forEach(row => {
        slaData[row.sla] = (slaData[row.sla] || 0) + 1;
      });
      const colors = ['#ff6b6b', '#ffa500', '#4ecdc4', '#45b7d1', '#96ceb4', '#dda15e'];
      rptState.charts['sla'] = new Chart(slaCatCtx, {
        type: 'pie',
        data: {
          labels: Object.keys(slaData),
          datasets: [{ data: Object.values(slaData), backgroundColor: colors }]
        }
      });
    }
  }

  function renderTables() {
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
    }
    
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
