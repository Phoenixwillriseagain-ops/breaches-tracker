// reporter-processor.js - Enhanced processor with AOS filtering and callback support
(function() {
  'use strict';
  const { CONFIG } = window.BT || {};
  const { MAX_DISPLAY_ROWS } = CONFIG || { MAX_DISPLAY_ROWS: 100 };

  // Column name mapping for actual Google Sheets columns
  const COLUMN_MAPPING = {
    'Incident Ticket': 'ticket',
    'Element eingegangen am': 'dateReceived',
    'Element gelöst am': 'dateResolved',
    'DATE_CLOSED': 'dateClosed',
    'Element Status': 'status',
    'Ticket Gruppe': 'ticketGroup',
    'SLA Kennzahl': 'slaCode',
    'SLA Prüfung': 'slaReview',
    'SLA Prüfung Name': 'slaName',
    'SLA Prüfung ist OK/NOK?': 'slaResult'
  };

  window.RPT = {
    allData: [],
    filtered: [],
    aosFiltered: [],
    charts: {},
    columnMap: {},
    uniqueValues: {},
    loadFile: function(file) { loadFile(file); },
    clearFilters: function() {
      document.querySelectorAll('.filter-group select').forEach(s => s.value = 'All');
      applyFilters();
      renderCharts();
      renderTables();
    },
    resetToUpload: function() { location.reload(); },
    exportFiltered: function() { alert('Export feature coming soon'); },
    exportReport: function() { alert('Export feature coming soon'); },
    // Callback function for custom data processing
    processDataCallback: function(data, callback) {
      // This allows external systems to hook into data processing
      if (typeof callback === 'function') {
        return callback(data);
      }
      return data;
    }
  };

  function clean(v) {
    return String(v == null ? '' : v).trim();
  }

  function extractWeek(v) {
    if (!v) return 'Unknown';
    const s = String(v).toLowerCase();
    const weekMatch = s.match(/w(\d+)|week[\s-]*(\d+)/);
    if (weekMatch) return 'W' + (weekMatch[1] || weekMatch[2]);
    const monthMatch = s.match(/jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec|\d{1,2}[\/\-]\d{1,2}/);
    if (monthMatch) return monthMatch[0];
    if (s.includes('summer time') || s.includes('gmt')) {
      return s.split('(')[0].trim().split(' ').slice(0, 3).join(' ');
    }
    return clean(v).substring(0, 20);
  }

  function normSla(v) {
    if (!v) return 'Unknown';
    const lower = String(v).toLowerCase();
    if (lower === '1' || lower.includes('1.')) return 'SLA-1';
    if (lower === '2' || lower.includes('2.')) return 'SLA-2';
    if (lower === '3' || lower.includes('3.')) return 'SLA-3';
    if (lower === 'unknown' || lower === '') return 'Unknown';
    if (lower.includes('critical')) return 'Critical';
    if (lower.includes('high')) return 'High';
    if (lower.includes('medium')) return 'Medium';
    if (lower.includes('low')) return 'Low';
    return clean(v).toUpperCase();
  }

  function normBoolLike(v) {
    const s = clean(v).toLowerCase();
    if (!s) return 'N';
    if (['t', 'true', 'yes', 'y', '1'].indexOf(s) !== -1) return 'Y';
    return 'N';
  }

  // Helper function to detect AOS Portal Issues
  function isAosIssue(ticketGroup) {
    if (!ticketGroup) return false;
    const lower = String(ticketGroup).toLowerCase();
    return lower.includes('aos') || lower.includes('portal');
  }

  // Callback-enabled remote data processor
  function remoteDataProcessor(rowData, callback) {
    const processed = {
      ticket: clean(rowData.ticket || ''),
      dateReceived: clean(rowData.dateReceived || ''),
      dateResolved: clean(rowData.dateResolved || ''),
      dateClosed: clean(rowData.dateClosed || ''),
      status: clean(rowData.status || 'N/A'),
      ticketGroup: clean(rowData.ticketGroup || 'Unknown'),
      month: extractWeek(rowData.month),
      sla: normSla(rowData.sla),
      isAos: isAosIssue(rowData.ticketGroup),
      language: clean(rowData.language || 'Unknown'),
      reason: clean(rowData.reason || 'Unknown'),
      excluded: normBoolLike(rowData.excluded)
    };
    
    // If callback provided, use it for custom processing
    if (typeof callback === 'function') {
      return callback(processed);
    }
    return processed;
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
            if (lower.includes('gruppe') || lower.includes('group')) colMap.ticketGroup = key;
            if (lower.includes('reason') || lower.includes('breach') || lower.includes('type')) colMap.reason = key;
            if (lower.includes('lang')) colMap.lang = key;
            if (lower.includes('app')) colMap.type = key;
            if (lower.includes('status')) colMap.status = key;
            if (lower.includes('excluded')) colMap.excluded = key;
            if (lower.includes('eingegangen') || lower.includes('received')) colMap.dateReceived = key;
            if (lower.includes('gelöst') || lower.includes('resolved')) colMap.dateResolved = key;
            if (lower.includes('closed')) colMap.dateClosed = key;
          });
        }
        const ticket = clean(r[colMap.ticket] || '');
        if (!ticket) return;
        
        const rawRow = {
          ticket: ticket,
          month: extractWeek(r[colMap.month]),
          sla: normSla(r[colMap.sla]),
          kml: normBoolLike(r[colMap.kml]),
          ticketGroup: clean(r[colMap.ticketGroup] || 'Unknown'),
          status: clean(r[colMap.status] || 'N/A'),
          breachType: clean(r[colMap.type] || 'Unknown'),
          language: clean(r[colMap.lang] || 'Unknown'),
          reason: clean(r[colMap.reason] || 'Unknown'),
          dateReceived: clean(r[colMap.dateReceived] || ''),
          dateResolved: clean(r[colMap.dateResolved] || ''),
          dateClosed: clean(r[colMap.dateClosed] || ''),
          excluded: normBoolLike(r[colMap.excluded])
        };
        
        // Process through callback-enabled remote processor
        const prow = remoteDataProcessor(rawRow, null);
        data.push(prow);
      });
    });

    console.log('Loaded', data.length, 'records');
    window.RPT.columnMap = colMap;
    window.RPT.allData = data;
    window.RPT.filtered = data.slice();
    
    // Filter AOS Portal Issues separately
    window.RPT.aosFiltered = data.filter(row => row.isAos === true);
    console.log('AOS Portal Issues found:', window.RPT.aosFiltered.length);
    
    populateUniqueValues();
    updateFilterDropdowns();
    applyFilters();
    renderCharts();
    renderTables();
  }

  function populateUniqueValues() {
    window.RPT.uniqueValues = {
      months: [],
      slas: [],
      languages: [],
      ticketGroups: [],
      excluded: ['Y', 'N']
    };
    window.RPT.allData.forEach(row => {
      if (window.RPT.uniqueValues.months.indexOf(row.month) === -1) {
        window.RPT.uniqueValues.months.push(row.month);
      }
      if (window.RPT.uniqueValues.slas.indexOf(row.sla) === -1) {
        window.RPT.uniqueValues.slas.push(row.sla);
      }
      if (window.RPT.uniqueValues.languages.indexOf(row.language) === -1) {
        window.RPT.uniqueValues.languages.push(row.language);
      }
      if (window.RPT.uniqueValues.ticketGroups.indexOf(row.ticketGroup) === -1) {
        window.RPT.uniqueValues.ticketGroups.push(row.ticketGroup);
      }
    });
    window.RPT.uniqueValues.months.sort();
    window.RPT.uniqueValues.slas.sort();
    window.RPT.uniqueValues.languages.sort();
    window.RPT.uniqueValues.ticketGroups.sort();
  }

  function updateFilterDropdowns() {
    const weekSelect = document.getElementById('filter-week');
    const slaSelect = document.getElementById('filter-sla');
    const langSelect = document.getElementById('filter-lang');
    const exclSelect = document.getElementById('filter-excl');

    if (weekSelect) {
      const html = '<option value="All">All</option>' +
        window.RPT.uniqueValues.months.map(m => `<option value="${m}">${m}</option>`).join('');
      weekSelect.innerHTML = html;
    }
    if (slaSelect) {
      const html = '<option value="All">All</option>' +
        window.RPT.uniqueValues.slas.map(s => `<option value="${s}">${s}</option>`).join('');
      slaSelect.innerHTML = html;
    }
    if (langSelect) {
      const html = '<option value="All">All</option>' +
        window.RPT.uniqueValues.languages.map(l => `<option value="${l}">${l}</option>`).join('');
      langSelect.innerHTML = html;
    }
    if (exclSelect) {
      exclSelect.innerHTML = '<option value="All">All</option><option value="Y">Excluded</option><option value="N">Counted</option>';
    }
  }

  function applyFilters() {
    const weekFilter = document.getElementById('filter-week')?.value || 'All';
    const slaFilter = document.getElementById('filter-sla')?.value || 'All';
    const langFilter = document.getElementById('filter-lang')?.value || 'All';
    const excludedFilter = document.getElementById('filter-excl')?.value || 'All';

    window.RPT.filtered = window.RPT.allData.filter(row => {
      if (weekFilter !== 'All' && row.month !== weekFilter) return false;
      if (slaFilter !== 'All' && row.sla !== slaFilter) return false;
      if (langFilter !== 'All' && row.language !== langFilter) return false;
      if (excludedFilter === 'Y' && row.excluded !== 'Y') return false;
      if (excludedFilter === 'N' && row.excluded !== 'N') return false;
      return true;
    });
  }

  function renderCharts() {
    const colors = ['#ff6b6b', '#ffa500', '#4ecdc4', '#45b7d1', '#96ceb4', '#dda15e', '#a8dadc', '#f1faee'];
    const data = window.RPT.filtered;
    if (!data.length) return;

    // Chart 1: Breaches per Week
    const bpwCtx = document.getElementById('chart-week')?.getContext('2d');
    if (bpwCtx) {
      if (window.RPT.charts['week']) window.RPT.charts['week'].destroy();
      const weekData = {};
      data.forEach(row => {
        weekData[row.month] = (weekData[row.month] || 0) + 1;
      });
      const labels = Object.keys(weekData).sort();
      window.RPT.charts['week'] = new Chart(bpwCtx, {
        type: 'bar',
        data: {
          labels: labels,
          datasets: [{
            label: 'Breaches',
            data: labels.map(l => weekData[l]),
            backgroundColor: '#0099cc'
          }]
        }
      });
    }

    // Chart 2: SLA Category
    const slaCatCtx = document.getElementById('chart-sla')?.getContext('2d');
    if (slaCatCtx) {
      if (window.RPT.charts['sla']) window.RPT.charts['sla'].destroy();
      const slaData = {};
      data.forEach(row => {
        slaData[row.sla] = (slaData[row.sla] || 0) + 1;
      });
      window.RPT.charts['sla'] = new Chart(slaCatCtx, {
        type: 'pie',
        data: {
          labels: Object.keys(slaData),
          datasets: [{
            data: Object.values(slaData),
            backgroundColor: colors
          }]
        }
      });
    }

    // Chart 3: Language Split
    const langCtx = document.getElementById('chart-lang')?.getContext('2d');
    if (langCtx) {
      if (window.RPT.charts['lang']) window.RPT.charts['lang'].destroy();
      const langData = {};
      data.forEach(row => {
        langData[row.language] = (langData[row.language] || 0) + 1;
      });
      window.RPT.charts['lang'] = new Chart(langCtx, {
        type: 'doughnut',
        data: {
          labels: Object.keys(langData),
          datasets: [{
            data: Object.values(langData),
            backgroundColor: colors
          }]
        }
      });
    }

    // Chart 4: Excluded vs Counted
    const exclCtx = document.getElementById('chart-excl')?.getContext('2d');
    if (exclCtx) {
      if (window.RPT.charts['excl']) window.RPT.charts['excl'].destroy();
      const exclData = { 'Excluded': 0, 'Counted': 0 };
      data.forEach(row => {
        if (row.excluded === 'Y') exclData['Excluded']++;
        else exclData['Counted']++;
      });
      window.RPT.charts['excl'] = new Chart(exclCtx, {
        type: 'pie',
        data: {
          labels: Object.keys(exclData),
          datasets: [{
            data: Object.values(exclData),
            backgroundColor: ['#ff6b6b', '#51cf66']
          }]
        }
      });
    }
  }

  function renderTables() {
    const data = window.RPT.filtered;

    // AOS Portal Issues Table - filtered to show ONLY AOS issues
    const aosTable = document.getElementById('aos-table');
    if (aosTable) {
      aosTable.innerHTML = '';
      const aosData = data.filter(row => row.isAos === true);
      
      if (aosData.length > 0) {
        const thead = document.createElement('thead');
        thead.innerHTML = '<tr><th>Ticket</th><th>Group</th><th>Status</th><th>SLA</th><th>Week</th></tr>';
        aosTable.appendChild(thead);
        const tbody = document.createElement('tbody');
        aosData.slice(0, MAX_DISPLAY_ROWS).forEach(row => {
          const tr = document.createElement('tr');
          tr.innerHTML = `<td>${clean(row.ticket)}</td><td>${clean(row.ticketGroup)}</td><td>${clean(row.status)}</td><td>${clean(row.sla)}</td><td>${clean(row.month)}</td>`;
          tbody.appendChild(tr);
        });
        aosTable.appendChild(tbody);
      } else {
        const emptyRow = document.createElement('tr');
        emptyRow.innerHTML = '<td colspan="5" style="text-align: center; padding: 20px;">No AOS Portal Issues found in current filters</td>';
        aosTable.appendChild(emptyRow);
      }
    }

    // KM-1 Table
    const km1Table = document.getElementById('km1-table');
    if (km1Table) {
      km1Table.innerHTML = '';
      const km1Data = data.filter(r => r.kml === 'Y');
      if (km1Data.length > 0) {
        const thead = document.createElement('thead');
        thead.innerHTML = '<tr><th>Ticket</th><th>Reason</th><th>Language</th><th>Week</th></tr>';
        km1Table.appendChild(thead);
        const tbody = document.createElement('tbody');
        km1Data.slice(0, MAX_DISPLAY_ROWS).forEach(row => {
          const tr = document.createElement('tr');
          tr.innerHTML = `<td>${clean(row.ticket)}</td><td>${clean(row.reason)}</td><td>${clean(row.language)}</td><td>${clean(row.month)}</td>`;
          tbody.appendChild(tr);
        });
        km1Table.appendChild(tbody);
      } else {
        const emptyRow = document.createElement('tr');
        emptyRow.innerHTML = '<td colspan="4" style="text-align: center; padding: 20px;">No KM-1 data available</td>';
        km1Table.appendChild(emptyRow);
      }
    }
  }

  window.addEventListener('load', function() {
    const fileInput = document.getElementById('reporter-file-input');
    if (fileInput) {
      fileInput.addEventListener('change', function(e) {
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
