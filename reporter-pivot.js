// reporter-pivot.js — Semi-dynamic pivot table for the Breaches Reporter
// Plug-in module: called by renderTables() after data is ready.
// Reads window.RPT.filtered. Uses countUniqueTickets() for all cell values.

(function () {
  'use strict';

  /* ---- Field definitions ---- */
  var PIVOT_FIELDS = [
    { key: 'sla',      label: 'SLA' },
    { key: 'week',     label: 'Week' },
    { key: 'queue',    label: 'Queue' },
    { key: 'tool',     label: 'Tool' },
    { key: 'language', label: 'Language' },
    { key: 'sheet',    label: 'Sheet' },
    { key: 'excluded', label: 'Excluded' },
  ];

  /* ---- State ---- */
  var _rowField = 'week';
  var _colField = 'sla';

  /* ---- Public entry point called by renderTables() ---- */
  window.renderPivot = function () {
    var container = document.getElementById('pivot-container');
    if (!container) return;
    container.innerHTML = _buildPivotHTML(window.RPT ? window.RPT.filtered : []);
    _bindPivotEvents();
  };

  /* ================================================================
     BUILD
     ================================================================ */
  function _buildPivotHTML(data) {
    /* Field selector row */
    var rowOpts = PIVOT_FIELDS.map(function (f) {
      return '<option value="' + f.key + '"' + (f.key === _rowField ? ' selected' : '') + '>' + f.label + '</option>';
    }).join('');
    var colOpts = PIVOT_FIELDS.map(function (f) {
      return '<option value="' + f.key + '"' + (f.key === _colField ? ' selected' : '') + '>' + f.label + '</option>';
    }).join('');

    var controls =
      '<div class="pvt-controls">' +
        '<label class="pvt-ctrl-lbl">Rows&nbsp;&nbsp;' +
          '<select id="pvt-row-sel" class="pvt-sel">' + rowOpts + '</select>' +
        '</label>' +
        '<span class="pvt-ctrl-sep">&#215;</span>' +
        '<label class="pvt-ctrl-lbl">Columns&nbsp;&nbsp;' +
          '<select id="pvt-col-sel" class="pvt-sel">' + colOpts + '</select>' +
        '</label>' +
        '<button class="pvt-export-btn" id="pvt-export-btn" title="Export pivot to Excel">&#8659; Export Pivot</button>' +
      '</div>';

    /* Guard: same field selected for both axes */
    if (_rowField === _colField) {
      return '<div class="sc pvt-wrap">' +
        '<div class="sc-head"><h2>Pivot Table</h2></div>' +
        '<div class="sc-body">' + controls +
        '<p class="pvt-msg pvt-warn">&#9888;&#xFE0E; Row and Column fields must be different.</p>' +
        '</div></div>';
    }

    /* Collect axis values */
    var rowVals = _uniqueVals(data, _rowField);
    var colVals = _uniqueVals(data, _colField);

    if (!data.length || !rowVals.length || !colVals.length) {
      return '<div class="sc pvt-wrap">' +
        '<div class="sc-head"><h2>Pivot Table</h2></div>' +
        '<div class="sc-body">' + controls +
        '<p class="pvt-msg">No data available for the current filter.</p>' +
        '</div></div>';
    }

    /* Build lookup: rowVal -> colVal -> [rows] */
    var matrix = {};
    rowVals.forEach(function (rv) {
      matrix[rv] = {};
      colVals.forEach(function (cv) { matrix[rv][cv] = []; });
    });
    data.forEach(function (r) {
      var rv = r[_rowField] || '(blank)';
      var cv = r[_colField] || '(blank)';
      if (matrix[rv] && matrix[rv][cv]) matrix[rv][cv].push(r);
    });

    /* Row totals & column totals */
    var rowTotals = {};
    rowVals.forEach(function (rv) {
      var allRows = [];
      colVals.forEach(function (cv) { allRows = allRows.concat(matrix[rv][cv]); });
      rowTotals[rv] = _countUniq(allRows);
    });
    var colTotals = {};
    colVals.forEach(function (cv) {
      var allRows = [];
      rowVals.forEach(function (rv) { allRows = allRows.concat(matrix[rv][cv]); });
      colTotals[cv] = _countUniq(allRows);
    });
    var grandTotal = _countUniq(data);

    /* Max value (for heat shading) */
    var maxCell = 0;
    rowVals.forEach(function (rv) {
      colVals.forEach(function (cv) {
        var v = _countUniq(matrix[rv][cv]);
        if (v > maxCell) maxCell = v;
      });
    });
    if (maxCell === 0) maxCell = 1;

    /* Render table */
    var colFieldLabel = _fieldLabel(_colField);
    var rowFieldLabel = _fieldLabel(_rowField);

    /* Header row */
    var thead = '<thead><tr>' +
      '<th class="pvt-th pvt-th-corner">' + _esc(rowFieldLabel) + ' \ ' + _esc(colFieldLabel) + '</th>' +
      colVals.map(function (cv) { return '<th class="pvt-th pvt-th-col">' + _esc(cv) + '</th>'; }).join('') +
      '<th class="pvt-th pvt-th-total">Total</th>' +
      '</tr></thead>';

    /* Body rows */
    var tbody = '<tbody>' + rowVals.map(function (rv) {
      var cells = colVals.map(function (cv) {
        var v = _countUniq(matrix[rv][cv]);
        return '<td class="pvt-cell" style="' + _heatStyle(v, maxCell) + '" title="' +
          _esc(rowFieldLabel) + ': ' + _esc(rv) + ' | ' + _esc(colFieldLabel) + ': ' + _esc(cv) + ' — ' + v + ' unique ticket(s)">' +
          (v > 0 ? v : '') + '</td>';
      }).join('');
      return '<tr>' +
        '<td class="pvt-row-lbl">' + _esc(rv) + '</td>' +
        cells +
        '<td class="pvt-total-cell pvt-row-total">' + rowTotals[rv] + '</td>' +
      '</tr>';
    }).join('') +
    /* Grand total row */
    '<tr class="pvt-grand-row">' +
      '<td class="pvt-grand-lbl">Grand Total</td>' +
      colVals.map(function (cv) {
        return '<td class="pvt-total-cell pvt-col-total">' + colTotals[cv] + '</td>';
      }).join('') +
      '<td class="pvt-total-cell pvt-grand-total">' + grandTotal + '</td>' +
    '</tr>' +
    '</tbody>';

    var badge = rowVals.length + ' row' + (rowVals.length !== 1 ? 's' : '') +
      ' &bull; ' + colVals.length + ' col' + (colVals.length !== 1 ? 's' : '') +
      ' &bull; ' + grandTotal + ' unique ticket(s)';

    return '<div class="sc pvt-wrap">' +
      '<div class="sc-head"><h2>Pivot Table</h2><span class="badge">' + badge + '</span></div>' +
      '<div class="sc-body">' +
        controls +
        '<div class="pvt-table-scroll">' +
          '<table class="pvt-table">' + thead + tbody + '</table>' +
        '</div>' +
        '<p class="pvt-note">&#9432; Values = unique Incident Tickets per cell. Cells with 0 are left blank. Heat shading is relative to the highest cell value.</p>' +
      '</div>' +
    '</div>';
  }

  /* ================================================================
     EVENTS
     ================================================================ */
  function _bindPivotEvents() {
    var rowSel = document.getElementById('pvt-row-sel');
    var colSel = document.getElementById('pvt-col-sel');
    var expBtn = document.getElementById('pvt-export-btn');

    if (rowSel) rowSel.addEventListener('change', function () {
      _rowField = this.value;
      window.renderPivot();
    });
    if (colSel) colSel.addEventListener('change', function () {
      _colField = this.value;
      window.renderPivot();
    });
    if (expBtn) expBtn.addEventListener('click', _exportPivot);
  }

  /* ================================================================
     EXPORT
     ================================================================ */
  function _exportPivot() {
    if (typeof XLSX === 'undefined') { alert('SheetJS not loaded — cannot export.'); return; }
    var data = window.RPT ? window.RPT.filtered : [];
    if (!data.length) { alert('No data to export.'); return; }
    if (_rowField === _colField) { alert('Row and Column fields must be different.'); return; }

    var rowVals = _uniqueVals(data, _rowField);
    var colVals = _uniqueVals(data, _colField);

    /* Build matrix */
    var matrix = {};
    rowVals.forEach(function (rv) {
      matrix[rv] = {};
      colVals.forEach(function (cv) { matrix[rv][cv] = []; });
    });
    data.forEach(function (r) {
      var rv = r[_rowField] || '(blank)';
      var cv = r[_colField] || '(blank)';
      if (matrix[rv] && matrix[rv][cv]) matrix[rv][cv].push(r);
    });

    var rowFieldLabel = _fieldLabel(_rowField);
    var colFieldLabel = _fieldLabel(_colField);

    /* AoA (array of arrays) for SheetJS */
    var aoa = [];
    /* Header */
    aoa.push([rowFieldLabel + ' \\ ' + colFieldLabel].concat(colVals).concat(['Total']));
    /* Data rows */
    rowVals.forEach(function (rv) {
      var rowTotal = 0;
      var cells = colVals.map(function (cv) {
        var v = _countUniq(matrix[rv][cv]);
        rowTotal += v;
        return v;
      });
      aoa.push([rv].concat(cells).concat([rowTotal]));
    });
    /* Grand total row */
    var grandRow = ['Grand Total'].concat(colVals.map(function (cv) {
      var allRows = [];
      rowVals.forEach(function (rv) { allRows = allRows.concat(matrix[rv][cv]); });
      return _countUniq(allRows);
    })).concat([_countUniq(data)]);
    aoa.push(grandRow);

    var ws = XLSX.utils.aoa_to_sheet(aoa);
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Pivot');
    var fname = 'Pivot_' + rowFieldLabel + '_x_' + colFieldLabel + '_' +
      new Date().toISOString().slice(0, 10).replace(/-/g, '') + '.xlsx';
    XLSX.writeFile(wb, fname);
  }

  /* ================================================================
     HELPERS
     ================================================================ */
  function _uniqueVals(data, field) {
    var seen = {}, vals = [];
    data.forEach(function (r) {
      var v = r[field] || '(blank)';
      if (!seen[v]) { seen[v] = 1; vals.push(v); }
    });
    /* Numeric sort for week; alpha for everything else */
    if (field === 'week') vals.sort(function (a, b) { return Number(a) - Number(b) || a.localeCompare(b); });
    else vals.sort(function (a, b) { return String(a).localeCompare(String(b)); });
    return vals;
  }

  function _countUniq(rows) {
    if (window.RPT && window.RPT.countUniqueTickets) return window.RPT.countUniqueTickets(rows);
    var s = new Set();
    (rows || []).forEach(function (r) { if (r.ticket) s.add(r.ticket); });
    return s.size;
  }

  function _fieldLabel(key) {
    var f = PIVOT_FIELDS.filter(function (x) { return x.key === key; })[0];
    return f ? f.label : key;
  }

  function _heatStyle(val, maxVal) {
    if (val === 0) return 'background:var(--pvt-zero,#f0eeea);color:var(--faint,#bab9b4)';
    var ratio = Math.pow(val / maxVal, 0.6);
    var BR = 247, BG = 246, BB = 242;
    var TR = 1,   TG = 105, TB = 111;
    var r = Math.round(BR + ratio * (TR - BR));
    var g = Math.round(BG + ratio * (TG - BG));
    var b = Math.round(BB + ratio * (TB - BB));
    var lum = 0.299 * r + 0.587 * g + 0.114 * b;
    return 'background:rgb(' + r + ',' + g + ',' + b + ');color:' + (lum < 140 ? '#ffffff' : '#28251d');
  }

  function _esc(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

})();
