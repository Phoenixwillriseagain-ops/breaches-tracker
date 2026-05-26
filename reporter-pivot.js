// reporter-pivot.js — Semi-dynamic pivot table + per-SLA panels
(function () {
  'use strict';

  var PIVOT_FIELDS = [
    { key: 'sla',      label: 'SLA' },
    { key: 'week',     label: 'Week' },
    { key: 'queue',    label: 'Queue' },
    { key: 'tool',     label: 'Tool' },
    { key: 'language', label: 'Language' },
    { key: 'sheet',    label: 'Sheet' },
    { key: 'excluded', label: 'Excluded' },
  ];

  var SLA_PANELS = [
    { sla: 'KSL-4',  rgb: [0,100,148],  combos: [['week','queue'],['week','tool'],['week','language']] },
    { sla: 'KM-1',   rgb: [14,78,120],  combos: [['week','queue'],['week','tool'],['week','language']] },
    { sla: 'KSL-5a', rgb: [150,66,25],  combos: [['week','queue'],['week','tool'],['week','language']] },
    { sla: 'KM-2',   rgb: [122,57,187], combos: [['week','queue'],['week','tool'],['week','language']] },
  ];

  var _rowField = 'week';
  var _colField = 'sla';

  window.renderPivot = function () {
    var container = document.getElementById('pivot-container');
    if (!container) return;
    var data = window.RPT ? window.RPT.filtered : [];
    container.innerHTML =
      _buildPivotHTML(data) +
      _buildSlaPanels(data);
    _bindPivotEvents();
  };

  /* ================================================================
     MAIN PIVOT
     ================================================================ */
  function _buildPivotHTML(data) {
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

    if (_rowField === _colField) {
      return '<div class="sc pvt-wrap">' +
        '<div class="sc-head"><h2>Pivot Table</h2></div>' +
        '<div class="sc-body">' + controls +
        '<p class="pvt-msg pvt-warn">&#9888;&#xFE0E; Row and Column fields must be different.</p>' +
        '</div></div>';
    }

    var rowVals = _uniqueVals(data, _rowField);
    var colVals = _uniqueVals(data, _colField);

    if (!data.length || !rowVals.length || !colVals.length) {
      return '<div class="sc pvt-wrap">' +
        '<div class="sc-head"><h2>Pivot Table</h2></div>' +
        '<div class="sc-body">' + controls +
        '<p class="pvt-msg">No data available for the current filter.</p>' +
        '</div></div>';
    }

    var matrix = _buildMatrix(data, rowVals, colVals, _rowField, _colField);
    var rowTotals = _rowTotals(matrix, rowVals, colVals);
    var colTotals = _colTotals(matrix, rowVals, colVals);
    var grandTotal = _countUniq(data);
    var maxCell = _maxCell(matrix, rowVals, colVals);

    var colFieldLabel = _fieldLabel(_colField);
    var rowFieldLabel = _fieldLabel(_rowField);

    var thead = '<thead><tr>' +
      '<th class="pvt-th pvt-th-corner">' + _esc(rowFieldLabel) + ' \ ' + _esc(colFieldLabel) + '</th>' +
      colVals.map(function (cv) { return '<th class="pvt-th pvt-th-col">' + _esc(cv) + '</th>'; }).join('') +
      '<th class="pvt-th pvt-th-total">Total</th>' +
      '</tr></thead>';

    var tbody = '<tbody>' + rowVals.map(function (rv) {
      var cells = colVals.map(function (cv) {
        var v = _countUniq(matrix[rv][cv]);
        return '<td class="pvt-cell" style="' + _heatStyle(v, maxCell, null) + '" title="' +
          _esc(rowFieldLabel) + ': ' + _esc(rv) + ' | ' + _esc(colFieldLabel) + ': ' + _esc(cv) + ' — ' + v + ' unique ticket(s)">' +
          (v > 0 ? v : '') + '</td>';
      }).join('');
      return '<tr>' +
        '<td class="pvt-row-lbl">' + _esc(rv) + '</td>' +
        cells +
        '<td class="pvt-total-cell pvt-row-total">' + rowTotals[rv] + '</td>' +
      '</tr>';
    }).join('') +
    '<tr class="pvt-grand-row">' +
      '<td class="pvt-grand-lbl">Grand Total</td>' +
      colVals.map(function (cv) {
        return '<td class="pvt-total-cell pvt-col-total">' + colTotals[cv] + '</td>';
      }).join('') +
      '<td class="pvt-total-cell pvt-grand-total">' + grandTotal + '</td>' +
    '</tr></tbody>';

    var badge = rowVals.length + ' row' + (rowVals.length !== 1 ? 's' : '') +
      ' &bull; ' + colVals.length + ' col' + (colVals.length !== 1 ? 's' : '') +
      ' &bull; ' + grandTotal + ' unique ticket(s)';

    return '<div class="sc pvt-wrap">' +
      '<div class="sc-head"><h2>Pivot Table</h2><span class="badge">' + badge + '</span></div>' +
      '<div class="sc-body">' +
        controls +
        '<div class="pvt-table-scroll"><table class="pvt-table">' + thead + tbody + '</table></div>' +
        '<p class="pvt-note">&#9432; Values = unique Incident Tickets per cell. Cells with 0 are left blank. Heat shading relative to highest cell value.</p>' +
      '</div>' +
    '</div>';
  }

  /* ================================================================
     PER-SLA PANELS
     Each SLA gets a card with three mini fixed pivots:
     Week x Queue | Week x Tool | Week x Language
     ================================================================ */
  function _buildSlaPanels(data) {
    var activeData = data.filter(function (r) { return r.excluded === 'N'; });

    var panelCards = SLA_PANELS.map(function (cfg) {
      var slaRows = activeData.filter(function (r) { return r.sla === cfg.sla; });
      var total = _countUniq(slaRows);

      if (!slaRows.length) {
        return '<div class="sc sla-pvt-panel">' +
          '<div class="sc-head"><h2>' + _esc(cfg.sla) + ' Breakdown</h2>' +
          '<span class="badge">no active data</span></div>' +
          '<div class="sc-body"><p class="pvt-msg">No active records for ' + _esc(cfg.sla) + ' in the current filter.</p></div>' +
          '</div>';
      }

      var accent = 'rgb(' + cfg.rgb[0] + ',' + cfg.rgb[1] + ',' + cfg.rgb[2] + ')';

      var miniPivots = cfg.combos.map(function (combo) {
        return _buildMiniPivot(slaRows, combo[0], combo[1], cfg.rgb);
      }).join('');

      return '<div class="sc sla-pvt-panel">' +
        '<div class="sc-head">' +
          '<h2 style="color:' + accent + '">' + _esc(cfg.sla) + ' — Drill-down</h2>' +
          '<span class="badge">' + total + ' unique active ticket(s)</span>' +
        '</div>' +
        '<div class="sc-body">' +
          '<div class="sla-pvt-grid">' + miniPivots + '</div>' +
        '</div>' +
      '</div>';
    }).join('');

    return '<div class="sla-pvt-section" style="display:flex;flex-direction:column;gap:16px;margin-top:16px">' +
      '<div class="sla-pvt-section-hdr" style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:var(--muted)">Per-SLA Drill-down</div>' +
      panelCards +
    '</div>';
  }

  function _buildMiniPivot(data, rowKey, colKey, rgb) {
    var rowVals = _uniqueVals(data, rowKey);
    var colVals = _uniqueVals(data, colKey);
    var rowLabel = _fieldLabel(rowKey);
    var colLabel = _fieldLabel(colKey);

    if (!rowVals.length || !colVals.length) {
      return '<div class="mini-pvt-wrap"><p class="pvt-msg" style="font-size:11px">No data for ' + _esc(rowLabel) + ' x ' + _esc(colLabel) + '</p></div>';
    }

    var matrix = _buildMatrix(data, rowVals, colVals, rowKey, colKey);
    var maxC   = _maxCell(matrix, rowVals, colVals);

    var thead = '<thead><tr>' +
      '<th class="pvt-th pvt-th-corner" style="font-size:10px">' + _esc(rowLabel) + ' \ ' + _esc(colLabel) + '</th>' +
      colVals.map(function (cv) {
        var disp = String(cv).length > 12 ? String(cv).slice(0, 11) + '…' : cv;
        return '<th class="pvt-th pvt-th-col" style="font-size:10px" title="' + _esc(cv) + '">' + _esc(disp) + '</th>';
      }).join('') +
      '<th class="pvt-th pvt-th-total" style="font-size:10px">Total</th>' +
    '</tr></thead>';

    var rowTotals = _rowTotals(matrix, rowVals, colVals);
    var colTotals = _colTotals(matrix, rowVals, colVals);
    var grandTotal = _countUniq(data);

    var tbody = '<tbody>' + rowVals.map(function (rv) {
      var cells = colVals.map(function (cv) {
        var v = _countUniq(matrix[rv][cv]);
        return '<td class="pvt-cell" style="' + _heatStyle(v, maxC, rgb) + ';padding:3px 6px;height:24px;font-size:11px" title="' +
          _esc(rowLabel) + ': ' + _esc(rv) + ' | ' + _esc(colLabel) + ': ' + _esc(cv) + ' — ' + v + '">' +
          (v > 0 ? v : '') + '</td>';
      }).join('');
      return '<tr>' +
        '<td class="pvt-row-lbl" style="font-size:11px;padding:3px 8px">' + _esc(rv) + '</td>' +
        cells +
        '<td class="pvt-total-cell pvt-row-total" style="font-size:11px;padding:3px 8px">' + rowTotals[rv] + '</td>' +
      '</tr>';
    }).join('') +
    '<tr class="pvt-grand-row">' +
      '<td class="pvt-grand-lbl" style="font-size:10px;padding:4px 8px">Total</td>' +
      colVals.map(function (cv) {
        return '<td class="pvt-total-cell pvt-col-total" style="font-size:11px;padding:3px 8px">' + colTotals[cv] + '</td>';
      }).join('') +
      '<td class="pvt-total-cell pvt-grand-total" style="font-size:11px;padding:3px 8px">' + grandTotal + '</td>' +
    '</tr></tbody>';

    return '<div class="mini-pvt-wrap">' +
      '<div class="mini-pvt-title">' + _esc(rowLabel) + ' &times; ' + _esc(colLabel) + '</div>' +
      '<div class="pvt-table-scroll"><table class="pvt-table">' + thead + tbody + '</table></div>' +
    '</div>';
  }

  /* ================================================================
     EVENTS
     ================================================================ */
  function _bindPivotEvents() {
    var rowSel = document.getElementById('pvt-row-sel');
    var colSel = document.getElementById('pvt-col-sel');
    var expBtn = document.getElementById('pvt-export-btn');
    if (rowSel) rowSel.addEventListener('change', function () { _rowField = this.value; window.renderPivot(); });
    if (colSel) colSel.addEventListener('change', function () { _colField = this.value; window.renderPivot(); });
    if (expBtn) expBtn.addEventListener('click', _exportPivot);
  }

  /* ================================================================
     EXPORT
     ================================================================ */
  function _exportPivot() {
    if (typeof XLSX === 'undefined') { alert('SheetJS not loaded.'); return; }
    var data = window.RPT ? window.RPT.filtered : [];
    if (!data.length) { alert('No data to export.'); return; }
    if (_rowField === _colField) { alert('Row and Column fields must be different.'); return; }

    var rowVals = _uniqueVals(data, _rowField);
    var colVals = _uniqueVals(data, _colField);
    var matrix  = _buildMatrix(data, rowVals, colVals, _rowField, _colField);
    var rowFieldLabel = _fieldLabel(_rowField);
    var colFieldLabel = _fieldLabel(_colField);

    var aoa = [[rowFieldLabel + ' \\ ' + colFieldLabel].concat(colVals).concat(['Total'])];
    rowVals.forEach(function (rv) {
      var rowTotal = 0;
      var cells = colVals.map(function (cv) {
        var v = _countUniq(matrix[rv][cv]); rowTotal += v; return v;
      });
      aoa.push([rv].concat(cells).concat([rowTotal]));
    });
    aoa.push(['Grand Total'].concat(colVals.map(function (cv) {
      var all = []; rowVals.forEach(function (rv) { all = all.concat(matrix[rv][cv]); });
      return _countUniq(all);
    })).concat([_countUniq(data)]));

    var ws = XLSX.utils.aoa_to_sheet(aoa);
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Pivot');
    XLSX.writeFile(wb, 'Pivot_' + rowFieldLabel + '_x_' + colFieldLabel + '_' +
      new Date().toISOString().slice(0, 10).replace(/-/g, '') + '.xlsx');
  }

  /* ================================================================
     HELPERS
     ================================================================ */
  function _buildMatrix(data, rowVals, colVals, rowKey, colKey) {
    var m = {};
    rowVals.forEach(function (rv) {
      m[rv] = {};
      colVals.forEach(function (cv) { m[rv][cv] = []; });
    });
    data.forEach(function (r) {
      var rv = r[rowKey] || '(blank)';
      var cv = r[colKey] || '(blank)';
      if (m[rv] && m[rv][cv]) m[rv][cv].push(r);
    });
    return m;
  }

  function _rowTotals(matrix, rowVals, colVals) {
    var t = {};
    rowVals.forEach(function (rv) {
      var all = []; colVals.forEach(function (cv) { all = all.concat(matrix[rv][cv]); });
      t[rv] = _countUniq(all);
    });
    return t;
  }

  function _colTotals(matrix, rowVals, colVals) {
    var t = {};
    colVals.forEach(function (cv) {
      var all = []; rowVals.forEach(function (rv) { all = all.concat(matrix[rv][cv]); });
      t[cv] = _countUniq(all);
    });
    return t;
  }

  function _maxCell(matrix, rowVals, colVals) {
    var m = 0;
    rowVals.forEach(function (rv) {
      colVals.forEach(function (cv) {
        var v = _countUniq(matrix[rv][cv]); if (v > m) m = v;
      });
    });
    return m || 1;
  }

  function _uniqueVals(data, field) {
    var seen = {}, vals = [];
    data.forEach(function (r) {
      var v = r[field] || '(blank)';
      if (!seen[v]) { seen[v] = 1; vals.push(v); }
    });
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

  function _heatStyle(val, maxVal, rgb) {
    if (val === 0) return 'background:var(--pvt-zero,#f0eeea);color:var(--faint,#bab9b4)';
    var ratio = Math.pow(val / maxVal, 0.6);
    var BR = 247, BG = 246, BB = 242;
    var TR = rgb ? rgb[0] : 1;
    var TG = rgb ? rgb[1] : 105;
    var TB = rgb ? rgb[2] : 111;
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
