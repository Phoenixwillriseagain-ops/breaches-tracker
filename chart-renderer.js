// chart-renderer.js – Inline bar-charts for Breaches Tracker (no external lib)
(function () {
  'use strict';

  var PALETTE = [
    '#01696f','#437a22','#964219','#7a39bb',
    '#006494','#d19900','#a12c7b','#da7101'
  ];

  function topN(map, n) {
    return Object.entries(map)
      .sort(function (a, b) { return b[1] - a[1]; })
      .slice(0, n);
  }

  function groupBy(rows, col) {
    var map = {};
    rows.forEach(function (r) {
      var k = (r[col] || 'Unknown').trim() || 'Unknown';
      map[k] = (map[k] || 0) + 1;
    });
    return map;
  }

  function barChart(entries, color) {
    if (!entries.length)
      return '<p style="color:var(--muted);font-size:12px;padding:8px 0;">No data</p>';
    var max = entries[0][1];
    return '<div style="display:flex;flex-direction:column;gap:6px;">'
      + entries.map(function (e) {
          var pct = max ? Math.round(e[1] / max * 100) : 0;
          return '<div style="display:flex;align-items:center;gap:8px;font-size:12px;">'
            + '<div style="width:140px;min-width:80px;text-align:right;color:var(--muted);'
            +           'overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="' + e[0] + '">' + e[0] + '</div>'
            + '<div style="flex:1;height:12px;background:var(--divider);border-radius:4px;overflow:hidden;">'
            +   '<div style="width:' + pct + '%;height:100%;background:' + color + ';border-radius:4px;transition:width .35s;"></div>'
            + '</div>'
            + '<div style="width:32px;font-weight:600;text-align:right;">' + e[1] + '</div>'
            + '</div>';
        }).join('')
      + '</div>';
  }

  function card(title, body) {
    return '<div style="background:var(--surface);border:1px solid var(--border);'
      +              'border-radius:10px;overflow:hidden;box-shadow:0 1px 3px rgba(0,0,0,.05);">'
      + '<div style="padding:10px 16px;border-bottom:1px solid var(--border);'
      +              'font-size:13px;font-weight:600;">' + title + '</div>'
      + '<div style="padding:14px 16px;">' + body + '</div>'
      + '</div>';
  }

  function render() {
    var area = document.getElementById('charts-area');
    if (!area) return;

    // Flatten all tab rows
    var state = window.BT && window.BT.trackerState;
    if (!state || !state.tabs) { area.style.display = 'none'; return; }
    var allRows = [];
    Object.values(state.tabs).forEach(function (arr) {
      allRows = allRows.concat(arr);
    });
    if (!allRows.length) { area.style.display = 'none'; return; }

    var charts = [
      { title: 'By SLA Code',  col: 'SLA_Code',   n: 8,  color: PALETTE[0] },
      { title: 'By Queue',     col: 'Queue',       n: 8,  color: PALETTE[1] },
      { title: 'By Tool',      col: 'Tool',        n: 8,  color: PALETTE[2] },
      { title: 'By Priority',  col: 'Priority',    n: 8,  color: PALETTE[3] },
      { title: 'By Language',  col: 'ISO_Language', n: 8, color: PALETTE[4] },
      { title: 'By Topic',     col: 'TOPIC',       n: 8,  color: PALETTE[5] },
    ];

    var html = '<div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(360px,1fr));gap:16px;margin:16px 20px 20px;">'
      + charts.map(function (c) {
          var entries = topN(groupBy(allRows, c.col), c.n);
          return card(c.title, barChart(entries, c.color));
        }).join('')
      + '</div>';

    area.innerHTML = html;
    area.style.display = '';
  }

  // Hook into the global render lifecycle
  var _origRender = window.render;
  window.render = function () {
    if (typeof _origRender === 'function') _origRender.apply(this, arguments);
    render();
  };

  window.BT_CHARTS = { render: render };
})();
