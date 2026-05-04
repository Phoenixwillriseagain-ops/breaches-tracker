// reporter-charts.js - Chart.js rendering for Reporter page with lazy loading
(function() {
  'use strict';

  var rpt = window.RPT;
  var charts = { _rendered: {} };
  window.RPT_CHARTS = charts;

  // Get theme-aware colors
  function getChartColors() {
    var cs = window.getComputedStyle(document.documentElement);
    return {
      primary: cs.getPropertyValue('--primary').trim() || '#01696f',
      success: cs.getPropertyValue('--success').trim() || '#437a22',
      warn: cs.getPropertyValue('--warn').trim() || '#964219',
      error: cs.getPropertyValue('--error').trim() || '#a12c7b',
      muted: cs.getPropertyValue('--muted').trim() || '#7a7974',
      light: cs.getPropertyValue('--faint').trim() || '#bab9b4',
    };
  }

  // Label sanitizer: replaces null/undefined/empty with fallback
  function safeLabel(v, fallback) {
    var s = String(v == null ? '' : v).trim();
    return s || fallback;
  }

  // Create/Update chart with lazy rendering
  function mkChart(canvasId, type, labels, datasets, title) {
    var canvas = document.getElementById(canvasId);
    if (!canvas) return;
    var container = canvas.closest('[id^="tab-"]');
    if (container && container.style.display === 'none') {
      charts._rendered[canvasId] = { canvasId: canvasId, type: type, labels: labels, datasets: datasets, title: title };
      return;
    }
    charts._rendered[canvasId] = null;
    var ctx = canvas.getContext('2d');
    var colors = getChartColors();
    if (charts[canvasId]) { charts[canvasId].destroy(); }
    var data = { labels: labels, datasets: datasets };
    var opts = {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: true, position: 'top' },
        title: { display: false },
        tooltip: { enabled: true },
      },
      layout: { padding: { top: 10, bottom: 10 } },
    };
    charts[canvasId] = new Chart(ctx, { type: type, data: data, options: opts });
  }

  // Aggregate helpers
  function groupBy(data, key) {
    var map = {};
    data.forEach(function(d) {
      var k = safeLabel(d[key], 'Unknown');
      map[k] = (map[k] || 0) + 1;
    });
    return map;
  }

  function topN(map, n) {
    return Object.keys(map).sort(function(a, b) { return map[b] - map[a]; }).slice(0, n);
  }

  function colorsArray(n, colors) {
    var pal = [colors.primary, colors.success, colors.warn, colors.error, colors.muted, colors.light];
    var res = [];
    for (var i = 0; i < n; i++) res.push(pal[i % pal.length]);
    return res;
  }

  // Render Overview charts
  function renderOverview() {
    if (!rpt || !rpt.filtered || !rpt.filtered.length) return;
    var d = rpt.filtered;
    var colors = getChartColors();
    var weekMap = groupBy(d, 'week');
    var wkKeys = Object.keys(weekMap).sort();
    mkChart('chart-week', 'bar', wkKeys, [{ data: wkKeys.map(function(k) { return weekMap[k]; }), backgroundColor: colors.primary }]);
    var slaMap = groupBy(d, 'sla_code');
    var slaKeys = topN(slaMap, 10);
    mkChart('chart-sla', 'doughnut', slaKeys, [{ data: slaKeys.map(function(k) { return slaMap[k]; }), backgroundColor: colorsArray(slaKeys.length, colors) }]);
    var langMap = groupBy(d, 'language');
    var langKeys = Object.keys(langMap);
    mkChart('chart-lang', 'doughnut', langKeys, [{ data: langKeys.map(function(k) { return langMap[k]; }), backgroundColor: colorsArray(langKeys.length, colors) }]);
    var excl = d.filter(function(r) { return r.excluded === '1'; }).length;
    var counted = d.length - excl;
    mkChart('chart-excl', 'pie', ['Counted', 'Excluded'], [{ data: [counted, excl], backgroundColor: [colors.success, colors.error] }]);
  }

  // Render Reasons tab
  function renderReasons() {
    if (!rpt || !rpt.filtered || !rpt.filtered.length) return;
    var d = rpt.filtered;
    var colors = getChartColors();
    var reasonMap = groupBy(d, 'reason');
    var topKeys = topN(reasonMap, 10);
    mkChart('chart-reasons', 'bar', topKeys, [{ data: topKeys.map(function(k) { return reasonMap[k]; }), backgroundColor: colors.success }]);
    var rsla = {};
    d.forEach(function(r) {
      var k = safeLabel(r.reason, 'Unknown') + '|' + safeLabel(r.sla_code, 'Unknown');
      rsla[k] = (rsla[k] || 0) + 1;
    });
    var rslaKeys = topN(rsla, 10);
    mkChart('chart-reasons-sla', 'doughnut', rslaKeys.map(function(k) { return k.replace('|', ' / '); }), [{ data: rslaKeys.map(function(k) { return rsla[k]; }), backgroundColor: colorsArray(rslaKeys.length, colors) }]);
    var tableData = topKeys.map(function(k) { return { Reason: k, Count: reasonMap[k] }; });
    if (rpt.buildTable) rpt.buildTable('reasons-table', tableData, ['Reason', 'Count']);
  }

  // Render Apps tab
  function renderApps() {
    if (!rpt || !rpt.filtered || !rpt.filtered.length) return;
    var d = rpt.filtered;
    var colors = getChartColors();
    var appMap = groupBy(d, 'application');
    var appKeys = topN(appMap, 10);
    mkChart('chart-apps', 'bar', appKeys, [{ data: appKeys.map(function(k) { return appMap[k]; }), backgroundColor: colors.primary }]);
    var asla = {};
    d.forEach(function(r) {
      var k = safeLabel(r.application, 'Unknown') + '|' + safeLabel(r.sla_code, 'Unknown');
      asla[k] = (asla[k] || 0) + 1;
    });
    var aslaKeys = topN(asla, 10);
    mkChart('chart-apps-sla', 'doughnut', aslaKeys.map(function(k) { return k.replace('|', ' / '); }), [{ data: aslaKeys.map(function(k) { return asla[k]; }), backgroundColor: colorsArray(aslaKeys.length, colors) }]);
    var tableData = appKeys.map(function(k) { return { Application: k, Count: appMap[k] }; });
    if (rpt.buildTable) rpt.buildTable('apps-table', tableData, ['Application', 'Count']);
  }

  // Render AOS tab
  function renderAOS() {
    if (!rpt || !rpt.filtered || !rpt.filtered.length) return;
    var d = rpt.filtered;
    var colors = getChartColors();
    // Use normalized aos === '1' (processor now outputs exactly '1' or '0')
    var aosData = d.filter(function(r) { return r.aos === '1'; });
    var total = d.length;
    var aosCount = aosData.length;
    var kp = document.getElementById('aos-kpis');
    if (kp) {
      function uCount(data, key) {
        var s = {};
        data.forEach(function(d) { if (d[key]) s[d[key]] = true; });
        return Object.keys(s).length;
      }
      kp.innerHTML = [
        { l: 'Total', v: total, s: 'records' },
        { l: 'AOS Issues', v: aosCount, s: 'found' },
        { l: 'AOS %', v: total ? Math.round(aosCount / total * 100) : 0, s: 'of all' },
        { l: 'Agents', v: uCount(aosData, 'agent'), s: 'unique' },
      ].map(function(k) {
        return '<div class="kpi-card"><div class="kpi-value">' + k.v + '</div><div class="kpi-label">' + k.l + '</div><div class="kpi-sub">' + k.s + '</div></div>';
      }).join('');
    }
    var weekMap = groupBy(aosData, 'week');
    var wkKeys = Object.keys(weekMap).sort();
    mkChart('chart-aos-week', 'bar', wkKeys, [{ data: wkKeys.map(function(k) { return weekMap[k]; }), backgroundColor: colors.error }]);
    var slaMap = groupBy(aosData, 'sla_code');
    var slaKeys = Object.keys(slaMap);
    mkChart('chart-aos-sla', 'doughnut', slaKeys, [{ data: slaKeys.map(function(k) { return slaMap[k]; }), backgroundColor: colorsArray(slaKeys.length, colors) }]);
    var tableData = aosData.map(function(r) {
      return { ticket: r.ticket, sla_code: r.sla_code, week: r.week, reason: r.reason, aos: r.aos };
    });
    if (rpt.buildTable) rpt.buildTable('aos-table', tableData, ['ticket', 'sla_code', 'week', 'reason', 'aos']);
  }

  // Render KM-1 tab
  function renderKM1() {
    if (!rpt || !rpt.filtered || !rpt.filtered.length) return;
    // Use normalized sla_code === 'KM-1'
    var d = rpt.filtered.filter(function(r) { return r.sla_code === 'KM-1'; });
    if (!d.length) return;
    var colors = getChartColors();
    var rcMap = {};
    d.forEach(function(r) {
      rcMap['Remote'] = (rcMap['Remote'] || 0) + (r.remote && r.remote !== '0' ? 1 : 0);
      rcMap['Callback'] = (rcMap['Callback'] || 0) + (r.callback && r.callback !== '0' ? 1 : 0);
    });
    mkChart('chart-km1-rc', 'bar', Object.keys(rcMap), [{ data: [rcMap['Remote'] || 0, rcMap['Callback'] || 0], backgroundColor: [colors.primary, colors.warn] }]);
    var reasonMap = groupBy(d, 'reason');
    var topKeys = topN(reasonMap, 10);
    mkChart('chart-km1-reasons', 'bar', topKeys, [{ data: topKeys.map(function(k) { return reasonMap[k]; }), backgroundColor: colors.success }]);
    var tableData = d.map(function(r) {
      return { ticket: r.ticket, reason: r.reason, remote: r.remote, callback: r.callback, week: r.week, language: r.language };
    });
    if (rpt.buildTable) rpt.buildTable('km1-table', tableData, ['ticket', 'reason', 'remote', 'callback', 'week', 'language']);
  }

  // Re-render lazy charts when tab becomes visible
  function renderPending() {
    var rendered = charts._rendered;
    Object.keys(rendered).forEach(function(id) {
      var r = rendered[id];
      if (r) { mkChart(r.canvasId, r.type, r.labels, r.datasets); }
    });
  }

  // Exports
  charts.renderOverview = renderOverview;
  charts.renderReasons = renderReasons;
  charts.renderApps = renderApps;
  charts.renderAOS = renderAOS;
  charts.renderKM1 = renderKM1;
  charts.renderPending = renderPending;
})();
