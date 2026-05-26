// reporter-processor.js — V2 model processor for Breaches Reporter
(function() {
  'use strict';

  window.RPT = {
    allData: [],
    filtered: [],
    aosFiltered: [],
    uniqueValues: {},

    loadFile: function(file) { window.RPT.loadFiles([file]); },

    loadFiles: function(workbooks) {
      _startMultiImport(workbooks);
    },

    applyFilters: function() { _applyFilters(); },

    clearData: function() {
      window.RPT.allData     = [];
      window.RPT.filtered    = [];
      window.RPT.aosFiltered = [];
      window.RPT.uniqueValues= {};
      _hideImportLog();
      var lbl = document.getElementById('lang-trigger-label');
      if (lbl) lbl.textContent = 'All';
      var list = document.getElementById('lang-dd-list');
      if (list) list.innerHTML = '';
      var up = document.getElementById('upload-section');
      var dp = document.getElementById('data-section');
      var sb = document.getElementById('rpt-sidebar');
      var eb = document.getElementById('export-btns');
      if(up) up.style.display = '';
      if(dp) { dp.style.display = 'none'; }
      if(sb) sb.style.display  = 'none';
      if(eb) eb.style.display  = 'none';
      var rc = document.getElementById('record-count');
      if(rc) rc.textContent = '';
    },

    clearAll: function() { window.RPT.clearData(); },

    clearFilters: function() {
      ['filter-week','filter-sla','filter-excl','filter-aos','filter-tool','filter-queue','filter-sheet']
        .forEach(function(id){
          var el=document.getElementById(id);
          if(el) el.value='All';
        });
      document.querySelectorAll('#lang-dd-list input[type=checkbox]').forEach(function(cb){
        cb.checked = true;
      });
      if (typeof window._updateLangLabel === 'function') window._updateLangLabel();
      _applyFilters();
    },

    resetToUpload: function() { window.RPT.clearData(); },

    exportFiltered: function() {
      if(window.EXPORTER) window.EXPORTER.exportAsExcel(window.RPT.filtered);
      else alert('Exporter not loaded');
    },

    exportReport: function() {
      if(window.EXPORTER) window.EXPORTER.exportAsExcel(
        window.RPT.allData,
        'Breaches_FullReport_'+new Date().toISOString().slice(0,10).replace(/-/g,'')+'.xlsx'
      );
      else alert('Exporter not loaded');
    },

    countUniqueTickets: function(data) {
      var seen = new Set();
      (data || []).forEach(function(r) { if (r.ticket) seen.add(r.ticket); });
      return seen.size;
    },
  };

  /* ---- helpers ---- */
  function clean(v){ return String(v==null?'':v).trim(); }

  function formatDate(val){
    if(!val) return '';
    var d=new Date(val);
    if(isNaN(d.getTime())) return clean(val);
    var t=new Date(d.getTime()+3*3600*1000);
    function p(n){return String(n).padStart(2,'0');}
    return p(t.getUTCDate())+'.'+p(t.getUTCMonth()+1)+'.'+t.getUTCFullYear()+
           ' '+p(t.getUTCHours())+':'+p(t.getUTCMinutes())+':'+p(t.getUTCSeconds());
  }

  function normBool(v){
    var s=clean(v).toLowerCase();
    return(['t','true','yes','y','1'].indexOf(s)!==-1)?'Y':'N';
  }

  /*
   * Column resolver — builds a case-insensitive, whitespace-normalised lookup
   * from the actual keys present in a parsed row object.
   *
   * Usage:
   *   var col = _makeColResolver(rows[0]);
   *   col(row, 'Incident Ticket')  // finds key regardless of case/spacing
   *
   * Each canonical name maps to one or more accepted aliases tried in order.
   */
  var COL_ALIASES = {
    'Incident Ticket':    ['incident ticket', 'incidentticket', 'ticket', 'inc ticket', 'incident_ticket'],
    'DATE_CLOSE':         ['date_close', 'dateclose', 'date close', 'closed date', 'date closed'],
    'DATE_TIME_Breach':   ['date_time_breach', 'datetimebreach', 'datetime breach', 'breach datetime', 'breach_datetime', 'breach date'],
    'DATE_TIME_Breach UTC': ['date_time_breach utc', 'date_time_breach_utc', 'breach datetime utc'],
    'Status':             ['status'],
    'Queue':              ['queue'],
    'Priority':           ['priority'],
    'ISO_Language':       ['iso_language', 'isolanguage', 'language', 'lang', 'iso language'],
    'Tool':               ['tool'],
    'TOPIC':              ['topic'],
    'SLA_Code':           ['sla_code', 'slacode', 'sla code', 'sla'],
    'SLA_N':              ['sla_n', 'slan', 'sla n', 'sla number'],
    'Breach_Description': ['breach_description', 'breachdescription', 'breach description', 'breach desc'],
    'COMPASS ID':         ['compass id', 'compassid', 'compass_id'],
    'Reason':             ['reason'],
    'Action':             ['action'],
    'AOS':                ['aos'],
    'Agent':              ['agent'],
    'BMS ID':             ['bms id', 'bmsid', 'bms_id'],
    'Comment':            ['comment', 'comments'],
    'AOS Issue':          ['aos issue', 'aosissue', 'aos_issue'],
    'Excluded':           ['excluded', 'exclude'],
    'Jira':               ['jira', 'jira ticket', 'jira id', 'jira_ticket'],
    'Week':               ['week', 'wk', 'week number', 'week no'],
    'Unique':             ['unique'],
  };

  function _makeColResolver(sampleRow) {
    // Build a normalised-key → actual-key map from whatever the row has
    var keyMap = {};
    Object.keys(sampleRow).forEach(function(k) {
      keyMap[k.toLowerCase().replace(/\s+/g,' ').trim()] = k;
    });

    return function get(row, canonicalName) {
      // 1. Try exact match first (fastest, covers most cases)
      if (row[canonicalName] !== undefined) return row[canonicalName];

      // 2. Try each alias
      var aliases = COL_ALIASES[canonicalName] || [];
      for (var i = 0; i < aliases.length; i++) {
        var actualKey = keyMap[aliases[i]];
        if (actualKey !== undefined && row[actualKey] !== undefined) {
          return row[actualKey];
        }
      }

      // 3. Fallback: case-insensitive scan of the canonical name itself
      var normCanon = canonicalName.toLowerCase().replace(/\s+/g,' ').trim();
      var fallbackKey = keyMap[normCanon];
      if (fallbackKey !== undefined) return row[fallbackKey];

      return '';
    };
  }

  /* ---- Import log ---- */
  function _showImportLog(lines) {
    var el = document.getElementById('rpt-import-log');
    if (!el) return;
    el.innerHTML = lines.map(function(l) {
      var icon  = l.type==='ok' ? '\u2713' : l.type==='warn' ? '\u26a0' : '\u2715';
      var color = l.type==='ok' ? 'var(--success,#437a22)' :
                  l.type==='warn' ? 'var(--warning,#964219)' : 'var(--error,#a12c7b)';
      return '<span style="color:'+color+';margin-right:8px;">'+icon+'</span>'+l.msg;
    }).join('<br>');
    el.style.display = 'block';
  }
  function _hideImportLog() {
    var el = document.getElementById('rpt-import-log');
    if (el) el.style.display = 'none';
  }

  /* ---- Multi-workbook import ---- */
  function _startMultiImport(workbooks) {
    _hideImportLog();
    var logLines = [];

    if (!Array.isArray(workbooks) || !workbooks.length) {
      logLines.push({ type: 'error', msg: 'No workbooks provided.' });
      _showImportLog(logLines);
      return;
    }

    workbooks.forEach(function(wb) {
      if (!wb || !wb.SheetNames) {
        logLines.push({ type: 'error', msg: 'Invalid workbook object received.' });
        return;
      }
      try {
        var result = _mergeWorkbook(wb, logLines);
        var name = wb._filename || ('workbook ' + (logLines.length + 1));
        logLines.push({
          type: result.added > 0 ? 'ok' : 'warn',
          msg: name + ' \u2014 ' + result.added + ' row(s) imported (' + result.uniqueTickets + ' unique ticket(s))'
             + (result.colWarning ? ' \u26a0 ' + result.colWarning : '')
        });
      } catch(err) {
        logLines.push({ type: 'error', msg: (wb._filename || 'workbook') + ' \u2014 ' + err.message });
      }
    });

    _populateUnique();
    _updateDropdowns();
    _buildLangDropdown();
    _applyFiltersOnly();
    window.renderTables();
    _showImportLog(logLines);
    console.log('[RPT] total allData:', window.RPT.allData.length, '| filtered:', window.RPT.filtered.length);
  }

  /* ---- Merge one workbook into allData ---- */
  function _mergeWorkbook(wb) {
    var added = 0;
    var ticketsInFile = new Set();
    var colWarning = '';

    wb.SheetNames.forEach(function(sn) {
      if (sn === 'Instructions') return;
      var rows = XLSX.utils.sheet_to_json(wb.Sheets[sn], { defval:'' });
      if (!rows.length) return;

      // Build column resolver once per sheet using the first row as sample
      var col = _makeColResolver(rows[0]);

      // Warn once if the ticket column couldn't be found by exact match
      var sampleTicket = col(rows[0], 'Incident Ticket');
      if (sampleTicket === '' && rows.length > 1) {
        // Check if resolved via alias by logging actual keys
        var actualKeys = Object.keys(rows[0]).join(', ');
        colWarning = 'Header "Incident Ticket" not found exactly — resolved via alias. Actual headers: ' + actualKeys.slice(0, 120);
      }

      rows.forEach(function(r) {
        var ticket = clean(col(r, 'Incident Ticket'));
        if (!ticket) return;
        ticketsInFile.add(ticket);

        var aosF     = normBool(col(r, 'AOS'));
        var aosIRaw  = col(r, 'AOS Issue');
        var aosI     = normBool(aosIRaw);
        var breachRaw = col(r, 'DATE_TIME_Breach') || col(r, 'DATE_TIME_Breach UTC') || '';

        window.RPT.allData.push({
          ticket:         ticket,
          dateClosed:     formatDate(col(r, 'DATE_CLOSE')),
          dateTimeBreach: formatDate(breachRaw),
          status:         clean(col(r, 'Status'))   || 'N/A',
          queue:          clean(col(r, 'Queue')),
          priority:       clean(col(r, 'Priority')),
          language:       clean(col(r, 'ISO_Language')) || 'Unknown',
          tool:           clean(col(r, 'Tool'))         || 'Unknown',
          topic:          clean(col(r, 'TOPIC')),
          sla:            clean(col(r, 'SLA_Code'))     || 'Unknown',
          slaN:           clean(col(r, 'SLA_N')),
          breachDesc:     clean(col(r, 'Breach_Description')),
          compassId:      clean(col(r, 'COMPASS ID')),
          reason:         clean(col(r, 'Reason')),
          action:         clean(col(r, 'Action')),
          aos:            aosF,
          agent:          clean(col(r, 'Agent')),
          bmsId:          clean(col(r, 'BMS ID')),
          comment:        clean(col(r, 'Comment')),
          aosIssue:       clean(aosIRaw),
          excluded:       normBool(col(r, 'Excluded')),
          jira:           clean(col(r, 'Jira')),
          week:           clean(col(r, 'Week')),
          unique:         clean(col(r, 'Unique')),
          sheet:          sn,
          isAos:          aosF==='Y' || aosI==='Y',
        });
        added++;
      });
    });

    window.RPT.aosFiltered = window.RPT.allData.filter(function(r){return r.isAos;});
    return { added: added, uniqueTickets: ticketsInFile.size, colWarning: colWarning };
  }

  function _populateUnique(){
    var uv={weeks:[],slas:[],languages:[],tools:[],queues:[],sheets:[]};
    window.RPT.allData.forEach(function(r){
      if(r.week     && uv.weeks.indexOf(r.week)<0)         uv.weeks.push(r.week);
      if(r.sla      && uv.slas.indexOf(r.sla)<0)           uv.slas.push(r.sla);
      if(r.language && uv.languages.indexOf(r.language)<0) uv.languages.push(r.language);
      if(r.tool     && uv.tools.indexOf(r.tool)<0)         uv.tools.push(r.tool);
      if(r.queue    && uv.queues.indexOf(r.queue)<0)       uv.queues.push(r.queue);
      if(r.sheet    && uv.sheets.indexOf(r.sheet)<0)       uv.sheets.push(r.sheet);
    });
    uv.weeks.sort(function(a,b){return Number(a)-Number(b);});
    uv.slas.sort();uv.languages.sort();uv.tools.sort();uv.queues.sort();
    window.RPT.uniqueValues=uv;
  }

  /* ---- Build language checkbox list ---- */
  function _buildLangDropdown() {
    var list = document.getElementById('lang-dd-list');
    if (!list) return;
    var langs = (window.RPT.uniqueValues.languages || []).slice().sort();
    list.innerHTML = langs.map(function(lang) {
      var id = 'lang-cb-' + lang.replace(/[^a-zA-Z0-9]/g,'_');
      return '<label class="lang-dd-item" for="'+id+'">'+
        '<input type="checkbox" id="'+id+'" value="'+_escAttr(lang)+'" checked '+
        'onchange="window._onLangChange()">'+
        '<span>'+_esc(lang)+'</span>'+
        '</label>';
    }).join('');
    if (typeof window._updateLangLabel === 'function') window._updateLangLabel();
  }

  function _updateDropdowns(){
    var uv=window.RPT.uniqueValues;
    function fill(id,arr){
      var el=document.getElementById(id);
      if(!el)return;
      el.innerHTML='<option value="All">All</option>'+
        arr.map(function(v){return'<option value="'+v+'">'+v+'</option>';}).join('');
    }
    fill('filter-week', uv.weeks);
    fill('filter-sla',  uv.slas);
    fill('filter-tool', uv.tools);
    fill('filter-queue',uv.queues);
    fill('filter-sheet',uv.sheets);
    var fe=document.getElementById('filter-excl');
    if(fe)fe.innerHTML='<option value="All">All</option><option value="Y">Excluded</option><option value="N">Counted</option>';
    var fa=document.getElementById('filter-aos');
    if(fa)fa.innerHTML='<option value="All">All</option><option value="Y">AOS Only</option><option value="N">Non-AOS</option>';

    var up=document.getElementById('upload-section');
    var dp=document.getElementById('data-section');
    if(up) up.style.display='none';
    if(dp) dp.style.display='flex';
  }

  function _applyFiltersOnly(){
    function g(id){var e=document.getElementById(id);return e?e.value:'All';}
    var wF=g('filter-week'), sF=g('filter-sla'),
        eF=g('filter-excl'), aF=g('filter-aos'),  tF=g('filter-tool'),
        qF=g('filter-queue'),shF=g('filter-sheet');

    var selectedLangs = (typeof window.getSelectedLanguages === 'function')
      ? window.getSelectedLanguages()
      : null;
    var totalLangs = document.querySelectorAll('#lang-dd-list input[type=checkbox]').length;
    var filterLang = selectedLangs && selectedLangs.size < totalLangs;

    window.RPT.filtered=window.RPT.allData.filter(function(r){
      if(wF !=='All'&&r.week     !==wF)  return false;
      if(sF !=='All'&&r.sla      !==sF)  return false;
      if(filterLang && !selectedLangs.has(r.language)) return false;
      if(eF !=='All'&&r.excluded !==eF)  return false;
      if(aF !=='All'&&(r.isAos?'Y':'N')!==aF) return false;
      if(tF !=='All'&&r.tool     !==tF)  return false;
      if(qF !=='All'&&r.queue    !==qF)  return false;
      if(shF!=='All'&&r.sheet    !==shF) return false;
      return true;
    });
  }

  function _applyFilters(){
    _applyFiltersOnly();
    if(typeof window.renderTables==='function') window.renderTables();
  }

  function _esc(s){return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');}
  function _escAttr(s){return String(s).replace(/"/g,'&quot;').replace(/'/g,'&#39;');}

  document.addEventListener('DOMContentLoaded',function(){
    ['filter-week','filter-sla','filter-excl',
     'filter-aos','filter-tool','filter-queue','filter-sheet']
      .forEach(function(id){
        var el=document.getElementById(id);
        if(el) el.addEventListener('change',_applyFilters);
      });
  });

})();
