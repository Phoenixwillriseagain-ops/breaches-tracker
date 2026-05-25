// reporter-processor.js — V2 model processor for Breaches Reporter
(function() {
  'use strict';

  window.RPT = {
    allData: [],
    filtered: [],
    aosFiltered: [],
    uniqueValues: {},

    // Legacy single-file entry point — now routes through multi-import
    loadFile: function(file) { window.RPT.loadFiles([file]); },

    // Multi-file entry point
    loadFiles: function(files) { _startMultiImport(files); },

    // Public method called by language dropdown checkboxes
    applyFilters: function() { _applyFilters(); },

    clearData: function() {
      window.RPT.allData     = [];
      window.RPT.filtered    = [];
      window.RPT.aosFiltered = [];
      window.RPT.uniqueValues= {};
      _hideImportLog();
      // Reset lang dropdown label
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

    clearFilters: function() {
      ['filter-week','filter-sla','filter-excl','filter-aos','filter-tool','filter-queue','filter-sheet']
        .forEach(function(id){
          var el=document.getElementById(id);
          if(el) el.value='All';
        });
      // Reset language: check all
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

    // Returns the count of unique Incident Tickets in a given dataset array.
    // V2 files always have multiple rows per ticket (one row per breach event).
    // This is the only correct way to count tickets — never use data.length.
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
    var t=new Date(d.getTime()+3*3600*1000); // UTC+3 (EEST)
    function p(n){return String(n).padStart(2,'0');}
    return p(t.getUTCDate())+'.'+p(t.getUTCMonth()+1)+'.'+t.getUTCFullYear()+
           ' '+p(t.getUTCHours())+':'+p(t.getUTCMinutes())+':'+p(t.getUTCSeconds());
  }

  function normBool(v){
    var s=clean(v).toLowerCase();
    return(['t','true','yes','y','1'].indexOf(s)!==-1)?'Y':'N';
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

  /* ---- Multi-file sequential import ---- */
  function _startMultiImport(files) {
    _hideImportLog();
    var logLines = [];
    var idx = 0;

    function next() {
      if (idx >= files.length) {
        _populateUnique();
        _updateDropdowns();
        _buildLangDropdown();
        _applyFiltersOnly();
        window.renderTables();
        _showImportLog(logLines);
        console.log('[RPT] total allData:', window.RPT.allData.length);
        return;
      }
      var file = files[idx++];
      if (!/\.xlsx$/i.test(file.name)) {
        logLines.push({ type:'error', msg: file.name + ' \u2014 only .xlsx files are accepted' });
        next();
        return;
      }
      var reader = new FileReader();
      reader.onload = function(e) {
        try {
          var wb = XLSX.read(e.target.result, { type:'array', cellDates:true });
          var result = _mergeWorkbook(wb);
          logLines.push({ type: 'ok', msg: file.name + ' \u2014 ' + result.added + ' row(s) imported (' + result.uniqueTickets + ' unique ticket(s))' });
        } catch(err) {
          logLines.push({ type:'error', msg: file.name + ' \u2014 ' + err.message });
        }
        next();
      };
      reader.onerror = function() {
        logLines.push({ type:'error', msg: file.name + ' \u2014 read error' });
        next();
      };
      reader.readAsArrayBuffer(file);
    }

    next();
  }

  /* ---- Merge one workbook into allData ---- */
  // NOTE: We intentionally do NOT deduplicate rows by ticket number here.
  // V2 data has multiple rows per ticket — one row per breach event.
  // Row-level dedup would silently discard valid breach timestamps and
  // corrupt the heatmap and dashboard counts.
  // Unique ticket counting happens at render time via countUniqueTickets().
  function _mergeWorkbook(wb) {
    var added = 0;
    // Track unique tickets in THIS workbook for the import log message only
    var ticketsInFile = new Set();

    wb.SheetNames.forEach(function(sn) {
      if (sn === 'Instructions') return;
      var rows = XLSX.utils.sheet_to_json(wb.Sheets[sn], { defval:'' });
      rows.forEach(function(r) {
        var ticket = clean(r['Incident Ticket']||'');
        if (!ticket) return;
        ticketsInFile.add(ticket);

        var aosF = normBool(r['AOS']||'');
        var aosI = normBool(r['AOS Issue']||'');
        // KSL-4 uses 'DATE_TIME_Breach'; KM-1 uses 'DATE_TIME_Breach UTC'
        var breachRaw = r['DATE_TIME_Breach'] || r['DATE_TIME_Breach UTC'] || '';
        window.RPT.allData.push({
          ticket:         ticket,
          dateClosed:     formatDate(r['DATE_CLOSE']||''),
          dateTimeBreach: formatDate(breachRaw),
          status:         clean(r['Status']||'N/A'),
          queue:          clean(r['Queue']||''),
          priority:       clean(r['Priority']||''),
          language:       clean(r['ISO_Language']||'Unknown'),
          tool:           clean(r['Tool']||'Unknown'),
          topic:          clean(r['TOPIC']||''),
          sla:            clean(r['SLA_Code']||'Unknown'),
          slaN:           clean(r['SLA_N']||''),
          breachDesc:     clean(r['Breach_Description']||''),
          compassId:      clean(r['COMPASS ID']||''),
          reason:         clean(r['Reason']||''),
          aos:            aosF,
          agent:          clean(r['Agent']||''),
          bmsId:          clean(r['BMS ID']||''),
          comment:        clean(r['Comment']||''),
          aosIssue:       aosI,
          excluded:       normBool(r['Excluded']||''),
          jira:           clean(r['Jira']||''),
          week:           clean(r['Week']||''),
          unique:         clean(r['Unique']||''),
          sheet:          sn,
          isAos:          aosF==='Y'||aosI==='Y',
        });
        added++;
      });
    });
    window.RPT.aosFiltered = window.RPT.allData.filter(function(r){return r.isAos;});
    return { added: added, uniqueTickets: ticketsInFile.size };
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
    // Numeric sort for week numbers; alpha sort for the rest
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
    var c=document.getElementById('record-count');
    if(c) c.textContent=window.RPT.filtered.length+' records';
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
