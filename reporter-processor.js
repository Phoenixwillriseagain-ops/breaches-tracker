// reporter-processor.js — V2 model processor for Breaches Reporter
(function() {
  'use strict';

  window.RPT = {
    allData: [],
    filtered: [],
    aosFiltered: [],
    uniqueValues: {},

    loadFile: function(file) { _loadFile(file); },

    clearFilters: function() {
      ['filter-week','filter-sla','filter-lang','filter-excl',
       'filter-aos','filter-tool','filter-queue','filter-sheet']
        .forEach(function(id){
          var el=document.getElementById(id);
          if(el) el.value='All';
        });
      _applyFilters();
    },

    resetToUpload: function() { location.reload(); },

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

  /* ---- processWorkbook ---- */
  function _processWorkbook(wb){
    var data=[];
    wb.SheetNames.forEach(function(sn){
      if(sn==='Instructions') return;
      var rows=XLSX.utils.sheet_to_json(wb.Sheets[sn],{defval:''});
      rows.forEach(function(r){
        var ticket=clean(r['Incident Ticket']||'');
        if(!ticket) return;
        var aosF=normBool(r['AOS']||'');
        var aosI=normBool(r['AOS Issue']||'');
        data.push({
          ticket:         ticket,
          dateClosed:     formatDate(r['DATE_CLOSE']||''),
          dateTimeBreach: formatDate(r['DATE_TIME_Breach']||''),
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
      });
    });

    console.log('[RPT] loaded',data.length,'records');
    window.RPT.allData    = data;
    window.RPT.filtered   = data.slice();
    window.RPT.aosFiltered= data.filter(function(r){return r.isAos;});

    _populateUnique();
    _updateDropdowns();
    _applyFiltersOnly();

    window.renderTables();
  }

  function _loadFile(file){
    if(!file) return;
    var reader=new FileReader();
    reader.onload=function(e){
      try{
        var wb=XLSX.read(e.target.result,{type:'array',cellDates:true});
        _processWorkbook(wb);
      }catch(err){
        console.error('[RPT] parse error',err);
        alert('Error loading file: '+err.message);
      }
    };
    reader.readAsArrayBuffer(file);
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
    uv.weeks.sort();uv.slas.sort();uv.languages.sort();uv.tools.sort();uv.queues.sort();
    window.RPT.uniqueValues=uv;
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
    fill('filter-lang', uv.languages);
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
    var wF=g('filter-week'), sF=g('filter-sla'),  lF=g('filter-lang'),
        eF=g('filter-excl'), aF=g('filter-aos'),  tF=g('filter-tool'),
        qF=g('filter-queue'),shF=g('filter-sheet');
    window.RPT.filtered=window.RPT.allData.filter(function(r){
      if(wF !=='All'&&r.week     !==wF)  return false;
      if(sF !=='All'&&r.sla      !==sF)  return false;
      if(lF !=='All'&&r.language !==lF)  return false;
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

  document.addEventListener('DOMContentLoaded',function(){
    ['filter-week','filter-sla','filter-lang','filter-excl',
     'filter-aos','filter-tool','filter-queue','filter-sheet']
      .forEach(function(id){
        var el=document.getElementById(id);
        if(el) el.addEventListener('change',_applyFilters);
      });
  });

})();
