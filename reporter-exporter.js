// reporter-exporter.js - Export data in V2 model format
(function() {
  'use strict';

  function formatDate(val) {
    if (!val) return '';
    const d = new Date(val);
    if (isNaN(d.getTime())) return String(val);
    const TZ_OFFSET_MS = 2 * 60 * 60 * 1000;
    const local = new Date(d.getTime() + TZ_OFFSET_MS);
    const dd  = String(local.getUTCDate()).padStart(2, '0');
    const mm  = String(local.getUTCMonth() + 1).padStart(2, '0');
    const yyyy = local.getUTCFullYear();
    const hh  = String(local.getUTCHours()).padStart(2, '0');
    const min = String(local.getUTCMinutes()).padStart(2, '0');
    const ss  = String(local.getUTCSeconds()).padStart(2, '0');
    return `${dd}.${mm}.${yyyy} ${hh}:${min}:${ss}`;
  }

  const V2_HEADERS = [
    'Incident Ticket','DATE_CLOSE','Status','Queue','Priority',
    'ISO_Language','Tool','TOPIC','SLA_Code','SLA_N',
    'Breach_Description','DATE_TIME_Breach','Munich time','COMPASS ID',
    'Reason','AOS','Agent','BMS ID','Comment','AOS Issue',
    'Excluded','Jira','Week','Unique',
  ];

  function rowToArray(row) {
    return [
      row.ticket        || '',
      row.dateClosed    || '',
      row.status        || '',
      row.queue         || '',
      row.priority      || '',
      row.language      || '',
      row.tool          || '',
      row.topic         || '',
      row.sla           || '',
      row.slaN          || '',
      row.breachDesc    || '',
      row.dateTimeBreach|| '',
      row.munichTime    || '',
      row.compassId     || '',
      row.reason        || '',
      row.aos           || '',
      row.agent         || '',
      row.bmsId         || '',
      row.comment       || '',
      row.aosIssue      || '',
      row.excluded      || '',
      row.jira          || '',
      row.week          || '',
      row.unique        || '',
    ];
  }

  window.EXPORTER = {
    exportAsExcel: function(data, filename) {
      if (!data || !data.length) { alert('No data to export'); return; }
      const ws_data = [V2_HEADERS].concat(data.map(rowToArray));
      const ws = XLSX.utils.aoa_to_sheet(ws_data);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Breaches');
      XLSX.writeFile(wb, filename || `Breaches-Export-${new Date().toISOString().slice(0,10).replace(/-/g,'')}.xlsx`);
    },

    exportAsCSV: function(data, filename) {
      if (!data || !data.length) { alert('No data to export'); return; }
      const escape = v => {
        const s = String(v || '');
        return (s.includes(',') || s.includes('"') || s.includes('\n'))
          ? '"' + s.replace(/"/g, '""') + '"' : s;
      };
      let csv = V2_HEADERS.map(escape).join(',') + '\n';
      data.forEach(row => { csv += rowToArray(row).map(escape).join(',') + '\n'; });
      const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = filename || `Breaches-Export-${new Date().toISOString().slice(0,10).replace(/-/g,'')}.csv`;
      a.style.display = 'none';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
    },

    exportAosOnly: function(data) {
      const sub = data.filter(r => r.isAos);
      if (!sub.length) { alert('No AOS Portal Issues to export'); return; }
      this.exportAsExcel(sub, `AOS-Portal-Issues-${new Date().toISOString().slice(0,10).replace(/-/g,'')}.xlsx`);
    },

    exportKSL5aOnly: function(data) {
      const sub = data.filter(r => r.sla === 'KSL-5a');
      if (!sub.length) { alert('No KSL-5a records to export'); return; }
      this.exportAsExcel(sub, `KSL-5a-Breaches-${new Date().toISOString().slice(0,10).replace(/-/g,'')}.xlsx`);
    },

    exportExcludedOnly: function(data) {
      const sub = data.filter(r => r.excluded === 'Y');
      if (!sub.length) { alert('No excluded records to export'); return; }
      this.exportAsExcel(sub, `Excluded-Records-${new Date().toISOString().slice(0,10).replace(/-/g,'')}.xlsx`);
    },
  };
})();
