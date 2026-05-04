// reporter-exporter.js - Export data in optimized Template format
(function() {
  'use strict';

  window.EXPORTER = {
    // Export filtered data as Excel file in Template-Example format
    exportAsExcel: function(data, filename) {
      if (!data || !data.length) {
        alert('No data to export');
        return;
      }
      const ws_data = [];
      ws_data.push(['Incident Ticket', 'Date Received', 'Date Resolved', 'Date Closed', 'Status', 'Ticket Group', 'Language', 'Week', 'SLA Category', 'AOS Portal', 'KM-1', 'Excluded', 'Reason']);
      data.forEach(row => {
        ws_data.push([
          row.ticket || '',
          row.dateReceived || '',
          row.dateResolved || '',
          row.dateClosed || '',
          row.status || 'N/A',
          row.ticketGroup || 'Unknown',
          row.language || 'Unknown',
          row.month || 'Unknown',
          row.sla || 'Unknown',
          row.isAos ? 'Y' : 'N',
          row.kml || 'N',
          row.excluded || 'N',
          row.reason || ''
        ]);
      });
      const ws = XLSX.utils.aoa_to_sheet(ws_data);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Breaches');
      const date = new Date().toISOString().split('T')[0];
      XLSX.writeFile(wb, filename || `Breaches-Export-${date}.xlsx`);
    },

    // Export as CSV
    exportAsCSV: function(data, filename) {
      if (!data || !data.length) {
        alert('No data to export');
        return;
      }
      let csv = 'Incident Ticket,Date Received,Date Resolved,Date Closed,Status,Ticket Group,Language,Week,SLA Category,AOS Portal,KM-1,Excluded,Reason\n';
      data.forEach(row => {
        const values = [
          this.escapeCsv(row.ticket || ''),
          this.escapeCsv(row.dateReceived || ''),
          this.escapeCsv(row.dateResolved || ''),
          this.escapeCsv(row.dateClosed || ''),
          this.escapeCsv(row.status || 'N/A'),
          this.escapeCsv(row.ticketGroup || 'Unknown'),
          this.escapeCsv(row.language || 'Unknown'),
          this.escapeCsv(row.month || 'Unknown'),
          this.escapeCsv(row.sla || 'Unknown'),
          row.isAos ? 'Y' : 'N',
          row.kml || 'N',
          row.excluded || 'N',
          this.escapeCsv(row.reason || '')
        ];
        csv += values.join(',') + '\n';
      });
      const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
      const link = document.createElement('a');
      const url = URL.createObjectURL(blob);
      link.setAttribute('href', url);
      link.setAttribute('download', filename || `Breaches-Export-${new Date().toISOString().split('T')[0]}.csv`);
      link.style.visibility = 'hidden';
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    },

    // Export AOS Portal Issues only
    exportAosOnly: function(data, filename) {
      const aosData = data.filter(row => row.isAos === true);
      if (!aosData.length) {
        alert('No AOS Portal Issues to export');
        return;
      }
      this.exportAsExcel(aosData, filename || `AOS-Portal-Issues-${new Date().toISOString().split('T')[0]}.xlsx`);
    },

    // Export KM-1 records only
    exportKM1Only: function(data, filename) {
      const km1Data = data.filter(row => row.kml === 'Y');
      if (!km1Data.length) {
        alert('No KM-1 records to export');
        return;
      }
      this.exportAsExcel(km1Data, filename || `KM-1-Analysis-${new Date().toISOString().split('T')[0]}.xlsx`);
    },

    // Export excluded records
    exportExcludedOnly: function(data, filename) {
      const excludedData = data.filter(row => row.excluded === 'Y');
      if (!excludedData.length) {
        alert('No excluded records to export');
        return;
      }
      this.exportAsExcel(excludedData, filename || `Excluded-Records-${new Date().toISOString().split('T')[0]}.xlsx`);
    },

    // Helper function to escape CSV values
    escapeCsv: function(value) {
      if (!value) return '';
      const str = String(value);
      if (str.includes(',') || str.includes('"') || str.includes('\n')) {
        return '"' + str.replace(/"/g, '""') + '"';
      }
      return str;
    }
  };
})();
