/* reporter-importer.js
   Handles the Import button, file-input, drag-and-drop zone,
   and the Clear-All button in reporter.html.
   Depends on: XLSX (cdn), reporter-processor.js

   Flow:
     1. User selects / drops .xlsx file(s)
     2. FileReader reads each file as ArrayBuffer
     3. XLSX.read() parses it into a SheetJS workbook
     4. Workbook is tagged with ._filename for the import log
     5. Array of workbooks passed to window.RPT.loadFiles()
        (reporter-processor.js owns all data merging from this point)
*/
(function () {
  'use strict';

  /* ── helpers ─────────────────────────────────────────────────── */
  function setStatus(msg, type) {
    var el = document.getElementById('upload-status');
    if (!el) return;
    el.textContent = msg;
    el.className = type || '';
  }

  /* ── file processing ─────────────────────────────────────────── */
  function processFiles(files) {
    if (!files || !files.length) return;

    if (typeof XLSX === 'undefined') {
      setStatus('SheetJS library not loaded — cannot read .xlsx files.', 'err');
      return;
    }
    if (typeof window.RPT === 'undefined' || typeof window.RPT.loadFiles !== 'function') {
      setStatus('Reporter processor not ready. Please refresh the page.', 'err');
      return;
    }

    setStatus('Reading file(s)…', '');

    var fileArray = Array.prototype.slice.call(files);
    var workbooks = [];
    var errors    = [];
    var pending   = fileArray.length;

    fileArray.forEach(function (file) {
      if (!/\.xlsx$/i.test(file.name)) {
        errors.push(file.name + ' — only .xlsx files are accepted');
        pending--;
        if (pending === 0) done();
        return;
      }

      var reader = new FileReader();

      reader.onload = function (e) {
        try {
          var wb = XLSX.read(e.target.result, { type: 'array', cellDates: true });
          wb._filename = file.name;   // tag for import log in processor
          workbooks.push(wb);
        } catch (err) {
          errors.push(file.name + ' — ' + err.message);
        }
        pending--;
        if (pending === 0) done();
      };

      reader.onerror = function () {
        errors.push(file.name + ' — file read error');
        pending--;
        if (pending === 0) done();
      };

      reader.readAsArrayBuffer(file);
    });

    function done() {
      setStatus('');

      if (errors.length) {
        // Show any pre-parse errors in the import log area
        var logEl = document.getElementById('rpt-import-log');
        if (logEl) {
          logEl.innerHTML = errors.map(function (e) {
            return '<span style="color:var(--error,#a12c7b);margin-right:8px;">&#10007;</span>' + e;
          }).join('<br>');
          logEl.style.display = 'block';
        }
      }

      if (!workbooks.length) {
        if (!errors.length) setStatus('No valid .xlsx files selected.', 'err');
        return;
      }

      // Hand off to processor — it handles merging, filter rebuild, and renderTables()
      window.RPT.loadFiles(workbooks);
    }
  }

  /* ── DOM wiring ──────────────────────────────────────────────── */
  document.addEventListener('DOMContentLoaded', function () {

    /* Import button → opens file picker */
    var importBtn = document.getElementById('import-btn');
    var fileInput = document.getElementById('rpt-file-input');
    if (importBtn && fileInput) {
      importBtn.addEventListener('click', function () {
        fileInput.value = '';
        fileInput.click();
      });
      fileInput.addEventListener('change', function () {
        processFiles(fileInput.files);
      });
    }

    /* Clear All button */
    var clearBtn = document.getElementById('rpt-clear-btn');
    if (clearBtn) {
      clearBtn.addEventListener('click', function () {
        if (typeof window.RPT !== 'undefined' && typeof window.RPT.clearAll === 'function') {
          window.RPT.clearAll();
        }
      });
    }

    /* Drop zone */
    var dropZone = document.getElementById('drop-zone');
    if (dropZone) {
      dropZone.addEventListener('click', function () {
        if (fileInput) { fileInput.value = ''; fileInput.click(); }
      });
      dropZone.addEventListener('dragover', function (e) {
        e.preventDefault();
        dropZone.classList.add('drag-over');
      });
      dropZone.addEventListener('dragleave', function () {
        dropZone.classList.remove('drag-over');
      });
      dropZone.addEventListener('drop', function (e) {
        e.preventDefault();
        dropZone.classList.remove('drag-over');
        var files = e.dataTransfer && e.dataTransfer.files;
        if (files && files.length) processFiles(files);
      });
    }
  });

}());
