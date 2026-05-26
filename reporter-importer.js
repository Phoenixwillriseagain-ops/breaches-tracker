/* reporter-importer.js
   Handles the Import button, file-input, drag-and-drop zone,
   and the Clear-All button in reporter.html.
   Depends on: XLSX (cdn), reporter-processor.js
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

  function showLog(lines) {
    var el = document.getElementById('rpt-import-log');
    if (!el) return;
    if (!lines || !lines.length) { el.style.display = 'none'; el.innerHTML = ''; return; }
    el.style.display = 'block';
    el.innerHTML = lines.map(function (l) { return '<div>' + l + '</div>'; }).join('');
  }

  function showData() {
    var up = document.getElementById('upload-section');
    var ds = document.getElementById('data-section');
    if (up) up.style.display = 'none';
    if (ds) { ds.style.display = 'flex'; ds.classList.add('visible'); }
  }

  function hideData() {
    var up = document.getElementById('upload-section');
    var ds = document.getElementById('data-section');
    var sb = document.getElementById('rpt-sidebar');
    var eb = document.getElementById('export-btns');
    var cb = document.getElementById('rpt-clear-btn');
    var rc = document.getElementById('record-count');
    if (up) up.style.display = '';
    if (ds) { ds.style.display = 'none'; ds.classList.remove('visible'); }
    if (sb) sb.style.display = 'none';
    if (eb) eb.style.display = 'none';
    if (cb) cb.style.display = 'none';
    if (rc) rc.textContent = '';
    showLog([]);
    setStatus('');
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
    var results = [];
    var pending = fileArray.length;

    fileArray.forEach(function (file) {
      var reader = new FileReader();
      reader.onload = function (e) {
        try {
          var wb = XLSX.read(e.target.result, { type: 'array' });
          results.push({ name: file.name, wb: wb });
        } catch (err) {
          results.push({ name: file.name, error: err.message });
        }
        pending--;
        if (pending === 0) finalize(results);
      };
      reader.onerror = function () {
        results.push({ name: file.name, error: 'File read error' });
        pending--;
        if (pending === 0) finalize(results);
      };
      reader.readAsArrayBuffer(file);
    });
  }

  function finalize(results) {
    var logs = [];
    var workbooks = [];
    results.forEach(function (r) {
      if (r.error) {
        logs.push('&#10060; ' + r.name + ' — ' + r.error);
      } else {
        logs.push('&#9989; ' + r.name + ' loaded (' + r.wb.SheetNames.length + ' sheet(s))');
        workbooks.push(r.wb);
      }
    });

    if (!workbooks.length) {
      setStatus('No valid files could be read.', 'err');
      showLog(logs);
      return;
    }

    try {
      var summary = window.RPT.loadFiles(workbooks);
      if (summary && summary.log) logs = logs.concat(summary.log);
      showLog(logs);
      showData();
      setStatus('');
      if (typeof window.renderTables === 'function') window.renderTables();
    } catch (err) {
      setStatus('Error processing file(s): ' + err.message, 'err');
      showLog(logs.concat(['&#9888; Processing error: ' + err.message]));
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
        hideData();
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
