// reporter-uploader.js — wires the import button, drop zone and Clear All button
(function () {
  'use strict';

  document.addEventListener('DOMContentLoaded', function () {

    var importBtn  = document.getElementById('import-btn');
    var fileInput  = document.getElementById('rpt-file-input');
    var dropZone   = document.getElementById('drop-zone');
    var clearBtn   = document.getElementById('rpt-clear-btn');
    var statusEl   = document.getElementById('upload-status');

    /* ---- helpers ---- */
    function setStatus(msg, type) {
      if (!statusEl) return;
      statusEl.textContent = msg;
      statusEl.className   = type || '';
    }

    function handleFiles(files) {
      if (!files || !files.length) return;
      var xlsxFiles = Array.prototype.filter.call(files, function (f) {
        return /\.xlsx$/i.test(f.name);
      });
      if (!xlsxFiles.length) {
        setStatus('No .xlsx files found — please select at least one .xlsx file.', 'err');
        return;
      }
      setStatus('Importing ' + xlsxFiles.length + ' file(s)…', '');
      window.RPT.loadFiles(xlsxFiles);
    }

    /* ---- Import button ---- */
    if (importBtn && fileInput) {
      importBtn.addEventListener('click', function () {
        fileInput.value = '';        // allow re-selecting the same file
        fileInput.click();
      });
      fileInput.addEventListener('change', function () {
        handleFiles(fileInput.files);
      });
    }

    /* ---- Drop zone click ---- */
    if (dropZone && fileInput) {
      dropZone.addEventListener('click', function () {
        fileInput.value = '';
        fileInput.click();
      });
    }

    /* ---- Drag-and-drop ---- */
    if (dropZone) {
      dropZone.addEventListener('dragover', function (e) {
        e.preventDefault();
        e.stopPropagation();
        dropZone.classList.add('drag-over');
      });

      dropZone.addEventListener('dragleave', function (e) {
        e.preventDefault();
        e.stopPropagation();
        dropZone.classList.remove('drag-over');
      });

      dropZone.addEventListener('drop', function (e) {
        e.preventDefault();
        e.stopPropagation();
        dropZone.classList.remove('drag-over');
        handleFiles(e.dataTransfer.files);
      });
    }

    /* ---- Clear All button ---- */
    if (clearBtn) {
      clearBtn.addEventListener('click', function () {
        if (!confirm('Clear all loaded data and return to the upload screen?')) return;
        window.RPT.clearData();
        setStatus('', '');
        if (fileInput) fileInput.value = '';
      });
    }

    /* ---- Global drag-and-drop on the whole page (outside the drop zone) ---- */
    document.addEventListener('dragover', function (e) { e.preventDefault(); });
    document.addEventListener('drop', function (e) {
      // Only handle if the drop zone itself didn't catch it
      if (dropZone && dropZone.contains(e.target)) return;
      e.preventDefault();
      handleFiles(e.dataTransfer.files);
    });

  });

})();
