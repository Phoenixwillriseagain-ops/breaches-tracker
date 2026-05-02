// ui-renderer.js - UI rendering for Breaches Tracker with optimizations
(function() {
  'use strict';

  const { TAB_DEFS, OUT_COLS, MANUAL_COLS, CONFIG } = window.BT;
  const { MAX_DISPLAY_ROWS } = CONFIG;
  const state = window.BT.trackerState;

  // Escape HTML to prevent XSS
  function escapeHTML(str) {
    const d = document.createElement('div');
    d.textContent = String(str);
    return d.innerHTML;
  }

  // Render KPI cards using DocumentFragment
  function renderKPIs() {
    const container = document.getElementById('kpi-area');
    if (!container) return;

    const total = Object.values(state.tabs).reduce(function(sum, arr) { return sum + arr.length; }, 0);
    const frag = document.createDocumentFragment();

    // Total card
    const totalCard = document.createElement('div');
    totalCard.className = 'kpi';
    totalCard.innerHTML = '<div class="kpi-label">Total</div><div class="kpi-value">' + total + '</div><div class="kpi-sub">all categories</div>';
    frag.appendChild(totalCard);

    // Per-tab cards
    TAB_DEFS.forEach(function(t) {
      const card = document.createElement('div');
      card.className = 'kpi';
      card.innerHTML = '<div class="kpi-label">' + t.label + '</div><div class="kpi-value">' + state.tabs[t.id].length + '</div><div class="kpi-sub">breaches</div>';
      frag.appendChild(card);
    });

    container.innerHTML = '';
    container.appendChild(frag);
  }

  // Render tab buttons with badges
  function renderTabs() {
    const container = document.getElementById('tabs-bar');
    if (!container) return;

    const frag = document.createDocumentFragment();
    TAB_DEFS.forEach(function(t) {
      const btn = document.createElement('button');
      btn.className = 'tab-btn' + (t.id === state.activeTab ? ' active' : '');
      btn.setAttribute('data-tab', t.id);
      btn.innerHTML = t.label + '<span class="tab-badge">' + state.tabs[t.id].length + '</span>';
      btn.addEventListener('click', function() { switchTab(t.id); });
      frag.appendChild(btn);
    });

    container.innerHTML = '';
    container.appendChild(frag);
  }

  // Switch active tab
  function switchTab(id) {
    state.activeTab = id;
    renderTabs();
    renderTable(id);
  }

  // Sort data by column
  function sortData(tabId, col) {
    const data = state.tabs[tabId];
    if (!data || !data.length) return;

    if (state.sortCol === col) {
      state.sortDir = state.sortDir === 'asc' ? 'desc' : 'asc';
    } else {
      state.sortCol = col;
      state.sortDir = 'asc';
    }

    const dir = state.sortDir === 'asc' ? 1 : -1;
    data.sort(function(a, b) {
      const va = String(a[col] || '').toLowerCase();
      const vb = String(b[col] || '').toLowerCase();
      if (va < vb) return -1 * dir;
      if (va > vb) return 1 * dir;
      return 0;
    });

    renderTable(tabId);
    renderTableHeaders(tabId);
  }

  // Render table headers with sort indicators
  function renderTableHeaders(tabId) {
    const tbl = document.getElementById('data-table');
    if (!tbl || !tbl.tHead) return;
    const ths = tbl.tHead.querySelectorAll('th');
    ths.forEach(function(th, i) {
      th.className = '';
      th.removeAttribute('data-sort');
      if (MANUAL_COLS.indexOf(OUT_COLS[i]) >= 0) th.classList.add('manual-col');
      if (OUT_COLS[i] === state.sortCol) {
        th.classList.add(state.sortDir === 'asc' ? 'sorted-asc' : 'sorted-desc');
      }
      th.addEventListener('click', function() { sortData(tabId, OUT_COLS[i]); });
    });
  }

  // Render table using DocumentFragment (optimized DOM)
  function renderTable(tabId) {
    const container = document.querySelector('.table-wrap');
    if (!container) return;

    const data = state.tabs[tabId] || [];
    const rowCount = document.getElementById('row-count');
    if (rowCount) rowCount.textContent = data.length + ' records';

    const tbl = document.getElementById('data-table');
    if (!tbl) return;

    tbl.innerHTML = '';

    if (!data.length) {
      tbl.innerHTML = '<tr><td colspan="20" style="text-align:center;padding:3rem;color:var(--muted)">No breaches in this category</td></tr>';
      return;
    }

    // Build thead
    const thead = document.createElement('thead');
    const tr = document.createElement('tr');
    OUT_COLS.forEach(function(col) {
      const th = document.createElement('th');
      th.textContent = col;
      if (MANUAL_COLS.indexOf(col) >= 0) th.classList.add('manual-col');
      if (col === state.sortCol) th.classList.add(state.sortDir === 'asc' ? 'sorted-asc' : 'sorted-desc');
      th.addEventListener('click', function() { sortData(tabId, col); });
      tr.appendChild(th);
    });
    thead.appendChild(tr);
    tbl.appendChild(thead);

    // Build tbody with DocumentFragment
    const tbody = document.createElement('tbody');
    const frag = document.createDocumentFragment();
    const limit = Math.min(data.length, MAX_DISPLAY_ROWS);
    const isTruncated = data.length > MAX_DISPLAY_ROWS;

    for (let i = 0; i < limit; i++) {
      const row = data[i];
      const tr = document.createElement('tr');
      OUT_COLS.forEach(function(col) {
        const td = document.createElement('td');
        const val = row[col] || '';
        td.textContent = val;
        if (MANUAL_COLS.indexOf(col) >= 0) td.classList.add('manual-col');
        td.title = String(val);
        tr.appendChild(td);
      });
      frag.appendChild(tr);
    }

    tbody.appendChild(frag);
    tbl.appendChild(tbody);

    if (isTruncated) {
      const note = document.createElement('tr');
      note.innerHTML = '<td colspan="20" style="text-align:center;color:var(--muted);font-size:0.75rem;">Showing ' + limit + ' of ' + data.length + ' records. Export for full data.</td>';
      tbody.appendChild(note);
    }
  }

  // Show/hide empty state
  function showEmptyState(show) {
    const empty = document.getElementById('empty-state');
    const table = document.getElementById('data-table');
    const kpi = document.getElementById('kpi-area');
    const tabs = document.getElementById('tabs-bar');
    const toolbar = document.querySelector('.toolbar');
    const tableWrap = document.querySelector('.table-wrap');

    if (empty) empty.style.display = show ? '' : 'none';
    if (table) table.style.display = show ? 'none' : '';
    if (kpi) kpi.style.display = show ? 'none' : '';
    if (tabs) tabs.style.display = show ? 'none' : '';
    if (toolbar) toolbar.style.display = show ? 'none' : '';
    if (tableWrap) tableWrap.style.display = show ? 'none' : '';
  }

  // Toggle export buttons visibility
  function updateExportButtons() {
    const total = Object.values(state.tabs).reduce(function(sum, arr) { return sum + arr.length; }, 0);
    const exportAll = document.getElementById('export-all-btn');
    if (exportAll) exportAll.style.display = total > 0 ? '' : 'none';
  }

  // Main render orchestration
  function render() {
    const total = Object.values(state.tabs).reduce(function(sum, arr) { return sum + arr.length; }, 0);
    showEmptyState(total === 0);
    if (total > 0) {
      if (!state.activeTab) state.activeTab = TAB_DEFS[0].id;
      renderKPIs();
      renderTabs();
      renderTable(state.activeTab);
      updateExportButtons();
    }
  }

  // Theme toggle
  function initTheme() {
    const btn = document.querySelector('[data-theme-toggle]');
    const root = document.documentElement;
    let dark = window.matchMedia('(prefers-color-scheme: dark)').matches;
    let current = dark ? 'dark' : 'light';
    root.setAttribute('data-theme', current);
    if (btn) {
      btn.addEventListener('click', function() {
        current = current === 'dark' ? 'light' : 'dark';
        root.setAttribute('data-theme', current);
      });
    }
  }

  // Keyboard shortcuts
  function initKeyboardShortcuts() {
    document.addEventListener('keydown', function(e) {
      // Ctrl+E: Export tab
      if (e.ctrlKey && e.key === 'e') {
        e.preventDefault();
        if (state.activeTab && state.tabs[state.activeTab]) {
          window.BT.exportTab(state.activeTab);
        }
      }
      // Ctrl+Shift+E: Export all
      if (e.ctrlKey && e.shiftKey && e.key === 'E') {
        e.preventDefault();
        window.BT.exportAll();
      }
      // Ctrl+F: Focus search/file input
      if (e.ctrlKey && e.key === 'f' && !document.querySelector('input[type=text]:focus')) {
        e.preventDefault();
        const input = document.getElementById('file-input');
        if (input) input.click();
      }
      // Tab switching with Ctrl+1-8
      if (e.ctrlKey && e.key >= '1' && e.key <= '8') {
        const idx = parseInt(e.key) - 1;
        if (TAB_DEFS[idx]) {
          e.preventDefault();
          switchTab(TAB_DEFS[idx].id);
        }
      }
    });
  }

  // Expose to global
  window.render = render;
  window.renderKPIs = renderKPIs;
  window.renderTabs = renderTabs;
  window.renderTable = renderTable;
  window.switchTab = switchTab;
  window.BT.renderEmptyState = showEmptyState;
  window.BT.updateExportButtons = updateExportButtons;
  window.BT.initTheme = initTheme;
  window.BT.initKeyboardShortcuts = initKeyboardShortcuts;


  	// Table search/filter function
	function searchTable(searchValue) {
		const table = document.querySelector('.data-table');
		if (!table) return;
		
		const rows = table.querySelectorAll('tbody tr');
		const lowerSearch = searchValue.toLowerCase();
		
		rows.forEach(row => {
			const text = row.textContent.toLowerCase();
			row.style.display = text.includes(lowerSearch) ? '' : 'none';
		});
	}
	
	window.searchTable = searchTable;
})();
