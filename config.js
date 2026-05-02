// config.js - Shared configuration for Breaches Tracker
// Constants, tab definitions, column mappings, and app settings

(function() {
  'use strict';

  // Tab definitions for index.html (SLA + language splits)
  const TAB_DEFS = [
    { id: 'ksl4-wri', label: 'KSL-4 WRI', code: 'KSL-4', lang: '!de' },
    { id: 'ksl4-de',  label: 'KSL-4 DE',  code: 'KSL-4', lang: 'de' },
    { id: 'km1-gr',   label: 'KM-1 GR',   code: 'KM-1',  lang: '!de', nokFilter: true },
    { id: 'km1-de',   label: 'KM-1 DE',   code: 'KM-1',  lang: 'de',  nokFilter: true },
    { id: 'ksl5a-gr', label: 'KSL-5a GR', code: 'KSL-5a', lang: '!de' },
    { id: 'ksl5a-de', label: 'KSL-5a DE', code: 'KSL-5a', lang: 'de' },
    { id: 'km2-gr',   label: 'KM-2 GR',   code: 'KM-2',  lang: '!de' },
    { id: 'km2-de',   label: 'KM-2 DE',   code: 'KM-2',  lang: 'de' },
  ];

  // Column index mappings for XLSX input
  const C_XLSX = {
    ticket: 0,
    date_close: 3,
    status: 4,
    queue: 5,
    priority: 6,
    lang: 12,
    tool: 14,
    topic: 15,
    sla_code: 19,
    sla_n: 24,
    breach_desc: 25,
    breach_dt: 26,
    nok: 27,
  };

  // Column index mappings for CSV input
  const C_CSV = {
    sla_code: 0,
    priority: 8,
    sla_n: 10,
    breach_desc: 11,
    breach_dt: 12,
    nok: 13,
    tool: 15,
    date_close: 18,
    status: 19,
    lang: 20,
    ticket: 21,
    queue: 22,
    topic: 27,
  };

  // Output column headers (standardized order for export)
  const OUT_COLS = [
    'Incident Ticket', 'DATE_CLOSE', 'Status', 'Queue', 'Priority',
    'ISO_Language', 'Tool', 'TOPIC', 'SLA_Code', 'SLA_N',
    'Breach_Description', 'DATE_TIME_Breach', 'Agent', 'BMS ID',
    'Comment if excluded', 'Additional comment', 'Excluded',
    'Jira', 'Week', 'Unique',
  ];

  // Columns that are manually editable (highlighted in table)
  const MANUAL_COLS = [
    'Agent', 'BMS ID', 'Comment if excluded', 'Additional comment',
    'Excluded', 'Jira', 'Week', 'Unique',
  ];

  // Debounce delay for filter inputs (ms)
  const FILTER_DEBOUNCE_MS = 200;

  // Max rows to display in tables (prevents DOM overload)
  const MAX_DISPLAY_ROWS = 500;

  // Export filename prefixes
  const EXPORT_PREFIXES = {
    tab: 'Breaches_',
    all: 'Breaches_All_',
    report: 'Breaches_FullReport_',
  };

  // Chart color palettes (theme-aware, used by reporter)
  const CHART_COLORS = {
    primary: ['#01696f', '#4f98a3', '#0c4e54', '#227f8b'],
    success: ['#437a22', '#6daa45'],
    warn:    ['#964219', '#bb653b'],
    error:   ['#a12c7b', '#d163a7'],
    muted:   ['#7a7974', '#797876'],
    light:   ['#cedcd8', '#313b3b'],
  };

  // Reporter tab definitions
  const REPORTER_TABS = [
    { id: 'overview', label: 'Overview' },
    { id: 'reasons',  label: 'Breach Reasons' },
    { id: 'apps',     label: 'Applications' },
    { id: 'aos',      label: 'AOS Portal Issues' },
    { id: 'km1',      label: 'KM-1 Detail' },
    { id: 'all',      label: 'All Records' },
  ];

  // Reporter SLA codes for that page
  const REPORTER_SLA_CODES = ['KSL-4', 'KM-1', 'KSL-5a', 'KM-2', 'AOS'];

  // LocalStorage key for reporter data persistence
  const REPORTER_STORAGE_KEY = 'reporter_data';

  // Expose to global scope for HTML scripts
  window.BT = window.BT || {};
  window.BT.TAB_DEFS = TAB_DEFS;
  window.BT.C_XLSX = C_XLSX;
  window.BT.C_CSV = C_CSV;
  window.BT.OUT_COLS = OUT_COLS;
  window.BT.MANUAL_COLS = MANUAL_COLS;
  window.BT.REPORTER_TABS = REPORTER_TABS;
  window.BT.REPORTER_SLA_CODES = REPORTER_SLA_CODES;
  window.BT.REPORTER_STORAGE_KEY = REPORTER_STORAGE_KEY;
  window.BT.CONFIG = {
    FILTER_DEBOUNCE_MS,
    MAX_DISPLAY_ROWS,
    EXPORT_PREFIXES,
    CHART_COLORS,
  };

})();
