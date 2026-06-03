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

  // ─── V2 XLSX column mappings (0-based index) ─────────────────────────────────────
  // V2 "Breaches" file headers (all sheets share this structure):
  // 0:Incident Ticket  1:DATE_CLOSE  2:Status  3:Queue  4:Priority
  // 5:ISO_Language  6:Tool  7:TOPIC  8:SLA_Code  9:SLA_N
  // 10:Breach_Description  11:DATE_TIME_Breach
  // 12:COMPASS ID  13:Reason  14:AOS  15:Agent  16:BMS ID
  // 17:Comment  18:AOS Issue  19:Excluded  20:Jira  21:Week  22:Unique
  const C_XLSX = {
    ticket:     0,
    date_close: 1,
    status:     2,
    queue:      3,
    priority:   4,
    lang:       5,
    tool:       6,
    topic:      7,
    sla_code:   8,
    sla_n:      9,
    breach_desc:10,
    breach_dt:  11,
    compass_id: 12,
    reason:     13,
    aos:        14,
    agent:      15,
    bms_id:     16,
    comment:    17,
    aos_issue:  18,
    excluded:   19,
    jira:       20,
    week:       21,
    unique:     22,
    nok:        18,
  };

  // ─── CSV column mappings (ServiceNow SLAs-Details export) ────────────────────────
  // Headers (0-based) from "3093-SLAs-Details" ServiceNow export:
  //  0: SLA Kennzahl                          → sla_code
  //  8: Ticket Priorität                      → priority
  // 10: SLA Prüfung                           → sla_n
  // 11: SLA Prüfung Name                      → breach_desc
  // 12: SLA Prüfung Zeitpunkt                 → breach_dt
  // 13: SLA Prüfung ist OK/NOK?               → nok
  // 19: Element geschlossen am                → date_close
  // 20: Element Status                        → status
  // 21: Incident Element Supportsprache       → lang  (ISO code: de, en, fr …)
  // 22: Incident Ticket                       → ticket
  // 23: Ticket Gruppe                         → queue
  // 28: KM Anwendung                          → tool / topic  (#SW#AOS, #SW#ISTA …)
  const C_CSV = {
    sla_code:    0,
    priority:    8,
    sla_n:       10,
    breach_desc: 11,
    breach_dt:   12,
    nok:         13,
    date_close:  19,
    status:      20,
    lang:        21,
    ticket:      22,
    queue:       23,
    tool:        28,
    topic:       28,
  };

  // Output column headers (standardized order for export)
  const OUT_COLS = [
    'Incident Ticket', 'DATE_CLOSE', 'Status', 'Queue', 'Priority',
    'ISO_Language', 'Tool', 'TOPIC', 'SLA_Code', 'SLA_N',
    'Breach_Description', 'DATE_TIME_Breach', 'COMPASS ID',
    'Reason', 'AOS', 'Agent', 'BMS ID', 'Comment', 'AOS Issue',
    'Excluded', 'Jira', 'Week', 'Unique',
  ];

  // Columns that are manually editable (highlighted in table)
  const MANUAL_COLS = [
    'Agent', 'BMS ID', 'Comment', 'AOS Issue',
    'Excluded', 'Jira', 'Week', 'Unique',
  ];

  const FILTER_DEBOUNCE_MS = 200;
  const MAX_DISPLAY_ROWS = 500;

  const EXPORT_PREFIXES = {
    tab: 'Breaches_',
    all: 'Breaches_All_',
    report: 'Breaches_FullReport_',
  };

  const CHART_COLORS = {
    primary: ['#01696f', '#4f98a3', '#0c4e54', '#227f8b'],
    success: ['#437a22', '#6daa45'],
    warn:    ['#964219', '#bb653b'],
    error:   ['#a12c7b', '#d163a7'],
    muted:   ['#7a7974', '#797876'],
    light:   ['#cedcd8', '#313b3b'],
  };

  const REPORTER_TABS = [
    { id: 'overview', label: 'Overview' },
    { id: 'reasons',  label: 'Breach Reasons' },
    { id: 'apps',     label: 'Applications' },
    { id: 'aos',      label: 'AOS Portal Issues' },
    { id: 'km1',      label: 'KM-1 Detail' },
    { id: 'all',      label: 'All Records' },
  ];

  const REPORTER_SLA_CODES = ['KSL-4', 'KM-1', 'KSL-5a', 'KM-2'];

  const REPORTER_STORAGE_KEY = 'reporter_data';

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
