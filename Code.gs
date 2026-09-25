/**
 * Desk Stats: front desk tap counter for Motlow State's four campus libraries.
 *
 * One standalone Apps Script project. setup() builds the "Desk Stats"
 * workbook (every tab, heading, formula and starter row) and remembers its
 * id. doGet() serves the tap page (Index.html) and the live board
 * (Board.html, ?view=week) and the trends page (Dashboard.html, ?view=trends). Every
 * total anywhere adds up the Count column: a tap is +1, a subtraction is -1,
 * a typed-in history day is its number. Nothing counts rows. Everything a screen does
 * goes through the functions below; nothing else writes to the workbook.
 */

const APP = {
  name: 'Desk Stats',
  build: 'DS-2026-09-23-01', /* the pages carry the same stamp; a mismatch is reported on screen */
  timeZone: 'America/Chicago',  /* every date in this script is worked out in this zone, never in the project's own clock setting */
  campuses: ['Smyrna', 'Moore County', 'Fayetteville', 'McMinnville'],
  blocks: ['7:30 AM - Noon', 'Noon - 4:30 PM', '4:30 PM - Close'],
  modes: ['In person', 'Remote'],
  categories: [
    ['D2L', 'Aiding a patron with D2L navigation and usage. 5 to 15 minutes.'],
    ['Tech Support', 'Software, connectivity, printers, and copiers. 5 to 15 minutes.'],
    ['Directional', 'Limited to how to get from one place to another. Under 1 minute.'],
    ['General Ref', 'All questions not in other categories; hours, etc. Under 1 minute.'],
    ['Library Search', 'Performing or instructing a library search. 5 to 15 minutes.'],
    ['Reference & Instruction', 'Reference or library instruction to a single patron. 10 to 20 minutes.'],
    ['XR Lab', 'Scheduling or handling an XR lab appointment. 1 minute to 1 hour.'],
    ['3d Printing', 'Consulting on or printing a requested item. 5 to 30 minutes.'],
  ],
  referral: 'Any interaction requiring in-depth research or more than 20 minutes of time: please refer the patron to a subject specialist or librarian.',
};

const SHEETS = {
  log: { name: 'Log', headers: ['When', 'Date', 'Hour', 'Weekday', 'Campus', 'Category', 'Mode', 'Count', 'Recorded by', 'Source', 'Tap id'] },
  gate: { name: 'Gate', headers: ['Date', 'Campus', 'Block', 'Count', 'Recorded by', 'Source', 'When entered'] },
  gateHistory: { name: 'Gate history', headers: ['When entered', 'Date', 'Campus', 'Block', 'Count', 'Replaced', 'Recorded by', 'Source', 'Entry id'] },
  buttons: { name: 'Buttons', headers: ['Category', 'Helper text', 'Order', 'Show', 'Campus'] },
  campuses: { name: 'Campuses', headers: ['Campus', 'Block 1', 'Block 2', 'Block 3', 'Block 4'] },
};

const WEEKDAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

/* ------------------------------------------------------------------ */
/* Setup: run once from the editor. Safe to run again; it fills gaps.  */
/* ------------------------------------------------------------------ */

function setup() {
  const props = PropertiesService.getScriptProperties();
  let ss = null;
  const existing = props.getProperty('WORKBOOK_ID');
  if (existing) {
    try { ss = SpreadsheetApp.openById(existing); } catch (err) { ss = null; }
  }
  if (!ss) {
    ss = SpreadsheetApp.create(APP.name);
    props.setProperty('WORKBOOK_ID', ss.getId());
  }
  ss.setSpreadsheetTimeZone(APP.timeZone);

  const log = ensureSheet_(ss, SHEETS.log);
  log.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  log.getRange('B:B').setNumberFormat('yyyy-mm-dd');

  const gate = ensureSheet_(ss, SHEETS.gate);
  gate.getRange('A:A').setNumberFormat('yyyy-mm-dd');
  gate.getRange('G:G').setNumberFormat('yyyy-mm-dd hh:mm:ss');

  const gateHistory = ensureSheet_(ss, SHEETS.gateHistory);
  gateHistory.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  gateHistory.getRange('B:B').setNumberFormat('yyyy-mm-dd');

  const buttons = ensureSheet_(ss, SHEETS.buttons);
  let added = [];
  if (buttons.getLastRow() < 2) {
    buttons.getRange(2, 1, APP.categories.length, 5)
      .setValues(APP.categories.map((c, i) => [c[0], c[1], i + 1, 'yes', '']));
  } else {
    added = addMissingButtons_(buttons);
  }
  buttons.getRange('D2:D200').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['yes', 'no'], true).setAllowInvalid(false).build());
  buttons.setColumnWidth(2, 420);

  const campuses = ensureSheet_(ss, SHEETS.campuses);
  if (campuses.getLastRow() < 2) {
    campuses.getRange(2, 1, APP.campuses.length, 5)
      .setValues(APP.campuses.map(c => [c, APP.blocks[0], APP.blocks[1], APP.blocks[2], '']));
  }

  readCampuses_(ss).forEach(c => buildWeeklySheet_(ss, c));
  buildSummarySheet_(ss);

  const first = ss.getSheetByName('Sheet1');
  if (first && ss.getSheets().length > 1) ss.deleteSheet(first);
  ss.setActiveSheet(log);
  ss.moveActiveSheet(1);

  CacheService.getScriptCache().remove('config');
  if (added.length) Logger.log('Added to the Buttons tab: ' + added.join(', '));
  Logger.log('Desk Stats workbook: ' + ss.getUrl());
  return ss.getUrl();
}

/**
 * Starter kinds of help that the Buttons tab does not have yet go in at the end of its list, so a kind added to
 * the code reaches an existing workbook by running setup again. A kind already there, under any capitalisation,
 * is left exactly as it is. Returns the names it added.
 */
function addMissingButtons_(sh) {
  const last = sh.getLastRow();
  const rows = last < 2 ? [] : sh.getRange(2, 1, last - 1, 5).getValues();
  const have = {};
  let order = 0;
  rows.forEach(r => {
    const name = String(r[0]).trim().toLowerCase();
    if (name) have[name] = true;
    order = Math.max(order, Number(r[2]) || 0);
  });
  const missing = APP.categories.filter(c => !have[c[0].toLowerCase()]);
  if (missing.length) sh.getRange(last + 1, 1, missing.length, 5).setValues(missing.map((c, i) => [c[0], c[1], order + i + 1, 'yes', '']));
  return missing.map(c => c[0]);
}

/** The Buttons tab as a list, in its Order: every kind of help, shown or not, with its helper text and campus. */
function buttonRows_(ss) {
  return readRows_(ss, SHEETS.buttons)
    .filter(r => String(r[0]).trim())
    .sort((a, b) => Number(a[2]) - Number(b[2]))
    .map(r => ({ category: String(r[0]).trim(), help: String(r[1]), campus: String(r[4] || '').trim(), show: String(r[3]).trim().toLowerCase() !== 'no' }));
}

function ensureSheet_(ss, spec) {
  let sh = ss.getSheetByName(spec.name);
  if (!sh) sh = ss.insertSheet(spec.name);
  sh.getRange(1, 1, 1, spec.headers.length).setValues([spec.headers]).setFontWeight('bold');
  sh.setFrozenRows(1);
  return sh;
}

/** One tab per campus that reproduces the paper weekly sheet with live formulas. */
function buildWeeklySheet_(ss, campus) {
  let sh = ss.getSheetByName(campus.name);
  if (!sh) sh = ss.insertSheet(campus.name);
  sh.clear();
  const q = s => '"' + String(s).replace(/"/g, '""') + '"';
  const dayCols = ['C', 'D', 'E', 'F', 'G', 'H', 'I'];
  const blank = () => ['', '', '', '', '', '', '', '', '', ''];
  const rows = [];
  rows.push([campus.name + ': weekly sheet', '', 'Type any date in the week you want to see into the yellow cell (B2).', '', '', '', '', '', '', '']);
  rows.push(['Any date in the week', '=TODAY()', '', '', '', '', '', '', '', '']);
  rows.push(['Week starting Sunday', '=B2-WEEKDAY(B2)+1', '', '', '', '', '', '', '', '']);
  rows.push(['', '', '=$B$3', '=$B$3+1', '=$B$3+2', '=$B$3+3', '=$B$3+4', '=$B$3+5', '=$B$3+6', '']);
  rows.push(['Category', 'Mode', 'Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'TOTAL']);
  let r = 6;
  const kinds = buttonRows_(ss).filter(b => !b.campus || b.campus === campus.name).map(b => b.category);
  kinds.forEach(c => {
    APP.modes.forEach(m => {
      const row = [c, m];
      dayCols.forEach(col => row.push(
        '=SUMIFS(Log!$H:$H,Log!$E:$E,' + q(campus.name) + ',Log!$F:$F,$A' + r + ',Log!$G:$G,$B' + r + ',Log!$B:$B,' + col + '$4)'));
      row.push('=SUM(C' + r + ':I' + r + ')');
      rows.push(row);
      r += 1;
    });
  });
  const inPerson = ['TOTAL (In person questions only)', ''];
  dayCols.forEach(col => inPerson.push(
    '=SUMIFS(Log!$H:$H,Log!$E:$E,' + q(campus.name) + ',Log!$G:$G,"In person",Log!$B:$B,' + col + '$4)'));
  inPerson.push('=SUM(C' + r + ':I' + r + ')');
  rows.push(inPerson);
  r += 1;
  const all = ['TOTAL (all questions)', ''];
  dayCols.forEach(col => all.push(
    '=SUMIFS(Log!$H:$H,Log!$E:$E,' + q(campus.name) + ',Log!$B:$B,' + col + '$4)'));
  all.push('=SUM(C' + r + ':I' + r + ')');
  rows.push(all);
  r += 1;
  rows.push(blank());
  r += 1;
  const gateStart = r;
  campus.blocks.forEach(b => {
    const row = ['Gate Count', b];
    dayCols.forEach(col => row.push(
      '=SUMIFS(Gate!$D:$D,Gate!$B:$B,' + q(campus.name) + ',Gate!$C:$C,$B' + r + ',Gate!$A:$A,' + col + '$4)'));
    row.push('=SUM(C' + r + ':I' + r + ')');
    rows.push(row);
    r += 1;
  });
  const gateTotal = ['Gate Count', 'GC Total'];
  dayCols.forEach(col => gateTotal.push('=SUM(' + col + gateStart + ':' + col + (r - 1) + ')'));
  gateTotal.push('=SUM(C' + r + ':I' + r + ')');
  rows.push(gateTotal);

  sh.getRange(1, 1, rows.length, 10).setValues(rows);
  sh.getRange('A1').setFontWeight('bold').setFontSize(14);
  sh.getRange('B2').setBackground('#ffe100').setNumberFormat('yyyy-mm-dd');
  sh.getRange('B3').setNumberFormat('yyyy-mm-dd');
  sh.getRange('C4:I4').setNumberFormat('mmm d').setFontColor('#494949');
  sh.getRange('A5:J5').setFontWeight('bold').setBackground('#212721').setFontColor('#ffffff');
  sh.getRange(gateStart - 3, 1, 2, 10).setFontWeight('bold');
  sh.getRange(r, 1, 1, 10).setFontWeight('bold');
  sh.setFrozenRows(5);
  sh.setColumnWidth(1, 250);
  sh.setColumnWidth(2, 130);
}

/** Month-at-a-glance for the director: by campus, by category, busiest hours. */
function buildSummarySheet_(ss) {
  let sh = ss.getSheetByName('Summary');
  if (!sh) sh = ss.insertSheet('Summary');
  sh.clear();
  const campuses = readCampuses_(ss).map(c => c.name);
  const range = (col, sheet) => sheet + '!$' + col + ':$' + col;
  const rows = [];
  rows.push(['Summary', '', 'Type any date in the month you want to see into the yellow cell (B2).', '', '', '']);
  rows.push(['Any date in the month', '=TODAY()', '', '', '', '']);
  rows.push(['Month starts', '=DATE(YEAR(B2),MONTH(B2),1)', '', '', '', '']);
  rows.push(['Month ends', '=EOMONTH(B2,0)', '', '', '', '']);
  rows.push(['', '', '', '', '', '']);
  rows.push(['By campus', 'In person', 'Remote', 'All questions', 'Gate total', '']);
  let r = 7;
  const inMonth = (dateCol, sheet) => ',' + range(dateCol, sheet) + ',">="&$B$3,' + range(dateCol, sheet) + ',"<="&$B$4';
  campuses.forEach(c => {
    rows.push([
      c,
      '=SUMIFS(Log!$H:$H,Log!$E:$E,$A' + r + ',Log!$G:$G,B$6' + inMonth('B', 'Log') + ')',
      '=SUMIFS(Log!$H:$H,Log!$E:$E,$A' + r + ',Log!$G:$G,C$6' + inMonth('B', 'Log') + ')',
      '=SUMIFS(Log!$H:$H,Log!$E:$E,$A' + r + inMonth('B', 'Log') + ')',
      '=SUMIFS(Gate!$D:$D,Gate!$B:$B,$A' + r + inMonth('A', 'Gate') + ')',
      '',
    ]);
    r += 1;
  });
  rows.push(['All campuses', '=SUM(B7:B' + (r - 1) + ')', '=SUM(C7:C' + (r - 1) + ')', '=SUM(D7:D' + (r - 1) + ')', '=SUM(E7:E' + (r - 1) + ')', '']);
  r += 1;
  rows.push(['', '', '', '', '', '']);
  r += 1;
  const catHeader = r;
  rows.push(['By category (all campuses)', 'In person', 'Remote', 'Total', '', '']);
  r += 1;
  buttonRows_(ss).forEach(b => {
    rows.push([
      b.category,
      '=SUMIFS(Log!$H:$H,Log!$F:$F,$A' + r + ',Log!$G:$G,B$' + catHeader + inMonth('B', 'Log') + ')',
      '=SUMIFS(Log!$H:$H,Log!$F:$F,$A' + r + ',Log!$G:$G,C$' + catHeader + inMonth('B', 'Log') + ')',
      '=B' + r + '+C' + r,
      '', '',
    ]);
    r += 1;
  });
  rows.push(['', '', '', '', '', '']);
  r += 1;
  const hourHeader = r;
  rows.push(['Busiest hours (taps only; typed-in history carries no hour)'].concat(campuses).concat(['']).slice(0, 6));
  r += 1;
  for (let h = 7; h <= 21; h += 1) {
    const row = [h];
    campuses.forEach((c, i) => {
      const col = String.fromCharCode(66 + i);
      row.push('=SUMIFS(Log!$H:$H,Log!$E:$E,' + col + '$' + hourHeader + ',Log!$C:$C,$A' + r + inMonth('B', 'Log') + ')');
    });
    while (row.length < 6) row.push('');
    rows.push(row);
    r += 1;
  }
  sh.getRange(1, 1, rows.length, 6).setValues(rows);
  sh.getRange('A1').setFontWeight('bold').setFontSize(14);
  sh.getRange('B2').setBackground('#ffe100').setNumberFormat('yyyy-mm-dd');
  sh.getRange('B3:B4').setNumberFormat('yyyy-mm-dd');
  [6, catHeader, hourHeader].forEach(h => sh.getRange(h, 1, 1, 6).setFontWeight('bold').setBackground('#212721').setFontColor('#ffffff'));
  sh.getRange(hourHeader + 1, 1, 15, 1).setNumberFormat('0":00"');
  sh.setColumnWidth(1, 300);
}

/* ------------------------------------------------------------------ */
/* Serving the two pages                                                */
/* ------------------------------------------------------------------ */

function doGet(e) {
  const view = (e && e.parameter && e.parameter.view) || 'tap';
  const file = (view === 'board' || view === 'week') ? 'Board' : ((view === 'dashboard' || view === 'trends') ? 'Dashboard' : 'Index');
  const t = HtmlService.createTemplateFromFile(file);
  let config;
  try {
    config = getConfig();
  } catch (err) {
    config = { error: String(err && err.message ? err.message : err) };
  }
  t.configJson = JSON.stringify(config);
  return t.evaluate()
    .setTitle(APP.name)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ------------------------------------------------------------------ */
/* Reading                                                              */
/* ------------------------------------------------------------------ */

function getWorkbook_() {
  const id = PropertiesService.getScriptProperties().getProperty('WORKBOOK_ID');
  if (!id) throw new Error('Desk Stats has not been set up yet. In the script editor, run the function named setup once.');
  return SpreadsheetApp.openById(id);
}

function readRows_(ss, spec) {
  const sh = ss.getSheetByName(spec.name);
  if (!sh) throw new Error('The tab named ' + spec.name + ' is missing from the workbook. Run setup again to put it back.');
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, spec.headers.length).getValues();
}

function readCampuses_(ss) {
  return readRows_(ss, SHEETS.campuses)
    .filter(r => String(r[0]).trim())
    .map(r => ({ name: String(r[0]).trim(), blocks: r.slice(1, 5).map(v => String(v).trim()).filter(Boolean) }));
}

/** Everything a page needs to draw itself. Cached for a minute so taps stay quick. */
function getConfig() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('config');
  if (hit) {
    const parsed = JSON.parse(hit);
    parsed.serverNow = Date.now();
    parsed.today = todayKey_(new Date());
    parsed.webAppUrl = webAppUrl_();
    parsed.build = APP.build;
    return parsed;
  }
  const ss = getWorkbook_();
  const buttons = buttonRows_(ss).filter(b => b.show).map(b => ({ category: b.category, help: b.help, campus: b.campus }));
  const config = {
    app: APP.name,
    campuses: readCampuses_(ss),
    buttons: buttons,
    modes: APP.modes,
    referral: APP.referral,
    timeZone: APP.timeZone,
    sheetUrl: ss.getUrl(),
    serverNow: Date.now(),
    today: todayKey_(new Date()),
  };
  cache.put('config', JSON.stringify(config), 60);
  config.webAppUrl = webAppUrl_();
  config.build = APP.build;
  return config;
}

/** The page's own full address. Links inside the page must carry it, because Google
 *  shows the page inside a frame whose hidden address is not the one in the browser bar. */
function webAppUrl_() {
  try { return ScriptApp.getService().getUrl() || ''; } catch (err) { return ''; }
}

/** Today's tallies for one campus, so a fresh device shows the right counts. */
function getTodayCounts(campusName) {
  const ss = getWorkbook_();
  const today = todayKey_(new Date());
  const taps = {};
  const tapIds = [];
  logRows_(ss).forEach(row => {
    if (row.date !== today || row.campus !== campusName) return;
    const key = row.category + '|' + row.mode;
    taps[key] = (taps[key] || 0) + row.count;
    if (row.tapId) tapIds.push(row.tapId);
  });
  const gate = {};
  gateRows_(ss).forEach(row => {
    if (row.date !== today || row.campus !== campusName) return;
    gate[row.block] = row.count;
  });
  return { date: today, campus: campusName, taps: taps, tapIds: tapIds, gate: gate, serverNow: Date.now(), build: APP.build };
}

/**
 * Everything the This week page needs, in one call: every campus-day since the first recorded one
 * (questions per kind and mode, gate count per block), so the page can show any week at once without
 * asking again, plus today, the first and current week, and the campuses, kinds and modes.
 */
function getBoard() {
  const ss = getWorkbook_();
  const config = getConfig();
  const grouped = groupDaily_(logRows_(ss), gateRows_(ss));
  const today = todayKey_(new Date());
  const thisWeek = addDaysKey_(today, -weekdayIndex_(today));
  const isKey = s => /^\d{4}-\d{2}-\d{2}$/.test(String(s || ''));
  let first = '';
  grouped.daily.forEach(d => { if (isKey(d.d) && (!first || d.d < first)) first = d.d; });
  const firstWeek = first ? addDaysKey_(first, -weekdayIndex_(first)) : thisWeek;
  return {
    daily: grouped.daily, today: today, thisWeek: thisWeek, firstWeek: firstWeek,
    campuses: config.campuses, categories: config.buttons.map(b => b.category), modes: APP.modes, serverNow: Date.now(),
  };
}

function logRows_(ss) {
  const sh = ss.getSheetByName(SHEETS.log.name);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 2, last - 1, 10).getValues().map(v => ({
    date: keyOf_(v[0]),
    hour: (v[1] === '' || v[1] === null) ? null : Number(v[1]),
    campus: String(v[3]).trim(), category: String(v[4]).trim(), mode: String(v[5]).trim(), count: Number(v[6]) || 0,
    source: String(v[8] || '').trim(), tapId: String(v[9] || '').trim(),
  }));
}

/* ------------------------------------------------------------------ */
/* The dashboard: one call returns every day's totals; the page does   */
/* the rest, so filters change instantly without another round trip.  */
/* ------------------------------------------------------------------ */

function getDashboard() {
  const ss = getWorkbook_();
  const config = getConfig();
  const grouped = groupDaily_(logRows_(ss), gateRows_(ss));
  const seen = {};
  grouped.daily.forEach(d => {
    Object.keys(d.ip).forEach(c => { seen[c] = true; });
    Object.keys(d.rm).forEach(c => { seen[c] = true; });
  });
  const categories = config.buttons.map(b => b.category);
  Object.keys(seen).forEach(c => { if (categories.indexOf(c) < 0) categories.push(c); });
  const blocks = [];
  config.campuses.forEach(c => c.blocks.forEach(b => { if (blocks.indexOf(b) < 0) blocks.push(b); }));
  return {
    daily: grouped.daily,
    hours: grouped.hours,
    campuses: config.campuses.map(c => c.name),
    categories: categories,
    blocks: blocks,
    today: todayKey_(new Date()),
    serverNow: Date.now(),
  };
}

/** Pure: folds row lists into one record per campus per day. Tested outside Google too. */
function groupDaily_(logs, gates) {
  const days = {};
  const hours = {};
  const dayOf = (date, campus) => {
    const k = date + '|' + campus;
    if (!days[k]) days[k] = { d: date, c: campus, ip: {}, rm: {}, g: {} };
    return days[k];
  };
  logs.forEach(row => {
    if (!row.date || !row.campus || !row.category) return;
    const d = dayOf(row.date, row.campus);
    const bucket = row.mode === 'Remote' ? d.rm : d.ip;
    bucket[row.category] = (bucket[row.category] || 0) + row.count;
    if (row.hour !== null && row.hour !== undefined && row.hour >= 0 && row.hour < 24) {
      if (!hours[row.campus]) hours[row.campus] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0];
      hours[row.campus][Math.floor(row.hour)] += row.count;
    }
  });
  gates.forEach(row => {
    if (!row.date || !row.campus || !row.block) return;
    const d = dayOf(row.date, row.campus);
    d.g[row.block] = (d.g[row.block] || 0) + row.count;
  });
  const daily = Object.keys(days).sort().map(k => days[k]);
  return { daily: daily, hours: hours };
}

/* ------------------------------------------------------------------ */
/* One-off: pull the typed-in history (the LimeSurvey years) into the  */
/* sheet. Upload history-log.csv and history-gate.csv to Google Drive, */
/* then run this once. Running it again adds nothing twice.            */
/* ------------------------------------------------------------------ */

function importHistory() {
  const ss = getWorkbook_();
  const lines = [
    importHistoryFile_(ss, 'history-log.csv', SHEETS.log, 10, r => ['', dateFromKey_(r[1]), '', r[3], r[4], r[5], r[6], Number(r[7]) || 0, '', r[9], '']),
    importHistoryFile_(ss, 'history-gate.csv', SHEETS.gate, 6, r => [dateFromKey_(r[0]), r[1], r[2], Number(r[3]) || 0, '', r[5], '']),
  ];
  CacheService.getScriptCache().remove('config');
  Logger.log(lines.join('\n'));
  return lines.join('\n');
}

function importHistoryFile_(ss, fileName, spec, sourceColumn, shape) {
  const files = DriveApp.getFilesByName(fileName);
  if (!files.hasNext()) return fileName + ': not found in your Google Drive. Drag it into drive.google.com, then run importHistory again.';
  const rows = Utilities.parseCsv(files.next().getBlob().getDataAsString('UTF-8'));
  const header = rows.shift() || [];
  if (header.join('|') !== spec.headers.join('|')) return fileName + ': its first line does not match the ' + spec.name + ' tab headings, so nothing was imported from it.';
  const sh = ss.getSheetByName(spec.name);
  const already = {};
  const last = sh.getLastRow();
  if (last >= 2) sh.getRange(2, sourceColumn, last - 1, 1).getValues().forEach(v => { already[String(v[0])] = true; });
  const out = [];
  const skipped = {};
  rows.forEach(r => {
    if (!r.length || !r.join('').trim()) return;
    const source = String(r[sourceColumn - 1]);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(String(r[spec === SHEETS.log ? 1 : 0]))) return;
    if (already[source]) { skipped[source] = true; return; }
    out.push(shape(r));
  });
  if (out.length) sh.getRange(sh.getLastRow() + 1, 1, out.length, spec.headers.length).setValues(out);
  const skippedCount = Object.keys(skipped).length;
  return fileName + ': ' + out.length + ' rows added to the ' + spec.name + ' tab' + (skippedCount ? '; ' + skippedCount + ' weeks skipped because they were already there' : '') + '.';
}

function gateRows_(ss) {
  const sh = ss.getSheetByName(SHEETS.gate.name);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 4).getValues().map(v => ({
    date: keyOf_(v[0]), campus: String(v[1]).trim(), block: String(v[2]).trim(), count: Number(v[3]) || 0,
  }));
}

/* ------------------------------------------------------------------ */
/* Writing                                                              */
/* ------------------------------------------------------------------ */

/** One tap: one row in Log, stamped with Google's clock, never twice for the same tap id. */
function recordTap(tap) {
  return recordTaps([tap])[0];
}

/**
 * Saves a burst of taps in order under one lock and returns one answer per tap, in the same
 * order: {tapId, ok, delta} saved; {ok, duplicate} already there (a resend); {ok:false, refused,
 * reason} a take-off that would push today's number below zero; {ok:false, permanent, reason}
 * a tap the sheet can never accept (unknown campus or button).
 */
function recordTaps(list) {
  const items = Array.isArray(list) ? list.slice(0, 50) : [];
  if (!items.length) return [];
  const config = getConfig();
  const ss = getWorkbook_();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = ss.getSheetByName(SHEETS.log.name);
    const now = new Date();
    const seen = recentIds_(sh, 11, 800);
    const netsByDay = {};
    const netsFor = day => { if (!netsByDay[day]) netsByDay[day] = todayNets_(sh, day); return netsByDay[day]; };
    const rows = [];
    const results = items.map(raw => {
      const t = raw || {};
      const campus = String(t.campus || '').trim();
      const category = String(t.category || '').trim();
      const mode = String(t.mode || '').trim();
      const tapId = String(t.tapId || '').trim();
      const delta = Number(t.delta) === -1 ? -1 : 1;
      if (!tapId) return { tapId: tapId, ok: false, permanent: true, reason: 'A tap arrived without an id. Reload the page.' };
      if (!config.campuses.some(c => c.name === campus)) return { tapId: tapId, ok: false, permanent: true, reason: 'The campus "' + campus + '" is not in the Campuses tab. Reload the page.' };
      if (!config.buttons.some(b => b.category === category)) return { tapId: tapId, ok: false, permanent: true, reason: 'The button "' + category + '" is not in the Buttons tab. Reload the page.' };
      if (APP.modes.indexOf(mode) < 0) return { tapId: tapId, ok: false, permanent: true, reason: 'Unknown mode: ' + mode };
      if (seen[tapId]) return { tapId: tapId, ok: true, duplicate: true, delta: delta };
      const when = tapTime_(t.at, now);
      const day = todayKey_(when);
      const net = netsFor(day);
      const key = campus + '|' + category + '|' + mode;
      if (delta === -1 && (net[key] || 0) <= 0) {
        return { tapId: tapId, ok: false, refused: true, delta: delta, reason: 'Nothing to take off: the count for ' + category + ', ' + mode + ' at ' + campus + ' on ' + day + ' is already 0 in the sheet.' };
      }
      net[key] = (net[key] || 0) + delta;
      seen[tapId] = true;
      rows.push([when, dateOnly_(when), hourOf_(when), weekday_(when), campus, category, mode, delta, whoAmI_(t.device), 'tap', tapId]);
      return { tapId: tapId, ok: true, delta: delta };
    });
    if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, SHEETS.log.headers.length).setValues(rows);
    return results;
  } finally {
    lock.releaseLock();
  }
}

/**
 * The moment of the tap itself, as the device saw it (its clock is aligned to Google's when the page opens), so two
 * taps that travel in one batch still carry their own times. The save time stands in when the moment is missing or
 * implausible: more than a week old, or in the future.
 */
function tapTime_(at, now) {
  const ms = Number(at);
  if (!ms || !isFinite(ms)) return now;
  if (ms > now.getTime() + 60000 || ms < now.getTime() - 7 * 86400000) return now;
  return new Date(ms);
}

/** The ids in one column of the newest rows, so a resend of the same tap or gate entry is recognised and not written twice. */
function recentIds_(sh, column, maxRows) {
  const seen = {};
  const last = sh.getLastRow();
  if (last < 2) return seen;
  const n = Math.min(last - 1, maxRows);
  sh.getRange(last - n + 1, column, n, 1).getValues().forEach(v => {
    const id = String(v[0] || '').trim();
    if (id) seen[id] = true;
  });
  return seen;
}

/** One day's net count (adds minus take-offs, plus corrections) for every campus, kind and mode, in one read. */
function todayNets_(sh, today) {
  const net = {};
  const last = sh.getLastRow();
  if (last < 2) return net;
  sh.getRange(2, 2, last - 1, 7).getValues().forEach(v => {
    if (keyOf_(v[0]) !== today) return;
    const key = String(v[3]).trim() + '|' + String(v[4]).trim() + '|' + String(v[5]).trim();
    net[key] = (net[key] || 0) + (Number(v[6]) || 0);
  });
  return net;
}

/** The last few taps at one campus, newest first: time, who, kind, mode, +1 or -1. */
function getRecent(campus, limit) {
  const ss = getWorkbook_();
  const sh = ss.getSheetByName(SHEETS.log.name);
  const want = Math.min(Math.max(Number(limit) || 10, 1), 50);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const n = Math.min(last - 1, 800);
  const rows = sh.getRange(last - n + 1, 1, n, 11).getValues();
  const out = [];
  for (let i = rows.length - 1; i >= 0 && out.length < want; i -= 1) {
    const v = rows[i];
    if (String(v[9]).trim() !== 'tap') continue;
    if (campus && String(v[4]).trim() !== String(campus).trim()) continue;
    const when = (v[0] && typeof v[0].getTime === 'function') ? v[0] : null;
    out.push({
      when: when ? Utilities.formatDate(when, APP.timeZone, 'EEE MMM d, h:mm:ss a') : String(v[0]),
      who: String(v[8] || ''), category: String(v[5]), mode: String(v[6]), delta: Number(v[7]) || 0, tapId: String(v[10] || ''),
    });
  }
  return out;
}

/**
 * One gate count for a campus, day and time block. The Gate tab keeps one row per block with the
 * current number (entering it again replaces that row, so every total stays right), and every
 * entry, first or replacement, is also added to the Gate history tab with the number it replaced;
 * nothing there is ever overwritten. A resend of an entry already saved changes nothing.
 */
function recordGate(entry) {
  const g = entry || {};
  const campus = String(g.campus || '').trim();
  const block = String(g.block || '').trim();
  const count = Number(g.count);
  const dateKey = String(g.date || todayKey_(new Date())).trim();
  const entryId = String(g.entryId || '').trim();
  const config = getConfig();
  const campusRow = config.campuses.filter(c => c.name === campus)[0];
  if (!campusRow) throw new Error('The campus "' + campus + '" is not in the Campuses tab of the sheet.');
  if (campusRow.blocks.indexOf(block) < 0) throw new Error('The time block "' + block + '" is not listed for ' + campus + ' in the Campuses tab. Reload the page.');
  if (!(count >= 0) || Math.floor(count) !== count) throw new Error('The gate count must be a whole number, zero or more.');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dateKey)) throw new Error('Bad date for the gate count: ' + dateKey);
  if (dateKey > todayKey_(new Date())) throw new Error('That day has not happened yet.');
  const ss = getWorkbook_();
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = ss.getSheetByName(SHEETS.gate.name);
    const history = gateHistorySheet_(ss);
    const now = new Date();
    const who = whoAmI_(g.device);
    const duplicate = entryId ? recentIds_(history, 9, 300)[entryId] === true : false;
    let replaced = false;
    let previous = '';
    if (!duplicate) {
      const last = sh.getLastRow();
      if (last >= 2) {
        const rows = sh.getRange(2, 1, last - 1, 4).getValues();
        for (let i = rows.length - 1; i >= 0; i -= 1) {
          if (keyOf_(rows[i][0]) === dateKey && String(rows[i][1]).trim() === campus && String(rows[i][2]).trim() === block) {
            previous = Number(rows[i][3]) || 0;
            sh.getRange(i + 2, 4, 1, 4).setValues([[count, who, 'tap', now]]);
            replaced = true;
            break;
          }
        }
      }
      if (!replaced) sh.appendRow([dateFromKey_(dateKey), campus, block, count, who, 'tap', now]);
      history.appendRow([now, dateFromKey_(dateKey), campus, block, count, previous, who, 'tap', entryId]);
    }
    const blocks = {};
    let dayTotal = 0;
    gateRows_(ss).forEach(row => {
      if (row.date !== dateKey || row.campus !== campus) return;
      blocks[row.block] = row.count;
      dayTotal += row.count;
    });
    return { ok: true, replaced: replaced, duplicate: duplicate, date: dateKey, blocks: blocks, dayTotal: dayTotal };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Changes one cell of the weekly grid: the questions of one kind and mode at one campus on one day.
 * Either sets it to a number (count) or nudges it up or down by a few (delta, the arrows on the
 * This week page). Nothing already in the sheet is touched: the change is written as one correction
 * row (Count = the difference, Source = edit, no clock hour, who made it), so every total still adds
 * up the Count column and the correction is on record. A nudge that would push the day below zero is
 * refused. A resend of the same edit changes nothing.
 */
function setLogCell(edit) {
  const e = edit || {};
  const campus = String(e.campus || '').trim();
  const category = String(e.category || '').trim();
  const mode = String(e.mode || '').trim();
  const dateKey = String(e.date || '').trim();
  const nudging = e.delta !== undefined && e.delta !== null && e.delta !== '';
  const delta = nudging ? Number(e.delta) : 0;
  const wanted = nudging ? null : Number(e.count);
  const editId = String(e.editId || '').trim();
  const config = getConfig();
  if (!config.campuses.some(c => c.name === campus)) throw new Error('The campus "' + campus + '" is not in the Campuses tab. Reload the page.');
  if (!config.buttons.some(b => b.category === category)) throw new Error('The button "' + category + '" is not in the Buttons tab. Reload the page.');
  if (APP.modes.indexOf(mode) < 0) throw new Error('Unknown mode: ' + mode);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dateKey)) throw new Error('Bad date for the change: ' + dateKey);
  if (dateKey > todayKey_(new Date())) throw new Error('That day has not happened yet.');
  if (nudging) {
    if (Math.floor(delta) !== delta || !delta || Math.abs(delta) > 500) throw new Error('A change by arrows must be a whole number of taps, up to 500 either way.');
  } else if (!(wanted >= 0) || Math.floor(wanted) !== wanted) throw new Error('The number must be a whole number, zero or more.');
  const ss = getWorkbook_();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = ss.getSheetByName(SHEETS.log.name);
    const now = new Date();
    const key = campus + '|' + category + '|' + mode;
    const duplicate = editId ? recentIds_(sh, 11, 800)[editId] === true : false;
    const current = todayNets_(sh, dateKey)[key] || 0;
    let diff = 0;
    if (!duplicate) {
      if (nudging) {
        if (current + delta < 0) throw new Error('Nothing to take off: ' + category + ', ' + mode + ' at ' + campus + ' on ' + dateKey + ' is already 0 in the sheet.');
        diff = delta;
      } else diff = wanted - current;
    }
    if (diff !== 0) {
      sh.appendRow([now, dateFromKey_(dateKey), '', weekday_(dateFromKey_(dateKey)), campus, category, mode, diff, whoAmI_(e.device), 'edit', editId]);
    }
    return { ok: true, duplicate: duplicate, date: dateKey, campus: campus, category: category, mode: mode, count: current + diff, diff: diff };
  } finally {
    lock.releaseLock();
  }
}

/** The Gate history tab, created on the spot in a workbook from before it existed, so nobody has to run setup again. */
function gateHistorySheet_(ss) {
  const existing = ss.getSheetByName(SHEETS.gateHistory.name);
  if (existing) return existing;
  const sh = ensureSheet_(ss, SHEETS.gateHistory);
  sh.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  sh.getRange('B:B').setNumberFormat('yyyy-mm-dd');
  return sh;
}

/* ------------------------------------------------------------------ */
/* Weekly backup copy (optional; run installWeeklyBackup once to arm)   */
/* ------------------------------------------------------------------ */

function installWeeklyBackup() {
  const already = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'backupNow');
  if (already) return 'The weekly backup was already switched on.';
  ScriptApp.newTrigger('backupNow').timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(3).create();
  return 'Weekly backup switched on: every Sunday around 3 am a dated copy of Log, Gate and Gate history lands in the folder "Desk Stats backups".';
}

function backupNow() {
  try {
    const ss = getWorkbook_();
    const folderName = 'Desk Stats backups';
    const folders = DriveApp.getFoldersByName(folderName);
    const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
    const copy = SpreadsheetApp.create(APP.name + ' backup ' + todayKey_(new Date()));
    [SHEETS.log, SHEETS.gate, SHEETS.gateHistory].forEach(spec => {
      const src = ss.getSheetByName(spec.name);
      if (!src) return;
      const values = src.getDataRange().getValues();
      const dst = copy.insertSheet(spec.name);
      if (values.length) dst.getRange(1, 1, values.length, values[0].length).setValues(values);
    });
    const first = copy.getSheetByName('Sheet1');
    if (first) copy.deleteSheet(first);
    DriveApp.getFileById(copy.getId()).moveTo(folder);
    Logger.log('Backup written: ' + copy.getUrl());
  } catch (err) {
    Logger.log('Backup failed: ' + err);
  }
}

/* ------------------------------------------------------------------ */
/* Small helpers                                                        */
/* ------------------------------------------------------------------ */

function findRecentRow_(sh, column, value, maxRows) {
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const start = Math.max(2, last - maxRows + 1);
  const values = sh.getRange(start, column, last - start + 1, 1).getValues();
  for (let i = values.length - 1; i >= 0; i -= 1) {
    if (String(values[i][0]) === value) return start + i;
  }
  return 0;
}

function whoAmI_(device) {
  let email = '';
  try { email = Session.getActiveUser().getEmail() || ''; } catch (err) { email = ''; }
  return email || String(device || '');
}

/* Dates. A "key" is the text yyyy-mm-dd in the app's zone. Nothing here reads the
   project's own clock setting, so the sheet dates stay right whatever that is set to. */
function dateOnly_(d) {
  return dateFromKey_(todayKey_(d));
}
function dateFromKey_(key) {
  /* Midnight of that day in the app's zone, found with the clock formatter alone:
     start at noon UTC of the same calendar day and step an hour at a time until the
     formatter reads that day at hour 00. Works across daylight-saving changes. */
  const k = String(key).trim();
  const p = k.split('-').map(Number);
  const want = k + ' 00';
  let guess = new Date(Date.UTC(p[0], p[1] - 1, p[2], 12));
  for (let i = 0; i < 48; i += 1) {
    const stamp = Utilities.formatDate(guess, APP.timeZone, 'yyyy-MM-dd HH');
    if (stamp === want) return guess;
    guess = new Date(guess.getTime() - 3600000 * (stamp > want ? 1 : -1));
  }
  return guess;
}

/**
 * One-off repair, safe to run again: re-dates every tap row from its own time stamp,
 * re-dates gate rows that an earlier version filed one day early, and removes gate
 * rows that were saved twice for the same campus, day and block (the newest stays).
 */
function repairDates() {
  const ss = getWorkbook_();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const lines = [];
    const log = ss.getSheetByName(SHEETS.log.name);
    const last = log.getLastRow();
    let fixedTaps = 0;
    if (last >= 2) {
      const rows = log.getRange(2, 1, last - 1, 10).getValues();
      const out = rows.map(r => {
        const when = r[0];
        if (String(r[9]).trim() !== 'tap' || !(when && typeof when.getTime === 'function')) return [r[1], r[2], r[3]];
        const key = todayKey_(when);
        const hour = hourOf_(when);
        const day = weekday_(when);
        if (keyOf_(r[1]) !== key || Number(r[2]) !== hour || String(r[3]) !== day) fixedTaps += 1;
        return [dateFromKey_(key), hour, day];
      });
      log.getRange(2, 2, out.length, 3).setValues(out);
    }
    lines.push('Log: ' + fixedTaps + ' tap rows re-dated from their time stamp.');

    const gate = ss.getSheetByName(SHEETS.gate.name);
    const glast = gate.getLastRow();
    let fixedGate = 0;
    let removed = 0;
    if (glast >= 2) {
      const rows = gate.getRange(2, 1, glast - 1, 7).getValues();
      const dates = rows.map(r => {
        const when = r[6];
        if (String(r[5]).trim() !== 'tap' || !(when && typeof when.getTime === 'function')) return [r[0]];
        const whenKey = todayKey_(when);
        if (keyOf_(r[0]) === addDaysKey_(whenKey, -1)) { fixedGate += 1; return [dateFromKey_(whenKey)]; }
        return [r[0]];
      });
      gate.getRange(2, 1, dates.length, 1).setValues(dates);
      const newest = {};
      rows.forEach((r, i) => {
        const key = keyOf_(dates[i][0]) + '|' + String(r[1]).trim() + '|' + String(r[2]).trim();
        const stamp = (r[6] && typeof r[6].getTime === 'function') ? r[6].getTime() : 0;
        if (!newest[key] || stamp >= newest[key].stamp) newest[key] = { row: i + 2, stamp: stamp };
      });
      const keep = {};
      Object.keys(newest).forEach(k => { keep[newest[k].row] = true; });
      for (let i = rows.length - 1; i >= 0; i -= 1) {
        const rowNumber = i + 2;
        if (!keep[rowNumber]) { gate.deleteRow(rowNumber); removed += 1; }
      }
    }
    lines.push('Gate: ' + fixedGate + ' rows moved to the right day; ' + removed + ' duplicate rows removed.');
    CacheService.getScriptCache().remove('config');
    Logger.log(lines.join('\n'));
    return lines.join('\n');
  } finally {
    lock.releaseLock();
  }
}
function hourOf_(d) {
  return Number(Utilities.formatDate(d, APP.timeZone, 'H'));
}
function weekday_(d) {
  return Utilities.formatDate(d, APP.timeZone, 'EEEE');
}
function weekdayIndex_(key) {
  const p = String(key).split('-').map(Number);
  return new Date(Date.UTC(p[0], p[1] - 1, p[2])).getUTCDay();
}
function addDaysKey_(key, n) {
  const p = String(key).split('-').map(Number);
  const d = new Date(Date.UTC(p[0], p[1] - 1, p[2] + n));
  const mm = d.getUTCMonth() + 1, dd = d.getUTCDate();
  return d.getUTCFullYear() + '-' + (mm < 10 ? '0' : '') + mm + '-' + (dd < 10 ? '0' : '') + dd;
}

function todayKey_(d) {
  return Utilities.formatDate(d, APP.timeZone, 'yyyy-MM-dd');
}

/* Google's date formatter is slow and the tabs hold thousands of date cells but only a few hundred distinct days:
   each day's key is worked out once per call and remembered. Exact, because the memory is keyed on the cell's own instant. */
const KEY_CACHE_ = {};
function keyOf_(cell) {
  if (cell && typeof cell.getTime === 'function') {
    const ms = cell.getTime();
    if (!KEY_CACHE_[ms]) KEY_CACHE_[ms] = Utilities.formatDate(cell, APP.timeZone, 'yyyy-MM-dd');
    return KEY_CACHE_[ms];
  }
  const s = String(cell).trim();
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return m[3] + '-' + ('0' + m[1]).slice(-2) + '-' + ('0' + m[2]).slice(-2);
  return s;
}
