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
  timeZone: 'America/Chicago',
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
  ],
  referral: 'Any interaction requiring in-depth research or more than 20 minutes of time: please refer the patron to a subject specialist or librarian.',
};

const SHEETS = {
  log: { name: 'Log', headers: ['When', 'Date', 'Hour', 'Weekday', 'Campus', 'Category', 'Mode', 'Count', 'Recorded by', 'Source', 'Tap id'] },
  gate: { name: 'Gate', headers: ['Date', 'Campus', 'Block', 'Count', 'Recorded by', 'Source', 'When entered'] },
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

  const buttons = ensureSheet_(ss, SHEETS.buttons);
  if (buttons.getLastRow() < 2) {
    buttons.getRange(2, 1, APP.categories.length, 5)
      .setValues(APP.categories.map((c, i) => [c[0], c[1], i + 1, 'yes', '']));
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
  Logger.log('Desk Stats workbook: ' + ss.getUrl());
  return ss.getUrl();
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
  APP.categories.forEach(c => {
    APP.modes.forEach(m => {
      const row = [c[0], m];
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
  APP.categories.forEach(c => {
    rows.push([
      c[0],
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
    return parsed;
  }
  const ss = getWorkbook_();
  const buttons = readRows_(ss, SHEETS.buttons)
    .filter(r => String(r[0]).trim() && String(r[3]).trim().toLowerCase() !== 'no')
    .sort((a, b) => Number(a[2]) - Number(b[2]))
    .map(r => ({ category: String(r[0]).trim(), help: String(r[1]), campus: String(r[4] || '').trim() }));
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
  logRows_(ss).forEach(row => {
    if (row.date !== today || row.campus !== campusName) return;
    const key = row.category + '|' + row.mode;
    taps[key] = (taps[key] || 0) + row.count;
  });
  const gate = {};
  gateRows_(ss).forEach(row => {
    if (row.date !== today || row.campus !== campusName) return;
    gate[row.block] = row.count;
  });
  return { date: today, campus: campusName, taps: taps, gate: gate, serverNow: Date.now() };
}

/** The current week's paper grid for one campus or for all of them. */
function getBoard(campusName) {
  const ss = getWorkbook_();
  const config = getConfig();
  const now = new Date();
  const today = todayKey_(now);
  const weekStart = dateOnly_(now);
  weekStart.setDate(weekStart.getDate() - weekStart.getDay());
  const days = [];
  for (let i = 0; i < 7; i += 1) {
    const d = new Date(weekStart.getTime());
    d.setDate(d.getDate() + i);
    days.push(todayKey_(d));
  }
  const wanted = config.campuses.filter(c => !campusName || campusName === 'all' || c.name === campusName);
  const logs = logRows_(ss);
  const gates = gateRows_(ss);
  const boards = wanted.map(c => {
    const cells = {};
    const inPerson = [0, 0, 0, 0, 0, 0, 0];
    const all = [0, 0, 0, 0, 0, 0, 0];
    let todayTaps = 0;
    logs.forEach(row => {
      if (row.campus !== c.name) return;
      const idx = days.indexOf(row.date);
      if (idx < 0) return;
      const key = row.category + '|' + row.mode;
      if (!cells[key]) cells[key] = [0, 0, 0, 0, 0, 0, 0];
      cells[key][idx] += row.count;
      all[idx] += row.count;
      if (row.mode === 'In person') inPerson[idx] += row.count;
      if (row.date === today) todayTaps += row.count;
    });
    const gate = {};
    c.blocks.forEach(b => { gate[b] = [0, 0, 0, 0, 0, 0, 0]; });
    const gateTotals = [0, 0, 0, 0, 0, 0, 0];
    gates.forEach(row => {
      if (row.campus !== c.name) return;
      const idx = days.indexOf(row.date);
      if (idx < 0) return;
      if (!gate[row.block]) gate[row.block] = [0, 0, 0, 0, 0, 0, 0];
      gate[row.block][idx] += row.count;
      gateTotals[idx] += row.count;
    });
    return {
      campus: c.name, blocks: c.blocks, cells: cells, inPerson: inPerson, all: all,
      gate: gate, gateTotals: gateTotals, todayTaps: todayTaps, todayGate: gateTotals[days.indexOf(today)],
    };
  });
  return { days: days, today: today, weekStart: days[0], boards: boards, categories: config.buttons.map(b => b.category), modes: APP.modes, serverNow: Date.now() };
}

function logRows_(ss) {
  const sh = ss.getSheetByName(SHEETS.log.name);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 2, last - 1, 7).getValues().map(v => ({
    date: keyOf_(v[0]),
    hour: (v[1] === '' || v[1] === null) ? null : Number(v[1]),
    campus: String(v[3]).trim(), category: String(v[4]).trim(), mode: String(v[5]).trim(), count: Number(v[6]) || 0,
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
    importHistoryFile_(ss, 'history-gate.csv', SHEETS.gate, 5, r => [dateFromKey_(r[0]), r[1], r[2], Number(r[3]) || 0, '', r[5], '']),
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
  const t = tap || {};
  const campus = String(t.campus || '').trim();
  const category = String(t.category || '').trim();
  const mode = String(t.mode || '').trim();
  const tapId = String(t.tapId || '').trim();
  const delta = Number(t.delta) === -1 ? -1 : 1;
  const config = getConfig();
  if (!config.campuses.some(c => c.name === campus)) throw new Error('The campus "' + campus + '" is not in the Campuses tab. Reload the page.');
  if (!config.buttons.some(b => b.category === category)) throw new Error('The button "' + category + '" is not in the Buttons tab. Reload the page.');
  if (APP.modes.indexOf(mode) < 0) throw new Error('Unknown mode: ' + mode);
  if (!tapId) throw new Error('The tap had no id; reload the page.');
  const ss = getWorkbook_();
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = ss.getSheetByName(SHEETS.log.name);
    if (findRecentRow_(sh, 11, tapId, 400) > 0) return { ok: true, duplicate: true, tapId: tapId, delta: delta };
    const now = new Date();
    if (delta === -1) {
      const net = todayNet_(sh, todayKey_(now), campus, category, mode);
      if (net <= 0) return { ok: false, refused: true, tapId: tapId, delta: delta, net: 0, reason: 'Nothing to take off: today\'s count for ' + category + ', ' + mode + ' at ' + campus + ' is already 0 in the sheet.' };
    }
    sh.appendRow([now, dateOnly_(now), now.getHours(), weekday_(now), campus, category, mode, delta, whoAmI_(t.device), 'tap', tapId]);
    return { ok: true, tapId: tapId, delta: delta, when: now.toISOString() };
  } finally {
    lock.releaseLock();
  }
}

/** Today's net count (adds minus subtractions) for one campus, kind and mode. */
function todayNet_(sh, today, campus, category, mode) {
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const rows = sh.getRange(2, 2, last - 1, 7).getValues();
  let net = 0;
  for (let i = rows.length - 1; i >= 0; i -= 1) {
    const v = rows[i];
    if (keyOf_(v[0]) !== today) continue;
    if (String(v[3]).trim() === campus && String(v[4]).trim() === category && String(v[5]).trim() === mode) net += Number(v[6]) || 0;
  }
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
    const when = v[0] instanceof Date ? v[0] : null;
    out.push({
      when: when ? Utilities.formatDate(when, APP.timeZone, 'EEE MMM d, h:mm:ss a') : String(v[0]),
      who: String(v[8] || ''), category: String(v[5]), mode: String(v[6]), delta: Number(v[7]) || 0, tapId: String(v[10] || ''),
    });
  }
  return out;
}

function recordGate(entry) {
  const g = entry || {};
  const campus = String(g.campus || '').trim();
  const block = String(g.block || '').trim();
  const count = Number(g.count);
  const dateKey = String(g.date || todayKey_(new Date())).trim();
  const config = getConfig();
  const campusRow = config.campuses.filter(c => c.name === campus)[0];
  if (!campusRow) throw new Error('The campus "' + campus + '" is not in the Campuses tab of the sheet.');
  if (campusRow.blocks.indexOf(block) < 0) throw new Error('The time block "' + block + '" is not listed for ' + campus + ' in the Campuses tab. Reload the page.');
  if (!(count >= 0) || Math.floor(count) !== count) throw new Error('The gate count must be a whole number, zero or more.');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dateKey)) throw new Error('Bad date for the gate count: ' + dateKey);
  const ss = getWorkbook_();
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = ss.getSheetByName(SHEETS.gate.name);
    const now = new Date();
    const who = whoAmI_(g.device);
    const last = sh.getLastRow();
    let replaced = false;
    if (last >= 2) {
      const keys = sh.getRange(2, 1, last - 1, 3).getValues();
      for (let i = keys.length - 1; i >= 0; i -= 1) {
        if (keyOf_(keys[i][0]) === dateKey && String(keys[i][1]).trim() === campus && String(keys[i][2]).trim() === block) {
          sh.getRange(i + 2, 4, 1, 4).setValues([[count, who, 'tap', now]]);
          replaced = true;
          break;
        }
      }
    }
    if (!replaced) sh.appendRow([dateFromKey_(dateKey), campus, block, count, who, 'tap', now]);
    const blocks = {};
    let dayTotal = 0;
    gateRows_(ss).forEach(row => {
      if (row.date !== dateKey || row.campus !== campus) return;
      blocks[row.block] = row.count;
      dayTotal += row.count;
    });
    return { ok: true, replaced: replaced, date: dateKey, blocks: blocks, dayTotal: dayTotal };
  } finally {
    lock.releaseLock();
  }
}

/* ------------------------------------------------------------------ */
/* Weekly backup copy (optional; run installWeeklyBackup once to arm)   */
/* ------------------------------------------------------------------ */

function installWeeklyBackup() {
  const already = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'backupNow');
  if (already) return 'The weekly backup was already switched on.';
  ScriptApp.newTrigger('backupNow').timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(3).create();
  return 'Weekly backup switched on: every Sunday around 3 am a dated copy of Log and Gate lands in the folder "Desk Stats backups".';
}

function backupNow() {
  try {
    const ss = getWorkbook_();
    const folderName = 'Desk Stats backups';
    const folders = DriveApp.getFoldersByName(folderName);
    const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
    const copy = SpreadsheetApp.create(APP.name + ' backup ' + todayKey_(new Date()));
    [SHEETS.log, SHEETS.gate].forEach(spec => {
      const src = ss.getSheetByName(spec.name);
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

function dateOnly_(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function dateFromKey_(key) {
  const parts = key.split('-').map(Number);
  return new Date(parts[0], parts[1] - 1, parts[2]);
}

function todayKey_(d) {
  return Utilities.formatDate(d, APP.timeZone, 'yyyy-MM-dd');
}

function keyOf_(cell) {
  if (cell instanceof Date) return Utilities.formatDate(cell, APP.timeZone, 'yyyy-MM-dd');
  const s = String(cell).trim();
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return m[3] + '-' + ('0' + m[1]).slice(-2) + '-' + ('0' + m[2]).slice(-2);
  return s;
}
