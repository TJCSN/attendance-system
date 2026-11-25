/**
 * Seahawk training backend (Google Apps Script)
 *
 * Tabs (created automatically):
 *   Availability - one row per member per week (upserted)
 *   Attendance   - one row per training date (upserted)
 *   Config       - key/value settings shared by all pages
 *
 * Roster     - (you maintain) Name, Category, PaddlingSide, Gender, Active
 *
 * GET  -> { success, data: [availability], attendance: [...], config: {...}, roster: [...] }
 * POST -> JSON body with "action":
 *           "upsert"     (default) member availability
 *           "attendance" actual attendance for one training date
 *           "config"     { key, value } shared setting
 */

var HEADERS = {
  Availability: ['Timestamp', 'Name', 'Category', 'PaddlingSide', 'Gender', 'Days', 'WeekStart', 'WeekNumber'],
  Attendance: ['Date', 'DayName', 'Display', 'Attended', 'AttendedDetails', 'Timestamp'],
  Config: ['Key', 'Value', 'Updated']
};

// ---------- helpers ----------

function getSheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS[name]);
    sh.setFrozenRows(1);
    // Keep date-like columns as plain text so Sheets never converts them
    if (name === 'Availability') sh.getRange('G:G').setNumberFormat('@');
    if (name === 'Attendance') sh.getRange('A:A').setNumberFormat('@');
  }
  return sh;
}

function str_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(v === null || v === undefined ? '' : v).trim();
}

function parse_(v, fallback) {
  try { return JSON.parse(v || ''); } catch (e) { return fallback; }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Find the 1-based row number matching a predicate on column values, or -1
function findRow_(sh, matchFn) {
  var last = sh.getLastRow();
  if (last < 2) return -1;
  var values = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
  for (var i = 0; i < values.length; i++) {
    if (matchFn(values[i])) return i + 2;
  }
  return -1;
}

function writeRow_(sh, row, values) {
  if (row === -1) {
    sh.appendRow(values);
  } else {
    sh.getRange(row, 1, 1, values.length).setValues([values]);
  }
}

// ---------- speed: short-lived cache + recent weeks only ----------

var CACHE_SECONDS = 20;   // GET results are reused for 20s; any save clears them at once
var RECENT_DAYS = 14;     // normal GET returns only the last 2 weeks of availability (pages show the current week); ?all=1 returns everything

function recentCutoff_() {
  var d = new Date();
  d.setDate(d.getDate() - RECENT_DAYS);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function cacheGet_(key) {
  var cache = CacheService.getScriptCache();
  var n = Number(cache.get(key + ':n') || 0);
  if (!n) return null;
  var keys = [];
  for (var i = 0; i < n; i++) keys.push(key + ':' + i);
  var parts = cache.getAll(keys);
  var out = '';
  for (var j = 0; j < keys.length; j++) {
    if (parts[keys[j]] === undefined) return null;
    out += parts[keys[j]];
  }
  return out;
}

function cachePut_(key, str) {
  if (str.length > 300000) return;          // too big to cache, just skip
  var size = 30000;                         // stays under the 100KB-per-value limit even with non-English names
  var n = Math.ceil(str.length / size);
  var obj = {};
  for (var i = 0; i < n; i++) obj[key + ':' + i] = str.substr(i * size, size);
  obj[key + ':n'] = String(n);
  CacheService.getScriptCache().putAll(obj, CACHE_SECONDS);
}

function cacheClear_() {
  CacheService.getScriptCache().removeAll(['get:recent:n', 'get:all:n']);
}

// ---------- readers ----------

function getAvailability_(all) {
  var sh = getSheet_('Availability');
  var last = sh.getLastRow();
  if (last < 2) return [];
  var rows = sh.getRange(2, 1, last - 1, HEADERS.Availability.length).getValues();
  // Same person + same week: keep only the latest submission
  var latest = {};
  var order = [];
  var cutoff = recentCutoff_();
  rows.forEach(function (r) {
    if (!r[1]) return;
    var wk = str_(r[6]);
    if (!all && /^\d{4}-\d{2}-\d{2}$/.test(wk) && wk < cutoff) return;   // skip old weeks unless ?all=1
    var item = {
      timestamp: str_(r[0]),
      name: str_(r[1]),
      category: str_(r[2]),
      paddlingSide: str_(r[3]),
      gender: str_(r[4]),
      days: parse_(r[5], []),
      weekStart: str_(r[6]),
      weekNumber: r[7]
    };
    var key = item.name.toLowerCase() + '|' + item.weekStart;
    if (!latest[key]) order.push(key);
    if (!latest[key] || item.timestamp >= latest[key].timestamp) latest[key] = item;
  });
  return order.map(function (k) { return latest[k]; });
}

function getAttendance_() {
  var sh = getSheet_('Attendance');
  var last = sh.getLastRow();
  if (last < 2) return [];
  var rows = sh.getRange(2, 1, last - 1, HEADERS.Attendance.length).getValues();
  var out = [];
  rows.forEach(function (r) {
    if (!r[0]) return;
    out.push({
      date: str_(r[0]),
      dayName: str_(r[1]),
      display: str_(r[2]),
      attended: parse_(r[3], []),
      attendedDetails: parse_(r[4], []),
      timestamp: str_(r[5])
    });
  });
  return out;
}

// Reads the "Roster" tab you maintain. Column order doesn't matter, only the headings.
// Active: blank / Yes / Y / TRUE / 1 = active;  No / N / FALSE / 0 = hidden from the dropdown.
function cap_(v) {
  var s = String(v === null || v === undefined ? '' : v).trim();
  return s ? s.charAt(0).toUpperCase() + s.slice(1).toLowerCase() : '';
}

function isActive_(v) {
  var s = String(v === null || v === undefined ? '' : v).trim().toLowerCase();
  return !(s === 'no' || s === 'n' || s === 'false' || s === '0' || s === 'inactive');
}

function getRoster_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Roster');
  if (!sh || sh.getLastRow() < 2) return [];
  var values = sh.getDataRange().getValues();
  var head = values[0].map(function (h) { return String(h).trim().toLowerCase().replace(/\s+/g, ''); });
  var iName = head.indexOf('name');
  if (iName < 0) return [];
  var iCat = head.indexOf('category');
  var iSide = head.indexOf('paddlingside');
  var iGen = head.indexOf('gender');
  var iAct = head.indexOf('active');

  var seen = {};
  var out = [];
  for (var i = 1; i < values.length; i++) {
    var name = String(values[i][iName] || '').trim();
    if (!name) continue;
    var key = name.toLowerCase();
    if (seen[key]) continue;           // ignore repeated names, first row wins
    seen[key] = true;
    out.push({
      name: name,
      category: iCat < 0 ? '' : String(values[i][iCat] || '').trim(),
      paddlingSide: iSide < 0 ? '' : cap_(values[i][iSide]),
      gender: iGen < 0 ? '' : cap_(values[i][iGen]),
      active: iAct < 0 ? true : isActive_(values[i][iAct])
    });
  }
  return out;
}

function getConfig_() {
  var sh = getSheet_('Config');
  var last = sh.getLastRow();
  var cfg = {};
  if (last < 2) return cfg;
  sh.getRange(2, 1, last - 1, 2).getValues().forEach(function (r) {
    if (r[0]) cfg[String(r[0])] = parse_(r[1], r[1]);
  });
  return cfg;
}

// ---------- web app entry points ----------

function doGet(e) {
  try {
    var all = !!(e && e.parameter && e.parameter.all === '1');
    var key = all ? 'get:all' : 'get:recent';
    var hit = null;
    try { hit = cacheGet_(key); } catch (e1) {}
    if (hit) {
      return ContentService.createTextOutput(hit).setMimeType(ContentService.MimeType.JSON);
    }
    var payload = JSON.stringify({
      success: true,
      data: getAvailability_(all),    // kept as "data" for the existing member form
      attendance: getAttendance_(),
      config: getConfig_(),
      roster: getRoster_()
    });
    try { cachePut_(key, payload); } catch (e2) {}
    return ContentService.createTextOutput(payload).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return json_({ success: false, error: String(err) });
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // prevents two simultaneous submits overwriting each other
    var data = JSON.parse(e.postData.contents);
    var action = data.action || 'upsert';

    if (action === 'upsert') {
      var sh = getSheet_('Availability');
      var name = str_(data.name).toLowerCase();
      var week = str_(data.weekStart);
      var row = findRow_(sh, function (r) {
        return str_(r[1]).toLowerCase() === name && str_(r[6]) === week;
      });
      writeRow_(sh, row, [
        new Date().toISOString(),
        str_(data.name),
        data.category || '',
        data.paddlingSide || '',
        data.gender || '',
        JSON.stringify(data.days || []),
        week,
        data.weekNumber || ''
      ]);
      return json_({ success: true, message: row === -1 ? 'Saved' : 'Updated' });
    }

    if (action === 'attendance') {
      var ash = getSheet_('Attendance');
      var date = str_(data.date);
      var arow = findRow_(ash, function (r) { return str_(r[0]) === date; });
      writeRow_(ash, arow, [
        date,
        data.dayName || '',
        data.display || '',
        JSON.stringify(data.attended || []),
        JSON.stringify(data.attendedDetails || []),
        new Date().toISOString()
      ]);
      return json_({ success: true, message: 'Attendance saved', attendance: getAttendance_() });
    }

    if (action === 'config') {
      var csh = getSheet_('Config');
      var key = str_(data.key);
      if (!key) throw new Error('Missing config key');
      var crow = findRow_(csh, function (r) { return str_(r[0]) === key; });
      writeRow_(csh, crow, [key, JSON.stringify(data.value), new Date().toISOString()]);
      return json_({ success: true, message: 'Config saved' });
    }

    throw new Error('Unknown action: ' + action);
  } catch (err) {
    return json_({ success: false, error: String(err) });
  } finally {
    try { cacheClear_(); } catch (e3) {}   // the next GET must see the new data
    try { lock.releaseLock(); } catch (e2) {}
  }
}

// ---------- one-off: copy old submissions into the new Availability tab ----------
// Run once from the editor (Run > migrateOldData) if you want past weeks kept.
function migrateOldData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var old = ss.getSheets().filter(function (s) {
    return ['Availability', 'Attendance', 'Config'].indexOf(s.getName()) === -1;
  })[0];
  if (!old) return;
  var target = getSheet_('Availability');
  old.getDataRange().getValues().forEach(function (r, i) {
    var looksLikeHeader = i === 0 && String(r[1]).toLowerCase() === 'name';
    if (looksLikeHeader || !r[1]) return;
    // old layout: timestamp, name, category, side, days, weekStart, weekNumber
    var dup = findRow_(target, function (t) {
      return str_(t[1]).toLowerCase() === str_(r[1]).toLowerCase() && str_(t[6]) === str_(r[5]);
    });
    writeRow_(target, dup, [str_(r[0]), str_(r[1]), r[2], r[3], '', r[4], str_(r[5]), r[6]]);
  });
}

// ---------- optional: physically delete older duplicate rows from the Availability tab ----------
// Run once from the editor (Run > removeDuplicateRows). Reading already ignores duplicates,
// so this only tidies the sheet.
function removeDuplicateRows() {
  var sh = getSheet_('Availability');
  var last = sh.getLastRow();
  if (last < 2) return;
  var rows = sh.getRange(2, 1, last - 1, HEADERS.Availability.length).getValues();
  var newest = {};
  rows.forEach(function (r, i) {
    var key = str_(r[1]).toLowerCase() + '|' + str_(r[6]);
    if (!newest[key] || str_(r[0]) >= str_(rows[newest[key]][0])) newest[key] = i;
  });
  for (var i = rows.length - 1; i >= 0; i--) {   // bottom-up so row numbers stay valid
    var key = str_(rows[i][1]).toLowerCase() + '|' + str_(rows[i][6]);
    if (newest[key] !== i) sh.deleteRow(i + 2);
  }
}
