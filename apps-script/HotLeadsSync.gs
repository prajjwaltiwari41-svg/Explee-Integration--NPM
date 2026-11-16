/**
 * Explee hot leads -> Google Sheets
 *
 * Every 10 minutes, fetches new hot leads from Explee and adds them to the
 * "Hot Leads" tab of this spreadsheet. A lead that replies again is updated
 * in place (matched on Person ID), never duplicated.
 *
 * One-time setup: paste this file into Extensions > Apps Script, put your
 * API key below, choose the "setup" function and press Run.
 *
 * You can add your own columns to the RIGHT of the synced ones (Status,
 * Owner, Next step...). The sync only writes its own columns, so your notes
 * are never overwritten. Sorting or filtering the rows is fine too.
 */

var EXPLEE_API_KEY = 'PASTE_YOUR_EXPLEE_API_KEY_HERE';
var EXPLEE_PROJECT_ID = 50789;
var SHEET_NAME = 'Hot Leads';
var EVERY_MINUTES = 10;

var API_BASE = 'https://api.explee.com/public/api/v1';
var PAGE_SIZE = 200;
var CURSOR_KEY = 'explee_since';

// Synced columns, in order. Column A (Person ID) is the match key.
var COLUMNS = [
  ['Person ID',      function (l) { return l.person_id; }],
  ['Name',           function (l) { return l.erased ? '(removed at contact\'s request)' : l.name; }],
  ['Email',          function (l) { return l.email; }],
  ['Job Title',      function (l) { return l.job_title; }],
  ['Company',        function (l) { return l.company_name; }],
  ['Company Domain', function (l) { return l.company_domain; }],
  ['LinkedIn',       function (l) { return l.linkedin_url; }],
  ['Country',        function (l) { return l.country; }],
  ['Phone',          function (l) { return l.phone; }],
  ['Why Hot (their reply)', function (l) { return l.why_hot; }],
  ['Became Hot At',  function (l) { return toDate_(l.became_hot_at); }],
  ['Replied At',     function (l) { return toDate_(l.replied_at); }],
  ['Explee Note',    function (l) { return l.note; }],
  ['Campaign ID',    function (l) { return l.campaign_id; }],
  ['Last Synced',    function () { return new Date(); }]
];
var DATE_COLUMNS = ['Became Hot At', 'Replied At', 'Last Synced'];

/** Run this once by hand. Creates the tab, the 10-minute timer, and does a first sync. */
function setup() {
  if (!EXPLEE_API_KEY || EXPLEE_API_KEY.indexOf('PASTE_') === 0) {
    throw new Error('Put your Explee API key in EXPLEE_API_KEY at the top of the script first.');
  }
  getSheet_();
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'syncHotLeads') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('syncHotLeads').timeBased().everyMinutes(EVERY_MINUTES).create();
  syncHotLeads();
}

/** Called by the timer. Safe to run by hand any time. */
function syncHotLeads() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return; // a previous run is still going
  try {
    var props = PropertiesService.getScriptProperties();
    var since = props.getProperty(CURSOR_KEY); // empty on first run = bring in all existing hot leads
    var leads = fetchHotLeads_(since);
    if (leads.length === 0) return;

    // Apply oldest first, so if a lead appears twice the newest version wins.
    leads.sort(function (a, b) { return time_(a.became_hot_at) - time_(b.became_hot_at); });

    var sheet = getSheet_();
    upsertLeads_(sheet, leads);

    // Advance the cursor only after the sheet was written.
    var newest = since;
    leads.forEach(function (l) {
      if (l.became_hot_at && (!newest || time_(l.became_hot_at) > time_(newest))) newest = l.became_hot_at;
    });
    if (newest) props.setProperty(CURSOR_KEY, newest);
  } finally {
    lock.releaseLock();
  }
}

/** Start over: forget the last-checked time so the next run re-reads every hot lead. */
function resetCursor() {
  PropertiesService.getScriptProperties().deleteProperty(CURSOR_KEY);
}

function fetchHotLeads_(since) {
  var all = [];
  var offset = 0;
  for (var page = 0; page < 100; page++) {
    var params = {
      project_id: EXPLEE_PROJECT_ID,
      source: 'agent',
      limit: PAGE_SIZE,
      offset: offset
    };
    if (since) params.since = since;
    var data = apiGet_('/autogtm/hot-leads', params);
    all = all.concat(data.leads || []);
    if (!data.has_more || data.next_offset === null || data.next_offset === undefined) break;
    offset = data.next_offset;
  }
  return all;
}

function apiGet_(path, params) {
  var qs = Object.keys(params).map(function (k) {
    return encodeURIComponent(k) + '=' + encodeURIComponent(params[k]);
  }).join('&');
  var url = API_BASE + path + (qs ? '?' + qs : '');
  var lastErr;
  for (var attempt = 0; attempt < 3; attempt++) {
    var res = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { 'X-API-Key': EXPLEE_API_KEY, 'Accept': 'application/json' },
      muteHttpExceptions: true
    });
    var code = res.getResponseCode();
    if (code === 200) return JSON.parse(res.getContentText());
    lastErr = 'Explee API returned ' + code + ': ' + res.getContentText().slice(0, 500);
    if (code !== 429 && code < 500) break; // bad key, no credit, etc. - retrying won't help
    Utilities.sleep(2000 * (attempt + 1));
  }
  throw new Error(lastErr);
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    var first = ss.getSheets()[0];
    var a1 = String(first.getRange(1, 1).getValue());
    if (ss.getSheets().length === 1 && (first.getLastRow() === 0 || a1 === COLUMNS[0][0])) {
      sheet = first.setName(SHEET_NAME); // reuse the blank/starter tab
    } else {
      sheet = ss.insertSheet(SHEET_NAME);
    }
  }
  var headers = COLUMNS.map(function (c) { return c[0]; });
  var current = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  if (current.join('|') !== headers.join('|')) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sheet.setFrozenRows(1);
    // Text columns stay text (keeps phone numbers like +44... and replies starting with "=" intact).
    var maxRows = sheet.getMaxRows();
    COLUMNS.forEach(function (c, i) {
      var fmt = DATE_COLUMNS.indexOf(c[0]) >= 0 ? 'yyyy-mm-dd hh:mm' : '@';
      sheet.getRange(2, i + 1, Math.max(maxRows - 1, 1), 1).setNumberFormat(fmt);
    });
    sheet.setColumnWidth(10, 400); // Why Hot
  }
  return sheet;
}

function upsertLeads_(sheet, leads) {
  var width = COLUMNS.length;
  var lastRow = sheet.getLastRow();
  var rowById = {};
  if (lastRow > 1) {
    var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (ids[i][0] !== '') rowById[String(ids[i][0])] = i + 2;
    }
  }

  var appended = [];
  var appendedIndex = {};
  leads.forEach(function (lead) {
    var values = COLUMNS.map(function (c) { return clean_(c[1](lead)); });
    var id = String(lead.person_id);
    if (rowById[id]) {
      sheet.getRange(rowById[id], 1, 1, width).setValues([values]);
    } else if (appendedIndex[id] !== undefined) {
      appended[appendedIndex[id]] = values;
    } else {
      appendedIndex[id] = appended.length;
      appended.push(values);
    }
  });

  if (appended.length) {
    var start = sheet.getLastRow() + 1;
    if (start + appended.length - 1 > sheet.getMaxRows()) {
      sheet.insertRowsAfter(sheet.getMaxRows(), start + appended.length - 1 - sheet.getMaxRows());
      COLUMNS.forEach(function (c, i) {
        var fmt = DATE_COLUMNS.indexOf(c[0]) >= 0 ? 'yyyy-mm-dd hh:mm' : '@';
        sheet.getRange(start, i + 1, appended.length, 1).setNumberFormat(fmt);
      });
    }
    sheet.getRange(start, 1, appended.length, width).setValues(appended);
  }
}

function clean_(v) {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) return v;
  if (typeof v === 'string' && v.charAt(0) === '=') return "'" + v; // never run as a formula
  return v;
}

function toDate_(s) {
  if (!s) return '';
  var d = new Date(s);
  return isNaN(d.getTime()) ? s : d;
}

function time_(s) {
  var t = new Date(s).getTime();
  return isNaN(t) ? 0 : t;
}
