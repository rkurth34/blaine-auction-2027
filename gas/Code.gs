/** @OnlyCurrentDoc */
/**
 * Blaine 2027 Auction Data: Apps Script backend.
 * Bound to the "Blaine 2027 Auction Data" sheet. The sheet is the source of truth;
 * the pages on Vercel read and write it through this web app.
 *
 * Every request is a POST with a text/plain JSON body that carries the team passcode
 * as key. The passcode lives in Script Properties (PASSCODE), never in the repo.
 *
 * Actions:
 *   read      { tabs: [name] }                 -> { tabs: { name: [rows] } }
 *   upsert    { tab, row }                     -> { row }        insert or replace by id
 *   patch     { tab, id, data }                -> { row }        merge fields into one row
 *   remove    { tab, id }                      -> { ok }
 *   claim     { tab, id, asker, by }           -> { row } or { taken: name }
 *   setting   { name, value }                  -> { ok }         Tracker settings tab
 *   saveOrg   { chart, rev }                   -> { rev }        Org chart JSON + readable tabs
 *   readOrg   {}                               -> { chart, rev }
 */

var TABS = ['Contacts', 'Business donations', 'Sponsors', 'Family donations', 'Hosted events',
  'Class art', 'Tracker settings', 'Tools', 'Org chart', 'Timeline'];
var ORG_DATA = 'Org chart data';

function doGet() {
  return out({ ok: true, service: 'blaine-2027-auction' });
}

function doPost(e) {
  var res;
  try {
    var b = JSON.parse(e.postData.contents);
    var pass = PropertiesService.getScriptProperties().getProperty('PASSCODE');
    if (!pass || String(b.key || '') !== pass) return out({ error: 'passcode' });
    res = handle(b);
  } catch (err) {
    res = { error: String(err && err.message || err) };
  }
  return out(res);
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function handle(b) {
  switch (b.action) {
    case 'read': return { tabs: readTabs(b.tabs || TABS) };
    case 'upsert': return locked(function () { return { row: upsert(b.tab, b.row) }; });
    case 'patch': return locked(function () { return { row: patch(b.tab, b.id, b.data) }; });
    case 'remove': return locked(function () { return { ok: remove(b.tab, b.id) }; });
    case 'claim': return locked(function () { return claim(b.tab, b.id, b.asker, b.by); });
    case 'setting': return locked(function () { setSetting(b.name, b.value); return { ok: true }; });
    case 'readOrg': return readOrg();
    case 'saveOrg': return locked(function () { return saveOrg(b.chart); });
    default: return { error: 'unknown action' };
  }
}

function locked(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function ss() { return SpreadsheetApp.getActiveSpreadsheet(); }

function sheet(name) {
  if (TABS.indexOf(name) === -1) throw new Error('unknown tab');
  var sh = ss().getSheetByName(name);
  if (!sh) throw new Error('missing tab ' + name);
  return sh;
}

function cell(v) {
  if (v instanceof Date) return v.toISOString();
  return v;
}

// Keep text as text: a leading = + @ would otherwise become a formula.
function safe(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'string' && /^[=+@]/.test(v)) return "'" + v;
  return v;
}

function table(name) {
  var sh = sheet(name);
  var values = sh.getDataRange().getValues();
  var headers = values[0].map(String);
  return { sh: sh, headers: headers, rows: values.slice(1) };
}

function toObj(headers, r) {
  var o = {};
  for (var i = 0; i < headers.length; i++) if (headers[i]) o[headers[i]] = cell(r[i]);
  return o;
}

function readTabs(names) {
  var res = {};
  names.forEach(function (n) {
    var t = table(n);
    res[n] = t.rows.filter(function (r) { return r.join('') !== ''; }).map(function (r) { return toObj(t.headers, r); });
  });
  return res;
}

function findRow(t, id) {
  var col = t.headers.indexOf('id');
  if (col === -1) throw new Error('no id column');
  for (var i = 0; i < t.rows.length; i++) if (String(t.rows[i][col]) === String(id)) return i;
  return -1;
}

function writeRow(t, i, obj) {
  var line = t.headers.map(function (h) { return safe(obj[h]); });
  t.sh.getRange(i + 2, 1, 1, line.length).setValues([line]);
}

function upsert(name, row) {
  var t = table(name);
  if (!row.id) row.id = 'u-' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
  var i = findRow(t, row.id);
  if (i === -1) {
    t.sh.appendRow(t.headers.map(function (h) { return safe(row[h]); }));
    return row;
  }
  writeRow(t, i, row);
  return row;
}

function patch(name, id, data) {
  var t = table(name);
  var i = findRow(t, id);
  if (i === -1) throw new Error('row not found');
  var obj = toObj(t.headers, t.rows[i]);
  for (var k in data) if (t.headers.indexOf(k) !== -1 && k !== 'id') obj[k] = data[k];
  writeRow(t, i, obj);
  return obj;
}

function remove(name, id) {
  var t = table(name);
  var i = findRow(t, id);
  if (i === -1) return false;
  t.sh.deleteRow(i + 2);
  return true;
}

function claim(name, id, asker, by) {
  var t = table(name);
  var i = findRow(t, id);
  if (i === -1) throw new Error('row not found');
  var obj = toObj(t.headers, t.rows[i]);
  if (obj.asker && obj.asker !== asker) return { taken: obj.asker, row: obj };
  obj.asker = asker;
  obj.updatedAt = new Date().toISOString();
  obj.updatedBy = by || asker;
  writeRow(t, i, obj);
  return { row: obj };
}

function setSetting(key, value) {
  if (!key) throw new Error('missing setting name');
  var t = table('Tracker settings');
  for (var i = 0; i < t.rows.length; i++) {
    if (String(t.rows[i][0]) === String(key)) {
      t.sh.getRange(i + 2, 2).setValue(safe(value));
      return;
    }
  }
  t.sh.appendRow([key, safe(value)]);
}

/* ---------- org chart: full JSON kept in a data tab, readable copy in Org chart + Timeline ---------- */

function orgSheet() {
  var sh = ss().getSheetByName(ORG_DATA);
  if (!sh) {
    sh = ss().insertSheet(ORG_DATA);
    sh.getRange(1, 1, 1, 2).setValues([['rev', 'json (split across cells, do not edit)']]);
  }
  return sh;
}

function readOrg() {
  var sh = orgSheet();
  var last = sh.getLastRow();
  if (last < 2) return { chart: null, rev: '' };
  var vals = sh.getRange(2, 1, last - 1, 2).getValues();
  var json = vals.map(function (r) { return r[1]; }).join('');
  return { chart: json ? JSON.parse(json) : null, rev: String(vals[0][0]) };
}

function saveOrg(chart) {
  var sh = orgSheet();
  var json = JSON.stringify(chart);
  var parts = [];
  for (var i = 0; i < json.length; i += 40000) parts.push(json.slice(i, i + 40000));
  var rev = new Date().toISOString();
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 2).clearContent();
  sh.getRange(2, 1, parts.length, 2).setValues(parts.map(function (p, n) { return [n === 0 ? rev : '', p]; }));
  sh.getRange(2, 2, parts.length, 1).setNumberFormat('@');
  writeOrgTabs(chart);
  return { rev: rev };
}

function writeOrgTabs(c) {
  var rows = [];
  function seats(side, group, r) {
    (r.seats || []).forEach(function (s) {
      rows.push([side, group, r.id, r.name, s.id, s.n, s.s, s.e, s.note, r.owns, r.lots, r.flag, r.y25, r.notes]);
    });
  }
  seats('Leadership', '', c.chairs);
  (c.branches || []).forEach(function (b) {
    (b.groups || []).forEach(function (g) { (g.roles || []).forEach(function (r) { seats(b.name, g.name, r); }); });
  });
  ((c.rail && c.rail.roles) || []).forEach(function (r) { seats(c.rail.name, '', r); });
  replaceBody('Org chart', rows);
  replaceBody('Timeline', (c.steps || []).map(function (s) { return [s.id, s.when, s.title, s.text]; }));
}

function replaceBody(name, rows) {
  var sh = sheet(name);
  var width = sh.getLastColumn();
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, width).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, rows[0].length).setValues(rows.map(function (r) { return r.map(safe); }));
}

function authorize() {
  return ss().getName();
}
