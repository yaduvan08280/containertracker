/**
 * Container Tracker — data layer.
 * 5 sheets only: Containers (one wide row per container), Users, Approvals, Log, Config.
 * Reads are served from CacheService (chunked JSON); writes update the sheet row(s)
 * and the cache in the same locked operation so readers never hit the sheet twice.
 */

var SH = { C: 'Containers', U: 'Users', A: 'Approvals', L: 'Log', G: 'Config' };
var CACHE_TTL = 21600; // 6h (max)
var CHUNK = 90000;
var TZ = Session.getScriptTimeZone() || 'Asia/Kolkata';

var _ss = null, _cols = null, _users = null, _cfg = null, _list = null;

function ss_() {
  if (_ss) return _ss;
  var id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  _ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActive();
  if (!_ss) throw new Error('No spreadsheet. Run setup() from the script editor first.');
  return _ss;
}
function sheet_(name) { return ss_().getSheetByName(name); }

/* ---------------------------------------------------------------- columns */

/** Ordered [key, header, kind] for the Containers sheet. kind: d = date(ms), n = number, s = string */
function containerCols_() {
  if (_cols) return _cols;
  var F = DEF.fields, out = [];
  function add(k, l, kind) { out.push([k, l, kind]); }
  function kindOf(k) { var t = F[k].t; return (t === 'date' || t === 'datetime') ? 'd' : (t === 'money' || t === 'int') ? 'n' : 's'; }
  add('id', 'Container ID', 's'); add('bookingId', 'Booking ID', 's'); add('seq', 'Seq', 'n');
  add('status', 'Status', 's'); add('stage', 'Current Step', 's');
  add('createdAt', 'Booking Timestamp', 'd'); add('createdBy', 'Booked By', 's');
  DEF.steps.forEach(function (s) {
    s.fields.forEach(function (k) { add(k, F[k].l + (s.n > 1 && /^Remarks?$/.test(F[k].l) ? ' (' + s.short + ')' : ''), kindOf(k)); });
    add('s' + s.n + '_at', 'S' + s.n + ' ' + s.short + ' — Done At', 'd');
    add('s' + s.n + '_by', 'S' + s.n + ' ' + s.short + ' — Done By', 's');
    add('s' + s.n + '_due', 'S' + s.n + ' ' + s.short + ' — Deadline', 'd');
  });
  DEF.sections.forEach(function (s) { s.fields.forEach(function (k) { add(k, F[k].l, kindOf(k)); }); });
  add('docs_at', 'Docs Completed At', 'd'); add('docs_by', 'Docs Completed By', 's');
  DEF.chargeGroups.forEach(function (g) {
    g.items.forEach(function (it) {
      add(it[0], g.name + ' › ' + it[1], 'n');
      add(it[0] + '_p1', g.name + ' › ' + it[1] + ' — Photo 1', 's');
      add(it[0] + '_p2', g.name + ' › ' + it[1] + ' — Photo 2', 's');
    });
  });
  for (var i = 1; i <= DEF.extraSlots; i++) {
    add('ex' + i + '_desc', 'Extra Expense ' + i + ' — Description', 's');
    add('ex' + i + '_amt', 'Extra Expense ' + i + ' — Amount', 'n');
    add('ex' + i + '_p1', 'Extra Expense ' + i + ' — Photo 1', 's');
    add('ex' + i + '_p2', 'Extra Expense ' + i + ' — Photo 2', 's');
  }
  add('statusAt', 'Status Changed At', 'd'); add('statusBy', 'Status Changed By', 's'); add('statusNote', 'Status Note', 's');
  add('updatedAt', 'Last Updated', 'd');
  _cols = out;
  return out;
}

var USER_COLS = ['id', 'username', 'name', 'role', 'mobile', 'telegramChatId', 'steps', 'active', 'mustChange', 'passHash', 'salt', 'linkCode', 'createdAt'];
var APPR_COLS = ['id', 'createdAt', 'containerId', 'containerLabel', 'bookingNo', 'field', 'fieldLabel', 'oldValue', 'newValue', 'reason', 'requestedBy', 'status', 'decidedBy', 'decidedAt', 'note'];
var LOG_COLS = ['Timestamp', 'Type', 'Container ID', 'Booking No', 'Container No', 'User', 'Detail', 'Old Value', 'New Value'];

/* ---------------------------------------------------------------- chunked cache */

function cachePut_(key, obj) {
  var c = CacheService.getScriptCache(), s = JSON.stringify(obj), n = Math.ceil(s.length / CHUNK) || 1, m = {};
  for (var i = 0; i < n; i++) m[key + '_' + i] = s.substr(i * CHUNK, CHUNK);
  m[key] = String(n);
  try { c.putAll(m, CACHE_TTL); } catch (e) { c.remove(key); }
}
function cacheGet_(key) {
  var c = CacheService.getScriptCache(), n = c.get(key);
  if (!n) return null;
  var keys = []; for (var i = 0; i < +n; i++) keys.push(key + '_' + i);
  var got = c.getAll(keys), s = '';
  for (var j = 0; j < keys.length; j++) { if (got[keys[j]] == null) return null; s += got[keys[j]]; }
  try { return JSON.parse(s); } catch (e) { return null; }
}
function cacheDrop_(key) { CacheService.getScriptCache().remove(key); }

function version_() { return Number(PropertiesService.getScriptProperties().getProperty('VER') || 0); }
function bumpVersion_() { var v = Date.now(); PropertiesService.getScriptProperties().setProperty('VER', String(v)); return v; }

/* ---------------------------------------------------------------- containers */

/** All containers as compact objects ({key: value} for non-empty cells, plus _r = sheet row). */
function loadContainers_() {
  if (_list) return _list;
  var cached = cacheGet_('C');
  if (cached) { _list = cached; return _list; }
  var v0 = version_();
  var sh = sheet_(SH.C), last = sh.getLastRow(), lastCol = sh.getLastColumn();
  var list = [];
  if (last >= 2) {
    var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    var byLabel = {}; containerCols_().forEach(function (c) { byLabel[c[1]] = c; });
    var map = hdr.map(function (h) { return byLabel[h] || null; });
    var vals = sh.getRange(2, 1, last - 1, lastCol).getValues();
    for (var r = 0; r < vals.length; r++) {
      var row = vals[r];
      if (!row[0]) continue;
      var o = { _r: r + 2 };
      for (var i = 0; i < row.length; i++) {
        var col = map[i], v = row[i];
        if (v === '' || v === null) continue;
        if (!col) { (o._x = o._x || {})[i] = v instanceof Date ? { d: v.getTime() } : v; continue; } // user-added column: keep it
        if (v instanceof Date) v = v.getTime();
        else if (col[2] === 'n') v = Number(v);
        else if (col[2] === 'd') { var t = new Date(v).getTime(); if (isNaN(t)) continue; v = t; }
        else v = String(v);
        o[col[0]] = v;
      }
      list.push(o);
    }
  }
  _list = list;
  if (version_() === v0) cachePut_('C', list); // don't cache if a writer raced us
  return list;
}

function findContainer_(id) {
  var list = loadContainers_();
  for (var i = 0; i < list.length; i++) if (list[i].id === id) return list[i];
  throw new Error('Container not found: ' + id);
}

function headerKeys_(sh) {
  var byLabel = {}; containerCols_().forEach(function (c) { byLabel[c[1]] = c; });
  return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (h) { return byLabel[h] || null; });
}

function rowOf_(o, hdr) {
  return hdr.map(function (col, i) {
    if (!col) { var x = o._x && o._x[i]; return x === undefined ? '' : (x && x.d ? new Date(x.d) : x); } // don't wipe columns we don't own
    var v = o[col[0]];
    if (v === undefined || v === null) return '';
    if (col[2] === 'd') return new Date(Number(v));
    return v;
  });
}

/** Persist changed containers (new ones have no _r). Must be called inside a lock. */
function saveContainers_(changed) {
  if (!changed.length) return;
  var sh = sheet_(SH.C), hdr = headerKeys_(sh), now = Date.now();
  var existing = [], fresh = [];
  changed.forEach(function (c) {
    c.updatedAt = now;
    var st = ctState(c, getRules_(), now);
    c.stage = c.status && c.status !== 'ACTIVE' ? c.status : (st.cur ? st.cur + ' · ' + DEF.steps[st.cur - 1].name : 'Pipeline complete');
    (c._r ? existing : fresh).push(c);
  });
  // group contiguous rows into single setValues calls
  existing.sort(function (a, b) { return a._r - b._r; });
  for (var i = 0; i < existing.length;) {
    var j = i; while (j + 1 < existing.length && existing[j + 1]._r === existing[j]._r + 1) j++;
    var rows = []; for (var k = i; k <= j; k++) rows.push(rowOf_(existing[k], hdr));
    sh.getRange(existing[i]._r, 1, rows.length, hdr.length).setValues(rows);
    i = j + 1;
  }
  if (fresh.length) {
    var start = Math.max(sh.getLastRow(), 1) + 1;
    sh.getRange(start, 1, fresh.length, hdr.length).setValues(fresh.map(function (c) { return rowOf_(c, hdr); }));
    fresh.forEach(function (c, idx) { c._r = start + idx; });
    var list = loadContainers_();
    fresh.forEach(function (c) { list.push(c); });
  }
  SpreadsheetApp.flush(); // make rows visible before VER moves, so a racing reader can't cache the old sheet
  bumpVersion_();
  cachePut_('C', loadContainers_());
}

/* ---------------------------------------------------------------- users */

function loadUsers_() {
  if (_users) return _users;
  var cached = cacheGet_('U');
  if (cached) { _users = cached; return _users; }
  var sh = sheet_(SH.U), last = sh.getLastRow(), out = [];
  if (last >= 2) {
    sh.getRange(2, 1, last - 1, USER_COLS.length).getValues().forEach(function (row, i) {
      if (!row[0]) return;
      var u = { _r: i + 2 };
      USER_COLS.forEach(function (k, j) { u[k] = row[j] instanceof Date ? row[j].getTime() : row[j]; });
      u.steps = String(u.steps || '').split(',').map(function (s) { return s.trim(); }).filter(String);
      u.active = u.active === true || String(u.active).toUpperCase() === 'TRUE';
      u.mustChange = u.mustChange === true || String(u.mustChange).toUpperCase() === 'TRUE';
      u.username = String(u.username).toLowerCase();
      u.mobile = String(u.mobile || ''); u.telegramChatId = String(u.telegramChatId || '');
      out.push(u);
    });
  }
  _users = out;
  cachePut_('U', out);
  return out;
}

function saveUser_(u) {
  var sh = sheet_(SH.U);
  var row = USER_COLS.map(function (k) {
    if (k === 'steps') return (u.steps || []).join(',');
    if (k === 'createdAt' && typeof u[k] === 'number') return new Date(u[k]);
    return u[k] === undefined ? '' : u[k];
  });
  if (u._r) sh.getRange(u._r, 1, 1, row.length).setValues([row]);
  else { u._r = Math.max(sh.getLastRow(), 1) + 1; sh.getRange(u._r, 1, 1, row.length).setValues([row]); loadUsers_().push(u); }
  cachePut_('U', loadUsers_());
  bumpVersion_();
}

function userByName_(username) {
  username = String(username || '').toLowerCase().trim();
  var us = loadUsers_();
  for (var i = 0; i < us.length; i++) if (us[i].username === username) return us[i];
  return null;
}
function userById_(id) {
  var us = loadUsers_();
  for (var i = 0; i < us.length; i++) if (us[i].id === id) return us[i];
  return null;
}
function publicUser_(u) {
  return { id: u.id, username: u.username, name: u.name, role: u.role, mobile: u.mobile, steps: u.steps, active: u.active, tg: !!u.telegramChatId };
}

/* ---------------------------------------------------------------- config (key | JSON value) */

function loadCfg_() {
  if (_cfg) return _cfg;
  var cached = cacheGet_('G');
  if (cached) { _cfg = cached; return _cfg; }
  var sh = sheet_(SH.G), last = sh.getLastRow(), out = {};
  if (last >= 2) sh.getRange(2, 1, last - 1, 2).getValues().forEach(function (r) {
    if (!r[0]) return;
    try { out[r[0]] = JSON.parse(r[1]); } catch (e) { out[r[0]] = r[1]; }
  });
  _cfg = out;
  cachePut_('G', out);
  return out;
}
function setCfg_(key, value) {
  var sh = sheet_(SH.G), last = sh.getLastRow(), json = JSON.stringify(value);
  var keys = last >= 2 ? sh.getRange(2, 1, last - 1, 1).getValues().map(function (r) { return r[0]; }) : [];
  var i = keys.indexOf(key);
  if (i >= 0) sh.getRange(i + 2, 2).setValue(json);
  else sh.getRange(last + 1, 1, 1, 2).setValues([[key, json]]);
  loadCfg_()[key] = value;
  cachePut_('G', _cfg);
}

function getRules_() {
  var r = loadCfg_().rules || {}, out = {};
  Object.keys(DEF.defaultRules).forEach(function (k) { out[k] = r[k] || DEF.defaultRules[k]; });
  return out;
}
function getSettings_() {
  var s = loadCfg_().settings || {};
  return {
    remindBeforeHours: s.remindBeforeHours != null ? Number(s.remindBeforeHours) : 2,
    escalateRepeatHours: s.escalateRepeatHours != null ? Number(s.escalateRepeatHours) : 4,
    quietStart: s.quietStart != null ? Number(s.quietStart) : 22,
    quietEnd: s.quietEnd != null ? Number(s.quietEnd) : 7,
    notifyAdminOnDelay: s.notifyAdminOnDelay !== false,
    nudgeAssigneeOnOverdue: s.nudgeAssigneeOnOverdue !== false,
    botUsername: s.botUsername || ''
  };
}

/* ---------------------------------------------------------------- approvals */

function loadApprovals_() {
  var cached = cacheGet_('A');
  if (cached) return cached;
  var sh = sheet_(SH.A), last = sh.getLastRow(), out = [];
  if (last >= 2) sh.getRange(2, 1, last - 1, APPR_COLS.length).getValues().forEach(function (row, i) {
    if (!row[0]) return;
    var a = { _r: i + 2 };
    APPR_COLS.forEach(function (k, j) { a[k] = row[j] instanceof Date ? row[j].getTime() : row[j]; });
    if (a.newValue !== '' && a.newValue != null) { try { a.newValue = JSON.parse(a.newValue); } catch (e) {} }
    if (a.oldValue !== '' && a.oldValue != null) { try { a.oldValue = JSON.parse(a.oldValue); } catch (e) {} }
    out.push(a);
  });
  cachePut_('A', out);
  return out;
}
function saveApproval_(a) {
  var sh = sheet_(SH.A), list = loadApprovals_();
  var row = APPR_COLS.map(function (k) {
    var v = a[k];
    if (k === 'newValue' || k === 'oldValue') return v === undefined || v === '' ? '' : JSON.stringify(v);
    if ((k === 'createdAt' || k === 'decidedAt') && v) return new Date(v);
    return v === undefined ? '' : v;
  });
  if (a._r) {
    sh.getRange(a._r, 1, 1, row.length).setValues([row]);
    for (var i = 0; i < list.length; i++) if (list[i].id === a.id) list[i] = a;
  } else {
    a._r = Math.max(sh.getLastRow(), 1) + 1;
    sh.getRange(a._r, 1, 1, row.length).setValues([row]);
    list.push(a);
  }
  cachePut_('A', list);
  bumpVersion_();
}

/* ---------------------------------------------------------------- log */

var _logBuf = [];
function log_(type, c, user, detail, oldV, newV, key) {
  _logBuf.push([new Date(), type, c ? c.id : '', c ? (c.bookingNo || '') : '', c ? (c.containerNo || '') : '',
    user ? user.name : 'System', detail || '', fmtVal_(oldV, key), fmtVal_(newV, key)]);
}
function flushLog_() {
  if (!_logBuf.length) return;
  var sh = sheet_(SH.L);
  sh.getRange(sh.getLastRow() + 1, 1, _logBuf.length, LOG_COLS.length).setValues(_logBuf);
  _logBuf = [];
}
function fmtVal_(v, key) {
  if (v === undefined || v === null || v === '') return '';
  if (key && DEF.fields[key]) {
    var t = DEF.fields[key].t;
    if (t === 'date') return Utilities.formatDate(new Date(Number(v)), TZ, 'dd MMM yyyy');
    if (t === 'datetime') return Utilities.formatDate(new Date(Number(v)), TZ, 'dd MMM yyyy HH:mm');
    if (t === 'money') return '₹ ' + Number(v).toLocaleString('en-IN');
  }
  return String(v);
}

/** History rows for one container — located with TextFinder so the Log can grow large. */
function historyFor_(id) {
  var sh = sheet_(SH.L), last = sh.getLastRow();
  if (last < 2) return [];
  var hits = sh.getRange(2, 3, last - 1, 1).createTextFinder(id).matchEntireCell(true).findAll();
  if (!hits.length) return [];
  var rows = hits.map(function (r) { return r.getRow(); });
  var lo = Math.min.apply(null, rows), hi = Math.max.apply(null, rows);
  var want = {}; rows.forEach(function (r) { want[r] = 1; });
  var vals = sh.getRange(lo, 1, hi - lo + 1, LOG_COLS.length).getValues(), out = [];
  vals.forEach(function (v, i) {
    if (!want[lo + i]) return;
    out.push({ ts: v[0] instanceof Date ? v[0].getTime() : v[0], type: v[1], user: v[5], detail: v[6], oldV: String(v[7]), newV: String(v[8]) });
  });
  return out.reverse();
}
