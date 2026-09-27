/**
 * Container Tracker — web app entry, auth and API.
 * Every public function takes the session token first; the browser calls them via google.script.run.
 */

/* ================================================================ web */

function doGet(e) {
  var P = PropertiesService.getScriptProperties();
  if (!P.getProperty('SECRET')) setup();
  try {
    var url = ScriptApp.getService().getUrl();
    if (url && /\/exec$/.test(url) && P.getProperty('APP_URL') !== url) P.setProperty('APP_URL', url);
  } catch (err) {}
  var t = HtmlService.createTemplateFromFile('Index');
  t.deep = String((e && e.parameter && e.parameter.c) || '').replace(/[^\w-]/g, '');
  return t.evaluate()
    .setTitle('Container Tracker')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(name) { return HtmlService.createHtmlOutputFromFile(name).getContent(); }

/** Shared definitions + deadline maths, injected into the page so client and server agree. */
function sharedJs_() {
  return 'var DEF=' + JSON.stringify(DEF) + ';\n' + ctDue.toString() + '\n' + ctState.toString() + '\n';
}

function appUrl_() {
  return PropertiesService.getScriptProperties().getProperty('APP_URL') || (function () { try { return ScriptApp.getService().getUrl(); } catch (e) { return ''; } })();
}

/* ================================================================ setup */

/** Run once from the editor (Run ▸ setup). Safe to re-run: it only adds what is missing. */
function setup() {
  var P = PropertiesService.getScriptProperties();
  var id = P.getProperty('SHEET_ID'), ss = null;
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) {} }
  if (!ss) ss = SpreadsheetApp.getActive();
  if (!ss) ss = SpreadsheetApp.create('Container Tracker — Database');
  P.setProperty('SHEET_ID', ss.getId());
  _ss = ss;
  try { ss.setSpreadsheetTimeZone(TZ); } catch (e) {}

  var cols = containerCols_();
  ensureSheet_(SH.C, cols.map(function (c) { return c[1]; }), true);
  ensureSheet_(SH.U, USER_COLS);
  ensureSheet_(SH.A, APPR_COLS);
  ensureSheet_(SH.L, LOG_COLS);
  ensureSheet_(SH.G, ['Key', 'Value (JSON)']);
  var def = ss.getSheetByName('Sheet1');
  if (def && def.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(def);

  // Column formats: text columns stay text (keeps leading zeros in invoice / bill numbers), dates readable.
  var shC = ss.getSheetByName(SH.C), hdr = shC.getRange(1, 1, 1, shC.getLastColumn()).getValues()[0];
  var txt = [], dts = [];
  cols.forEach(function (c) {
    var i = hdr.indexOf(c[1]);
    if (i < 0) return;
    var a = colA1_(i + 1) + '2:' + colA1_(i + 1);
    if (c[2] === 's') txt.push(a); else if (c[2] === 'd') dts.push(a);
  });
  if (txt.length) shC.getRangeList(txt).setNumberFormat('@');
  if (dts.length) shC.getRangeList(dts).setNumberFormat('dd-mmm-yyyy hh:mm');
  ss.getSheetByName(SH.U).getRange('A2:L').setNumberFormat('@');
  ss.getSheetByName(SH.A).getRange('H2:H').setNumberFormat('@');
  ss.getSheetByName(SH.A).getRange('I2:I').setNumberFormat('@');
  ss.getSheetByName(SH.L).getRange('A:A').setNumberFormat('dd-mmm-yyyy hh:mm:ss');
  ss.getSheetByName(SH.L).getRange('C2:I').setNumberFormat('@');

  if (!P.getProperty('SECRET')) P.setProperty('SECRET', Utilities.getUuid() + Utilities.getUuid());
  ['C', 'U', 'G', 'A'].forEach(cacheDrop_);
  _users = null; _cfg = null; _list = null;

  if (!loadUsers_().length) {
    var all = DEF.steps.map(function (s) { return s.key; });
    [
      ['admin', 'Admin', 'admin', 'Admin@123', []],
      ['pc01', 'PC01', 'pc', 'Pc01@123', all],
      ['harish', 'Harish Ji', 'user', 'Harish@123', ['TRACK']],
      ['kuldeep', 'Kuldeep Ji', 'user', 'Kuldeep@123', ['CHARGES']],
      ['dhruv', 'Dhruv Ji', 'user', 'Dhruv@123', ['DOCS']]
    ].forEach(function (d) { saveUser_(newUserObj_(d[0], d[1], d[2], d[3], d[4], '')); });
  }
  var cfg = loadCfg_();
  if (!cfg.rules) setCfg_('rules', DEF.defaultRules);
  if (!cfg.settings) setCfg_('settings', { remindBeforeHours: 2, escalateRepeatHours: 4, quietStart: 22, quietEnd: 7, notifyAdminOnDelay: true, nudgeAssigneeOnOverdue: true });
  if (!cfg.reminders) setCfg_('reminders', []);
  uploadFolder_();
  installTriggers_();
  bumpVersion_();
  P.setProperty('FULL', String(Date.now()));
  return 'Container Tracker is ready. Deploy ▸ New deployment ▸ Web app.';
}

function ensureSheet_(name, headers, appendMissing) {
  var ss = ss_(), sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getMaxColumns() < headers.length) sh.insertColumnsAfter(sh.getMaxColumns(), headers.length - sh.getMaxColumns());
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else if (appendMissing) {
    var have = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    var miss = headers.filter(function (h) { return have.indexOf(h) < 0; });
    if (miss.length) {
      if (sh.getMaxColumns() < have.length + miss.length) sh.insertColumnsAfter(sh.getMaxColumns(), have.length + miss.length - sh.getMaxColumns());
      sh.getRange(1, have.length + 1, 1, miss.length).setValues([miss]);
    }
  }
  var n = sh.getLastColumn();
  sh.getRange(1, 1, 1, n).setFontWeight('bold').setBackground('#232220').setFontColor('#F2C14E').setWrap(true);
  sh.setFrozenRows(1);
  if (name === SH.C) sh.setFrozenColumns(2);
  return sh;
}

function colA1_(n) { var s = ''; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }

function uploadFolder_() {
  var P = PropertiesService.getScriptProperties(), id = P.getProperty('FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  var f = DriveApp.createFolder('Container Tracker — Uploads');
  P.setProperty('FOLDER_ID', f.getId());
  return f;
}

function installTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'cronTick') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('cronTick').timeBased().everyMinutes(5).create();
}

/** Simple trigger: someone edited the sheet by hand → drop caches so the app reloads it. */
function onEdit(e) {
  try {
    var n = e && e.range && e.range.getSheet().getName();
    if ([SH.C, SH.U, SH.G, SH.A].indexOf(n) < 0) return;
    CacheService.getScriptCache().removeAll(['C', 'U', 'G', 'A']);
    var P = PropertiesService.getScriptProperties(), now = String(Date.now());
    P.setProperties({ VER: now, FULL: now });
  } catch (err) {}
}

/* ================================================================ auth */

function hash_(pw, salt) {
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + pw, Utilities.Charset.UTF_8));
}
function newUserObj_(username, name, role, pw, steps, mobile) {
  var salt = Utilities.getUuid().slice(0, 12);
  return {
    id: 'U' + Utilities.getUuid().slice(0, 8).toUpperCase(), username: username.toLowerCase(), name: name, role: role,
    mobile: mobile || '', telegramChatId: '', steps: steps || [], active: true, mustChange: true,
    passHash: hash_(pw, salt), salt: salt, linkCode: '', createdAt: new Date()
  };
}
function sign_(s) {
  var sec = PropertiesService.getScriptProperties().getProperty('SECRET');
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(s, sec)).replace(/=+$/, '');
}
function makeToken_(u) {
  var exp = Date.now() + 30 * 864e5, body = u.id + '.' + exp;
  return body + '.' + sign_(body + '.' + String(u.passHash).slice(0, 10));
}
function auth_(token) {
  var p = String(token || '').split('.');
  if (p.length !== 3 || Number(p[1]) < Date.now()) throw new Error('AUTH');
  var u = userById_(p[0]);
  if (!u || !u.active) throw new Error('AUTH');
  if (sign_(p[0] + '.' + p[1] + '.' + String(u.passHash).slice(0, 10)) !== p[2]) throw new Error('AUTH');
  return u;
}
function isAdmin_(u) { return u.role === 'admin'; }
function can_(u, stepKey) { return isAdmin_(u) || u.steps.indexOf(stepKey) >= 0; }
function need_(ok, msg) { if (!ok) throw new Error(msg || 'You are not allowed to do this.'); }

function withLock_(fn) {
  var l = LockService.getScriptLock();
  if (!l.tryLock(25000)) throw new Error('Server is busy — please try again.');
  try {
    _list = null;
    var r = fn();
    flushLog_();
    SpreadsheetApp.flush();
    return r;
  } catch (e) { _logBuf = []; throw e; }
  finally { l.releaseLock(); }
}

/* ================================================================ session API */

function login(username, password) {
  var u = userByName_(username);
  if (!u || !u.active || hash_(String(password || ''), u.salt) !== u.passHash) throw new Error('Invalid username or password.');
  var token = makeToken_(u);
  return { token: token, boot: boot(token) };
}

/** App open. With `since` (browser already has a cached snapshot) only the delta is returned. */
function boot(token, since) {
  var u = auth_(token);
  countVisit_(u);
  if (since) { var r = sync(token, since); if (!r.same) return r; }
  return payload_(u, !since, since);
}

/** Cheap poll: returns nothing heavy if nothing changed since `since`. */
function sync(token, since) {
  var u = auth_(token), P = PropertiesService.getScriptProperties(), now = Date.now(); // stamp BEFORE reading VER, or a write landing in between is never sent
  var ver = Number(P.getProperty('VER') || 0), full = Number(P.getProperty('FULL') || 0);
  since = Number(since) || 0;
  if (ver && ver < since) return { same: true, now: now };
  if (full >= since) return payload_(u, true);
  return payload_(u, false, since);
}

function payload_(u, full, since) {
  var now = Date.now(), list = loadContainers_();
  // 60s overlap: updatedAt is stamped before a (possibly slow) sheet write finishes
  var cs = full ? list : list.filter(function (c) { return (c.updatedAt || 0) >= since - 60000; });
  var cut = now - 14 * 864e5;
  return {
    full: !!full, now: now,
    me: { id: u.id, username: u.username, name: u.name, role: u.role, mobile: u.mobile, steps: u.steps, tg: !!u.telegramChatId, mustChange: u.mustChange },
    users: loadUsers_().map(publicUser_),
    rules: getRules_(),
    settings: getSettings_(),
    containers: cs,
    approvals: loadApprovals_().filter(function (a) { return a.status === 'PENDING' || (a.decidedAt || 0) > cut || (a.createdAt || 0) > cut; }),
    url: appUrl_()
  };
}

function countVisit_(u) {
  var l = LockService.getScriptLock();
  if (!l.tryLock(1500)) return;
  try {
    var P = PropertiesService.getScriptProperties(), k = 'VIS_' + Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
    var m = JSON.parse(P.getProperty(k) || '{}');
    m[u.id] = (m[u.id] || 0) + 1;
    P.setProperty(k, JSON.stringify(m));
  } finally { l.releaseLock(); }
}

/* ================================================================ pipeline API */

function norm_(key, v) {
  var f = DEF.fields[key];
  if (!f) throw new Error('Unknown field ' + key);
  if (v === null || v === undefined) return '';
  if (typeof v === 'string') v = v.trim();
  if (v === '') return '';
  switch (f.t) {
    case 'date': case 'datetime':
      var t = typeof v === 'number' ? v : new Date(v).getTime();
      if (isNaN(t)) throw new Error('Invalid date for ' + f.l);
      return t;
    case 'money':
      var n = Number(String(v).replace(/[,₹\s]/g, ''));
      if (isNaN(n)) throw new Error('Invalid amount for ' + f.l);
      return n;
    case 'int':
      var i = parseInt(v, 10);
      if (isNaN(i)) throw new Error('Invalid number for ' + f.l);
      return i;
    case 'select':
      if (DEF.locations.indexOf(String(v)) < 0) throw new Error('Pick a valid ' + f.l);
      return String(v);
    case 'file':
      v = String(v);
      if (!/^https:\/\/\S+$/.test(v)) throw new Error(f.l + ' must be an https:// link.'); // blocks javascript: links (stored XSS)
      return v;
    default:
      v = String(v);
      return f.up ? v.toUpperCase() : v;
  }
}
function same_(a, b) { return String(a === undefined ? '' : a) === String(b === undefined ? '' : b); }
function label_(c) { return c.containerNo || ('#' + c.seq + ' of ' + (c.numContainers || '?') + ' · ' + (c.bookingNo || '')); }

function createBooking(token, d, force) {
  var u = auth_(token);
  need_(can_(u, 'S1'), 'Only users assigned to “Booking Details” can create bookings.');
  var vals = {};
  DEF.steps[0].fields.forEach(function (k) { vals[k] = norm_(k, d[k]); if (DEF.fields[k].req && vals[k] === '') throw new Error(DEF.fields[k].l + ' is required.'); });
  need_(vals.numContainers >= 1 && vals.numContainers <= 200, 'No. of containers must be between 1 and 200.');
  return withLock_(function () {
    var list = loadContainers_();
    if (!force) {
      var dup = list.filter(function (c) { return c.bookingNo === vals.bookingNo && c.status !== 'CANCELLED'; });
      if (dup.length) return { confirm: 'Booking ' + vals.bookingNo + ' already exists with ' + dup.length + ' container(s). Create another booking with the same number?' };
    }
    var now = Date.now(), P = PropertiesService.getScriptProperties();
    var bSeq = Number(P.getProperty('SEQ_B') || 0) + 1, cSeq = Number(P.getProperty('SEQ_C') || 0);
    var bookingId = 'B' + ('00000' + bSeq).slice(-5), made = [];
    for (var i = 1; i <= vals.numContainers; i++) {
      cSeq++;
      var c = { id: 'CT' + ('000000' + cSeq).slice(-6), bookingId: bookingId, seq: i, status: 'ACTIVE', createdAt: now, createdBy: u.name, s1_at: now, s1_by: u.name };
      Object.keys(vals).forEach(function (k) { c[k] = vals[k]; });
      made.push(c);
    }
    P.setProperties({ SEQ_B: String(bSeq), SEQ_C: String(cSeq) });
    saveContainers_(made);
    made.forEach(function (c) { log_('BOOKING', c, u, 'Booking created (' + c.seq + ' of ' + c.numContainers + ')'); });
    queueTask_('S2', made);
    return { containers: made };
  });
}

function addContainer(token, bookingId) {
  var u = auth_(token);
  need_(can_(u, 'S1'));
  return withLock_(function () {
    var list = loadContainers_(), sib = list.filter(function (c) { return c.bookingId === bookingId; });
    need_(sib.length, 'Booking not found.');
    var P = PropertiesService.getScriptProperties(), cSeq = Number(P.getProperty('SEQ_C') || 0) + 1, now = Date.now();
    var src = sib[0], n = sib.length + 1;
    var c = { id: 'CT' + ('000000' + cSeq).slice(-6), bookingId: bookingId, seq: n, status: 'ACTIVE', createdAt: now, createdBy: u.name, s1_at: now, s1_by: u.name };
    DEF.steps[0].fields.forEach(function (k) { c[k] = src[k]; });
    P.setProperty('SEQ_C', String(cSeq));
    sib.forEach(function (s) { s.numContainers = n; });
    c.numContainers = n;
    saveContainers_(sib.concat([c]));
    log_('BOOKING', c, u, 'Container added to booking (now ' + n + ')');
    queueTask_('S2', [c]);
    return { containers: sib.concat([c]) };
  });
}

function completeStep(token, id, n, values, applyBooking, force) {
  var u = auth_(token);
  n = Number(n);
  var step = DEF.steps[n - 1];
  need_(step && n > 1, 'Invalid step.');
  need_(can_(u, step.key), 'You are not assigned to “' + step.name + '”.');
  values = values || {};
  var vals = {};
  step.fields.forEach(function (k) {
    vals[k] = norm_(k, values[k]);
    if (DEF.fields[k].req && vals[k] === '') throw new Error(DEF.fields[k].l + ' is required.');
  });
  return withLock_(function () {
    var list = loadContainers_(), c = findContainer_(id), rules = getRules_(), now = Date.now();
    need_(c.status === 'ACTIVE', 'This container is ' + c.status + '.');
    var st = ctState(c, rules, now);
    need_(st.cur === n, 'This step is not open. Current step: ' + (st.cur ? DEF.steps[st.cur - 1].name : 'none') + '.');
    if (vals.containerNo && !force) {
      var dup = list.filter(function (o) { return o.id !== c.id && o.containerNo === vals.containerNo && o.status === 'ACTIVE'; })[0];
      if (dup) return { confirm: 'Container ' + vals.containerNo + ' is already active under booking ' + dup.bookingNo + '. Save anyway?' };
    }
    var targets = [c];
    if (applyBooking && step.bulk) {
      list.forEach(function (o) {
        if (o.id !== c.id && o.bookingId === c.bookingId && o.status === 'ACTIVE' && ctState(o, rules, now).cur === n) targets.push(o);
      });
    }
    targets.forEach(function (t) {
      var due = ctDue(t, n, rules);
      Object.keys(vals).forEach(function (k) { if (vals[k] === '') delete t[k]; else t[k] = vals[k]; });
      t['s' + n + '_at'] = now; t['s' + n + '_by'] = u.name;
      if (due) t['s' + n + '_due'] = due;
      log_('STEP', t, u, 'Completed: ' + step.name + (due && now > due ? ' (LATE)' : ''), '', Object.keys(vals).filter(function (k) { return vals[k] !== ''; })
        .map(function (k) { return DEF.fields[k].l + ': ' + fmtVal_(vals[k], k); }).join(' | '));
    });
    saveContainers_(targets);
    if (n < 10) queueTask_('S' + (n + 1), targets);
    if (n === 7) queueTask_('DOCS', targets);
    return { containers: targets };
  });
}

/** Edit fields. Empty → saved directly. Already filled → sent to admin for approval (admins save directly). */
function saveFields(token, id, changes, reason) {
  var u = auth_(token);
  return withLock_(function () { return saveFieldsLocked_(u, id, changes, reason); });
}

function saveFieldsLocked_(u, id, changes, reason) {
  var list = loadContainers_(), c = findContainer_(id), changed = {}, made = [];
  var admin = isAdmin_(u), pend = loadApprovals_();
  need_(c.status === 'ACTIVE' || admin, 'This container is ' + c.status + ' — only admin can edit it.');
  Object.keys(changes || {}).forEach(function (k) {
    var f = DEF.fields[k];
    need_(f && !f.ro, 'Field cannot be edited: ' + k);
    need_(can_(u, f.own), 'You are not allowed to edit “' + f.l + '”.');
    var sn = /^S(\d+)$/.exec(f.own);
    if (sn && !admin) need_(c['s' + sn[1] + '_at'], 'Complete “' + DEF.steps[sn[1] - 1].name + '” first.');
    var nv = norm_(k, changes[k]), ov = c[k] === undefined ? '' : c[k];
    if (same_(nv, ov)) return;
    if (ov === '' || admin) {
      var targets = f.own === 'S1' ? list.filter(function (o) { return o.bookingId === c.bookingId; }) : [c];
      targets.forEach(function (t) {
        var old = t[k] === undefined ? '' : t[k];
        if (nv === '') delete t[k]; else t[k] = nv;
        changed[t.id] = t;
        log_('EDIT', t, u, f.l + (admin && ov !== '' ? ' (admin edit)' : ''), old, nv, k);
      });
    } else {
      var a = pend.filter(function (x) { return x.status === 'PENDING' && x.containerId === c.id && x.field === k; })[0] || {};
      a.id = a.id || 'R' + Date.now().toString(36).toUpperCase() + Math.floor(Math.random() * 1e3);
      a.createdAt = Date.now(); a.containerId = c.id; a.containerLabel = label_(c); a.bookingNo = c.bookingNo || '';
      a.field = k; a.fieldLabel = f.l; a.oldValue = ov; a.newValue = nv; a.reason = reason || '';
      a.requestedBy = u.name; a.status = 'PENDING'; a.decidedBy = ''; a.decidedAt = ''; a.note = '';
      saveApproval_(a);
      made.push(a);
      log_('REQUEST', c, u, 'Change requested: ' + f.l + (reason ? ' — ' + reason : ''), ov, nv, k);
    }
  });
  // mark documents complete
  if (!c.docs_at && c.sbNo && c.sbPhoto && c.blNo && c.blPhoto) { c.docs_at = Date.now(); c.docs_by = u.name; changed[c.id] = c; log_('STEP', c, u, 'Shipping Bill & BL documents completed'); }
  var out = Object.keys(changed).map(function (k) { return changed[k]; });
  saveContainers_(out);
  if (made.length) queueApprovalNotice_(made, u);
  return { containers: out, approvals: made };
}

function uploadFile(token, id, key, fileName, mime, b64, reason) {
  var u = auth_(token), f = DEF.fields[key];
  need_(f && f.t === 'file', 'Not a file field.');
  need_(can_(u, f.own), 'You are not allowed to upload “' + f.l + '”.');
  var c = findContainer_(id);
  var ext = (String(fileName).match(/\.(\w{1,5})$/) || [, mime === 'application/pdf' ? 'pdf' : 'jpg'])[1];
  var name = (c.containerNo || c.id) + '_' + key + '_' + Utilities.formatDate(new Date(), TZ, 'yyyyMMdd_HHmmss') + '.' + ext;
  var file = uploadFolder_().createFile(Utilities.newBlob(Utilities.base64Decode(b64), mime || 'application/octet-stream', name));
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
  var ch = {}; ch[key] = file.getUrl();
  return withLock_(function () { return saveFieldsLocked_(u, id, ch, reason); });
}

function getHistory(token, id) {
  auth_(token);
  return historyFor_(String(id));
}

/* ================================================================ approvals */

function decideApproval(token, apprId, approve, note) {
  var u = auth_(token);
  need_(isAdmin_(u), 'Only admin can approve changes.');
  return withLock_(function () {
    var a = loadApprovals_().filter(function (x) { return x.id === apprId; })[0];
    need_(a && a.status === 'PENDING', 'This request was already handled.');
    var out = [];
    if (approve) {
      var list = loadContainers_(), c = findContainer_(a.containerId), f = DEF.fields[a.field];
      var targets = f.own === 'S1' ? list.filter(function (o) { return o.bookingId === c.bookingId; }) : [c];
      targets.forEach(function (t) {
        var old = t[a.field];
        if (a.newValue === '' || a.newValue === null) delete t[a.field]; else t[a.field] = a.newValue;
        log_('APPROVED', t, u, f.l + ' change by ' + a.requestedBy + ' approved', old, a.newValue, a.field);
      });
      if (!c.docs_at && c.sbNo && c.sbPhoto && c.blNo && c.blPhoto) { c.docs_at = Date.now(); c.docs_by = a.requestedBy; }
      saveContainers_(targets);
      out = targets;
    } else {
      log_('REJECTED', findContainer_(a.containerId), u, a.fieldLabel + ' change by ' + a.requestedBy + ' rejected' + (note ? ' — ' + note : ''), a.oldValue, a.newValue, a.field);
    }
    a.status = approve ? 'APPROVED' : 'REJECTED'; a.decidedBy = u.name; a.decidedAt = Date.now(); a.note = note || '';
    saveApproval_(a);
    var req = loadUsers_().filter(function (x) { return x.name === a.requestedBy; })[0];
    if (req && req.telegramChatId) queueMsg_(req.telegramChatId, (approve ? '✅' : '❌') + ' Your change to <b>' + escH_(a.fieldLabel) + '</b> on ' + escH_(a.containerLabel) + ' was <b>' + a.status.toLowerCase() + '</b> by ' + escH_(u.name) + (note ? '\n“' + escH_(note) + '”' : ''));
    return { containers: out, approval: a };
  });
}

function withdrawApproval(token, apprId) {
  var u = auth_(token);
  return withLock_(function () {
    var a = loadApprovals_().filter(function (x) { return x.id === apprId; })[0];
    need_(a && a.status === 'PENDING', 'Request not pending.');
    need_(a.requestedBy === u.name || isAdmin_(u));
    a.status = 'WITHDRAWN'; a.decidedBy = u.name; a.decidedAt = Date.now();
    saveApproval_(a);
    return { approval: a };
  });
}

/* ================================================================ admin: container status */

function setStatus(token, id, status, note, force) {
  var u = auth_(token);
  need_(isAdmin_(u), 'Only admin can close or cancel a container.');
  need_(['ACTIVE', 'CLOSED', 'CANCELLED'].indexOf(status) >= 0, 'Bad status.');
  if (status === 'CANCELLED') need_(String(note || '').trim(), 'Give a reason for cancelling.');
  return withLock_(function () {
    var c = findContainer_(id);
    if (status === 'CLOSED' && !force) {
      var w = [];
      if (!c.s10_at) w.push('pipeline is not complete (no Gate-In)');
      if (!(c.sbNo && c.sbPhoto && c.blNo && c.blPhoto)) w.push('Shipping Bill / BL documents are incomplete');
      if (w.length) return { confirm: 'This container ' + w.join(' and ') + '. Close it anyway?' };
    }
    var old = c.status;
    c.status = status; c.statusAt = Date.now(); c.statusBy = u.name; c.statusNote = note || '';
    saveContainers_([c]);
    log_('STATUS', c, u, 'Status → ' + status + (note ? ' — ' + note : ''), old, status);
    if (status !== 'ACTIVE') PropertiesService.getScriptProperties().deleteProperty('NF_' + c.id);
    return { containers: [c] };
  });
}

/* ================================================================ admin: users / workflow / settings */

function adminData(token) {
  var u = auth_(token);
  need_(isAdmin_(u));
  var P = PropertiesService.getScriptProperties();
  return {
    users: loadUsers_().map(function (x) {
      return { id: x.id, username: x.username, name: x.name, role: x.role, mobile: x.mobile, telegramChatId: x.telegramChatId, steps: x.steps, active: x.active, mustChange: x.mustChange };
    }),
    rules: getRules_(), settings: getSettings_(),
    tgSet: !!P.getProperty('TG_TOKEN'), url: appUrl_(),
    triggers: ScriptApp.getProjectTriggers().filter(function (t) { return t.getHandlerFunction() === 'cronTick'; }).length,
    sheetUrl: ss_().getUrl()
  };
}

function saveUserAdmin(token, d) {
  var me = auth_(token);
  need_(isAdmin_(me));
  return withLock_(function () {
    _users = null;
    var uname = String(d.username || '').toLowerCase().trim();
    need_(/^[a-z0-9._-]{3,}$/.test(uname), 'Username: at least 3 characters (letters, numbers, . _ -).');
    need_(String(d.name || '').trim(), 'Name is required.');
    need_(DEF.roles[d.role], 'Pick a role.');
    var clash = userByName_(uname);
    var u = d.id ? userById_(d.id) : null;
    if (clash && (!u || clash.id !== u.id)) throw new Error('Username already taken.');
    if (!u) {
      need_(String(d.password || '').length >= 6, 'Password must be at least 6 characters.');
      u = newUserObj_(uname, d.name, d.role, d.password, [], d.mobile);
    } else if (d.password) {
      need_(String(d.password).length >= 6, 'Password must be at least 6 characters.');
      u.salt = Utilities.getUuid().slice(0, 12); u.passHash = hash_(d.password, u.salt); u.mustChange = true;
    }
    if (u.id === me.id) { need_(d.role === 'admin' && d.active !== false, 'You cannot remove your own admin access.'); }
    u.username = uname; u.name = String(d.name).trim(); u.role = d.role;
    u.mobile = String(d.mobile || '').trim(); u.telegramChatId = String(d.telegramChatId || '').trim();
    u.steps = (d.steps || []).filter(function (s) { return DEF.stepKeys.indexOf(s) >= 0; });
    u.active = d.active !== false;
    saveUser_(u);
    log_('ADMIN', null, me, 'User saved: ' + u.username + ' (' + u.role + ', steps: ' + u.steps.join(',') + ')');
    return adminData(token);
  });
}

function setStepAssignees(token, stepKey, userIds) {
  var me = auth_(token);
  need_(isAdmin_(me));
  need_(DEF.stepKeys.indexOf(stepKey) >= 0, 'Unknown step.');
  return withLock_(function () {
    _users = null;
    loadUsers_().forEach(function (u) {
      var has = u.steps.indexOf(stepKey) >= 0, want = userIds.indexOf(u.id) >= 0;
      if (has === want) return;
      u.steps = want ? u.steps.concat([stepKey]) : u.steps.filter(function (s) { return s !== stepKey; });
      saveUser_(u);
    });
    log_('ADMIN', null, me, 'Assignees for ' + stepKey + ' set');
    return adminData(token);
  });
}

function saveRules(token, rules) {
  var me = auth_(token);
  need_(isAdmin_(me));
  var out = {};
  Object.keys(DEF.defaultRules).forEach(function (k) {
    var r = rules[k] || DEF.defaultRules[k];
    var src = r.src === 'field' && k !== 'DOCS' ? 'field' : 'after';
    if (src === 'field') need_(DEF.deadlineFields.indexOf(r.field) >= 0, 'Pick a date field for ' + k);
    out[k] = { src: src, field: src === 'field' ? r.field : '', hours: Number(r.hours) || 0 };
  });
  return withLock_(function () {
    setCfg_('rules', out);
    bumpVersion_();
    PropertiesService.getScriptProperties().setProperty('FULL', String(Date.now()));
    log_('ADMIN', null, me, 'Deadline rules updated');
    return adminData(token);
  });
}

function saveSettings(token, s, tgToken) {
  var me = auth_(token);
  need_(isAdmin_(me));
  return withLock_(function () {
    var cur = loadCfg_().settings || {};
    ['remindBeforeHours', 'escalateRepeatHours', 'quietStart', 'quietEnd'].forEach(function (k) { if (s[k] !== undefined && s[k] !== '') cur[k] = Number(s[k]); });
    ['notifyAdminOnDelay', 'nudgeAssigneeOnOverdue'].forEach(function (k) { if (s[k] !== undefined) cur[k] = !!s[k]; });
    if (tgToken) {
      var res = JSON.parse(UrlFetchApp.fetch('https://api.telegram.org/bot' + tgToken + '/getMe', { muteHttpExceptions: true }).getContentText());
      need_(res.ok, 'Telegram rejected this bot token.');
      PropertiesService.getScriptProperties().setProperty('TG_TOKEN', tgToken);
      try { UrlFetchApp.fetch('https://api.telegram.org/bot' + tgToken + '/deleteWebhook', { muteHttpExceptions: true }); } catch (e) {}
      cur.botUsername = res.result.username;
    }
    setCfg_('settings', cur);
    bumpVersion_();
    log_('ADMIN', null, me, 'Settings updated');
    return adminData(token);
  });
}

function testTelegram(token, chatId) {
  var u = auth_(token);
  var id = (isAdmin_(u) && chatId) || u.telegramChatId; // non-admins can only test their own chat
  need_(id, 'No Telegram chat linked.');
  var r = tgSend_(id, '✅ <b>Container Tracker</b> test message for ' + escH_(u.name) + '.');
  need_(r && r.ok, 'Telegram error: ' + (r && r.description || 'bot token not set'));
  return true;
}

function adminInstallTriggers(token) { need_(isAdmin_(auth_(token))); installTriggers_(); return adminData(token); }
function adminRunCheck(token) { need_(isAdmin_(auth_(token))); return cronTick(true); }
function adminClearCache(token) {
  need_(isAdmin_(auth_(token)));
  ['C', 'U', 'G', 'A'].forEach(cacheDrop_);
  var now = String(Date.now());
  PropertiesService.getScriptProperties().setProperties({ VER: now, FULL: now });
  _list = null; _users = null; _cfg = null;
  return true;
}

function getVisits(token, days) {
  var me = auth_(token);
  need_(me.role === 'admin' || me.role === 'pc');
  days = Math.min(Number(days) || 14, 60);
  var P = PropertiesService.getScriptProperties().getProperties(), out = [];
  for (var i = days - 1; i >= 0; i--) {
    var d = Utilities.formatDate(new Date(Date.now() - i * 864e5), TZ, 'yyyy-MM-dd');
    out.push({ d: d, v: JSON.parse(P['VIS_' + d] || '{}') });
  }
  return out;
}

/* ================================================================ profile */

function changePassword(token, oldPw, newPw) {
  var u = auth_(token);
  need_(hash_(String(oldPw || ''), u.salt) === u.passHash, 'Current password is wrong.');
  need_(String(newPw || '').length >= 6, 'New password must be at least 6 characters.');
  return withLock_(function () {
    _users = null;
    var x = userById_(u.id);
    x.salt = Utilities.getUuid().slice(0, 12); x.passHash = hash_(newPw, x.salt); x.mustChange = false;
    saveUser_(x);
    return { token: makeToken_(x) };
  });
}

function tgLinkCode(token) {
  var u = auth_(token), s = getSettings_();
  need_(s.botUsername, 'Admin has not connected a Telegram bot yet.');
  return withLock_(function () {
    _users = null;
    var x = userById_(u.id);
    x.linkCode = String(Math.floor(100000 + Math.random() * 900000));
    saveUser_(x);
    return { bot: s.botUsername, code: x.linkCode, url: 'https://t.me/' + s.botUsername + '?start=' + x.linkCode };
  });
}

function tgCheck(token) {
  var u = auth_(token);
  pollTelegram_();
  _users = null; cacheDrop_('U');
  return { linked: !!userById_(u.id).telegramChatId };
}

/* ================================================================ reminders (self reminders) */

function listReminders(token) {
  var u = auth_(token);
  return (loadCfg_().reminders || []).filter(function (r) { return r.owner === u.id; });
}
function saveReminder(token, r) {
  var u = auth_(token);
  need_(String(r.text || '').trim(), 'Write what to remind you about.');
  var at = Number(r.at);
  need_(at > 0, 'Pick a date & time.');
  return withLock_(function () {
    _cfg = null; cacheDrop_('G');
    var all = loadCfg_().reminders || [];
    var ex = r.id ? all.filter(function (x) { return x.id === r.id && x.owner === u.id; })[0] : null;
    var o = ex || { id: 'M' + Date.now().toString(36), owner: u.id, createdAt: Date.now() };
    o.text = String(r.text).trim(); o.at = at; o.repeat = ['none', 'daily', 'weekly'].indexOf(r.repeat) >= 0 ? r.repeat : 'none';
    o.containerId = r.containerId || ''; o.active = true;
    if (!ex) all.push(o);
    setCfg_('reminders', all.filter(function (x) { return x.active || (x.lastSent || 0) > Date.now() - 30 * 864e5; }));
    return listReminders(token);
  });
}
function deleteReminder(token, id) {
  var u = auth_(token);
  return withLock_(function () {
    _cfg = null; cacheDrop_('G');
    setCfg_('reminders', (loadCfg_().reminders || []).filter(function (x) { return !(x.id === id && x.owner === u.id); }));
    return listReminders(token);
  });
}

/** Called by the browser right after a save so Telegram sends never slow down the save itself. */
function flushNotify(token) { auth_(token); flushQueue_(); return true; }
