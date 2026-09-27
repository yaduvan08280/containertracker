/**
 * Container Tracker — Telegram notifications and the 5-minute scheduler.
 * Messages are queued in Script Properties (NQ_*) and sent in parallel with fetchAll,
 * so a user's save never waits on Telegram.
 */

function tgToken_() { return PropertiesService.getScriptProperties().getProperty('TG_TOKEN') || ''; }
function escH_(s) { return String(s === undefined || s === null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
function fmtT_(ms) { return ms ? Utilities.formatDate(new Date(Number(ms)), TZ, 'dd MMM yyyy, HH:mm') : '—'; }
function dur_(ms) {
  ms = Math.abs(ms);
  var m = Math.round(ms / 60000), d = Math.floor(m / 1440), h = Math.floor((m % 1440) / 60), mm = m % 60;
  return ((d ? d + 'd ' : '') + (h ? h + 'h ' : '') + (!d ? mm + 'm' : '')).trim();
}

function tgSend_(chatId, text) {
  var tok = tgToken_();
  if (!tok || !chatId) return null;
  var r = UrlFetchApp.fetch('https://api.telegram.org/bot' + tok + '/sendMessage', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    payload: JSON.stringify({ chat_id: chatId, text: text, parse_mode: 'HTML', disable_web_page_preview: true })
  });
  try { return JSON.parse(r.getContentText()); } catch (e) { return null; }
}

function queueMsg_(chatId, text) {
  if (!chatId || !text) return;
  // 2500 chars stays under the 9KB property limit even with multi-byte text; a failed alert must never fail the save that queued it
  try { PropertiesService.getScriptProperties().setProperty('NQ_' + Date.now() + '_' + Math.floor(Math.random() * 1e6), JSON.stringify({ c: String(chatId), t: String(text).slice(0, 2500) })); }
  catch (e) { console.warn('queueMsg_', e); }
}

/** Send everything queued. Messages to the same chat are merged to cut noise. */
function flushQueue_() {
  var tok = tgToken_(), P = PropertiesService.getScriptProperties();
  var all = P.getProperties(), keys = Object.keys(all).filter(function (k) { return k.indexOf('NQ_') === 0; }).sort();
  if (!keys.length) return 0;
  var l = LockService.getScriptLock();
  if (!l.tryLock(8000)) return 0; // another flush / write is running — the next flush picks these up
  try {
    all = P.getProperties();
    keys = Object.keys(all).filter(function (k) { return k.indexOf('NQ_') === 0; }).sort();
    keys.forEach(function (k) { P.deleteProperty(k); });
  } finally { l.releaseLock(); }
  if (!tok) return 0;
  var byChat = {};
  keys.forEach(function (k) {
    try { var m = JSON.parse(all[k]); (byChat[m.c] = byChat[m.c] || []).push(m.t); } catch (e) {}
  });
  var reqs = [];
  Object.keys(byChat).forEach(function (chat) {
    var buf = '';
    byChat[chat].forEach(function (t) {
      if (buf && (buf.length + t.length) > 3800) { reqs.push(req(chat, buf)); buf = ''; }
      buf += (buf ? '\n\n━━━━━━━━━━\n\n' : '') + t;
    });
    if (buf) reqs.push(req(chat, buf));
  });
  function req(chat, text) {
    return { url: 'https://api.telegram.org/bot' + tok + '/sendMessage', method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      payload: JSON.stringify({ chat_id: chat, text: text, parse_mode: 'HTML', disable_web_page_preview: true }) };
  }
  for (var i = 0; i < reqs.length; i += 25) UrlFetchApp.fetchAll(reqs.slice(i, i + 25));
  return reqs.length;
}

function assignees_(stepKey) {
  return loadUsers_().filter(function (u) { return u.active && u.steps.indexOf(stepKey) >= 0; });
}
function stepName_(stepKey) {
  if (stepKey === 'DOCS') return 'Shipping Bill & BL upload';
  var n = Number(stepKey.slice(1));
  return DEF.steps[n - 1].name;
}
function bookingLine_(c) {
  return 'Booking <b>' + escH_(c.bookingNo) + '</b> · ' + escH_(c.party || '') + (c.line ? ' · ' + escH_(c.line) : '') + (c.vessel ? ' / ' + escH_(c.vessel) : '');
}
function contList_(cs) {
  var names = cs.map(function (c) { return c.containerNo ? '<code>' + escH_(c.containerNo) + '</code>' : '#' + c.seq; });
  return (cs.length > 1 ? 'Containers (' + cs.length + '): ' : 'Container: ') + names.join(', ');
}
function openLink_(c) { var u = appUrl_(); return u ? '\n<a href="' + u + '?c=' + c.id + '">Open in Container Tracker</a>' : ''; }

/** "New task" message to everyone assigned to stepKey, grouped per booking. */
function queueTask_(stepKey, cs) {
  var who = assignees_(stepKey).filter(function (u) { return u.telegramChatId; });
  if (!who.length || !cs.length) return;
  var rules = getRules_(), now = Date.now(), groups = {};
  cs.forEach(function (c) { (groups[c.bookingId] = groups[c.bookingId] || []).push(c); });
  Object.keys(groups).forEach(function (b) {
    var g = groups[b], dues = g.map(function (c) { var s = ctState(c, rules, now); return stepKey === 'DOCS' ? s.docsDue : s.due; }).filter(Boolean);
    var due = dues.length ? Math.min.apply(null, dues) : null;
    var text = '🔔 <b>New task — ' + escH_(stepName_(stepKey)) + '</b>\n' + bookingLine_(g[0]) + '\n' + contList_(g) +
      '\n⏰ Deadline: <b>' + (due ? fmtT_(due) : 'waiting for date') + '</b>' + openLink_(g[0]);
    who.forEach(function (u) { queueMsg_(u.telegramChatId, text); });
  });
}

function queueApprovalNotice_(reqs, u) {
  var admins = loadUsers_().filter(function (x) { return x.active && x.role === 'admin' && x.telegramChatId; });
  if (!admins.length) return;
  var text = '✏️ <b>Change approval needed</b>\nRequested by ' + escH_(u.name) + '\n' + reqs.map(function (a) {
    return '• ' + escH_(a.containerLabel) + ' — <b>' + escH_(a.fieldLabel) + '</b>: “' + escH_(fmtVal_(a.oldValue, a.field)) + '” → “' + escH_(fmtVal_(a.newValue, a.field)) + '”';
  }).join('\n') + (reqs[0].reason ? '\nReason: ' + escH_(reqs[0].reason) : '');
  admins.forEach(function (a) { queueMsg_(a.telegramChatId, text); });
}

/* ================================================================ scheduler */

/** Time-driven every 5 minutes (see installTriggers_). */
function cronTick(manual) {
  var out = { linked: 0, alerts: 0, reminders: 0, sent: 0 };
  try { out.linked = pollTelegram_(); } catch (e) { console.warn('poll', e); }
  var now = Date.now(), s = getSettings_();
  var hour = Number(Utilities.formatDate(new Date(), TZ, 'H'));
  var quiet = s.quietStart !== s.quietEnd && (s.quietStart > s.quietEnd ? (hour >= s.quietStart || hour < s.quietEnd) : (hour >= s.quietStart && hour < s.quietEnd));
  if (!quiet || manual === true) out.alerts = checkDeadlines_(now, s);
  out.reminders = checkReminders_(now);
  out.sent = flushQueue_();
  if (hour === 3) { try { pruneVisits_(); } catch (e) {} }
  return out;
}

function checkDeadlines_(now, s) {
  var P = PropertiesService.getScriptProperties(), props = P.getProperties(), rules = getRules_();
  var list = loadContainers_(), users = loadUsers_().filter(function (u) { return u.active; });
  var watchers = users.filter(function (u) { return u.telegramChatId && (u.role === 'pc' || (s.notifyAdminOnDelay && u.role === 'admin')); });
  var remindMs = s.remindBeforeHours * 3600000, repeatMs = Math.max(s.escalateRepeatHours, 0.5) * 3600000;
  var events = {}, upd = {}, alive = {};

  list.forEach(function (c) {
    if (c.status !== 'ACTIVE') return;
    alive['NF_' + c.id] = 1;
    var st = ctState(c, rules, now), key = 'NF_' + c.id, state, dirty = false;
    try { state = JSON.parse(props[key] || '{}'); } catch (e) { state = {}; }
    function check(stepKey, due, tag) {
      if (!due) return;
      if (now < due && due - now <= remindMs && state['r' + tag] !== due) {
        push('remind', stepKey, due, c); state['r' + tag] = due; dirty = true;
      }
      if (now > due && now - (state['o' + tag] || 0) >= repeatMs) {
        push('over', stepKey, due, c); state['o' + tag] = now; dirty = true;
      }
    }
    if (st.cur) check('S' + st.cur, st.due, String(st.cur));
    if (st.docsNeeded) check('DOCS', st.docsDue, 'D');
    if (dirty) upd[key] = JSON.stringify(state);
  });
  function push(type, stepKey, due, c) {
    var k = type + '|' + stepKey + '|' + c.bookingId;
    (events[k] = events[k] || { type: type, stepKey: stepKey, due: due, cs: [] }).cs.push(c);
    events[k].due = Math.min(events[k].due, due);
  }

  var n = 0;
  Object.keys(events).forEach(function (k) {
    var ev = events[k], c0 = ev.cs[0], who = assignees_(ev.stepKey), name = stepName_(ev.stepKey);
    if (ev.type === 'remind') {
      var t = '⏰ <b>Reminder — ' + escH_(name) + '</b>\n' + bookingLine_(c0) + '\n' + contList_(ev.cs) +
        '\nDue in <b>' + dur_(ev.due - now) + '</b> (' + fmtT_(ev.due) + ')' + openLink_(c0);
      who.forEach(function (u) { if (u.telegramChatId) queueMsg_(u.telegramChatId, t); });
    } else {
      var call = who.length ? who.map(function (u) { return '• ' + escH_(u.name) + ' — ' + (u.mobile ? '<b>' + escH_(u.mobile) + '</b>' : '<i>no mobile saved</i>'); }).join('\n')
        : '⚠️ Nobody is assigned to this step — assign someone in Admin ▸ Workflow.';
      var alert = '🚨 <b>DELAY — follow-up needed</b>\nStep: <b>' + escH_(name) + '</b>\n' + bookingLine_(c0) + '\n' + contList_(ev.cs) +
        '\nDeadline: ' + fmtT_(ev.due) + ' · <b>overdue by ' + dur_(now - ev.due) + '</b>\n📞 Please call:\n' + call + openLink_(c0);
      watchers.forEach(function (u) { queueMsg_(u.telegramChatId, alert); });
      if (s.nudgeAssigneeOnOverdue) {
        var nudge = '❗ <b>Overdue — ' + escH_(name) + '</b>\n' + bookingLine_(c0) + '\n' + contList_(ev.cs) +
          '\nDeadline was ' + fmtT_(ev.due) + ' (' + dur_(now - ev.due) + ' ago). Please complete it now.' + openLink_(c0);
        who.forEach(function (u) { if (u.telegramChatId) queueMsg_(u.telegramChatId, nudge); });
      }
    }
    n++;
  });
  if (Object.keys(upd).length) P.setProperties(upd);
  Object.keys(props).forEach(function (k) { if (k.indexOf('NF_') === 0 && !alive[k]) P.deleteProperty(k); });
  return n;
}

function checkReminders_(now) {
  var due = (loadCfg_().reminders || []).filter(function (r) { return r.active && r.at <= now; });
  if (!due.length) return 0;
  var l = LockService.getScriptLock();
  if (!l.tryLock(20000)) return 0;
  try {
    _cfg = null; cacheDrop_('G');
    var list = loadCfg_().reminders || [], fired = 0; // re-read inside the lock so we never drop a just-added reminder
    list.forEach(function (r) {
      if (!r.active || r.at > now) return;
      var u = userById_(r.owner);
      if (u && u.telegramChatId) {
        var c = null;
        if (r.containerId) { try { c = findContainer_(r.containerId); } catch (e) {} }
        queueMsg_(u.telegramChatId, '🔔 <b>Your reminder</b>\n' + escH_(r.text) + (c ? '\n' + bookingLine_(c) + '\n' + contList_([c]) + openLink_(c) : ''));
      }
      r.lastSent = now; fired++;
      if (r.repeat === 'daily' || r.repeat === 'weekly') { var step = r.repeat === 'daily' ? 864e5 : 7 * 864e5; while (r.at <= now) r.at += step; }
      else r.active = false;
    });
    if (fired) setCfg_('reminders', list);
    return fired;
  } finally { l.releaseLock(); }
}

/** Reads bot messages; "/start 123456" links a Telegram chat to the user with that code. */
function pollTelegram_() {
  var tok = tgToken_();
  if (!tok) return 0;
  var l = LockService.getScriptLock();
  if (!l.tryLock(15000)) return 0; // one poller at a time, so an update is never handled twice
  try {
    var P = PropertiesService.getScriptProperties(), off = Number(P.getProperty('TG_OFFSET') || 0);
    var res = JSON.parse(UrlFetchApp.fetch('https://api.telegram.org/bot' + tok + '/getUpdates?timeout=0&offset=' + off, { muteHttpExceptions: true }).getContentText());
    if (!res.ok || !res.result.length) return 0;
    var linked = 0;
    res.result.forEach(function (upd) {
      off = upd.update_id + 1;
      var m = upd.message;
      if (!m || !m.text) return;
      var chat = String(m.chat.id), code = (/^\/(?:start|link)\s+(\d{6})/.exec(m.text.trim()) || [])[1];
      if (code) {
        _users = null; cacheDrop_('U');
        var u = loadUsers_().filter(function (x) { return String(x.linkCode) === code; })[0];
        if (u) {
          u.telegramChatId = chat; u.linkCode = '';
          saveUser_(u); linked++;
          tgSend_(chat, '✅ Linked! You will now receive <b>Container Tracker</b> alerts as <b>' + escH_(u.name) + '</b>.');
        } else tgSend_(chat, 'That code is not valid any more. Open Container Tracker ▸ Profile ▸ Link Telegram for a new one.');
      } else if (/^\/start/.test(m.text)) {
        tgSend_(chat, 'Hi! To get alerts, open <b>Container Tracker</b> ▸ Profile ▸ <b>Link Telegram</b> and tap the link shown there.\nYour chat id: <code>' + chat + '</code>');
      }
    });
    P.setProperty('TG_OFFSET', String(off));
    return linked;
  } finally { l.releaseLock(); }
}

/** Drop visit counters older than 120 days so Script Properties stay small. */
function pruneVisits_() {
  var P = PropertiesService.getScriptProperties(), cut = Utilities.formatDate(new Date(Date.now() - 120 * 864e5), TZ, 'yyyy-MM-dd');
  Object.keys(P.getProperties()).forEach(function (k) { if (k.indexOf('VIS_') === 0 && k.slice(4) < cut) P.deleteProperty(k); });
}
