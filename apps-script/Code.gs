/**
 * GLugo To-Do — Google Apps Script backend
 * Keeps your tasks, notes, drawings, photos and PDFs in your own Google Drive
 * (folder "GLugo To-Do"), sends the daily email, and makes a backup every day.
 *
 * Setup (one time, on the Mac in Chrome):
 *  1. script.google.com → New project → paste this whole file over Code.gs → Save.
 *  2. Deploy → New deployment → type "Web app" → Execute as: Me, Who has access: Anyone → Deploy.
 *     Allow access when Google asks. Copy the Web app URL.
 *  3. Back in the editor, pick "setup" in the function menu → Run.
 *     It moves your current list in, turns on the daily email and the daily backup,
 *     and emails you your connection key.
 *  4. Open the app, paste the Web app URL and the key. Do this once per device.
 *
 * The key lives only in this script's settings, never in the app's public code.
 */

const APP_URL = 'https://glugoelectric.github.io/glugo-todo/';
const FOLDER_NAME = 'GLugo To-Do';
const DATA_NAME = 'data.json';
const IMPORT_NAME = 'glugo-todo-import.json';
const KEEP_BACKUPS = 30;
const EMAIL_HOUR = 5; // daily email around 5:45 am (script time zone)

/* ---------------- setup ---------------- */
function setup() {
  const props = PropertiesService.getScriptProperties();
  let key = props.getProperty('KEY');
  if (!key) {
    key = Utilities.getUuid().replace(/-/g, '').slice(0, 20);
    props.setProperty('KEY', key);
  }
  const folder = appFolder_();
  const imported = importOld_(folder);
  dailyBackup();

  ScriptApp.getProjectTriggers().forEach(t => {
    if (['dailyEmail', 'dailyBackup'].includes(t.getHandlerFunction())) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('dailyEmail').timeBased().everyDays(1).atHour(EMAIL_HOUR).nearMinute(45).create();
  ScriptApp.newTrigger('dailyBackup').timeBased().everyDays(1).atHour(2).create();

  let url = '';
  try { url = ScriptApp.getService().getUrl() || ''; } catch (e) {}
  const me = Session.getEffectiveUser().getEmail();
  const connect = url ? APP_URL + '#connect=' + encodeURIComponent(url) + '|' + key : '';
  MailApp.sendEmail({
    to: me,
    subject: 'GLugo To-Do: your connection key',
    htmlBody:
      '<div style="font:16px -apple-system,Arial,sans-serif">' +
      '<p>Your GLugo To-Do key:</p><p style="font:700 22px monospace">' + key + '</p>' +
      (url ? '<p>Connection code (copy all of it):<br><span style="font:13px monospace;word-break:break-all;background:#f2f4f7;display:block;padding:8px;border-radius:6px">' + esc_(url + '|' + key) + '</span></p>' +
        '<p><a href="' + esc_(connect) + '" style="display:inline-block;background:#1f6feb;color:#fff;padding:12px 20px;border-radius:10px;text-decoration:none;font-weight:700">Connect this device</a></p>' +
        '<p style="color:#666;font-size:14px">Tap Connect on each device (iPhone, iPad, Mac). If it doesn\'t connect, open the app and paste the URL and key.</p>' : '') +
      (imported ? '<p style="color:#666;font-size:14px">Moved in ' + imported + ' tasks and notes from the old app.</p>' : '') +
      '<p style="color:#666;font-size:13px">Keep this email private. Anyone with the key and URL can open your list.</p></div>'
  });
  Logger.log('Key: ' + key);
  Logger.log('URL: ' + url);
  Logger.log('Imported: ' + imported);
}

/* ---------------- storage ---------------- */
function appFolder_() {
  const it = DriveApp.getFoldersByName(FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME);
}
function sub_(name) {
  const f = appFolder_(), it = f.getFoldersByName(name);
  return it.hasNext() ? it.next() : f.createFolder(name);
}
function dataFile_() {
  const f = appFolder_(), it = f.getFilesByName(DATA_NAME);
  if (it.hasNext()) return it.next();
  return f.createFile(DATA_NAME, JSON.stringify(empty_()), 'application/json');
}
function empty_() { return { todos: {}, notes: {}, sketches: {}, folders: [], rev: 0 }; }
function read_() {
  const d = JSON.parse(dataFile_().getBlob().getDataAsString() || '{}');
  const e = empty_();
  Object.keys(e).forEach(k => { if (d[k] == null) d[k] = e[k]; });
  return d;
}
function write_(d) {
  d.rev = (d.rev || 0) + 1;
  dataFile_().setContent(JSON.stringify(d));
  return d.rev;
}
function importOld_(folder) {
  const it = DriveApp.getFilesByName(IMPORT_NAME);
  if (!it.hasNext()) return 0;
  const file = it.next();
  const old = JSON.parse(file.getBlob().getDataAsString());
  const d = read_();
  let n = 0;
  ['todos', 'notes', 'sketches'].forEach(k => {
    Object.keys(old[k] || {}).forEach(id => { if (!d[k][id]) { d[k][id] = old[k][id]; if (k !== 'sketches') n++; } });
  });
  (old.folders || []).forEach(f => { if (!d.folders.includes(f)) d.folders.push(f); });
  write_(d);
  file.setName(IMPORT_NAME.replace('.json', ' (moved in).json'));
  file.moveTo(folder);
  return n;
}

/* ---------------- web app ---------------- */
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function keyOk_(k) {
  const key = PropertiesService.getScriptProperties().getProperty('KEY');
  return !!key && k === key;
}
function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'bad request' }); }
  if (!keyOk_(req.k)) return out_({ ok: false, error: 'wrong key' });
  const lock = LockService.getScriptLock();
  try {
    if (['put', 'del', 'folders', 'batch'].includes(req.a)) lock.waitLock(20000);
    switch (req.a) {
      case 'load': { const d = read_(); return out_({ ok: true, data: d }); }
      case 'batch': {
        // ops: [{a:'put', kind, id, rec} | {a:'del', kind, id} | {a:'folders', folders}]
        const d = read_();
        (req.ops || []).forEach(op => applyOp_(d, op));
        const rev = write_(d);
        return out_({ ok: true, rev: rev, data: d });
      }
      case 'upload': {
        const blob = Utilities.newBlob(Utilities.base64Decode(req.b64), req.mime || 'application/octet-stream', req.name || 'file');
        const f = sub_('files').createFile(blob);
        return out_({ ok: true, id: f.getId() });
      }
      case 'file': {
        const f = ownFile_(req.id);
        if (!f) return out_({ ok: false, error: 'not found' });
        const b = f.getBlob();
        return out_({ ok: true, mime: b.getContentType(), name: f.getName(), b64: Utilities.base64Encode(b.getBytes()) });
      }
      case 'trash': {
        const f = ownFile_(req.id);
        if (f) f.setTrashed(true);
        return out_({ ok: true });
      }
      case 'backups': {
        const list = [], it = sub_('Backups').getFiles();
        while (it.hasNext()) { const f = it.next(); list.push({ id: f.getId(), name: f.getName(), t: f.getDateCreated().getTime() }); }
        list.sort((a, b) => b.t - a.t);
        return out_({ ok: true, list: list });
      }
      case 'restore': {
        lock.waitLock(20000);
        const backups = sub_('Backups');
        let src = null; const it = backups.getFiles();
        while (it.hasNext()) { const f = it.next(); if (f.getId() === req.id) src = f; }
        if (!src) return out_({ ok: false, error: 'backup not found' });
        const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HHmm');
        dataFile_().makeCopy('before-restore-' + stamp + '.json', backups); // the current list is kept too
        const d = JSON.parse(src.getBlob().getDataAsString());
        write_(d);
        return out_({ ok: true, data: d });
      }
      default: return out_({ ok: false, error: 'unknown action' });
    }
  } catch (err) {
    return out_({ ok: false, error: String(err && err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}
const KINDS_ = ['todos', 'notes', 'sketches'];
function applyOp_(d, op) {
  if (op.a === 'folders' && Array.isArray(op.folders)) { d.folders = op.folders.map(String).slice(0, 200); return; }
  if (!KINDS_.includes(op.kind) || !op.id) return;
  if (op.a === 'put' && op.rec) d[op.kind][op.id] = op.rec;
  if (op.a === 'del') delete d[op.kind][op.id];
}
function ownFile_(id) {
  if (!id) return null;
  try {
    const f = DriveApp.getFileById(id), files = sub_('files').getId();
    const ps = f.getParents();
    while (ps.hasNext()) if (ps.next().getId() === files) return f;
  } catch (e) {}
  return null;
}

/* Done links in the daily email */
function doGet(e) {
  const p = e && e.parameter || {};
  if (p.a === 'done' && keyOk_(p.k) && p.id) {
    const lock = LockService.getScriptLock(); lock.waitLock(20000);
    let text = '';
    try {
      const d = read_(), t = d.todos[p.id];
      if (t) { t.done = true; t.doneAt = Date.now(); t.updated = Date.now(); text = t.text; write_(d); }
    } finally { lock.releaseLock(); }
    return HtmlService.createHtmlOutput(
      '<meta name="viewport" content="width=device-width,initial-scale=1">' +
      '<div style="font:20px -apple-system,Arial,sans-serif;text-align:center;padding:40px 16px">' +
      (text ? '<div style="font-size:44px">✅</div><p><b>Done:</b> ' + esc_(text) + '</p>' : '<p>That task was already removed.</p>') +
      '<p><a href="' + APP_URL + '" target="_top" style="display:inline-block;background:#1f6feb;color:#fff;padding:14px 22px;border-radius:10px;text-decoration:none;font-weight:700">Open my list</a></p></div>'
    ).setTitle('GLugo To-Do');
  }
  return HtmlService.createHtmlOutput('<p style="font:18px Arial;padding:30px">GLugo To-Do is running.</p>').setTitle('GLugo To-Do');
}

/* ---------------- daily email ---------------- */
function dailyEmail() {
  const d = read_();
  const key = PropertiesService.getScriptProperties().getProperty('KEY');
  let url = ''; try { url = ScriptApp.getService().getUrl() || ''; } catch (e) {}
  const all = Object.keys(d.todos).map(id => Object.assign({ id: id }, d.todos[id]));
  const open = all.filter(t => !t.done).sort((a, b) => (b.urgent ? 1 : 0) - (a.urgent ? 1 : 0) || (a.created || 0) - (b.created || 0));
  const urgent = open.filter(t => t.urgent).length;
  const doneDay = all.filter(t => t.done && t.doneAt > Date.now() - 864e5).length;
  const subject = open.length ? 'GLugo To-Do: ' + open.length + ' open' + (urgent ? ' · ' + urgent + ' urgent' : '') : 'GLugo To-Do: all done 👍';
  const btn = 'display:inline-block;background:#1a7f37;color:#fff;padding:6px 12px;border-radius:8px;text-decoration:none;font-weight:700;font-size:14px';
  const rows = open.map(t =>
    '<tr><td style="padding:10px 8px 10px 0;border-bottom:1px solid #eee;vertical-align:top">' +
    (t.urgent ? '<span style="color:#cf222e;font-weight:700">URGENT · </span>' : '') +
    '<a href="' + APP_URL + '#item-' + encodeURIComponent(t.id) + '" style="color:#1b1f24;font-weight:600;text-decoration:none">' + esc_(t.text) + '</a>' +
    ' <span style="color:#888;font-size:14px">· ' + esc_(t.area || 'Home') + '</span></td>' +
    '<td style="padding:10px 0;border-bottom:1px solid #eee;text-align:right;vertical-align:top;white-space:nowrap">' +
    (url && key ? '<a href="' + url + '?a=done&id=' + encodeURIComponent(t.id) + '&k=' + key + '" style="' + btn + '">✓ Done</a>' : '') +
    '</td></tr>').join('');
  const html = '<div style="font:16px -apple-system,Arial,sans-serif;max-width:560px">' +
    '<p><a href="' + APP_URL + '" style="display:inline-block;background:#1f6feb;color:#fff;padding:12px 20px;border-radius:10px;text-decoration:none;font-weight:700">Open my list</a></p>' +
    (open.length ? '<table style="width:100%;border-collapse:collapse">' + rows + '</table>' : '<p>Nothing open. 👍</p>') +
    (doneDay ? '<p>Done in the last day: ' + doneDay + '</p>' : '') +
    '<p style="color:#888;font-size:13px">Tap ✓ Done to check a task off, or tap the task to open it.</p></div>';
  const text = open.map(t => (t.urgent ? 'URGENT · ' : '') + t.text + ' · ' + (t.area || 'Home')).join('\n') + '\n\n' + APP_URL;
  MailApp.sendEmail({ to: Session.getEffectiveUser().getEmail(), subject: subject, htmlBody: html, body: text });
}

/* ---------------- daily backup ---------------- */
function dailyBackup() {
  const backups = sub_('Backups');
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const it = backups.getFilesByName('data-' + stamp + '.json');
  while (it.hasNext()) it.next().setTrashed(true);
  dataFile_().makeCopy('data-' + stamp + '.json', backups);
  const list = [];
  const all = backups.getFiles();
  while (all.hasNext()) list.push(all.next());
  list.sort((a, b) => b.getDateCreated() - a.getDateCreated());
  list.slice(KEEP_BACKUPS).forEach(f => f.setTrashed(true));
}

function esc_(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }
