/**
 * Athena Run Club – backend (Google Apps Script)
 * The data lives in the Google Sheet this script is attached to.
 *
 * Setup (once):
 *  1. In the Sheet: Extensions → Apps Script, paste this file, Save.
 *  2. Run the function `setup` once (approve the permissions it asks for).
 *  3. Project Settings → Script properties → add  ADMIN_PIN = <a secret PIN for the coach>
 *  4. Deploy → New deployment → type "Web app"
 *       Execute as: Me      Who has access: Anyone
 *     Copy the Web app URL and paste it into config.js  (window.ARC_API = "…")
 *
 * Optional – automatic Instagram posts without any outside service:
 *  5. Get a long-lived Instagram access token for the club account
 *     (Instagram professional account + Meta app with "Instagram API with Instagram login").
 *  6. Script properties → add  IG_TOKEN = <the token>
 *  7. Run the function `installTriggers` once: it refreshes the token automatically every week.
 */

const TABS = {
  sessions: ['id', 'date', 'time', 'title', 'place', 'city', 'mapUrl', 'level', 'distance', 'notes', 'cancelled'],
  updates:  ['id', 'date', 'title', 'text', 'pinned'],
  settings: ['key', 'value']
};

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(TABS).forEach(function (name) {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    sh.getRange(1, 1, sh.getMaxRows(), Math.max(sh.getMaxColumns(), TABS[name].length)).setNumberFormat('@'); // keep everything as plain text
    sh.getRange(1, 1, 1, TABS[name].length).setValues([TABS[name]]).setFontWeight('bold');
    sh.setFrozenRows(1);
  });
  const st = ss.getSheetByName('settings');
  if (st.getLastRow() < 2) {
    st.getRange(2, 1, 7, 2).setValues([
      ['city', ''], ['meetingPoint', ''], ['mapUrl', ''], ['whatsapp', ''],
      ['instagram', 'athena.runclub'], ['whoCanJoin', ''], ['whatToBring', '']
    ]);
  }
  const s1 = ss.getSheetByName('Sheet1') || ss.getSheetByName('Sayfa1');
  if (s1 && s1.getLastRow() === 0 && ss.getSheets().length > 3) ss.deleteSheet(s1);
}

function doGet() {
  const data = readAll();
  data.instagram = instagramPosts();
  return json(Object.assign({ ok: true }, data));
}

/* ---------- Instagram (optional) ---------- */
function instagramPosts() {
  const token = PropertiesService.getScriptProperties().getProperty('IG_TOKEN');
  if (!token) return [];
  const cache = CacheService.getScriptCache();
  const hit = cache.get('ig');
  if (hit) return JSON.parse(hit);
  try {
    const url = 'https://graph.instagram.com/me/media?fields=id,caption,media_type,media_url,thumbnail_url,permalink,timestamp&limit=9&access_token=' + encodeURIComponent(token);
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const j = JSON.parse(res.getContentText());
    const posts = (j.data || []).map(function (p) {
      return { permalink: p.permalink, caption: (p.caption || '').slice(0, 300), timestamp: p.timestamp,
               mediaType: p.media_type, image: p.media_type === 'VIDEO' ? (p.thumbnail_url || '') : (p.media_url || '') };
    });
    cache.put('ig', JSON.stringify(posts), 1800); // 30 minutes
    return posts;
  } catch (err) { return []; }
}

function refreshInstagramToken() {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('IG_TOKEN');
  if (!token) return;
  const res = UrlFetchApp.fetch('https://graph.instagram.com/refresh_access_token?grant_type=ig_refresh_token&access_token=' + encodeURIComponent(token), { muteHttpExceptions: true });
  const j = JSON.parse(res.getContentText());
  if (j.access_token) props.setProperty('IG_TOKEN', j.access_token);
}

function installTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'refreshInstagramToken') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('refreshInstagramToken').timeBased().everyWeeks(1).create();
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json({ ok: false, error: 'bad_request' }); }
  const pin = PropertiesService.getScriptProperties().getProperty('ADMIN_PIN');
  if (!pin || String(body.pin) !== String(pin)) return json({ ok: false, error: 'pin' });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (body.op === 'save' && body.type === 'settings') saveSettings(body.item || {});
    else if (body.op === 'save' && TABS[body.type]) upsert(body.type, body.item || {});
    else if (body.op === 'delete' && TABS[body.type]) remove(body.type, body.id);
    // op 'ping' just checks the PIN
    return json({ ok: true, data: readAll() });
  } finally {
    lock.releaseLock();
  }
}

/* ---------- helpers ---------- */
function sheet(name) { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name); }

function rows(name) {
  const sh = sheet(name), cols = TABS[name];
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, cols.length).getDisplayValues().map(function (r) {
    const o = {}; cols.forEach(function (c, i) { o[c] = r[i]; }); return o;
  }).filter(function (o) { return o[cols[0]] !== ''; });
}

function readAll() {
  const settings = {};
  rows('settings').forEach(function (r) { settings[r.key] = r.value; });
  return { settings: settings, sessions: rows('sessions'), updates: rows('updates') };
}

function clean(v) { v = v === undefined || v === null ? '' : String(v); return v.slice(0, 2000); }

function upsert(name, item) {
  const sh = sheet(name), cols = TABS[name];
  if (!item.id) item.id = Utilities.getUuid().slice(0, 8);
  const values = [cols.map(function (c) { return clean(item[c]); })];
  const ids = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getDisplayValues().map(function (r) { return r[0]; }) : [];
  const i = ids.indexOf(String(item.id));
  const row = i >= 0 ? i + 2 : sh.getLastRow() + 1;
  sh.getRange(row, 1, 1, cols.length).setNumberFormat('@').setValues(values);
}

function remove(name, id) {
  const sh = sheet(name);
  if (sh.getLastRow() < 2) return;
  const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getDisplayValues().map(function (r) { return r[0]; });
  const i = ids.indexOf(String(id));
  if (i >= 0) sh.deleteRow(i + 2);
}

function saveSettings(obj) {
  const sh = sheet('settings');
  const keys = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getDisplayValues().map(function (r) { return r[0]; }) : [];
  Object.keys(obj).forEach(function (k) {
    const i = keys.indexOf(k);
    if (i >= 0) sh.getRange(i + 2, 2).setNumberFormat('@').setValue(clean(obj[k]));
    else { sh.appendRow([k, clean(obj[k])]); keys.push(k); }
  });
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
