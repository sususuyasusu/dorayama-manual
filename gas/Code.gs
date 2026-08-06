/**
 * dorayama-manual-api : staff manual backend (Google Apps Script web app)
 * Data lives in hidden tabs of the seizou spreadsheet:
 *   _manual_content : category | type | label | text | photo | src
 *   _manual_updates : date | title | note | category | by
 *   _app_config     : key | value   (key "manual_edit_code" = edit passcode)
 * Web app: execute as Me / access Anyone. Frontend is a static site (GitHub Pages).
 * POSTs use Content-Type text/plain to avoid CORS preflight.
 * Perf: single SpreadsheetApp.openById per invocation + ScriptCache for GET /manual.
 */
var SHEET_ID = '1PRDhGP_4xiO_ZjJP3NB9Id3PmaPa5W7hNyrqFQ5EyqM';
var T_CONTENT = '_manual_content';
var T_UPDATES = '_manual_updates';
var T_CONFIG = '_app_config';
var CACHE_KEY = 'manual_payload_v1';
var CACHE_TTL_SEC = 120;

var _SS = null;
function _ss() {
  if (!_SS) _SS = SpreadsheetApp.openById(SHEET_ID);
  return _SS;
}

function _sh(name) {
  var sh = _ss().getSheetByName(name);
  if (!sh) throw new Error('tab not found: ' + name);
  return sh;
}

function _values(name) {
  var v = _sh(name).getDataRange().getValues();
  return v.length > 1 ? v.slice(1) : [];
}

function _json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function _code() {
  var rows = _values(T_CONFIG);
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === 'manual_edit_code') {
      return String(rows[i][1]).trim();
    }
  }
  return '';
}

function _invalidateCache() {
  CacheService.getScriptCache().remove(CACHE_KEY);
}

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || 'manual';
  if (action === 'manual') {
    var cache = CacheService.getScriptCache();
    var hit = cache.get(CACHE_KEY);
    if (hit) {
      return ContentService.createTextOutput(hit).setMimeType(ContentService.MimeType.JSON);
    }
    var payload = JSON.stringify({
      ok: true,
      content: _values(T_CONTENT),
      updates: _values(T_UPDATES),
      today: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd')
    });
    cache.put(CACHE_KEY, payload, CACHE_TTL_SEC);
    return ContentService.createTextOutput(payload).setMimeType(ContentService.MimeType.JSON);
  }
  if (action === 'ping') return _json({ ok: true, t: String(new Date()) });
  return _json({ ok: false, error: 'unknown action' });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(20000);
  try {
    var p = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (String(p.code || '').trim() !== _code() || !_code()) {
      return _json({ ok: false, error: 'BAD_CODE' });
    }
    var today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');
    if (p.action === 'add') {
      // insert content row right after the last row of the same category
      var sh = _sh(T_CONTENT);
      var col = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
      var last = -1;
      for (var i = 0; i < col.length; i++) {
        if (String(col[i][0]).trim() === String(p.cat).trim()) last = i + 1;
      }
      var APP_NOTE = 'アプリ追記';
      var row = [p.cat, p.type, p.label || '', p.text, '', today + (p.by ? ' ' + p.by : '') + ' ' + APP_NOTE];
      if (last === -1) {
        sh.appendRow(row);
      } else {
        sh.insertRowAfter(last);
        sh.getRange(last + 1, 1, 1, row.length).setValues([row]);
      }
      _sh(T_UPDATES).appendRow([today, p.upd_title || '', p.upd_note || '', p.cat, p.by || '']);
      _invalidateCache();
      return _json({ ok: true });
    }
    if (p.action === 'announce') {
      _sh(T_UPDATES).appendRow([today, p.title, p.note || '', p.cat || '', p.by || '']);
      _invalidateCache();
      return _json({ ok: true });
    }
    return _json({ ok: false, error: 'unknown action' });
  } catch (err) {
    return _json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
