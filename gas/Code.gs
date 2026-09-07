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
var T_QA = '_ai_questions';
var CACHE_KEY = 'manual_payload_v1';
var CACHE_TTL_SEC = 120;
var QA_HEADERS = ['id', '日時', '質問', '質問者', '状態', '回答', '回答日時'];
var QA_MAX_RETURN = 40;

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

/** 質問箱のタブ。無ければ見出しつきで作る（手作業でのタブ追加を不要にするため） */
function _shQA() {
  var sh = _ss().getSheetByName(T_QA);
  if (!sh) {
    sh = _ss().insertSheet(T_QA);
    sh.appendRow(QA_HEADERS);
    sh.hideSheet();
  }
  return sh;
}

function _qaRows() {
  var sh = _shQA();
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, QA_HEADERS.length).getValues();
}

function _fmt(d) {
  return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
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

  // 質問箱: 直近のやり取り（画面表示用。合言葉なしで読める）
  if (action === 'qa') {
    var rows = _qaRows();
    var out = rows.slice(Math.max(0, rows.length - QA_MAX_RETURN)).map(function (r) {
      return { id: String(r[0]), at: String(r[1]), q: String(r[2]), by: String(r[3]),
               state: String(r[4]), a: String(r[5]), a_at: String(r[6]) };
    });
    out.reverse();
    return _json({ ok: true, qa: out });
  }

  // 質問箱: まだ答えていないもの（Mac側の定期タスクが取りに来る。合言葉が要る）
  if (action === 'pending') {
    if (String((e.parameter && e.parameter.code) || '').trim() !== _code() || !_code()) {
      return _json({ ok: false, error: 'BAD_CODE' });
    }
    var pend = _qaRows().filter(function (r) { return String(r[4]).trim() !== '回答済'; })
      .map(function (r) {
        return { id: String(r[0]), at: String(r[1]), q: String(r[2]), by: String(r[3]) };
      });
    return _json({ ok: true, pending: pend });
  }

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
    // 質問箱: スタッフが質問を出す
    if (p.action === 'ask') {
      var q = String(p.q || '').trim();
      if (!q) return _json({ ok: false, error: 'EMPTY' });
      if (q.length > 500) q = q.slice(0, 500);
      var now = new Date();
      var id = 'q' + now.getTime();
      _shQA().appendRow([id, _fmt(now), q, String(p.by || '').slice(0, 30), '未回答', '', '']);
      return _json({ ok: true, id: id });
    }

    // 質問箱: 回答を書き戻す（Mac側の定期タスクが使う）
    if (p.action === 'answer') {
      var sh2 = _shQA();
      var rows2 = _qaRows();
      for (var j = 0; j < rows2.length; j++) {
        if (String(rows2[j][0]) === String(p.id)) {
          sh2.getRange(j + 2, 6, 1, 2).setValues([[String(p.a || ''), _fmt(new Date())]]);
          sh2.getRange(j + 2, 5).setValue('回答済');
          return _json({ ok: true });
        }
      }
      return _json({ ok: false, error: 'ID_NOT_FOUND' });
    }

    return _json({ ok: false, error: 'unknown action' });
  } catch (err) {
    return _json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
