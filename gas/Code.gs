/**
 * dorayama-manual-api : staff manual backend (Google Apps Script web app)
 * Data lives in hidden tabs of the seizou spreadsheet:
 *   _manual_content : category | type | label | text | photo | src
 *   _manual_updates : date | title | note | category | by
 *   _app_config     : key | value   (key "manual_edit_code" = edit passcode)
 *   _manual_history : 日時 | 操作(追加/修正/削除) | カテゴリ | 前3列 | 後3列 | 記入者 | 取り込み済み
 *                     (アプリでの追加・修正・削除の記録。元に戻す時の控えであり、
 *                      生成スクリプトへ取り込む前の全置換(seed)を止める目印にもなる)
 * Web app: execute as Me / access Anyone. Frontend is a static site (GitHub Pages).
 * POSTs use Content-Type text/plain to avoid CORS preflight.
 * Perf: single SpreadsheetApp.openById per invocation + ScriptCache for GET /manual.
 */
var SHEET_ID = '1PRDhGP_4xiO_ZjJP3NB9Id3PmaPa5W7hNyrqFQ5EyqM';
var T_CONTENT = '_manual_content';
var T_UPDATES = '_manual_updates';
var T_CONFIG = '_app_config';
var T_QA = '_ai_questions';
var T_FACTS = '_ai_facts';
var FACTS_HEADERS = ['キー', '見出し', '答え', '更新日時'];
var T_HISTORY = '_manual_history';
var HISTORY_HEADERS = ['日時', '操作', 'カテゴリ', '種類(前)', 'ラベル(前)', '内容(前)',
                       '種類(後)', 'ラベル(後)', '内容(後)', '記入者', '取り込み済み'];
var APP_NOTE = 'アプリ追記';
var EDIT_NOTE = 'アプリ修正';
// 追加できる種類／修正で入れ替えてよい種類（文章だけの行）
var SIMPLE_TYPES = ['見出し', '項目', '箇条書き', '注意', 'ポイント', 'メモ', '禁止'];
// 種類は変えず、中身だけ直せる行（流れ・カード・写真）。冷蔵庫の図は直せない
var FIXED_TYPES = ['流れ', 'カード', '写真'];
// ラベル欄を使う種類（それ以外はラベルを空にして保存する）
var LABEL_TYPES = ['項目', 'カード', '写真'];
var MAX_TEXT = 3000;
var MAX_LABEL = 300;
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

/** 即答用の「あらかじめ用意した答え」のタブ。無ければ作る */
function _shFacts() {
  var sh = _ss().getSheetByName(T_FACTS);
  if (!sh) {
    sh = _ss().insertSheet(T_FACTS);
    sh.appendRow(FACTS_HEADERS);
    sh.hideSheet();
  }
  return sh;
}

/** セルの値を文字列にする。日付として取り込まれていた場合も日本時間で整える */
function _s(v) {
  if (v instanceof Date) return _fmt(v);
  return String(v === null || v === undefined ? '' : v);
}

/**
 * 質問の日時。id は 'q' + 受付時刻のミリ秒 なので、そこから作るのが確実。
 * 日時セルはスプレッドシートが日付として取り込むと時差がずれるため当てにしない。
 */
function _atOf(id, cell) {
  var m = String(id || '').match(/^q(\d{10,})$/);
  if (m) return _fmt(new Date(Number(m[1])));
  return _s(cell);
}

/** 追加・修正・削除の記録タブ。無ければ作る（文字として記録するため全列を文字形式にする） */
function _shHistory() {
  var sh = _ss().getSheetByName(T_HISTORY);
  if (!sh) {
    sh = _ss().insertSheet(T_HISTORY);
    sh.getRange('A:K').setNumberFormat('@');
    sh.appendRow(HISTORY_HEADERS);
    sh.hideSheet();
  }
  return sh;
}

/** b / a = [種類, ラベル, 内容]（前／後。無ければ null） */
function _logHistory(op, cat, b, a, by) {
  b = b || ['', '', ''];
  a = a || ['', '', ''];
  var sh = _shHistory();
  var row = [_fmt(new Date()), op, cat, String(b[0]), String(b[1]), String(b[2]),
             String(a[0]), String(a[1]), String(a[2]), String(by || ''), ''];
  sh.appendRow(row);
  var rng = sh.getRange(sh.getLastRow(), 1, 1, row.length);
  rng.setNumberFormat('@');
  rng.setValues([row]);
}

/** お知らせを1件足す。日付が日付型に化けて1日ずれないよう、文字として書く */
function _appendUpdate(today, title, note, cat, by) {
  var sh = _sh(T_UPDATES);
  var row = [today, title || '', note || '', cat || '', by || ''];
  sh.appendRow(row);
  var rng = sh.getRange(sh.getLastRow(), 1, 1, row.length);
  rng.setNumberFormat('@');
  rng.setValues([row]);
}

/** 本文の1行を、文字のまま書く（「7/15」などが日付に化けないように） */
function _putContentRow(sh, rowNo, row) {
  var rng = sh.getRange(rowNo, 1, 1, row.length);
  rng.setNumberFormat('@');
  rng.setValues([row]);
}

/**
 * 画面が見ていた行(before)を本文タブから探す。vals は見出し行を含む全行。
 * 行番号 n がそのまま合っていればそれ。誰かが先に行を足し引きしてずれていたら、
 * 同じカテゴリで中身が完全に同じ行を探す（複数あれば n に近い方）。無ければ 0。
 */
function _findRow(vals, cat, n, before) {
  before = before || {};
  function same(i) {
    var r = vals[i];
    return String(r[0]).trim() === cat &&
      String(r[1]) === String(before.type || '') &&
      String(r[2]) === String(before.label || '') &&
      String(r[3]) === String(before.text || '') &&
      String(r[4]) === String(before.photo || '');
  }
  if (n >= 2 && n <= vals.length && same(n - 1)) return n;
  var best = 0;
  for (var i = 1; i < vals.length; i++) {
    if (same(i) && (!best || Math.abs(i + 1 - n) < Math.abs(best - n))) best = i + 1;
  }
  return best;
}

/** 修正の中身の検査。問題なければ空文字、あれば理由の符号 */
function _checkEdit(oldType, newType, label, text) {
  if (oldType === '冷蔵庫') return 'LOCKED';
  if (FIXED_TYPES.indexOf(oldType) >= 0) {
    if (newType !== oldType) return 'BAD_TYPE';
  } else if (SIMPLE_TYPES.indexOf(newType) < 0) {
    return 'BAD_TYPE';
  }
  if (newType === '写真') {
    if (!String(label).trim()) return 'EMPTY';
  } else if (!String(text).trim()) {
    return 'EMPTY';
  }
  if (newType === 'カード' && !String(label).trim()) return 'EMPTY';
  if (String(text).length > MAX_TEXT || String(label).length > MAX_LABEL) return 'TOO_LONG';
  return '';
}

/**
 * 追加。p.sec（見出しの文字）があればその見出しの最後へ、無ければカテゴリの最後へ入れる。
 * お知らせは p.upd_title がある時だけ出す。
 */
function _addRow(p, today) {
  var cat = String(p.cat || '').trim();
  var type = String(p.type || '').trim();
  var label = String(p.label || '');
  var text = String(p.text || '');
  if (!cat) return { ok: false, error: 'EMPTY' };
  if (SIMPLE_TYPES.indexOf(type) < 0) return { ok: false, error: 'BAD_TYPE' };
  if (!text.trim()) return { ok: false, error: 'EMPTY' };
  if (text.length > MAX_TEXT || label.length > MAX_LABEL) return { ok: false, error: 'TOO_LONG' };
  if (LABEL_TYPES.indexOf(type) < 0) label = '';

  var sh = _sh(T_CONTENT);
  var vals = sh.getDataRange().getValues();
  var last = 0;                       // そのカテゴリの最後の行（シート上の行番号）
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === cat) last = i + 1;
  }
  var at = last;
  var sec = String(p.sec || '').trim();
  if (sec && last) {
    for (var h = 1; h < vals.length; h++) {
      if (String(vals[h][0]).trim() === cat && String(vals[h][1]).trim() === '見出し' &&
          String(vals[h][3]).trim() === sec) {
        var j = h + 1;                // 見出しの次の行から、次の見出し(またはカテゴリの終わり)の手前まで
        while (j < vals.length && String(vals[j][0]).trim() === cat &&
               String(vals[j][1]).trim() !== '見出し') j++;
        at = j;                       // j はその区切りに入らない最初の行の添字 → 最後の行の行番号は j
        break;
      }
    }
  }
  var by = String(p.by || '').slice(0, 30);
  var src = today + (by ? ' ' + by : '') + ' ' + APP_NOTE;
  var row = [cat, type, label, text, '', src];
  _logHistory('追加', cat, null, [type, label, text], by);
  if (!at) {
    sh.appendRow(row);
    at = sh.getLastRow();
  } else {
    sh.insertRowAfter(at);
    at = at + 1;
  }
  _putContentRow(sh, at, row);
  if (String(p.upd_title || '').trim()) _appendUpdate(today, p.upd_title, p.upd_note, cat, by);
  _invalidateCache();
  return { ok: true, n: at };
}

/** 修正。画面が見ていた内容(before)と今のシートが違えば CHANGED を返して何も変えない */
function _editRow(p, today) {
  var cat = String(p.cat || '').trim();
  var a = p.after || {};
  var sh = _sh(T_CONTENT);
  var vals = sh.getDataRange().getValues();
  var hit = _findRow(vals, cat, Number(p.n), p.before);
  if (!hit) return { ok: false, error: 'CHANGED' };
  var old = vals[hit - 1];
  var oldType = String(old[1]).trim();
  var newType = String(a.type || '').trim();
  var label = String(a.label || '');
  var text = String(a.text || '');
  var bad = _checkEdit(oldType, newType, label, text);
  if (bad) return { ok: false, error: bad };
  if (LABEL_TYPES.indexOf(newType) < 0) label = '';
  if (newType === '写真') text = String(old[3]);          // 写真は説明文（ラベル）だけ直せる
  if (newType === String(old[1]) && label === String(old[2]) && text === String(old[3])) {
    return { ok: true, n: hit, unchanged: true };
  }
  var by = String(p.by || '').slice(0, 30);
  var mark = today + (by ? ' ' + by : '') + ' ' + EDIT_NOTE;
  var kept = String(old[5] || '').split(' / ').filter(function (s) {
    return s && s.indexOf(EDIT_NOTE) < 0;
  });
  kept.push(mark);
  _logHistory('修正', cat, [old[1], old[2], old[3]], [newType, label, text], by);
  _putContentRow(sh, hit, [old[0], newType, label, text, old[4], kept.join(' / ')]);
  if (p.announce && String(p.upd_title || '').trim()) {
    _appendUpdate(today, p.upd_title, p.upd_note, cat, by);
  }
  _invalidateCache();
  return { ok: true, n: hit };
}

/** 削除。消した中身は記録タブに残る（元に戻す時はそこから写す） */
function _deleteRow(p) {
  var cat = String(p.cat || '').trim();
  var sh = _sh(T_CONTENT);
  var vals = sh.getDataRange().getValues();
  var hit = _findRow(vals, cat, Number(p.n), p.before);
  if (!hit) return { ok: false, error: 'CHANGED' };
  var old = vals[hit - 1];
  if (String(old[1]).trim() === '冷蔵庫') return { ok: false, error: 'LOCKED' };
  _logHistory('削除', cat, [old[1], old[2], old[3]], null, String(p.by || '').slice(0, 30));
  sh.deleteRow(hit);
  _invalidateCache();
  return { ok: true, n: hit };
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

  // 催事報告アプリ（本部の一覧。合言葉が要る）→ Saiji.gs
  if (action === 'saiji_list') return _json(_saijiList(e));

  // 質問箱: 直近のやり取り（画面表示用。合言葉なしで読める）
  if (action === 'qa') {
    var rows = _qaRows();
    var out = rows.slice(Math.max(0, rows.length - QA_MAX_RETURN)).map(function (r) {
      return { id: _s(r[0]), at: _atOf(r[0], r[1]), q: _s(r[2]), by: _s(r[3]),
               state: _s(r[4]), a: _s(r[5]), a_at: _s(r[6]) };
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
        return { id: _s(r[0]), at: _atOf(r[0], r[1]), q: _s(r[2]), by: _s(r[3]) };
      });
    return _json({ ok: true, pending: pend });
  }

  // 即答用データ（画面が起動時に読む。合言葉なし）
  if (action === 'facts') {
    var fsh = _shFacts();
    var frows = fsh.getLastRow() < 2 ? []
      : fsh.getRange(2, 1, fsh.getLastRow() - 1, FACTS_HEADERS.length).getValues();
    var facts = frows.map(function (r) {
      return { k: _s(r[0]), title: _s(r[1]), a: _s(r[2]), at: _s(r[3]) };
    }).filter(function (x) { return x.k && x.a; });
    return _json({ ok: true, facts: facts });
  }

  return _json({ ok: false, error: 'unknown action' });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(20000);
  try {
    var p = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    // 催事報告アプリ（販売員の送信。派遣会社の人も使うので合言葉なし）→ Saiji.gs
    if (p.action === 'saiji_submit') return _json(_saijiSubmit(p));
    if (String(p.code || '').trim() !== _code() || !_code()) {
      return _json({ ok: false, error: 'BAD_CODE' });
    }
    var today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');
    // 本文の追加・修正・削除（記録タブにも残す）。
    // 画面は通信が不安定だと同じ送信をやり直すので、送信ごとの番号(rid)で「もう実行済み」を見分け、二重に反映しない
    if (p.action === 'add' || p.action === 'edit' || p.action === 'delete') {
      var rid = String(p.rid || '').slice(0, 60);
      var cache = CacheService.getScriptCache();
      if (rid) {
        var seen = cache.get('rid_' + rid);
        if (seen) return _json(JSON.parse(seen));
      }
      var res = p.action === 'add' ? _addRow(p, today)
        : p.action === 'edit' ? _editRow(p, today) : _deleteRow(p);
      if (rid && res.ok) cache.put('rid_' + rid, JSON.stringify(res), 600);
      return _json(res);
    }
    if (p.action === 'announce') {
      _appendUpdate(today, p.title, p.note, p.cat, p.by);
      _invalidateCache();
      return _json({ ok: true });
    }
    // 質問箱: スタッフが質問を出す
    // p.a があれば「その場で答えが出た質問」＝回答済として記録する（何を聞かれたかは残す）
    if (p.action === 'ask') {
      var q = String(p.q || '').trim();
      if (!q) return _json({ ok: false, error: 'EMPTY' });
      if (q.length > 500) q = q.slice(0, 500);
      var now = new Date();
      var id = 'q' + now.getTime();
      var instant = String(p.a || '').trim();
      _shQA().appendRow([id, _fmt(now), q, String(p.by || '').slice(0, 30),
                         instant ? '回答済' : '未回答', instant, instant ? _fmt(now) : '']);
      return _json({ ok: true, id: id });
    }

    // 即答用データの入れ替え（Mac側が毎日流し込む）
    if (p.action === 'facts_set') {
      var rows = p.facts || [];
      var fsh = _shFacts();
      fsh.clear();
      fsh.appendRow(FACTS_HEADERS);
      if (rows.length) {
        var stamp = _fmt(new Date());
        var out = rows.map(function (r) {
          return [String(r.k || ''), String(r.title || ''), String(r.a || ''), stamp];
        });
        var rng = fsh.getRange(2, 1, out.length, FACTS_HEADERS.length);
        // 日時列を日付として取り込まれると時差でずれるので、文字列として書く
        rng.setNumberFormat('@');
        rng.setValues(out);
      }
      return _json({ ok: true, n: rows.length });
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


/* ===================== 催事報告アプリ（saiji/report.html） ===================== */
/**
 * 催事報告アプリ（saiji/report.html）の受け口。上の doGet / doPost から呼ばれる。
 *
 * 流れ:
 *   販売員のスマホ → POST {action:'saiji_submit', kind, data, rid}
 *     1) 隠しタブ _saiji_reports に1行残す（全項目をJSONでも保存）
 *     2) kind === 'daily'（閉店後の日報）の時だけ、今までのGoogleフォーム
 *        「どら山 売上日報提出」へ同じ内容を回答として送る。
 *        → フォーム回答シート → 夜のAirメイト自動入力・速報・予測は今まで通り動く
 *   本部 → GET ?action=saiji_list&date=YYYY-MM-DD&code=合言葉 で会場ごとの状況を見る
 *
 * 合言葉: _app_config の saiji_hq_code（無ければ manual_edit_code と同じ）
 * 販売員側は合言葉なしで送れる（派遣会社の人にも配れるように）。
 * 送信ごとの番号 rid で二重送信を見分ける（電波が悪い会場で再送されても1回分）。
 */
var SAIJI_FORM_SHEET_ID = '1v0w_oAmbTmw3t9oOhCFVWv87ysm2gpOnNQbGAlALpIc';
var T_SAIJI = '_saiji_reports';
var SAIJI_HEADERS = ['受信日時', 'rid', '種類', '日付', '会場', '販売員', '所属',
                     '売上税込', '売上税抜', '客数', 'フォーム送信', '中身(JSON)'];
var SAIJI_KINDS = ['open', 'delivery', 'progress', 'daily', 'trouble'];
var SAIJI_KIND_JA = { open: '開店準備', delivery: '入荷検品', progress: '途中報告',
                      daily: '閉店日報', trouble: '困りごと' };
// フォームの商品列（設問タイトルに含まれる言葉 → アプリの商品キー）
var SAIJI_PRODUCTS = [
  ['黒どら', 'kuro'], ['白どら', 'shiro'], ['あんバター', 'anbutter'],
  ['旬どら', 'shun'], ['皮だけ', 'kawa'], ['生どら', 'nama']
];

function _saijiCode() {
  var rows = _values(T_CONFIG);
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === 'saiji_hq_code') {
      var v = String(rows[i][1]).trim();
      if (v) return v;
    }
  }
  return _code();
}

function _shSaiji() {
  var sh = _ss().getSheetByName(T_SAIJI);
  if (!sh) {
    sh = _ss().insertSheet(T_SAIJI);
    sh.getRange('A:L').setNumberFormat('@');
    sh.appendRow(SAIJI_HEADERS);
    sh.setFrozenRows(1);
  }
  return sh;
}

function _num(v) {
  var s = String(v === null || v === undefined ? '' : v)
    .replace(/[０-９]/g, function (c) { return String.fromCharCode(c.charCodeAt(0) - 0xFEE0); })
    .replace(/[,，¥￥円\s]/g, '');
  var m = s.match(/^-?\d+/);
  return m ? Number(m[0]) : 0;
}

function _clip(v, n) { return String(v === null || v === undefined ? '' : v).slice(0, n); }

/** 'YYYY-MM-DD' / 'YYYY/M/D' → {y,m,d}。だめなら null */
function _ymd(s) {
  var m = String(s || '').match(/^(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})$/);
  if (!m) return null;
  return { y: Number(m[1]), m: Number(m[2]), d: Number(m[3]) };
}

function _ridSeen(sh, rid) {
  if (!rid) return false;
  var cache = CacheService.getScriptCache();
  if (cache.get('srid_' + rid)) return true;
  var last = sh.getLastRow();
  if (last < 2) return false;
  var from = Math.max(2, last - 300);   // 直近300件だけ見る（十分）
  var ids = sh.getRange(from, 2, last - from + 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === rid) return true;
  return false;
}

/** 販売員の送信を受ける */
function _saijiSubmit(p) {
  var kind = String(p.kind || '');
  if (SAIJI_KINDS.indexOf(kind) < 0) return { ok: false, error: 'BAD_KIND' };
  var d = p.data || {};
  var rid = _clip(p.rid, 60);
  var ymd = _ymd(d.date);
  if (!ymd) return { ok: false, error: 'BAD_DATE' };
  var venue = _clip(d.venue, 60).trim();
  if (!venue) return { ok: false, error: 'NO_VENUE' };

  var sh = _shSaiji();
  if (_ridSeen(sh, rid)) return { ok: true, dup: true };

  var formNote = '';
  if (kind === 'daily') {
    if (_num(d.salesIncl) <= 0) return { ok: false, error: 'NO_SALES' };
    try {
      formNote = _saijiToForm(d);
    } catch (err) {
      formNote = '失敗: ' + String(err).slice(0, 200);
    }
  }
  var row = [_fmt(new Date()), rid, SAIJI_KIND_JA[kind] || kind,
             Utilities.formatString('%04d/%02d/%02d', ymd.y, ymd.m, ymd.d),
             venue, _clip(d.staff, 40), _clip(d.agency, 40),
             d.salesIncl === undefined ? '' : String(_num(d.salesIncl)),
             d.salesExcl === undefined ? '' : String(_num(d.salesExcl)),
             d.customers === undefined ? '' : String(_num(d.customers)),
             formNote, _clip(JSON.stringify(d), 45000)];
  sh.appendRow(row);
  var rng = sh.getRange(sh.getLastRow(), 1, 1, row.length);
  rng.setNumberFormat('@');
  rng.setValues([row]);
  if (rid) CacheService.getScriptCache().put('srid_' + rid, '1', 21600);
  return { ok: true, form: formNote };
}

/** 設問タイトル → 入れる値（入れない設問は null） */
function _saijiValueFor(title, d) {
  var t = String(title || '').replace(/[\s　]/g, '');
  var p = d.products || {};
  for (var i = 0; i < SAIJI_PRODUCTS.length; i++) {
    if (t.indexOf(SAIJI_PRODUCTS[i][0]) >= 0) {
      var x = p[SAIJI_PRODUCTS[i][1]] || {};
      if (t.indexOf('在庫') >= 0 || t.indexOf('残') >= 0) return String(_num(x.left));
      if (t.indexOf('販売') >= 0) return String(_num(x.sold));
      return null;
    }
  }
  var ymd = _ymd(d.date);
  if (t.indexOf('手入力') >= 0) return ymd.y + '/' + ymd.m + '/' + ymd.d;
  if (t.indexOf('日付') >= 0) return { date: ymd };
  if (t.indexOf('店舗') >= 0 || t.indexOf('会場') >= 0) return String(d.venue || '');
  if (t.indexOf('税込') >= 0) return String(_num(d.salesIncl));
  if (t.indexOf('税抜') >= 0) return String(_num(d.salesExcl));
  if (t.indexOf('客数') >= 0) return String(_num(d.customers));
  if (t.indexOf('所感') >= 0 || t.indexOf('来店') >= 0) return String(d.formComment || d.comment || '');
  return null;
}

/** 選択肢の設問で、表記の近い選択肢を選ぶ */
function _pickChoice(choices, v) {
  var norm = function (s) { return String(s).replace(/[\s　]/g, '').replace(/松阪屋/g, '松坂屋'); };
  var nv = norm(v);
  for (var i = 0; i < choices.length; i++) if (norm(choices[i].getValue()) === nv) return choices[i].getValue();
  for (var j = 0; j < choices.length; j++) {
    var c = norm(choices[j].getValue());
    if (c && (c.indexOf(nv) >= 0 || nv.indexOf(c) >= 0)) return choices[j].getValue();
  }
  return null;
}

/** 今までのGoogleフォームへ、日報を回答として送る。結果のひとことを返す */
function _saijiToForm(d) {
  var url = SpreadsheetApp.openById(SAIJI_FORM_SHEET_ID).getFormUrl();
  if (!url) return '未送信: フォームが見つからない';
  var form = FormApp.openByUrl(url);
  var resp = form.createResponse();
  var items = form.getItems();
  var used = 0, skipped = [];
  for (var i = 0; i < items.length; i++) {
    var it = items[i];
    var v = _saijiValueFor(it.getTitle(), d);
    if (v === null || v === '') continue;
    var type = it.getType();
    try {
      if (type === FormApp.ItemType.DATE) {
        if (!v.date) continue;
        resp.withItemResponse(it.asDateItem().createResponse(new Date(v.date.y, v.date.m - 1, v.date.d)));
      } else if (v && v.date) {
        continue;
      } else if (type === FormApp.ItemType.TEXT) {
        resp.withItemResponse(it.asTextItem().createResponse(v));
      } else if (type === FormApp.ItemType.PARAGRAPH_TEXT) {
        resp.withItemResponse(it.asParagraphTextItem().createResponse(v));
      } else if (type === FormApp.ItemType.LIST) {
        var li = it.asListItem();
        var c1 = _pickChoice(li.getChoices(), v);
        if (!c1) { skipped.push(it.getTitle()); continue; }
        resp.withItemResponse(li.createResponse(c1));
      } else if (type === FormApp.ItemType.MULTIPLE_CHOICE) {
        var mc = it.asMultipleChoiceItem();
        var c2 = _pickChoice(mc.getChoices(), v);
        if (!c2 && !mc.hasOtherOption()) { skipped.push(it.getTitle()); continue; }
        resp.withItemResponse(mc.createResponse(c2 || v));
      } else {
        continue;
      }
      used++;
    } catch (err) {
      skipped.push(it.getTitle());
    }
  }
  resp.submit();
  return 'OK ' + used + '問' + (skipped.length ? '（入らず: ' + skipped.join('、').slice(0, 120) + '）' : '');
}

/** 本部用: 指定日の報告一覧（合言葉が要る） */
function _saijiList(e) {
  var prm = (e && e.parameter) || {};
  if (String(prm.code || '').trim() !== _saijiCode() || !_saijiCode()) {
    return { ok: false, error: 'BAD_CODE' };
  }
  var ymd = _ymd(prm.date);
  if (!ymd) return { ok: false, error: 'BAD_DATE' };
  var day = Utilities.formatString('%04d/%02d/%02d', ymd.y, ymd.m, ymd.d);
  var since = new Date(ymd.y, ymd.m - 1, ymd.d - 7);
  var sinceStr = Utilities.formatString('%04d/%02d/%02d', since.getFullYear(), since.getMonth() + 1, since.getDate());
  var sh = _shSaiji();
  var last = sh.getLastRow();
  if (last < 2) return { ok: true, date: day, items: [], recentVenues: [] };
  var from = Math.max(2, last - 1500);
  var vals = sh.getRange(from, 1, last - from + 1, SAIJI_HEADERS.length).getValues();
  var items = [], recent = {};
  for (var i = 0; i < vals.length; i++) {
    var r = vals[i];
    var rd = _s(r[3]).slice(0, 10);
    if (rd >= sinceStr && rd < day && _s(r[2]) === '閉店日報') recent[_s(r[4])] = rd;
    if (rd !== day) continue;
    var body = {};
    try { body = JSON.parse(_s(r[11]) || '{}'); } catch (err) { body = {}; }
    items.push({ at: _s(r[0]), kind: _s(r[2]), venue: _s(r[4]), staff: _s(r[5]), agency: _s(r[6]),
                 salesIncl: _s(r[7]), salesExcl: _s(r[8]), customers: _s(r[9]), form: _s(r[10]),
                 data: body });
  }
  var rv = Object.keys(recent).map(function (k) { return { venue: k, last: recent[k] }; });
  return { ok: true, date: day, items: items, recentVenues: rv };
}
