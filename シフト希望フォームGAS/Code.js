// ===== フォームの決まりごと（募集期間・締切・名簿・祝日はシフト表から読む） =====
var STORE_NAME = 'ペッパーランチ飯塚店';
var NOTE_CHOICES = ['6時間以内希望', 'ロング希望', '試験休み', '旅行のため', '未定'];
var NOTE_SEPARATOR = '／';            // 備考欄に「選択／手入力」の形で書く
// スタッフが選べる時間の幅。シフト表の「シフト希望の募集設定」で店舗ごとに変えられる
var TIME_RANGE_CELLS = { startFrom: 'E42', startTo: 'E43', endFrom: 'E44', endTo: 'E45' };
var TIME_RANGE_DEFAULT = { start: { from: 9 * 60, to: 19 * 60 + 30 }, end: { from: 12 * 60, to: 21 * 60 + 30 } };

// シフト表（37期飯塚店シフト）
var SHIFT_SPREADSHEET_ID = '1WXZ-hOSbPC4h7Atg75hnGp2Z3KLOggS6vTbE0SAj850';
var FIRST_WRITABLE_SHEET = '261019';  // これより前の週シートには書き込まない
var ROSTER_START_ROW = 4;             // 名前はA4〜A25
var ROSTER_ROWS = 22;
var INPUT_COLS = 15;                  // B〜P列（出勤・退勤×7日＋備考）
var HOLIDAY_SHEET_NAME = '祝日';

// シフト表の「シフト希望設定」シート（シフト表GASのメニューで作る。並びはシフト表GAS側と同じにすること）
var REQUEST_SETTINGS_SHEET = 'シフト希望設定';
var SETTING_START_CELL = 'B3';        // 今回だけの開始週（月曜日）yyyy-MM-dd。空なら自動
var SETTING_WEEKS_CELL = 'B4';        // 今回だけの週数
var SETTING_DEADLINE_CELL = 'B5';     // 今回だけの締切 yyyy-MM-dd HH:mm
var SETTING_STORE_NAME_CELL = 'E6';   // 店舗名（シフト表のメニューで設定する）
var BASE_CLOSING_DAY_CELL = 'E3';     // 基本設定：締め日（毎月この日。社員の締め日）
var BASE_DEADLINE_WEEKS_CELL = 'E4';  // 基本設定：スタッフ入力締切＝募集が始まる何週間前の日曜日
var BASE_DEADLINE_TIME_CELL = 'E5';   // 基本設定：締切の時刻
var DEFAULT_BASE = { closingDay: 15, deadlineWeeks: 2, deadlineTime: '23:59' };
var CAST_CLOSING_DAY_CELL = 'E7';     // キャストの締め日（店長・社員以外の人の出勤の区切り）
var DEFAULT_CAST_CLOSING_DAY = 20;
var ROLE_SHEET_NAME = 'スタッフ区分';  // A列＝名前、B列＝区分。店長・社員だけが載っていて、載っていない人はキャスト
var EMPLOYEE_ROLES = ['店長', '社員'];
var PAY_DAY_CAP_HOURS = 8;            // 概算の給与は1日この時間までで数える
var HISTORY_SHEET_NAME = '出勤履歴';  // A日付・C名前（シフト表GASが週シートの右側から作る。1人1日1行）
var PENDING_FIRST_ROW = 9;            // まだ週シートが無い週の希望
var PENDING_COLS = 19;                // シート名・名前・時間14・備考・送信日時・端末ID
var PENDING_DEVICE_INDEX = 18;

// LINE公式アカウント（お知らせはシフト表GASが送る。ここは「送り先の登録」を受け取るだけ）
var LINE_TOKEN_CELL = 'E9';         // チャネルアクセストークン
var LINE_TARGET_CELL = 'E10';       // お知らせを送る先のID
var LINE_REGISTER_CELL = 'E12';     // 登録を受け付ける期限（シフト表のメニューで開く。うしろに |test が付くとお試し用）
var LINE_TEST_TARGET_CELL = 'E32';  // お試しの送り先のID（お試し公開のときだけ使われる）
var LINE_REPLY_URL = 'https://api.line.me/v2/bot/message/reply';

// 記録を書き込むスプレッドシート（14380@hotpalette.com の「ペッパーランチ飯塚店」フォルダ内）
var STORE_SPREADSHEET_ID = '1tgFvH1k3WxsMb_fG2C_m6vMu_n9KTKGgdScSk7jNxGM'; // シフト希望_アクセス記録（飯塚店）

// 同じ端末・同じ名前から短時間に何回も来たら止める
var LIMITS = {
  submit: { max: 3, seconds: 300 },
  event: { max: 30, seconds: 300 }
};

var SHEETS = {
  log: {
    name: '操作の記録',
    headers: ['日時', '端末の持ち主', '操作', '選んだ名前', '内容', '注意', '端末ID', 'IPアドレス', '機種', 'ブラウザ情報']
  },
  submit: {
    name: '送信の記録',
    headers: ['送信日時', '端末の持ち主', '名前', '注意', '出勤日数', '合計時間', '1週', '2週', '3週', '4週', '5週', '端末ID', 'IPアドレス', '機種', 'ブラウザ情報', 'データ']
  },
  devices: {
    name: '端末の持ち主',
    headers: ['端末ID', '持ち主（店長が選ぶ）', '候補（よく選ばれた名前）', 'この端末で選ばれた名前', '機種', '最初に使った日時', '最後に使った日時', '使った回数']
  }
};
var LOG_DEVICE_COL = 7;      // 操作の記録の「端末ID」列
var SUBMIT_DEVICE_COL = 12;  // 送信の記録の「端末ID」列
var OWNER_COL = 2;           // 両シートの「端末の持ち主」列

var EVENT_ACTIONS = { open: '画面を開いた', select: '名前を選択', change: '名前を変更' };
var ACTION_SUBMIT = '送信';
var ACTION_FAILED = '送信できなかった';
var DOW_LABELS = ['月', '火', '水', '木', '金', '土', '日'];

// Claudeの自動確認（試験用ブラウザ）で入った記録の目印
var TEST_BROWSER_MARK = 'HeadlessChrome';
// 自動確認の記録は、確かめる時間を残して1時間たったら、ふつうの操作のついでに自動で消す
var TEST_RECORD_KEEP_MINUTES = 60;

// LINEログイン（LIFF）から来ると、「?」のあとの指定が liff.state という1つの箱に
// まとめられて届くことがある。ほどいて、ふつうの指定と同じように読めるようにする。
function applyLiffState_(e) {
  try {
    if (!e || !e.parameter || !e.parameter['liff.state']) return e;
    String(e.parameter['liff.state']).replace(/^[?#]/, '').split('&').forEach(function(pair) {
      if (!pair) return;
      var cut = pair.indexOf('=');
      var key = decodeURIComponent(cut < 0 ? pair : pair.slice(0, cut));
      if (!key || e.parameter[key]) return;                       // もとの指定があればそちらを使う
      e.parameter[key] = cut < 0 ? '' : decodeURIComponent(pair.slice(cut + 1).replace(/\+/g, ' '));
    });
  } catch (err) {}
  return e;
}

// ページを開くたびにシフト表を開き直すと、それだけで2〜3秒かかる。
// 中身は全員に同じもの（募集の期間・名簿・祝日）なので、少しの間だけ覚えておく。
// 店長が設定を変えても、この秒数がすぎれば新しい内容になる。
var PAGE_CACHE_SEC = 120;
// できたシフトを見る画面は、シフト表の週シートを何枚も読むので重い（実測10〜11秒）。
// 中身は全員に同じなので長めに覚えておき、店長が公開したときに消す（下の ?flush=）
// シフトは毎日変わるものではないので、長く覚えておく。
// 公開（確定・もう一度知らせる）のたびにシフト表GASが ?flush= を呼んで消すので、古いまま出ることはない。
// ただし万一その連絡が届かなかったときのために、覚えておく上限を決めておく
var SHIFT_PAGE_CACHE_SEC = 21600;              // すぐ出せるほう（CacheService）の上限＝6時間
var SHIFT_PAGE_KEEP_DAYS = 4;                  // 長く覚えておくほう（文書の保管場所）の上限＝4日
var SHIFT_PAGE_PROP_KEY = 'SHIFTPAGE_SAVED';   // 保管場所の鍵（中身は分けて入れる）
var SHIFT_PAGE_CHUNK = 7000;                   // 1つに入れる文字数
var SHIFT_PAGE_CACHE_KEY = 'SHIFTPAGE_V1';

function pageCache_() {
  try { return CacheService.getScriptCache(); } catch (err) { return null; }
}

// key で覚えておく。無ければ build() で作って覚える
// 募集の設定（全員に同じ内容）。開くたびにシフト表を開き直さない
function cachedConfig_() {
  return remember_('CONFIG_V2', function() {
    return loadConfig_(SpreadsheetApp.openById(SHIFT_SPREADSHEET_ID));
  });
}

function remember_(key, build) {
  var cache = pageCache_();
  if (cache) {
    var hit = cache.get(key);
    if (hit) { try { return JSON.parse(hit); } catch (err) {} }
  }
  var made = build();
  if (cache) { try { cache.put(key, JSON.stringify(made), PAGE_CACHE_SEC); } catch (err) {} }
  return made;
}

// ===== 長く覚えておく入れ物（文書の保管場所）=====
// CacheService は最長6時間なので、それより長くはこちらに置く。
// 大きいと1つに入らないので、7000文字ずつに分けて入れる
function longStore_() {
  try { return PropertiesService.getScriptProperties(); } catch (err) { return null; }
}

function saveLong_(obj) {
  var props = longStore_();
  if (!props) return false;
  try {
    var text = JSON.stringify(obj);
    var before = Number(props.getProperty(SHIFT_PAGE_PROP_KEY + '_n') || 0);
    var n = Math.ceil(text.length / SHIFT_PAGE_CHUNK);
    var map = {};
    map[SHIFT_PAGE_PROP_KEY + '_n'] = String(n);
    for (var i = 0; i < n; i++) map[SHIFT_PAGE_PROP_KEY + '_' + i] = text.substr(i * SHIFT_PAGE_CHUNK, SHIFT_PAGE_CHUNK);
    props.setProperties(map, false);
    for (var j = n; j < before; j++) props.deleteProperty(SHIFT_PAGE_PROP_KEY + '_' + j);   // 前より短くなった分を消す
    return true;
  } catch (err) { return false; }
}

function loadLong_() {
  var props = longStore_();
  if (!props) return null;
  try {
    var all = props.getProperties();
    var n = Number(all[SHIFT_PAGE_PROP_KEY + '_n'] || 0);
    if (!n) return null;
    var text = '';
    for (var i = 0; i < n; i++) {
      var part = all[SHIFT_PAGE_PROP_KEY + '_' + i];
      if (part === undefined) return null;      // 欠けていたら使わない
      text += part;
    }
    return JSON.parse(text);
  } catch (err) { return null; }
}

function clearLong_() {
  var props = longStore_();
  if (!props) return;
  try {
    var n = Number(props.getProperty(SHIFT_PAGE_PROP_KEY + '_n') || 0);
    props.deleteProperty(SHIFT_PAGE_PROP_KEY + '_n');
    for (var i = 0; i < n; i++) props.deleteProperty(SHIFT_PAGE_PROP_KEY + '_' + i);
  } catch (err) {}
}

// 覚えておいた内容を消す。店長が公開したときにシフト表GASから呼ばれる
function forgetPages_() {
  clearLong_();
  var cache = pageCache_();
  if (!cache) return false;
  try { cache.removeAll([SHIFT_PAGE_CACHE_KEY, 'CONFIG_V2']); } catch (err) { return false; }
  return true;
}

// LINEログインから戻ってきたとき（?code=…&state=…）、本人のLINEユーザーIDを受け取り、
// もとの行き先（state）に戻す。うまくいかなければ、今までどおり名前を選んでもらう。
function applyLineLogin_(e) {
  if (!e || !e.parameter || !e.parameter.code) return e;
  var why = '';                               // うまくいかなかった理由（画面に出して手がかりにする）
  try {
    var shiftSs = SpreadsheetApp.openById(SHIFT_SPREADSHEET_ID);
    var id = shiftSetting_(shiftSs, LINE_LOGIN_ID_CELL).replace(/[^0-9]/g, '');   // 桁区切りが付いていても読む
    var secret = shiftSetting_(shiftSs, LINE_LOGIN_SECRET_CELL);
    if (!id || !secret) {
      why = '設定が入っていません（ID' + (id ? '有' : '無') + '／シークレット' + (secret ? '有' : '無') + '）';
    } else {
      var got = UrlFetchApp.fetch(LINE_TOKEN_URL, {
        method: 'post', muteHttpExceptions: true,
        payload: {
          grant_type: 'authorization_code',
          code: String(e.parameter.code),
          redirect_uri: WEB_APP_URL,
          client_id: id,
          client_secret: secret
        }
      });
      var body = JSON.parse(got.getContentText() || '{}');
      if (!body.access_token) {
        why = '合い言葉の交換に失敗（' + got.getResponseCode() + '：'
          + (body.error_description || body.error || '理由なし') + '）';
      } else {
        var who = UrlFetchApp.fetch(LINE_PROFILE_URL, {
          muteHttpExceptions: true, headers: { Authorization: 'Bearer ' + body.access_token }
        });
        var prof = JSON.parse(who.getContentText() || '{}');
        if (prof.userId) e.parameter.lid = String(prof.userId);
        else why = '名前の取り出しに失敗（' + who.getResponseCode() + '）';
      }
    }
  } catch (err) {
    why = 'つまずき：' + err.message;
    console.error('LINEログイン: ' + err.message);
  }
  if (why) e.parameter.lerr = why;
  String(e.parameter.state || '').replace(/^\?/, '').split('&').forEach(function(one) {
    var at = one.indexOf('=');
    if (at <= 0) return;
    var key = decodeURIComponent(one.slice(0, at));
    if (key === 'lid' || key === 'code' || key === 'state') return;   // 外から入れさせない
    if (e.parameter[key] === undefined) e.parameter[key] = decodeURIComponent(one.slice(at + 1));
  });
  return e;
}

function doGet(e) {
  e = applyLiffState_(e);
  e = applyLineLogin_(e);
  // 速さの調べ用。スプレッドシートを一切読まずに返す（中身は無い）
  if (e && e.parameter && e.parameter.ping) return HtmlService.createHtmlOutput('ok');
  // 店長が公開したとき、シフト表GASから呼ばれる。覚えていた内容を消して作り直させる
  if (e && e.parameter && e.parameter.flush) {
    return ContentService.createTextOutput(forgetPages_() ? 'ok' : 'ng');
  }
  // ?slow=1 で、どこに時間がかかっているかを段階ごとに返す（遅いと言われたときの調べ用）
  if (e && e.parameter && e.parameter.slow) {
    return ContentService.createTextOutput(JSON.stringify(shiftPageTiming_(), null, 1))
      .setMimeType(ContentService.MimeType.JSON);
  }
  // ?lineowners=1 で、LINEログインの登録が何人ぶん入っているかだけを返す（調査用。名前は返さない）
  if (e && e.parameter && e.parameter.lineowners) return lineOwnerCount_();
  // カレンダーアプリが読みに来たとき（?cal=その人の鍵）は、その人のシフトだけを返す
  if (e && e.parameter && e.parameter.cal) return calendarFeed_(String(e.parameter.cal));
  // スタッフが「できたシフトを見る」ページ（?shift=週シート名。週の指定が無くてもよい）
  if (e && e.parameter && e.parameter.shift) return shiftPage_(String(e.parameter.shift));
  // 検体キットを出した人が押すページ（?kit=2026-09）
  if (e && e.parameter && e.parameter.kit) return kitPage_(String(e.parameter.kit), lineIdOf_(e), lineErrOf_(e));
  // 店長だけが見る、締め期間ごとの人件費（?boss=1）
  if (e && e.parameter && e.parameter.boss) return bossPage_(lineIdOf_(e));
  var config;
  try {
    config = cachedConfig_();
  } catch (e) {
    console.error('設定の読み込みに失敗: ' + e.message);
    config = { closed: true };
  }
  if (!config.storeName) config.storeName = STORE_NAME;
  var template = HtmlService.createTemplateFromFile('index');
  template.configJson = JSON.stringify(config).replace(/</g, '\u003c');
  template.lineIdJson = JSON.stringify(lineIdOf_(e));
  var out = template.evaluate()
    .setTitle(linkTitle_(config))
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    // 他のページに埋め込んで表示すると、Google の「Apps Script ユーザーが作成」の注意書きが出なくなるため許可する
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return out;   // 説明文(description・og:description)はGoogleに取り除かれるので、案内は見出しに入れる
}

// LINEなどにリンクを貼った時に出る見出し。「シフト希望 10/19〜11/15 締切11/1(日)23:59｜店名」
function linkTitle_(config) {
  var title = 'シフト希望';
  if (!config.closed && config.startDate) {
    var end = addDays_(config.startDate, config.weekCount * 7 - 1);
    title += ' ' + mdText_(config.startDate) + '〜' + mdText_(end);
    // 締切は画面の中に大きく出しているので、題名には入れない（LINEの履歴で読みやすくする）
  }
  return title + '｜' + config.storeName;
}

function mdText_(ymd) {
  var p = ymd.split('-').map(Number);
  return p[1] + '/' + p[2];
}

// LINEからの連絡を受け取る入り口（このフォームと同じアドレス）。
// シフト表の設定で「登録を受け付ける」を押してから10分の間だけ、送られてきたトークのIDを送り先として覚える。
// （いつでも受け付けると、アドレスを知った人に送り先を書き換えられてしまうため）
function doPost(e) {
  try {
    var events = JSON.parse(e.postData.contents).events || [];
    // LINEの「検証」ボタンは中身が空のデータを送る。ここで返事が遅れるとタイムアウトになるので、すぐ返す
    if (!events.length) return ContentService.createTextOutput('OK');
    var settings = SpreadsheetApp.openById(SHIFT_SPREADSHEET_ID).getSheetByName(REQUEST_SETTINGS_SHEET);
    if (!settings) return ContentService.createTextOutput('OK');
    // 期限のうしろに |test が付いていたら、本番ではなく「お試しの送り先」として登録する
    var raw = String(settings.getRange(LINE_REGISTER_CELL).getDisplayValue()).trim();
    var parts = raw.split('|');
    var until = String(parts[0]).trim();
    var forTest = String(parts[1] || '').trim() === 'test';
    var open = /^\d{4}-\d{2}-\d{2} \d{2}:\d{2}$/.test(until) && jstTime_(until) > Date.now();
    events.forEach(function(event) {
      var source = event.source || {};
      var id = source.groupId || source.roomId || source.userId;
      if (!id || !open) return;
      settings.getRange(forTest ? LINE_TEST_TARGET_CELL : LINE_TARGET_CELL).setValue(id);
      settings.getRange(LINE_REGISTER_CELL).setValue('');   // 1回登録したら受付を閉じる
      var token = String(settings.getRange(LINE_TOKEN_CELL).getDisplayValue()).trim();
      if (token && event.replyToken) {
        lineReply_(token, event.replyToken, forTest
          ? 'お試しの送り先として、このトークを登録しました。「お試し公開」のときだけ、ここに届きます。'
          : 'シフト希望のお知らせ先として、このトークを登録しました。');
      }
    });
  } catch (err) {
    console.error('doPost: ' + err.message);
  }
  return ContentService.createTextOutput('OK');
}

function lineReply_(token, replyToken, text) {
  UrlFetchApp.fetch(LINE_REPLY_URL, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify({ replyToken: replyToken, messages: [{ type: 'text', text: text }] }),
    muteHttpExceptions: true
  });
}

// GASエディタから実行すると、記録先の準備（シート作成）とGoogleの許可を行う
function setup() {
  var ss = getStore_();
  Logger.log('記録用スプレッドシート: ' + ss.getName() + '  ' + ss.getUrl());
}

// GASエディタから実行：自動確認で入った記録を今すぐ全部消す（ふだんは cleanupTestRecords_ が自動で消す）
function removeTestRecords() {
  var counts = withLock_(function() {
    return cleanupTestRecords_(getStore_(), Date.now() + 1);
  });
  Logger.log('消した行数: ' + JSON.stringify(counts));
}

function isTestBrowser_(ua) {
  return String(ua).indexOf(TEST_BROWSER_MARK) !== -1;
}

// cutoffMs より前の、自動確認の記録（記録2枚・その端末の持ち主の行・保留中の希望）を消す。スタッフの記録は消さない
function cleanupTestRecords_(ss, cutoffMs) {
  var removedDevices = {};
  var keptDevices = {};
  var result = {};

  [SHEETS.log, SHEETS.submit].forEach(function(def) {
    var sheet = sheet_(ss, def.name);
    var uaIndex = def.headers.indexOf('ブラウザ情報');
    var deviceIndex = def.headers.indexOf('端末ID');
    var last = sheet.getLastRow();
    result[def.name] = 0;
    if (last < 2) return;
    var values = sheet.getRange(2, 1, last - 1, def.headers.length).getValues();
    for (var i = values.length - 1; i >= 0; i--) {
      if (!isTestBrowser_(values[i][uaIndex])) continue;
      var device = String(values[i][deviceIndex]);
      var time = values[i][0] instanceof Date ? values[i][0].getTime() : 0;
      if (time >= cutoffMs) { keptDevices[device] = true; continue; }
      removedDevices[device] = true;
      sheet.deleteRow(i + 2);
      result[def.name]++;
    }
  });

  // まだ新しい記録が残っている端末は、持ち主の行と保留を残す
  var targets = {};
  Object.keys(removedDevices).forEach(function(d) { if (!keptDevices[d]) targets[d] = true; });
  var hasTargets = Object.keys(targets).length > 0;

  result[SHEETS.devices.name] = hasTargets
    ? deleteRowsWhere_(sheet_(ss, SHEETS.devices.name), 2, 0, function(r) { return targets[String(r[0])]; })
    : 0;
  result[REQUEST_SETTINGS_SHEET] = 0;
  if (hasTargets) {
    var settings = SpreadsheetApp.openById(SHIFT_SPREADSHEET_ID).getSheetByName(REQUEST_SETTINGS_SHEET);
    if (settings) {
      result[REQUEST_SETTINGS_SHEET] = deleteRowsWhere_(settings, PENDING_FIRST_ROW, PENDING_DEVICE_INDEX, function(r) {
        return targets[String(r[PENDING_DEVICE_INDEX])];
      });
    }
  }
  return result;
}

function deleteRowsWhere_(sheet, firstRow, lastIndex, match) {
  var last = sheet.getLastRow();
  if (last < firstRow) return 0;
  var values = sheet.getRange(firstRow, 1, last - firstRow + 1, lastIndex + 1).getValues();
  var removed = 0;
  for (var i = values.length - 1; i >= 0; i--) {
    if (!match(values[i])) continue;
    sheet.deleteRow(firstRow + i);
    removed++;
  }
  return removed;
}

/* ================= フォームから呼ばれる ================= */

// 画面を開いた・名前を選択・名前を変更 の記録（失敗してもフォームの操作は止めない）
function logEvent(info) {
  info = info || {};
  var action = EVENT_ACTIONS[info.action];
  var device = validDeviceId_(info.deviceId);
  if (!action || !device) return { ok: false };
  var name = info.action === 'open' ? '' : cleanText_(info.name, 40);
  if (info.action !== 'open' && !name) return { ok: false };

  return withLock_(function() {
    if (overLimit_('ev:' + device, LIMITS.event)) return { ok: false };
    writeRecord_(getStore_(), {
      action: action,
      name: name,
      detail: info.action === 'change' ? '変更前の名前' : '',
      device: device, ip: cleanIp_(info.ip), model: cleanText_(info.model, 60), ua: cleanText_(info.ua, 300)
    });
    return { ok: true };
  });
}

// 名前を選んだ人の今の希望（週シートに入っている時間、シートが無い週は保留中の希望）を返す
function getMyRequests(request) {
  request = request || {};
  var device = validDeviceId_(request.deviceId);
  if (!device) throw new Error('端末の情報が読み取れませんでした。ページを開き直してください。');

  return withLock_(function() {
    if (overLimit_('rd:' + device, LIMITS.event)) throw new Error('短い時間に何回も読み込まれました。少し待ってから開き直してください。');
    var shiftSs = SpreadsheetApp.openById(SHIFT_SPREADSHEET_ID);
    var config = loadConfig_(shiftSs);
    checkConfig_(config, request.startDate);
    var name = String(request.name || '');
    if (config.staff.indexOf(name) === -1) throw new Error('名前が名簿にありません。最初の画面から選び直してください。');
    var lineId = String(request.lineId || '');
    // 金額を出してよいかは、結びつけより先に決める。
    // その場で結びつけた名前を「本人」と認めると、他人の名前を選ぶだけで金額が見えてしまう
    var canPay = payVisible_(lineId, device, name);
    if (lineId && !lineOwner_(lineId)) bindMe(lineId, name);   // まだ結びついていない人だけ覚える
    return {
      weeks: readWeekStates_(shiftSs, config, name).map(function(s) {
        return { shifts: s.shifts, note: s.note };
      }),
      totals: closingPeriods_(shiftSs, config, name, canPay),
      recruits: openRecruits_(config.sheetNames, name)   // 自分が応募ずみかも分かる形で
    };
  });
}

// 希望の送信。変えた週だけシフト表に書き込み（シートが無い週は保留）、記録は送るたびに1行ずつ残す
function submitRequest(payload) {
  payload = payload || {};
  var rec = {
    name: cleanText_(payload.name, 40),
    device: validDeviceId_(payload.deviceId),
    ip: cleanIp_(payload.ip), model: cleanText_(payload.model, 60), ua: cleanText_(payload.ua, 300)
  };

  try {
    return withLock_(function() {
      var shiftSs = SpreadsheetApp.openById(SHIFT_SPREADSHEET_ID);
      var config = loadConfig_(shiftSs);
      checkConfig_(config, payload.startDate);
      if (config.readOnly) throw new Error('締切を過ぎています。次の募集が始まるまでお待ちください。');
      if (config.staff.indexOf(rec.name) === -1) throw new Error('名前が名簿にありません。最初の画面から選び直してください。');
      if (!rec.device) throw new Error('端末の情報が読み取れませんでした。ページを開き直してください。');

      var states = readWeekStates_(shiftSs, config, rec.name);
      var parsed = parseSubmission_(payload, config, states);
      if (overLimit_('sub:d:' + rec.device, LIMITS.submit) || overLimit_('sub:n:' + rec.name, LIMITS.submit)) {
        throw new Error('短い時間に何回も送信されたため、受付を止めています。5分ほど待ってから送ってください。');
      }

      var now = new Date();
      var nowText = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
      var settings = sheet_(shiftSs, REQUEST_SETTINGS_SHEET);
      var result = { saved: [], waiting: [], notReflected: [] };
      var outcome = {};

      parsed.dirtyWeeks.forEach(function(w) {
        var state = states[w];
        var values = buildRowValues_(parsed.shifts[w], parsed.notes[w]);
        var label = (w + 1) + '週';
        if (!state.writable) {
          result.notReflected.push(label);
          outcome[w] = '→ 反映できず（' + FIRST_WRITABLE_SHEET + 'より前の週）';
        } else if (state.sheet && state.row > 0) {
          state.sheet.getRange(state.row, 2, 1, INPUT_COLS).setValues([values]);
          result.saved.push(label);
          outcome[w] = '→ シート' + state.sheetName + 'に反映';
        } else if (state.sheet) {
          result.notReflected.push(label);
          outcome[w] = '→ 反映できず（シート' + state.sheetName + 'に名前が無い）';
        } else {
          upsertPending_(settings, state.sheetName, rec.name, values, nowText, rec.device);
          result.waiting.push(label);
          outcome[w] = '→ シート' + state.sheetName + 'の作成時に反映';
        }
      });

      // 変えていない週は今の内容のまま、合計と記録に使う
      var finalShifts = states.map(function(s, w) { return parsed.dirtyWeeks.indexOf(w) !== -1 ? parsed.shifts[w] : s.shifts; });
      var finalNotes = states.map(function(s, w) { return parsed.dirtyWeeks.indexOf(w) !== -1 ? parsed.notes[w] : s.note; });
      var totals = sumShifts_(finalShifts);

      var extra = [];
      if (config.pastDeadline) extra.push('締切後の送信');
      if (result.notReflected.length) extra.push('シートに反映できない週：' + result.notReflected.join('・'));

      rec.action = ACTION_SUBMIT;
      rec.detail = totals.days + '日・' + totals.hours + '時間';
      rec.extraWarning = extra.join(' ／ ');
      var store = getStore_();
      var warning = writeRecord_(store, rec, now);

      var sheet = sheet_(store, SHEETS.submit.name);
      sheet.appendRow(safeRow_([now, '', rec.name, warning, totals.days, totals.hours]
        .concat(weekSummaries_(config, finalShifts, finalNotes, outcome))
        .concat([rec.device, rec.ip, rec.model, rec.ua, JSON.stringify({
          startDate: config.startDate, weekCount: config.weekCount, dirtyWeeks: parsed.dirtyWeeks,
          shifts: finalShifts, notes: finalNotes
        })])));
      setOwnerFormula_(sheet, SUBMIT_DEVICE_COL);

      result.ok = true;
      result.receivedAt = Utilities.formatDate(now, 'Asia/Tokyo', 'M/d HH:mm');
      return result;
    });
  } catch (e) {
    recordFailure_(rec, e.message);
    throw e;
  }
}

/* ================= シフト表の読み書き ================= */

// 募集設定・名簿・祝日をシフト表から読む。設定が無ければ募集していない扱い
function loadConfig_(shiftSs) {
  var tz = shiftSs.getSpreadsheetTimeZone();
  var settings = shiftSs.getSheetByName(REQUEST_SETTINGS_SHEET);
  if (!settings) return { closed: true };

  var period = currentPeriod_(settings, tz, new Date());
  if (!period) return { closed: true };
  var start = period.start, weekCount = period.weeks, deadline = period.deadline;

  var newest = shiftSs.getSheets()
    .map(function(s) { return s.getName(); })
    .filter(function(n) { return /^\d{6}$/.test(n); })
    .sort()
    .pop();
  var staff = newest ? rosterNames_(shiftSs.getSheetByName(newest)) : [];
  if (!staff.length) return { closed: true };

  var sheetNames = [];
  for (var w = 0; w < weekCount; w++) sheetNames.push(sheetNameOf_(addDays_(start, w * 7)));

  var validDeadline = /^\d{4}-\d{2}-\d{2} \d{2}:\d{2}$/.test(deadline);
  return {
    closed: false,
    storeName: String(settings.getRange(SETTING_STORE_NAME_CELL).getDisplayValue()).trim() || STORE_NAME,
    startDate: start,
    weekCount: weekCount,
    sheetNames: sheetNames,
    deadline: validDeadline ? deadlineLabel_(deadline) : '',
    pastDeadline: validDeadline && Utilities.parseDate(deadline, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm').getTime() < Date.now(),
    readOnly: !!period.readOnly,          // 締切後の「見るだけ」。送信も時間の確定もできない
    holidays: holidaysInPeriod_(shiftSs, start, weekCount),
    staff: staff,
    liffId: String(settings.getRange(LIFF_ID_CELL).getDisplayValue()).trim(),
    noteChoices: NOTE_CHOICES,
    timeRange: timeRange_(settings),
    recruits: openRecruits_(sheetNames, '')
  };
}

function checkConfig_(config, startDate) {
  if (config.closed) throw new Error('現在、シフト希望の募集はありません。');
  if (String(startDate || '') !== config.startDate) throw new Error('募集期間が変わりました。ページを開き直してから入力してください。');
}

// 今の募集期間。今回だけ手で決めた期間が締切前ならそれを使い、無ければ基本設定から自動で決める
// ※シフト表GAS（currentRequestPeriod_）と同じ決まり。変える時は両方を直すこと
function currentPeriod_(settings, tz, now) {
  var start = settingText_(settings.getRange(SETTING_START_CELL).getValue(), tz, 'yyyy-MM-dd');
  var weeks = Number(settings.getRange(SETTING_WEEKS_CELL).getValue());
  var deadline = settingText_(settings.getRange(SETTING_DEADLINE_CELL).getValue(), tz, 'yyyy-MM-dd HH:mm');
  if (/^\d{4}-\d{2}-\d{2}$/.test(start) && weeks >= 1 && weeks <= 5 &&
      /^\d{4}-\d{2}-\d{2} \d{2}:\d{2}$/.test(deadline) && jstTime_(deadline) > now.getTime()) {
    return { start: start, weeks: weeks, deadline: deadline, manual: true };
  }
  var pair = autoPeriods_(readBase_(settings), now);
  // 締切の瞬間に次の月へ切り替わるとスタッフが戸惑うので、
  // 次の募集のお知らせ（1回目）が届くまでは、締め切ったばかりの期間を「見るだけ」で出す
  if (pair.next && pair.closed && now.getTime() < noticeStartMs_(settings, pair.next)) {
    return { start: pair.closed.start, weeks: pair.closed.weeks, deadline: pair.closed.deadline,
      manual: false, readOnly: true };
  }
  return pair.next;
}

function readBase_(settings) {
  var closing = Number(settings.getRange(BASE_CLOSING_DAY_CELL).getDisplayValue());
  var weeksText = String(settings.getRange(BASE_DEADLINE_WEEKS_CELL).getDisplayValue()).trim();
  var time = String(settings.getRange(BASE_DEADLINE_TIME_CELL).getDisplayValue()).trim();
  return {
    closingDay: closing >= 1 && closing <= 31 ? Math.floor(closing) : DEFAULT_BASE.closingDay,
    deadlineWeeks: Number(weeksText) >= 1 && Number(weeksText) <= 4 ? Math.floor(Number(weeksText)) : DEFAULT_BASE.deadlineWeeks,
    deadlineTime: /^\d{2}:\d{2}$/.test(time) ? time : DEFAULT_BASE.deadlineTime
  };
}

// 自動の募集期間：前の期間の最終日の翌日（月曜）〜締め日が入っている週の日曜日。
// スタッフ入力締切＝募集が始まる◯週間前の日曜日（初日の前日の日曜から◯−1週さかのぼる）。
// 締切を過ぎた期間と、書き込まない週（FIRST_WRITABLE_SHEETより前）から始まる期間は飛ばす
function autoPeriod_(base, now) {
  return autoPeriods_(base, now).next;
}

// 次に募集する期間と、締切を過ぎたばかりの期間（画面に出し続けるほう）の両方を返す
function autoPeriods_(base, now) {
  var ym = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM').split('-').map(Number);
  var closed = null;
  for (var i = -2; i <= 13; i++) {
    var end = periodEnd_(ym[0], ym[1] + i, base.closingDay);
    var start = addDays_(periodEnd_(ym[0], ym[1] + i - 1, base.closingDay), 1);
    if (sheetNameOf_(start) < FIRST_WRITABLE_SHEET) continue;
    var deadline = addDays_(start, -(base.deadlineWeeks * 7 + 1)) + ' ' + base.deadlineTime;
    var one = { start: start, weeks: (daysBetween_(start, end) + 1) / 7, deadline: deadline, manual: false };
    if (jstTime_(deadline) <= now.getTime()) { closed = one; continue; }   // 締切ずみ。直近のものを覚えておく
    return { next: one, closed: closed };
  }
  return { next: null, closed: closed };
}

// 次の募集の「1回目のお知らせ」を送る日（ここで画面が新しい期間に切り替わる）
var NOTICE_DAYS_CELL_FORM = 'E13';     // お知らせ：送る日（締切の何日前か。シフト表GASが書く）
var NOTICE_FIRST_DAYS_DEFAULT = 14;
function noticeStartMs_(settings, next) {
  var days = String(settings.getRange(NOTICE_DAYS_CELL_FORM).getDisplayValue()).split(/[,、\s]+/)
    .map(Number).filter(function(n) { return n >= 0 && n <= 60 && Math.floor(n) === n; });
  var first = days.length ? Math.max.apply(null, days) : NOTICE_FIRST_DAYS_DEFAULT;
  return jstTime_(next.deadline) - first * 86400000;
}

// その月の締め日が入っている週の日曜日（締め日が日曜ならその日）。月の日数を超える締め日は末日にする
function periodEnd_(year, month, closingDay) {
  var lastDay = new Date(Date.UTC(year, month, 0)).getUTCDate();
  var d = new Date(Date.UTC(year, month - 1, Math.min(closingDay, lastDay)));
  d.setUTCDate(d.getUTCDate() + (7 - d.getUTCDay()) % 7);
  return d.getUTCFullYear() + '-' + pad2_(d.getUTCMonth() + 1) + '-' + pad2_(d.getUTCDate());
}

function daysBetween_(fromYmd, toYmd) {
  var a = fromYmd.split('-').map(Number), b = toYmd.split('-').map(Number);
  return Math.round((Date.UTC(b[0], b[1] - 1, b[2]) - Date.UTC(a[0], a[1] - 1, a[2])) / 86400000);
}

function jstTime_(text) {
  return Utilities.parseDate(text, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm').getTime();
}

function settingText_(value, tz, pattern) {
  return value instanceof Date ? Utilities.formatDate(value, tz, pattern) : String(value == null ? '' : value).trim();
}

function rosterNames_(sheet) {
  return sheet.getRange(ROSTER_START_ROW, 1, ROSTER_ROWS, 1).getValues()
    .map(function(r) { return String(r[0]).trim(); })
    .filter(function(n) { return n; });
}

// すでに読んである中身（A列をふくむ）から、名簿（A4〜A25）を取り出す
function rosterFromGrid_(values) {
  var out = [];
  for (var i = 0; i < ROSTER_ROWS; i++) {
    var row = values[ROSTER_START_ROW - 1 + i] || [];
    var name = String(row[0] == null ? '' : row[0]).trim();
    if (name) out.push(name);
  }
  return out;
}

// 「祝日」シートの日付は、シフト表GASが日本時間の0時として書き込んでいる。
// シフト表の時間帯（アメリカ西海岸）で読むと前日になるため、日本時間で読む（シフト表GASの色付けと同じ）
function holidaysInPeriod_(shiftSs, start, weekCount) {
  var sheet = shiftSs.getSheetByName(HOLIDAY_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var end = addDays_(start, weekCount * 7 - 1);
  var out = [];
  sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().forEach(function(r) {
    if (!(r[0] instanceof Date)) return;
    var ymd = Utilities.formatDate(r[0], 'Asia/Tokyo', 'yyyy-MM-dd');
    if (ymd < start || ymd > end) return;
    var p = ymd.split('-').map(Number);
    out.push(p[1] + '-' + p[2]);
  });
  return out;
}



/* ================= 足りないところの募集（スタッフ側） =================
   シフト表GASが「募集と応募」シート（アクセス記録）に書いた枠を読み、
   シフト希望の画面に赤く出す。応募はこのシートの「応募」列に積む。
   誰が応募したかは、ほかのスタッフには見せない（人数だけ）。                        */
var RECRUIT_SHEET_NAME = '募集と応募';
var RECRUIT_COLS = { id: 1, week: 2, ymd: 3, dow: 4, from: 5, to: 6, people: 7, state: 8,
  applied: 9, filled: 10, made: 11, kind: 12, choices: 13, targets: 14 };
var RECRUIT_READ_COLS = 14;
var RECRUIT_OPEN = '募集中';
var RECRUIT_KIND_STRETCH = '延ばす';   // その日もう入っている人に、早く来る／遅くまでを頼む枠

function recruitSheet_() {
  try {
    return getStore_().getSheetByName(RECRUIT_SHEET_NAME);
  } catch (err) { return null; }
}

// 「名前|日時」の行を名前の一覧にする
function recruitNames_(text) {
  return String(text == null ? '' : text).split('\n')
    .map(function(line) { return String(line).split('|')[0].trim(); })
    .filter(function(n) { return n; });
}

// 候補の時間。シートには「16:00-21:15」の行で入っている
function recruitChoices_(text, from, to) {
  var out = [];
  String(text == null ? '' : text).split('\n').forEach(function(line) {
    var m = /^\s*(\d{1,2}:\d{2})\s*-\s*(\d{1,2}:\d{2})\s*$/.exec(line);
    if (m) out.push({ from: m[1], to: m[2], text: m[1] + '-' + m[2] });
  });
  if (!out.length && from && to) out.push({ from: from, to: to, text: from + '-' + to });
  return out;
}

// 「名前|日時|16:00-21:15」の行を細かく分ける
function recruitEntries_(text) {
  return String(text == null ? '' : text).split('\n').map(function(line) {
    var p = String(line).split('|');
    return { name: String(p[0]).trim(), at: String(p[1] || '').trim(), choice: String(p[2] || '').trim() };
  }).filter(function(x) { return x.name; });
}

// いま募集している枠。weekNames を渡すとその週だけ。name を渡すと「自分が応募ずみか」も付ける
function openRecruits_(weekNames, name) {
  var sheet = recruitSheet_();
  if (!sheet || sheet.getLastRow() < 2) return [];
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, RECRUIT_READ_COLS).getDisplayValues();
  var key = name ? normalizeName_(name) : '';
  var out = [];
  rows.forEach(function(r, i) {
    if (String(r[RECRUIT_COLS.state - 1]).trim() !== RECRUIT_OPEN) return;
    var week = String(r[RECRUIT_COLS.week - 1]).trim();
    if (weekNames && weekNames.length && weekNames.indexOf(week) === -1) return;
    var kind = String(r[RECRUIT_COLS.kind - 1]).trim();
    var targets = recruitNames_(r[RECRUIT_COLS.targets - 1]);
    // 「今いる人に頼む」枠は、声をかける人にだけ見せる（名前を渡さない店長用の読みは全部返す）
    if (kind === RECRUIT_KIND_STRETCH && key &&
        !targets.some(function(n) { return normalizeName_(n) === key; })) return;
    var from = String(r[RECRUIT_COLS.from - 1]).trim();
    var to = String(r[RECRUIT_COLS.to - 1]).trim();
    var entries = recruitEntries_(r[RECRUIT_COLS.applied - 1]);
    var filled = recruitNames_(r[RECRUIT_COLS.filled - 1]);
    var people = Number(r[RECRUIT_COLS.people - 1]) || 1;
    var mine = key ? entries.filter(function(e) { return normalizeName_(e.name) === key; })[0] : null;
    out.push({
      row: i + 2,
      id: String(r[RECRUIT_COLS.id - 1]).trim(),
      ymd: String(r[RECRUIT_COLS.ymd - 1]).trim(),
      dow: String(r[RECRUIT_COLS.dow - 1]).trim(),
      from: from, to: to,
      kind: kind,
      stretch: kind === RECRUIT_KIND_STRETCH,
      targets: targets,
      choices: recruitChoices_(r[RECRUIT_COLS.choices - 1], from, to),
      people: people,
      left: Math.max(0, people - filled.length),
      applied: entries.length,                       // 人数だけ。名前はスタッフに出さない
      mine: !!mine,
      myChoice: mine ? mine.choice : ''
    });
  });
  return out;
}

// 画面から呼ぶ：この枠に「入れます」。押しただけでは確定しない（店長が入れる）
// ★末尾に _ を付けると画面から呼べなくなるため付けない
function applyRecruit(request) {
  request = request || {};
  var device = validDeviceId_(request.deviceId);
  if (!device) throw new Error('端末の情報が読み取れませんでした。ページを開き直してください。');
  return withLock_(function() {
    if (overLimit_('rc:' + device, LIMITS.event)) throw new Error('短い時間に何回も押されました。少し待ってから試してください。');
    var shiftSs = SpreadsheetApp.openById(SHIFT_SPREADSHEET_ID);
    var config = loadConfig_(shiftSs);
    if (config.closed) throw new Error('現在、シフト希望の募集はありません。');
    var name = String(request.name || '');
    if (config.staff.indexOf(name) === -1) throw new Error('名前が名簿にありません。最初の画面から選び直してください。');
    var sheet = recruitSheet_();
    if (!sheet) throw new Error('いま募集している枠はありません。');
    var id = String(request.id || '');
    var hit = openRecruits_(config.sheetNames, name).filter(function(r) { return r.id === id; })[0];
    if (!hit) throw new Error('この枠は締め切られました。画面を開き直してください。');
    if (hit.mine) return { ok: true, already: true };
    // どの時間で入れるかを選んでもらっている。選ばれていなければ最初の候補
    var want = String(request.choice || '');
    var pick = hit.choices.filter(function(c) { return c.text === want; })[0] || hit.choices[0];
    if (!pick) throw new Error('時間が分かりませんでした。画面を開き直してください。');
    var cell = sheet.getRange(hit.row, RECRUIT_COLS.applied);
    var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
    var text = String(cell.getDisplayValue()).trim();
    cell.setValue((text ? text + '\n' : '') + name + '|' + now + '|' + pick.text);
    logEvent({ deviceId: device, ip: request.ip, model: request.model, ua: request.ua,
      name: name, action: 'apply' });
    return { ok: true, already: false, choice: pick.text };
  });
}


// 右の日別シフト表（名前の列）。シフト表GASの DAY_BLOCKS と同じ場所。月〜日の順
var DAY_NAME_RANGES = ['T24:T39', 'T43:T58', 'T62:T77', 'DD4:DD19', 'DD24:DD39', 'DD43:DD58', 'DD62:DD77'];

// 週シート名（261019）から、その週の月曜日（2026-10-19）
function weekStartYmdOf_(sheetName) {
  return '20' + sheetName.slice(0, 2) + '-' + sheetName.slice(2, 4) + '-' + sheetName.slice(4, 6);
}

// 店長用：いま募集している枠と、応募した人（名前つき）
function bossRecruits(lineId) {
  var shiftSs = SpreadsheetApp.openById(SHIFT_SPREADSHEET_ID);
  if (!bossNameOf_(shiftSs, lineId)) return { ok: false };
  var sheet = recruitSheet_();
  if (!sheet || sheet.getLastRow() < 2) return { ok: true, slots: [] };
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, RECRUIT_READ_COLS).getDisplayValues();
  var slots = [];
  rows.forEach(function(r) {
    if (String(r[RECRUIT_COLS.state - 1]).trim() !== RECRUIT_OPEN) return;
    var ymd = String(r[RECRUIT_COLS.ymd - 1]).trim();
    var filled = recruitNames_(r[RECRUIT_COLS.filled - 1]);
    var people = Number(r[RECRUIT_COLS.people - 1]) || 1;
    var from = String(r[RECRUIT_COLS.from - 1]).trim();
    var to = String(r[RECRUIT_COLS.to - 1]).trim();
    var choices = recruitChoices_(r[RECRUIT_COLS.choices - 1], from, to);
    slots.push({
      id: String(r[RECRUIT_COLS.id - 1]).trim(),
      week: String(r[RECRUIT_COLS.week - 1]).trim(),
      ymd: ymd, label: ymdLabel_(ymd), dow: String(r[RECRUIT_COLS.dow - 1]).trim(),
      from: from, to: to,
      kind: String(r[RECRUIT_COLS.kind - 1]).trim(),
      stretch: String(r[RECRUIT_COLS.kind - 1]).trim() === RECRUIT_KIND_STRETCH,
      targets: recruitNames_(r[RECRUIT_COLS.targets - 1]),
      choices: choices,
      people: people, left: Math.max(0, people - filled.length), filled: filled,
      applied: recruitEntries_(r[RECRUIT_COLS.applied - 1])
        .filter(function(a) { return filled.indexOf(a.name) === -1; })
        .map(function(a) {
          var pick = choices.filter(function(c) { return c.text === a.choice; })[0] || choices[0];
          return { name: a.name, at: a.at, choice: a.choice,
            time: pick ? pick.from + '〜' + pick.to : '' };
        })
    });
  });
  slots.sort(function(a, b) { return a.ymd < b.ymd ? -1 : (a.ymd > b.ymd ? 1 : 0); });
  return { ok: true, slots: slots };
}

// 店長用：応募した人をその枠に入れる。
// 左の入力表（転写しても消えない）と、右の日別シフト表（すぐ見える）の両方に書く
function fillRecruit(lineId, id, name, choice) {
  var shiftSs = SpreadsheetApp.openById(SHIFT_SPREADSHEET_ID);
  if (!bossNameOf_(shiftSs, lineId)) return { ok: false, message: 'この操作は店長だけができます。' };
  return withLock_(function() {
    var sheet = recruitSheet_();
    if (!sheet) return { ok: false, message: '募集がありません。' };
    var hit = openRecruits_(null, '').filter(function(r) { return r.id === String(id); })[0];
    if (!hit) return { ok: false, message: 'この枠はもう終わっています。' };
    var weekName = String(sheet.getRange(hit.row, RECRUIT_COLS.week).getDisplayValue()).trim();
    var weekSheet = shiftSs.getSheetByName(weekName);
    if (!weekSheet) return { ok: false, message: '週のシート（' + weekName + '）が見つかりません。' };
    var day = daysBetween_(weekStartYmdOf_(weekName), hit.ymd);
    if (!(day >= 0 && day <= 6)) return { ok: false, message: '日付がこの週に入っていません。' };
    var row = findRosterRow_(weekSheet, String(name));
    if (row === -1) return { ok: false, message: String(name) + ' さんが名簿にありません。' };

    // その人が選んだ時間を使う（店長が別の候補を指しても、そちらを使う）
    var entry = recruitEntries_(sheet.getRange(hit.row, RECRUIT_COLS.applied).getDisplayValue())
      .filter(function(e) { return normalizeName_(e.name) === normalizeName_(String(name)); })[0];
    var want = String(choice || (entry && entry.choice) || '');
    var pick = hit.choices.filter(function(c) { return c.text === want; })[0] || hit.choices[0];
    if (!pick) return { ok: false, message: '時間が分かりません。' };
    var from = toSheetNumber_(pick.from), to = toSheetNumber_(pick.to);
    // ① 左の入力表（ここに入れておけば、もう一度転写しても消えない）
    weekSheet.getRange(row, 2 + day * 2, 1, 2).setValues([[from, to]]);
    // ② 右の日別シフト表の空いている行
    var names = weekSheet.getRange(DAY_NAME_RANGES[day]);
    var list = names.getValues();
    var at = -1;
    for (var i = 0; i < list.length; i++) {
      var cur = String(list[i][0]).trim();
      if (normalizeName_(cur) === normalizeName_(String(name))) { at = i; break; }
      if (!cur && at === -1) at = i;
    }
    if (at >= 0) {
      weekSheet.getRange(names.getRow() + at, names.getColumn(), 1, 3)
        .setValues([[String(name), from, to]]);
    }

    // ③ 募集の記録を更新。人数ぶん埋まったら終わりにする
    var filledCell = sheet.getRange(hit.row, RECRUIT_COLS.filled);
    var filled = recruitNames_(filledCell.getDisplayValue());
    if (filled.indexOf(String(name)) === -1) filled.push(String(name));
    filledCell.setValue(filled.join('\n'));
    if (filled.length >= hit.people) sheet.getRange(hit.row, RECRUIT_COLS.state).setValue('終了');
    return { ok: true, message: String(name) + ' さんを ' + pick.from + '〜' + pick.to + ' に入れました。',
      left: Math.max(0, hit.people - filled.length) };
  });
}

// 店長用：この枠の募集をやめる
function closeRecruit(lineId, id) {
  var shiftSs = SpreadsheetApp.openById(SHIFT_SPREADSHEET_ID);
  if (!bossNameOf_(shiftSs, lineId)) return { ok: false, message: 'この操作は店長だけができます。' };
  var sheet = recruitSheet_();
  if (!sheet) return { ok: false, message: '募集がありません。' };
  var hit = openRecruits_(null, '').filter(function(r) { return r.id === String(id); })[0];
  if (!hit) return { ok: false, message: 'この枠はもう終わっています。' };
  sheet.getRange(hit.row, RECRUIT_COLS.state).setValue('止めた');
  return { ok: true, message: '募集をやめました。' };
}

/* ================= 店長用：締め期間ごとの人件費（概算） =================
   スマホで見るための、店長だけのページ（?boss=1）。
   ・出勤履歴（週シートの右側から自動で積まれる）を、その人の締め期間ぶん足す
   ・実働時間は休憩を引いたあとの時間。時給は締め期間の初日で決まる
   ・「確定」＝店長が確定して公開した週、または書庫へ移した過ぎた週
   　「作成中」＝まだ公開していない週（交通費はまだ付けない）                        */
var BOSS_ROLE = '店長';               // このページを開けるのは、スタッフ区分がこれの人だけ

// このLINEの人が店長なら、その名前を返す。ちがえば空
function bossNameOf_(shiftSs, lineId) {
  var who = lineOwner_(lineId);
  if (!who) return '';
  var roles = shiftSs.getSheetByName(ROLE_SHEET_NAME);
  if (!roles || roles.getLastRow() < 2) return '';
  var rows = roles.getRange(2, 1, roles.getLastRow() - 1, 2).getValues();
  var key = normalizeName_(who);
  for (var i = 0; i < rows.length; i++) {
    if (normalizeName_(rows[i][0]) === key) {
      return String(rows[i][1]).trim() === BOSS_ROLE ? who : '';
    }
  }
  return '';
}

// 名前ごとの区分・時給を1回で読む（人数ぶんシートを読まないように）
function bossPayInfos_(shiftSs) {
  var out = {};
  var sheet = shiftSs.getSheetByName(ROLE_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return out;
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, ROLE_WAGE_FROM_COL).getValues();
  rows.forEach(function(r) {
    var name = String(r[0]).trim();
    if (!name) return;
    out[normalizeName_(name)] = {
      role: String(r[1]).trim(),
      wage: Number(r[ROLE_WAGE_COL - 1]) || 0,
      prevWage: Number(r[ROLE_PREV_WAGE_COL - 1]) || 0,
      wageFrom: (r[ROLE_WAGE_FROM_COL - 1] instanceof Date) ? r[ROLE_WAGE_FROM_COL - 1] : null
    };
  });
  return out;
}

// 時間（1230 や Date や "12:30"）を「12:30」の形にする。空なら空
function bossTime_(v) {
  if (v === '' || v === null || v === undefined) return '';
  if (v instanceof Date) return ('0' + v.getHours()).slice(-2) + ':' + ('0' + v.getMinutes()).slice(-2);
  var t = String(v).trim();
  if (/^\d{1,2}:\d{2}$/.test(t)) return ('0' + t).slice(-5);
  var n = Number(t.replace(/[^0-9]/g, ''));
  if (!n && n !== 0) return '';
  return ('0' + Math.floor(n / 100)).slice(-2) + ':' + ('0' + (n % 100)).slice(-2);
}

function bossYen_(n) {
  return Math.round(n);
}

// 店長用のまとめ。ym は 'yyyy-MM'（空なら今月）
function bossSummary_(shiftSs, ym) {
  var settings = shiftSs.getSheetByName(REQUEST_SETTINGS_SHEET);
  var tz = shiftSs.getSpreadsheetTimeZone();
  var todayYmd = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var month = /^\d{4}-\d{2}$/.test(String(ym)) ? String(ym) : todayYmd.slice(0, 7);
  var y = Number(month.slice(0, 4)), m = Number(month.slice(5, 7));
  var rule = readFareRule_(settings);
  var published = publishedWeeks_(shiftSs);
  var inBook = shiftSs.getSheets().map(function(s) { return s.getName(); })
    .filter(function(n) { return /^\d{6}$/.test(n); });
  var infos = bossPayInfos_(shiftSs);
  var empDay = closingDayOf_(settings, true);
  var castDay = closingDayOf_(settings, false);

  // 出勤履歴を1回だけ読む
  var history = shiftSs.getSheetByName(HISTORY_SHEET_NAME);
  var rows = (history && history.getLastRow() >= 2)
    ? history.getRange(2, 1, history.getLastRow() - 1, HISTORY_HEADERS_LEN).getValues() : [];

  var people = {};
  function personOf(name) {
    var key = normalizeName_(name);
    if (people[key]) return people[key];
    var known = !!infos[key];          // スタッフ区分に行があるか（無い人は一覧に出さない）
    var info = infos[key] || { role: '', wage: 0, prevWage: 0, wageFrom: null };
    var employee = EMPLOYEE_ROLES.indexOf(info.role) !== -1;
    var day = employee ? empDay : castDay;
    var end = closingDate_(y, m, day);
    var start = addDays_(closingDate_(y, m - 1, day), 1);
    people[key] = {
      known: known,
      name: String(name).trim(),
      role: info.role || 'キャスト',
      employee: employee,
      wage: wageForPeriod_(info, start, tz),
      start: start, end: end,
      range: ymdLabel_(start) + '〜' + ymdLabel_(end),
      label: closingLabel_(end),
      fixedDays: 0, fixedHours: 0, draftDays: 0, draftHours: 0,
      days: {}, rows: []
    };
    return people[key];
  }

  // 名簿の人は0日でも出す（抜けに気づけるように）。
  // ただしスタッフ区分に載っていない人（時給が分からない人）は出さない
  var roster = [];
  try { roster = cachedConfig_().staff || []; } catch (err) { roster = []; }
  roster.forEach(function(n) { personOf(n); });

  rows.forEach(function(r) {
    if (!(r[0] instanceof Date)) return;
    var name = String(r[2]).trim();
    if (!name) return;
    var hours = Number(r[HISTORY_HOURS_COL - 1]);
    if (!(hours > 0)) return;
    var ymd = dateToYmd_(r[0]);
    var p = personOf(name);
    if (ymd < p.start || ymd > p.end) return;
    var src = String(r[HISTORY_SOURCE_COL - 1]).trim();
    // 確定＝公開ずみの週、または このファイルにもう無い（書庫へ移した過ぎた）週
    var fixed = published.indexOf(src) !== -1 || inBook.indexOf(src) === -1;
    if (!p.days[ymd]) { p.days[ymd] = true; if (fixed) p.fixedDays++; else p.draftDays++; }
    if (fixed) p.fixedHours += hours; else p.draftHours += hours;
    p.rows.push({
      ymd: ymd, label: ymdLabel_(ymd), dow: String(r[1]).trim(),
      start: bossTime_(r[3]), end: bossTime_(r[4]),
      breakStart: bossTime_(r[5]), breakEnd: bossTime_(r[6]),
      hours: Math.round(hours * 100) / 100,
      pay: bossYen_(p.wage * hours),
      fixed: fixed
    });
  });

  var list = [];
  var unknown = [];                   // スタッフ区分に載っていないのに出勤がある人（気づけるように数だけ出す）
  Object.keys(people).forEach(function(k) {
    var p = people[k];
    if (!p.known) {
      if (p.fixedDays + p.draftDays > 0) {
        unknown.push({ name: p.name, days: p.fixedDays + p.draftDays });
      }
      return;                         // スタッフ区分に無い人は一覧に出さない（時給が分からないため）
    }
    p.rows.sort(function(a, b) { return a.ymd < b.ymd ? -1 : 1; });
    var fixedPay = bossYen_(p.wage * p.fixedHours);
    var draftPay = bossYen_(p.wage * p.draftHours);
    // 交通費は確定した出勤日数で計算する。作成中まで入れたときの見込みも一緒に出す
    var fare = p.employee ? 0 : fareFor_(rule, p.fixedDays);
    var fareAll = p.employee ? 0 : fareFor_(rule, p.fixedDays + p.draftDays);
    list.push({
      name: p.name, role: p.role, employee: p.employee, wage: p.wage,
      label: p.label, range: p.range,
      fixedDays: p.fixedDays, fixedHours: Math.round(p.fixedHours * 10) / 10, fixedPay: fixedPay,
      draftDays: p.draftDays, draftHours: Math.round(p.draftHours * 10) / 10, draftPay: draftPay,
      fare: fare, fareAll: fareAll,
      total: fixedPay + draftPay + fareAll,
      rows: p.rows
    });
  });
  list.sort(function(a, b) {
    return (b.total - a.total) || a.name.localeCompare(b.name, 'ja');
  });

  var sum = { people: 0, fixedPay: 0, draftPay: 0, fare: 0, fareAll: 0, total: 0, fixedDays: 0, draftDays: 0,
    fixedHours: 0, draftHours: 0 };
  list.forEach(function(p) {
    if (p.fixedDays + p.draftDays > 0) sum.people++;
    sum.fixedPay += p.fixedPay; sum.draftPay += p.draftPay;
    sum.fare += p.fare; sum.fareAll += p.fareAll; sum.total += p.total;
    sum.fixedDays += p.fixedDays; sum.draftDays += p.draftDays;
    sum.fixedHours += p.fixedHours; sum.draftHours += p.draftHours;
  });
  sum.fixedHours = Math.round(sum.fixedHours * 10) / 10;
  sum.draftHours = Math.round(sum.draftHours * 10) / 10;

  return {
    ok: true, ym: month,
    storeName: String(settings ? settings.getRange(SETTING_STORE_NAME_CELL).getDisplayValue() : '').trim() || STORE_NAME,
    fareKind: rule.kind, fareDays: rule.days, fareAmount: rule.amount,
    unknown: unknown,
    closing: { employee: empDay, cast: castDay },
    people: list, sum: sum
  };
}

// 画面（boss.html）から月を変えたときに呼ぶ。毎回、店長かどうかを確かめる
function bossPay(lineId, ym) {
  var shiftSs = SpreadsheetApp.openById(SHIFT_SPREADSHEET_ID);
  if (!bossNameOf_(shiftSs, lineId)) return { ok: false };
  return bossSummary_(shiftSs, ym);
}

function bossPage_(lineId) {
  var shiftSs = SpreadsheetApp.openById(SHIFT_SPREADSHEET_ID);
  var boss = bossNameOf_(shiftSs, lineId);
  var template = HtmlService.createTemplateFromFile('boss');
  template.lineIdJson = JSON.stringify(String(lineId || ''));
  template.dataJson = JSON.stringify(boss ? bossSummary_(shiftSs, '') : { ok: false })
    .replace(/</g, '\u003c');
  template.recruitJson = JSON.stringify(boss ? bossRecruits(lineId) : { ok: false })
    .replace(/</g, '\u003c');
  var settings = shiftSs.getSheetByName(REQUEST_SETTINGS_SHEET);
  var storeName = String(settings ? settings.getRange(SETTING_STORE_NAME_CELL).getDisplayValue() : '').trim() || STORE_NAME;
  return template.evaluate()
    .setTitle('店長ページ（人件費・募集）｜' + storeName)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ================= 締め期間ごとの合計（参考） ================= */

// 画面に出ている週にかかる、この人の締め期間（店長・社員は社員の締め日、キャストはキャストの締め日で区切る）。
// 表示は「11月20日〆」のように締める日で出す（何日までの分かが分かるように）。
// 店長・社員は休日日数、キャストは出勤日数を数えるので、showOff で伝える。
// 締め期間のうち画面に出ていない日（募集期間より前の、シフトが決まっている日など）の出勤日数は「出勤履歴」から先に数えておく
// （画面に出ている日は、画面の入力で数える。休日日数は「締め期間の日数 − 出勤日数」で画面が出す）
// 金額の目安を見せてよい相手か。LINEで本人と分かっているか、この端末の持ち主がその人のときだけ。
// 名前を選んだだけの人（ほかの人の名前を選んで見ている人）には見せない
function payVisible_(lineId, deviceId, name) {
  var key = normalizeName_(name);
  var fromLine = lineOwner_(lineId);
  if (fromLine && normalizeName_(fromLine) === key) return true;
  var owner = deviceOwner_(deviceId);
  return !!(owner && normalizeName_(owner) === key);
}

// showPay ＝ 金額の目安を出してよいとき（本人と分かっているとき）だけ時給を渡す
function closingPeriods_(shiftSs, config, name, showPay) {
  var settings = shiftSs.getSheetByName(REQUEST_SETTINGS_SHEET);
  var employee = isEmployee_(shiftSs, name);
  var formStart = config.startDate;
  var formEnd = addDays_(config.startDate, config.weekCount * 7 - 1);
  var tzNow = shiftSs.getSpreadsheetTimeZone();
  var payInfo = (employee || !showPay) ? null : staffPayInfo_(shiftSs, name);   // 店長・社員は月給なので金額を出さない
  var periods = closingRanges_(formStart, formEnd, closingDayOf_(settings, employee)).map(function(p) {
    return { start: p.start, end: p.end, label: closingLabel_(p.end), days: daysBetween_(p.start, p.end) + 1,
      fixedDays: 0, fixedHours: 0, wage: wageForPeriod_(payInfo, p.start, tzNow) };
  });
  // 概算の給与は「1日 PAY_DAY_CAP_HOURS 時間まで」で数える（長い希望を出した日があっても現実的な額にする）
  var result = { showOff: employee, periods: periods, capHours: PAY_DAY_CAP_HOURS };

  var history = shiftSs.getSheetByName(HISTORY_SHEET_NAME);
  if (!history || history.getLastRow() < 2) return result;
  var tz = tzNow;                              // 出勤履歴の日付は、シフト表の時間帯の0時で入っている
  var counted = {};
  history.getRange(2, 1, history.getLastRow() - 1, HISTORY_HOURS_COL).getValues().forEach(function(r) {
    if (!(r[0] instanceof Date) || String(r[2]).trim() !== name) return;
    var ymd = Utilities.formatDate(r[0], tz, 'yyyy-MM-dd');
    if (ymd >= formStart && ymd <= formEnd) return;   // 画面に出ている日は、画面の入力のほうで数える
    var period = periods.filter(function(p) { return ymd >= p.start && ymd <= p.end; })[0];
    if (!period) return;
    if (!counted[ymd]) { counted[ymd] = true; period.fixedDays++; }   // 同じ日が2行あっても1日
    var hours = Number(r[HISTORY_HOURS_COL - 1]);                     // 時間のほうは足す
    if (hours > 0) period.fixedHours += hours;
  });
  return result;
}

// 「スタッフ区分」シートに店長・社員として載っている人か（載っていない人はキャスト）
function isEmployee_(shiftSs, name) {
  var roles = shiftSs.getSheetByName(ROLE_SHEET_NAME);
  return !!roles && roles.getLastRow() >= 2 &&
    roles.getRange(2, 1, roles.getLastRow() - 1, 2).getValues().some(function(r) {
      return String(r[0]).trim() === name && EMPLOYEE_ROLES.indexOf(String(r[1]).trim()) !== -1;
    });
}

// 店長・社員は社員の締め日（E3）、キャストはキャストの締め日（E7）
function closingDayOf_(settings, employee) {
  if (employee) return readBase_(settings).closingDay;
  var cast = Number(settings.getRange(CAST_CLOSING_DAY_CELL).getDisplayValue());
  return cast >= 1 && cast <= 31 ? Math.floor(cast) : DEFAULT_CAST_CLOSING_DAY;
}

// fromYmd〜toYmd にかかる締め期間（前の月の締め日の翌日〜その月の締め日）を古い順に
function closingRanges_(fromYmd, toYmd, closingDay) {
  var p = fromYmd.split('-').map(Number);
  var out = [];
  for (var i = 0; i <= 13; i++) {
    var end = closingDate_(p[0], p[1] + i, closingDay);
    var start = addDays_(closingDate_(p[0], p[1] + i - 1, closingDay), 1);
    if (start > toYmd) break;
    if (end >= fromYmd) out.push({ start: start, end: end });
  }
  return out;
}

// 締め期間の見出し。締める日まで出す（例：2026-11-20 → 「11月20日〆」）
function closingLabel_(endYmd) {
  var p = endYmd.split('-');
  return Number(p[1]) + '月' + Number(p[2]) + '日〆';
}

// その月の締め日（月の日数を超える締め日は末日）。month は13や0でもよい
function closingDate_(year, month, closingDay) {
  var first = new Date(Date.UTC(year, month - 1, 1));
  var y = first.getUTCFullYear(), m = first.getUTCMonth() + 1;
  var lastDay = new Date(Date.UTC(y, m, 0)).getUTCDate();
  return y + '-' + pad2_(m) + '-' + pad2_(Math.min(closingDay, lastDay));
}

// 週ごとに、この人の今の内容を読む（週シートがあればシート、無ければ保留中の希望）
function readWeekStates_(shiftSs, config, name) {
  var settings = sheet_(shiftSs, REQUEST_SETTINGS_SHEET);
  applyAllPending_(shiftSs, settings);
  var pending = readPending_(settings);

  return config.sheetNames.map(function(sheetName) {
    var state = {
      sheetName: sheetName, sheet: null, row: -1,
      shifts: {}, note: { choice: '', text: '' },
      writable: sheetName >= FIRST_WRITABLE_SHEET
    };
    var sheet = shiftSs.getSheetByName(sheetName);
    if (sheet) {
      state.sheet = sheet;
      state.row = findRosterRow_(sheet, name);
      if (state.row > 0) {
        var range = sheet.getRange(state.row, 2, 1, INPUT_COLS);
        fillState_(state, range.getValues()[0], range.getDisplayValues()[0]);
      }
    } else {
      var own = pending.filter(function(p) { return p.sheetName === sheetName && p.name === name; })[0];
      if (own) fillState_(state, own.values, []);
    }
    return state;
  });
}

function fillState_(state, values, displays) {
  for (var d = 0; d < 7; d++) {
    var s = sheetTime_(values[d * 2], displays[d * 2]);
    var e = sheetTime_(values[d * 2 + 1], displays[d * 2 + 1]);
    if (s && e) state.shifts[d] = { s: s, e: e };
  }
  state.note = parseNote_(values[14]);
}

function findRosterRow_(sheet, name) {
  var names = sheet.getRange(ROSTER_START_ROW, 1, ROSTER_ROWS, 1).getValues();
  for (var i = 0; i < names.length; i++) {
    if (String(names[i][0]).trim() === name) return ROSTER_START_ROW + i;
  }
  return -1;
}

function readPending_(settings) {
  var last = settings.getLastRow();
  if (last < PENDING_FIRST_ROW) return [];
  return settings.getRange(PENDING_FIRST_ROW, 1, last - PENDING_FIRST_ROW + 1, PENDING_COLS).getValues()
    .map(function(r, i) {
      return { row: PENDING_FIRST_ROW + i, sheetName: String(r[0]), name: String(r[1]).trim(), values: r.slice(2, 2 + INPUT_COLS) };
    })
    .filter(function(p) { return p.sheetName && p.name; });
}

// 週シートがもうできている週の保留中の希望は、シートに入れて保留から消す
// （通常は「次週シート作成」がその場で行う。作成と送信が重なった時の取りこぼし防止）
function applyAllPending_(shiftSs, settings) {
  var pending = readPending_(settings);
  for (var i = pending.length - 1; i >= 0; i--) {
    var p = pending[i];
    if (!/^\d{6}$/.test(p.sheetName) || p.sheetName < FIRST_WRITABLE_SHEET) continue;
    var sheet = shiftSs.getSheetByName(p.sheetName);
    if (!sheet) continue;
    var row = findRosterRow_(sheet, p.name);
    if (row === -1) continue; // 名簿に名前が無い希望は消さずに残す
    sheet.getRange(row, 2, 1, INPUT_COLS).setValues([p.values]);
    settings.deleteRow(p.row);
  }
}

function upsertPending_(settings, sheetName, name, values, nowText, device) {
  var rowValues = [sheetName, name].concat(values).concat([nowText, device]);
  var existing = readPending_(settings).filter(function(p) { return p.sheetName === sheetName && p.name === name; })[0];
  var row = existing ? existing.row : Math.max(settings.getLastRow() + 1, PENDING_FIRST_ROW);
  settings.getRange(row, 1).setNumberFormat('@');   // シート名が数字にならないように
  settings.getRange(row, 1, 1, PENDING_COLS).setValues([safeRow_(rowValues)]);
}

// セルの時間を「09:00」の形にする（900 のような数字・時刻形式・文字のどれでも読む）
function sheetTime_(value, display) {
  if (value === '' || value === null || value === undefined) return '';
  if (typeof value === 'number') {
    var h = Math.floor(value / 100), m = value % 100;
    return (value === Math.floor(value) && h <= 29 && m < 60) ? pad2_(h) + ':' + pad2_(m) : '';
  }
  var shown = /(\d{1,2}):(\d{2})/.exec(String(display || ''));
  if (shown) return pad2_(Number(shown[1])) + ':' + shown[2];
  var text = String(value).trim();
  var colon = /^(\d{1,2}):(\d{2})$/.exec(text);
  if (colon) return pad2_(Number(colon[1])) + ':' + colon[2];
  var digits = /^(\d{3,4})$/.exec(text);
  if (digits) {
    var n = Number(digits[1]);
    return (n % 100 < 60) ? pad2_(Math.floor(n / 100)) + ':' + pad2_(n % 100) : '';
  }
  return '';
}

// シフト表の書き方（出勤・退勤は 900 のような数字、備考は「選択／手入力」）
function buildRowValues_(shifts, note) {
  var values = [];
  for (var d = 0; d < 7; d++) {
    var a = shifts[d];
    values.push(a ? toSheetNumber_(a.s) : '', a ? toSheetNumber_(a.e) : '');
  }
  values.push(note.choice && note.text ? note.choice + NOTE_SEPARATOR + note.text : (note.choice || note.text));
  return values;
}

function parseNote_(value) {
  var text = String(value == null ? '' : value).trim();
  for (var i = 0; i < NOTE_CHOICES.length; i++) {
    var c = NOTE_CHOICES[i];
    if (text === c) return { choice: c, text: '' };
    if (text.indexOf(c + NOTE_SEPARATOR) === 0) return { choice: c, text: text.slice(c.length + NOTE_SEPARATOR.length) };
  }
  return { choice: '', text: text };
}

function toSheetNumber_(t) {
  var p = String(t).split(':');
  return Number(p[0]) * 100 + Number(p[1]);
}

/* ================= 送信内容のチェック ================= */

function parseSubmission_(payload, config, states) {
  var dirty = Array.isArray(payload.dirtyWeeks) ? payload.dirtyWeeks : [];
  var dirtyWeeks = [];
  dirty.forEach(function(w) {
    if (typeof w !== 'number' || w !== Math.floor(w) || w < 0 || w >= config.weekCount) throw new Error('週の内容が正しくありません。');
    if (dirtyWeeks.indexOf(w) === -1) dirtyWeeks.push(w);
  });
  dirtyWeeks.sort();

  var shifts = states.map(function() { return {}; });
  var days = payload.days || {};
  Object.keys(days).forEach(function(key) {
    var m = /^(\d)-(\d)$/.exec(key);
    if (!m || Number(m[1]) >= config.weekCount || Number(m[2]) > 6) throw new Error('日付の内容が正しくありません。');
    var w = Number(m[1]), d = Number(m[2]);
    if (dirtyWeeks.indexOf(w) === -1) return;
    var v = days[key] || {};
    var current = states[w].shifts[d];
    // 新しく選んだ時間は決まりの範囲内だけ。シートに元から入っていた時間はそのまま送り返してよい
    var unchanged = current && current.s === v.s && current.e === v.e;
    if (!unchanged && !validShift_(v.s, v.e, config.timeRange)) throw new Error('時間の内容が正しくありません。');
    shifts[w][d] = { s: String(v.s), e: String(v.e) };
  });

  var notes = states.map(function(s) { return s.note; });
  dirtyWeeks.forEach(function(w) {
    var n = (payload.notes && payload.notes[w]) || {};
    var choice = String(n.choice || '');
    if (choice && NOTE_CHOICES.indexOf(choice) === -1) throw new Error('備考の選択が正しくありません。');
    notes[w] = { choice: choice, text: cleanText_(n.text, 200, true).replace(/\n/g, ' ') };
  });
  return { dirtyWeeks: dirtyWeeks, shifts: shifts, notes: notes };
}

function sumShifts_(weekShifts) {
  var days = 0, minutes = 0;
  weekShifts.forEach(function(shifts) {
    Object.keys(shifts).forEach(function(d) {
      var s = anyMinutes_(shifts[d].s), e = anyMinutes_(shifts[d].e);
      if (s === null || e === null || e <= s) return;
      days++;
      minutes += e - s;
    });
  });
  return { days: days, hours: Math.round(minutes / 60 * 100) / 100 };
}

// 断った送信も「送信できなかった」として理由つきで残す（記録に失敗しても元のエラーを優先）
function recordFailure_(rec, reason) {
  try {
    withLock_(function() {
      if (overLimit_('ev:' + (rec.device || 'none'), LIMITS.event)) return;
      writeRecord_(getStore_(), {
        action: ACTION_FAILED, name: rec.name, detail: reason,
        device: rec.device, ip: rec.ip, model: rec.model, ua: rec.ua
      });
    });
  } catch (e) {
    console.error('送信できなかった記録の書き込みに失敗: ' + e.message);
  }
}

/* ================= 記録の書き込み ================= */

// 操作の記録に1行足し、端末の持ち主の一覧も更新する。注意の文を返す
function writeRecord_(ss, rec, now) {
  now = now || new Date();
  if (!isTestBrowser_(rec.ua)) cleanupTestRecords_(ss, now.getTime() - TEST_RECORD_KEEP_MINUTES * 60000);
  var logSheet = sheet_(ss, SHEETS.log.name);
  var deviceSheet = sheet_(ss, SHEETS.devices.name);
  var usage = readUsage_(logSheet);
  var devices = readDevices_(deviceSheet);

  var warning = (rec.action === ACTION_SUBMIT || rec.action === EVENT_ACTIONS.select || rec.action === ACTION_FAILED)
    ? buildWarning_(rec.name, rec.device, usage, devices) : '';
  if (rec.extraWarning) warning = warning ? warning + ' ／ ' + rec.extraWarning : rec.extraWarning;

  logSheet.appendRow(safeRow_([now, '', rec.action, rec.name, rec.detail || '', warning, rec.device, rec.ip, rec.model, rec.ua]));
  setOwnerFormula_(logSheet, LOG_DEVICE_COL);

  if (rec.device) {
    if (rec.action === ACTION_SUBMIT || rec.action === EVENT_ACTIONS.select) addUsage_(usage, rec.name, rec.device);
    updateDevice_(deviceSheet, devices[rec.device], rec, usage, now);
  }
  return warning;
}

// 「端末の持ち主」列は、持ち主の一覧を見に行く式にする（あとから持ち主を決めても過去の行に反映される）
function setOwnerFormula_(sheet, deviceCol) {
  sheet.getRange(sheet.getLastRow(), OWNER_COL)
    .setFormulaR1C1('=IFERROR(""&VLOOKUP(R[0]C' + deviceCol + ",'" + SHEETS.devices.name + "'!C1:C2,2,FALSE),\"\")");
}

// これまでの「名前を選択」「送信」から、端末ごとに選ばれた名前の回数と、名前ごとの端末を集める
function readUsage_(logSheet) {
  var usage = { namesByDevice: {}, devicesByName: {} };
  var last = logSheet.getLastRow();
  if (last < 2) return usage;
  logSheet.getRange(2, 3, last - 1, SHEETS.log.headers.length - 2).getValues().forEach(function(r) {
    var action = String(r[0]);
    if (action !== ACTION_SUBMIT && action !== EVENT_ACTIONS.select) return;
    if (isTestBrowser_(r[r.length - 1])) return; // 自動確認の記録は注意の判定に使わない
    addUsage_(usage, String(r[1]), String(r[LOG_DEVICE_COL - 3]));
  });
  return usage;
}

function addUsage_(usage, name, device) {
  if (!name || !device) return;
  var counts = usage.namesByDevice[device] || (usage.namesByDevice[device] = {});
  counts[name] = (counts[name] || 0) + 1;
  (usage.devicesByName[name] || (usage.devicesByName[name] = {}))[device] = true;
}

function readDevices_(deviceSheet) {
  var devices = {};
  var last = deviceSheet.getLastRow();
  if (last < 2) return devices;
  deviceSheet.getRange(2, 1, last - 1, SHEETS.devices.headers.length).getValues().forEach(function(r, i) {
    if (!r[0]) return;
    devices[String(r[0])] = { row: i + 2, owner: String(r[1]), uses: Number(r[7]) || 0 };
  });
  return devices;
}

function updateDevice_(deviceSheet, entry, rec, usage, now) {
  var counts = usage.namesByDevice[rec.device] || {};
  var names = Object.keys(counts).sort(function(a, b) { return counts[b] - counts[a]; });
  var candidate = names.length ? names[0] : '';
  var summary = names.map(function(n) { return n + '×' + counts[n]; }).join('、');

  if (entry) {
    // 「最初に使った日時」(F列)は書き直さない
    deviceSheet.getRange(entry.row, 3, 1, 3).setValues([[candidate, summary, rec.model]]);
    deviceSheet.getRange(entry.row, 7, 1, 2).setValues([[now, entry.uses + 1]]);
    return;
  }
  deviceSheet.appendRow(safeRow_([rec.device, '', candidate, summary, rec.model, now, now, 1]));
  deviceSheet.getRange(deviceSheet.getLastRow(), 2).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(currentStaff_(), true).setAllowInvalid(false).build());
}

// 持ち主のプルダウン用の名簿（シフト表の最新の週シートから）
function currentStaff_() {
  var config = loadConfig_(SpreadsheetApp.openById(SHIFT_SPREADSHEET_ID));
  return config.staff && config.staff.length ? config.staff : ['（名簿を読めませんでした）'];
}

// 持ち主が決まっている端末は持ち主と比べる。決まっていなければ使われ方から注意を出す
function buildWarning_(name, device, usage, devices) {
  if (!name) return '';
  var entry = devices[device];
  if (entry && entry.owner) return entry.owner === name ? '' : '持ち主（' + entry.owner + '）と違う名前を選択';

  var messages = [];
  var others = Object.keys(usage.namesByDevice[device] || {}).filter(function(n) { return n !== name; });
  if (others.length) messages.push('この端末で他の名前も使用：' + others.join('、'));
  var hasOwnDevice = Object.keys(devices).some(function(d) { return devices[d].owner === name; });
  if (hasOwnDevice) {
    messages.push('本人の端末として登録されていない端末');
  } else if (Object.keys(usage.devicesByName[name] || {}).some(function(d) { return d !== device; })) {
    messages.push('この名前は別の端末でも使用');
  }
  return messages.join(' ／ ');
}

// 週ごとに「10/26(月) 17:00-21:30」の形で読みやすくまとめる（5週分の列）。変えていない週は「変更なし」
function weekSummaries_(config, weekShifts, weekNotes, outcome) {
  var out = [];
  for (var w = 0; w < 5; w++) {
    if (w >= config.weekCount) { out.push(''); continue; }
    if (!outcome[w]) { out.push('変更なし'); continue; }
    var lines = [];
    for (var d = 0; d < 7; d++) {
      var a = weekShifts[w][d];
      if (!a) continue;
      var p = addDays_(config.startDate, w * 7 + d).split('-').map(Number);
      lines.push(p[1] + '/' + p[2] + '(' + DOW_LABELS[d] + ') ' + a.s + '-' + a.e);
    }
    if (!lines.length) lines.push('出勤希望なし');
    if (weekNotes[w].choice) lines.push('選択：' + weekNotes[w].choice);
    if (weekNotes[w].text) lines.push('備考：' + weekNotes[w].text);
    lines.push(outcome[w]);
    out.push(lines.join('\n'));
  }
  return out;
}

/* ================= 記録用スプレッドシート ================= */

function getStore_() {
  if (!STORE_SPREADSHEET_ID) throw new Error('記録先のスプレッドシートが設定されていません。');
  var ss = SpreadsheetApp.openById(STORE_SPREADSHEET_ID);
  // 会社アカウントで作ったファイルは時間帯が日本でないことがあり、記録の時刻がずれるため揃える
  if (ss.getSpreadsheetTimeZone() !== 'Asia/Tokyo') ss.setSpreadsheetTimeZone('Asia/Tokyo');
  ensureSheets_(ss);
  return ss;
}

// 必要なシートが無ければ作る（新しい空のファイルの「シート1」は名前を変えて使う）
function ensureSheets_(ss) {
  Object.keys(SHEETS).forEach(function(key) {
    var def = SHEETS[key];
    if (ss.getSheetByName(def.name)) return;
    var blank = ss.getSheets().filter(function(s) {
      return s.getLastRow() === 0 && !Object.keys(SHEETS).some(function(k) { return SHEETS[k].name === s.getName(); });
    })[0];
    var sheet = blank ? blank.setName(def.name) : ss.insertSheet(def.name);
    sheet.getRange(1, 1, 1, def.headers.length).setValues([def.headers]).setFontWeight('bold').setBackground('#EFF0ED');
    sheet.setFrozenRows(1);
    if (key === 'submit') sheet.getRange('G:K').setWrap(true);
    if (key === 'devices') sheet.getRange('B1').setBackground('#FFE9A8');
  });
}

function sheet_(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error('シート「' + name + '」が見つかりません。');
  return sheet;
}

/* ================= 共通 ================= */

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) throw new Error('混み合っているため受け付けられませんでした。少し待ってからもう一度送ってください。');
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function overLimit_(key, limit) {
  var cache = CacheService.getScriptCache();
  var count = Number(cache.get(key) || 0);
  if (count >= limit.max) return true;
  cache.put(key, String(count + 1), limit.seconds);
  return false;
}

// 'yyyy-MM-dd' に日数を足す（時間帯の影響を受けないように日付だけで計算する）
function addDays_(ymd, days) {
  var p = ymd.split('-').map(Number);
  var d = new Date(Date.UTC(p[0], p[1] - 1, p[2] + days));
  return d.getUTCFullYear() + '-' + pad2_(d.getUTCMonth() + 1) + '-' + pad2_(d.getUTCDate());
}

// '2026-10-26' → '261026'（週シートの名前）
function sheetNameOf_(ymd) {
  return ymd.slice(2).replace(/-/g, '');
}

// '2026-10-21 18:00' → '10/21(水) 18:00'
function deadlineLabel_(text) {
  var p = text.slice(0, 10).split('-').map(Number);
  var dow = new Date(Date.UTC(p[0], p[1] - 1, p[2])).getUTCDay();
  return p[1] + '/' + p[2] + '(' + ['日', '月', '火', '水', '木', '金', '土'][dow] + ') ' + text.slice(11);
}

function pad2_(n) {
  return (n < 10 ? '0' : '') + n;
}

// スタッフが送ってくる時間。15分きざみまで読む（幅の端が21:15のような店があるため）
function toMinutes_(t) {
  var m = /^(\d{2}):(00|15|30|45)$/.exec(String(t));
  return m ? Number(m[1]) * 60 + Number(m[2]) : null;
}

function anyMinutes_(t) {
  var m = /^(\d{1,2}):(\d{2})$/.exec(String(t));
  return m ? Number(m[1]) * 60 + Number(m[2]) : null;
}

// スタッフが選べる時間の幅を読む（30分きざみ。空や書き間違いは初期値）
function timeRange_(settings) {
  function read(cell, fallback) {
    var m = /^(\d{1,2}):(00|15|30|45)$/.exec(String(settings.getRange(cell).getDisplayValue()).trim());
    var v = m && Number(m[1]) <= 23 ? Number(m[1]) * 60 + Number(m[2]) : null;
    return v === null ? fallback : v;
  }
  var d = TIME_RANGE_DEFAULT;
  var range = {
    start: { from: read(TIME_RANGE_CELLS.startFrom, d.start.from), to: read(TIME_RANGE_CELLS.startTo, d.start.to) },
    end: { from: read(TIME_RANGE_CELLS.endFrom, d.end.from), to: read(TIME_RANGE_CELLS.endTo, d.end.to) }
  };
  // 早い→遅いが逆になっていたら初期値に戻す（画面が空にならないように）
  if (range.start.to <= range.start.from) range.start = { from: d.start.from, to: d.start.to };
  if (range.end.to <= range.end.from) range.end = { from: d.end.from, to: d.end.to };
  return range;
}

// スタッフが選べるのは30分きざみ。ただし幅の端（いちばん早い・いちばん遅い）だけは
// 21:15 のような15分きざみの時間も選べる（閉店時間に合わせるため）
function inTimeRange_(v, part) {
  if (v < part.from || v > part.to) return false;
  return v % 30 === 0 || v === part.from || v === part.to;
}

function validShift_(s, e, range) {
  var r = range || TIME_RANGE_DEFAULT;
  var sm = toMinutes_(s), em = toMinutes_(e);
  if (sm === null || em === null) return false;
  if (!inTimeRange_(sm, r.start) || !inTimeRange_(em, r.end)) return false;
  return em > sm;
}

function validDeviceId_(v) {
  var s = String(v || '');
  return /^[A-Za-z0-9-]{8,64}$/.test(s) ? s : '';
}

function cleanIp_(v) {
  var s = String(v || '');
  return /^[0-9A-Fa-f:.]{3,45}$/.test(s) ? s : '';
}

function cleanText_(v, maxLength, allowNewline) {
  var s = String(v == null ? '' : v);
  s = allowNewline ? s.replace(/[ -	-]/g, ' ') : s.replace(/[ -]/g, ' ');
  return s.trim().slice(0, maxLength);
}

// 「=」などで始まる文字が数式として動かないようにする
function safeRow_(row) {
  return row.map(function(v) {
    return (typeof v === 'string' && /^[=+\-@]/.test(v)) ? "'" + v : v;
  });
}

/* ================= カレンダー配信（自分のシフトをカレンダーアプリに出す） ================= */

// 鍵を作るための秘密の文字。変えると、全員のカレンダー登録URLが使えなくなる
var CALENDAR_SECRET = '739be42f457eb4bdce51a7d9292861c2bcc453fb3306b677';
var CALENDAR_KEY_LENGTH = 16;
var CALENDAR_PAST_DAYS = 28;          // 何日前の週まで載せるか（過去の出勤も少し残す）
// 週シートの日ブロック（シフト表GASの ADJUST_BLOCKS と同じ）。head＝見出しの行、col＝時間の升目が始まる列
var CALENDAR_BLOCKS = [
  { head: 23, col: 25 }, { head: 42, col: 25 }, { head: 61, col: 25 },
  { head: 3, col: 113 }, { head: 23, col: 113 }, { head: 42, col: 113 }, { head: 61, col: 113 }
];
var CALENDAR_PEOPLE = 16;             // 1日に並ぶ人数
// 1枚の週シートから読む範囲。時間帯の升目（数式が入っていて重い）は読まない。
// 左は A列（名簿）〜X列（左の日ブロックの休終）をまとめて1回、右は5列ぶんを1回
var CALENDAR_READ_ROWS = 80;
var CALENDAR_LEFT_COLS = 24;          // A〜X
var CALENDAR_RIGHT_COL = 108;         // 右の日ブロックの名前の列（113-5）
var CALENDAR_RIGHT_COLS = 5;
// このウェブアプリ自身のアドレス（カレンダー登録やお知らせのリンクに使う。作り直さない決まり）
var SELF_URL = 'https://script.google.com/macros/s/AKfycbwFtgsWnq75UnttAODig9q3c89E1Iek0Oow-lV8tHzF_Zp5qBQe2DXPahAeksqkd4i2/exec';
var PUBLISHED_WEEKS_CELL = 'E26';     // シフト希望設定：確定して公開した週（シフト表GASが書く）

// 名前から、その人だけのURLの鍵を作る（同じ名前なら毎回同じ鍵）
function calendarKey_(name) {
  var bytes = Utilities.computeHmacSha256Signature(normalizeName_(name), CALENDAR_SECRET, Utilities.Charset.UTF_8);
  return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/, '').slice(0, CALENDAR_KEY_LENGTH);
}

function normalizeName_(name) {
  return String(name == null ? '' : name).replace(/[\s　]+/g, '');
}

function calendarFeed_(key) {
  var shiftSs = SpreadsheetApp.openById(SHIFT_SPREADSHEET_ID);
  var weeks = readPublishedWeeks_(shiftSs, new Date());
  var who = '';
  namesInWeeks_(weeks).some(function(n) {
    if (calendarKey_(n) === key) { who = n; return true; }
    return false;
  });
  var storeName = STORE_NAME;
  try { storeName = loadConfig_(shiftSs).storeName || STORE_NAME; } catch (err) {}
  var ics = buildIcs_(who ? personShifts_(weeks, who) : [], storeName, who);
  return ContentService.createTextOutput(ics).setMimeType(ContentService.MimeType.ICAL);
}

// 公開した週に出てくる名前（表に書かれたままの形）
function namesInWeeks_(weeks) {
  var seen = {}, names = [];
  weeks.forEach(function(wk) {
    wk.days.forEach(function(people) {
      people.forEach(function(p) {
        var key = normalizeName_(p.name);
        if (seen[key]) return;
        seen[key] = true;
        names.push(p.name);
      });
    });
  });
  return names;
}

// 載せる週：今日から CALENDAR_PAST_DAYS 日前以降で、店長が「確定する」を押して公開した週。
// まだ一度も公開していない間は、希望を受け付けている週より前を載せる（今までどおり）
// fromYmd（'20260921' の形）を渡すと、その日より前に終わる週は読まない。
// 読んでから捨てるのは無駄（1枚読むのに1秒ほどかかる）。カレンダー配信は渡さないので今までどおり
function calendarWeekSheets_(shiftSs, today, fromYmd) {
  var from = new Date(today.getTime() - (CALENDAR_PAST_DAYS + 7) * 86400000);
  var published = publishedWeeks_(shiftSs);
  return shiftSs.getSheets().filter(function(sh) {
    var n = sh.getName();
    if (!/^\d{6}$/.test(n)) return false;
    if (published.length ? published.indexOf(n) === -1 : n >= FIRST_WRITABLE_SHEET) return false;
    if (weekStartOf_(n) < from) return false;
    if (fromYmd && weekEndYmdOf_(n) < String(fromYmd)) return false;
    return true;
  });
}

// 店長が確定して公開した週（シフト希望設定の E26。シフト表GASが書きます）
function publishedWeeks_(shiftSs) {
  try {
    var settings = shiftSs.getSheetByName(REQUEST_SETTINGS_SHEET);
    if (!settings) return [];
    return String(settings.getRange(PUBLISHED_WEEKS_CELL).getDisplayValue())
      .split('|').map(function(x) { return x.trim(); }).filter(function(x) { return /^\d{6}$/.test(x); });
  } catch (err) {
    return [];
  }
}

function weekStartOf_(sheetName) {
  return new Date(2000 + Number(sheetName.slice(0, 2)), Number(sheetName.slice(2, 4)) - 1, Number(sheetName.slice(4, 6)));
}

// その人のシフトだけを取り出す（カレンダー用）
function personShifts_(weeks, name) {
  var key = normalizeName_(name);
  var out = [];
  weeks.forEach(function(wk) {
    wk.days.forEach(function(people, day) {
      people.forEach(function(p) {
        if (normalizeName_(p.name) !== key) return;
        out.push({ ymd: wk.ymds[day], start: p.start, end: p.end, breakStart: p.breakStart, breakEnd: p.breakEnd });
      });
    });
  });
  return out;
}

// 公開した週を、週ごと・曜日ごとに読む
// [{ sheet, ymds:[7], labels:[7], days:[[{name,start,end,breakStart,breakEnd}]×7] }]
function readPublishedWeeks_(shiftSs, today, fromYmd) {
  var weeks = [];
  calendarWeekSheets_(shiftSs, today, fromYmd).forEach(function(sh) {
    var monday = weekStartOf_(sh.getName());
    var week = { sheet: sh.getName(), ymds: [], labels: [], days: [[], [], [], [], [], [], []] };
    // 速さの肝は「読みに行く回数」。1回につき0.2秒ほどかかるので、
    // 1枚につき2回（左と右の値）だけにする。名簿も左の中に入っている。
    // 「見た目」は時間がDateや文字で入っていたときだけ読む（ふつうは900のような数字なので読まない）
    var leftRange = sh.getRange(1, 1, CALENDAR_READ_ROWS, CALENDAR_LEFT_COLS);
    var rightRange = sh.getRange(1, CALENDAR_RIGHT_COL, CALENDAR_READ_ROWS, CALENDAR_RIGHT_COLS);
    var v1 = leftRange.getValues(), v2 = rightRange.getValues(), d1 = null, d2 = null;
    function isLeft(col) { return col <= CALENDAR_LEFT_COLS; }
    function val(row, col) {
      var r = (isLeft(col) ? v1 : v2)[row - 1] || [];
      return r[isLeft(col) ? col - 1 : col - CALENDAR_RIGHT_COL];
    }
    function disp(row, col) {
      if (isLeft(col)) { if (!d1) d1 = leftRange.getDisplayValues(); }
      else if (!d2) d2 = rightRange.getDisplayValues();
      var r = (isLeft(col) ? d1 : d2)[row - 1] || [];
      return r[isLeft(col) ? col - 1 : col - CALENDAR_RIGHT_COL];
    }
    function timeAt(row, col) {
      var v = val(row, col);
      if (v === '' || v === null || v === undefined) return '';
      if (typeof v === 'number') return sheetTime_(v, '');
      return sheetTime_(v, disp(row, col));           // ここだけ見た目が要る
    }

    var order = {};                                   // 名簿（週シートの名前の欄）の並び順
    rosterFromGrid_(v1).forEach(function(n, i) { order[normalizeName_(n)] = i; });
    CALENDAR_BLOCKS.forEach(function(b, day) {
      var nameCol = b.col - 5;
      var date = new Date(monday.getFullYear(), monday.getMonth(), monday.getDate() + day);
      week.ymds[day] = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyyMMdd');
      week.labels[day] = (date.getMonth() + 1) + '/' + date.getDate();
      for (var r = b.head + 1; r <= b.head + CALENDAR_PEOPLE; r++) {
        var name = String(val(r, nameCol) || '').trim();
        var start = timeAt(r, nameCol + 1), end = timeAt(r, nameCol + 2);
        if (!name || !start || !end) continue;
        week.days[day].push({
          name: name, start: start, end: end,
          breakStart: timeAt(r, nameCol + 3), breakEnd: timeAt(r, nameCol + 4)
        });
      }
      // 並びは週シートの名前の欄と同じ順（名簿順）。名簿に無い人は後ろに回し、同じ位置なら早い出勤から
      week.days[day].sort(function(a, b2) {
        var ia = order.hasOwnProperty(normalizeName_(a.name)) ? order[normalizeName_(a.name)] : 9999;
        var ib = order.hasOwnProperty(normalizeName_(b2.name)) ? order[normalizeName_(b2.name)] : 9999;
        return ia - ib || toSheetNumber_(a.start) - toSheetNumber_(b2.start);
      });
    });
    weeks.push(week);
  });
  weeks.sort(function(a, b) { return a.sheet < b.sheet ? -1 : 1; });
  return weeks;
}

function buildIcs_(shifts, storeName, name) {
  var stamp = Utilities.formatDate(new Date(), 'UTC', "yyyyMMdd'T'HHmmss'Z'");
  var lines = [
    'BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//' + icsText_(storeName) + '//shift//JA',
    'CALSCALE:GREGORIAN', 'METHOD:PUBLISH',
    'X-WR-CALNAME:' + icsText_('シフト（' + storeName + '）'),
    'X-WR-TIMEZONE:Asia/Tokyo',
    'REFRESH-INTERVAL;VALUE=DURATION:PT1H', 'X-PUBLISHED-TTL:PT1H',
    'BEGIN:VTIMEZONE', 'TZID:Asia/Tokyo', 'BEGIN:STANDARD', 'DTSTART:19700101T000000',
    'TZOFFSETFROM:+0900', 'TZOFFSETTO:+0900', 'TZNAME:JST', 'END:STANDARD', 'END:VTIMEZONE'
  ];
  shifts.forEach(function(x) {
    var brk = x.breakStart && x.breakEnd ? '休憩' + trimZero_(x.breakStart) + '〜' + trimZero_(x.breakEnd) : '休憩なし';
    lines.push(
      'BEGIN:VEVENT',
      // 日付と名前で決まるID。シフトが変わっても同じ予定として書き換わる
      'UID:' + x.ymd + '-' + calendarKey_(name) + '@shift',
      'DTSTAMP:' + stamp,
      'DTSTART;TZID=Asia/Tokyo:' + x.ymd + 'T' + x.start.replace(':', '') + '00',
      'DTEND;TZID=Asia/Tokyo:' + x.ymd + 'T' + x.end.replace(':', '') + '00',
      // スマホの月表示は名前の先頭6文字ほどしか見えないので、短い「出勤-退勤」を先頭に置く（9-18 休16-17）
      'SUMMARY:' + icsText_(shortTime_(x.start) + '-' + shortTime_(x.end) + ' ' + (x.breakStart && x.breakEnd ? '休' + shortTime_(x.breakStart) + '-' + shortTime_(x.breakEnd) : '休なし')),
      'LOCATION:' + icsText_(storeName),
      'DESCRIPTION:' + icsText_('出勤 ' + trimZero_(x.start) + '〜' + trimZero_(x.end) + '\n' + brk),
      'END:VEVENT');
  });
  lines.push('END:VCALENDAR');
  return lines.map(foldIcsLine_).join('\r\n') + '\r\n';
}

// 09:00→9、21:30→21:30
function shortTime_(t) {
  return trimZero_(t).replace(/:00$/, '');
}

function trimZero_(t) {
  return String(t).replace(/^0(\d):/, '$1:');
}

function icsText_(text) {
  return String(text).replace(/\\/g, '\\\\').replace(/;/g, '\\;').replace(/,/g, '\\,').replace(/\r?\n/g, '\\n');
}

// 1行75バイトまでの決まりに合わせて折り返す（日本語は1文字3バイト）
function foldIcsLine_(line) {
  var out = [], cur = '', bytes = 0;
  for (var i = 0; i < line.length; i++) {
    var ch = line.charAt(i);
    var b = ch.charCodeAt(0) < 0x80 ? 1 : 3;
    if (bytes + b > (out.length ? 74 : 75)) { out.push(cur); cur = ''; bytes = 0; }
    cur += ch; bytes += b;
  }
  out.push(cur);
  return out.join('\r\n ');
}

/* ================= できたシフトを見るページ ================= */

/* ================= 自分の給与（概算） =================
   スタッフのページで、締め期間ごとの勤務時間と概算の給与を見せる。
   ・他の人のぶんが見えないように、生年月日の月日4桁（0131など）が合った人にだけ返す
   ・社員（スタッフ区分シートに載っている人）には出さない
   ・元にするのは「出勤履歴」シート（店長が転写した日のぶんだけ入っている）
   ・割増（深夜・残業）は考えない。実際の支給額とは違うことがある */
var ROLE_BIRTH_COL = 3;             // スタッフ区分：C列＝生年月日
var ROLE_WAGE_COL = 5;              // E列＝今の時給
var ROLE_PREV_WAGE_COL = 6;         // F列＝前の時給
var ROLE_WAGE_FROM_COL = 7;         // G列＝今の時給はいつから
var HISTORY_DATE_COL = 1;           // 出勤履歴：A列＝日付
var HISTORY_NAME_COL = 3;           // C列＝名前
var HISTORY_HOURS_COL = 8;          // H列＝実働時間
var HISTORY_SOURCE_COL = 9;         // I列＝元シート（どの週から来たか）
var HISTORY_HEADERS_LEN = 9;        // 出勤履歴の列数
var FARE_KIND_CELL = 'E33';         // 交通費の出し方
var FARE_DAYS_CELL = 'E34';
var FARE_AMOUNT_CELL = 'E35';

// 店長・社員の名前（この人たちには給与を出さない）
function employeeNames_(shiftSs) {
  var roles = shiftSs.getSheetByName(ROLE_SHEET_NAME);
  if (!roles || roles.getLastRow() < 2) return [];
  return roles.getRange(2, 1, roles.getLastRow() - 1, 2).getValues()
    .filter(function(r) { return EMPLOYEE_ROLES.indexOf(String(r[1]).trim()) !== -1; })
    .map(function(r) { return String(r[0]).trim(); })
    .filter(function(n) { return n; });
}

// 日付のマスをそのまま「2026-09-21」にする（時差でずれないよう、時間帯の変換はしない）
function dateToYmd_(date) {
  return date.getFullYear() + '-' + ('0' + (date.getMonth() + 1)).slice(-2) + '-' + ('0' + date.getDate()).slice(-2);
}

// 「9/21(月)」のような見せ方
function ymdLabel_(ymd) {
  var p = ymd.split('-').map(Number);
  var dow = ['日', '月', '火', '水', '木', '金', '土'][new Date(p[0], p[1] - 1, p[2]).getDay()];
  return p[1] + '/' + p[2] + '(' + dow + ')';
}

// 名前と生年月日から、その人の時給・生年月日を読む
function staffPayInfo_(shiftSs, name) {
  var sheet = shiftSs.getSheetByName(ROLE_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return null;
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, ROLE_WAGE_FROM_COL).getValues();
  var key = normalizeName_(name);
  for (var i = 0; i < rows.length; i++) {
    if (normalizeName_(rows[i][0]) !== key) continue;
    return {
      birth: (rows[i][ROLE_BIRTH_COL - 1] instanceof Date) ? rows[i][ROLE_BIRTH_COL - 1] : null,
      wage: Number(rows[i][ROLE_WAGE_COL - 1]) || 0,
      prevWage: Number(rows[i][ROLE_PREV_WAGE_COL - 1]) || 0,
      wageFrom: (rows[i][ROLE_WAGE_FROM_COL - 1] instanceof Date) ? rows[i][ROLE_WAGE_FROM_COL - 1] : null
    };
  }
  return null;
}

// その締め期間に使う時給（期間の初日が「いつから」より前なら、前の時給）
function wageForPeriod_(info, startYmd, tz) {
  if (!info) return 0;
  if (info.wageFrom && info.prevWage > 0 && startYmd < dateToYmd_(info.wageFrom)) {
    return info.prevWage;
  }
  return info.wage;
}

// 交通費の決まり
function readFareRule_(settings) {
  return {
    kind: String(settings.getRange(FARE_KIND_CELL).getDisplayValue()).trim(),
    days: Number(settings.getRange(FARE_DAYS_CELL).getDisplayValue()) || 0,
    amount: Number(settings.getRange(FARE_AMOUNT_CELL).getDisplayValue()) || 0
  };
}

function fareFor_(rule, workDays) {
  if (!rule || !workDays) return 0;
  if (rule.kind === '日数以上で決まった額') return workDays >= rule.days ? rule.amount : 0;
  if (rule.kind === '出勤1日ごと') return workDays * rule.amount;
  return 0;
}

// 今日を含む締め期間と、その1つ前（今月分・先月分）
function payPeriods_(settings, todayYmd) {
  var day = closingDayOf_(settings, false);          // キャストの締め日
  var ranges = closingRanges_(addDays_(todayYmd, -80), todayYmd, day);
  return ranges.slice(-2);
}

// 出勤履歴から、その人の日ごとの実働時間を締め期間ごとに足す
function payTotals_(shiftSs, name, periods) {
  var out = periods.map(function(p) { return { start: p.start, end: p.end, days: 0, hours: 0 }; });
  var history = shiftSs.getSheetByName(HISTORY_SHEET_NAME);
  if (!history || history.getLastRow() < 2) return out;
  var tz = shiftSs.getSpreadsheetTimeZone();
  var key = normalizeName_(name);
  var counted = {};
  history.getRange(2, 1, history.getLastRow() - 1, HISTORY_HOURS_COL).getValues().forEach(function(r) {
    if (!(r[HISTORY_DATE_COL - 1] instanceof Date)) return;
    if (normalizeName_(r[HISTORY_NAME_COL - 1]) !== key) return;
    var hours = Number(r[HISTORY_HOURS_COL - 1]);
    if (!(hours > 0)) return;
    var ymd = dateToYmd_(r[HISTORY_DATE_COL - 1]);
    for (var i = 0; i < out.length; i++) {
      if (ymd >= out[i].start && ymd <= out[i].end) {
        if (!counted[ymd]) { counted[ymd] = true; out[i].days++; }   // 同じ日が2行あっても1日
        out[i].hours += hours;                                       // 時間のほうは足す
        return;
      }
    }
  });
  return out;
}

// この端末の持ち主（「シフト希望_アクセス記録」の端末の持ち主シートで、店長が選んだ人）。
// 分からなければ空。給与は「この端末の持ち主 ＝ 選んだ名前」のときだけ見せる
function deviceOwner_(deviceId) {
  var id = String(deviceId == null ? '' : deviceId).trim();
  if (!id) return '';
  try {
    var sheet = getStore_().getSheetByName(SHEETS.devices.name);
    if (!sheet || sheet.getLastRow() < 2) return '';
    var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, OWNER_COL).getValues();
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === id) return String(rows[i][OWNER_COL - 1]).trim();
    }
  } catch (err) {}
  return '';
}

// 画面から呼ぶ：この端末の持ち主で、生年月日も合っていれば、その人の今月分・先月分を返す
function myPayInfo(name, birthCode, deviceId) {
  var shiftSs = SpreadsheetApp.openById(SHIFT_SPREADSHEET_ID);
  var settings = shiftSs.getSheetByName(REQUEST_SETTINGS_SHEET);
  if (!settings) return { ok: false, reason: 'ほかの理由', message: '今は見られません。店長に伝えてください。' };
  if (isEmployee_(shiftSs, name)) {
    return { ok: false, reason: '社員', message: 'この画面では、アルバイトの方の給与だけを出しています。' };
  }
  // ①この端末の持ち主かどうか（他の人のスマホからは見られない）
  var owner = deviceOwner_(deviceId);
  if (!owner) {
    return { ok: false, reason: '端末なし', message:
      'このスマホはまだ登録されていません。\nシフト希望のフォームを一度開いてから、店長に伝えてください。' };
  }
  if (normalizeName_(owner) !== normalizeName_(name)) {
    return { ok: false, reason: '端末ちがい', message:
      'このスマホは、ほかの方のものとして登録されています。\n自分のスマホで開いてください。' };
  }
  var info = staffPayInfo_(shiftSs, name);
  if (!info || !info.birth) {
    return { ok: false, reason: '生年月日なし', message: '生年月日がまだ登録されていません。店長に伝えてください。' };
  }
  var tz = shiftSs.getSpreadsheetTimeZone();
  var right = dateToYmd_(info.birth).slice(5).replace('-', '');
  if (String(birthCode).replace(/[^0-9]/g, '') !== right) {
    return { ok: false, reason: '合わない', message: '生年月日が合いません。誕生日の月日を4けたで入れてください（1月31日なら 0131）。' };
  }
  if (!info.wage) {
    return { ok: false, reason: '時給なし', message: '時給がまだ登録されていません。店長に伝えてください。' };
  }

  var todayYmd = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var periods = payPeriods_(settings, todayYmd);
  var rule = readFareRule_(settings);
  var totals = payTotals_(shiftSs, name, periods);
  return {
    ok: true,
    periods: totals.map(function(t, i) {
      var wage = wageForPeriod_(info, t.start, tz);
      var pay = Math.round(wage * t.hours);
      var fare = fareFor_(rule, t.days);
      return {
        label: Number(t.end.split('-')[1]) + '月' + Number(t.end.split('-')[2]) + '日〆',
        range: ymdLabel_(t.start) + '〜' + ymdLabel_(t.end),
        startYmd: t.start,
        endYmd: t.end,
        days: t.days,
        hours: Math.round(t.hours * 10) / 10,
        wage: wage,
        pay: pay,
        fare: fare,
        total: pay + fare,
        ongoing: todayYmd <= t.end
      };
    }).reverse()
  };
}

// ページに出す週は「今週から先」だけにする。終わった週はLINEに残っているPDFで見てもらう
// （カレンダーのほうは実績として少し前まで残す＝CALENDAR_PAST_DAYS）
// 画面に出す週。今週からが基本。ただし「締め日までの合計」を出すため、
// いまの締め期間のはじめ（fromYmd）がもっと前なら、そこまでさかのぼって出す。
// ※ ymds は '20261012' の形。前はここを '2026-10-12' の形で比べていたので、どの週も残っていた
// 画面に出す「いちばん前の日」。今週の月曜が基本。締め期間のはじめがもっと前なら、そこまで
function weekFloorYmd_(today, fromYmd) {
  var d = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  d.setDate(d.getDate() - ((d.getDay() + 6) % 7));        // 今週の月曜
  var from = d.getFullYear() + ('0' + (d.getMonth() + 1)).slice(-2) + ('0' + d.getDate()).slice(-2);
  if (fromYmd && String(fromYmd) < from) from = String(fromYmd);
  return from;
}

function weeksFromThisWeek_(weeks, today, fromYmd) {
  var from = weekFloorYmd_(today, fromYmd);
  return weeks.filter(function(wk) { return wk.ymds[6] >= from; });   // 週の最終日がその日以降なら出す
}

// 週シート名（261019）の、その週の日曜日を '20261025' の形で
function weekEndYmdOf_(sheetName) {
  var monday = weekStartOf_(sheetName);
  var sunday = new Date(monday.getFullYear(), monday.getMonth(), monday.getDate() + 6);
  return sunday.getFullYear() + ('0' + (sunday.getMonth() + 1)).slice(-2) + ('0' + sunday.getDate()).slice(-2);
}

// '2026-10-12' を '20261012' に
function tightYmd_(ymd) { return String(ymd).replace(/-/g, ''); }

// その日（'20261012' の形）がふくまれる締め期間
function periodOfYmd_(ymdTight, closingDay) {
  var dash = ymdTight.slice(0, 4) + '-' + ymdTight.slice(4, 6) + '-' + ymdTight.slice(6, 8);
  var list = closingRanges_(dash, dash, closingDay);
  var p = list[list.length - 1];
  return { from: tightYmd_(p.start), to: tightYmd_(p.end), label: closingLabel_(p.end) };
}

// シートから読んだ中身を取り出す。覚えてあればそれを使う（重いのはここだけ）
// ※ 覚えておくのは「シートに書いてある事実」だけ。今日がいつか・どの週を出すか は毎回その場で決める。
//    そうしないと、日付が変わったときに覚えた内容が使えなくなる
function shiftSheetData_(now) {
  var cache = pageCache_();
  var fresh = null;

  if (cache) {
    var hit = cache.get(SHIFT_PAGE_CACHE_KEY);
    if (hit) { try { fresh = JSON.parse(hit); } catch (err) {} }
  }
  if (!fresh) {
    var saved = loadLong_();
    if (saved && (now.getTime() - Number(saved.builtAt || 0)) < SHIFT_PAGE_KEEP_DAYS * 86400000) {
      fresh = saved;
      if (cache) { try { cache.put(SHIFT_PAGE_CACHE_KEY, JSON.stringify(fresh), SHIFT_PAGE_CACHE_SEC); } catch (err) {} }
    }
  }
  if (!fresh) {
    fresh = readShiftSheets_(now);
    if (cache) { try { cache.put(SHIFT_PAGE_CACHE_KEY, JSON.stringify(fresh), SHIFT_PAGE_CACHE_SEC); } catch (err) {} }
    saveLong_(fresh);
  }
  return fresh;
}

function shiftPage_(weekParam) {
  var now = new Date();
  var data = shiftPageData_(now);
  data.openWeek = /^\d{6}$/.test(weekParam) ? weekParam : '';

  var template = HtmlService.createTemplateFromFile('shift');
  template.dataJson = JSON.stringify(data).replace(/</g, '\u003c');
  return template.evaluate()
    .setTitle('シフト｜' + data.storeName)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// どこで時間を使っているかを測る。?slow=1 から呼ぶ
function shiftPageTiming_() {
  var out = {}, t0 = new Date().getTime(), t = t0;
  function lap(name) { var now = new Date().getTime(); out[name] = (now - t) / 1000; t = now; }

  var shiftSs = SpreadsheetApp.openById(SHIFT_SPREADSHEET_ID);
  lap('1_シフト表をひらく');

  var settings = shiftSs.getSheetByName(REQUEST_SETTINGS_SHEET);
  var closingDays = { emp: closingDayOf_(settings, true), cast: closingDayOf_(settings, false) };
  lap('2_締め日の設定をよむ');

  var now = new Date();
  var weeks = readPublishedWeeks_(shiftSs, now);
  lap('3_週シートをよむ（' + weeks.length + '枚）');

  var names = staffNames_(shiftSs, weeks);
  lap('4_名前の一覧');

  holidayLabels_(shiftSs, weeks);
  lap('5_祝日');

  employeeNames_(shiftSs);
  lap('6_店長・社員');

  names.forEach(function(n) { calendarKey_(n); });
  lap('7_カレンダーの鍵');

  out['0_ぜんぶ'] = (new Date().getTime() - t0) / 1000;
  return out;
}

// シフト表を実際に読む（重い処理はここだけ）。ここで作った物を覚えておく
function readShiftSheets_(now) {
  var shiftSs = SpreadsheetApp.openById(SHIFT_SPREADSHEET_ID);
  var settings = shiftSs.getSheetByName(REQUEST_SETTINGS_SHEET);
  // 締め日は人によって違う（店長・社員＝E3／キャスト＝E7）
  var closingDays = { emp: closingDayOf_(settings, true), cast: closingDayOf_(settings, false) };
  // 店舗名はマス1つ。ここで設定ぜんぶを読み直すと、そのぶん遅くなる
  var storeName = STORE_NAME;
  try {
    if (settings) storeName = String(settings.getRange(SETTING_STORE_NAME_CELL).getDisplayValue()).trim() || STORE_NAME;
  } catch (err) {}

  var weeks = readPublishedWeeks_(shiftSs, now);
  var names = staffNames_(shiftSs, weeks);
  var keys = {};
  names.forEach(function(n) { keys[n] = calendarKey_(n); });

  return {
    builtAt: now.getTime(),
    storeName: storeName,
    closing: closingDays,
    names: names,
    keys: keys,
    holidays: holidayLabels_(shiftSs, weeks),
    weeks: weeks.map(function(wk) {
      return { sheet: wk.sheet, ymds: wk.ymds, labels: wk.labels, days: wk.days };
    }),
    employees: employeeNames_(shiftSs)          // この人たちには給与のボタンを出さない
  };
}

// 画面に渡す中身。重いところは覚えてある物を使い、今日に合わせた分だけここで足す
function shiftPageData_(now) {
  var base = shiftSheetData_(now);
  var todayYmd = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyyMMdd');
  var closingDays = base.closing || { emp: DEFAULT_BASE.closingDay, cast: DEFAULT_CAST_CLOSING_DAY };
  var firstYmd = [periodOfYmd_(todayYmd, closingDays.emp).from,
    periodOfYmd_(todayYmd, closingDays.cast).from].sort()[0];
  var floorYmd = weekFloorYmd_(now, firstYmd);     // 合計に要る、いちばん前の日
  var thisMonday = weekFloorYmd_(now, null);       // 今週の月曜

  var weeks = base.weeks.filter(function(wk) { return wk.ymds[6] >= floorYmd; })
    .map(function(wk) {
      return { sheet: wk.sheet, ymds: wk.ymds, labels: wk.labels, days: wk.days,
        // 終わった週はタブに出さない（合計には使う）
        past: wk.ymds[6] < thisMonday,
        // その週がふくまれる締め期間。合計を「締め日まで」で出すのに使う
        period: { emp: periodOfYmd_(wk.ymds[0], closingDays.emp),
          cast: periodOfYmd_(wk.ymds[0], closingDays.cast) } };
    });

  return {
    storeName: base.storeName,
    today: todayYmd,
    openWeek: '',
    names: base.names,
    keys: base.keys,
    holidays: base.holidays,
    weeks: weeks,
    selfUrl: SELF_URL,
    employees: base.employees
  };
}

// 名前の一覧＝一番新しい週シートの名簿＋公開した週に出てくる名前
function staffNames_(shiftSs, weeks) {
  var names = [];
  var seen = {};
  function add(n) {
    var key = normalizeName_(n);
    if (!key || seen[key]) return;
    seen[key] = true;
    names.push(String(n).trim());
  }
  var sheets = shiftSs.getSheets().filter(function(sh) { return /^\d{6}$/.test(sh.getName()); })
    .sort(function(a, b) { return a.getName() < b.getName() ? 1 : -1; });
  if (sheets.length) rosterNames_(sheets[0]).forEach(add);
  namesInWeeks_(weeks).forEach(add);
  return names;
}

// 公開した週の中にある祝日（{ '20261012': 'スポーツの日' } の形）
function holidayLabels_(shiftSs, weeks) {
  var out = {};
  if (!weeks.length) return out;
  var sheet = shiftSs.getSheetByName(HOLIDAY_SHEET_NAME);
  if (!sheet) return out;
  var last = sheet.getLastRow();
  if (last < 1) return out;
  var first = weeks[0].ymds[0], end = weeks[weeks.length - 1].ymds[6];
  sheet.getRange(1, 1, last, 2).getValues().forEach(function(r) {
    if (!(r[0] instanceof Date)) return;
    var ymd = Utilities.formatDate(r[0], 'Asia/Tokyo', 'yyyyMMdd');
    if (ymd < first || ymd > end) return;
    out[ymd] = String(r[1] || '祝日');
  });
  return out;
}

// ===================== 検体キットの提出チェック（?kit=2026-09） =====================
// 誰が出した・出していないは健康にかかわること。この画面には自分の分しか出さない。
// 記録はシフト表の「提出チェック」シート（店長だけが見る）に書く。
var KIT_SHEET_NAME = '提出チェック';
var KIT_VIEW_HEADERS = ['番号', '名前', 'LINE登録', '提出不要', '日にち', 'だれが入れたか'];
var KIT_FIRST_ROW = 3;                      // 1行目＝見出しの帯、2行目＝列の名前、3行目から中身
var KIT_OLD_HEADERS = ['月', '名前', '提出した日', 'だれが入れたか'];   // 前の形。1回だけの移しかえで読む
var KIT_OLD_LOG_COL = 7;                    // 前の形でG列に記録を置いていた分
var KIT_DOW_JP = ['日', '月', '火', '水', '木', '金', '土'];
var KIT_DUE_DAY_CELL = 'E18';        // シフト希望設定：毎月のおしらせの期日（何日）
var LIFF_ID_CELL = 'E39';            // シフト希望設定：LINEログイン（LIFF）のID
var LINE_LOGIN_ID_CELL = 'E40';      // シフト希望設定：LINEログインのチャネルID
var LINE_LOGIN_SECRET_CELL = 'E41';  // 　　　〃　　　　チャネルシークレット
var LINE_TOKEN_URL = 'https://api.line.me/oauth2/v2.1/token';
var LINE_PROFILE_URL = 'https://api.line.me/v2/profile';
var WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbwFtgsWnq75UnttAODig9q3c89E1Iek0Oow-lV8tHzF_Zp5qBQe2DXPahAeksqkd4i2/exec';
var LINE_OWNER_SHEET = 'LINEの持ち主';
var LINE_OWNER_HEADERS = ['LINEユーザーID', '名前', '最初に使った日時', '最後に使った日時', '使った回数'];

// シフト表の設定を1つ読む（無ければ空）
function shiftSetting_(shiftSs, cell) {
  try {
    var sheet = shiftSs.getSheetByName('シフト希望設定');
    return sheet ? String(sheet.getRange(cell).getDisplayValue()).trim() : '';
  } catch (err) { return ''; }
}

// LINEのユーザーIDと名前の対応表（アクセス記録の中。使うときに作る）
function lineOwnerSheet_() {
  var ss = getStore_();
  var sheet = ss.getSheetByName(LINE_OWNER_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(LINE_OWNER_SHEET);
    sheet.getRange(1, 1, 1, LINE_OWNER_HEADERS.length).setValues([LINE_OWNER_HEADERS]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// LINEのユーザーIDから名前を引く
function lineOwner_(lineId) {
  var id = String(lineId == null ? '' : lineId).trim();
  if (!id) return '';
  try {
    var sheet = lineOwnerSheet_();
    var last = sheet.getLastRow();
    if (last < 2) return '';
    var rows = sheet.getRange(2, 1, last - 1, 2).getValues();
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === id) return String(rows[i][1]).trim();
    }
  } catch (err) {}
  return '';
}

// LINEのユーザーIDと名前を結びつける（1回目だけ書き、2回目からは回数と日時を足す）
function bindLineOwner_(lineId, name) {
  var id = String(lineId == null ? '' : lineId).trim();
  var clean = String(name == null ? '' : name).trim();
  if (!id || !clean) return;
  try {
    var sheet = lineOwnerSheet_();
    var last = sheet.getLastRow();
    var rows = last >= 2 ? sheet.getRange(2, 1, last - 1, 1).getValues() : [];
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === id) {
        var row = i + 2;
        sheet.getRange(row, 2).setValue(clean);
        sheet.getRange(row, 4).setValue(new Date());
        sheet.getRange(row, 5).setValue(Number(sheet.getRange(row, 5).getValue() || 0) + 1);
        return;
      }
    }
    sheet.appendRow([id, clean, new Date(), new Date(), 1]);
  } catch (err) {}
}

// 置き場所は「シフト希望_アクセス記録」。飯塚店シフトはスタッフ全員が開くので置かない。
function kitSheet_() {
  var ss = getStore_();
  var sheet = ss.getSheetByName(KIT_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(KIT_SHEET_NAME);
  kitEnsureLayout_(sheet);
  return sheet;
}

// 提出チェックシートは「当月の一覧」だけを持つ。提出の履歴は残さない。
// 検体の結果は業者から届く書類が証明になるので、ここで見たいのは「全員そろって送れるか」だけ。
// A1＝この一覧がどの月ぶんか。C列＝LINEログインの登録が済んでいるか（月をまたいで残る）。
// D列＝提出不要（店長の手入力。実習や長期休みの人。月をまたいで残る）。
// E列に日にちが入っていれば提出ずみ。F列＝だれが入れたか。
function kitEnsureLayout_(sheet) {
  if (String(sheet.getRange(2, 4).getValue()).trim() === KIT_VIEW_HEADERS[3]) return false;
  var last = sheet.getLastRow();
  var ym = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM');
  var keep = [];                              // [名前, 提出不要, 日にち, だれが入れたか]
  if (last >= 2) {
    if (String(sheet.getRange(2, 3).getValue()).trim() === KIT_VIEW_HEADERS[2]) {
      var same = kitYm_(sheet.getRange(1, 1).getValue()) === ym;      // 前の形（提出不要の列が無い一覧）
      sheet.getRange(KIT_FIRST_ROW, 2, Math.max(last - KIT_FIRST_ROW + 1, 1), 4).getValues().forEach(function(r) {
        if (String(r[0]).trim()) keep.push([String(r[0]).trim(), '',
          same ? String(r[2]).trim() : '', same ? String(r[3]).trim() : '']);
      });
    } else {
      var logCol = String(sheet.getRange(2, KIT_OLD_LOG_COL).getValue()).trim() === KIT_OLD_HEADERS[0] ? KIT_OLD_LOG_COL
        : (String(sheet.getRange(1, 1).getValue()).trim() === KIT_OLD_HEADERS[0] ? 1 : 0);
      if (logCol) {                           // もっと前の形（記録が並んでいた）。当月ぶんだけ写す
        var logRow = logCol === 1 ? 2 : KIT_FIRST_ROW;
        sheet.getRange(logRow, logCol, Math.max(last - logRow + 1, 1), KIT_OLD_HEADERS.length).getValues().forEach(function(r) {
          if (kitYm_(r[0]) === ym && String(r[1]).trim()) {
            keep.push([String(r[1]).trim(), '', kitDayLabel_(r[2]), String(r[3] || '')]);
          }
        });
      }
    }
  }
  if (last >= 1) sheet.getRange(1, 1, last, 10).clearContent();
  sheet.getRange(1, 1).setNumberFormat('@');
  sheet.getRange(1, 1).setValue(ym);
  sheet.getRange(2, 1, 1, KIT_VIEW_HEADERS.length).setValues([KIT_VIEW_HEADERS]).setFontWeight('bold');
  sheet.getRange(2, 4).setNote('実習や長期の休みで、この月は出さなくてよい人に何か書いてください（例：実習）。'
    + '書いた人は人数から外れます。戻ってきたら消してください。');
  if (keep.length) {
    sheet.getRange(KIT_FIRST_ROW, 1, keep.length, KIT_VIEW_HEADERS.length).setValues(keep.map(function(r) {
      return ['', r[0], '', r[1], r[2], r[3]];
    }));
  }
  sheet.setFrozenRows(2);
  [46, 130, 90, 90, 96, 150].forEach(function(w, i) { sheet.setColumnWidth(i + 1, w); });
  return true;
}

// シートに書いてある月を '2026-09' の形にそろえる（日付で入ってしまった行も読めるように）
function kitYm_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM');
  var s = String(v == null ? '' : v).trim();
  if (/^[0-9]{4}-[0-9]{2}/.test(s)) return s.slice(0, 7);
  var d = new Date(s);                        // 「Tue Sep 01 2026 …」の形で入ってしまった分も読む
  return isNaN(d.getTime()) ? s.slice(0, 7) : Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM');
}

// この一覧が持っている月（A1）
function kitSheetMonth_(sheet) { return kitYm_(sheet.getRange(1, 1).getValue()); }

// 一覧に今書いてある中身 { そろえた名前: {name, exempt, day, by} }。
// sameMonth が false のときは、日にちとだれが入れたかは持ち越さない（月が変わったら白紙に戻す）
function kitTableMap_(sheet, sameMonth) {
  var map = {};
  var last = sheet.getLastRow();
  if (last < KIT_FIRST_ROW) return map;
  sheet.getRange(KIT_FIRST_ROW, 2, last - KIT_FIRST_ROW + 1, 5).getValues().forEach(function(r) {
    var name = String(r[0]).trim();
    if (!name) return;
    map[normalizeName_(name)] = { name: name, exempt: String(r[2]).trim(),
      day: sameMonth ? String(r[3]).trim() : '', by: sameMonth ? String(r[4]).trim() : '' };
  });
  return map;
}

// その月にもう出した人 { そろえた名前: {name, day, by} }。別の月の一覧なら空
function kitMonthMap_(ym) {
  var sheet = kitSheet_();
  var map = {};
  if (kitSheetMonth_(sheet) !== String(ym)) return map;
  var table = kitTableMap_(sheet, true);
  for (var key in table) if (table[key].day) map[key] = table[key];
  return map;
}

// LINEログインの登録が済んでいる人
function kitLineRegistered_() {
  var map = {};
  try {
    var sheet = getStore_().getSheetByName(LINE_OWNER_SHEET);
    if (!sheet) return map;
    var last = sheet.getLastRow();
    if (last < 2) return map;
    sheet.getRange(2, 2, last - 1, 1).getValues().forEach(function(r) {
      if (String(r[0]).trim()) map[normalizeName_(r[0])] = true;
    });
  } catch (e) {}
  return map;
}

// その月にもう出した人（名前のゆれをそろえた形で持つ）
function kitDoneMap_(ym) {
  var hits = kitMonthMap_(ym);
  var map = {};
  for (var key in hits) map[key] = hits[key].name;
  return map;
}

// いちばん新しい週シートの名簿
function kitStaffNames_(shiftSs) {
  var newest = null;                       // 名前がいちばん大きい週シート（＝いちばん新しい週）
  shiftSs.getSheets().forEach(function(sh) {
    if (!/^\d{6}$/.test(sh.getName())) return;
    if (!newest || sh.getName() > newest.getName()) newest = sh;
  });
  return newest ? rosterNames_(newest) : [];
}

// 今月（指定が無い・おかしいときに使う）
function kitThisMonth_() {
  return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM');
}

// 期日の見出し「9月25日(金)」。シフト表の設定にある「毎月この日」を使う
function kitDueLabel_(shiftSs, ym) {
  var day = 0;
  try {
    var settings = shiftSs.getSheetByName('シフト希望設定');
    if (settings) day = Number(String(settings.getRange(KIT_DUE_DAY_CELL).getDisplayValue()).trim());
  } catch (err) {}
  var p = ym.split('-').map(Number);
  if (!(day >= 1 && day <= 31)) return p[1] + '月';
  var last = new Date(p[0], p[1], 0).getDate();
  var due = new Date(p[0], p[1] - 1, Math.min(day, last));
  return (due.getMonth() + 1) + '月' + due.getDate() + '日(' + DOW_LABELS[(due.getDay() + 6) % 7] + ')';
}

// 【1回だけ】GASエディタでこれを実行し、出てくる画面で「許可」を押す。
// LINEに問い合わせる権限（UrlFetchApp）を承認するためだけの関数。
// メニューやページからは承認できないので、必ずエディタから実行する。
function authorizeNow() {
  var res = UrlFetchApp.fetch(LINE_PROFILE_URL, {
    muteHttpExceptions: true, headers: { Authorization: "Bearer dummy" }
  });
  SpreadsheetApp.openById(SHIFT_SPREADSHEET_ID).getName();
  getStore_().getName();
  return "通信の権限OK（LINEの返事：" + res.getResponseCode() + "）。これでLINEログインが使えます。";
}

// LINEログインで分かったユーザーID（分からなければ空）
function lineIdOf_(e) {
  return (e && e.parameter && e.parameter.lid) ? String(e.parameter.lid) : '';
}

// LINEログインを通ったのにうまくいかなかったときの理由（通っていなければ空）
function lineErrOf_(e) {
  return (e && e.parameter && e.parameter.lerr) ? String(e.parameter.lerr) : '';
}

function kitPage_(param, lineId, lineErr) {
  var ym = /^\d{4}-\d{2}$/.test(param) ? param : kitThisMonth_();
  var data = remember_('KIT_PAGE_V1_' + ym, function() {
    var shiftSs = SpreadsheetApp.openById(SHIFT_SPREADSHEET_ID);
    var name = STORE_NAME;
    try { name = loadConfig_(shiftSs).storeName || STORE_NAME; } catch (err) {}
    return {
      storeName: name,
      month: ym,
      monthLabel: Number(ym.split('-')[1]) + '月',
      dueLabel: kitDueLabel_(shiftSs, ym),
      names: kitStaffNames_(shiftSs),
      liffId: shiftSetting_(shiftSs, LIFF_ID_CELL)
    };
  });
  var storeName = data.storeName;
  var template = HtmlService.createTemplateFromFile('kit');
  template.dataJson = JSON.stringify(data).replace(/</g, '\u003c');
  template.lineIdJson = JSON.stringify(String(lineId || ''));
  template.lineErrJson = JSON.stringify(String(lineErr || ''));
  return template.evaluate()
    .setTitle('検体キットの提出｜' + storeName)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 画面から呼ぶ：このスマホの持ち主と、その人がもう出しているか
function kitWho(ym, deviceId, lineId) {
  var month = /^\d{4}-\d{2}$/.test(ym) ? ym : kitThisMonth_();
  // LINEログインで分かった人を先に見る（機種を変えても変わらないため）
  var fromLine = lineOwner_(lineId);
  var owner = fromLine || deviceOwner_(deviceId);
  // LINEを通ってきたのに、まだ名前と結びついていないときは、この場で覚える。
  // （ボタンを押さずに閉じる人もいるので、開いた時点で結びつけておく）
  if (lineId && !fromLine && owner) bindLineOwner_(lineId, owner);
  if (!owner) return { owner: '', done: false, by: '' };
  return { owner: owner, by: (fromLine || lineId) ? 'LINE' : '端末',
    done: !!kitDoneMap_(month)[normalizeName_(owner)] };
}

// 画面から呼ぶ：提出したと記録する
function kitSubmit(ym, name, deviceId, lineId) {
  var month = /^\d{4}-\d{2}$/.test(ym) ? ym : kitThisMonth_();
  var clean = String(name == null ? '' : name).trim();
  if (!clean) throw new Error('名前を選んでください。');
  var shiftSs = SpreadsheetApp.openById(SHIFT_SPREADSHEET_ID);
  var known = kitStaffNames_(shiftSs).filter(function(n) { return normalizeName_(n) === normalizeName_(clean); });
  if (!known.length) throw new Error('名簿にありません。店長に伝えてください。');
  bindLineOwner_(lineId, known[0]);            // 次からは名前を選ばなくてよくなる
  var owner = lineOwner_(lineId) || deviceOwner_(deviceId);
  var by = lineId ? 'LINEログイン'
    : (owner && normalizeName_(owner) === normalizeName_(clean) ? '本人' : '本人（名前を選んだ）');
  var res = kitWriteSummary_(shiftSs, month, { name: known[0], by: by });
  return { ok: true, already: res.already, name: known[0] };
}

// 提出した日の見出し「9/25(金)」
function kitDayLabel_(d) {
  if (!(d instanceof Date)) return String(d == null ? '' : d);
  return (d.getMonth() + 1) + '/' + d.getDate() + '(' + KIT_DOW_JP[d.getDay()] + ')';
}

// 店長が見る一覧を、提出チェックシートに書く。mark を渡すとその人を提出ずみにしてから書く。
// 名前が出るのはこのシートだけ（店長しか開かない）。LINEには人数しか出さない。
function kitWriteSummary_(shiftSs, ym, mark) {
  var sheet = kitSheet_();
  var table = kitTableMap_(sheet, kitSheetMonth_(sheet) === String(ym));
  var already = false;
  if (mark && mark.name) {
    var markKey = normalizeName_(mark.name);
    var was = table[markKey];
    already = !!(was && was.day);
    if (!already) {
      table[markKey] = { name: String(mark.name).trim(), exempt: was ? was.exempt : '',
        day: kitDayLabel_(new Date()), by: String(mark.by || '本人') };
    }
  }
  var names = kitStaffNames_(shiftSs);
  var line = kitLineRegistered_();
  var rows = [];
  var exemptRows = [];
  var target = 0;
  var done = 0;
  var lineDone = 0;
  var offCount = 0;
  for (var i = 0; i < names.length; i++) {
    var key = normalizeName_(names[i]);
    var hit = table[key] || {};
    if (line[key]) lineDone++;
    if (hit.exempt) { offCount++; exemptRows.push(KIT_FIRST_ROW + i); }
    else { target++; if (hit.day) done++; }
    rows.push([i + 1, names[i], line[key] ? '○' : '', hit.exempt || '', hit.day || '', hit.by || '']);
  }
  sheet.getRange(1, 1).setNumberFormat('@');
  sheet.getRange(1, 1, 1, KIT_VIEW_HEADERS.length).setValues([[
    String(ym),
    '期日 ' + kitDueLabel_(shiftSs, ym),
    'LINE登録 ' + lineDone + '人',
    offCount ? '提出不要 ' + offCount + '人' : '',
    target + '人中 ' + done + '人が提出',
    '未提出 ' + (target - done) + '人'
  ]]);
  sheet.getRange(2, 1, 1, KIT_VIEW_HEADERS.length).setValues([KIT_VIEW_HEADERS]);
  if (rows.length) {
    sheet.getRange(KIT_FIRST_ROW, 1, rows.length, KIT_VIEW_HEADERS.length).setValues(rows).setFontColor('#000000');
    exemptRows.forEach(function(r) { sheet.getRange(r, 1, 1, KIT_VIEW_HEADERS.length).setFontColor('#999999'); });
    sheet.getRange(KIT_FIRST_ROW, 3, rows.length, 1).setHorizontalAlignment('center');
  }
  var below = KIT_FIRST_ROW + rows.length;              // 名簿が減ったときの余りを消す
  var last = sheet.getLastRow();
  if (last >= below) sheet.getRange(below, 1, last - below + 1, KIT_VIEW_HEADERS.length).clearContent();
  sheet.getRange(1, 1, 2, KIT_VIEW_HEADERS.length).setFontWeight('bold');
  sheet.getRange(1, 6).setFontColor(target - done > 0 ? '#C0392B' : '#1E8449');
  return { total: target, done: done, left: target - done, off: offCount, already: already };
}

// ===================== LINEログインで本人を確かめる（シフト希望ページ用） =====================
// LINEのユーザーIDは機種を変えても同じなので、端末IDのように消えない。
// 分からないとき（LINE以外から開いた・まだ結びついていない）は空を返し、今までどおり名前を選んでもらう。
function whoAmI(lineId) {
  var name = lineOwner_(lineId);
  return { name: name || '', by: name ? 'LINE' : '' };
}

// LINEログインの登録が何人ぶんあるかを返す（人数と日付だけ。名前は出さない）
function lineOwnerCount_() {
  var out = { sheet: false, count: 0, first: '', last: '' };
  try {
    var sheet = getStore_().getSheetByName(LINE_OWNER_SHEET);
    if (sheet) {
      out.sheet = true;
      var last = sheet.getLastRow();
      if (last >= 2) {
        var rows = sheet.getRange(2, 1, last - 1, LINE_OWNER_HEADERS.length).getValues()
          .filter(function(r) { return String(r[0]).trim() !== ''; });
        out.count = rows.length;
        rows.forEach(function(r) {
          var f = r[2] instanceof Date ? Utilities.formatDate(r[2], 'Asia/Tokyo', 'yyyy-MM-dd HH:mm') : '';
          var l = r[3] instanceof Date ? Utilities.formatDate(r[3], 'Asia/Tokyo', 'yyyy-MM-dd HH:mm') : '';
          if (f && (!out.first || f < out.first)) out.first = f;
          if (l && l > out.last) out.last = l;
        });
      }
    }
  } catch (err) { out.error = err.message; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

// 名前を選んだときに、LINEのユーザーIDと結びつける（次からは選ばなくてよくなる）
function bindMe(lineId, name) {
  var id = String(lineId == null ? '' : lineId).trim();
  var clean = String(name == null ? '' : name).trim();
  if (!id || !clean) return false;
  // すでに誰かと結びついているLINEは、別の名前で上書きしない。
  // 上書きできると、ほかの人の名前を選ぶだけでその人になりすませてしまう（2026-09-28）。
  // 間違って覚えた時は、店長が「LINEの持ち主」シートで直す
  var already = lineOwner_(id);
  if (already) return normalizeName_(already) === normalizeName_(clean);
  var staff = [];
  try { staff = cachedConfig_().staff || []; } catch (err) { return false; }
  var known = staff.filter(function(n) { return normalizeName_(n) === normalizeName_(clean); });
  if (!known.length) return false;                     // 名簿にない名前は結びつけない
  bindLineOwner_(id, known[0]);
  return true;
}
