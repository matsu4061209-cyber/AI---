/*******************************************************
 * 予定カレンダー（クラウド保存版）サーバー側
 * データは専用スプレッドシートに自動保存。
 * 初回実行時にスプレッドシートを自動作成し、IDを記憶する。
 *******************************************************/

var PROP_KEY = 'CAL_SHEET_ID';

// ウェブアプリを開いたときに画面を返す
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('予定カレンダー')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// データ用スプレッドシートを取得（無ければ作る）
function getSS_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty(PROP_KEY);
  var ss = null;
  if (id) {
    try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; }
  }
  if (!ss) {
    ss = SpreadsheetApp.create('予定カレンダー データ');
    props.setProperty(PROP_KEY, ss.getId());
  }
  ensureSheets_(ss);
  return ss;
}

// 必要なシートと見出しを用意
function ensureSheets_(ss) {
  var ev = ss.getSheetByName('予定');
  if (!ev) { ev = ss.insertSheet('予定'); ev.appendRow(['id', 'date', 'text']); }
  var rc = ss.getSheetByName('毎月');
  if (!rc) { rc = ss.insertSheet('毎月'); rc.appendRow(['id', 'day', 'text']); }
  var cp = ss.getSheetByName('期間');
  if (!cp) { cp = ss.insertSheet('期間'); cp.appendRow(['id', 'start', 'end', 'text', 'color']); }
  // 日付・曜日列を「文字列」に固定（勝手に日付変換されないように）
  ev.getRange('B:B').setNumberFormat('@');
  rc.getRange('B:B').setNumberFormat('@');
  cp.getRange('B:C').setNumberFormat('@');
  // 自動でできる空シートを削除
  var s1 = ss.getSheetByName('シート1') || ss.getSheetByName('Sheet1');
  if (s1 && ss.getSheets().length > 1) { ss.deleteSheet(s1); }
}

// 全データを読み出して返す
function readAll_() {
  var ss = getSS_();
  var events = {};
  var evVals = ss.getSheetByName('予定').getDataRange().getValues();
  for (var i = 1; i < evVals.length; i++) {
    var id = evVals[i][0], date = String(evVals[i][1]), text = evVals[i][2];
    if (!id || !date) continue;
    if (!events[date]) events[date] = [];
    events[date].push({ id: String(id), text: String(text) });
  }
  var recurring = [];
  var rcVals = ss.getSheetByName('毎月').getDataRange().getValues();
  for (var j = 1; j < rcVals.length; j++) {
    var rid = rcVals[j][0];
    if (!rid) continue;
    var dayRaw = String(rcVals[j][1]);
    // 数字だけなら数値、それ以外は文字列のまま（last / nth:.. / even:.. / odd:..）
    var day = /^\d+$/.test(dayRaw) ? Number(dayRaw) : dayRaw;
    recurring.push({
      id: String(rid),
      day: day,
      text: String(rcVals[j][2])
    });
  }
  var periods = [];
  var cpVals = ss.getSheetByName('期間').getDataRange().getValues();
  for (var p = 1; p < cpVals.length; p++) {
    var pid = cpVals[p][0];
    if (!pid) continue;
    periods.push({
      id: String(pid),
      start: String(cpVals[p][1]),
      end: String(cpVals[p][2]),
      text: String(cpVals[p][3]),
      color: String(cpVals[p][4])
    });
  }
  return { events: events, recurring: recurring, periods: periods };
}

// ---- 画面から呼ばれる関数（毎回、最新の全データを返す） ----

function apiGetAll() { return readAll_(); }

function apiAddPeriod(start, end, text, color) {
  text = String(text || '').trim();
  if (!text || !start || !end) return readAll_();
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    getSS_().getSheetByName('期間')
      .appendRow([Utilities.getUuid(), String(start), String(end), text, String(color || '#3f7fd6')]);
  } finally { lock.releaseLock(); }
  return readAll_();
}

function apiDeletePeriod(id) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = getSS_().getSheetByName('期間');
    var vals = sh.getDataRange().getValues();
    for (var i = vals.length - 1; i >= 1; i--) {
      if (String(vals[i][0]) === String(id)) sh.deleteRow(i + 1);
    }
  } finally { lock.releaseLock(); }
  return readAll_();
}

function apiAddEvent(date, text) {
  text = String(text || '').trim();
  if (!text) return readAll_();
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    getSS_().getSheetByName('予定').appendRow([Utilities.getUuid(), String(date), text]);
  } finally { lock.releaseLock(); }
  return readAll_();
}

function apiDeleteEvent(id) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = getSS_().getSheetByName('予定');
    var vals = sh.getDataRange().getValues();
    for (var i = vals.length - 1; i >= 1; i--) {
      if (String(vals[i][0]) === String(id)) sh.deleteRow(i + 1);
    }
  } finally { lock.releaseLock(); }
  return readAll_();
}

function apiAddRecurring(day, text) {
  text = String(text || '').trim();
  if (!text) return readAll_();
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    getSS_().getSheetByName('毎月').appendRow([Utilities.getUuid(), String(day), text]);
  } finally { lock.releaseLock(); }
  return readAll_();
}

function apiDeleteRecurring(id) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = getSS_().getSheetByName('毎月');
    var vals = sh.getDataRange().getValues();
    for (var i = vals.length - 1; i >= 1; i--) {
      if (String(vals[i][0]) === String(id)) sh.deleteRow(i + 1);
    }
  } finally { lock.releaseLock(); }
  return readAll_();
}
