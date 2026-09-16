// ===== フォームの決まりごと（募集期間・締切・名簿・祝日はシフト表から読む） =====
var STORE_NAME = 'ペッパーランチ飯塚店';
var NOTE_CHOICES = ['6時間以内希望', 'ロング希望', '試験休み', '旅行のため', '未定'];
var NOTE_SEPARATOR = '／';            // 備考欄に「選択／手入力」の形で書く
var START_HOURS = { first: 9, last: 19 };   // 出勤は 9:00〜19:30
var END_HOURS = { first: 12, last: 21 };    // 退勤は 12:00〜21:30（21:00はなし）

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
var BASE_DEADLINE_DAYS_CELL = 'E4';   // 基本設定：締切＝期間の最終日の何日前
var BASE_DEADLINE_TIME_CELL = 'E5';   // 基本設定：締切の時刻
var DEFAULT_BASE = { closingDay: 15, deadlineDays: 14, deadlineTime: '23:59' };
var CAST_CLOSING_DAY_CELL = 'E7';     // キャストの締め日（店長・社員以外の人の出勤の区切り）
var DEFAULT_CAST_CLOSING_DAY = 20;
var ROLE_SHEET_NAME = 'スタッフ区分';  // A列＝名前、B列＝区分。店長・社員だけが載っていて、載っていない人はキャスト
var EMPLOYEE_ROLES = ['店長', '社員'];
var HISTORY_SHEET_NAME = '出勤履歴';  // A日付・C名前（シフト表GASが週シートの右側から作る。1人1日1行）
var PENDING_FIRST_ROW = 9;            // まだ週シートが無い週の希望
var PENDING_COLS = 19;                // シート名・名前・時間14・備考・送信日時・端末ID
var PENDING_DEVICE_INDEX = 18;

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

function doGet() {
  var config;
  try {
    config = loadConfig_(SpreadsheetApp.openById(SHIFT_SPREADSHEET_ID));
  } catch (e) {
    console.error('設定の読み込みに失敗: ' + e.message);
    config = { closed: true };
  }
  if (!config.storeName) config.storeName = STORE_NAME;
  var template = HtmlService.createTemplateFromFile('index');
  template.configJson = JSON.stringify(config).replace(/</g, '\u003c');
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
    if (config.deadline) title += ' 締切' + config.deadline;
  }
  return title + '｜' + config.storeName;
}

function mdText_(ymd) {
  var p = ymd.split('-').map(Number);
  return p[1] + '/' + p[2];
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
    return {
      weeks: readWeekStates_(shiftSs, config, name).map(function(s) {
        return { shifts: s.shifts, note: s.note };
      }),
      totals: closingPeriods_(shiftSs, config, name)
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
    holidays: holidaysInPeriod_(shiftSs, start, weekCount),
    staff: staff,
    noteChoices: NOTE_CHOICES,
    startHours: START_HOURS,
    endHours: END_HOURS
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
  return autoPeriod_(readBase_(settings), now);
}

function readBase_(settings) {
  var closing = Number(settings.getRange(BASE_CLOSING_DAY_CELL).getDisplayValue());
  var daysText = String(settings.getRange(BASE_DEADLINE_DAYS_CELL).getDisplayValue()).trim();
  var time = String(settings.getRange(BASE_DEADLINE_TIME_CELL).getDisplayValue()).trim();
  return {
    closingDay: closing >= 1 && closing <= 31 ? Math.floor(closing) : DEFAULT_BASE.closingDay,
    deadlineDays: daysText !== '' && Number(daysText) >= 0 && Number(daysText) <= 28 ? Math.floor(Number(daysText)) : DEFAULT_BASE.deadlineDays,
    deadlineTime: /^\d{2}:\d{2}$/.test(time) ? time : DEFAULT_BASE.deadlineTime
  };
}

// 自動の募集期間：前の期間の最終日の翌日（月曜）〜締め日が入っている週の日曜日。
// 締切＝期間の最終日の◯日前。締切を過ぎた期間と、書き込まない週（FIRST_WRITABLE_SHEETより前）から始まる期間は飛ばす
function autoPeriod_(base, now) {
  var ym = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM').split('-').map(Number);
  for (var i = -1; i <= 13; i++) {
    var end = periodEnd_(ym[0], ym[1] + i, base.closingDay);
    var start = addDays_(periodEnd_(ym[0], ym[1] + i - 1, base.closingDay), 1);
    if (sheetNameOf_(start) < FIRST_WRITABLE_SHEET) continue;
    var deadline = addDays_(end, -base.deadlineDays) + ' ' + base.deadlineTime;
    if (jstTime_(deadline) <= now.getTime()) continue;
    return { start: start, weeks: (daysBetween_(start, end) + 1) / 7, deadline: deadline, manual: false };
  }
  return null;
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

/* ================= 締め期間ごとの合計（参考） ================= */

// 画面に出ている週にかかる、この人の締め期間（店長・社員は社員の締め日、キャストはキャストの締め日で区切る）。
// 表示は「11月〆」のように締める月で出す。店長・社員は休日日数、キャストは出勤日数を数えるので、showOff で伝える。
// 締め期間のうち画面に出ていない日（募集期間より前の、シフトが決まっている日など）の出勤日数は「出勤履歴」から先に数えておく
// （画面に出ている日は、画面の入力で数える。休日日数は「締め期間の日数 − 出勤日数」で画面が出す）
function closingPeriods_(shiftSs, config, name) {
  var settings = shiftSs.getSheetByName(REQUEST_SETTINGS_SHEET);
  var employee = isEmployee_(shiftSs, name);
  var formStart = config.startDate;
  var formEnd = addDays_(config.startDate, config.weekCount * 7 - 1);
  var periods = closingRanges_(formStart, formEnd, closingDayOf_(settings, employee)).map(function(p) {
    return { start: p.start, end: p.end, label: Number(p.end.split('-')[1]) + '月〆', days: daysBetween_(p.start, p.end) + 1, fixedDays: 0 };
  });
  var result = { showOff: employee, periods: periods };

  var history = shiftSs.getSheetByName(HISTORY_SHEET_NAME);
  if (!history || history.getLastRow() < 2) return result;
  var tz = shiftSs.getSpreadsheetTimeZone();   // 出勤履歴の日付は、シフト表の時間帯の0時で入っている
  var counted = {};
  history.getRange(2, 1, history.getLastRow() - 1, 3).getValues().forEach(function(r) {
    if (!(r[0] instanceof Date) || String(r[2]).trim() !== name) return;
    var ymd = Utilities.formatDate(r[0], tz, 'yyyy-MM-dd');
    if ((ymd >= formStart && ymd <= formEnd) || counted[ymd]) return;
    var period = periods.filter(function(p) { return ymd >= p.start && ymd <= p.end; })[0];
    if (!period) return;
    counted[ymd] = true;
    period.fixedDays++;
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
    if (!unchanged && !validShift_(v.s, v.e)) throw new Error('時間の内容が正しくありません。');
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

function toMinutes_(t) {
  var m = /^(\d{2}):(00|30)$/.exec(String(t));
  return m ? Number(m[1]) * 60 + Number(m[2]) : null;
}

function anyMinutes_(t) {
  var m = /^(\d{1,2}):(\d{2})$/.exec(String(t));
  return m ? Number(m[1]) * 60 + Number(m[2]) : null;
}

function validShift_(s, e) {
  var sm = toMinutes_(s), em = toMinutes_(e);
  if (sm === null || em === null) return false;
  if (sm < START_HOURS.first * 60 || sm > START_HOURS.last * 60 + 30) return false;
  if (em < END_HOURS.first * 60 || em > END_HOURS.last * 60 + 30 || em === END_HOURS.last * 60) return false;
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
