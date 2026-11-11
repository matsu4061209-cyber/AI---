// スタッフ名簿の場所（4行目〜25行目＝22人分）
var ROSTER_START_ROW = 4;
var ROSTER_ROWS = 22;

// 店長・社員の登録先（非表示シート）。登録されていない人はキャスト扱い
var ROLE_SHEET_NAME = 'スタッフ区分';
var ROLE_MANAGER = '店長';
var ROLE_EMPLOYEE = '社員';

// 名前で時間数を計算する仕組みを入れる最初の週シート（これより前のシートには触らない）
var FIRST_ROLE_SHEET = '261019';

// 上部集計の行
var ROW_CAST_HOURS = 9;
var ROW_MANAGER_HOURS = 11;
var ROW_EMPLOYEE_HOURS = 12;

// 曜日ごとの「明細の名前(番号1〜16)」「1行ごとの時間合計」「上部集計の列」
var DAY_BLOCKS = [
  { col: 'Y',  names: 'T24:T39',   hours: 'CI24:CI39' }, // 月
  { col: 'AI', names: 'T43:T58',   hours: 'CI43:CI58' }, // 火
  { col: 'AS', names: 'T62:T77',   hours: 'CI62:CI77' }, // 水
  { col: 'BC', names: 'DD4:DD19',  hours: 'FS4:FS19' },  // 木
  { col: 'BM', names: 'DD24:DD39', hours: 'FS24:FS39' }, // 金
  { col: 'BW', names: 'DD43:DD58', hours: 'FS43:FS58' }, // 土
  { col: 'CG', names: 'DD62:DD77', hours: 'FS62:FS77' }  // 日
];

function createNextWeekSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const currentSheet = ss.getActiveSheet();

  try {
    createWeekSheetFrom_(ss, currentSheet);

  } catch (e) {
    // ★エラーの本当の原因を突き止めるための一時的な処理。
    //   画面に詳しい内容をそのまま表示し、あとで確認できるようにも保存しておく。
    var detail = '次週シート作成でエラー: ' + e.message + '\n' + (e.stack || '');
    PropertiesService.getDocumentProperties().setProperty('lastError', detail);
    Logger.log(detail);
    SpreadsheetApp.getUi().alert(detail);
  }
}

// 1枚の週シートから、その次の週のシートを作る（「次週シート作成」ボタンからも、毎日の自動点検からも使う）
function createWeekSheetFrom_(ss, currentSheet) {
  // 0. 複製する前に「今のシート」の余分な空行・空列を縮めておく
  //    （ワークブック全体のセル数が上限(1000万)に近づくと複製自体が失敗するため、
  //      次週シート作成のたびに毎回自動でこれを行い、二度と同じ問題が起きないようにする）
  trimSheetSize_(currentSheet);

  // 1. T1セルから日付を取得
  const currentDate = new Date(currentSheet.getRange("T1").getValue());

  // 2. 7日後を計算する。
  //    ★24時間×7で足すと、夏時間が切り替わる週（11月の初め・3月の初め）に1時間ずれて前日になってしまう。
  //      そのため「日付だけで7日進めて、その日の0時にする」形にしている（ファイルの時間帯で数える）。
  const tz = ss.getSpreadsheetTimeZone();
  const nextYmd = addDaysYmd_(Utilities.formatDate(currentDate, tz, 'yyyy-MM-dd'), 7);
  const nextDate = Utilities.parseDate(nextYmd + ' 00:00', tz, 'yyyy-MM-dd HH:mm');

  // 3. YYMMDD形式の名前を作成 (例: 260629)
  const nextSheetName = nextYmd.slice(2).replace(/-/g, '');

  // 4. シートを複製して名前を設定
  const newSheet = currentSheet.copyTo(ss).setName(nextSheetName);

  // 5. 新しいシートの T1 に新しい日付を書き込む
  newSheet.getRange("T1").setValue(nextDate);

  // 6. シフト時間入力表(B4:P25)をクリア。
  //    ただし「時間を固定するスタッフを設定」でチェックされている人は消さない
  var protectedNames = getProtectedStaffNames_();
  var rosterNames = newSheet.getRange(ROSTER_START_ROW, 1, ROSTER_ROWS, 1).getValues();
  for (var i = 0; i < rosterNames.length; i++) {
    var name = rosterNames[i][0];
    if (name && protectedNames.indexOf(name) !== -1) continue; // この人は消さない
    newSheet.getRange(ROSTER_START_ROW + i, 2, 1, 15).clearContent(); // B〜P列
  }

  // 6-2. スマホのシフト希望フォームで、この週の希望が先に送られていればシートに入れる
  applyPendingRequests_(newSheet);

  // 7. 右側の日別シフト表（転写ボタンが書き込む場所）も空にする
  const clearRangesStr = ["T24:X39", "T43:X58", "T62:X77", "DD4:DH19", "DD24:DH39", "DD43:DH58", "DD62:DH77"];
  clearRangesStr.forEach(rangeStr => newSheet.getRange(rangeStr).clearContent());

  // 7-2. ポジションの色は新しい週に持ち越さない
  resetPositionColorsOnCopy_(newSheet);
  fixDayTimeFormat_(newSheet);      // 出勤・退勤・休入・休終を「900→9:00」の書式にそろえる

  // 8. 祝日の色付けを新しいシートにも反映する（「祝日」シートは既にあるものを使う）
  applyHolidayColoring_(newSheet);

  // 9. シフト時間入力表(B4:O25)の表示形式を、常に「9:00」のようにコロン付きになるよう整える
  fixTimeFormat_(newSheet);

  // 9-2. その週の売上予算（5行目）を「売上予算」シートから入れる
  applyWeekBudget_(newSheet, readBudgets_(ss));

  // 9-3. その週の「目安」を「目安人数」シートから入れる
  applyWeekTargets_(newSheet, readTargets_(), holidayKeys_(), ss.getSpreadsheetTimeZone());

  // 10. 一番左に移動してアクティブにする
  ss.setActiveSheet(newSheet);
  ss.moveActiveSheet(1);

  // 11. 終わった週の出勤を「出勤履歴」に取り込み、設定・集計のシートの余白も詰めておく
  refreshWeekHistory_(currentSheet);
  [REQUEST_SETTINGS_SHEET, ROLE_SHEET_NAME, HISTORY_SHEET_NAME, MONTHLY_SHEET_NAME, BUDGET_SHEET_NAME, TARGET_SHEET_NAME, ADJUST_SHEET_NAME].forEach(function(name) {
    var s = ss.getSheetByName(name);
    if (s) trimSheetSize_(s);
  });
  // 名簿が増減していることがあるので、日別の表も作り直す（シートを作ってある時だけ）
  if (ss.getSheetByName(DAILY_SHEET_NAME)) refreshDailyHours_();

  // 12. 直近の週だけを残し、古い週は書庫ファイルへ移す（ファイルが重くならないようにする）
  archiveOldWeeks_(true);

  return newSheet;
}

// シフト時間入力表(B4:O25)の表示形式を「900」と入力しても「9:00」と表示されるように統一する
// ★誤って結合されてしまったセル（例：L5等）があると書式設定でエラーになるため、
//   結合を自動で解除してから表示形式を設定する
function fixTimeFormat_(sheet) {
  var range = sheet.getRange(ROSTER_START_ROW, 2, ROSTER_ROWS, 14); // B4:O25
  var mergedRanges = range.getMergedRanges();
  mergedRanges.forEach(function(m) { m.breakApart(); });
  range.setNumberFormat('0":"00');
}

// 右側の日別シフト表の「出勤・退勤・休入・休終」も、900と打てば 9:00 と出るようにする。
// 転写は「書いた行の出勤・退勤」しか直していなかったので、休憩の2列と、
// 誰も入っていない行の書式が抜けていた＝場所によって 1600 のまま出る、の正体。MARKER_DAY_TIME_FORMAT
function fixDayTimeFormat_(week) {
  ADJUST_BLOCKS.forEach(function(block) {
    var range = week.getRange(block.head + 1, block.col - 4, ADJUST_ROWS, 4);   // 出勤〜休終
    range.getMergedRanges().forEach(function(m) { m.breakApart(); });
    range.setNumberFormat('0":"00');
  });
}

// 「16:00」と打つと本物の時刻になってしまうことがある。この表は 1600 という数字で持つ決まりなので、
// 時刻のまま置いておくと升目にも出ず、決まりの確認も効かなくなる。打ち方に関わらず数字に直す。
// 中身ではなく「画面の見え方」から読む（古い日付の時間帯のズレに引っかからないため）
function fixTypedTimes_(week, range) {
  var r0 = range.getRow(), c0 = range.getColumn();
  var r1 = r0 + range.getNumRows() - 1, c1 = c0 + range.getNumColumns() - 1;
  var fixed = 0;
  ADJUST_BLOCKS.forEach(function(block) {
    var left = block.col - 4, right = block.col - 1;
    if (c1 < left || c0 > right) return;
    var top = Math.max(r0, block.head + 1), bottom = Math.min(r1, block.head + ADJUST_ROWS);
    if (top > bottom) return;
    var from = Math.max(c0, left), to = Math.min(c1, right);
    var area = week.getRange(top, from, bottom - top + 1, to - from + 1);
    var values = area.getValues(), shown = area.getDisplayValues(), changed = false;
    for (var i = 0; i < values.length; i++) {
      for (var j = 0; j < values[i].length; j++) {
        if (!(values[i][j] instanceof Date)) continue;
        var m = /^(\d{1,2}):(\d{2})/.exec(String(shown[i][j]).trim());
        if (!m) continue;
        values[i][j] = Number(m[1]) * 100 + Number(m[2]);
        changed = true;
        fixed++;
      }
    }
    if (changed) { area.setValues(values); area.setNumberFormat('0":"00'); }
  });
  return fixed;
}

// 休入に時間を入れたら、休終をその1時間後にしておく（休憩はたいてい1時間のため）。
// 入れるのは「休終が空のとき」と「休終が休入より後になっていないとき」だけ。
// 手で入れた休終は上書きしない。MARKER_BREAK_PLUS_HOUR
var BREAK_DEFAULT_MINUTES = 60;

function fillBreakEnd_(week, range) {
  var r0 = range.getRow(), c0 = range.getColumn();
  var r1 = r0 + range.getNumRows() - 1, c1 = c0 + range.getNumColumns() - 1;
  var filled = 0;
  ADJUST_BLOCKS.forEach(function(block) {
    var breakCol = block.col - 2;                 // 休入（休終はその右どなり）
    if (c1 < breakCol || c0 > breakCol) return;
    var top = Math.max(r0, block.head + 1), bottom = Math.min(r1, block.head + ADJUST_ROWS);
    if (top > bottom) return;
    var rows = bottom - top + 1;
    var startRange = week.getRange(top, breakCol, rows, 1);
    var endRange = week.getRange(top, breakCol + 1, rows, 1);
    var starts = startRange.getValues(), startsShown = startRange.getDisplayValues();
    var ends = endRange.getValues(), endsShown = endRange.getDisplayValues();
    var changed = false;
    for (var i = 0; i < rows; i++) {
      var from = minutesOfTime_(timeNumber_(starts[i][0], startsShown[i][0]));
      if (from === null) continue;                // 休入が空・読めないときは何もしない
      var now = minutesOfTime_(timeNumber_(ends[i][0], endsShown[i][0]));
      if (now !== null && now > from) continue;   // すでに入っている休終は触らない
      var to = from + BREAK_DEFAULT_MINUTES;
      ends[i][0] = Math.floor(to / 60) * 100 + (to % 60);
      changed = true;
      filled++;
    }
    if (changed) endRange.setValues(ends).setNumberFormat('0":"00');
  });
  return filled;
}

// メニューが無くても実行できるように、Appsスクリプトエディタから直接実行するための入り口
// （実行後、対象シートの出勤・退勤欄はすべて「9:00」のようなコロン付き表示になる）
function fixTimeFormatNow() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  fixTimeFormat_(sheet);
  Logger.log('時間の表示形式を直しました。（対象シート: ' + sheet.getName() + '）');
}

// 1枚のシートについて、「データが入っていない余分な行・列」だけを削除する共通処理。
// データが入っている部分（実際に使っている範囲）は一切削除しない。
// ★createNextWeekSheet からも毎回自動で呼ばれるため、今後は手動でのメンテナンス不要。
function trimSheetSize_(sheet) {
  var ROW_BUFFER = 30; // 実際に使っている行数より、余裕として残しておく行数
  var COL_BUFFER = 15; // 同じく列数の余裕

  if (sheet.getName() === '祝日') {
    // 祝日シートは行数(日付データ)を使うので対象外。列だけ縮める
    var hCols = sheet.getMaxColumns();
    var targetHCols = 10;
    if (hCols > targetHCols) {
      sheet.deleteColumns(targetHCols + 1, hCols - targetHCols);
    }
    return;
  }

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var maxRows = sheet.getMaxRows();
  var maxCols = sheet.getMaxColumns();
  var targetRows = Math.max(lastRow + ROW_BUFFER, 50);
  var targetCols = Math.max(lastCol + COL_BUFFER, 50);

  if (maxRows > targetRows) {
    sheet.deleteRows(targetRows + 1, maxRows - targetRows);
  }
  if (maxCols > targetCols) {
    sheet.deleteColumns(targetCols + 1, maxCols - targetCols);
  }
}

// ★今回のセル数上限トラブルを解消するための、1回だけ実行する整理用スクリプト。
//   既存の全シート（過去の週シートを含む）に対して trimSheetSize_ をまとめて実行する。
function trimSheetSizes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logLines = [];
  ss.getSheets().forEach(function(sheet) {
    var before = sheet.getMaxRows() + '×' + sheet.getMaxColumns();
    trimSheetSize_(sheet);
    var after = sheet.getMaxRows() + '×' + sheet.getMaxColumns();
    logLines.push(sheet.getName() + ': ' + before + ' → ' + after);
  });
  Logger.log('シート整理が完了しました。\n' + logLines.join('\n'));
}

// スプレッドシートを開いたときに「スタッフ管理」メニューを追加する
// メニューに出すポジション名（onOpenではシートを作らないよう、読むだけ・失敗したら初期値）
function positionNamesForMenu_() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REQUEST_SETTINGS_SHEET);
    if (!sheet) return POSITION_DEFAULTS.map(function(p) { return p.name; });
    return positionCells_(sheet)
      .map(function(r, i) { return String(r[0]).trim() || POSITION_DEFAULTS[i].name; });
  } catch (e) {
    return POSITION_DEFAULTS.map(function(p) { return p.name; });
  }
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var posNames = positionNamesForMenu_();
  // 上から「毎週やること」→「月ごと」→「設定のまとまり」の順。
  // ふだん使わないものは下のまとまりに入れて、毎週使うものを探しやすくする
  ui.createMenu('スタッフ管理')
    // ----- 毎週のシフト作り -----
    .addItem('希望と今の表をくらべる', 'openRequestCompare')
    .addItem('決まりを点検する（未成年・休憩）', 'checkWeekLawNow')
    .addItem('足りないところを募集する', 'openRecruit')
    .addItem('目安人数を週シートに入れ直す', 'refreshWeekTargets')
    .addSeparator()
    // ----- できたシフトを知らせる -----
    .addItem('お試し公開（スタッフには送らない）', 'publishWeekTest')
    .addItem('確定する（スタッフに公開・お知らせ）', 'confirmAdjust')
    .addItem('もう一度知らせる（PDFを作り直す）', 'publishWeekAgain')
    .addSeparator()
    // ----- 月ごと -----
    .addItem('月次集計を開く', 'openMonthlySummary')
    .addItem('日別の勤務時間を開く', 'openDailyHours')
    .addItem('毎月のおしらせ', 'openMonthlyNotice')
    .addItem('古い週を書庫へ移す', 'archiveOldWeeks')
    .addSeparator()
    // ----- まとまり（ふだんは開かない設定） -----
    .addSubMenu(ui.createMenu('スタッフの設定')
      .addItem('スタッフを追加', 'addStaffMember')
      .addItem('スタッフを削除', 'removeStaffMember')
      .addSeparator()
      .addItem('店長・社員を設定', 'setStaffRoles')
      .addItem('時間を固定するスタッフを設定', 'setProtectedStaff')
      .addSeparator()
      .addItem('スタッフの生年月日', 'openMinorSettings')
      .addItem('時給と交通費の設定', 'openWageSettings'))
    .addSubMenu(ui.createMenu('目安人数と売上予算')
      .addItem('目安人数を開く', 'openTargetSheet')
      .addItem('売上予算を入れ直す', 'refreshWeekBudgets'))
    .addSubMenu(ui.createMenu('シフト希望とお知らせ')
      .addItem('シフト希望の募集設定', 'openShiftRequestSettings')
      .addItem('お知らせの文面・日程', 'openNoticeSettings')
      .addSeparator()
      .addItem('毎月のおしらせをお試しで送る（スタッフには送りません）', 'sendMonthlyNoticeTest'))
    .addSubMenu(ui.createMenu('ポジション色分け')
      .addItem(posNames[0] + 'に塗る', 'paintPosition1')
      .addItem(posNames[1] + 'に塗る', 'paintPosition2')
      .addItem(posNames[2] + 'に塗る', 'paintPosition3')
      .addItem(posNames[3] + 'に塗る', 'paintPosition4')
      .addItem(posNames[4] + 'に塗る', 'paintPosition5')
      .addItem('選んだところの色を戻す', 'clearPositionPaint')
      .addSeparator()
      .addItem('週ぜんぶ水色にもどす', 'clearAllPositionColors')
      .addItem('名前と色を変える', 'openPositionSettings')
      .addSeparator()
      .addItem('ボタンを画面に置く', 'placePositionButtons')
      .addItem('ボタンを消す', 'clearPositionButtons'))
    .addItem('お店のはじめ方（新しい店舗）', 'openStoreSetup')
    .addSubMenu(ui.createMenu('LINEのつなぎこみ')
      .addItem('LINE通知の設定（はじめて入れる・入れ替える）', 'openLineSettings')
      .addItem('LINEログインの設定（はじめて入れる・入れ替える）', 'openLiffSettings'))
    .addSubMenu(ui.createMenu('困ったとき・やり直し')
      .addItem('色を塗り直す（臨時用）', 'refreshWeekColors')
      .addItem('出勤履歴を作り直す（臨時用）', 'rebuildAttendanceHistory')
      .addSeparator()
      .addItem('自動でシート作成・お知らせを再開する', 'startShiftAutomation')
      .addItem('自動を止める（臨時用）', 'stopShiftAutomation'))
    .addToUi();
}

// ===================== 店長・社員・キャストの時間数 =====================

// 「スタッフ区分」シートを取得（無ければ非表示で作る）
// 前の版は、D列に「数式の自動チェック」の結果を書いていた（今はJ列に書く）。
// D列は「高校生」の印に使うので、古い書き込みが残っていたら見出しごと直す
function fixOldRoleHeader_(sheet) {
  if (String(sheet.getRange('D1').getDisplayValue()).trim() !== '数式の自動チェック') return;
  var old = sheet.getRange('D2').getDisplayValue();
  sheet.getRange('D1:D2').clearContent();
  sheet.getRange(1, ROLE_BIRTH_COL, 1, 2).setValues([['生年月日', '高校生']]);
  if (old) sheet.getRange('J1:J2').setValues([['数式の自動チェック'], [old]]);   // 点検の結果はJ列へ移す
}

function getRoleSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(ROLE_SHEET_NAME);
  if (sheet) fixOldRoleHeader_(sheet);
  if (!sheet) {
    var active = ss.getActiveSheet(); // シート作成で表示中のシートが切り替わるため先に覚えておく
    sheet = ss.insertSheet(ROLE_SHEET_NAME);
    sheet.getRange('A1:B1').setValues([['名前', '区分']]);
    ss.setActiveSheet(active);
    sheet.hideSheet();
  }
  return sheet;
}

// 登録済みの区分を { 名前: '店長' or '社員' } の形で取得
function getStaffRoles_() {
  var sheet = getRoleSheet_();
  var map = {};
  var last = sheet.getLastRow();
  if (last < 2) return map;
  sheet.getRange(2, 1, last - 1, 2).getValues().forEach(function(r) {
    if (r[0]) map[r[0]] = r[1];
  });
  return map;
}

function getRosterNames_(sheet) {
  return sheet.getRange(ROSTER_START_ROW, 1, ROSTER_ROWS, 1).getValues()
    .map(function(r) { return r[0]; })
    .filter(function(n) { return n; });
}

function escapeHtml_(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function roleOption_(value, label, current) {
  return '<option value="' + value + '"' + (value === current ? ' selected' : '') + '>' + label + '</option>';
}

// A列の名簿を一覧表示し、1人ずつ「キャスト/社員/店長」をプルダウンで選ぶ画面を開く
function setStaffRoles() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  var names = getRosterNames_(sheet);
  if (names.length === 0) {
    ui.alert('登録されているスタッフがいません。');
    return;
  }

  var roles = getStaffRoles_();
  var rowsHtml = names.map(function(name, i) {
    var current = roles[name] || '';
    return '<div style="display:flex;justify-content:space-between;align-items:center;padding:6px 0;border-bottom:1px solid #eee;">' +
      '<span>' + escapeHtml_(name) + '</span>' +
      '<select data-i="' + i + '" style="padding:4px;font-size:14px;">' +
        roleOption_('', 'キャスト', current) +
        roleOption_(ROLE_EMPLOYEE, ROLE_EMPLOYEE, current) +
        roleOption_(ROLE_MANAGER, ROLE_MANAGER, current) +
      '</select></div>';
  }).join('');

  var html = '<div style="font-family:sans-serif;padding:8px;font-size:14px;">' +
    '<p style="margin-top:0;">店長・社員の人を選んでください。それ以外はキャストとして計算します。</p>' +
    '<div style="max-height:380px;overflow-y:auto;">' + rowsHtml + '</div>' +
    '<div style="margin-top:16px;text-align:right;">' +
    '<button onclick="google.script.host.close()" style="margin-right:8px;padding:6px 12px;">キャンセル</button>' +
    '<button id="saveBtn" onclick="doSave()" style="padding:6px 12px;">保存する</button>' +
    '</div>' +
    '</div>' +
    '<script>' +
    'function doSave(){' +
    '  document.getElementById("saveBtn").disabled = true;' +
    '  var selects = document.querySelectorAll("select[data-i]");' +
    '  var values = Array.prototype.map.call(selects, function(s){ return s.value; });' +
    '  google.script.run' +
    '    .withSuccessHandler(function(){ google.script.host.close(); })' +
    '    .withFailureHandler(function(e){ alert(e.message); document.getElementById("saveBtn").disabled = false; })' +
    '    .saveStaffRoles(values);' +
    '}' +
    '</script>';

  var output = HtmlService.createHtmlOutput(html).setWidth(360).setHeight(520);
  ui.showModalDialog(output, '店長・社員を設定');
}

// setStaffRoles の画面から呼ばれる保存処理
// ★末尾に _ を付けると google.script.run から呼べなくなるため付けない
function saveStaffRoles(roleList) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var names = getRosterNames_(sheet);
  if (names.length !== roleList.length) {
    throw new Error('名簿が変わったため保存できませんでした。画面を開き直してください。');
  }

  var rows = [];
  names.forEach(function(name, i) {
    if (roleList[i] === ROLE_MANAGER || roleList[i] === ROLE_EMPLOYEE) rows.push([name, roleList[i]]);
  });

  var roleSheet = getRoleSheet_();
  var last = roleSheet.getLastRow();
  if (last >= 2) roleSheet.getRange(2, 1, last - 1, 2).clearContent();
  if (rows.length > 0) roleSheet.getRange(2, 1, rows.length, 2).setValues(rows);
}

// 明細の名前が、登録された区分(role)に一致する行の時間だけを合計する数式
function roleHoursFormula_(namesA1, hoursA1, registryNamesA1, registryRolesA1, role) {
  return '=SUM(IFERROR(FILTER(' + hoursA1 + ',ISNUMBER(MATCH(' + namesA1 + ',FILTER(' +
    registryNamesA1 + ',' + registryRolesA1 + '="' + role + '"),0))),0))';
}

// "T24:T39" → "$T$24:$T$39"
function absA1_(a1) {
  return a1.replace(/([A-Z]+)(\d+)/g, '$$$1$$$2');
}

function setFormulaChecked_(sheet, a1, formula) {
  var cell = sheet.getRange(a1);
  var merged = cell.getMergedRanges();
  if (merged.length > 0 && merged[0].getA1Notation().split(':')[0] !== a1) {
    throw new Error(sheet.getName() + ' の ' + a1 + ' は結合セルの途中のため数式を書けません');
  }
  cell.setFormula(formula);
}

// 1枚の週シートの「キャスト・店長・社員時間数」を名前で計算する数式に置き換える
function applyRoleHoursFormulas_(sheet) {
  var regNames = "'" + ROLE_SHEET_NAME + "'!$A$2:$A";
  var regRoles = "'" + ROLE_SHEET_NAME + "'!$B$2:$B";

  DAY_BLOCKS.forEach(function(b) {
    var names = absA1_(b.names);
    var hours = absA1_(b.hours);
    var managerCell = b.col + ROW_MANAGER_HOURS;
    var employeeCell = b.col + ROW_EMPLOYEE_HOURS;
    setFormulaChecked_(sheet, managerCell, roleHoursFormula_(names, hours, regNames, regRoles, ROLE_MANAGER));
    setFormulaChecked_(sheet, employeeCell, roleHoursFormula_(names, hours, regNames, regRoles, ROLE_EMPLOYEE));
    // キャスト = その日の全員の時間 − 店長 − 社員
    setFormulaChecked_(sheet, b.col + ROW_CAST_HOURS, '=SUM(' + hours + ')-' + managerCell + '-' + employeeCell);
  });

  // 週の累計（社員時間数だけ数式が入っていなかったため追加）
  setFormulaChecked_(sheet, 'CQ' + ROW_EMPLOYEE_HOURS, '=SUM(Y' + ROW_EMPLOYEE_HOURS + ':CP' + ROW_EMPLOYEE_HOURS + ')');
}

// 数式が正しく計算されるかを、一時シートでテストする（テスト後に一時シートは削除）
function selfTestRoleFormula_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var active = ss.getActiveSheet();
  var test = ss.insertSheet('_区分チェック');
  try {
    test.getRange('A1:B4').setValues([['テスト店長', 8], ['テスト社員', 6], ['テストキャスト', 4], ['', 0]]);
    test.getRange('D1:E2').setValues([['テスト店長', ROLE_MANAGER], ['テスト社員', ROLE_EMPLOYEE]]);
    var n = '$A$1:$A$4', h = '$B$1:$B$4', rn = '$D$1:$D$10', rr = '$E$1:$E$10';
    test.getRange('G1').setFormula(roleHoursFormula_(n, h, rn, rr, ROLE_MANAGER));
    test.getRange('G2').setFormula(roleHoursFormula_(n, h, rn, rr, ROLE_EMPLOYEE));
    test.getRange('G3').setFormula('=SUM(' + h + ')-G1-G2');
    test.getRange('G4').setFormula(roleHoursFormula_(n, h, rn, rr, '該当なし'));
    SpreadsheetApp.flush();
    var got = test.getRange('G1:G4').getValues().map(function(r) { return r[0]; });
    var ok = got[0] === 8 && got[1] === 6 && got[2] === 4 && got[3] === 0;
    return (ok ? 'OK' : 'NG ' + JSON.stringify(got)) + ' ' + Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd HH:mm');
  } finally {
    ss.deleteSheet(test);
    ss.setActiveSheet(active);
  }
}

// ★1回だけAppsスクリプトエディタから実行する。
//   261019以降の週シートすべてに、名前で時間数を計算する数式を入れる（261012以前には触らない）
function setupRoleHours() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var roleSheet = getRoleSheet_();
  var targets = ss.getSheets().filter(function(s) {
    var name = s.getName();
    return /^\d{6}$/.test(name) && name >= FIRST_ROLE_SHEET;
  });
  if (targets.length === 0) {
    throw new Error(FIRST_ROLE_SHEET + ' 以降の週シートが見つかりません');
  }

  targets.forEach(function(s) { applyRoleHoursFormulas_(s); });

  var result = selfTestRoleFormula_();
  // D列は「高校生」の印に使うので、点検の結果はJ列に書く
  roleSheet.getRange('J1:J2').setValues([['数式の自動チェック'], [result]]);
  Logger.log('対象シート: ' + targets.map(function(s) { return s.getName(); }).join(', ') + ' / チェック結果: ' + result);
}

// 「次週シート作成」で時間を消さない人の一覧を取得（未設定なら松山・奥田を初期値とする）
function getProtectedStaffNames_() {
  var raw = PropertiesService.getDocumentProperties().getProperty('protectedStaff');
  if (!raw) return ['松山 貴志', '奥田 恵'];
  try {
    var parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch (e) {
    return [];
  }
}

// チェックボックスで「時間を固定するスタッフ」を選ぶ画面を開く
function setProtectedStaff() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var ui = SpreadsheetApp.getUi();

  var names = sheet.getRange(ROSTER_START_ROW, 1, ROSTER_ROWS, 1).getValues();
  var protectedNames = getProtectedStaffNames_();

  var rowsHtml = '';
  var hasAny = false;
  for (var i = 0; i < names.length; i++) {
    var name = names[i][0];
    if (!name) continue;
    hasAny = true;
    var checked = protectedNames.indexOf(name) !== -1 ? 'checked' : '';
    rowsHtml += '<label style="display:block;padding:4px 0;">' +
      '<input type="checkbox" value="' + name + '" ' + checked + '> ' + name +
      '</label>';
  }

  if (!hasAny) {
    ui.alert('登録されているスタッフがいません。');
    return;
  }

  var html = '<div style="font-family:sans-serif;padding:8px;font-size:14px;">' +
    '<p>「次週シート作成」の時に時間を消さない人にチェックを付けてください</p>' +
    '<div style="max-height:300px;overflow-y:auto;">' + rowsHtml + '</div>' +
    '<div style="margin-top:16px;text-align:right;">' +
    '<button onclick="google.script.host.close()" style="margin-right:8px;padding:6px 12px;">キャンセル</button>' +
    '<button onclick="doSave()" style="padding:6px 12px;">保存する</button>' +
    '</div>' +
    '</div>' +
    '<script>' +
    'function doSave(){' +
    '  var boxes = document.querySelectorAll("input[type=checkbox]:checked");' +
    '  var names = Array.prototype.map.call(boxes, function(b){ return b.value; });' +
    '  google.script.run.withSuccessHandler(function(){ google.script.host.close(); }).saveProtectedStaff(names);' +
    '}' +
    '</script>';

  var output = HtmlService.createHtmlOutput(html).setWidth(320).setHeight(400);
  ui.showModalDialog(output, '時間を固定するスタッフ');
}

// setProtectedStaff の画面から呼ばれる保存処理
// ★末尾に _ を付けると google.script.run から呼べなくなるため付けない
function saveProtectedStaff(names) {
  PropertiesService.getDocumentProperties().setProperty('protectedStaff', JSON.stringify(names));
}

// 名簿(4〜25行目)の空いている行に新しいスタッフの名前を入力する
// 週シート以外（日別勤務時間・月次集計など）でも押せてしまうと、そのシートの4〜25行目を
// 名簿と勘違いして書き換えてしまうので、週シートかどうかを先に確かめる。MARKER_STAFF_WEEK_GUARD
function addStaffMember() {
  var sheet = activeWeekSheet_();
  if (!sheet) return;
  var ui = SpreadsheetApp.getUi();

  var names = sheet.getRange(ROSTER_START_ROW, 1, ROSTER_ROWS, 1).getValues();
  var emptyRow = -1;
  for (var i = 0; i < names.length; i++) {
    if (!names[i][0]) {
      emptyRow = ROSTER_START_ROW + i;
      break;
    }
  }

  if (emptyRow === -1) {
    ui.alert('名簿の枠がいっぱいです。先に「スタッフを削除」で退社した人の枠を空けてください。');
    return;
  }

  var resp = ui.prompt('スタッフを追加', '新しいスタッフの名前を入力してください', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  var name = resp.getResponseText().trim();
  if (!name) return;

  sheet.getRange(emptyRow, 1).setValue(name);
  ui.alert(name + ' さんを ' + emptyRow + '行目 に追加しました。');
}

// 名簿(4〜25行目)を一覧表示し、プルダウンで選んで削除する画面を開く
// 週シート以外でも押せてしまうと、そのシートの4〜25行目の文字が「スタッフの名前」として
// プルダウンに出てしまう（日別勤務時間には名前が並ぶので見分けがつかない）。先に確かめる。
// MARKER_STAFF_WEEK_GUARD
function removeStaffMember() {
  var sheet = activeWeekSheet_();
  if (!sheet) return;
  var ui = SpreadsheetApp.getUi();

  var names = sheet.getRange(ROSTER_START_ROW, 1, ROSTER_ROWS, 1).getValues();
  var list = [];
  for (var i = 0; i < names.length; i++) {
    if (names[i][0]) list.push({ row: ROSTER_START_ROW + i, name: names[i][0] });
  }

  if (list.length === 0) {
    ui.alert('登録されているスタッフがいません。');
    return;
  }

  var options = list.map(function(item) {
    return '<option value="' + item.row + '">' + item.name + '</option>';
  }).join('');

  var html = '<div style="font-family:sans-serif;padding:8px;font-size:14px;">' +
    '<p style="margin:0 0 6px;">削除する人を選んでください</p>' +
    '<p style="margin:0 0 10px;color:#555;font-size:12px;">今週から先の週ぜんぶの名簿（名前と時間）と、' +
    '区分・生年月日・時給の設定から消します。過ぎた週と、右の日別シフト表はそのまま残します</p>' +
    '<select id="targetRow" style="width:100%;padding:6px;font-size:14px;">' + options + '</select>' +
    '<div id="msg" style="margin-top:10px;font-size:12.5px;white-space:pre-wrap;"></div>' +
    '<div style="margin-top:12px;text-align:right;">' +
    '<button onclick="google.script.host.close()" style="margin-right:8px;padding:6px 12px;">閉じる</button>' +
    '<button id="go" onclick="doDelete()" style="padding:6px 12px;">削除する</button>' +
    '</div>' +
    '</div>' +
    '<script>' +
    'function doDelete(){' +
    '  var sel = document.getElementById("targetRow");' +
    '  var btn = document.getElementById("go");' +
    '  var msg = document.getElementById("msg");' +
    '  if (!sel.value) return;' +
    '  btn.disabled = true; msg.style.color = "#555"; msg.textContent = "消しています…";' +
    '  google.script.run' +
    '    .withSuccessHandler(function(r){' +
    '      btn.disabled = false;' +
    '      msg.style.color = "#1155CC";' +
    '      msg.textContent = r.text;' +
    '      if (sel.selectedIndex >= 0) sel.remove(sel.selectedIndex);' +
    '      if (!sel.options.length) btn.disabled = true;' +
    '    })' +
    '    .withFailureHandler(function(e){' +
    '      btn.disabled = false; msg.style.color = "#C00000";' +
    '      msg.textContent = "消せませんでした：" + ((e && e.message) ? e.message : e);' +
    '    })' +
    '    .deleteStaffRow(Number(sel.value));' +
    '}' +
    '</script>';

  var output = HtmlService.createHtmlOutput(html).setWidth(340).setHeight(280);
  ui.showModalDialog(output, 'スタッフを削除');
}

// 1枚の週シートの名簿から、その名前の行を消して下を1行ずつ上に詰める。消した数を返す
// 名前の突き合わせは flatName_（全角・半角の空白を取る）。下から見るので、詰めても番号がずれない
function removeRosterName_(week, flat) {
  var lastRosterRow = ROSTER_START_ROW + ROSTER_ROWS - 1;
  var names = week.getRange(ROSTER_START_ROW, 1, ROSTER_ROWS, 1).getValues();
  var removed = 0;
  for (var i = ROSTER_ROWS - 1; i >= 0; i--) {
    if (flatName_(names[i][0]) !== flat) continue;
    var row = ROSTER_START_ROW + i;
    if (row < lastRosterRow) {
      // copyTo は値だけでなく「時刻の表示形式」などの書式もまとめて移動するため、
      // getValues/setValues（値だけ）ではなくこちらを使う
      week.getRange(row + 1, 1, lastRosterRow - row, 16)
        .copyTo(week.getRange(row, 1, lastRosterRow - row, 16));
    }
    week.getRange(lastRosterRow, 1, 1, 16).clearContent();
    removed++;
  }
  return removed;
}

// スタッフ区分シート（A=名前 B=区分 C=生年月日 D=高校生 E=今の時給 F=前の時給 G=いつから）
// から、その人の行を消して下を1行ずつ上に詰める。
// ★行ごと削除（deleteRow）は使わない。J列に「数式の自動チェック」の控えが入っているため。
// MARKER_STAFF_DELETE_ALL
function removeRoleRow_(flat) {
  var sheet = getRoleSheet_();
  var last = sheet.getLastRow();
  if (last < 2) return 0;
  var names = sheet.getRange(2, 1, last - 1, 1).getValues();
  var removed = 0;
  for (var i = names.length - 1; i >= 0; i--) {
    if (flatName_(names[i][0]) !== flat) continue;
    var row = i + 2;
    if (row < last) {
      sheet.getRange(row + 1, 1, last - row, ROLE_WAGE_FROM_COL)
        .copyTo(sheet.getRange(row, 1, last - row, ROLE_WAGE_FROM_COL));
    }
    sheet.getRange(last, 1, 1, ROLE_WAGE_FROM_COL).clearContent();
    last--;                      // 1行詰めたので、いちばん下が1つ上がる
    removed++;
  }
  return removed;
}

// 「時間を固定するスタッフ」（この書類の覚えがき）から外す
function removeProtectedStaff_(flat) {
  var props = PropertiesService.getDocumentProperties();
  var raw = props.getProperty('protectedStaff');
  if (!raw) return 0;
  var list = [];
  try { list = JSON.parse(raw) || []; } catch (err) { return 0; }
  var kept = list.filter(function(n) { return flatName_(n) !== flat; });
  if (kept.length === list.length) return 0;
  props.setProperty('protectedStaff', JSON.stringify(kept));
  return list.length - kept.length;
}

// まだ週シートが無い週にフォームから届いていた希望（保留）を消す。
// 専用シートなので、行ごと消して詰めてよい（下から見るので番号がずれない）
function removePendingFor_(flat) {
  var pending = getPendingSheet_();
  var last = pending.getLastRow();
  if (last < PENDING_FIRST_ROW) return 0;
  var rows = pending.getRange(PENDING_FIRST_ROW, 1, last - PENDING_FIRST_ROW + 1, 2).getValues();
  var removed = 0;
  for (var i = rows.length - 1; i >= 0; i--) {
    if (flatName_(rows[i][1]) !== flat) continue;
    pending.deleteRow(PENDING_FIRST_ROW + i);
    removed++;
  }
  return removed;
}

// その週の右の日別シフト表（DAY_BLOCKS の名前の欄）に、その人が入っているか。
// 曜日ごとに7回読むと、週の枚数だけ時間がかかるので、同じ列はまとめて1回で読む
function weekHasNameInDays_(week, flat) {
  var byCol = {};
  DAY_BLOCKS.forEach(function(block) {
    var m = /^([A-Z]+)(\d+):[A-Z]+(\d+)$/.exec(block.names);
    if (!m) return;
    var col = m[1], top = Number(m[2]), bottom = Number(m[3]);
    if (!byCol[col]) byCol[col] = { top: top, bottom: bottom };
    byCol[col].top = Math.min(byCol[col].top, top);
    byCol[col].bottom = Math.max(byCol[col].bottom, bottom);
  });
  var cols = Object.keys(byCol);
  for (var i = 0; i < cols.length; i++) {
    var span = byCol[cols[i]];
    var vals = week.getRange(cols[i] + span.top + ':' + cols[i] + span.bottom).getValues();
    for (var j = 0; j < vals.length; j++) {
      if (flatName_(vals[j][0]) === flat) return true;
    }
  }
  return false;
}

// 週シートの名前（261116）を 'yyyy-MM-dd' に直す
function weekNameToYmd_(name) {
  return '20' + name.slice(0, 2) + '-' + name.slice(2, 4) + '-' + name.slice(4, 6);
}

// removeStaffMember のプルダウンから呼ばれる削除の実処理
// ★末尾に _ を付けると google.script.run から呼べなくなるため付けない
// 辞めた人の名前が、シフトを作るときに出てこないようにするのが目的なので、
// 「今日が入る週から先ぜんぶ」の名簿（名前と時間）と、スタッフ区分・生年月日・時給・
// 時間を固定するスタッフ・まだ入れていない週の希望を、まとめて消す。
// 過ぎた週は実績なので触らない。右の日別シフト表は、その後の調整で動かすので触らない（残りは知らせる）。
// MARKER_STAFF_DELETE_ALL
function deleteStaffRow(row) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var week = ss.getActiveSheet();
  // 6けたの週シート（261116）だけを相手にする。古い4けたの名前の週は、日付の大小を
  // くらべられないので触らない（実績として残す）
  if (!/^\d{6}$/.test(week.getName())) {
    throw new Error('週シート（261116 のような名前）を開いてから押してください。');
  }
  var name = String(week.getRange(row, 1).getValue()).trim();
  if (!name) throw new Error('その行には名前が入っていません。いちど閉じて、開き直してください。');
  var flat = flatName_(name);
  var today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  var monday = addDaysYmd_(today, -((dowOfYmd_(today) + 6) % 7));   // 今日が入る週の月曜
  var done = [], left = [];
  ss.getSheets().forEach(function(s) {
    var n = s.getName();
    if (!/^\d{6}$/.test(n) || weekNameToYmd_(n) < monday) return;   // 過ぎた週は実績として残す
    if (removeRosterName_(s, flat)) done.push(n);
    if (weekHasNameInDays_(s, flat)) left.push(n);
  });
  var role = removeRoleRow_(flat);
  var fixed = removeProtectedStaff_(flat);
  var pending = removePendingFor_(flat);
  // 名簿から消えた人が、月次集計や日別勤務時間に残らないように作り直す
  try {
    if (ss.getSheetByName(MONTHLY_SHEET_NAME)) refreshMonthlySummary_();
    if (ss.getSheetByName(DAILY_SHEET_NAME)) refreshDailyHours_();
  } catch (err) {
    Logger.log('スタッフ削除のあとの作り直しに失敗: ' + err.message);
  }
  var lines = [name + ' さんを消しました'];
  lines.push('・名簿（名前と時間）：' + (done.length ? done.length + '枚（' + done.join('・') + '）' : 'なし'));
  lines.push('・区分・生年月日・時給：' + (role ? '消しました' : 'もともと入っていません'));
  if (fixed) lines.push('・時間を固定するスタッフ：外しました');
  if (pending) lines.push('・まだ入れていない週の希望：' + pending + '件');
  if (left.length) {
    lines.push('※ 右の日別シフト表にはまだ名前が入っています：' + left.join('・') + '（そのままで大丈夫です）');
  }
  return { name: name, done: done, left: left, role: role, fixed: fixed, pending: pending,
    from: monday, text: lines.join('\n') };
}

// ===================== スマホのシフト希望フォームとの連携 =====================
// フォーム（別のGAS）もこの「シフト希望設定」シートを同じ並びで読み書きする。並びを変える時は両方を直すこと
var REQUEST_SETTINGS_SHEET = 'シフト希望設定';
var REQUEST_PENDING_COLS = 19;       // シート名・名前・時間14・備考・送信日時・端末ID
/* 保留（まだ週シートが無い週）の希望は、専用シート「保留の希望」に置く。
   前は「シフト希望設定」シートの9行目以降に置いていた。しかしそのシートは
   D3:D49＝設定の説明／E3:E49＝設定の値 で、保留のデータと行が重なる。
   いまは保留が設定の下（50行目以降）に積まれているので壊れていないが、
   設定をもう1つ足すと 50行目＝保留の1行目 と重なって壊れる（時間幅やページURLがずれる）。
   専用シートに分ければ、行を消しても詰めても安全。MARKER_PENDING_OWN_SHEET          */
var PENDING_SHEET_NAME = '保留の希望';
var PENDING_FIRST_ROW = 2;           // 1行目は見出し
var OLD_PENDING_HEADER_ROW = 8;      // 前の見出し行（D8・E8は「書庫のURL」の設定なので触らない）
var OLD_PENDING_FIRST_ROW = 9;       // 前の置き場所。残っている分を移すときだけ見る
var REQUEST_INPUT_COLS = 15;         // 週シートのB〜P列
/* ----- お店ごとに変わる2つ。設定シートに入っていればそれを使う -----
   新しいお店では、メニュー「お店のはじめ方」から入れてもらう。
   設定が空のときは、下の初期値（飯塚店）のまま動く＝今までの店はそのまま。   */
// 下の初期値（REQUEST_FORM_URL_DEFAULT・RECORD_SPREADSHEET_ID_DEFAULT）の持ち主のファイル。
// コピーして作った別のお店では、この初期値は「よそのお店の値」なので、必ず入れ替えてもらう
var HOME_SHIFT_SPREADSHEET_ID = '1WXZ-hOSbPC4h7Atg75hnGp2Z3KLOggS6vTbE0SAj850';
var STORE_FORM_URL_CELL = 'E46';      // このお店のシフト希望ページのアドレス
var STORE_RECORD_ID_CELL = 'E47';     // このお店のアクセス記録ファイルのID
var REQUEST_FORM_URL_DEFAULT = 'https://script.google.com/macros/s/AKfycbwFtgsWnq75UnttAODig9q3c89E1Iek0Oow-lV8tHzF_Zp5qBQe2DXPahAeksqkd4i2/exec';
var STORE_SETUP_CACHE = {};           // 1回の実行の中では読み直さない

function storeSetting_(cell, pattern, fallback) {
  if (STORE_SETUP_CACHE[cell] === undefined) {
    var text = '';
    try {
      text = String(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REQUEST_SETTINGS_SHEET)
        .getRange(cell).getDisplayValue()).trim();
    } catch (err) { text = ''; }
    STORE_SETUP_CACHE[cell] = pattern.test(text) ? text : fallback;
  }
  return STORE_SETUP_CACHE[cell];
}

// このお店のシフト希望ページのアドレス
function requestFormUrl_() {
  return storeSetting_(STORE_FORM_URL_CELL,
    /^https:\/\/script\.google\.com\/macros\/s\/[A-Za-z0-9_-]+\/exec$/, REQUEST_FORM_URL_DEFAULT);
}

// このお店のアクセス記録ファイルのID
function recordSpreadsheetId_() {
  return storeSetting_(STORE_RECORD_ID_CELL, /^[A-Za-z0-9_-]{30,}$/, RECORD_SPREADSHEET_ID_DEFAULT);
}
var DOW_JP = ['日', '月', '火', '水', '木', '金', '土'];

// 募集期間の決まり。B3:B5＝今回だけ手で決めた期間（空なら自動）、E3:E5＝基本設定（店舗ごとの決まり）
var REQUEST_BASE_CLOSING_CELL = 'E3';   // 締め日（毎月この日。社員の締め日）
var REQUEST_BASE_WEEKS_CELL = 'E4';     // スタッフ入力締切＝募集が始まる何週間前の日曜日
var REQUEST_BASE_TIME_CELL = 'E5';      // 締切の時刻
var REQUEST_STORE_NAME_CELL = 'E6';     // 店舗名（見出しなどに使う）
var REQUEST_CAST_CLOSING_CELL = 'E7';   // キャストの締め日（月次集計に使う）
// スタッフがフォームで選べる時間の幅。店舗ごとに営業時間が違うので設定で変えられる
var REQUEST_START_FROM_CELL = 'E42';    // 出勤で選べる いちばん早い時間
var REQUEST_START_TO_CELL = 'E43';      // 出勤で選べる いちばん遅い時間
var REQUEST_END_FROM_CELL = 'E44';      // 退勤で選べる いちばん早い時間
var REQUEST_END_TO_CELL = 'E45';        // 退勤で選べる いちばん遅い時間
var REQUEST_DEFAULT_BASE = { closingDay: 15, deadlineWeeks: 2, deadlineTime: '23:59', castClosingDay: 20,
  startFrom: '9:00', startTo: '19:30', endFrom: '12:00', endTo: '21:30' };

// 保留（まだ週シートが無い週）の希望シートを取得（無ければ非表示で作る）。
// 取るたびに、前の置き場所に残っている分を引っ越す（1回で終わる・自動）。MARKER_PENDING_OWN_SHEET
function getPendingSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(PENDING_SHEET_NAME);
  if (!sheet) {
    var active = ss.getActiveSheet();   // シート作成で表示中のシートが変わるので先に覚えておく
    sheet = ss.insertSheet(PENDING_SHEET_NAME);
    ss.setActiveSheet(active);
    var headers = ['シート名', '名前'];
    ['月', '火', '水', '木', '金', '土', '日'].forEach(function(d) { headers.push(d + ' 出勤', d + ' 退勤'); });
    headers.push('備考', '送信日時', '端末ID');
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#EFF0ED');
    sheet.setFrozenRows(1);
    sheet.getRange('A:A').setNumberFormat('@');     // シート名（261116）が日付や数字に化けないように
    sheet.getRange('S:S').setNumberFormat('@');     // 端末ID
    sheet.getRange('A1').setNote('「次週シート作成」でその週のシートができると、ここから週シートへ入って行が消えます');
    sheet.hideSheet();
  }
  movePendingFromSettings_(ss, sheet);
  return sheet;
}

// 前の置き場所（シフト希望設定シートの9行目以降）に残っている保留を、専用シートへ移す。
// A列が6けたの週シート名で、名前が入っている行だけが保留のデータ（設定の行はA列が空）
function movePendingFromSettings_(ss, pendingSheet) {
  var settings = ss.getSheetByName(REQUEST_SETTINGS_SHEET);
  if (!settings) return 0;
  var last = settings.getLastRow();
  if (last < OLD_PENDING_FIRST_ROW) return 0;
  var n = last - OLD_PENDING_FIRST_ROW + 1;
  // まずA・B列だけ見て、前の置き場所に残っているかを調べる（ふだんは0件なのでここで終わる）
  var head = settings.getRange(OLD_PENDING_FIRST_ROW, 1, n, 2).getValues();
  var at = [];
  for (var i = 0; i < n; i++) {
    if (/^\d{6}$/.test(String(head[i][0]).trim()) && String(head[i][1]).trim()) at.push(i);
  }
  if (!at.length) return 0;
  var rows = settings.getRange(OLD_PENDING_FIRST_ROW, 1, n, REQUEST_PENDING_COLS).getValues();
  pendingSheet.getRange(pendingSheet.getLastRow() + 1, 1, at.length, REQUEST_PENDING_COLS)
    .setValues(at.map(function(i) { return rows[i]; }));
  at.forEach(function(i) {
    settings.getRange(OLD_PENDING_FIRST_ROW + i, 1, 1, REQUEST_PENDING_COLS).clearContent();   // 行は消さない（設定が上にある）
  });
  Logger.log('保留の希望を ' + at.length + '件、専用シートへ移しました');
  return at.length;
}

// 「シフト希望設定」シートを取得（無ければ非表示で作る）
function getRequestSettingsSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(REQUEST_SETTINGS_SHEET);
  if (!sheet) {
    var active = ss.getActiveSheet(); // シート作成で表示中のシートが切り替わるため先に覚えておく
    sheet = ss.insertSheet(REQUEST_SETTINGS_SHEET);
    // 日付やシート名が勝手に日付・数字に変わらないよう、文字として扱う
    sheet.getRange('A:A').setNumberFormat('@');
    sheet.getRange('B3:B5').setNumberFormat('@');
    sheet.getRange('R:R').setNumberFormat('@');
    sheet.getRange('A1').setValue('シフト希望の募集設定（メニュー「スタッフ管理」→「シフト希望の募集設定」で変更します）');
    ss.setActiveSheet(active);
    sheet.hideSheet();
  }
  // 基本設定の欄。空の項目だけ初期値を入れる（前の版で作ったシートにも足す）
  sheet.getRange('E3:E28').setNumberFormat('@');
  sheet.getRange('E42:E47').setNumberFormat('@');   // 「9:00」が時刻に化けないように文字で持つ
  sheet.getRange('F22:F25').setNumberFormat('@');
  sheet.getRange('E48:F49').setNumberFormat('@');   // ポジション5と通知先（E26が埋まっているので離れた行）
  [[REQUEST_BASE_CLOSING_CELL, String(REQUEST_DEFAULT_BASE.closingDay)],
   [REQUEST_BASE_WEEKS_CELL, String(REQUEST_DEFAULT_BASE.deadlineWeeks)],
   [REQUEST_BASE_TIME_CELL, REQUEST_DEFAULT_BASE.deadlineTime],
   [REQUEST_CAST_CLOSING_CELL, String(REQUEST_DEFAULT_BASE.castClosingDay)],
   [MONTHLY_ON_CELL, '切'],
   [MONTHLY_DAY_CELL, String(MONTHLY_DEFAULT_DAY)],
   [MONTHLY_DAYS_CELL, MONTHLY_DEFAULT_DAYS.join(',')],
   [MONTHLY_TEXT_CELL, MONTHLY_DEFAULT_TEXT],
   [MONTHLY_HOUR_CELL, String(MONTHLY_DEFAULT_HOUR)],
   [REQUEST_START_FROM_CELL, REQUEST_DEFAULT_BASE.startFrom],
   [REQUEST_START_TO_CELL, REQUEST_DEFAULT_BASE.startTo],
   [REQUEST_END_FROM_CELL, REQUEST_DEFAULT_BASE.endFrom],
   [REQUEST_END_TO_CELL, REQUEST_DEFAULT_BASE.endTo]].forEach(function(pair) {
    if (String(sheet.getRange(pair[0]).getDisplayValue()).trim() === '') sheet.getRange(pair[0]).setValue(pair[1]);
  });
  // スタッフ入力締切の決め方を「期間の最終日の◯日前」から「募集が始まる◯週間前の日曜日」に変えたため、
  // 前の版で入っていた1〜4以外の数字（14など）は初期値に直す
  var weeksNow = Number(sheet.getRange(REQUEST_BASE_WEEKS_CELL).getDisplayValue());
  if (!(weeksNow >= 1 && weeksNow <= 4)) sheet.getRange(REQUEST_BASE_WEEKS_CELL).setValue(String(REQUEST_DEFAULT_BASE.deadlineWeeks));
  sheet.getRange('D3:D49').setValues([
    ['基本設定：社員の締め日（毎月この日）'],
    ['基本設定：スタッフ入力締切は募集が始まる何週間前の日曜日'],
    ['基本設定：スタッフ入力締切の時刻'],
    ['基本設定：店舗名'],
    ['基本設定：キャストの締め日（毎月この日）'],
    ['基本設定：書庫ファイルのURL（古い週の移動先。年が変わると自動で新しい書庫に切り替わります）'],
    ['LINE：チャネルアクセストークン（メニュー「LINE通知の設定」で入れます）'],
    ['LINE：お知らせを送る先のID（グループに招待すると自動で入ります）'],
    ['お知らせを送った控え（自動。消しても害はありません）'],
    ['LINE：送り先の登録を受け付ける期限（自動。10分たつと切れます）'],
    ['お知らせ：送る日（スタッフ入力締切の何日前か）'],
    ['お知らせ：送る時刻'],
    ['お知らせ：1回目の文'],
    ['お知らせ：2回目以降の文'],
    ['毎月のおしらせ：送る（入）／送らない（切）'],
    ['毎月のおしらせ：期日（毎月この日）'],
    ['毎月のおしらせ：送る日（期日の何日前か）'],
    ['毎月のおしらせ：文'],
    ['毎月のおしらせ：送る時刻'],
    ['ポジション1：名前（E列）と色（F列）'],
    ['ポジション2：名前（E列）と色（F列）'],
    ['ポジション3：名前（E列）と色（F列）'],
    ['ポジション4：名前（E列）と色（F列）'],
    ['確定してスタッフに公開した週（自動。この週だけがスタッフのページとカレンダーに出ます）'],
    ['シフトPDFの入れ物フォルダ（自動）'],
    ['アルバイト名簿ファイルのURL（生年月日を読みます。空ならスタッフ区分シートの入力を使います）'],
    ['1日に欲しい人数の目安（平日）。お知らせの{人手不足}で使います'],
    ['1日に欲しい人数の目安（土曜・日曜・祝日）。お知らせの{人手不足}で使います'],
    ['お知らせ：締切当日の文（送る日に0を入れると送ります）'],
    ['LINE：お試しの送り先（メニュー「お試し公開」だけがここへ送ります）'],
    ['交通費の出し方（メニュー「時給と交通費の設定」で選びます）'],
    ['交通費：何日以上はたらいたら出すか'],
    ['交通費：金額（円）'],
    ['契約が更新される月（例「3,9」。この日に時給が変わります）'],
    ['契約が更新される日（例「21」）'],
    ['毎月のおしらせ：提出チェック（人数／出さない／使わない）'],
    ['LINEログイン：LIFF ID（いまは使いません）'],
    ['LINEログイン：チャネルID（メニュー「LINEログインの設定」で入れます）'],
    ['LINEログイン：チャネルシークレット（同上）'],
    ['スタッフが選べる出勤の時間：いちばん早い'],
    ['スタッフが選べる出勤の時間：いちばん遅い'],
    ['スタッフが選べる退勤の時間：いちばん早い'],
    ['スタッフが選べる退勤の時間：いちばん遅い'],
    ['このお店のシフト希望ページのアドレス（メニュー「お店のはじめ方」で入れます）'],
    ['このお店のアクセス記録ファイルのID（同上）'],
    ['ポジション5：名前（E列）と色（F列）'],
    ['シフト希望が届いたときの通知先（メニュー「LINE通知の設定」で登録。空なら送りません）']
  ]);
  // 欲しい人数の目安。空のときだけ初期値を入れる
  if (String(sheet.getRange(NEED_WEEKDAY_CELL).getDisplayValue()).trim() === '') {
    sheet.getRange(NEED_WEEKDAY_CELL).setValue(NEED_DEFAULT.weekday);
  }
  if (String(sheet.getRange(NEED_WEEKEND_CELL).getDisplayValue()).trim() === '') {
    sheet.getRange(NEED_WEEKEND_CELL).setValue(NEED_DEFAULT.weekend);
  }
  // ポジションの名前と色。空のところだけ初期値を入れる
  POSITION_DEFAULTS.forEach(function(pos, i) {
    var nameCell = sheet.getRange(POSITION_ROWS[i], 5);
    var colorCell = sheet.getRange(POSITION_ROWS[i], 6);
    if (String(nameCell.getDisplayValue()).trim() === '') nameCell.setValue(pos.name);
    if (String(colorCell.getDisplayValue()).trim() === '') colorCell.setValue(pos.color);
  });
  sheet.getRange('A3:A5').setValues([['今回だけの開始週（空なら自動）'], ['今回だけの週数'], ['今回だけのスタッフ入力締切']]);
  // 前の版は、7行目に案内・8行目に保留の見出しを書いていた。保留は専用シートに移したので片づける。
  // ★8行目の D・E列は「書庫ファイルのURL」の設定なので触らない
  sheet.getRange('A7').setValue('※ まだ週シートが無い週の希望は「保留の希望」シート（ふだんは非表示）に入ります');
  sheet.getRange(OLD_PENDING_HEADER_ROW, 1, 1, 3).clearContent();                          // A〜C
  sheet.getRange(OLD_PENDING_HEADER_ROW, 6, 1, REQUEST_PENDING_COLS - 5).clearContent();   // F〜S
  return sheet;
}

// 新しく作った週シートに、フォームから先に送られていたその週の希望を入れ、保留から消す
function applyPendingRequests_(sheet) {
  var pending = getPendingSheet_();
  var last = pending.getLastRow();
  if (last < PENDING_FIRST_ROW) return;

  var rows = pending.getRange(PENDING_FIRST_ROW, 1, last - PENDING_FIRST_ROW + 1, REQUEST_PENDING_COLS).getValues();
  var roster = sheet.getRange(ROSTER_START_ROW, 1, ROSTER_ROWS, 1).getValues()
    .map(function(r) { return String(r[0]).trim(); });

  for (var i = rows.length - 1; i >= 0; i--) {
    var name = String(rows[i][1]).trim();
    if (String(rows[i][0]) !== sheet.getName() || !name) continue;
    var index = roster.indexOf(name);
    if (index === -1) continue; // 名簿に名前が無い希望は消さずに残す
    sheet.getRange(ROSTER_START_ROW + index, 2, 1, REQUEST_INPUT_COLS).setValues([rows[i].slice(2, 2 + REQUEST_INPUT_COLS)]);
    pending.deleteRow(PENDING_FIRST_ROW + i);   // 専用シートなので、行ごと消して詰めてよい
  }
}

// 'yyyy-MM-dd' に日数を足す（時間帯の影響を受けないように日付だけで計算する）
function addDaysYmd_(ymd, days) {
  var p = ymd.split('-').map(Number);
  var d = new Date(Date.UTC(p[0], p[1] - 1, p[2] + days));
  return d.getUTCFullYear() + '-' + ('0' + (d.getUTCMonth() + 1)).slice(-2) + '-' + ('0' + d.getUTCDate()).slice(-2);
}

function dowOfYmd_(ymd) {
  var p = ymd.split('-').map(Number);
  return new Date(Date.UTC(p[0], p[1] - 1, p[2])).getUTCDay();
}

// '2026-10-26' → '10/26(月)'
function ymdLabel_(ymd) {
  var p = ymd.split('-').map(Number);
  return p[1] + '/' + p[2] + '(' + DOW_JP[dowOfYmd_(ymd)] + ')';
}

// 基本設定（店舗ごとの決まり）を読む。空や変な値の欄は初期値
function readRequestBase_(sheet) {
  var closing = Number(sheet.getRange(REQUEST_BASE_CLOSING_CELL).getDisplayValue());
  var weeksText = String(sheet.getRange(REQUEST_BASE_WEEKS_CELL).getDisplayValue()).trim();
  var time = String(sheet.getRange(REQUEST_BASE_TIME_CELL).getDisplayValue()).trim();
  var cast = Number(sheet.getRange(REQUEST_CAST_CLOSING_CELL).getDisplayValue());
  return {
    closingDay: closing >= 1 && closing <= 31 ? Math.floor(closing) : REQUEST_DEFAULT_BASE.closingDay,
    deadlineWeeks: Number(weeksText) >= 1 && Number(weeksText) <= 4 ? Math.floor(Number(weeksText)) : REQUEST_DEFAULT_BASE.deadlineWeeks,
    deadlineTime: /^\d{2}:\d{2}$/.test(time) ? time : REQUEST_DEFAULT_BASE.deadlineTime,
    castClosingDay: cast >= 1 && cast <= 31 ? Math.floor(cast) : REQUEST_DEFAULT_BASE.castClosingDay,
    startFrom: stepTime_(sheet.getRange(REQUEST_START_FROM_CELL).getDisplayValue(), REQUEST_DEFAULT_BASE.startFrom),
    startTo: stepTime_(sheet.getRange(REQUEST_START_TO_CELL).getDisplayValue(), REQUEST_DEFAULT_BASE.startTo),
    endFrom: stepTime_(sheet.getRange(REQUEST_END_FROM_CELL).getDisplayValue(), REQUEST_DEFAULT_BASE.endFrom),
    endTo: stepTime_(sheet.getRange(REQUEST_END_TO_CELL).getDisplayValue(), REQUEST_DEFAULT_BASE.endTo)
  };
}

// 「9:00」「21:15」のような15分きざみの時間だけ受け取る。それ以外は初期値
function stepTime_(text, fallback) {
  var t = String(text == null ? '' : text).trim();
  var m = /^(\d{1,2}):(00|15|30|45)$/.exec(t);
  if (!m || Number(m[1]) > 23) return fallback;
  return Number(m[1]) + ':' + m[2];
}

// 0:00〜23:45 の15分ごとの選択肢（端の時間だけは15分・45分も決められる）
function stepTimeOptions_(selected) {
  var out = '';
  for (var h = 0; h < 24; h++) {
    ['00', '15', '30', '45'].forEach(function(mm) {
      var t = h + ':' + mm;
      out += '<option value="' + t + '"' + (t === selected ? ' selected' : '') + '>' + t + '</option>';
    });
  }
  return out;
}

// スタッフ入力締切の日＝募集が始まる◯週間前の日曜日（初日が月曜なので、初日の前日の日曜からさらに◯−1週さかのぼる）
function requestDeadlineOf_(startYmd, base) {
  return addDaysYmd_(startYmd, -(base.deadlineWeeks * 7 + 1));
}

// その月の締め日が入っている週の日曜日（締め日が日曜ならその日）。月の日数を超える締め日は末日にする
function requestPeriodEnd_(year, month, closingDay) {
  var lastDay = new Date(Date.UTC(year, month, 0)).getUTCDate();
  var d = new Date(Date.UTC(year, month - 1, Math.min(closingDay, lastDay)));
  d.setUTCDate(d.getUTCDate() + (7 - d.getUTCDay()) % 7);
  return addDaysYmd_(d.getUTCFullYear() + '-' + ('0' + (d.getUTCMonth() + 1)).slice(-2) + '-' + ('0' + d.getUTCDate()).slice(-2), 0);
}

function jstMillis_(text) {
  return Utilities.parseDate(text, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm').getTime();
}

function daysBetweenYmd_(fromYmd, toYmd) {
  var a = fromYmd.split('-').map(Number), b = toYmd.split('-').map(Number);
  return Math.round((Date.UTC(b[0], b[1] - 1, b[2]) - Date.UTC(a[0], a[1] - 1, a[2])) / 86400000);
}

// 自動の募集期間：前の期間の最終日の翌日（月曜）〜締め日が入っている週の日曜日。
// スタッフ入力締切＝募集が始まる◯週間前の日曜日。締切を過ぎた期間と、FIRST_ROLE_SHEET より前から始まる期間は飛ばす
function autoRequestPeriod_(base, now) {
  var ym = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM').split('-').map(Number);
  for (var i = -1; i <= 13; i++) {
    var end = requestPeriodEnd_(ym[0], ym[1] + i, base.closingDay);
    var start = addDaysYmd_(requestPeriodEnd_(ym[0], ym[1] + i - 1, base.closingDay), 1);
    if (start.slice(2).replace(/-/g, '') < FIRST_ROLE_SHEET) continue;
    var deadline = requestDeadlineOf_(start, base) + ' ' + base.deadlineTime;
    if (jstMillis_(deadline) <= now.getTime()) continue;
    return { start: start, weeks: (daysBetweenYmd_(start, end) + 1) / 7, deadline: deadline, manual: false };
  }
  return null;
}

// 今の募集期間。今回だけ手で決めた期間が締切前ならそれ、無ければ自動
// ※フォームのGAS（currentPeriod_）と同じ決まり。変える時は両方を直すこと
function currentRequestPeriod_(sheet, now) {
  var start = String(sheet.getRange('B3').getDisplayValue()).trim();
  var weeks = Number(sheet.getRange('B4').getDisplayValue());
  var deadline = String(sheet.getRange('B5').getDisplayValue()).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(start) && weeks >= 1 && weeks <= 5 &&
      /^\d{4}-\d{2}-\d{2} \d{2}:\d{2}$/.test(deadline) && jstMillis_(deadline) > now.getTime()) {
    return { start: start, weeks: weeks, deadline: deadline, manual: true };
  }
  return autoRequestPeriod_(readRequestBase_(sheet), now);
}

// '10/19(月)〜11/15(日)（4週）　締切 11/1(日) 23:59'
function requestPeriodText_(period) {
  if (!period) return '募集できる期間が見つかりません';
  return ymdLabel_(period.start) + '〜' + ymdLabel_(addDaysYmd_(period.start, period.weeks * 7 - 1)) +
    '（' + period.weeks + '週）　スタッフ入力締切 ' + ymdLabel_(period.deadline.slice(0, 10)) + ' ' + period.deadline.slice(11);
}

// 0:00〜23:30 の30分ごと＋23:59 の選択肢
function requestTimeOptions_(selected) {
  var times = [];
  for (var h = 0; h < 24; h++) times.push(('0' + h).slice(-2) + ':00', ('0' + h).slice(-2) + ':30');
  times.push('23:59');
  if (times.indexOf(selected) === -1) times.push(selected);
  return times.map(function(t) {
    return '<option value="' + t + '"' + (t === selected ? ' selected' : '') + '>' + t + '</option>';
  }).join('');
}

function numberOptions_(from, to, selected, format) {
  var out = '';
  for (var n = from; n <= to; n++) {
    out += '<option value="' + n + '"' + (n === selected ? ' selected' : '') + '>' + format.replace('#', n) + '</option>';
  }
  return out;
}

// メニュー「シフト希望の募集設定」：基本設定（締め日・締切）と、今回だけ期間を変える設定の画面を開く
function openShiftRequestSettings() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getRequestSettingsSheet_();
  var now = new Date();
  var base = readRequestBase_(sheet);
  var storeName = getStoreName_();
  var archiveUrl = String(sheet.getRange(ARCHIVE_URL_CELL).getDisplayValue()).trim();
  var current = currentRequestPeriod_(sheet, now);
  var auto = autoRequestPeriod_(base, now);
  var manualActive = !!(current && current.manual);
  var shown = manualActive ? current : auto;

  // 今回だけの期間の欄は、使っていなければ自動の期間を初期値にする
  var formStart = shown ? shown.start : '';
  var formWeeks = shown ? shown.weeks : 5;
  var formDeadlineDate = shown ? shown.deadline.slice(0, 10) : '';
  var formDeadlineTime = shown ? shown.deadline.slice(11) : base.deadlineTime;

  var today = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd');
  var thisMonday = addDaysYmd_(today, -((dowOfYmd_(today) + 6) % 7));
  var starts = [];
  for (var i = 0; i <= 20; i++) starts.push(addDaysYmd_(thisMonday, i * 7));
  if (formStart && starts.indexOf(formStart) === -1) starts.unshift(formStart);
  var startOptions = starts.filter(function(ymd) {
    return ymd.slice(2).replace(/-/g, '') >= FIRST_ROLE_SHEET;
  }).map(function(ymd) {
    var made = ss.getSheetByName(ymd.slice(2).replace(/-/g, '')) ? 'シート作成済み' : 'シート未作成';
    return '<option value="' + ymd + '"' + (ymd === formStart ? ' selected' : '') + '>' +
      ymdLabel_(ymd) + '〜' + ymdLabel_(addDaysYmd_(ymd, 6)) + '　' + made + '</option>';
  }).join('');

  // シートがまだ無いため保留になっている希望の件数
  var pendingCounts = {};
  var pendingSheet = getPendingSheet_();
  var last = pendingSheet.getLastRow();
  if (last >= PENDING_FIRST_ROW) {
    pendingSheet.getRange(PENDING_FIRST_ROW, 1, last - PENDING_FIRST_ROW + 1, 2).getValues().forEach(function(r) {
      if (r[0] && r[1]) pendingCounts[String(r[0])] = (pendingCounts[String(r[0])] || 0) + 1;
    });
  }
  var pendingText = Object.keys(pendingCounts).sort().map(function(n) { return n + '：' + pendingCounts[n] + '人'; }).join('、');

  // スタッフに送るリンクと、LINEにそのまま貼れる文（期間と締切が入る）。
  // リンクの末尾に期間を付けるのは、同じリンクだとLINEが前の見出しを覚えていて古い期間を出すため
  var linkUrl = staffPageUrl_(sheet, shown ? '?p=' + shown.start.slice(2).replace(/-/g, '') : '');
  var lineText = 'シフト希望の入力をお願いします。' +
    (shown ? '\n期間 ' + ymdLabel_(shown.start) + '〜' + ymdLabel_(addDaysYmd_(shown.start, shown.weeks * 7 - 1)) +
             '\nスタッフ入力締切 ' + ymdLabel_(shown.deadline.slice(0, 10)) + ' ' + shown.deadline.slice(11) : '') +
    '\n' + linkUrl;

  var label = 'display:block;font-weight:bold;margin-top:10px;';
  var select = 'width:100%;padding:6px;font-size:14px;';
  var html = '<div style="font-family:sans-serif;padding:4px 8px;font-size:14px;line-height:1.6;">' +
    '<div style="background:#FFF4E5;border-radius:6px;padding:8px 10px;">' +
    '<div style="font-size:12px;color:#555;">いまフォームで募集している期間（' + (manualActive ? '今回だけの設定' : '自動') + '）</div>' +
    '<div style="font-weight:bold;">' + escapeHtml_(requestPeriodText_(shown)) + '</div></div>' +

    '<div style="font-weight:bold;margin-top:14px;border-bottom:1px solid #ddd;">基本設定（店舗ごとの決まり）</div>' +
    '<label style="' + label + '">社員の締め日</label>' +
    '<select id="closingDay" style="' + select + '">' + numberOptions_(1, 31, base.closingDay, '毎月#日') + '</select>' +
    '<label style="' + label + '">スタッフ入力締切</label>' +
    '<select id="deadlineWeeks" style="' + select + '">' + numberOptions_(1, 4, base.deadlineWeeks, '募集が始まる#週間前の日曜日') + '</select>' +
    '<label style="' + label + '">スタッフ入力締切の時刻</label>' +
    '<select id="baseTime" style="' + select + '">' + requestTimeOptions_(base.deadlineTime) + '</select>' +
    '<label style="' + label + '">キャストの締め日（月次集計に使います）</label>' +
    '<select id="castClosingDay" style="' + select + '">' + numberOptions_(1, 31, base.castClosingDay, '毎月#日') + '</select>' +
    '<label style="' + label + '">スタッフが選べる時間（出勤）</label>' +
    '<div style="display:flex;gap:6px;align-items:center;">' +
    '<select id="startFrom" style="flex:1;padding:6px;font-size:14px;">' + stepTimeOptions_(base.startFrom) + '</select>' +
    '<span>〜</span>' +
    '<select id="startTo" style="flex:1;padding:6px;font-size:14px;">' + stepTimeOptions_(base.startTo) + '</select>' +
    '</div>' +
    '<label style="' + label + '">スタッフが選べる時間（退勤）</label>' +
    '<div style="display:flex;gap:6px;align-items:center;">' +
    '<select id="endFrom" style="flex:1;padding:6px;font-size:14px;">' + stepTimeOptions_(base.endFrom) + '</select>' +
    '<span>〜</span>' +
    '<select id="endTo" style="flex:1;padding:6px;font-size:14px;">' + stepTimeOptions_(base.endTo) + '</select>' +
    '</div>' +
    '<p style="margin:6px 0 0;font-size:12px;color:#555;">スマホのフォームで選べる時間の幅です。ここは15分きざみで決められます。' +
    'スタッフが選ぶのは30分きざみですが、<b>いちばん早い／遅い時間だけは、ここで決めた時間そのもの</b>も選べます' +
    '（例：退勤を21:15にすると、…20:30／21:00／21:15 から選べます）。</p>' +
    '<label style="' + label + '">店舗名</label>' +
    '<input id="storeName" value="' + escapeHtml_(storeName) + '" placeholder="例：ペッパーランチ飯塚店" style="' + select + '">' +
    '<label style="' + label + '">書庫ファイルのURL（' + ARCHIVE_KEEP_WEEKS + '週より前の週シートの移動先。年が変わると自動で新しい書庫を作って入れ替えます）</label>' +
    '<input id="archiveUrl" value="' + escapeHtml_(archiveUrl) + '" placeholder="空のスプレッドシートのURLを貼る" style="' + select + '">' +
    '<p style="margin:6px 0 0;font-size:12px;color:#555;">期間は「社員の締め日が入っている週の日曜日」で区切ります。スタッフ入力締切を過ぎると、フォームは次の期間に切り替わります。店舗名は月次集計の見出しとスマホのフォームに出ます。</p>' +

    '<label style="display:block;margin-top:14px;font-weight:bold;border-bottom:1px solid #ddd;">' +
    '<input type="checkbox" id="manual"' + (manualActive ? ' checked' : '') + ' onchange="toggleManual()"> 今回だけ期間を変える</label>' +
    '<div id="manualBox" style="' + (manualActive ? '' : 'opacity:.4;pointer-events:none;') + '">' +
    '<label style="' + label + '">開始週</label>' +
    '<select id="start" style="' + select + '">' + startOptions + '</select>' +
    '<label style="' + label + '">週数</label>' +
    '<select id="weeks" style="' + select + '">' + numberOptions_(1, 5, formWeeks, '#週') + '</select>' +
    '<label style="' + label + '">スタッフ入力締切</label>' +
    '<div style="display:flex;gap:6px;">' +
    '<input type="date" id="deadlineDate" value="' + formDeadlineDate + '" style="flex:2;padding:5px;font-size:14px;">' +
    '<select id="deadlineTime" style="flex:1;padding:6px;font-size:14px;">' + requestTimeOptions_(formDeadlineTime) + '</select>' +
    '</div>' +
    '<p style="margin:6px 0 0;font-size:12px;color:#555;">このスタッフ入力締切を過ぎると、自動の期間に戻ります。</p>' +
    '</div>' +

    '<p style="margin:14px 0 4px;font-size:12px;color:#555;">LINEに貼る文（タップで全部選ばれます）</p>' +
    '<textarea readonly onclick="this.select()" style="width:100%;height:84px;padding:5px;font-size:12px;">' + escapeHtml_(lineText) + '</textarea>' +
    '<p style="margin:10px 0 4px;font-size:12px;color:#555;">リンクだけ</p>' +
    '<input readonly value="' + linkUrl + '" onclick="this.select()" style="width:100%;padding:5px;font-size:12px;">' +
    '<p style="margin:10px 0 0;font-size:12px;color:#555;">シート作成待ちの希望：' + (pendingText ? escapeHtml_(pendingText) : 'なし') + '</p>' +
    '<div style="margin-top:14px;text-align:right;">' +
    '<button onclick="google.script.host.close()" style="margin-right:8px;padding:6px 12px;">キャンセル</button>' +
    '<button id="saveBtn" onclick="doSave()" style="padding:6px 12px;">保存する</button>' +
    '</div>' +
    '</div>' +
    '<script>' +
    'function toggleManual(){' +
    '  var on = document.getElementById("manual").checked;' +
    '  var box = document.getElementById("manualBox");' +
    '  box.style.opacity = on ? "1" : ".4";' +
    '  box.style.pointerEvents = on ? "auto" : "none";' +
    '}' +
    'function doSave(){' +
    '  document.getElementById("saveBtn").disabled = true;' +
    '  google.script.run' +
    '    .withSuccessHandler(function(){ google.script.host.close(); })' +
    '    .withFailureHandler(function(e){ alert(e.message); document.getElementById("saveBtn").disabled = false; })' +
    '    .saveShiftRequestSettings({' +
    '      closingDay: document.getElementById("closingDay").value,' +
    '      deadlineWeeks: document.getElementById("deadlineWeeks").value,' +
    '      baseTime: document.getElementById("baseTime").value,' +
    '      castClosingDay: document.getElementById("castClosingDay").value,' +
    '      startFrom: document.getElementById("startFrom").value,' +
    '      startTo: document.getElementById("startTo").value,' +
    '      endFrom: document.getElementById("endFrom").value,' +
    '      endTo: document.getElementById("endTo").value,' +
    '      storeName: document.getElementById("storeName").value,' +
    '      archiveUrl: document.getElementById("archiveUrl").value,' +
    '      manual: document.getElementById("manual").checked,' +
    '      start: document.getElementById("start").value,' +
    '      weeks: document.getElementById("weeks").value,' +
    '      deadlineDate: document.getElementById("deadlineDate").value,' +
    '      deadlineTime: document.getElementById("deadlineTime").value' +
    '    });' +
    '}' +
    '</script>';

  var output = HtmlService.createHtmlOutput(html).setWidth(400).setHeight(640);
  SpreadsheetApp.getUi().showModalDialog(output, 'シフト希望の募集設定');
}

// openShiftRequestSettings の画面から呼ばれる保存処理
// ★末尾に _ を付けると google.script.run から呼べなくなるため付けない
function saveShiftRequestSettings(settings) {
  settings = settings || {};
  var closingDay = Number(settings.closingDay);
  var deadlineWeeks = Number(settings.deadlineWeeks);
  var baseTime = String(settings.baseTime || '');
  var castClosingDay = Number(settings.castClosingDay);
  var storeName = String(settings.storeName == null ? '' : settings.storeName).trim().slice(0, 40);
  if (!(closingDay >= 1 && closingDay <= 31 && Math.floor(closingDay) === closingDay)) throw new Error('社員の締め日は1〜31日で選んでください。');
  if (!(castClosingDay >= 1 && castClosingDay <= 31 && Math.floor(castClosingDay) === castClosingDay)) throw new Error('キャストの締め日は1〜31日で選んでください。');
  if (!(deadlineWeeks >= 1 && deadlineWeeks <= 4 && Math.floor(deadlineWeeks) === deadlineWeeks)) throw new Error('スタッフ入力締切は1〜4週間前で選んでください。');
  if (!/^\d{2}:\d{2}$/.test(baseTime)) throw new Error('スタッフ入力締切の時刻を選んでください。');
  // 画面から届かなかった時は、今シートに入っている値をそのまま使う
  var hoursNow = readRequestBase_(getRequestSettingsSheet_());
  var hours = {};
  ['startFrom', 'startTo', 'endFrom', 'endTo'].forEach(function(key) {
    hours[key] = stepTime_(settings[key], hoursNow[key]);
  });
  if (stepMinutes_(hours.startTo) <= stepMinutes_(hours.startFrom)) throw new Error('出勤の時間は「早い→遅い」の順で選んでください。');
  if (stepMinutes_(hours.endTo) <= stepMinutes_(hours.endFrom)) throw new Error('退勤の時間は「早い→遅い」の順で選んでください。');
  if (stepMinutes_(hours.endTo) <= stepMinutes_(hours.startFrom)) throw new Error('退勤のいちばん遅い時間は、出勤のいちばん早い時間より後にしてください。');

  var values = [['', '', ''], ['', '', ''], ['', '', '']];
  if (settings.manual) {
    var start = String(settings.start || '');
    var weeks = Number(settings.weeks);
    var deadlineDate = String(settings.deadlineDate || '');
    var deadlineTime = String(settings.deadlineTime || '');
    if (!/^\d{4}-\d{2}-\d{2}$/.test(start) || dowOfYmd_(start) !== 1) throw new Error('開始週は月曜日を選んでください。');
    if (start.slice(2).replace(/-/g, '') < FIRST_ROLE_SHEET) throw new Error(FIRST_ROLE_SHEET + ' より前の週は選べません。');
    if (!(weeks >= 1 && weeks <= 5 && Math.floor(weeks) === weeks)) throw new Error('週数は1〜5で選んでください。');
    if (!/^\d{4}-\d{2}-\d{2}$/.test(deadlineDate) || !/^\d{2}:\d{2}$/.test(deadlineTime)) throw new Error('スタッフ入力締切の日付と時間を選んでください。');
    if (jstMillis_(deadlineDate + ' ' + deadlineTime) <= Date.now()) throw new Error('今回だけのスタッフ入力締切が、もう過ぎています。');
    values = [[start], [String(weeks)], [deadlineDate + ' ' + deadlineTime]];
  } else {
    values = [[''], [''], ['']]; // 今回だけの設定を消して自動に戻す
  }

  var sheet = getRequestSettingsSheet_();
  sheet.getRange('B3:B5').setNumberFormat('@').setValues(values);
  var archiveUrl = String(settings.archiveUrl == null ? '' : settings.archiveUrl).trim();
  if (archiveUrl && !/[-\w]{25,}/.test(archiveUrl)) throw new Error('書庫ファイルのURLが正しくありません。');
  sheet.getRange('E3:E8').setNumberFormat('@')
    .setValues([[String(closingDay)], [String(deadlineWeeks)], [baseTime], [storeName], [String(castClosingDay)], [archiveUrl]]);
  sheet.getRange('E42:E45').setNumberFormat('@')
    .setValues([[hours.startFrom], [hours.startTo], [hours.endFrom], [hours.endTo]]);

  // 月次集計の見出し・締め期間も新しい設定に合わせる
  try {
    if (SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MONTHLY_SHEET_NAME)) refreshMonthlySummary_();
    if (SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DAILY_SHEET_NAME)) refreshDailyHours_();
  } catch (e) {
    Logger.log('月次集計の更新に失敗: ' + e.message);
  }
}

// ===================== 売上予算（日別） =====================
// 週シートの5行目（Y5・AI5・AS5・BC5・BM5・BW5・CG5＝月〜日）に、その日の売上予算を入れる。
// 数字は「売上予算」シートに置く。日付ごとの予算があればそれを、無ければ曜日別の基準額を使う。
// どちらも決まっていない日は、今シートに入っている数字をそのままにする（勝手に消さない）。
// 将来は「売上予算」シートを売上予測から自動で埋めれば、シフト表側は直さずに済む。
var BUDGET_SHEET_NAME = '売上予算';
var BUDGET_DOW_ROW = 4;        // 曜日別の基準額（B4〜H4＝月〜日）
var BUDGET_HEAD_ROW = 7;       // 日別の一覧の見出し（A=日付、B=予算）
var BUDGET_FIRST_ROW = 8;      // 日別の一覧の1行目
var BUDGET_WEEK_ROW = 5;       // 週シートの予算の行
var BUDGET_DEFAULT_DOW = [90000, 90000, 90000, 90000, 100000, 200000, 200000];
// 日別の予算の「日付」は、いつも先の月末まで並べておく（店長は金額を入れるだけで済む）。
// 2か月先＝11月なら1月31日まで。毎日の自動とメニュー「売上予算を入れ直す」のたびに、
// 足りないぶんだけうしろに足す。金額は空のままなので、入っていない日のお知らせは今までどおり出る。
// MARKER_BUDGET_DATE_ROWS
var BUDGET_AHEAD_MONTHS = 2;
var BUDGET_DATE_FORMAT = 'yyyy/mm/dd"("ddd")"';   // 土日が一目で分かるように曜日も出す

// 「売上予算」シートを取得（無ければ作る）。曜日別の基準額は、今ある一番新しい週シートの5行目から写す
function getBudgetSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(BUDGET_SHEET_NAME);
  if (sheet) return sheet;

  var active = ss.getActiveSheet(); // シート作成で表示中のシートが切り替わるため先に覚えておく
  sheet = ss.insertSheet(BUDGET_SHEET_NAME);
  var set = "'" + REQUEST_SETTINGS_SHEET + "'!";
  sheet.getRange('A1').setFormula('=IF(' + set + REQUEST_STORE_NAME_CELL + '="","売上予算",' + set + REQUEST_STORE_NAME_CELL + '&" 売上予算")')
    .setFontWeight('bold').setFontSize(14);
  sheet.getRange('A3').setValue('曜日別の基準額');
  sheet.getRange(3, 2, 1, 7).setValues([WEEK_DOW_JP]).setFontWeight('bold').setBackground('#EFF0ED');
  sheet.getRange('A' + BUDGET_DOW_ROW).setValue('金額');
  sheet.getRange(BUDGET_DOW_ROW, 2, 1, 7).setValues([newestWeekBudget_(ss)]);
  sheet.getRange('A6').setValue('日別の予算（入れた日は、上の基準額より優先します）');
  sheet.getRange(BUDGET_HEAD_ROW, 1, 1, 2).setValues([['日付', '予算']]).setFontWeight('bold').setBackground('#EFF0ED');
  sheet.setFrozenRows(BUDGET_HEAD_ROW);
  sheet.getRange('A:A').setNumberFormat(BUDGET_DATE_FORMAT);
  sheet.getRange('B:B').setNumberFormat('#,##0');
  sheet.setColumnWidth(1, 150);
  ss.setActiveSheet(active);
  return sheet;
}

// 今ある一番新しい週シートの5行目（月〜日の予算）。取れない曜日は初期値を使う
function newestWeekBudget_(ss) {
  var newest = newestWeekSheet_(ss);
  return DAY_BLOCKS.map(function(block, i) {
    var value = newest ? Number(newest.getRange(block.col + BUDGET_WEEK_ROW).getValue()) : 0;
    return value > 0 ? value : BUDGET_DEFAULT_DOW[i];
  });
}

// その年月の末日。月が12を超えたら翌年として数える
function monthEndYmd_(year, month) {
  var d = new Date(Date.UTC(year, month, 0));
  return d.getUTCFullYear() + '-' + ('0' + (d.getUTCMonth() + 1)).slice(-2) + '-' + ('0' + d.getUTCDate()).slice(-2);
}

// 日別の予算の「日付」の行を、先の月末まで用意する（金額は空のまま＝店長が入れる）。
// すでにある行のうしろに足すだけなので、入れた数字も並びも触らない。足した日数を返す。
// 「来月ぶんを入れたいのに行が無い」を無くすため、毎日の自動とメニューの両方から呼ぶ
function ensureBudgetDateRows_(ss, throughYmd) {
  var sheet = getBudgetSheet_();
  var tz = ss.getSpreadsheetTimeZone();
  var todayYmd = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  var to = monthEndYmd_(Number(todayYmd.slice(0, 4)), Number(todayYmd.slice(5, 7)) + BUDGET_AHEAD_MONTHS);
  // 募集期間がそれより先まで続くときは、その月の末日まで伸ばす
  if (throughYmd && throughYmd > to) to = monthEndYmd_(Number(throughYmd.slice(0, 4)), Number(throughYmd.slice(5, 7)));

  // すでに入っている日付（二重に足さないため）と、最後に日付が入っている行
  var have = {}, lastRow = BUDGET_HEAD_ROW, lastYmd = '';
  var last = sheet.getLastRow();
  if (last >= BUDGET_FIRST_ROW) {
    sheet.getRange(BUDGET_FIRST_ROW, 1, last - BUDGET_FIRST_ROW + 1, 1).getValues().forEach(function(r, i) {
      if (!(r[0] instanceof Date)) return;
      var ymd = Utilities.formatDate(r[0], tz, 'yyyy-MM-dd');
      have[ymd] = true;
      lastRow = BUDGET_FIRST_ROW + i;
      lastYmd = ymd;
    });
  }
  // 足しはじめるのは「最後の日付の翌日」から。終わった月までさかのぼらない
  var from = todayYmd.slice(0, 8) + '01';
  if (lastYmd && addDaysYmd_(lastYmd, 1) > from) from = addDaysYmd_(lastYmd, 1);
  if (from > to) return 0;

  var rows = [];
  for (var ymd = from; ymd <= to; ymd = addDaysYmd_(ymd, 1)) {
    if (have[ymd]) continue;
    // このファイルの時間帯で0時にする（ずれるとシートの日付が1日前になる）
    rows.push([Utilities.parseDate(ymd + ' 00:00', tz, 'yyyy-MM-dd HH:mm')]);
  }
  if (!rows.length) return 0;

  var need = lastRow + rows.length;
  if (need > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), need - sheet.getMaxRows());
  sheet.getRange(lastRow + 1, 1, rows.length, 1).setValues(rows);
  sheet.getRange('A:A').setNumberFormat(BUDGET_DATE_FORMAT);
  return rows.length;
}

// 「売上予算」シートを読む。{ dow: 月〜日の基準額, byDate: { 'yyyy-MM-dd': 金額 } }
function readBudgets_(ss) {
  var sheet = getBudgetSheet_();
  var tz = ss.getSpreadsheetTimeZone();
  var dow = sheet.getRange(BUDGET_DOW_ROW, 2, 1, 7).getValues()[0].map(function(v) { return Number(v) > 0 ? Number(v) : 0; });
  var byDate = {};
  var last = sheet.getLastRow();
  if (last >= BUDGET_FIRST_ROW) {
    sheet.getRange(BUDGET_FIRST_ROW, 1, last - BUDGET_FIRST_ROW + 1, 2).getValues().forEach(function(r) {
      if (!(r[0] instanceof Date) || !(Number(r[1]) > 0)) return;
      byDate[Utilities.formatDate(r[0], tz, 'yyyy-MM-dd')] = Number(r[1]);
    });
  }
  return { dow: dow, byDate: byDate };
}

// 1枚の週シートの5行目に予算を入れる。入れた数（変わったセルの数）を返す
function applyWeekBudget_(sheet, budgets) {
  var start = sheet.getRange('T1').getValue();
  if (!(start instanceof Date)) return 0;
  var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  var startYmd = Utilities.formatDate(start, tz, 'yyyy-MM-dd');
  var today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  var changed = 0;
  DAY_BLOCKS.forEach(function(block, i) {
    var ymd = addDaysYmd_(startYmd, i);
    // 日別の予算を入れた日は、終わった日でもその数字にする。
    // 曜日別の基準額は、これからの日にだけ使う（終わった週に入っている予算を、あとから書き換えないため）
    var value = budgets.byDate.hasOwnProperty(ymd) ? budgets.byDate[ymd] : (ymd >= today ? budgets.dow[i] : 0);
    if (!(value > 0)) return;   // 決まっていない日は、今の数字をそのままにする
    var cell = sheet.getRange(block.col + BUDGET_WEEK_ROW);
    if (Number(cell.getValue()) === value) return;
    cell.setValue(value);
    changed++;
  });
  return changed;
}

// メニュー「売上予算を入れ直す」：今ある週シート全部に、売上予算シートの数字を入れ直す
function refreshWeekBudgets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var created = !ss.getSheetByName(BUDGET_SHEET_NAME);
  var sheet = getBudgetSheet_();
  var added = ensureBudgetDateRows_(ss);   // 先の月末までの日付を用意してから案内する
  if (created) {
    ss.setActiveSheet(sheet);
    ui.alert('「売上予算」シートを作りました。\n\n曜日別の基準額（4行目）と、変える日だけ日別の予算（8行目から）を入れてください。\n日付は' +
      BUDGET_AHEAD_MONTHS + 'か月先の月末まで並べてあります（金額だけ入れてください）。\n次に週シートを作るときから、5行目の予算に自動で入ります。');
    return;
  }

  var budgets = readBudgets_(ss);
  var weeks = weekSheets_(ss);
  var answer = ui.alert('売上予算を入れ直す',
    '今ある週シート ' + weeks.length + '枚の5行目に、「売上予算」シートの数字を入れ直します。\n\n' +
    '・日別の予算を入れた日は、その数字にします（終わった日もふくむ）\n' +
    '・入れていない日は、これからの日だけ曜日別の基準額にします\n' +
    '・終わった日で日別の予算が無い日は、今の数字のままにします\n\n入れ直してよいですか？',
    ui.ButtonSet.OK_CANCEL);
  if (answer !== ui.Button.OK) return;

  var changed = 0, sheets = 0;
  weeks.forEach(function(week) {
    var n = applyWeekBudget_(week, budgets);
    if (n) { changed += n; sheets++; }
  });
  ui.alert((changed ? sheets + '枚のシートで、' + changed + '日ぶんの予算を入れ直しました。' : '入れ直す予算はありませんでした（すでに同じ数字です）。') +
    (added ? '\n\n※「売上予算」シートに、' + added + '日ぶんの日付を用意しました。金額を入れてください。' : ''));
}

// ===================== 目安人数 =====================
// 週シートの各曜日にある「目安」の行（9:00から30分ごとに30マス）を、
// 「目安人数」シートに登録した1時間ごとの人数から自動で入れる。
// 目安も、すぐ下の「実時間」も、画面には人数で出す（4人なら4）。
// 右はしの1日の合計だけは人時のまま（人件費の目安に使うため）。
//   自動人数 ＝ その時間の売上（1日の平均売上×構成比）÷ 目標人時売上高。最低人数を下回らせない。
//   手打ち人数 ＝ 店長が入れる数字。週シートにはこちらを入れる（空なら自動人数）。
var TARGET_SHEET_NAME = '目安人数';
var TARGET_FIRST_HOUR = 9;               // 目安の先頭は9:00
var TARGET_HOURS = 15;                   // 9時〜23時
var TARGET_SLOTS = TARGET_HOURS * 2;     // 週シートの30分マスの数（30）
var TARGET_CATEGORIES = ['平日', '土曜', '日曜', '祝日'];
var TARGET_SPH_ROW = 3;                  // B3 目標人時売上高
var TARGET_MAXSPH_ROW = 4;               // B4 上限人時売上高
var TARGET_MIN_ROW = 5;                  // B5 最低人数
var TARGET_OPEN_ROW = 6;                 // B6・C6 人を置く時間（開始〜終了）
var TARGET_CAT_ROW = 7;                  // 7行目 区分の見出し
var TARGET_AVG_ROW = 8;                  // 8行目 1日の平均売上（自動）
var TARGET_SPH_ACTUAL_ROW = 9;           // 9行目 手打ちでの人時売上高
var TARGET_SUM_ROW = 10;                 // 10行目 手打ちの合計人時
var TARGET_HEAD_ROW = 11;                // 11行目 列の見出し
var TARGET_FIRST_ROW = 12;               // 12行目が9時
var TARGET_RATIO_SUM_ROW = 27;           // 27行目 構成比の合計
var TARGET_FIRST_COL = 2;                // B列から
var TARGET_COLS_PER_CAT = 3;             // 構成比・自動人数・手打ち人数
var TARGET_DEFAULT = { sph: 5000, maxSph: 10000, minPeople: 1, openFrom: 9, openTo: 22 };
// 最初に入れておく構成比（仮の数字。お店の実際の売れ方に直して使う）
var TARGET_DEFAULT_RATIO = [1, 2, 10, 16, 11, 6, 4, 4, 6, 11, 12, 9, 5, 3, 0];
// 週シートの「目安」の場所（月→日の順。左に月火水、右に木金土日）
var TARGET_BLOCKS = [
  { row: 40, col: 27 },    // 月曜 AA
  { row: 59, col: 27 },    // 火曜 AA
  { row: 78, col: 27 },    // 水曜 AA
  { row: 20, col: 115 },   // 木曜 DK
  { row: 40, col: 115 },   // 金曜 DK
  { row: 59, col: 115 },   // 土曜 DK
  { row: 78, col: 115 }    // 日曜 DK
];

// 列番号を「A」「AA」のような文字に直す（数式を組み立てるため）
function colLetter_(n) {
  var s = '';
  while (n > 0) {
    var m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = (n - m - 1) / 26;
  }
  return s;
}

// 構成比の読み方をそろえる。
// 「10」と入れた時はそのまま10％、「10%」と入れた時はセルの中身が0.1なので100倍する。
// どちらかは、1日ぶんを全部足した数で見分ける（％で入っていれば合計は100前後、割合なら1前後）。
function normalizeRatios_(list) {
  var nums = list.map(function(v) { return Number(v) > 0 ? Number(v) : 0; });
  var total = nums.reduce(function(a, b) { return a + b; }, 0);
  if (total > 0 && total <= 2) return nums.map(function(n) { return n * 100; });
  return nums;
}

// 自動人数＝その時間の売上 ÷ 目標人時売上高。0.5人きざみ。
// 人を置く時間の外は0人。中は、売上が少なくても最低人数は下回らせない。
function autoPeople_(avgSales, ratioPercent, sph, minPeople, inRange) {
  if (!inRange) return 0;
  var people = (Number(ratioPercent) > 0 && Number(sph) > 0)
    ? Number(avgSales) * Number(ratioPercent) / 100 / Number(sph) : 0;
  if (people < minPeople) people = minPeople;
  return Math.round(people * 2) / 2;
}

// 区分の1列目（構成比の列）の列番号
function targetCatCol_(index) {
  return TARGET_FIRST_COL + index * TARGET_COLS_PER_CAT;
}

// 「目安人数」シートを取得（無ければ作る）
// 前の版で作ったシートにも、あとから足した欄を入れる
function ensureTargetSettings_(sheet) {
  if (String(sheet.getRange(TARGET_OPEN_ROW, 1).getValue()).trim() === '') {
    sheet.getRange(TARGET_OPEN_ROW, 1).setValue('人を置く時間（開始〜終了。終了の時間には置きません。9 でも 21:15 でも入れられます）');
    sheet.getRange(TARGET_OPEN_ROW, 2).setValue(TARGET_DEFAULT.openFrom);
    sheet.getRange(TARGET_OPEN_ROW, 3).setValue(TARGET_DEFAULT.openTo);
  }
  if (String(sheet.getRange(TARGET_RATIO_SUM_ROW, 1).getValue()).trim() === '') {
    sheet.getRange(TARGET_RATIO_SUM_ROW, 1).setValue('構成比の合計（100 または 100% になるように）');
    TARGET_CATEGORIES.forEach(function(name, i) {
      var letter = colLetter_(targetCatCol_(i));
      sheet.getRange(TARGET_RATIO_SUM_ROW, targetCatCol_(i))
        .setFormula('=SUM(' + letter + TARGET_FIRST_ROW + ':' + letter + (TARGET_FIRST_ROW + TARGET_HOURS - 1) + ')');
    });
  }
  return sheet;
}

// ----- 目安人数シートの見た目（線と色） -----
// 中身（数字や数式）には触らない。色でひと目で分かるようにする。
//   うすい黄色＝自分で入れるところ／うすい灰色＝自動で出るところ
var TARGET_LAST_COL = TARGET_FIRST_COL + TARGET_CATEGORIES.length * TARGET_COLS_PER_CAT - 1;
var TARGET_CAT_COLORS = ['#EFEFEF', '#CFE2F3', '#F4CCCC', '#FCE5CD'];  // 平日・土曜・日曜・祝日
var TARGET_INPUT_COLOR = '#FFFDE7';   // 自分で入れるところ
var TARGET_AUTO_COLOR = '#F3F3F3';    // 自動で出るところ
var TARGET_LABEL_COLOR = '#EFEFEF';   // 左の見出し
var TARGET_HEAD_COLOR = '#595959';    // 見出しの帯
var TARGET_LINE_COLOR = '#B7B7B7';

function styleTargetSheet_(sheet) {
  var cols = TARGET_LAST_COL;
  var lastRow = TARGET_RATIO_SUM_ROW;
  var solid = SpreadsheetApp.BorderStyle.SOLID;
  var medium = SpreadsheetApp.BorderStyle.SOLID_MEDIUM;

  // 表題と注意書き。
  // ここをA列から結合すると、下の setFrozenColumns(1) が
  // 「結合されたセルの一部だけを含む列を固定することはできません」で止まる（2026-10-07に発生）。
  // 結合はやめて、色だけ行いっぱいに付け、文字は右へあふれさせる（見た目は同じ）。
  sheet.getRange(1, 1, 2, cols).breakApart();
  sheet.getRange(1, 1, 1, cols)
    .setBackground('#1C4587').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(13)
    .setWrap(false).setVerticalAlignment('middle');
  sheet.setRowHeight(1, 30);
  sheet.getRange(2, 1, 1, cols)
    .setFontColor('#666666').setFontSize(10).setWrap(false).setVerticalAlignment('middle');
  sheet.setRowHeight(2, 32);

  // 設定のところ（3〜6行目）
  sheet.getRange(TARGET_SPH_ROW, 1, 4, 1).setBackground(TARGET_LABEL_COLOR).setFontWeight('bold');
  sheet.getRange(TARGET_SPH_ROW, 2, 3, 1).setBackground(TARGET_INPUT_COLOR);
  sheet.getRange(TARGET_OPEN_ROW, 2, 1, 2).setBackground(TARGET_INPUT_COLOR);
  sheet.getRange(TARGET_SPH_ROW, 2, 2, 1).setNumberFormat('#,##0"円"');
  sheet.getRange(TARGET_OPEN_ROW, 2, 1, 2).setNumberFormat('@');   // 「21:15」と入れても時刻に変わらないように
  sheet.getRange(TARGET_SPH_ROW, 1, 4, 3)
    .setBorder(true, true, true, true, true, true, TARGET_LINE_COLOR, solid);

  // 表の全体に細い線を引いてから、区分ごとの区切りを太くする
  sheet.getRange(TARGET_CAT_ROW, 1, lastRow - TARGET_CAT_ROW + 1, cols)
    .setBorder(true, true, true, true, true, true, TARGET_LINE_COLOR, solid);

  // 左の見出しの列
  sheet.getRange(TARGET_CAT_ROW, 1, 4, 1).setBackground(TARGET_LABEL_COLOR).setFontWeight('bold');
  sheet.getRange(TARGET_FIRST_ROW, 1, TARGET_HOURS, 1)
    .setBackground(TARGET_LABEL_COLOR).setFontWeight('bold').setHorizontalAlignment('center');

  // 自動で出る3行（平均売上・人時売上高・合計人時）
  sheet.getRange(TARGET_AVG_ROW, 2, 3, cols - 1).setBackground(TARGET_AUTO_COLOR);

  // 列の見出し（11行目）
  sheet.getRange(TARGET_HEAD_ROW, 1, 1, cols)
    .setBackground(TARGET_HEAD_COLOR).setFontColor('#FFFFFF').setFontWeight('bold')
    .setHorizontalAlignment('center');

  // 区分ごと（平日・土曜・日曜・祝日）
  TARGET_CATEGORIES.forEach(function(name, i) {
    var c = targetCatCol_(i);
    sheet.getRange(TARGET_CAT_ROW, c, 1, TARGET_COLS_PER_CAT).breakApart().merge()
      .setBackground(TARGET_CAT_COLORS[i]).setFontWeight('bold').setHorizontalAlignment('center');
    sheet.getRange(TARGET_FIRST_ROW, c, TARGET_HOURS, 1).setBackground(TARGET_INPUT_COLOR);
    sheet.getRange(TARGET_FIRST_ROW, c + 1, TARGET_HOURS, 1)
      .setBackground(TARGET_AUTO_COLOR).setNumberFormat('0.0').setHorizontalAlignment('center');
    sheet.getRange(TARGET_FIRST_ROW, c + 2, TARGET_HOURS, 1)
      .setBackground(TARGET_INPUT_COLOR).setNumberFormat('0.0').setHorizontalAlignment('center');
    sheet.getRange(TARGET_CAT_ROW, c, lastRow - TARGET_CAT_ROW + 1, TARGET_COLS_PER_CAT)
      .setBorder(true, true, true, true, null, null, '#666666', medium);
  });

  // 構成比の合計の行
  sheet.getRange(TARGET_RATIO_SUM_ROW, 1, 1, cols).setBackground('#D9D9D9').setFontWeight('bold');

  // 合計が100（または100%）でなければ赤くする
  var sumCells = TARGET_CATEGORIES.map(function(name, i) {
    return colLetter_(targetCatCol_(i)) + TARGET_RATIO_SUM_ROW;
  });
  var keep = sheet.getConditionalFormatRules().filter(function(rule) {
    var hit = false;
    rule.getRanges().forEach(function(r) { if (sumCells.indexOf(r.getA1Notation()) >= 0) hit = true; });
    return !hit;
  });
  var head = sumCells[0];
  keep.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(' + head + '<>"",' + head + '<>100,ROUND(' + head + ',4)<>1)')
    .setBackground('#F4CCCC').setFontColor('#9C0006')
    .setRanges(sumCells.map(function(a1) { return sheet.getRange(a1); }))
    .build());
  sheet.setConditionalFormatRules(keep);

  // 幅と、見出しの固定
  sheet.setColumnWidth(1, 210);
  for (var c2 = TARGET_FIRST_COL; c2 <= cols; c2++) sheet.setColumnWidth(c2, 84);
  sheet.setFrozenRows(TARGET_HEAD_ROW);
  sheet.setFrozenColumns(1);
  return sheet;
}

function getTargetSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(TARGET_SHEET_NAME);
  if (sheet) return ensureTargetSettings_(sheet);

  var active = ss.getActiveSheet();
  sheet = ss.insertSheet(TARGET_SHEET_NAME);
  sheet.getRange('A1').setValue(TARGET_SHEET_NAME + '（メニュー「スタッフ管理」→「目安人数を開く」）');
  sheet.getRange('A2').setValue('※構成比は最初は仮の数字です。お店の実際の売れ方に直してください。手打ち人数が空の時は自動人数を使います。');
  sheet.getRange(TARGET_SPH_ROW, 1).setValue('目標人時売上高（円）');
  sheet.getRange(TARGET_SPH_ROW, 2).setValue(TARGET_DEFAULT.sph);
  sheet.getRange(TARGET_MAXSPH_ROW, 1).setValue('上限人時売上高（円。これを超えると人が少なすぎ）');
  sheet.getRange(TARGET_MAXSPH_ROW, 2).setValue(TARGET_DEFAULT.maxSph);
  sheet.getRange(TARGET_MIN_ROW, 1).setValue('最低人数（人を置く時間は、この人数を下回らせない）');
  sheet.getRange(TARGET_MIN_ROW, 2).setValue(TARGET_DEFAULT.minPeople);
  ensureTargetSettings_(sheet);

  sheet.getRange(TARGET_CAT_ROW, 1).setValue('区分');
  sheet.getRange(TARGET_AVG_ROW, 1).setValue('1日の平均売上（自動）');
  sheet.getRange(TARGET_SPH_ACTUAL_ROW, 1).setValue('手打ちでの人時売上高');
  sheet.getRange(TARGET_SUM_ROW, 1).setValue('手打ちの合計人時');
  sheet.getRange(TARGET_HEAD_ROW, 1).setValue('時間');

  // 今ある一番新しい週シートの目安から、手打ち人数の初期値をもらう（無ければ空のまま）
  var fromWeek = newestWeekTargets_(ss);

  TARGET_CATEGORIES.forEach(function(name, i) {
    var c = targetCatCol_(i);
    var letter = colLetter_(c), manual = colLetter_(c + 2);
    sheet.getRange(TARGET_CAT_ROW, c).setValue(name);
    sheet.getRange(TARGET_HEAD_ROW, c, 1, 3).setValues([['構成比%', '自動人数', '手打ち人数']]);
    sheet.getRange(TARGET_SUM_ROW, c)
      .setFormula('=SUM(' + manual + TARGET_FIRST_ROW + ':' + manual + (TARGET_FIRST_ROW + TARGET_HOURS - 1) + ')');
    sheet.getRange(TARGET_SPH_ACTUAL_ROW, c)
      .setFormula('=IFERROR(' + letter + TARGET_AVG_ROW + '/' + letter + TARGET_SUM_ROW + ',"")');

    var rows = [];
    for (var h = 0; h < TARGET_HOURS; h++) {
      rows.push([TARGET_DEFAULT_RATIO[h], '', fromWeek[name] ? fromWeek[name][h] : '']);
    }
    sheet.getRange(TARGET_FIRST_ROW, c, TARGET_HOURS, 3).setValues(rows);
  });

  var hours = [];
  for (var h = 0; h < TARGET_HOURS; h++) hours.push([(TARGET_FIRST_HOUR + h) + ':00']);
  sheet.getRange(TARGET_FIRST_ROW, 1, TARGET_HOURS, 1).setValues(hours);

  sheet.getRange(TARGET_AVG_ROW, 2, 1, 20).setNumberFormat('#,##0');
  sheet.getRange(TARGET_SPH_ACTUAL_ROW, 2, 1, 20).setNumberFormat('#,##0');
  styleTargetSheet_(sheet);
  refreshTargetSheet_(sheet);
  ss.setActiveSheet(active);
  return sheet;
}

// 一番新しい週シートの「目安」から、区分ごとの1時間あたり人数をもらう（目安は30分ぶんの人時なので2倍）
function newestWeekTargets_(ss) {
  var out = {};
  var sheet = newestWeekSheet_(ss);
  if (!sheet) return out;
  var pick = { 0: '平日', 5: '土曜', 6: '日曜' };
  Object.keys(pick).forEach(function(i) {
    var block = TARGET_BLOCKS[Number(i)];
    var values = sheet.getRange(block.row, block.col, 1, TARGET_SLOTS * 2 - 1).getValues()[0];
    var hours = [];
    for (var h = 0; h < TARGET_HOURS; h++) {
      var v = Number(values[h * 4]);      // 1時間の先頭の30分マス（目安は2列おきなので4つ飛ばし）
      hours.push(v > 0 ? v * 2 : '');
    }
    out[pick[i]] = hours;
  });
  if (out['日曜']) out['祝日'] = out['日曜'].slice();   // 祝日は日曜と同じから始める
  return out;
}

// 「祝日」シートの日付一覧
function holidayKeys_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('祝日');
  var keys = {};
  if (!sheet || sheet.getLastRow() < 2) return keys;
  sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().forEach(function(r) {
    if (r[0] instanceof Date) keys[Utilities.formatDate(r[0], Session.getScriptTimeZone(), 'yyyy-MM-dd')] = true;
  });
  return keys;
}

// その日がどの区分か。祝日＞日曜＞土曜＞平日の順で決める
function categoryOfYmd_(ymd, keys) {
  if (keys && keys[ymd]) return '祝日';
  var dow = dowOfYmd_(ymd);
  if (dow === 0) return '日曜';
  if (dow === 6) return '土曜';
  return '平日';
}

// 区分ごとの1日の平均売上を、売上予算シートから出す
function averageSalesByCategory_(ss) {
  var budgets = readBudgets_(ss);
  var keys = holidayKeys_();
  var sum = {}, count = {};
  TARGET_CATEGORIES.forEach(function(name) { sum[name] = 0; count[name] = 0; });
  Object.keys(budgets.byDate).forEach(function(ymd) {
    var cat = categoryOfYmd_(ymd, keys);
    sum[cat] += budgets.byDate[ymd];
    count[cat]++;
  });
  // 日別の予算が1件も無い区分は、曜日別の基準額から出す（平日は月〜金の平均）
  var out = {};
  TARGET_CATEGORIES.forEach(function(name) {
    if (count[name]) { out[name] = Math.round(sum[name] / count[name]); return; }
    if (name === '土曜') out[name] = budgets.dow[5];
    else if (name === '日曜' || name === '祝日') out[name] = budgets.dow[6];
    else {
      var weekday = budgets.dow.slice(0, 5).filter(function(v) { return v > 0; });
      out[name] = weekday.length ? Math.round(weekday.reduce(function(a, b) { return a + b; }, 0) / weekday.length) : 0;
    }
  });
  return out;
}

// 平均売上（売上予算シートから）と自動人数を、今の設定で入れ直す
function refreshTargetSheet_(sheet) {
  ensureTargetSettings_(sheet);
  var averages = averageSalesByCategory_(SpreadsheetApp.getActiveSpreadsheet());
  var sph = Number(sheet.getRange(TARGET_SPH_ROW, 2).getValue()) || TARGET_DEFAULT.sph;
  var minPeople = Number(sheet.getRange(TARGET_MIN_ROW, 2).getValue());
  if (!(minPeople >= 0)) minPeople = TARGET_DEFAULT.minPeople;
  var open = targetOpenHours_(sheet);
  TARGET_CATEGORIES.forEach(function(name, i) {
    var c = targetCatCol_(i);
    var avg = averages[name] || 0;
    sheet.getRange(TARGET_AVG_ROW, c).setValue(avg);
    var ratios = normalizeRatios_(sheet.getRange(TARGET_FIRST_ROW, c, TARGET_HOURS, 1).getValues()
      .map(function(r) { return r[0]; }));
    var autos = ratios.map(function(ratio, h) {
      var hour = TARGET_FIRST_HOUR + h;
      // その1時間が、人を置く時間と少しでも重なっていれば中とみなす（21:15 なら21時台も中）
      var inside = hour * 60 < open.toMin && (hour + 1) * 60 > open.fromMin;
      return [autoPeople_(avg, ratio, sph, minPeople, inside)];
    });
    sheet.getRange(TARGET_FIRST_ROW, c + 1, TARGET_HOURS, 1).setValues(autos);
  });
  return averages;
}

// 「22」（時だけ）と「21:15」（15分きざみ）の両方を分で読み取る
function openMinutesOf_(cell, fallbackHour) {
  var raw = cell.getValue();
  if (raw instanceof Date) return raw.getHours() * 60 + raw.getMinutes();
  var text = String(cell.getDisplayValue() == null ? '' : cell.getDisplayValue()).trim();
  var m = /^(\d{1,2}):(\d{2})$/.exec(text);
  if (m) {
    var min = Number(m[1]) * 60 + Number(m[2]);
    if (min >= 0 && min <= 24 * 60) return min;
  }
  var digits = String(text).replace(/[^0-9]/g, '');
  if (digits) {
    var hour = Number(digits);
    if (hour >= 0 && hour <= 24) return hour * 60;
  }
  return fallbackHour * 60;   // 空や書き間違いは初期値
}

// スタッフが退勤できる、いちばん遅い時間（分）。この後ろに人は置けない
function staffEndLimitMin_() {
  try {
    var settings = getRequestSettingsSheet_();
    return stepMinutes_(stepTime_(settings.getRange(REQUEST_END_TO_CELL).getDisplayValue(), REQUEST_DEFAULT_BASE.endTo));
  } catch (err) {
    return stepMinutes_(REQUEST_DEFAULT_BASE.endTo);
  }
}

// 人を置く時間（開始〜終了。終了の時間には置かない）。
// スタッフが退勤できる上限（シフト希望設定）より後ろは、誰も入れないので終わりにする
function targetOpenHours_(sheet) {
  var fromMin = openMinutesOf_(sheet.getRange(TARGET_OPEN_ROW, 2), TARGET_DEFAULT.openFrom);
  var toMin = openMinutesOf_(sheet.getRange(TARGET_OPEN_ROW, 3), TARGET_DEFAULT.openTo);
  if (!(toMin > fromMin)) {
    fromMin = TARGET_DEFAULT.openFrom * 60;
    toMin = TARGET_DEFAULT.openTo * 60;
  }
  var limit = staffEndLimitMin_();
  if (limit > fromMin && limit < toMin) toMin = limit;
  return { from: Math.floor(fromMin / 60), to: Math.ceil(toMin / 60), fromMin: fromMin, toMin: toMin };
}

// 不足を見る終わりの時間（分）。「人を置く時間の終了」と「退勤できる上限」の早いほう
function shiftEndLimitMin_() {
  var staff = staffEndLimitMin_();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TARGET_SHEET_NAME);
  if (!sheet) return staff;
  var toMin = openMinutesOf_(sheet.getRange(TARGET_OPEN_ROW, 3), TARGET_DEFAULT.openTo);
  return Math.min(staff, toMin);
}

// 目安人数シートを読む。手打ちが空の時は自動人数を使う
function readTargets_() {
  var sheet = getTargetSheet_();
  refreshTargetSheet_(sheet);          // 構成比や設定を直した直後でも合うようにする
  var values = sheet.getRange(1, 1, TARGET_FIRST_ROW + TARGET_HOURS, targetCatCol_(TARGET_CATEGORIES.length)).getValues();
  var byCat = {};
  TARGET_CATEGORIES.forEach(function(name, i) {
    var c = targetCatCol_(i) - 1;      // 0から数えた列
    var hours = [];
    for (var h = 0; h < TARGET_HOURS; h++) {
      var row = values[TARGET_FIRST_ROW - 1 + h] || [];
      var manual = Number(row[c + 2]);
      var auto = Number(row[c + 1]);
      hours.push(manual > 0 ? manual : (auto > 0 ? auto : 0));
    }
    byCat[name] = hours;
  });
  return {
    sph: Number(sheet.getRange(TARGET_SPH_ROW, 2).getValue()) || TARGET_DEFAULT.sph,
    maxSph: Number(sheet.getRange(TARGET_MAXSPH_ROW, 2).getValue()) || TARGET_DEFAULT.maxSph,
    minPeople: Number(sheet.getRange(TARGET_MIN_ROW, 2).getValue()),
    open: targetOpenHours_(sheet),
    byCat: byCat
  };
}

// 1枚の週シートの「目安」を入れ直す。
// 目安のマスは結合されていることがあるので、変わるマスだけ1つずつ書く。
function applyWeekTargets_(sheet, targets, keys, tz) {
  var top = sheet.getRange('T1').getValue();
  if (!(top instanceof Date)) return 0;
  useHeadcountDisplay_(sheet);                         // 古い「人時」表示の週は、先に人数表示へ直す
  var start = Utilities.formatDate(top, tz, 'yyyy-MM-dd');
  var open = targets.open || { fromMin: 0, toMin: 24 * 60 };
  var changed = 0;
  TARGET_BLOCKS.forEach(function(block, i) {
    var hours = targets.byCat[categoryOfYmd_(addDaysYmd_(start, i), keys)];
    if (!hours) return;
    var now = sheet.getRange(block.row, block.col, 1, TARGET_SLOTS * 2 - 1).getValues()[0];
    for (var s = 0; s < TARGET_SLOTS; s++) {
      // 30分のマスが「人を置く時間」と重なっていなければ0にする。
      // 21:15 までなら 21:30〜22:00 のマスは空＝そこに不足は出ない
      var slotFrom = TARGET_FIRST_HOUR * 60 + s * 30;
      var inside = slotFrom < open.toMin && slotFrom + 30 > open.fromMin;
      var people = inside ? Number(hours[Math.floor(s / 2)]) : 0;
      var want = people > 0 ? people : '';             // 画面には人数のまま出す
      if (String(now[s * 2]) === String(want)) continue;
      sheet.getRange(block.row, block.col + s * 2).setValue(want);
      changed++;
    }
  });
  return changed;
}

// 「目安」と「実時間」を人数で表示するように直す（前は30分ぶんの人時＝人数の半分だった）。
// 直したかどうかは、実時間の式に「*2」が付いているかで見分けるので、何回実行しても二重にはならない。
// 右はしの1日の合計は人時のままにしたいので、そこだけ「/2」を付ける。
function useHeadcountDisplay_(sheet) {
  var width = TARGET_SLOTS * 2 + 3;      // 実時間の先頭のマスから、右はしの合計まで
  var last = width - 3;                  // いちばん右のマス（0から数えた位置）
  var sumAt = width - 1;                 // 1日の合計（0から数えた位置）
  var fixed = 0;
  TARGET_BLOCKS.forEach(function(block) {
    var firstCol = block.col - 2;        // 実時間は目安より1マス（2列）左から始まる
    var actual = sheet.getRange(block.row + 1, firstCol, 1, width);
    var formulas = actual.getFormulas()[0];
    var joined = formulas.join('');
    if (!joined) return;                 // 実時間の式が無いシートは触らない
    if (joined.indexOf('*2') !== -1) return;          // もう人数表示になっている

    for (var c = 0; c <= last; c += 2) {
      if (formulas[c]) actual.getCell(1, c + 1).setFormula(formulas[c] + '*2');
    }
    if (formulas[sumAt]) actual.getCell(1, sumAt + 1).setFormula(formulas[sumAt] + '/2');

    // 目安のマスは結合されていることがあるので、1つずつ書く
    var guide = sheet.getRange(block.row, firstCol, 1, width);
    var values = guide.getValues()[0];
    var guideFormulas = guide.getFormulas()[0];
    for (var g = 0; g <= last; g += 2) {
      if (typeof values[g] === 'number' && values[g] > 0) guide.getCell(1, g + 1).setValue(values[g] * 2);
    }
    if (guideFormulas[sumAt]) guide.getCell(1, sumAt + 1).setFormula(guideFormulas[sumAt] + '/2');
    fixed++;
  });
  return fixed;
}

// 手打ちの人数だと人時売上高が上限を超えてしまう区分（人が少なすぎる区分）を挙げる
function tooLeanCategories_(ss, targets) {
  var averages = averageSalesByCategory_(ss);
  return TARGET_CATEGORIES.filter(function(name) {
    var total = (targets.byCat[name] || []).reduce(function(a, b) { return a + Number(b || 0); }, 0);
    if (!total || !averages[name]) return false;
    return averages[name] / total > targets.maxSph;
  });
}

// メニュー「目安人数を開く」
function openTargetSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getTargetSheet_();
  refreshTargetSheet_(sheet);
  styleTargetSheet_(sheet);
  ss.setActiveSheet(sheet);
  SpreadsheetApp.getUi().alert(
    '「目安人数」シートを開きました。\n\n' +
    '・構成比％：その時間に1日の売上の何％が売れるか（合計が100、または100%になるように）\n' +
    '　　「10」でも「10%」でも、どちらの書き方でも読み取ります。\n' +
    '・人を置く時間（B6〜C6）：この外の時間は0人になります（例 9〜22なら21時台まで）\n' +
    '　　「21:15」のように15分きざみでも入れられます。その場合、21:30〜22:00には人を置きません。\n' +
    '　　シフト希望設定の「退勤のいちばん遅い時間」より後ろには、どちらにしても人を置きません。\n' +
    '・自動人数：1日の平均売上×構成比÷目標人時売上高（人を置く時間は最低人数を下回りません）\n' +
    '・手打ち人数：実際に置く人数。週シートにはこちらが入ります（空なら自動人数）\n' +
    '　　週シートの「目安」には、この人数がそのまま出ます（4人なら4）。\n\n' +
    '直したら、メニューの「目安人数を週シートに入れ直す」を押してください。\n' +
    '自動人数は、このメニューを開いたときと入れ直したときに計算し直します。');
}

// メニュー「目安人数を週シートに入れ直す」（これから始まる週だけ。終わった週は触らない）
function refreshWeekTargets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var tz = ss.getSpreadsheetTimeZone();
  var today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var weeks = ss.getSheets().filter(function(s) {
    if (!/^\d{6}$/.test(s.getName())) return false;
    var top = s.getRange('T1').getValue();
    if (!(top instanceof Date)) return false;
    return addDaysYmd_(Utilities.formatDate(top, tz, 'yyyy-MM-dd'), 6) >= today;   // 終わっていない週だけ
  });
  if (!weeks.length) { ui.alert('入れ直す週シートがありません（終わった週は触りません）。'); return; }

  var names = weeks.map(function(s) { return s.getName(); }).join('、');
  if (ui.alert('次の週シートの「目安」を入れ直します。\n\n' + names + '\n\n終わった週は触りません。よろしいですか。',
      ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;

  var targets = readTargets_();
  var keys = holidayKeys_();
  var changed = 0, sheets = 0;
  weeks.forEach(function(week) {
    var n = applyWeekTargets_(week, targets, keys, tz);
    if (n) { changed += n; sheets++; }
  });
  var lean = tooLeanCategories_(ss, targets);
  ui.alert((changed ? sheets + '枚のシートで、' + changed + 'マスの目安を入れ直しました。'
                    : '入れ直す目安はありませんでした（すでに同じ数字です）。') +
    (lean.length ? '\n\n※' + lean.join('・') + ' は人時売上高が' + targets.maxSph +
      '円を超えます（人が少なすぎます）。「目安人数」シートの手打ち人数を見直してください。' : ''));
}

// ===================== 削った時間を薄く表示（確定まで残す） =====================
// 転写のあとに時間を削ると、右側の時間帯バーのマスがすぐ消えてしまう。
// 「調整前」の状態を控えておき、削ったマスを薄いグレーで残して、確定したら消す。
//   転写すると自動で「調整前」を記録し、時間を直すとその場で薄いグレーになる。
//   決まったら「確定する（薄い表示を消す）」。おかしくなったら「色を塗り直す（掃除）」。
// 数式・条件付き書式・列幅には一切触らない（背景色を塗って、確定で元の色に戻すだけ）。
var ADJUST_SHEET_NAME = '調整前の時間';
var ADJUST_ROWS = 16;              // 1つの曜日ブロックの人数
var ADJUST_GRID_COLS = 62;         // 15分きざみのマスの数（左はY〜CH、右はDI〜FR）
var ADJUST_COLOR = '#E0E0E0';      // 削ったマスの薄いグレー
// 自動で動く処理（onEditの塗り直しなど）が失敗しても、画面には何も出ない＝気づけない。
// 最後の失敗だけ、この控えシートの隅に残しておく（あとから原因を見るため）。MARKER_AUTO_FAIL_NOTE
var AUTO_FAIL_LABEL_CELL = 'G1';
var AUTO_FAIL_CELL = 'H1';
// 曜日ブロックの場所（月→日の順。head＝見出しの行、col＝升目の先頭列）
var ADJUST_BLOCKS = [
  { head: 23, col: 25 },    // 月曜
  { head: 42, col: 25 },    // 火曜
  { head: 61, col: 25 },    // 水曜
  { head: 3,  col: 113 },   // 木曜
  { head: 23, col: 113 },   // 金曜
  { head: 42, col: 113 },   // 土曜
  { head: 61, col: 113 }    // 日曜
];

// 控えのシート（非表示）を取得（無ければ作る）
function getAdjustSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(ADJUST_SHEET_NAME);
  if (sheet) return sheet;
  var active = ss.getActiveSheet();
  sheet = ss.insertSheet(ADJUST_SHEET_NAME);
  sheet.getRange(1, 1, 1, 5)
    .setValues([['シート名', '曜日', '記録日時', '調整前の在店（15分ごとに1か0）', '薄く塗ったマスの控え']])
    .setFontWeight('bold');
  sheet.getRange('A:A').setNumberFormat('@');
  ss.setActiveSheet(active);
  sheet.hideSheet();
  return sheet;
}

// 控えシートの「シート名/曜日」→行番号
function adjustRowMap_(sheet) {
  var map = {};
  var last = sheet.getLastRow();
  if (last < 2) return map;
  sheet.getRange(2, 1, last - 1, 2).getValues().forEach(function(r, i) {
    if (String(r[0]) === '') return;
    map[String(r[0]) + '/' + r[1]] = i + 2;
  });
  return map;
}

// 1つの曜日ブロックの「人がいるマス」を 1と0 の文字にする（人ごとに | で区切る）
function gridSnapshot_(week, block) {
  return week.getRange(block.head + 1, block.col, ADJUST_ROWS, ADJUST_GRID_COLS).getValues()
    .map(function(row) {
      return row.map(function(v) { return Number(v) > 0 ? '1' : '0'; }).join('');
    }).join('|');
}

// 今の状態を「調整前」として控えに書く
function saveAdjustSnapshots_(week) {
  var sheet = getAdjustSheet_();
  var map = adjustRowMap_(sheet);
  var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
  ADJUST_BLOCKS.forEach(function(block, i) {
    var row = map[week.getName() + '/' + i];
    if (!row) {
      row = Math.max(sheet.getLastRow() + 1, 2);
      sheet.getRange(row, 1, 1, 2).setValues([[week.getName(), i]]);
    }
    sheet.getRange(row, 3, 1, 3).setValues([[now, gridSnapshot_(week, block), '']]);
  });
}

// 薄いグレーを、塗る前の色に戻す
function clearAdjustMarks_(week, ours) {
  var sheet = getAdjustSheet_();
  var map = adjustRowMap_(sheet);
  var cleared = 0;
  ADJUST_BLOCKS.forEach(function(block, i) {
    var row = map[week.getName() + '/' + i];
    if (!row) return;
    var text = String(sheet.getRange(row, 5).getValue());
    if (!text) return;
    var keep = [];
    try { keep = JSON.parse(text); } catch (e) { keep = []; }
    if (keep.length) {
      var range = week.getRange(block.head + 1, block.col, ADJUST_ROWS, ADJUST_GRID_COLS);
      var colors = range.getBackgrounds();
      keep.forEach(function(item) {
        var back = String(item[2]);
        if (ours && isOurColor_(back, ours)) back = '#ffffff';   // 水色やポジションの色は返さない
        colors[item[0]][item[1]] = back;
        cleared++;
      });
      range.setBackgrounds(colors);
    }
    sheet.getRange(row, 5).setValue('');
  });
  return cleared;
}

// 開いているシートが週シートかどうか
function activeWeekSheet_() {
  var week = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (!/^\d{6}$/.test(week.getName())) {
    SpreadsheetApp.getUi().alert('週シート（261019 のような名前）を開いてから押してください。');
    return null;
  }
  return week;
}

// メニュー「確定する（薄い表示を消す）」
function confirmAdjust() {
  var week = activeWeekSheet_();
  if (!week) return;
  var ui = SpreadsheetApp.getUi();
  if (ui.alert('「' + week.getName() + '」（' + weekRangeLabel_(week.getName()) + '）を確定します。\n\n' +
      '・薄いグレーの表示を消して、今の時間を新しい「調整前」にします\n' +
      '・スタッフのページとカレンダーに、この週を公開します\n' +
      '・全員の表をPDFにして、LINEのグループにお知らせを送ります\n\nよろしいですか。',
      ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  syncRoleSheetBirths_();                    // 名簿の生年月日をスタッフ区分シートにも入れておく
  var found = checkWeekLaw_(week, true);
  var unknown = unknownBirthNames_(week);
  if ((found.length || unknown.length) && ui.alert('決まりの確認（' + week.getName() + '）',
      lawReportText_(found, unknown) + '\n\nこのまま確定して公開しますか。\n「いいえ」を押すと、直してから確定できます。',
      ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  ensureGridPaintable_(week);          // 自動の水色ルールが残っていれば、ここでも外す
  var cleared = clearAdjustMarks_(week, ourColors_());
  saveAdjustSnapshots_(week);
  repaintWeekColors_(week);
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '確定しました。薄い表示を' + cleared + 'マス消しました。続けて公開します。', 'シフト表', 6);
  SpreadsheetApp.flush();
  publishWeek_(week, false, false);
}

// メニュー用：もう一度お知らせを送る（PDFも作り直す）
function publishWeekAgain() {
  var week = activeWeekSheet_();
  if (!week) return;
  var ui = SpreadsheetApp.getUi();
  if (ui.alert('「' + week.getName() + '」（' + weekRangeLabel_(week.getName()) + '）のPDFを作り直して、\n' +
      'LINEのグループへもう一度お知らせを送ります。よろしいですか。',
      ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  publishWeek_(week, true, false);
}


// ===================== 確定した週をスタッフに公開する =====================
// 「確定する」を押すと、次の3つを行う。
//  ①その週を「公開した週」として覚える（スタッフのページとカレンダーは、この週だけを出す）
//  ②全員の表（R1:FU79）をPDFにして、ドライブの「シフトPDF」フォルダに置く（同じ週は同じファイルを作り直す＝リンクは変わらない）
//  ③LINEのグループへ、ページとPDFのリンクを送る（同じ週に二重で送らない。送り直しはメニューから）
var PUBLISHED_WEEKS_CELL = 'E26';     // 公開した週（新しい順ではなく、押した順に足す）
var PDF_FOLDER_CELL = 'E27';          // PDFの入れ物フォルダのID
var PUBLISHED_KEEP = 30;              // 覚えておく週の数
var PDF_RANGE = 'R1:GJ79';            // 印刷する範囲（全員の表。右のはみ出した文字が切れないよう、空の列を少し足している）
var PDF_FOLDER_NAME = 'シフトPDF';
// スタッフが見るページ（シフト希望フォームと同じアドレス。作り直さない決まり）
var STAFF_PAGE_URL = 'https://script.google.com/macros/s/AKfycbwFtgsWnq75UnttAODig9q3c89E1Iek0Oow-lV8tHzF_Zp5qBQe2DXPahAeksqkd4i2/exec';

// 週シート名 261019 → '10/19(月)〜10/25(日)'
function weekRangeLabel_(sheetName) {
  var start = '20' + sheetName.slice(0, 2) + '-' + sheetName.slice(2, 4) + '-' + sheetName.slice(4, 6);
  var end = addDaysYmd_(start, 6);
  return ymdLabel_(start) + '〜' + ymdLabel_(end);
}

// 週シート名 261019 → '2026-10-19〜10-25'（名前順＝日付順になり、年が入るので来年と重ならない）
function weekFileLabel_(sheetName) {
  var start = '20' + sheetName.slice(0, 2) + '-' + sheetName.slice(2, 4) + '-' + sheetName.slice(4, 6);
  var end = addDaysYmd_(start, 6);
  return start + '〜' + end.slice(5);
}

function publishedWeekList_(settings) {
  return String(settings.getRange(PUBLISHED_WEEKS_CELL).getDisplayValue())
    .split('|').map(function(x) { return x.trim(); }).filter(function(x) { return /^\d{6}$/.test(x); });
}

function markPublished_(settings, sheetName) {
  var list = publishedWeekList_(settings).filter(function(x) { return x !== sheetName; });
  list.push(sheetName);
  settings.getRange(PUBLISHED_WEEKS_CELL).setValue(list.slice(-PUBLISHED_KEEP).join('|'));
  flushStaffPageCache_();     // スタッフのページが覚えている古い内容を消す
}

// PDFの入れ物フォルダ。シフト表と同じ場所の「シフトPDF」の中に、年ごとのフォルダを作る
// （1年で52枚たまるため。「シフトPDF」のIDだけ覚えておく）
function pdfFolder_(settings, year) {
  var top = null;
  var id = String(settings.getRange(PDF_FOLDER_CELL).getDisplayValue()).trim();
  if (id) {
    try { top = DriveApp.getFolderById(id); } catch (e) {}
  }
  if (!top) {
    var parents = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId()).getParents();
    var parent = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
    var found = parent.getFoldersByName(PDF_FOLDER_NAME);
    top = found.hasNext() ? found.next() : parent.createFolder(PDF_FOLDER_NAME);
    settings.getRange(PDF_FOLDER_CELL).setValue(top.getId());
  }
  var yearName = String(year) + '年';
  var sub = top.getFoldersByName(yearName);
  return sub.hasNext() ? sub.next() : top.createFolder(yearName);
}

// 公開したら、スタッフのページが覚えている内容を消させる（古いシフトが出ないように）。
// つながらなくても公開そのものは止めない
function flushStaffPageCache_() {
  try {
    var url = requestFormUrl_();
    if (!url) return false;
    var res = UrlFetchApp.fetch(url + (url.indexOf('?') >= 0 ? '&' : '?') + 'flush=1',
      { muteHttpExceptions: true, followRedirects: true });
    return res.getResponseCode() === 200;
  } catch (e) {
    Logger.log('flushStaffPageCache_: ' + e.message);
    return false;
  }
}

// 週シートの決まった範囲をPDFにする
function weekPdfBlob_(week, fileName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export' +
    '?format=pdf&gid=' + week.getSheetId() +
    '&range=' + encodeURIComponent(PDF_RANGE) +
    '&portrait=false&scale=4&gridlines=false&printtitle=false&sheetnames=false' +
    '&pagenum=UNDEFINED&fzr=false&printnotes=false&top_margin=0.85&bottom_margin=0.2&left_margin=0.2&right_margin=0.2';
  var res = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) {
    throw new Error('PDFを作れませんでした（' + res.getResponseCode() + '）');
  }
  return res.getBlob().setName(fileName);
}

// 同じ名前のPDFがあれば中身を入れ替える（リンクが変わらないようにするため）
function savePdf_(folder, blob) {
  var files = folder.getFilesByName(blob.getName());
  if (files.hasNext()) {
    var file = files.next();
    var res = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files/' + file.getId() + '?uploadType=media', {
      method: 'patch',
      contentType: 'application/pdf',
      payload: blob.getBytes(),
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });
    if (res.getResponseCode() !== 200) throw new Error('PDFを入れ替えられませんでした（' + res.getResponseCode() + '）');
    return file;
  }
  var made = folder.createFile(blob);
  made.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return made;
}

// お知らせの文
// LINEに送るお知らせ。長いアドレスをそのまま並べると読みにくいので、
// 「自分の分だけ見る」「全員の表（PDF）」という2つのボタンが付いたカード（Flex Message）で送る。
// altText ＝ 通知や、カードを出せない古いLINEで代わりに出る文
function publishNoticeMessage_(week, pageUrl, pdfUrl) {
  var range = weekRangeLabel_(week.getName());
  return {
    type: 'flex',
    altText: range + 'のシフトができました',
    contents: {
      type: 'bubble',
      // 何通も並んだときに、どの週のぶんか一目で分かるよう、日にちを大きくして1行にまとめる
      body: {
        type: 'box',
        layout: 'vertical',
        contents: [
          { type: 'text', wrap: true, contents: [
            { type: 'span', text: range, weight: 'bold', size: 'lg', color: '#111111' },
            { type: 'span', text: ' のシフト', size: 'sm', color: '#666666' }
          ] }
        ]
      },
      footer: {
        type: 'box',
        layout: 'vertical',
        spacing: 'sm',
        contents: [
          { type: 'button', style: 'primary', height: 'sm',
            action: { type: 'uri', label: '自分の分だけ見る', uri: pageUrl } },
          { type: 'button', style: 'secondary', height: 'sm',
            action: { type: 'uri', label: '全員の表（PDF）', uri: pdfUrl } }
        ]
      }
    }
  };
}

// メニュー用：お試し。お試しの送り先を登録してあれば、そこへ本物と同じお知らせを送る。
// 登録していなければ、今までどおりLINEには何も送らない
function publishWeekTest() {
  var week = activeWeekSheet_();
  if (!week) return;
  var ui = SpreadsheetApp.getUi();
  var settings = getRequestSettingsSheet_();
  var testTarget = String(settings.getRange(LINE_TEST_TARGET_CELL).getDisplayValue()).trim();
  if (ui.alert('「' + week.getName() + '」のPDFを作って、スタッフのページに公開します。\n\n' +
      (testTarget ? 'LINEは「お試しの送り先」にだけ送ります。スタッフのグループには送りません。'
                  : 'LINEへのお知らせは送りません。\n'
                    + '（お試しの送り先を登録しておくと、そこへ本物と同じお知らせを送って見え方を確かめられます）') +
      '\n\nよろしいですか。',
      ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  publishWeek_(week, true, !testTarget, true);
}

// 公開する本体。resend＝true なら、もう一度LINEを送る。skipLine＝true ならLINEを送らない
function publishWeek_(week, resend, skipLine, useTestTarget) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settings = getRequestSettingsSheet_();
  var name = week.getName();
  var storeName = String(settings.getRange(REQUEST_STORE_NAME_CELL).getDisplayValue()).trim() || 'シフト';
  var fileName = storeName + ' シフト ' + weekFileLabel_(name) + '.pdf';

  var file;
  try {
    file = savePdf_(pdfFolder_(settings, '20' + name.slice(0, 2)), weekPdfBlob_(week, fileName));
  } catch (e) {
    ss.toast('PDFを作れませんでした：' + e.message, 'シフト表', 8);
    markPublished_(settings, name);   // ページとカレンダーには出す
    return;
  }
  markPublished_(settings, name);

  var pdfUrl = file.getUrl();
  var pageUrl = STAFF_PAGE_URL + '?shift=' + name;
  var sent = '';
  var key = 'shift/' + name;
  if (skipLine) {
    sent = 'お試しのため、LINEへのお知らせは送っていません。';
  } else if (resend || !noticeAlreadySent_(settings, key)) {
    var line = lineConfig_(settings, useTestTarget);
    if (!line.token || !line.target) {
      sent = 'LINEの設定がないため、お知らせは送っていません。';
    } else {
      try {
        linePush_(line, publishNoticeMessage_(week, pageUrl, pdfUrl));
        // お試しのときは「送った控え」を残さない（本番の確定でちゃんと送れるように）
        if (!resend && !useTestTarget) writeNoticeLog_(settings, key);
        sent = useTestTarget ? 'お試しの送り先に、本物と同じお知らせを送りました（スタッフのグループには送っていません）。'
                             : 'LINEのグループにお知らせを送りました。';
      } catch (e) {
        sent = 'LINEへ送れませんでした：' + e.message;
        noticeSendFailure_(settings, name + 'のシフト公開のお知らせ', e.message);
      }
    }
  } else {
    sent = 'この週のお知らせは送信済みです（送り直すときはメニューの「もう一度知らせる」）。';
  }
  showPublishResult_(week, pageUrl, pdfUrl, sent);
}

// 結果を出して、印刷用にPDFを開けるようにする
function showPublishResult_(week, pageUrl, pdfUrl, sent) {
  var html = '<div style="font-family:sans-serif;padding:6px 10px;font-size:14px;line-height:1.8;">' +
    '<div style="font-weight:bold;margin-bottom:6px;">' + weekRangeLabel_(week.getName()) + ' を公開しました</div>' +
    '<div>' + sent + '</div>' +
    '<div style="margin-top:10px;">' +
    '<a href="' + pdfUrl + '" target="_blank" style="display:inline-block;background:#1a73e8;color:#fff;' +
    'padding:8px 14px;border-radius:6px;text-decoration:none;font-weight:bold;">PDFを開いて印刷する</a></div>' +
    // スタッフのページのリンクは、そのまま押すとブラウザがアカウント番号を足して開けないことがあるため、
    // コピーして貼り付けられるようにする
    '<div style="margin-top:12px;font-size:12px;color:#555;">スタッフのページ（コピーして開いてください）</div>' +
    '<div style="display:flex;gap:6px;margin-top:4px;">' +
    '<input id="pageUrl" readonly value="' + pageUrl + '" style="flex:1;font-size:12px;padding:6px;">' +
    '<button onclick="var b=document.getElementById(\'pageUrl\');b.select();document.execCommand(\'copy\');' +
    'this.textContent=\'コピーしました\';" style="font-size:12px;padding:6px 10px;">コピー</button></div>' +
    '</div>';
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(460).setHeight(280), 'シフトの公開');
}

// ===================== 決まりの見張り（未成年・休憩） =====================
// 守る決まり（労働基準法）
//  ・18歳未満（年度で見る）と高校生 … 22時〜5時は働けない／1日8時間・週40時間まで（労基法60条・61条）
//  ・全員 … 実働6時間超で45分以上、8時間超で60分以上の休憩（労基法34条）
// 年度で見るのは、高校生の学年に合わせるため。4月1日の時点で18歳未満なら、その年度いっぱい対象にする。
var MINOR_NIGHT_START = 2200;      // これ以降の退勤は不可
var MINOR_NIGHT_END = 500;         // これより前の出勤は不可
var MINOR_DAY_HOURS = 8;           // 1日の実働の上限
var MINOR_WEEK_HOURS = 40;         // 1週の実働の上限（年少者は44時間の特例が使えない）
var BREAK_RULES = [{ over: 8, need: 60 }, { over: 6, need: 45 }];   // 実働◯時間超なら休憩◯分
// 決まりに反する人の名前を塗る色。2段階にする
//  濃い赤＝時間が長い・遅い（年少者の夜間/5時前/1日8時間超/1週40時間超、実働8時間超で休憩60分が足りない）
//  うすい赤＝休憩のつけ方（6時間超で45分が足りない、休憩が勤務時間の外）
var LAW_WARN_COLOR = '#FFC7CE';    // うすい赤
var LAW_HARD_COLOR = '#C00000';    // 濃い赤
var LAW_HARD_FONT = '#FFFFFF';     // 濃い赤のときの文字色（白）
var ROSTER_URL_CELL = 'E28';       // シフト希望設定：アルバイト名簿ファイルのURL（空ならスタッフ区分シートの入力を使う）
var ROLE_BIRTH_COL = 3;            // スタッフ区分：C列＝生年月日
var ROLE_SCHOOL_COL = 4;           // スタッフ区分：D列＝高校生（〇）
var MINOR_CACHE_KEY = 'MINOR_INFO_V1';
var MINOR_CACHE_SEC = 1800;        // 名簿の読み直しは30分に1回でよい

// 名前のゆれ（全角・半角の空白）をなくす
function flatName_(name) {
  return String(name == null ? '' : name).replace(/[\s　]+/g, '');
}

// その日が属する年度（4月1日はじまり）
function fiscalYearOf_(date) {
  return date.getMonth() + 1 >= 4 ? date.getFullYear() : date.getFullYear() - 1;
}

// 年度の4月1日の時点で18歳未満か（＝その年度いっぱいは高校生と同じ扱いにする）
function isMinorOn_(birth, date) {
  if (!(birth instanceof Date)) return false;
  var april1 = new Date(fiscalYearOf_(date), 3, 1);
  var age = april1.getFullYear() - birth.getFullYear();
  var beforeBirthday = (april1.getMonth() < birth.getMonth()) ||
    (april1.getMonth() === birth.getMonth() && april1.getDate() < birth.getDate());
  if (beforeBirthday) age--;
  return age < 18;
}

// 名前ごとの { birth, highSchool } を集める。名簿ファイルがあればそれを優先し、
// 無い人はシフト表の「スタッフ区分」シートに入れた分を使う（名簿が無くても単体で使えるようにするため）
function minorInfoMap_(useCache) {
  var cache = null;
  if (useCache !== false) {
    try { cache = CacheService.getScriptCache(); } catch (e) { cache = null; }
    if (cache) {
      var hit = cache.get(MINOR_CACHE_KEY);
      if (hit) return reviveMinorMap_(JSON.parse(hit));
    }
  }
  var map = readRoleSheetBirths_();
  var fromRoster = readRosterBirths_();
  // 名簿のほうを先に使うが、名簿で空いているところは「スタッフ区分」シートに入れた分を残す。
  // （名簿に生年月日が入っていない人の分まで消してしまうと、入れたはずの生年月日が未登録あつかいになる）
  Object.keys(fromRoster).forEach(function(key) {
    var now = map[key], add = fromRoster[key];
    if (!now) { map[key] = add; return; }
    map[key] = {
      name: add.name || now.name,
      birth: add.birth || now.birth,
      highSchool: add.highSchool || now.highSchool,
      kind: add.kind || now.kind,
      wage: add.wage || now.wage
    };
  });
  if (cache) {
    try {
      cache.put(MINOR_CACHE_KEY, JSON.stringify(dryMinorMap_(map)), MINOR_CACHE_SEC);
    } catch (e) {}
  }
  return map;
}

function dryMinorMap_(map) {
  var out = {};
  Object.keys(map).forEach(function(key) {
    out[key] = { name: map[key].name, birth: map[key].birth ? map[key].birth.getTime() : 0, highSchool: map[key].highSchool };
  });
  return out;
}

function reviveMinorMap_(plain) {
  var out = {};
  Object.keys(plain).forEach(function(key) {
    out[key] = { name: plain[key].name, birth: plain[key].birth ? new Date(plain[key].birth) : null, highSchool: plain[key].highSchool };
  });
  return out;
}

// シフト表の「スタッフ区分」シート（A名前・B区分・C生年月日・D高校生）
function readRoleSheetBirths_() {
  var sheet = getRoleSheet_();
  var map = {};
  var last = sheet.getLastRow();
  if (last < 2) return map;
  sheet.getRange(2, 1, last - 1, 4).getValues().forEach(function(r) {
    var key = flatName_(r[0]);
    if (!key) return;
    map[key] = {
      name: String(r[0]).trim(),
      birth: (r[ROLE_BIRTH_COL - 1] instanceof Date) ? r[ROLE_BIRTH_COL - 1] : null,
      highSchool: String(r[ROLE_SCHOOL_COL - 1]).trim() !== '' || String(r[1]).indexOf('高校') === 0
    };
  });
  return map;
}

// アルバイト名簿ファイル（「名前」「生年月日」「区分」の見出しがある表を探す）
// 名簿ファイルを開く。設定にURLがあればそれを、無ければシフト表と同じフォルダから
// 名前に「名簿」が入っている表を探す（他のお店にフォルダごと渡しても動くように）
function rosterBook_() {
  var url = String(getRequestSettingsSheet_().getRange(ROSTER_URL_CELL).getDisplayValue()).trim();
  if (url) {
    try { return openBookByUrl_(url); } catch (e) { return null; }
  }
  try {
    var parents = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId()).getParents();
    while (parents.hasNext()) {
      var files = parents.next().getFilesByType(MimeType.GOOGLE_SHEETS);
      var other = null;
      while (files.hasNext()) {
        var file = files.next();
        var name = file.getName();
        if (name.indexOf('アルバイト名簿') !== -1) return SpreadsheetApp.openById(file.getId());
        if (!other && name.indexOf('名簿') !== -1) other = file;
      }
      if (other) return SpreadsheetApp.openById(other.getId());
    }
  } catch (e) {}
  return null;
}

// 名簿のシートを見る順番。「在籍」の表を先に見る（退職した人と混ざらないように）
function rosterSheetsInOrder_(book) {
  var sheets = book.getSheets();
  var first = sheets.filter(function(s) { return s.getName().indexOf('在籍') !== -1; });
  return first.concat(sheets.filter(function(s) { return first.indexOf(s) === -1; }));
}

// 「\1,220」のような書き方からも数字だけ取り出す
function wageNumber_(value) {
  var n = Number(String(value == null ? '' : value).replace(/[^0-9.]/g, ''));
  return n > 0 ? Math.round(n) : 0;
}

function readRosterBirths_() {
  var map = {};
  var book = rosterBook_();
  if (!book) return map;
  rosterSheetsInOrder_(book).forEach(function(sheet) {
    if (Object.keys(map).length) return;                 // 最初に見つかった表だけ使う
    var last = sheet.getLastRow(), lastCol = sheet.getLastColumn();
    if (last < 2 || lastCol < 2) return;
    var values = sheet.getRange(1, 1, Math.min(last, 200), lastCol).getValues();
    for (var h = 0; h < Math.min(values.length, 5); h++) {
      var head = values[h].map(function(v) { return String(v).trim(); });
      var nameCol = head.indexOf('名前');
      var birthCol = head.indexOf('生年月日');
      if (nameCol < 0 || birthCol < 0) continue;
      var kindCol = head.indexOf('区分');
      var wageCol = head.indexOf('時給');
      var leftCol = head.indexOf('退社日');
      for (var r = h + 1; r < values.length; r++) {
        var key = flatName_(values[r][nameCol]);
        if (!key) continue;
        if (leftCol >= 0 && values[r][leftCol] instanceof Date) continue;     // 辞めた人は入れない
        var birth = values[r][birthCol];
        map[key] = {
          name: String(values[r][nameCol]).trim(),
          birth: (birth instanceof Date) ? birth : null,
          highSchool: kindCol >= 0 && String(values[r][kindCol]).indexOf('高校') === 0,
          kind: kindCol >= 0 ? String(values[r][kindCol]).trim() : '',
          wage: wageCol >= 0 ? wageNumber_(values[r][wageCol]) : 0
        };
      }
      return;
    }
  });
  return map;
}

function openBookByUrl_(url) {
  var id = /[-\w]{25,}/.exec(String(url));
  return id ? SpreadsheetApp.openById(id[0]) : null;
}

// 「9:30」を分に直す（30分きざみの設定を比べるのに使う）
function stepMinutes_(t) {
  var m = /^(\d{1,2}):(\d{2})$/.exec(String(t));
  return m ? Number(m[1]) * 60 + Number(m[2]) : -1;
}

// 時間の数字（900・2130）を分に直す
function minutesOfTime_(value) {
  var n = Number(value);
  if (!isFinite(n) || n <= 0) return null;
  var h = Math.floor(n / 100), m = n % 100;
  if (m >= 60 || h > 29) return null;
  return h * 60 + m;
}

// 1人・1日の決まり破りを調べる。戻り値は理由の配列（空なら問題なし）
// 返すのは { list: 理由のことば, hard: 濃い赤にするか }
function dayViolations_(person, info, date) {
  var reasons = [], hard = false;
  var start = minutesOfTime_(person.start), end = minutesOfTime_(person.end);
  if (start === null || end === null) return { list: reasons, hard: hard };
  if (end <= start) return { list: reasons, hard: hard };   // 書きかけは見ない
  var bStart = minutesOfTime_(person.breakStart), bEnd = minutesOfTime_(person.breakEnd);
  var breakMin = (bStart !== null && bEnd !== null && bEnd > bStart) ? bEnd - bStart : 0;
  var work = end - start - breakMin;

  var minor = !!info && (info.highSchool || isMinorOn_(info.birth, date));
  if (minor) {
    if (end > minutesOfTime_(MINOR_NIGHT_START)) { reasons.push('22時以降の勤務'); hard = true; }
    if (start < minutesOfTime_(MINOR_NIGHT_END)) { reasons.push('5時前の勤務'); hard = true; }
    if (work > MINOR_DAY_HOURS * 60) {
      reasons.push('1日8時間を超える勤務（' + (Math.round(work / 6) / 10) + '時間）');
      hard = true;
    }
  }
  for (var i = 0; i < BREAK_RULES.length; i++) {
    if (work > BREAK_RULES[i].over * 60) {
      if (breakMin < BREAK_RULES[i].need) {
        reasons.push('実働' + (Math.round(work / 6) / 10) + '時間に対して休憩' + breakMin + '分（' + BREAK_RULES[i].need + '分必要）');
        if (BREAK_RULES[i].over >= 8) hard = true;        // 8時間を超える日は濃い赤
      }
      break;
    }
  }
  if (breakMin > 0 && bStart !== null && (bStart < start || bEnd > end)) {
    reasons.push('休憩が勤務時間の外にある');
  }
  return { list: reasons, hard: hard };
}

// この表の時間は「900」のような数字で持つ決まり（minutesOfTime_ は数字しか読めない）。
// 「9:00」と打って本物の時刻で入っているセルは、画面の見え方から数字に直して読む。
// これをしないと、その行は決まりの点検から丸ごと外れ、8時間を超えていても赤くならない。
// MARKER_LAW_TIME_READ
function timeNumber_(value, shownText) {
  if (!(value instanceof Date)) return value;
  var m = /^(\d{1,2}):(\d{2})/.exec(String(shownText == null ? '' : shownText).trim());
  return m ? Number(m[1]) * 100 + Number(m[2]) : '';
}

// 週シートの1日ぶん（日ブロック）の人を読む
function blockPeople_(week, blockIndex) {
  var block = ADJUST_BLOCKS[blockIndex];
  var range = week.getRange(block.head + 1, block.col - 5, ADJUST_ROWS, 5);
  var values = range.getValues(), shown = range.getDisplayValues();
  return values.map(function(r, p) {
    return { row: block.head + 1 + p, name: String(r[0]).trim(),
      start: timeNumber_(r[1], shown[p][1]), end: timeNumber_(r[2], shown[p][2]),
      breakStart: timeNumber_(r[3], shown[p][3]), breakEnd: timeNumber_(r[4], shown[p][4]) };
  });
}

// 週シートの日付（月曜）
function weekMondayOf_(week) {
  var name = week.getName();
  if (/^\d{6}$/.test(name)) {
    return new Date(2000 + Number(name.slice(0, 2)), Number(name.slice(2, 4)) - 1, Number(name.slice(4, 6)));
  }
  var t1 = week.getRange('T1').getValue();
  return (t1 instanceof Date) ? t1 : new Date();
}

// 週ぜんぶを調べる。見つけた分の名前を塗り、理由をメモに入れる
function checkWeekLaw_(week, paint) {
  var info = minorInfoMap_();
  var monday = weekMondayOf_(week);
  var found = [];
  var weekWork = {};
  ADJUST_BLOCKS.forEach(function(block, day) {
    var date = new Date(monday.getFullYear(), monday.getMonth(), monday.getDate() + day);
    var people = blockPeople_(week, day);
    people.forEach(function(person) {
      // 名前を消した行は、前に付いた色とメモを消してから次へ（消し忘れて赤が残らないように）
      if (!person.name) { if (paint) paintLawWarning_(week, block, person.row, [], false); return; }
      var who = info[flatName_(person.name)] || null;
      var v = dayViolations_(person, who, date);
      var reasons = v.list;
      var start = minutesOfTime_(person.start), end = minutesOfTime_(person.end);
      if (start !== null && end !== null && end > start) {
        var bs = minutesOfTime_(person.breakStart), be = minutesOfTime_(person.breakEnd);
        var brk = (bs !== null && be !== null && be > bs) ? be - bs : 0;
        var key = flatName_(person.name);
        weekWork[key] = (weekWork[key] || 0) + (end - start - brk);
      }
      if (paint) paintLawWarning_(week, block, person.row, reasons, v.hard);
      if (reasons.length) {
        found.push({ day: day, name: person.name, date: date, reasons: reasons, hard: v.hard });
      }
    });
  });
  // 週の合計（未成年・高校生だけ）
  Object.keys(weekWork).forEach(function(key) {
    var who = info[key];
    var minor = who && (who.highSchool || isMinorOn_(who.birth, monday));
    if (minor && weekWork[key] > MINOR_WEEK_HOURS * 60) {
      found.push({ day: -1, name: who.name, date: monday, hard: true,
        reasons: ['1週40時間を超える勤務（' + (Math.round(weekWork[key] / 6) / 10) + '時間）'] });
    }
  });
  return found;
}

// この色は、決まりの見張りが塗ったものか
function isLawColor_(bg) {
  var c = String(bg || '').toUpperCase();
  return c === LAW_WARN_COLOR.toUpperCase() || c === LAW_HARD_COLOR.toUpperCase();
}

// 名前のマスに残った「決まりの確認」の赤とメモを、まとめて消す。
// 週シートをコピーすると、名前と時間（T24:X39など）は消えるのに、名前のマスの色とメモは
// 前の週のまま残ってしまう。これが「使っていないはずの赤が残る」の正体。MARKER_LAW_MARK_CLEAR
function clearLawMarks_(week) {
  var cleared = 0;
  ADJUST_BLOCKS.forEach(function(block) {
    var top = block.head + 1, col = block.col - 5;
    var range = week.getRange(top, col, ADJUST_ROWS, 1);
    var bgs = range.getBackgrounds(), notes = range.getNotes();
    for (var i = 0; i < ADJUST_ROWS; i++) {
      if (!isLawColor_(bgs[i][0]) && String(notes[i][0]).indexOf('決まりの確認') !== 0) continue;
      week.getRange(top + i, col).setBackground(null).setFontColor(null).setFontWeight(null).setNote('');
      cleared++;
    }
  });
  return cleared;
}

// 自動で動く処理は、1つ失敗しても次を続ける。失敗は控えシートの隅に残す
function autoStep_(what, run) {
  try { run(); return true; } catch (err) { noteAutoFailure_(what, err); return false; }
}

function noteAutoFailure_(what, err) {
  try {
    var text = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'MM/dd HH:mm') + '　' + what + '　' +
      (err && err.message ? err.message : String(err));
    Logger.log('自動処理の失敗: ' + text);
    var sheet = getAdjustSheet_();
    sheet.getRange(AUTO_FAIL_LABEL_CELL).setValue('自動処理の最後の失敗');
    sheet.getRange(AUTO_FAIL_CELL).setValue(text);
  } catch (e2) {}
}

// 名前のマスに色とメモを付ける／消す
function paintLawWarning_(week, block, row, reasons, hard) {
  var cell = week.getRange(row, block.col - 5);
  if (reasons.length) {
    cell.setBackground(hard ? LAW_HARD_COLOR : LAW_WARN_COLOR);
    cell.setFontColor(hard ? LAW_HARD_FONT : null);
    cell.setFontWeight(hard ? 'bold' : null);
    cell.setNote('決まりの確認：\n・' + reasons.join('\n・'));
  } else if (String(cell.getNote()).indexOf('決まりの確認') === 0 || isLawColor_(cell.getBackground())) {
    // メモを手で消してあっても、この色が残っていれば消す
    cell.setBackground(null);
    cell.setFontColor(null);
    cell.setFontWeight(null);
    cell.setNote('');
  }
}

// 入力したその場で調べる（onEditから呼ぶ。直した行のまわりだけ）
function checkEditedLaw_(week, range) {
  var r0 = range.getRow(), c0 = range.getColumn();
  var r1 = r0 + range.getNumRows() - 1;
  var c1 = c0 + range.getNumColumns() - 1;
  var info = null, monday = null, messages = [];
  ADJUST_BLOCKS.forEach(function(block, day) {
    if (c1 < block.col - 5 || c0 > block.col - 1) return;      // 名前〜休終の列に当たっていない
    if (!info) { info = minorInfoMap_(); monday = weekMondayOf_(week); }
    var date = new Date(monday.getFullYear(), monday.getMonth(), monday.getDate() + day);
    blockPeople_(week, day).forEach(function(person) {
      if (person.row < r0 || person.row > r1) return;
      // 名前を消した行は、前に付いた色とメモも消す
      if (!person.name) { paintLawWarning_(week, block, person.row, [], false); return; }
      var v = dayViolations_(person, info[flatName_(person.name)] || null, date);
      paintLawWarning_(week, block, person.row, v.list, v.hard);
      if (v.list.length) messages.push(DOW_JP[date.getDay()] + '曜 ' + person.name + '：' + v.list[0]);
    });
  });
  if (messages.length) {
    SpreadsheetApp.getActiveSpreadsheet().toast(messages.join(' ／ '), '決まりの確認', 8);
  }
  return messages.length;
}

// メニュー「決まりを点検する」
function checkWeekLawNow() {
  syncRoleSheetBirths_();                    // 名簿の生年月日をスタッフ区分シートにも入れておく（週シートでなくても先にやる）
  var week = activeWeekSheet_();
  if (!week) return;
  var found = checkWeekLaw_(week, true);
  var unknown = unknownBirthNames_(week);
  if (!found.length && !unknown.length) {
    SpreadsheetApp.getActiveSpreadsheet().toast('決まりに反するところはありませんでした。', '決まりの確認', 6);
    return;
  }
  SpreadsheetApp.getUi().alert('決まりの確認（' + week.getName() + '）', lawReportText_(found, unknown), SpreadsheetApp.getUi().ButtonSet.OK);
}

// 名簿から読めた生年月日・高校生を「スタッフ区分」シートにも書き写す。
// 仕組みは名簿を直接読むので無くても動くが、画面で見えないと確かめられないため。
// 生年月日は空いているマスだけ入れる（手で直した分は消さない）。高校生の印は名簿に合わせる。
function syncRoleSheetBirths_() {
  var fromRoster = readRosterBirths_();
  if (!Object.keys(fromRoster).length) return 0;
  var sheet = getRoleSheet_();
  var last = sheet.getLastRow();
  if (last < 2) return 0;
  var rows = sheet.getRange(2, 1, last - 1, ROLE_SCHOOL_COL).getValues();
  var changed = 0;
  for (var i = 0; i < rows.length; i++) {
    var who = fromRoster[flatName_(rows[i][0])];
    if (!who) continue;
    if (who.birth && !(rows[i][ROLE_BIRTH_COL - 1] instanceof Date)) {
      sheet.getRange(i + 2, ROLE_BIRTH_COL).setValue(who.birth);
      changed++;
    }
    if (who.kind) {                                  // 名簿に区分があるときだけ印を合わせる
      var mark = who.highSchool ? '〇' : '';
      if (String(rows[i][ROLE_SCHOOL_COL - 1]).trim() !== mark) {
        sheet.getRange(i + 2, ROLE_SCHOOL_COL).setValue(mark);
        changed++;
      }
    }
  }
  if (changed) { try { CacheService.getScriptCache().remove(MINOR_CACHE_KEY); } catch (e) {} }
  return changed;
}

// この週に出ている人のうち、生年月日が分からない人（＝未成年かどうか調べられない人）。
// 店長・社員は大人と分かっているので数えない。
function unknownBirthNames_(week) {
  var info = minorInfoMap_();
  var roles = {};
  var rolesRaw = getStaffRoles_();
  Object.keys(rolesRaw).forEach(function(n) { if (String(rolesRaw[n]).trim()) roles[flatName_(n)] = true; });
  var seen = {}, out = [];
  ADJUST_BLOCKS.forEach(function(block, day) {
    blockPeople_(week, day).forEach(function(person) {
      var key = flatName_(person.name);
      if (!key || seen[key] || roles[key]) return;
      if (minutesOfTime_(person.start) === null || minutesOfTime_(person.end) === null) return;   // 出勤の無い行は見ない
      var who = info[key];
      if (who && (who.birth || who.highSchool)) return;
      seen[key] = true;
      out.push(person.name);
    });
  });
  return out;
}

function lawReportText_(found, unknown) {
  var lines = found.map(function(x) {
    var day = x.day >= 0 ? (x.date.getMonth() + 1) + '/' + x.date.getDate() + '(' + DOW_JP[x.date.getDay()] + ') ' : '週の合計 ';
    return '・' + day + x.name + '：' + x.reasons.join('／');
  });
  if (unknown && unknown.length) {
    if (lines.length) lines.push('');
    lines.push('生年月日が未登録のため、未成年かどうか調べられませんでした。');
    lines.push('・' + unknown.join('、'));
    lines.push('（スタッフ管理 → スタッフの生年月日 で入れられます。');
    lines.push('　名前の書き方がシフト表と違っていても、別の人として出ます）');
  }
  return lines.join('\n');
}

// メニュー「スタッフの生年月日」：名簿が無くても使えるようにする入力画面
function openMinorSettings() {
  syncRoleSheetBirths_();                    // 名簿から読めた分は先にシートへ入れておく
  var settings = getRequestSettingsSheet_();
  var rosterUrl = String(settings.getRange(ROSTER_URL_CELL).getDisplayValue()).trim();
  var info = minorInfoMap_(false);
  var names = [];
  var newest = newestWeekSheet_(SpreadsheetApp.getActiveSpreadsheet());
  if (newest) getRosterNames_(newest).forEach(function(n) { if (String(n).trim()) names.push(String(n).trim()); });
  Object.keys(info).forEach(function(key) {
    if (!names.some(function(n) { return flatName_(n) === key; })) names.push(info[key].name);
  });

  var rows = names.map(function(name) {
    var who = info[flatName_(name)] || {};
    var ymd = who.birth ? Utilities.formatDate(who.birth, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '';
    return '<tr><td style="padding:2px 6px;">' + name + '</td>' +
      '<td><input type="date" value="' + ymd + '" data-name="' + name + '" class="birth" style="font-size:13px;"></td>' +
      '<td style="text-align:center;"><input type="checkbox" class="hs" data-name="' + name + '"' +
      (who.highSchool ? ' checked' : '') + '></td></tr>';
  }).join('');

  var html = '<div style="font-family:sans-serif;font-size:13px;padding:6px 10px;">' +
    '<p style="margin:0 0 8px;color:#555;">18歳未満（年度で判定）と高校生は、22時以降の勤務と1日8時間を超える勤務ができません。<br>' +
    'アルバイト名簿のURLを入れておくと、生年月日は名簿から自動で読みます。</p>' +
    '<div style="margin-bottom:8px;">名簿ファイルのURL<br>' +
    '<input id="rosterUrl" value="' + rosterUrl + '" style="width:100%;font-size:12px;padding:4px;"></div>' +
    '<table style="border-collapse:collapse;"><tr><th style="text-align:left;">名前</th><th>生年月日</th><th>高校生</th></tr>' +
    rows + '</table>' +
    '<div style="margin-top:10px;"><button onclick="save()" style="padding:6px 14px;font-weight:bold;">保存する</button>' +
    '<span id="msg" style="margin-left:8px;color:#1a73e8;"></span></div>' +
    '<script>' +
    'function save(){' +
    'var list=[];' +
    'var b=document.getElementsByClassName("birth");' +
    'for(var i=0;i<b.length;i++){' +
    'var n=b[i].getAttribute("data-name");' +
    'var hs=document.querySelector(\'.hs[data-name="\'+n+\'"]\').checked;' +
    'list.push({name:n,birth:b[i].value,highSchool:hs});}' +
    'document.getElementById("msg").textContent="保存中…";' +
    'google.script.run.withSuccessHandler(function(){document.getElementById("msg").textContent="保存しました";})' +
    '.saveMinorSettings(document.getElementById("rosterUrl").value,list);}' +
    '<\/script></div>';
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(460).setHeight(520), 'スタッフの生年月日');
}

// 入力画面からの保存（スタッフ区分シートのC・D列に書く）
function saveMinorSettings(rosterUrl, list) {
  var settings = getRequestSettingsSheet_();
  settings.getRange(ROSTER_URL_CELL).setValue(String(rosterUrl || '').trim());
  var sheet = getRoleSheet_();
  sheet.getRange('C1:D1').setValues([['生年月日', '高校生']]);
  var last = sheet.getLastRow();
  var names = last >= 2 ? sheet.getRange(2, 1, last - 1, 1).getValues().map(function(r) { return flatName_(r[0]); }) : [];
  (list || []).forEach(function(item) {
    var key = flatName_(item.name);
    var idx = names.indexOf(key);
    var row;
    if (idx >= 0) {
      row = idx + 2;
    } else {
      row = sheet.getLastRow() + 1;
      sheet.getRange(row, 1).setValue(item.name);
      names.push(key);
    }
    sheet.getRange(row, ROLE_BIRTH_COL).setValue(item.birth ? new Date(item.birth + 'T00:00:00+09:00') : '');
    sheet.getRange(row, ROLE_SCHOOL_COL).setValue(item.highSchool ? '〇' : '');
  });
  try { CacheService.getScriptCache().remove(MINOR_CACHE_KEY); } catch (e) {}
  return true;
}

// ===================== ポジションの色分け =====================
// MARKER_POSITION_RUN
// 時間帯バーの水色は、もともと「値が0.25なら#00B0F0」という条件付き書式で付いていた。
// 条件付き書式は直接塗った色より強いので、そのままでは人がいるマスに色を塗っても見えない。
// 前は「色分けを始める／やめる」でこのルールを外したり戻したりしていたが、
//   ・戻すときにルールを作り直すので、元と同じにならない（やめても元に戻らない）
//   ・「色分け中かどうか」の見分け方が2か所で食い違う（塗っても色が変わらない）
// という不具合が続いたので、モードそのものを無くした。
// いまは週シートからこのルールを外しておき、人がいるマスの水色も同じ色でGASが塗る。
// 見た目は今までと変わらず、いつでもポジションを塗れる。
// シートの条件付き書式に入っている水色（この色のルールを外す）
var POSITION_RULE_COLOR = '#00B0F0';
// 人がいて、まだポジションを決めていない時間の色。今までの自動の水色と同じ色にする
var POSITION_BASE_COLOR = '#00B0F0';
// ポジションの名前と色を入れておく行（E列＝名前・F列＝色）
// ※E26は「確定して公開した週」で埋まっているので、5つめだけ離れた48行目に置く
var POSITION_NAME_ROW = 22;                 // E22〜E25 名前、F22〜F25 色（1〜4）
var POSITION_ROWS = [22, 23, 24, 25, 48];
var POSITION_COUNT = POSITION_ROWS.length;
var POSITION_DEFAULTS = [
  { name: 'デシャップ',   color: '#E06666' },
  { name: 'キッチン',     color: '#4A86E8' },
  { name: 'ウォッシャー', color: '#A9D18E' },
  { name: '店長',         color: '#B4A7D6' },
  { name: '掃除',         color: '#FFD966' }
];

// 設定シートから、ポジションの行だけをまとめて読む（読みに行くのは1回）
function positionCells_(sheet) {
  var top = Math.min.apply(null, POSITION_ROWS);
  var bottom = Math.max.apply(null, POSITION_ROWS);
  var all = sheet.getRange(top, 5, bottom - top + 1, 2).getValues();
  return POSITION_ROWS.map(function(row) { return all[row - top]; });
}

// ポジションの名前と色を読む
function positionList_() {
  var settings = getRequestSettingsSheet_();
  return positionCells_(settings).map(function(r, i) {
    var name = String(r[0]).trim() || POSITION_DEFAULTS[i].name;
    var color = String(r[1]).trim();
    return { name: name, color: /^#[0-9A-Fa-f]{6}$/.test(color) ? color : POSITION_DEFAULTS[i].color };
  });
}

// その行・列が、時間帯バーの升目（人の行）の中かどうか
function inGridArea_(row, col) {
  for (var i = 0; i < ADJUST_BLOCKS.length; i++) {
    var b = ADJUST_BLOCKS[i];
    if (row > b.head && row <= b.head + ADJUST_ROWS &&
        col >= b.col && col < b.col + ADJUST_GRID_COLS) return true;
  }
  return false;
}

// このシートの「値が0.25なら水色」という自動の色づけを外しておく（まだ付いていれば）。
// 「色分け中かどうか」という状態は持たない。シートの中身だけを見て、必要なときに1回だけ外す。
// 外したあとは人がいるマスの水色もGASが同じ色で塗るので、見た目は変わらない。
function ensureGridPaintable_(week) {
  var keep = [], removed = 0;
  week.getConditionalFormatRules().forEach(function(rule) {
    if (isGridBlueRule_(rule)) { removed++; return; }
    keep.push(rule);
  });
  if (!removed) return false;
  week.setConditionalFormatRules(keep);
  repaintWeekColors_(week);
  return true;
}

// 「値が0.25なら水色」という、時間帯バーの自動の色づけかどうか
// （右はしの「計」の列にも同じルールが付いているが、計は0.25にならないので外しても見た目は変わらない）
function isGridBlueRule_(rule) {
  var bc = rule.getBooleanCondition();
  return !!(bc && String(bc.getCriteriaType()) === 'NUMBER_EQUAL_TO' &&
    Number(bc.getCriteriaValues()[0]) === 0.25 &&
    String(bc.getBackground()).toUpperCase() === POSITION_RULE_COLOR);
}

// その行・列が、日ブロックの「番号・名前」の列か
// （そこを選んだときは、その人の1日ぶんをまるごと指していることにする）
// 出勤・退勤・休憩の列は入れない。時間を直そうとクリックしただけで塗れてしまうため。
function personBlockAt_(row, col) {
  for (var i = 0; i < ADJUST_BLOCKS.length; i++) {
    var b = ADJUST_BLOCKS[i];
    if (row > b.head && row <= b.head + ADJUST_ROWS &&
        col >= b.col - 6 && col <= b.col - 5) return b;
  }
  return null;
}

// 決めた場所のうち、人がいるマスだけを塗る（ブロックごとにまとめて読み書きする）
// 選んだマスから右へ、人がいる間ずっと塗る。休憩や退勤でマスが空になるところで止まる。
// 　例）9:00〜18:00・休憩14:00〜15:00 で 9:00 を選ぶ → 9:00〜14:00 が塗られる
// 　　　同じ人の 15:00 を選ぶ → 15:00〜18:00 が塗られる
function paintCells_(week, targets, color) {
  var painted = 0;
  ADJUST_BLOCKS.forEach(function(block) {
    var seeds = [];
    for (var p = 0; p < ADJUST_ROWS; p++) {
      for (var c = 0; c < ADJUST_GRID_COLS; c++) {
        if (targets[(block.head + 1 + p) + ',' + (block.col + c)]) seeds.push([p, c]);
      }
    }
    if (!seeds.length) return;
    var range = week.getRange(block.head + 1, block.col, ADJUST_ROWS, ADJUST_GRID_COLS);
    var values = range.getValues();
    var colors = range.getBackgrounds();
    var done = {}, changed = 0;
    seeds.forEach(function(seed) {
      var p = seed[0];
      for (var c = seed[1]; c < ADJUST_GRID_COLS; c++) {
        if (!(Number(values[p][c]) > 0)) break;      // 休憩・退勤でおしまい
        if (done[p + ',' + c]) continue;             // もう塗ったマスは数えない
        done[p + ',' + c] = true;
        colors[p][c] = color;
        changed++;
      }
    });
    if (changed) { range.setBackgrounds(colors); painted += changed; }
  });
  return painted;
}

// 選んだ範囲のうち、塗りはじめる場所を決める
// ・時間帯バーの升目 … そのマス（そこから右へ、休憩の直前まで塗る）
// ・名前や番号の列   … その人の1日ぶん
function targetsOfRanges_(ranges) {
  var targets = {};
  ranges.forEach(function(range) {
    var r0 = range.getRow(), c0 = range.getColumn();
    var nr = range.getNumRows(), nc = range.getNumColumns();
    for (var i = 0; i < nr; i++) {
      for (var j = 0; j < nc; j++) {
        var row = r0 + i, col = c0 + j;
        if (inGridArea_(row, col)) { targets[row + ',' + col] = true; continue; }
        var block = personBlockAt_(row, col);
        if (block) {
          for (var c = 0; c < ADJUST_GRID_COLS; c++) targets[row + ',' + (block.col + c)] = true;
        }
      }
    }
  });
  return targets;
}

// 今えらんでいる範囲（飛び地に対応）
function activeRanges_(week) {
  var list = week.getActiveRangeList();
  if (list) return list.getRanges();
  var one = week.getActiveRange();
  return one ? [one] : [];
}

// ポジションのボタンを押したとき
// ・升目を1マス選んでいたら … そこから右へ、休憩の直前まで（休憩が無ければ退勤まで）塗る
// ・名前や番号を選んでいたら … その人の1日ぶんを塗る
// モードもペンも無い。押したときだけ塗るので、触っただけで塗れてしまうことはない。
function paintSelectedPosition_(color, label) {
  var week = activeWeekSheet_();
  if (!week) return;
  var book = SpreadsheetApp.getActiveSpreadsheet();
  ensureGridPaintable_(week);                  // 自動の水色ルールが残っていれば、ここで1回だけ外す
  var painted = paintCells_(week, targetsOfRanges_(activeRanges_(week)), color);
  if (!painted) {
    book.toast('塗るところが見つかりませんでした。\n' +
      '塗りはじめたい時間のマスを選んでから、もう一度押してください。', 'ポジション', 6);
    return;
  }
  book.toast(painted + 'マス（' + (Math.round(painted / 4 * 100) / 100) + '時間）を塗りました。\n' +
    '選んだマスから、休憩の直前まで（休憩が無ければ退勤まで）塗ります。', label, 6);
}

function paintPosition1() { var p = positionList_()[0]; paintSelectedPosition_(p.color, p.name); }
function paintPosition2() { var p = positionList_()[1]; paintSelectedPosition_(p.color, p.name); }
function paintPosition3() { var p = positionList_()[2]; paintSelectedPosition_(p.color, p.name); }
function paintPosition4() { var p = positionList_()[3]; paintSelectedPosition_(p.color, p.name); }
function paintPosition5() { var p = positionList_()[4]; paintSelectedPosition_(p.color, p.name); }
function clearPositionPaint() { paintSelectedPosition_(POSITION_BASE_COLOR, 'もとの水色'); }

// メニュー／ボタン「週ぜんぶ水色にもどす」
// 「水色にもどす」の、週まるごと版。戻る色はどちらも同じ水色
function clearAllPositionColors() {
  var week = activeWeekSheet_();
  if (!week) return;
  var ui = SpreadsheetApp.getUi();
  if (ui.alert('「' + week.getName() + '」で塗ったポジションの色を、週ぜんぶ水色にもどします。\n\n' +
      '（削った時間の薄いグレーは、そのまま残ります）\nよろしいですか。',
      ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  ensureGridPaintable_(week);
  clearOurGridColors_(week);
  repaintWeekColors_(week);
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '週ぜんぶを水色にもどしました。（削った時間の薄いグレーは残しています）', 'ポジション', 5);
}

// メニュー「ポジションの名前と色を変える」
function openPositionSettings() {
  var list = positionList_();
  var box = 'padding:6px;font-size:12px;';
  var rows = list.map(function(p, i) {
    return '<div style="display:flex;gap:8px;align-items:center;margin-top:8px;">' +
      '<input id="name' + i + '" value="' + escapeHtml_(p.name) + '" style="' + box + 'flex:1;">' +
      '<input id="color' + i + '" type="color" value="' + p.color + '" style="width:54px;height:32px;padding:0;border:none;">' +
      '</div>';
  }).join('');
  var html = '<div style="font-family:sans-serif;padding:4px 8px;font-size:14px;line-height:1.6;">' +
    '<p style="margin:0;font-size:12px;color:#555;">ポジションの名前と色です。メニューの表示は、次にスプレッドシートを開き直したときに変わります。</p>' +
    rows +
    '<p id="msg" style="margin:10px 0 0;font-size:12px;color:#186A3B;"></p>' +
    '<div style="margin-top:10px;text-align:right;">' +
    '<button id="saveBtn" onclick="doSave()" style="padding:6px 12px;">保存する</button></div></div>' +
    '<script>' +
    'function doSave(){' +
    '  var v = [];' +
    '  for (var i = 0; i < ' + POSITION_COUNT + '; i++) {' +
    '    v.push({ name: document.getElementById("name"+i).value, color: document.getElementById("color"+i).value });' +
    '  }' +
    '  document.getElementById("saveBtn").disabled = true;' +
    '  google.script.run' +
    '    .withSuccessHandler(function(t){ document.getElementById("msg").textContent = t; document.getElementById("saveBtn").disabled = false; })' +
    '    .withFailureHandler(function(e){ document.getElementById("msg").textContent = e.message; document.getElementById("saveBtn").disabled = false; })' +
    '    .savePositionSettings(v);' +
    '}' +
    '</script>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(380).setHeight(300), 'ポジションの名前と色');
}

function savePositionSettings(values) {
  var settings = getRequestSettingsSheet_();
  var rows = [];
  for (var i = 0; i < POSITION_COUNT; i++) {
    var v = values[i] || {};
    var name = String(v.name || '').trim();
    var color = String(v.color || '').trim();
    if (!name) throw new Error((i + 1) + '番目のポジションの名前を入れてください。');
    rows.push([name, /^#[0-9A-Fa-f]{6}$/.test(color) ? color.toUpperCase() : POSITION_DEFAULTS[i].color]);
  }
  POSITION_ROWS.forEach(function(row, i) { settings.getRange(row, 5, 1, 2).setValues([rows[i]]); });
  return '保存しました。メニューの表示は、開き直すと変わります。';
}



// ----- 色の塗り直し（ここ1か所で色を決める） -----
// マスの色は、いつでも次のルールで決まる。
//   今その人がいる … ポジションの色があればそのまま／無ければ水色
//   今はいない     … 調整前にはいた → 薄いグレー（元の色は控えるので、時間を戻せば返ってくる）
//                    調整前もいない → 白（ただし自分で付けた色は消さない）
function isSameColor_(a, b) { return String(a).toLowerCase() === String(b).toLowerCase(); }

// 仕組みが塗った色の一覧（これ以外＝自分で付けた色は、勝手に消さない）
function ourColors_() {
  var list = [ADJUST_COLOR, POSITION_BASE_COLOR, POSITION_RULE_COLOR];
  positionList_().forEach(function(p) { list.push(p.color); });
  return list;
}
function isOurColor_(color, list) {
  for (var i = 0; i < list.length; i++) if (isSameColor_(color, list[i])) return true;
  return false;
}

// 塗り直しに使う情報をまとめて読む（何度も読まないように）
function repaintContext_(week) {
  var sheet = getAdjustSheet_();
  var map = adjustRowMap_(sheet);
  return {
    sheet: sheet,
    map: map,
    ours: ourColors_()
  };
}

// 1つの曜日ブロックを塗り直す（persons を渡すとその人の行だけ）
function repaintBlock_(week, blockIndex, persons, ctx) {
  var block = ADJUST_BLOCKS[blockIndex];
  var recRow = ctx.map[week.getName() + '/' + blockIndex];
  var before = [], keepMap = {};
  if (recRow) {
    before = String(ctx.sheet.getRange(recRow, 4).getValue()).split('|');
    var text = String(ctx.sheet.getRange(recRow, 5).getValue());
    if (text) {
      try {
        JSON.parse(text).forEach(function(k) { keepMap[k[0] + ',' + k[1]] = k[2]; });
      } catch (e) { keepMap = {}; }
    }
  }
  var range = week.getRange(block.head + 1, block.col, ADJUST_ROWS, ADJUST_GRID_COLS);
  var values = range.getValues();
  var colors = range.getBackgrounds();
  var changed = false, keepChanged = false, grayed = 0, freed = 0;

  for (var p = 0; p < ADJUST_ROWS; p++) {
    if (persons && persons.indexOf(p) < 0) continue;
    var line = before[p] || '';
    for (var c = 0; c < ADJUST_GRID_COLS; c++) {
      var key = p + ',' + c;
      var color = String(colors[p][c]);
      if (Number(values[p][c]) > 0) {
        // 今その人がいる
        if (keepMap[key] !== undefined) {                 // 時間が戻った → 元の色を返す
          color = String(keepMap[key]);
          colors[p][c] = color; changed = true; freed++;
          delete keepMap[key]; keepChanged = true;
        }
        if (isSameColor_(color, ADJUST_COLOR) || isSameColor_(color, '#ffffff')) {
          colors[p][c] = POSITION_BASE_COLOR; changed = true;   // 戻った時間・増えた時間は水色にする
        }
      } else if (line.charAt(c) === '1') {
        // 今はいないが、調整前にはいた → 薄いグレー
        if (!isSameColor_(color, ADJUST_COLOR)) {
          keepMap[key] = color; keepChanged = true;
          colors[p][c] = ADJUST_COLOR; changed = true; grayed++;
        }
      } else {
        // もともといない
        if (keepMap[key] !== undefined) { delete keepMap[key]; keepChanged = true; }
        if (!isSameColor_(color, '#ffffff') && isOurColor_(color, ctx.ours)) {
          colors[p][c] = '#ffffff'; changed = true;
        }
      }
    }
  }
  if (changed) range.setBackgrounds(colors);
  if (recRow && keepChanged) {
    var list = [];
    Object.keys(keepMap).forEach(function(k) {
      var pc = k.split(',');
      list.push([Number(pc[0]), Number(pc[1]), keepMap[k]]);
    });
    ctx.sheet.getRange(recRow, 5).setValue(list.length ? JSON.stringify(list) : '');
  }
  return { grayed: grayed, freed: freed };
}

// 升目の「仕組みが塗った色」（薄い水色・濃い水色・ポジションの色・薄いグレー）を白に戻す
// 店長が自分で付けた色は残す
function clearOurGridColors_(week) {
  var ours = ourColors_();
  var cleared = 0;
  ADJUST_BLOCKS.forEach(function(block) {
    var range = week.getRange(block.head + 1, block.col, ADJUST_ROWS, ADJUST_GRID_COLS);
    var colors = range.getBackgrounds();
    var changed = false;
    for (var p = 0; p < ADJUST_ROWS; p++) {
      for (var c = 0; c < ADJUST_GRID_COLS; c++) {
        if (!isOurColor_(colors[p][c], ours)) continue;
        colors[p][c] = '#ffffff';
        changed = true;
        cleared++;
      }
    }
    if (changed) range.setBackgrounds(colors);
  });
  return cleared;
}

// 次の週シートは、シートごとコピーして作る。
// 先週のポジションの色を持ち越さないよう、新しいシートの色だけ消して塗り直す
// （今のシートの色はそのまま残す）。
function resetPositionColorsOnCopy_(newSheet) {
  ensureGridPaintable_(newSheet);      // 自動の水色ルールが残っていれば外す
  clearOurGridColors_(newSheet);       // 先週のポジションの色を消す
  clearLawMarks_(newSheet);            // 先週の「決まりの確認」の赤とメモも消す（名前は空になっているため）
  repaintWeekColors_(newSheet);        // いる時間を水色に塗り直す
  try { newSheet.setTabColor(null); } catch (e) {}
}

// 週シート全体を塗り直す
function repaintWeekColors_(week) {
  var ctx = repaintContext_(week);
  var total = { grayed: 0, freed: 0 };
  ADJUST_BLOCKS.forEach(function(block, i) {
    var r = repaintBlock_(week, i, null, ctx);
    total.grayed += r.grayed;
    total.freed += r.freed;
  });
  return total;
}

// 出勤・退勤・休憩を直したときに、その人の行だけ塗り直す（onEditから呼ぶ）
function repaintEditedRows_(week, range) {
  var r0 = range.getRow(), c0 = range.getColumn();
  var r1 = r0 + range.getNumRows() - 1;
  var c1 = c0 + range.getNumColumns() - 1;
  var ctx = null;
  ADJUST_BLOCKS.forEach(function(block, i) {
    if (c1 < block.col - 4 || c0 > block.col - 1) return;   // 出勤〜休終の列に当たっていない
    var persons = [];
    for (var p = 0; p < ADJUST_ROWS; p++) {
      var row = block.head + 1 + p;
      if (row >= r0 && row <= r1) persons.push(p);
    }
    if (!persons.length) return;
    if (!ctx) ctx = repaintContext_(week);
    repaintBlock_(week, i, persons, ctx);
  });
  return !!ctx;
}

// メニュー／ボタン「掃除（色を塗り直す）」
function refreshWeekColors() {
  var week = activeWeekSheet_();
  if (!week) return;
  ensureGridPaintable_(week);          // 自動の水色ルールが残っていれば、ここでも外す
  var t = repaintWeekColors_(week);
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '色を塗り直しました。削った時間 ' + t.grayed + 'マスを薄いグレーに、' +
    '戻った時間 ' + t.freed + 'マスを元の色にしました。', 'シフト表', 6);
}

// ----- 画面に置くボタン -----
// ポジションの色分けは「マスを選んでからボタンを押す」のが一番速い。
// 図形を手で作らなくてよいように、絵を持ったボタンを6個まとめて置く。
// 図形と同じで、押してもセルの選びは外れないので、選んだところがそのまま塗られる。
var POSITION_BUTTON_ROW = 19;                 // 置く場所（R19のあたり・18〜22行目は空き）
var POSITION_BUTTON_COL = 18;
var POSITION_BUTTON_W = 148;
var POSITION_BUTTON_H = 34;
var POSITION_BUTTON_GAP = 6;
var POSITION_BUTTON_PER_ROW = 4;              // 4個ずつの2段にする
var POSITION_BUTTON_ROW_GAP = 2;              // 2段目は2行下
// 1段目＝ポジション1〜4、2段目＝ポジション5（掃除）・水色にもどす・週ぜんぶ水色に・確定する
// ※「色を塗り直す」はメニュー（困ったとき・やり直し）に移した
var POSITION_BUTTON_FUNCS = ['paintPosition1', 'paintPosition2', 'paintPosition3', 'paintPosition4',
                             'paintPosition5', 'clearPositionPaint', 'clearAllPositionColors', 'confirmAdjust'];
// 前の版で置いていて、今は使わないボタン。置き直すときに一緒に外す
// （外さないと、古い「掃除（色を塗り直す）」のボタンが画面に残ってしまう）
var POSITION_BUTTON_RETIRED = ['refreshWeekColors'];
function isPositionButtonScript_(fn) {
  return POSITION_BUTTON_FUNCS.indexOf(fn) >= 0 || POSITION_BUTTON_RETIRED.indexOf(fn) >= 0;
}
var POSITION_BUTTON_IMAGES = [
  'iVBORw0KGgoAAAANSUhEUgAAAJQAAAAiCAYAAAC5r/C8AAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAALQSURBVHhe7ZyxTtxAFEX5tK1paalRGkpKSjqKNJS0tLShp+YHaBJFibKKoiiKIkVK4eisdKPJeGyPdyfxjLlPOmLXnrF3eWffGxu0R29OT7vXJyfGHAwuHfFgs9kYczC49EeoV8fHPeOMyQF3ekLx8+3ZmTGzCR2yUOZgLJQpioUyRbFQpiirF+rT9XX34+mpt70l3p2fdz+fn7tf22335e6ut78mViuURCIZxMerq96YVvh6f797PzxGrPcXF70xtdC0UGNBAghE+vbw0G1vbnZzSAafdLYpSbXz/fHRQv0PxkIifb69/WsOiSEpYSydoNxwy6sYJCI5SBfvG6petNC56zGERgRaL88/XF7untPGNCYn5p53KV6sUGMo4u0kNbU9BiFpUzxW61WlpHIS2j+Fxluof8yc0IKcRNPuSO5QmxtKONvZRlBdQBWL4LiqQpyH0DkICSHBclvX3PFL8+KECmOo3RFaxOsqMRW0LkTRmox2xpy4KnFMnQsxCB1/irnjl6ZZoVIo4u0xuqUQJymUR9VGx2W8ZKNShfsRC8JjaVx8blW5oQoZM3f80qxOKKpFvD0X2hgx1l4kl54jJRHLiQip4xCxfEMg0ZzxNbAaofTL33fxqrUTEVafmHBhroqWm3DJF17hjSHBY1lrZjVCsV6Zk6wQ1j6K1O2CECUZAZmHYFNzAPkQj8hpX5IvV9ZaWIVQ4dpHC+NctOglcipBvODOQX+LI1JtMEYfjn3ez9I0LxS/cCUrtzrppqYqBpFTZTSXSC24U4SvL3VVGULVU0slWpMJmhWKT314G2AqWZoTSqR5Y2umFKmruhTh6xuTPaxgRHwV2RLNCgW0KBKbW12AJFMFxm5uTkHCcxb/nINxOf/pwDF5bTlja6ZpoUx9WChTFAtlimKhTFEslCmKhTJFsVCmKEmh/GUZZl+SX5ZhzKHshPIXjplS4NJvDFQVNkYHeLwAAAAASUVORK5CYII=',   // デシャップ
  'iVBORw0KGgoAAAANSUhEUgAAAJQAAAAiCAYAAAC5r/C8AAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAKWSURBVHhe7Zy/i9RQFEb3TxsQESzd1k5sLKwEaxsbrUTEShDsxMZuG7ESS4ttrEXwB6uiKCJC5AS+4XGTTF42b2Xy8l047Gxykwncw3133sAcXLjxojl/5b4xs8GlA15sNhtjZoNLW6HOXbrWMc6YHHCnIxR/L97+YMxkUocslJmNhTJFsVCmKBbKFMVCmaJULdTVh58b4vHLH51z5myoWqjL9z61Qr16+7tzrgSHdz821x+dtNx6+rV5cPS9OXrzqzl+96d9HfPXQNVCgYKiq+DvT/62/8fcXG4++bK9765AOF2DcIg9532XQHVC0ZUouOQZijvPv3WuBY4jnGRg2eT/uGxyjKAbIYqC90UaniPNf/b65zaH/HheOYgXjy+J6oQaCgqPFBQMSdJrKCTy8VrdR8IhB6HzQyji8RTel+dQRKly7rHvVCeUOobkUWdAlJgrlKMCE5q7JNjYTDRFBp4t3i9X3H2nOqEiFG5MiNiVEAzS67UUqfC5kTszIRkxtBQvheqFkgCaTxCE7sMMlC59RF930Bym7nUWQjGvKeIyuDSqEyoO5enMEiMtNrl9XYxAvng8oojHc1AXVFdcMlUJNdY9JA3C5XQCOhoRP+FF1GFOs9+lzVci55n2naqEAhVWQzmS6SN+zN0Fkui6sUJLvKkdJpVp6dsFojqh+tAclDPPADJpqexbBmOuxMu9P2iZI6aKuM+sQigVL+cTVLpXNFZoOldurvK1carIeaYlsQqhtC0wNgulu9lTcnNkirv2Q7vlS2cVQulL4jFJ6Ba537cx89Bpdm2YpmiW4xlqFEmsQijz/7BQpigWyhTFQpmiWChTFAtlimKhTFF6hfKPZZjT0vtjGcbMpRXKPzhmSoFL/wBMJquinz3vVAAAAABJRU5ErkJggg==',   // キッチン
  'iVBORw0KGgoAAAANSUhEUgAAAJQAAAAiCAYAAAC5r/C8AAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAALLSURBVHhe7Zy9ihRBFEb3neYFxkAQEw1EVDBQlMVAWFBYUBEMBhRkxWwTAzFSMDA1NDP0iVpO47dca6q7q+3CnZr9gsNMd/10NffsvdW9MAfvvj7rnr4+NGYxuHTAl/V6bcxicOlMqDsPbmwZZ0wJuLMlFJ+ff50YM5vokIUyi7FQpioWylTFQpmqWChTlSaFOj457FarVf+ZtpnzxUIt4MOPTXf56qXu0fO73dsvx1vtFxELtQAkunbrSr8WgWRpv4tEE0Ldfnj9r6ANQb90bEQivv/2Yqst8vL0cZ95Pv180x+ffn/VHx9t7m/1BSRibjJV2jaUvVjr1HpbxEL9gcDfO7rZf0cA+iEWx/TnWO1z0NrS87qn9HzrNCVUTgRQwKNQJRLG/gjFOZWs2C7BYomlP+UO6YbK3JCInOccbWQ9UMbiHPMqO7bGXgilgMeSM1eoNCsxl+ZTZtt8fHLWXwKKoXIXxyFJugZBSWU92pNRZtP5WmAvhCrZpKvP2Dy0pdkElE1ymQgRWV+UDaI8MdtwTH/JxtyxHbGA78pwU4zd9/+mCaGmUMCHREEE2hVEZaEU5skFhzEKcimUsalgSy4dIyXnJKeFOgdiJkjb1E4ZAQVoTgAU5KEnvBxRhLG9UNyY6z7mirtrNC/UVCZQVmJPMlcogkyAGZMrdzm4jmTKvS6IaO2si3EINjVm12laKGWP+M5IcKwMkJaQEqGU2Ur7Q9ynpXuqHOmDwD7QpFBkC+2blH3SdmWWGNhSoZhPMuWe3tJrMZ+uV5KZ4lj65x4EWqUpoQhA/FcHGWioFCFCmiVKhIqvA8b2TbEcCsammXKK+FS3DzQlFCARf9FDT3RjSKixrEP54Rol8zMPfcdebk7BvcQnvdZpTiiz21goUxULZapioUxVLJSpioUyVbFQpipZofxjGeZfyf5YhjFL6YXyD46ZWuDSby5kqdQIgvaYAAAAAElFTkSuQmCC',   // ウォッシャー
  'iVBORw0KGgoAAAANSUhEUgAAAJQAAAAiCAYAAAC5r/C8AAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAIYSURBVHhe7Zy9SsRAFIX3nfICKcRGEOxUFGz8qRQUthC2UxBlQUEU1GYL2Uqx0NbS0tLXiZzADdk7czVrxmbuKT42MzszKebj3NkNZHA7fqmOjy4I6Q1cGuCiLEtCegOXGqE213cC4wjpAtwJhMLn6/SLkLlpO0ShSG8oFEkKhSJJoVAkKRSKJIVCKYqiqNH9FtOHj2o0vOrE3eVbMD833AqFzdUbDiDT4sJS0C/E1hEJfyM2PzfcCiXyzIteR4Ta3xvV1xaTm/dgbo64F0qXIUmurgJ0TaiN1d1gbo5QKCWU1W/BhJqFQhlCQRB9ftJjARNqFgplCBUjdqgWoWLfecStUEggiKBLkSWaRlJL1kEC6USz0GvlhFuhIABE0P1dhdLpNQ96rZxwK9TK8lqzue30ENH0GcoSTJe8mJDjk8dqe+uwep58BvNzw6VQ2FhsOqRCWydIDKtUaaGwNv4YRR/K6enovrmXLq854lIoJIakENq4tn6FaWE0+nt5FNOWcXhw5iKdgEuhpKxdnz/VbVxbj1vk0G0JJQmE+ZJMWB9pJG2kEyT2IJU7oZAgIoD0tdPEQguFM5EeA5EgGO6BMRAI6dQek3vpcycUwIYiMaQtMuhxQJc0AfMhx08HdgGCQTSMt+6TCy6FIv8HhSJJoVAkKRSKJIVCkaRQKJIUCkWSEhWKL8sgfyX6sgxC+lILxReOkVTApW/DAHf/kZph/wAAAABJRU5ErkJggg==',   // 店長
  'iVBORw0KGgoAAAANSUhEUgAAAJQAAAAiCAYAAAC5r/C8AAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAJ7SURBVHhe7ZyxaxRBFIfzP11pkQML4YKQkEKUQAiBYBnQPmAnWAjWwVLSBYJgmzIQCDaCpbW19cq33O94vN27JO7YzPyKj5t5ezO7sB/vvduD3bq7Oug+f9g3ZjK4tMVgPp8bMxlcWgl1erIYGGfMQ8CdgVB8dj9PjHk00SELZSZjoUxRLJQpioUyRbFQpijNCvXj20E3m8268/fP+/mf78d9bBNx/fWXF93Fp90V7MMeOs6Y/Y9fbQ/OXTPNC4UMcb6JuP7t66eD4x/PFqvjjBVHvnz+WrFQS6Hg1/XhICsBWSYLRUZS7Ov5fj9mPXPJhHQa8518DTXSpFBIFG+6MshY1onkPRRDuL2dJ32Ze/fm2WpflcB4rt83R4PrqYkmhcqi6ObnzLQpQ0koBOFTIonYXwHC6ZgyWY00KRTkkvfYHkpCqfQhFuLdXr4crBOcI/ZZNWKhlkLljLIOlSwJlfeJECf75XjNNC8UN1xN9UNgHeslFGOVO0pZlI8YpS7Gam/OmxOKDEOZUl8ENOV8xizDPGaXmJHinLHKnGL3ka+pJpoTCripapIlEaIh0LrskkvXmFDIFn/FEdO62n/diSaF4uaO9T45k2TiHhIKmSQnJQ+B9AhBQhFnXHu5gyaFgnVCaR6zS8xGQjH1X5RRZSr9nRP30DOumh8ZgIW6RyiymR5MxvVZMmUhZae4x7rjNdKkUAiiplwCZcEkgzJQ7qGUcTSnsUcWxGGs51NxHXE/h6oQ3WwEiM0yUo1llzGyUBH9aoSW/hiGJoUy/w8LZYpioUxRLJQpioUyRbFQpigWyhRlVCi/LMP8K6MvyzBmKr1QfuGYKQUu/QWEMhXfQCLDYgAAAABJRU5ErkJggg==',   // 掃除
  'iVBORw0KGgoAAAANSUhEUgAAAJQAAAAiCAYAAAC5r/C8AAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAOjSURBVHhe7ZwxixRBEIXvP+0f2NhEjAwvOENRAxO57AIRTAwOjPQy4QTDMxNRDAwODAQjEwMjwcx45S28492b6p3emd7dWa3gY+d6uqqnq15X9+zBHpx//LJ4/PwsSUYDLR3gYj6fJ8looKUrQR3dfdBRXJLUAO10BIXPy9+LJFkb1VAKKhlNCippSgoqaUoKKmlKCmrinJy+WDw9O++0T5VJCwqBnM1my6D6vRoePXm2uH14tHj1/rJzb5PgefHcb7/+6NyLuHd8suz/6eefa22Y/5vP3zr3Ihgr9Pd722QvBDVkhSKZsKUg4YO8+/6r01+BEGHX16/EnfsPl/beXiISw8uLD1c+ahZF5GMX/JOCUjFFIOFuo6wrCLVZRWkeFM/p64tr7Tdu3uoIxH324fabZi8EhYqh7dgOQLQNqJiQKG/Hao/sFPbDNXwgsZ4oT9YYQaES4r7OE77h06uk++zjvxQUtiQET+E2xcBQBGxDkvWMgvvaHz50m2O7b38ecPihvY4X4bZE7WvhnCAgnqmAL6YSueUJpdWtgUWl4NYQVRkKoVRNSnjV4CGY7UiwV4lVsNq43z50rgRiWjU2D/+A805BGUwIRMNVh2sGLhITQcWiILRCuEgAfXvio741lBaFUko2nlvniLFLcySlylkaY1tMTlB8w0I1YtBwWMUnVmFfoMcKimPp+auGoYLSM1/pXBgBAWo8WOH8YL9tJiUoVicECn9r0lne+xJN8cAH7ABFqucqCsAFxTGj5NfAbbnWnpWpb14O5hcJauh3dq2YjKAQGAaX37mooHCf54RVwaegaukTFJ6Jb3w1uD3htqYVBPNEX4jB/fTBReIM8dWSyQiKAdKA+Lak24MLgURb3ip8i3FBrONL7aPDNBeMv7G6+GqADxUVr0tx2RY7F5RWJj9wu6CAigr9PXHrCgpVTyuQJ5lj1Sad9t4e+fa/h6Ixqf13z6aYjKBcTCASFOD2EZV3nsNqBKWJYBu3IY6pr+aOPxfQZ9b5wC+3bCZdvwbp+9eKg3nCxquU99s2OxcUQOC9MulbU/Tm4jYK7ZAo+CkRJcLfCDEGrqPvtyJB9Z3hXOisziW8P5/J++36ME4mISiHKxes8ypNUAFgV5MsrySsWmMShMqhgi2NBShY76927h/ABvPDp2/7u2SSgkr2lxRU0pQUVNKUFFTSlBRU0pQUVNKUFFTSlFBQ+WMZyVDCH8tIkrEsBZU/OJa0Alr6CyqoI0Eopiu/AAAAAElFTkSuQmCC',   // 水色にもどす
  'iVBORw0KGgoAAAANSUhEUgAAAJQAAAAiCAYAAAC5r/C8AAAACXBIWXMAAAsTAAALEwEAmpwYAAAHhklEQVR4nO2Z52sVSxTAgw1rNCpiFxF7r9jiB8VEhagoYuy9oGLD3hA7dsWGXYxRFBuKvfcWjTXZv+Y8foc3y9y9JffmWdZ9++HozmR3dvac3ylzblp+fr6EEuog/yfpII1/HMcJJdSB/FcdhECFEMnP1EEIVAiUhECFEIhfdRBGKB8YwQmQhED5wAhOgCQEygdGcAIkIVA+MIITIPnlQL19+/aXfsDdu3fl8uXLcvr0adm6detvV+CePXvk8ePHCe+5ffu2Sirrzp8/X/Ly8v44IL4C6vDhw9KwYUN5//69jo8fPy4VKlSQRo0axZSyZcvKzZs3I9bYuXOnPHv2LOb6586dk5kzZ0rr1q3l5cuXusbHjx9j3gtwe/fuTXrvP378kNWrV8vu3bvVsADx6dMn9+/Tp0+XAwcOyJYtW2TEiBEJ18rNzVVJ9t1Hjx6VMmXKyPr162Xfvn0RkshBDx48KPfu3QsuUMXFxdK7d28ZOHCgXgNUp06d3L9v3LhR58w4IyMjCqi6devKqVOnYq6/ePFiyc7OlgkTJsidO3cUYMCKdW+bNm1k4cKFSe/9w4cPCuvQoUOlW7du6hjly5eX+vXrS1ZWlmzYsEGaNWsmX758kUmTJimAqQL17t27qDmiXe3atfWbZsyY4QrficOdOXMm7nvq1asnx44d0+sLFy7Itm3bdJ8rVqyQJUuWyKxZs9z1Eu3Xt0AhT58+VaMQOQxQRAtAA5amTZvq9Zs3b1IGas2aNZKZmVniHk6cOCHVqlVzI2VphD0Dz9WrVzVy3L9/X65du6aOUtKzubm5Mnr06Ig5olt6ero8f/7cnSsoKFDwiYz2vUTG5s2by8qVK+O+gyhOVDPfyBq8c+zYsQrQ3LlzFSoj379//zuBssUboVD02rVr3bEBCgUCEcLc0qVL3bEtU6ZMkRYtWrjjeOmgZ8+eMnv27FLvGyBbtWpVohEwVHp6eoTs2LFDv7NOnToKJPcdOnRIqlevrpCaZwsLC6Vr1676DCDY0rlzZ6latWrEHNHGfvf+/fulZcuWv8WOvgHqyJEj0r17d1VG27ZtNX00adJEr6lvDFBPnjyRQYMGqVSsWFGfMWNbeK5WrVrumIjhfScFe6VKleT169el3vfUqVOjgASARM8UFxcrHEQ0gEpLS9OosW7dOk1Nly5dcu998eKFtGvXThYsWKD3UTcaJyGNM0ddZeaWL18uDRo0iIJ58ODBwQUKhVMkAw0KwTvxor59+8Z9JtWUh2L79++fcB+ANnHiRDXsrVu3kt7/9evXtY7imn0DvinK+TagOHv2bNzn8/LypGbNmlJUVKRA8d0A1qVLFwXIW0sBDtfo6tGjR+7feKfRn5njvV6gpk2bJmPGjAkuUHYNZRTCiQgDk/oAxUSWKlWqJA0UBjLXRA3qs3jv5sRDIdu4cWOttTAo3p/MvgGGiGrGgGu/i9MfzhIvUmVnZ7uHAFOUs19zQIn33tICNXLkSJkzZ07E3NevX4MN1Lx58/TkBFAo3NyDoZMFCsNyZL9y5Yr06tUrSom2UJBidHpVjEk5sSIkJyI70tGOoGb5/PmzO/fq1Sudo54ycxwmKHhjrZeRkeG2MAxQGLh9+/ayaNGihECxZ9NK4XSJfvjfzPF3ajr7uZycHD312tD16NEj2ED169dPayVvhCIaJAsUhTdHaopclByv78Jpp3LlympcM3f+/Hldz3svhrB7SQMGDNA06b1v/PjxEY5AGqU+s1NYQUGBpkeK8VhtA6IPbQFOqMlEqGRl3Lhx6mhm3KdPHy3+AwsUx1o8DYVTdJ88eTIKAIzjBQpoTG8lFaFIRamxjtbfvn2LSKH0k0wqxLNJwfZx3ojpPdlzRDyiLtdEoMzMTI2MifpQN27cUOchutkpPBWgcMpNmza5YyL1qFGj3DHQmsgcKKComzDi5s2b9WeKWPeQ/4GNEG2nGZNWJk+enNI7Od4TieymqREKZSKVGS9btkznqIUowgGYkxh/o1Cm+2yagEQAjvD2ekRcoiUpkb1mZWVFABsLKATHobajVWDXYQD18OHDEr+RtGkDRE3XsWNHd0wkpSTAcR88eKANX3p9XFMu2DXZXwEUJxc+ipoB5dF/Qtn0lPAs708KCPNEAT7crEPKqlGjhjb2eH7IkCFuL2b48OGaMr01ya5du/T+WMUvdRytBgwMAPSDTF2EYenem3uZZ//AjrHKlSsX5RQYhiYn0Y+OdNG/EQfHIFWzRocOHTQlefdCVLZTo4ksOBgRExi8vTecBCellrLrRzrs/KxlCnG+hd8DcQCAN30xHIb6iyzxVwGFQumrGE/AuMCxatUq9XRv884W0qJ3LY7hHN+3b9+u0BFF+B/xHt8B2K6dbMHgGBnjYzQiS0nfQrsB4GP1ueJJYWGh9skQwKXeSua5ixcv6k85w4YNc2tMajqc04zREXr0Rhm+61d1wP84UKE4/0sdhED5wAhOgCQEygdGcAIkIVA+MIITIAmB8oERnABJCJQPjOAESEKgfGAEJ0ASAuUDIzgBkhAoHxjBCZCEQPnACE4QgQol1EH+T9LBP4a/ZrnQZpSlAAAAAElFTkSuQmCC',   // 週ぜんぶ水色に
  'iVBORw0KGgoAAAANSUhEUgAAAJQAAAAiCAYAAAC5r/C8AAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAL+SURBVHhe7ZwxzuIwEEb/o1FR0VFRUtJS0tLSUtPScwB6TsANOMauXqRvNTuxE0Mi/TaakZ4S4sQxzMvEScHPer3+s1qtgmAyuPTDymKxCILJ4NI/oZbLZc+4ICgBd3pCsdxsNkHwNtahECqYTAgVzEoIFcxKCBXMSggVzMpXC0Xc7/fe9hIul0t3/PF47LWVoj4Oh0Ov7VtpViiSNcT1eu2S+Xg8em22H5Lt24HjiNvt1muD0+nUG5OH/UqE0n658PvXTLNCTQnbz1gyc4Fwn47FCzY2hhJ5a6FZoWC/3//3gysx2+22Wz6fz2477WxTu+2jtIpYiDmFSkFlJDjPbrfrtddK00Kdz+fuR0cWPiOObkkEn0kGQYJ0e7N92H39bS0FfSjRfjyeT2QFycT3822107RQr9erq0JKNvIo4Uqk5lIIw2eqmu1DxzD5LglJ5fsROh/B+IhSobgw+D6tVSVLs0KpOtlIbfNRUllAYvqKNoQqko8SoSQToSXBU2pO3hppUijNnQh7G1JFIJCL7SQHWIfclU9CkU3zsU+E4jyMQbdgZNBY/L4WK1MuSqSsgSaFIukKKxTbx6pUThD1wS3LfvavHYaeuJApJZT6TOErkx4gaEN+zaf0gFE7TQrFD64qZYVSFdI6+yJEyW1OIupFpvrxMdSXku8j93LVypQTHVR5/fYaaVIoQZBgJdIK9e4tTy8yp9zyEMRKpfWhPtiHcyKyxGFp39BrbP7YGmleKBtWqFykkuuPQYxPhLLYed7YpFrvzXwglW3zx9VI80KRcBLGVcxyrEKxv+YowLoqg5VAj//vCEX1QwJfpfx+Hjtm+1lzstJ+aqBpoTR/stt8tUmFPcbPYZDKPi2yrnmYxc+LUlVmaDLuGQrGaC+CmmlaqBQSyouWg/lL6urXds1ffHihAClJPsvcXC0HIuupUMG5x1451MbXCRX8LiFUMCshVDArIVQwKyFUMCshVDArIVQwK0mh4s8ygk9J/llGEEylEyr+cCyYC1z6Cyx3/Ul8WimLAAAAAElFTkSuQmCC'   // 確定する
];

// すでに置いてあるボタンを外す（同じ命令が付いている絵だけ）
function removePositionButtons_(sheet) {
  var images;
  try { images = sheet.getImages(); } catch (e) { return 0; }
  var n = 0;
  images.forEach(function(img) {
    var fn = '';
    try { fn = String(img.getScript() || ''); } catch (e) { fn = ''; }
    if (isPositionButtonScript_(fn)) { img.remove(); n++; }
  });
  return n;
}

// ボタンの絵の色は決め打ちなので、設定シートの色もそれに合わせる
// （設定とボタンで色が食い違うと、押した色と見た目が変わってしまうため）
function syncPositionColors_() {
  var settings = getRequestSettingsSheet_();
  var now = positionCells_(settings);
  var changed = false;
  POSITION_DEFAULTS.forEach(function(pos, i) {
    if (String(now[i][1]).trim().toUpperCase() === pos.color.toUpperCase()) return;
    settings.getRange(POSITION_ROWS[i], 6).setValue(pos.color);
    changed = true;
  });
  return changed;
}

// 今このシートに置いてあるボタンの数
function countPositionButtons_(sheet) {
  var images;
  try { images = sheet.getImages(); } catch (e) { return 0; }
  return images.filter(function(img) {
    var fn = '';
    try { fn = String(img.getScript() || ''); } catch (e) { fn = ''; }
    return POSITION_BUTTON_FUNCS.indexOf(fn) >= 0;
  }).length;
}

// メニュー「ボタンを画面に置く」（今のシート）
function placePositionButtons() {
  var week = activeWeekSheet_();
  if (!week) return;
  syncPositionColors_();
  var placed = putPositionButtons_(week);
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'ボタンを' + placed + '個置きました。' +
    (placed < POSITION_BUTTON_FUNCS.length ? '足りないときはもう一度押してください。'
                                           : '画面に出ないときは、ページを再読み込みしてください。'),
    'ポジション', 8);
}

// 1枚のシートにボタンを置き直す（置けた数を返す）
function putPositionButtons_(week) {
  removePositionButtons_(week);
  SpreadsheetApp.flush();                       // 消したのを確定させてから置く
  for (var i = 0; i < POSITION_BUTTON_FUNCS.length; i++) {
    var blob = Utilities.newBlob(Utilities.base64Decode(POSITION_BUTTON_IMAGES[i]),
                                 'image/png', 'pos' + (i + 1) + '.png');
    var line = Math.floor(i / POSITION_BUTTON_PER_ROW);
    var slot = i % POSITION_BUTTON_PER_ROW;
    var img = week.insertImage(blob, POSITION_BUTTON_COL,
                               POSITION_BUTTON_ROW + line * POSITION_BUTTON_ROW_GAP,
                               slot * (POSITION_BUTTON_W + POSITION_BUTTON_GAP), 1);
    img.setWidth(POSITION_BUTTON_W);
    img.setHeight(POSITION_BUTTON_H);
    img.assignScript(POSITION_BUTTON_FUNCS[i]);
    SpreadsheetApp.flush();                     // 1個ずつ確定させる（まとめて置くと出ないことがある）
  }
  return countPositionButtons_(week);
}

// メニュー「ボタンを消す」
function clearPositionButtons() {
  var week = activeWeekSheet_();
  if (!week) return;
  var n = removePositionButtons_(week);
  SpreadsheetApp.getActiveSpreadsheet().toast(
    n ? 'ボタンを' + n + '個消しました。' : 'このシートにボタンはありませんでした。', 'ポジション', 5);
}

// ===================== 自動でシート作成・お知らせ =====================
// 1日1回（朝10時ごろ）に dailyShiftCheck が動き、次の3つを自動で行う。
//  ①今募集している期間の週シートが足りなければ作る（1日に AUTO_SHEETS_PER_DAY 枚まで）
//  ②その期間に日別の売上予算が入っていない日があれば、店長にメールで知らせる
//  ③スタッフ入力締切の 14日前・7日前・5日前・3日前に、LINEのグループへお知らせを送る
// このほか「毎月のおしらせ」（毎月の決まった日が期日のもの）は、別の予約 monthlyNoticeCheck で、決めた時刻に送る。
var LINE_TOKEN_CELL = 'E9';         // LINE公式アカウントのチャネルアクセストークン
var LINE_TARGET_CELL = 'E10';       // 送り先のID（グループに招待すると自動で入る）
var NOTICE_LOG_CELL = 'E11';        // 送った控え（同じお知らせを二重に送らないため）
var NOTICE_DAYS = [14, 7, 5, 3];    // スタッフ入力締切の何日前に送るか
var NOTICE_HOUR = 10;               // 自動で動かす時刻（朝10時）
var AUTO_SHEETS_PER_DAY = 2;        // 1日に自動で作る週シートの上限（時間切れを防ぐ）
var NOTICE_LOG_KEEP = 12;           // 控えに残す件数
// スマホのフォームの記録（誰が送ったか）を見るためのファイル
var RECORD_SPREADSHEET_ID_DEFAULT = '1tgFvH1k3WxsMb_fG2C_m6vMu_n9KTKGgdScSk7jNxGM';   // シフト希望_アクセス記録（飯塚店）
var RECORD_SUBMIT_SHEET = '送信の記録';
var RECORD_NAME_COL = 3;    // 名前の列
var RECORD_DATA_COL = 16;   // 送った中身（この期間のものかを見る）
var LINE_PUSH_URL = 'https://api.line.me/v2/bot/message/push';
var LINE_REGISTER_CELL = 'E12';     // 送り先の登録を受け付ける期限（他の人に書き換えられないように、10分間だけ受け付ける）
// お試しの送り先。本番のグループとは別のトーク（テスト用のグループなど）を入れておくと、
// 「お試し公開」のときだけ、そちらへ本物と同じお知らせが届く。本番には行かない
var LINE_TEST_TARGET_CELL = 'E32';
var LINE_NOTIFY_TARGET_CELL = 'E49';   // シフト希望が届いたときの通知先（店長が見るトーク）
var LINE_REGISTER_MINUTES = 10;

// 送り先の登録は3種類ある。期限のうしろに付ける印（|test・|notify）で見分ける
// ※フォームのGAS（doPost）と同じ決まり。変える時は両方を直すこと
var LINE_REGISTER_KINDS = {
  '':       { cell: LINE_TARGET_CELL,        label: 'お知らせの送り先',              where: 'スタッフのグループ' },
  'test':   { cell: LINE_TEST_TARGET_CELL,   label: 'お試しの送り先',                where: 'テスト用のトーク' },
  'notify': { cell: LINE_NOTIFY_TARGET_CELL, label: 'シフト希望が届いたときの通知先', where: '店長が見るトーク' }
};
// LINE Developers の「Webhook URL」に貼るアドレス（スマホのフォームと同じアドレスで受け取る）
// LINE Developers の「Webhook URL」に貼るアドレス＝シフト希望ページと同じ
function lineWebhookUrl_() { return requestFormUrl_(); }



/* ================= お店のはじめ方（新しい店舗の設定） =================
   新しいお店では、コードを触らずにここだけ埋めれば動くようにする。
   入っているか・まだかが一目で分かる形にして、上から順に案内する。     */

// いまの入り具合を調べる。画面と、保存したあとの点検の両方で使う
function storeSetupState_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settings = getRequestSettingsSheet_();
  function text(cell) { return String(settings.getRange(cell).getDisplayValue()).trim(); }
  var home = ss.getId() === HOME_SHIFT_SPREADSHEET_ID;
  // 初期値の持ち主のお店なら、その値をそのまま設定に書いておく（以後コードに頼らない）。
  // コピーして作った別のお店では書かない。よその店の値が入ってしまうため
  if (home) {
    if (!text(STORE_FORM_URL_CELL)) settings.getRange(STORE_FORM_URL_CELL).setValue(REQUEST_FORM_URL_DEFAULT);
    if (!text(STORE_RECORD_ID_CELL)) settings.getRange(STORE_RECORD_ID_CELL).setValue(RECORD_SPREADSHEET_ID_DEFAULT);
    STORE_SETUP_CACHE = {};
  }
  var url = text(STORE_FORM_URL_CELL);
  var record = text(STORE_RECORD_ID_CELL);
  var recordName = '';
  if (record) {
    try { recordName = SpreadsheetApp.openById(record).getName(); } catch (err) { recordName = ''; }
  }
  return {
    home: home,
    // 設定が空のまま動いている＝よそのお店の値を使っている。コピーして作ったお店では危ない
    borrowed: !home && (!url || !record),
    store: text(REQUEST_STORE_NAME_CELL),
    token: text(LINE_TOKEN_CELL),
    notifyTarget: text(LINE_NOTIFY_TARGET_CELL),
    loginId: text(LINE_LOGIN_ID_CELL),
    loginSecret: text(LINE_LOGIN_SECRET_CELL),
    url: url,
    urlOk: /^https:\/\/script\.google\.com\/macros\/s\/[A-Za-z0-9_-]+\/exec$/.test(url),
    record: record,
    recordName: recordName,
    shiftId: ss.getId(),
    shiftName: ss.getName()
  };
}

// 画面から呼ぶ保存。★末尾に _ を付けると画面から呼べなくなるため付けない
function saveStoreSetup(values) {
  values = values || {};
  var settings = getRequestSettingsSheet_();
  var store = String(values.store || '').trim();
  var url = String(values.url || '').trim();
  var record = String(values.record || '').trim();

  if (store) settings.getRange(REQUEST_STORE_NAME_CELL).setValue(store);
  if (url) {
    if (!/^https:\/\/script\.google\.com\/macros\/s\/[A-Za-z0-9_-]+\/exec$/.test(url)) {
      throw new Error('ページのアドレスは https://script.google.com/macros/s/… /exec の形で入れてください。');
    }
    settings.getRange(STORE_FORM_URL_CELL).setValue(url);
  }
  if (record) {
    var id = record;
    var m = /\/spreadsheets\/d\/([A-Za-z0-9_-]+)/.exec(record);   // URLを貼られてもIDを取り出す
    if (m) id = m[1];
    if (!/^[A-Za-z0-9_-]{30,}$/.test(id)) throw new Error('アクセス記録のIDが正しくありません。');
    try { SpreadsheetApp.openById(id); }
    catch (err) { throw new Error('そのIDのファイルを開けませんでした。共有されているか確かめてください。'); }
    settings.getRange(STORE_RECORD_ID_CELL).setValue(id);
  }
  STORE_SETUP_CACHE = {};        // 次から新しい値を読む
  return storeSetupState_();
}

// メニュー「お店のはじめ方（新しい店舗）」
function openStoreSetup() {
  var data = JSON.stringify(storeSetupState_()).replace(/</g, '\u003c');
  var html = [
    '<style>',
    '  body{ font-family:sans-serif; font-size:13.5px; line-height:1.7; margin:0; padding:6px 10px 10px; }',
    '  h3{ margin:0 0 4px; font-size:13.5px; }',
    '  p.note{ margin:2px 0 6px; font-size:11.5px; color:#555; }',
    '  .step{ border:1px solid #ddd; border-left:4px solid #ddd; border-radius:8px; padding:9px 11px; margin-bottom:8px; }',
    '  .step.ok{ border-left-color:#1E8449; background:#F6FBF7; }',
    '  .step.todo{ border-left-color:#C0392B; background:#FFFBFA; }',
    '  .mark{ float:right; font-size:11.5px; font-weight:bold; }',
    '  .step.ok .mark{ color:#1E8449; }',
    '  .step.todo .mark{ color:#C0392B; }',
    '  .warn{ background:#FDECEA; border:1px solid #C0392B; border-radius:8px;',
    '    padding:8px 11px; margin-bottom:9px; font-size:12px; color:#A0231A; }',
    '  input{ width:100%; box-sizing:border-box; padding:6px; font-size:12.5px; }',
    '  input[readonly]{ background:#F2F2F2; color:#333; }',
    '  ol{ margin:4px 0 0 18px; padding:0; font-size:12px; color:#333; }',
    '  .foot{ position:sticky; bottom:0; background:#fff; border-top:1px solid #eee;',
    '    padding-top:8px; display:flex; gap:8px; align-items:center; justify-content:flex-end; }',
    '  #msg{ flex:1; margin:0; font-size:11.5px; color:#186A3B; }',
    '  #saveBtn{ padding:6px 14px; font-size:13px; }',
    '</style>',
    '<p class="note">新しいお店で使うときは、上から順に埋めてください。',
    '<b>コードを書き換える必要はありません。</b>緑はもう入っているところです。</p>',
    '<div id="warn"></div>',
    '<div id="list"></div>',
    '<div class="foot"><p id="msg"></p><button id="saveBtn">保存する</button></div>',
    '<script>',
    'var D = ' + data + ';',
    'function esc(s){ return String(s == null ? "" : s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"); }',
    'function step(no, title, ok, body){',
    '  return "<div class=\\"step " + (ok ? "ok" : "todo") + "\\"><span class=\\"mark\\">" +',
    '    (ok ? "入っています" : "まだです") + "</span><h3>" + no + ". " + title + "</h3>" + body + "</div>";',
    '}',
    'function draw(){',
    '  document.getElementById("warn").innerHTML = D.borrowed',
    '    ? "<div class=\\"warn\\"><b>ほかのお店の値で動いています。</b><br>' +
      '2と3がまだ入っていないので、いまは別のお店のページとアクセス記録を使っています。' +
      'このまま使うと、記録が混ざります。必ず入れ替えてください。</div>"',
    '    : "";',
    '  var html = "";',
    '  html += step(1, "お店の名前", !!D.store,',
    '    "<p class=\\"note\\">お知らせや画面の見出しに出ます。</p>" +',
    '    "<input id=\\"store\\" value=\\"" + esc(D.store) + "\\" placeholder=\\"例）飯塚店\\">");',
    '  html += step(2, "このお店のシフト希望ページ", D.urlOk,',
    '    "<p class=\\"note\\">フォームのGASを<b>ウェブアプリとして公開</b>し、出てきたアドレスを貼ります。</p>" +',
    '    "<input id=\\"url\\" value=\\"" + esc(D.url) + "\\" placeholder=\\"https://script.google.com/macros/s/……/exec\\">" +',
    '    "<ol><li>フォームのGASを開く</li><li>デプロイ →「新しいデプロイ」→ ウェブアプリ</li>" +',
    '    "<li>アクセスできるユーザー＝全員</li><li>出てきたアドレスをここに貼る</li></ol>");',
    '  html += step(3, "アクセス記録のファイル", !!D.recordName,',
    '    "<p class=\\"note\\">LINEの持ち主・送信の記録・募集と応募を入れるファイルです。' +
      'お店ごとに1つ作ってください（ほかのお店と同じものを使わないでください）。</p>" +',
    '    "<input id=\\"record\\" value=\\"" + esc(D.record) + "\\" placeholder=\\"スプレッドシートのURLを貼ってもOKです\\">" +',
    '    (D.recordName ? "<p class=\\"note\\">いま見ているファイル：<b>" + esc(D.recordName) + "</b></p>" : ""));',
    '  html += step(4, "LINE通知（お知らせを送る）", !!D.token,',
    '    "<p class=\\"note\\">メニューの「LINEのつなぎこみ」→「LINE通知の設定」で入れます。</p>");',
    '  html += step(5, "シフト希望が届いたときの通知先", !!D.notifyTarget,',
    '    "<p class=\\"note\\">スタッフが希望を送るたびに、店長のトークへ知らせます。' +
      'メニューの「LINEのつなぎこみ」→「LINE通知の設定」の⑤で登録します。' +
      '登録しなければ通知は送りません。</p>");',
    '  html += step(6, "LINEログイン（本人を見分ける）", !!(D.loginId && D.loginSecret),',
    '    "<p class=\\"note\\">メニューの「LINEのつなぎこみ」→「LINEログインの設定」で入れます。</p>");',
    '  html += step(7, "フォーム側に、このシフト表を教える", true,',
    '    "<p class=\\"note\\">フォームのGASの <b>SHIFT_SPREADSHEET_ID</b> に、下のIDを入れます。' +
      'ここだけはフォーム側の作業です。</p>" +',
    '    "<input readonly value=\\"" + esc(D.shiftId) + "\\" onclick=\\"this.select()\\">" +',
    '    "<p class=\\"note\\">このファイル：" + esc(D.shiftName) + "</p>");',
    '  document.getElementById("list").innerHTML = html;',
    '}',
    'function show(t, err){ var m = document.getElementById("msg"); m.textContent = t; m.style.color = err ? "#B03A2E" : "#186A3B"; }',
    'document.getElementById("saveBtn").addEventListener("click", function(){',
    '  var btn = document.getElementById("saveBtn");',
    '  btn.disabled = true; show("保存しています…");',
    '  google.script.run',
    '    .withSuccessHandler(function(res){ D = res; draw(); show("保存しました。"); btn.disabled = false; })',
    '    .withFailureHandler(function(e){ show(e.message, true); btn.disabled = false; })',
    '    .saveStoreSetup({',
    '      store: document.getElementById("store").value,',
    '      url: document.getElementById("url").value,',
    '      record: document.getElementById("record").value });',
    '});',
    'draw();',
    '</script>'
  ].join('\n');
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(520).setHeight(640), 'お店のはじめ方');
}

/* ================= 祝日の一覧を新しくする（年に1回だけ） =================
   「祝日」シートは内閣府のCSVから取り込んだ日付の一覧。週シートの色付けと、
   目安人数の区分（平日・土曜・日曜・祝日）の両方がこれを見ている。
   内閣府は毎年2月ごろに翌年ぶんを足すので、3月に1回だけ取り直せば足りる。
   ただし、先の週シートを作ろうとしているのに その年の祝日が入っていなければ、
   3月を待たずにその場で取りに行く（取りこぼしを自分で直す）。          */
var HOLIDAY_SHEET_NAME = '祝日';
var HOLIDAY_REFRESH_MONTH = 3;              // 年に1回、この月に取り直す
var HOLIDAY_AHEAD_DAYS = 150;               // この先の日付の年まで入っていてほしい
var HOLIDAY_YEAR_KEY = 'holidayRefreshedYear';   // 最後に取り直した年
var HOLIDAY_TRY_KEY = 'holidayTriedOn';          // 最後に取りに行った日（1日1回まで）

// その年の祝日が「祝日」シートに入っているか
function hasHolidayYear_(year) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOLIDAY_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return false;
  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  for (var i = values.length - 1; i >= 0; i--) {     // 新しい日付は下にあるので下から探す
    if (values[i][0] instanceof Date && values[i][0].getFullYear() === year) return true;
  }
  return false;
}

// 取り直すかどうかを決めて、必要なら取りに行く。知らせる文を返す（何もしなければ空）
function maybeRefreshHolidays_() {
  var props = PropertiesService.getDocumentProperties();
  var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  var now = new Date();
  var today = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  var year = now.getFullYear();

  var aheadYear = new Date(now.getTime() + HOLIDAY_AHEAD_DAYS * 24 * 60 * 60 * 1000).getFullYear();
  var missing = !hasHolidayYear_(aheadYear);                               // 先の年がまだ無い
  var yearly = (now.getMonth() + 1) === HOLIDAY_REFRESH_MONTH &&
    Number(props.getProperty(HOLIDAY_YEAR_KEY) || 0) !== year;             // 3月の1回
  if (!missing && !yearly) return '';
  if (props.getProperty(HOLIDAY_TRY_KEY) === today) return '';             // 1日に何度も取りに行かない
  props.setProperty(HOLIDAY_TRY_KEY, today);

  try {
    refreshHolidayList_();                                                 // 読み取り用のほうにある
    props.setProperty(HOLIDAY_YEAR_KEY, String(year));
    return '祝日の一覧を取り込み直しました（' + aheadYear + '年ぶんまで）。';
  } catch (err) {
    // 取りに行けなくてもシフト作りは止めない。気づけるようにメールだけ出す
    try {
      MailApp.sendEmail(Session.getEffectiveUser().getEmail(),
        '【シフト表】祝日の一覧を取り込めませんでした',
        '内閣府の祝日CSVを取りに行きましたが、うまくいきませんでした。\n\n' +
        '理由：' + err.message + '\n\n' +
        'そのままでも今までの祝日は使えます。' + aheadYear + '年の祝日が入っていないと、\n' +
        'その日が平日あつかいになり、目安人数が少なめに出ます。');
    } catch (mailErr) {}
    return '';
  }
}

// メニュー「自動でシート作成・お知らせを始める」：毎日朝10時に動く予約を入れる
function startShiftAutomation() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'dailyShiftCheck') ScriptApp.deleteTrigger(t);
  });
  var settings = getRequestSettingsSheet_();
  var notice = readNoticeSettings_(settings);
  ScriptApp.newTrigger('dailyShiftCheck').timeBased().atHour(notice.hour).everyDays(1).create();
  var monthly = readMonthlyNotice_(settings);
  resetMonthlyTrigger_(monthly);
  SpreadsheetApp.getUi().alert('毎日 ' + notice.hour + '時ごろに、次の3つを自動で行います。\n\n' +
    '・募集している期間の週シートを作る（1日' + AUTO_SHEETS_PER_DAY + '枚まで）\n' +
    '・日別の売上予算が入っていない日があれば、メールで知らせる\n' +
    '・スタッフ入力締切の' + notice.days.join('日前・') + '日前に、LINEのグループへお知らせを送る\n\n' +
    (monthly.on ? '「毎月のおしらせ」は、毎月' + monthly.day + '日の' + monthly.days.join('日前・') + '日前、' + monthly.hour + '時ごろに送ります。\n\n'
                : '「毎月のおしらせ」は今は「送らない」になっています。\n\n') +
    '文面や日程はメニュー「お知らせの文面・日程」「毎月のおしらせ」で変えられます。\n' +
    '止めたいときは、メニューの「自動を止める」を選んでください。');
}

// メニュー「自動を止める」
function stopShiftAutomation() {
  var count = 0;
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'dailyShiftCheck' || t.getHandlerFunction() === 'monthlyNoticeCheck') {
      ScriptApp.deleteTrigger(t);
      count++;
    }
  });
  SpreadsheetApp.getUi().alert(count ? '自動を止めました。' : 'もともと自動は動いていません。');
}

// 毎日1回（予約から）動く本体
function dailyShiftCheck() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settings = getRequestSettingsSheet_();
  var now = new Date();
  var today = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd');
  kitRefreshSummary_(settings, today);        // 募集していない時期も書き直すので、期間の判定より前に置く
  var holiday = maybeRefreshHolidays_();      // 週シートを作る前に、祝日の一覧を新しくしておく
  var period = currentRequestPeriod_(settings, now);
  if (!period) return { created: [], missingBudget: [], notice: '', holiday: holiday };

  tidyFutureWeekColors_(ss, today);   // 色と名前の赤を、今の時間に合わせ直す（取りこぼしの回収）
  var created = ensurePeriodSheets_(ss, period);
  ensureBudgetDateRows_(ss, addDaysYmd_(period.start, period.weeks * 7 - 1));   // 入れる欄を先の月まで用意しておく
  var missing = missingBudgetDays_(ss, period);
  var notice = sendDeadlineNotice_(ss, settings, period, today);

  // 予算が足りない日は、週シートを作った日とお知らせを送った日にだけメールする（毎日は送らない）
  if (missing.length && (created.length || notice)) noticeMissingBudget_(ss, period, missing, settings, today);
  return { created: created, missingBudget: missing, notice: notice, holiday: holiday };
}

// これから先の週シート（今週をふくむ）の色と、名前の「決まりの確認」の赤を、今の時間に合わせ直す。
// 手で直したときの自動の塗り直しが取りこぼしても、翌朝にはそろう。変わるところが無ければ何も書かない
var TIDY_WEEKS_PER_DAY = 5;          // 1日にそろえ直す枚数（今週から先に向かって。組んでいる先の週まで届くように）

function tidyFutureWeekColors_(ss, today) {
  var monday = addDaysYmd_(today, -((dowOfYmd_(today) + 6) % 7));   // 今週の月曜
  var names = ss.getSheets().map(function(s) { return s.getName(); })
    .filter(function(n) {
      if (!/^\d{6}$/.test(n)) return false;
      return '20' + n.slice(0, 2) + '-' + n.slice(2, 4) + '-' + n.slice(4, 6) >= monday;
    })
    .sort().slice(0, TIDY_WEEKS_PER_DAY);
  names.forEach(function(name) {
    var week = ss.getSheetByName(name);
    autoStep_('時間の書式（毎日・' + name + '）', function() { fixDayTimeFormat_(week); });
    autoStep_('色の塗り直し（毎日・' + name + '）', function() { repaintWeekColors_(week); });
    autoStep_('決まりの確認（毎日・' + name + '）', function() { checkWeekLaw_(week, true); });
  });
  return names;
}

// 募集している期間の週シートを、足りなければ古い方から順に作る
function ensurePeriodSheets_(ss, period) {
  var created = [];
  while (created.length < AUTO_SHEETS_PER_DAY) {
    var wanted = '';
    for (var w = 0; w < period.weeks; w++) {
      var ymd = addDaysYmd_(period.start, w * 7);
      if (!ss.getSheetByName(ymd.slice(2).replace(/-/g, ''))) { wanted = ymd; break; }
    }
    if (!wanted) break;                       // 期間ぶんのシートがそろっている

    var source = newestWeekSheet_(ss);
    if (!source) break;
    var sourceYmd = Utilities.formatDate(source.getRange('T1').getValue(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
    if (sourceYmd >= wanted) break;           // 順番がおかしい時は何もしない
    created.push(createWeekSheetFrom_(ss, source).getName());   // 一番新しい週の次の週を1枚作る
  }
  return created;
}

// 日別の売上予算が入っていない日（曜日別の基準額で入る日）を返す
function missingBudgetDays_(ss, period) {
  var budgets = readBudgets_(ss);
  var out = [];
  for (var i = 0; i < period.weeks * 7; i++) {
    var ymd = addDaysYmd_(period.start, i);
    if (!budgets.byDate.hasOwnProperty(ymd)) out.push(ymd);
  }
  return out;
}

function noticeMissingBudget_(ss, period, missing, settings, today) {
  var key = 'budget/' + period.start + '/' + today;
  if (noticeAlreadySent_(settings, key)) return;
  var body = '「売上予算」シートに、日別の予算が入っていない日があります。\n' +
    '（入っていない日は、曜日別の基準額がそのまま入ります）\n' +
    '日付の行は先の月末まで用意してありますので、B列に金額を入れてください。\n\n' +
    '募集期間：' + ymdLabel_(period.start) + '〜' + ymdLabel_(addDaysYmd_(period.start, period.weeks * 7 - 1)) + '\n' +
    '足りない日（' + missing.length + '日）：\n' + missing.map(ymdLabel_).join('、') + '\n\n' + ss.getUrl();
  MailApp.sendEmail(Session.getEffectiveUser().getEmail(), '【シフト表】売上予算が入っていない日があります', body);
  writeNoticeLog_(settings, key);
}

// スタッフ入力締切が近づいたら、LINEのグループへお知らせを送る
function sendDeadlineNotice_(ss, settings, period, today) {
  var notice = readNoticeSettings_(settings);
  var left = daysBetweenYmd_(today, period.deadline.slice(0, 10));
  if (notice.days.indexOf(left) === -1) return '';
  var key = 'line/' + period.start + '/' + left;
  if (noticeAlreadySent_(settings, key)) return '';
  var what = 'シフト希望の催促（スタッフ入力締切の' + (left === 0 ? '当日' : left + '日前') + '）';
  var line = lineConfig_(settings);
  if (!line.token || !line.target) {
    noticeSendFailure_(settings, what, line.token ? '送り先のグループが登録されていません。' : 'チャネルアクセストークンが入っていません。');
    return '';
  }

  var card = {};
  var text = fillNoticeText_(noticeTemplateFor_(notice, left), ss, period, left, null, card);
  try {
    linePush_(line, noticeMessage_(text, card));
  } catch (e) {
    noticeSendFailure_(settings, what, e.message);
    return '';
  }
  writeNoticeLog_(settings, key);
  return text;
}

// 「まだ出していない人」と「途中の人」を数える。
// 出したかどうかは、①シート（または保留中）に何か入っている ②スマホのフォームの「送信の記録」にある、
// のどちらかで見る（①＝店長が直接打った分、②＝全部休み希望で送って何も入らない分）。
// 途中の人＝出してはいるが、時間も備考も入っていない週がある人。
// 記録ファイルが読めない時は、シートに何か入っているかどうかだけで数える（お知らせが止まらないように）。
function waitingCounts_(ss, period) {
  var newest = newestWeekSheet_(ss);
  if (!newest) return { notSubmitted: 0, partial: 0 };
  var names = getRosterNames_(newest).map(function(n) { return String(n).trim(); }).filter(function(n) { return n; });
  var filled = filledWeeksByName_(ss, period);       // 名前ごとに、何か入っている週の数
  var submitted = submittedNames_(period);           // 送信の記録から。読めない時は null

  var notSubmitted = 0, partial = 0;
  names.forEach(function(name) {
    var weeks = filled[name] || 0;
    // スマホで送った人か、シート（または保留中）にもう入っている人を「出した人」とする。
    // 店長が直接シートに打ち込んだ人も催促しないようにするため、両方を見る。
    var sent = weeks > 0 || (submitted ? !!submitted[name] : false);
    if (!sent) { notSubmitted++; return; }
    if (weeks < period.weeks) partial++;
  });
  return { notSubmitted: notSubmitted, partial: partial };
}

// 名前ごとに「時間か備考が入っている週の数」を数える（週シートがまだ無い週は、保留中の希望を見る）
function filledWeeksByName_(ss, period) {
  var filled = {};
  var pending = {};
  var pendingSheet = getPendingSheet_();
  var last = pendingSheet.getLastRow();
  if (last >= PENDING_FIRST_ROW) {
    pendingSheet.getRange(PENDING_FIRST_ROW, 1, last - PENDING_FIRST_ROW + 1, 2).getValues().forEach(function(r) {
      if (r[0] && r[1]) pending[String(r[0]) + '/' + String(r[1]).trim()] = true;
    });
  }

  for (var w = 0; w < period.weeks; w++) {
    var sheetName = addDaysYmd_(period.start, w * 7).slice(2).replace(/-/g, '');
    var sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      sheet.getRange(ROSTER_START_ROW, 1, ROSTER_ROWS, REQUEST_INPUT_COLS + 1).getValues().forEach(function(r) {
        var person = String(r[0]).trim();
        if (!person) return;
        for (var c = 1; c <= REQUEST_INPUT_COLS; c++) {   // B〜P列（出勤・退勤と備考）
          if (String(r[c]).trim() !== '') { filled[person] = (filled[person] || 0) + 1; return; }
        }
      });
    } else {
      Object.keys(pending).forEach(function(key) {
        if (key.indexOf(sheetName + '/') !== 0) return;
        var person = key.slice(sheetName.length + 1);
        filled[person] = (filled[person] || 0) + 1;
      });
    }
  }
  return filled;
}

// この募集期間に送信した人の一覧（記録ファイルが読めない時は null）
function submittedNames_(period) {
  try {
    var sheet = SpreadsheetApp.openById(recordSpreadsheetId_()).getSheetByName(RECORD_SUBMIT_SHEET);
    if (!sheet || sheet.getLastRow() < 2) return {};
    var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, RECORD_DATA_COL).getValues();
    var out = {};
    rows.forEach(function(r) {
      var name = String(r[RECORD_NAME_COL - 1]).trim();
      if (!name) return;
      if (String(r[RECORD_DATA_COL - 1]).indexOf('"startDate":"' + period.start + '"') === -1) return;
      out[name] = true;
    });
    return out;
  } catch (e) {
    Logger.log('送信の記録が読めませんでした: ' + e.message);
    return null;
  }
}

function lineConfig_(settings, useTestTarget) {
  return {
    token: String(settings.getRange(LINE_TOKEN_CELL).getDisplayValue()).trim(),
    target: String(settings.getRange(useTestTarget ? LINE_TEST_TARGET_CELL : LINE_TARGET_CELL).getDisplayValue()).trim()
  };
}

// LINEへ送る。うまくいかない時は理由をそのまま出す（トークンの間違い・無料通数切れなどが分かるように）
// message ＝ ふつうの文字列（そのまま文で送る）か、組み立てたメッセージ（ボタン付きカードなど）
function linePush_(line, message) {
  var body = typeof message === 'string' ? { type: 'text', text: message } : message;
  var res = UrlFetchApp.fetch(LINE_PUSH_URL, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + line.token },
    payload: JSON.stringify({ to: line.target, messages: [body] }),
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) {
    throw new Error('LINEへ送れませんでした（' + res.getResponseCode() + '）：' + res.getContentText());
  }
  return true;
}

// LINEへ送れなかったとき、店長へメールで知らせる。
// 自動で動いている時はエラーが画面に出ないので、気づけるようにするため。
// 同じ種類・同じ日のメールは1通だけ（控えで防ぐ）。
function noticeSendFailure_(settings, what, reason) {
  var key = 'fail/' + what + '/' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  if (noticeAlreadySent_(settings, key)) return;
  var body = 'LINEへお知らせを送れませんでした。スタッフには届いていません。\n\n' +
    '　種類：' + what + '\n' +
    '　理由：' + reason + '\n\n' +
    'よくある原因\n' +
    '　・無料で送れる通数（月200通）を使い切った\n' +
    '　・チャネルアクセストークンを作り直した、または期限が切れた\n' +
    '　・グループから公式アカウントが外された\n\n' +
    '確かめ方\n' +
    '　メニュー「スタッフ管理」→「LINE通知の設定」→「テスト送信」を押すと、今の状態が分かります。\n' +
    '　直したあと「今すぐ点検する」を押すと、その場で送り直せます。\n\n' +
    SpreadsheetApp.getActiveSpreadsheet().getUrl();
  MailApp.sendEmail(Session.getEffectiveUser().getEmail(), '【シフト表】LINEへお知らせを送れませんでした', body);
  writeNoticeLog_(settings, key);
}

function noticeAlreadySent_(settings, key) {
  return String(settings.getRange(NOTICE_LOG_CELL).getDisplayValue()).split('|').indexOf(key) !== -1;
}

function writeNoticeLog_(settings, key) {
  var log = String(settings.getRange(NOTICE_LOG_CELL).getDisplayValue()).split('|').filter(function(s) { return s; });
  log.push(key);
  settings.getRange(NOTICE_LOG_CELL).setValue(log.slice(-NOTICE_LOG_KEEP).join('|'));
}

// メニュー「LINE通知の設定」：トークンを入れる・受け取り先を確かめる・テスト送信する
function openLineSettings() {
  var settings = getRequestSettingsSheet_();
  var line = lineConfig_(settings);
  var testTarget = String(settings.getRange(LINE_TEST_TARGET_CELL).getDisplayValue()).trim();
  var notifyTarget = String(settings.getRange(LINE_NOTIFY_TARGET_CELL).getDisplayValue()).trim();
  var box = 'width:100%;padding:6px;font-size:12px;';
  var html = '<div style="font-family:sans-serif;padding:4px 8px;font-size:14px;line-height:1.6;">' +
    '<p style="margin:0 0 8px;font-size:12px;color:#555;">LINE公式アカウントから、スタッフのグループへお知らせを送るための設定です。</p>' +

    '<label style="display:block;font-weight:bold;">① LINE Developers に貼り付けるアドレス（Webhook URL）</label>' +
    '<input readonly value="' + lineWebhookUrl_() + '" onclick="this.select()" style="' + box + '">' +

    '<label style="display:block;font-weight:bold;margin-top:12px;">② チャネルアクセストークン</label>' +
    '<textarea id="token" style="' + box + 'height:70px;">' + escapeHtml_(line.token) + '</textarea>' +

    '<label style="display:block;font-weight:bold;margin-top:12px;">③ お知らせを送る先</label>' +
    '<p style="margin:2px 0;font-size:12px;color:#555;">下のボタンを押してから' + LINE_REGISTER_MINUTES + '分の間に、グループで何か一言送ると登録されます。</p>' +
    '<button onclick="doOpen()" style="padding:5px 10px;">送り先の登録を受け付ける（' + LINE_REGISTER_MINUTES + '分）</button>' +
    '<p style="margin:2px 0;font-size:12px;color:' + (line.target ? '#186A3B' : '#B03A2E') + ';">' +
    (line.target ? '登録済み（' + escapeHtml_(line.target.slice(0, 8)) + '…）' :
      'まだ登録されていません。公式アカウントをグループに招待し、そのグループで何か一言送ってください。') + '</p>' +

    '<label style="display:block;font-weight:bold;margin-top:12px;">④ お試しの送り先（なくても構いません）</label>' +
    '<p style="margin:2px 0;font-size:12px;color:#555;">テスト用のグループを登録しておくと、メニューの「お試し公開」で' +
    '<b>本物と同じお知らせがそちらにだけ</b>届きます。スタッフのグループには送りません。</p>' +
    '<button onclick="doOpenTest()" style="padding:5px 10px;">お試しの送り先を登録する（' + LINE_REGISTER_MINUTES + '分）</button>' +
    '<button onclick="doClearTest()" style="margin-left:6px;padding:5px 10px;">消す</button>' +
    '<p style="margin:2px 0;font-size:12px;color:#555;">' +
    (testTarget ? '登録済み（' + escapeHtml_(testTarget.slice(0, 8)) + '…）' : 'まだ登録されていません。') + '</p>' +

    '<label style="display:block;font-weight:bold;margin-top:12px;">⑤ シフト希望が届いたときの通知先</label>' +
    '<p style="margin:2px 0;font-size:12px;color:#555;">ここを登録しておくと、<b>スタッフがシフト希望を送るたびに</b>そのトークへ届きます。' +
    '店長だけが見るトークを登録してください（スタッフのグループに入れると、全員に名前が流れます）。</p>' +
    '<button onclick="doOpenNotify()" style="padding:5px 10px;">通知先を登録する（' + LINE_REGISTER_MINUTES + '分）</button>' +
    '<button onclick="doClearNotify()" style="margin-left:6px;padding:5px 10px;">消す</button>' +
    '<p style="margin:2px 0;font-size:12px;color:' + (notifyTarget ? '#186A3B' : '#555') + ';">' +
    (notifyTarget ? '登録済み（' + escapeHtml_(notifyTarget.slice(0, 8)) + '…）' :
      'まだ登録されていません。登録するまで通知は送りません。') + '</p>' +

    '<p id="msg" style="margin:10px 0 0;font-size:12px;color:#186A3B;"></p>' +
    '<div style="margin-top:12px;text-align:right;">' +
    '<button onclick="doTest()" style="margin-right:8px;padding:6px 12px;">テスト送信</button>' +
    '<button id="saveBtn" onclick="doSave()" style="padding:6px 12px;">保存する</button>' +
    '</div></div>' +
    '<script>' +
    'function show(t, err){ var m = document.getElementById("msg"); m.textContent = t; m.style.color = err ? "#B03A2E" : "#186A3B"; }' +
    'function doSave(){' +
    '  document.getElementById("saveBtn").disabled = true;' +
    '  google.script.run' +
    '    .withSuccessHandler(function(t){ show(t); document.getElementById("saveBtn").disabled = false; })' +
    '    .withFailureHandler(function(e){ show(e.message, true); document.getElementById("saveBtn").disabled = false; })' +
    '    .saveLineSettings(document.getElementById("token").value);' +
    '}' +
    'function doOpen(){' +
    '  google.script.run' +
    '    .withSuccessHandler(function(t){ show(t); })' +
    '    .withFailureHandler(function(e){ show(e.message, true); })' +
    '    .openLineRegister();' +
    '}' +
    'function doOpenTest(){' +
    '  google.script.run' +
    '    .withSuccessHandler(function(t){ show(t); })' +
    '    .withFailureHandler(function(e){ show(e.message, true); })' +
    '    .openLineRegisterTest();' +
    '}' +
    'function doOpenNotify(){' +
    '  google.script.run' +
    '    .withSuccessHandler(function(t){ show(t); })' +
    '    .withFailureHandler(function(e){ show(e.message, true); })' +
    '    .openLineRegisterNotify();' +
    '}' +
    'function doClearNotify(){' +
    '  google.script.run' +
    '    .withSuccessHandler(function(t){ show(t); })' +
    '    .withFailureHandler(function(e){ show(e.message, true); })' +
    '    .clearLineNotifyTarget();' +
    '}' +
    'function doClearTest(){' +
    '  google.script.run' +
    '    .withSuccessHandler(function(t){ show(t); })' +
    '    .withFailureHandler(function(e){ show(e.message, true); })' +
    '    .clearLineTestTarget();' +
    '}' +
    'function doTest(){' +
    '  show("送っています…");' +
    '  google.script.run' +
    '    .withSuccessHandler(function(t){ show(t); })' +
    '    .withFailureHandler(function(e){ show(e.message, true); })' +
    '    .sendLineTest();' +
    '}' +
    '</script>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(460).setHeight(560), 'LINE通知の設定');
}

function saveLineSettings(token) {
  getRequestSettingsSheet_().getRange(LINE_TOKEN_CELL).setValue(String(token == null ? '' : token).trim());
  return '保存しました。';
}

// 「送り先の登録を受け付ける」：これから10分の間だけ、LINEから来た連絡の送り先を登録する。
// kind で登録先が変わる（''＝お知らせ／'test'＝お試し／'notify'＝シフト希望が届いたときの通知）
function openLineRegister(kind) {
  kind = kind === true ? 'test' : String(kind || '');      // 前の版は true を渡していた
  if (!LINE_REGISTER_KINDS[kind]) kind = '';
  var k = LINE_REGISTER_KINDS[kind];
  var until = new Date(Date.now() + LINE_REGISTER_MINUTES * 60000);
  var value = Utilities.formatDate(until, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm') + (kind ? '|' + kind : '');
  getRequestSettingsSheet_().getRange(LINE_REGISTER_CELL).setValue(value);
  return LINE_REGISTER_MINUTES + '分の間、' + k.label + 'の登録を受け付けます。\n' +
    '登録したいトーク（' + k.where + '）で、何か一言送ってください。';
}

// 「お試しの送り先を登録する」
function openLineRegisterTest() {
  return openLineRegister('test');
}

// 「シフト希望が届いたときの通知先を登録する」
function openLineRegisterNotify() {
  return openLineRegister('notify');
}

// お試しの送り先を消す（お試し公開でLINEを送らないように戻す）
function clearLineTestTarget() {
  getRequestSettingsSheet_().getRange(LINE_TEST_TARGET_CELL).setValue('');
  return 'お試しの送り先を消しました。お試し公開ではLINEを送りません。';
}

// シフト希望の通知先を消す（通知を止める）
function clearLineNotifyTarget() {
  getRequestSettingsSheet_().getRange(LINE_NOTIFY_TARGET_CELL).setValue('');
  return '通知先を消しました。シフト希望が届いてもLINEは送りません。';
}

function sendLineTest() {
  var settings = getRequestSettingsSheet_();
  var line = lineConfig_(settings);
  if (!line.token) throw new Error('先にトークンを入れて「保存する」を押してください。');
  if (!line.target) throw new Error('送り先がまだ登録されていません。公式アカウントをグループに招待し、そのグループで何か一言送ってください。');
  linePush_(line, 'テスト送信です。この形でシフト希望のお知らせが届きます。');
  return '送りました。LINEを見てください。';
}

// メニュー「毎月のおしらせをお試しで送る」
// 本番と同じ形のカードを、お試しの送り先だけに1通送る（スタッフのグループには送らない）
function sendMonthlyNoticeTest() {
  var settings = getRequestSettingsSheet_();
  var monthly = readMonthlyNotice_(settings);
  var line = lineConfig_(settings, true);
  if (!line.token) throw new Error('先にメニュー「LINE通知の設定」でトークンを入れてください。');
  if (!line.target) throw new Error('お試しの送り先が登録されていません。メニュー「お試しの送り先を登録する」から登録してください。');
  var today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  var due = monthlyDueYmd_(today, monthly.day, 0);
  var left = daysBetweenYmd_(today, due);
  if (left < 0) { due = monthlyDueYmd_(today, monthly.day, 1); left = daysBetweenYmd_(today, due); }
  var index = 0;                                 // いちばん近い「送る日」の文を使う
  monthly.days.forEach(function(d, i) { if (d >= left) index = i; });
  var text = fillMonthlyText_(monthly.texts[index], due, left);
  var ym = kitMonthOf_(due);
  var mode = kitMode_(settings);
  var count = mode === '使わない' ? null : kitCount_(ym);
  var pageUrl = count ? staffPageUrl_(settings, '?kit=' + ym) : '';
  linePush_(line, monthlyMessage_(text, mode === '人数' ? count : null, pageUrl));
  var note = count ? '\n\nまだの人：' + count.left + '人（名簿 ' + count.total + '人）' : '';
  // ボタンのリンクも出す。今送ったカードが新しいものかどうかが、ここで分かる
  if (pageUrl) {
    note += '\n\n【ボタンのリンク】\n'
      + (pageUrl.indexOf('https://access.line.me') === 0
        ? '◯ LINEログインを通ります（このカードを押してください）'
        : '× LINEログインを通りません')
      + '\n' + pageUrl.slice(0, 80) + '…';
  }
  SpreadsheetApp.getUi().alert('お試しの送り先に送りました。スタッフのグループには送っていません。\n\n' + text + note);
  return text;
}

// ===== お知らせの文面・日程の設定（メニュー「お知らせの文面・日程」） =====
// 送る日・送る時刻・文面を、店長が画面から変えられるようにする。
// 文面の {期間} {締切} {リンク} {パソコン} {残り日数} {未提出} {途中} {入力状況} は、送るときに中身へ置き換わる。
var NOTICE_DAYS_CELL = 'E13';    // 送る日（例「14,7,5,3」＝締切の何日前）
var NOTICE_HOUR_CELL = 'E14';    // 送る時刻（0〜23時）
var NOTICE_FIRST_CELL = 'E15';   // 1回目（いちばん早い日）に送る文
var NOTICE_LATER_CELL = 'E16';   // 2回目以降に送る文
var NOTICE_LAST_CELL = 'E31';    // 締切の当日（送る日に0を入れたとき）に送る文
var NOTICE_DEFAULT_FIRST = 'シフト希望の入力をお願いします。\n期間 {期間}\nスタッフ入力締切 {締切}\n{リンク}';
var NOTICE_DEFAULT_LATER = '【あと{残り日数}日】シフト希望のスタッフ入力締切が近づいています。\n期間 {期間}\nスタッフ入力締切 {締切}\n{入力状況}\n\n{人手不足}\n{リンク}';
// 締切の当日に送る文。{お礼}と{人手不足}は必ずどちらか片方だけが出る
var NOTICE_DEFAULT_LAST = '【本日が締切】シフト希望のスタッフ入力締切は今日 {締切} です。\n期間 {期間}\n{入力状況}\n\n{お礼}{人手不足}\n{リンク}';
// 入力の様子をお知らせに出すのは、締切までこの日数以内になってから。
// それより前はまだほとんど入力が無く、急かすだけになるので
// {入力状況} も {人手不足}（人が少なめの日・とくに少ない日）も {お礼} も出さない
var NOTICE_STATUS_WITHIN_DAYS = 5;
// {リンク}{パソコン}を消したあと、残りがこの文字数までならその行ごと消す（「パソコンから→」など）
var NOTICE_LINK_LINE_MAX = 12;

// 「14,7,5,3」のような文字を数字の一覧にする（空の切れ端は落とす）
function splitDays_(text) {
  return String(text == null ? '' : text).split(/[,、\s]+/)
    .filter(function(t) { return t !== ''; })
    .map(Number)
    .filter(function(n) { return !isNaN(n); });
}

// 設定を読む。空や書き間違いのときは初期値を使う
function readNoticeSettings_(settings) {
  var daysText = String(settings.getRange(NOTICE_DAYS_CELL).getDisplayValue()).trim();
  var days = splitDays_(daysText).filter(function(n) { return n >= 0 && n <= 60 && Math.floor(n) === n; });
  days = days.filter(function(n, i) { return days.indexOf(n) === i; }).sort(function(a, b) { return b - a; });
  var hour = Number(String(settings.getRange(NOTICE_HOUR_CELL).getDisplayValue()).trim());
  var first = String(settings.getRange(NOTICE_FIRST_CELL).getValue());
  var later = String(settings.getRange(NOTICE_LATER_CELL).getValue());
  var last = String(settings.getRange(NOTICE_LAST_CELL).getValue());
  return {
    days: days.length ? days : NOTICE_DAYS.slice(),
    hour: hour >= 0 && hour <= 23 ? Math.floor(hour) : NOTICE_HOUR,
    first: first.trim() ? first : NOTICE_DEFAULT_FIRST,
    later: later.trim() ? later : NOTICE_DEFAULT_LATER,
    last: last.trim() ? last : NOTICE_DEFAULT_LAST
  };
}

// その日に使う文を選ぶ。締切の当日（残り0日）は専用の文、いちばん早い日は1回目の文
function noticeTemplateFor_(notice, left) {
  if (left === 0) return notice.last;
  return left === notice.days[0] ? notice.first : notice.later;
}


/* ================= 足りないところの募集 =================
   週シートの「目安」と「実時間」の差＝不足を30分ごとに拾い、続いているところを1つの枠にまとめる。
   店長が画面で選んだ枠だけを「募集と応募」シート（アクセス記録）に書き、
   スタッフのシフト希望の画面に赤く重ねて出す。週シートの中には枠を置かない
   （週シートに名前と時間を入れると、その時間の「実時間」が増えて不足が消えてしまうため）。   */
// ===================== 希望と今の表をくらべる =====================
// シフトを組んでいる最中に、人ごとの「出した希望の日数」と「今の表に入っている日数」を
// 並べて見るための浮く窓。シートは読むだけで、何も書き換えない。
//  希望 ＝ 名簿（A4:P25）。1日＝出勤と退勤の2列で、両方そろって1日と数える（フォームの fillState_ と同じ決まり）
//  今   ＝ 曜日ブロックの名前と右はしの「計」。計が0より大きい日だけ数える（出勤履歴と同じ決まり）
// 週シートを切りかえると、窓もその週にきりかわる（getActiveSheet を毎回見るため）。
// MARKER_REQUEST_COMPARE
var COMPARE_REFRESH_MS = 8000;   // 窓が自分で数え直す間かく
var COMPARE_EMP_OFF_TARGET = 9;    // 社員が1か月に取る公休の数
var COMPARE_CAST_DAYS_TARGET = 15; // これ以上の出勤で交通費が出る
var COMPARE_HEIGHT = 700;        // 窓の高さ（「もどす」でもこの高さに戻る）
var COMPARE_MIN_HEIGHT = 96;     // 「小さく」したときの高さ（見出しと合計だけ残る）

// その日が入っている締め期間（◯日締め）を 'yyyy-MM-dd' で返す。
// 締め日が月の日数より大きいときは末日にする（31日締めの2月など）
function closingPeriodOf_(ymd, closingDay) {
  var y = Number(ymd.slice(0, 4)), m = Number(ymd.slice(5, 7));
  function endOf(yy, mm) {
    var last = new Date(Date.UTC(yy, mm, 0)).getUTCDate();
    return yy + '-' + ('0' + mm).slice(-2) + '-' + ('0' + Math.min(closingDay, last)).slice(-2);
  }
  var end = endOf(y, m);
  if (ymd > end) { m++; if (m > 12) { m = 1; y++; } end = endOf(y, m); }
  var pm = m - 1, py = y;
  if (pm < 1) { pm = 12; py--; }
  return { start: addDaysYmd_(endOf(py, pm), 1), end: end };
}

// 出勤履歴から「名前→その期間に入った日数」を数える。同じ日に2行あっても1日と数える
function attendanceDaysIn_(ss, start, end) {
  var out = {}, days = {};
  var sheet = ss.getSheetByName(HISTORY_SHEET_NAME);
  if (!sheet) return days;
  var last = sheet.getLastRow();
  if (last < 2) return days;
  var tz = ss.getSpreadsheetTimeZone();
  sheet.getRange(2, 1, last - 1, HISTORY_HEADERS.length).getValues().forEach(function(r) {
    if (!(r[0] instanceof Date) || !(Number(r[7]) > 0)) return;
    var ymd = Utilities.formatDate(r[0], tz, 'yyyy-MM-dd');
    if (ymd < start || ymd > end) return;
    var key = flatName_(r[2]);
    if (!key) return;
    if (!out[key]) out[key] = {};
    out[key][ymd] = true;
  });
  Object.keys(out).forEach(function(k) { days[k] = Object.keys(out[k]).length; });
  return days;
}

// 「スタッフ区分」シートの 名前→区分（店長・社員・空＝スタッフ）
function roleOfNames_() {
  var sheet = getRoleSheet_();
  var map = {}, last = sheet.getLastRow();
  if (last < 2) return map;
  sheet.getRange(2, 1, last - 1, 2).getValues().forEach(function(r) {
    var key = flatName_(r[0]);
    if (key) map[key] = String(r[1]).trim();
  });
  return map;
}

// 'yyyy-MM-dd' → '11/16'
function shortYmd_(ymd) { return Number(ymd.slice(5, 7)) + '/' + Number(ymd.slice(8, 10)); }

// 窓の下に出す締め期間。その週のうちに締め日が来る＝切り替わりの週なら印を付ける
// （どちらの月で数えているのか、ひと目で分かるようにするため）。MARKER_COMPARE_TERM_LINE
function termInfo_(term, mondayYmd) {
  var sunday = addDaysYmd_(mondayYmd, 6);
  return { text: shortYmd_(term.start) + '〜' + shortYmd_(term.end),
    cross: term.end >= mondayYmd && term.end <= sunday };
}

function requestCompareData_(sheet) {
  // 同じ人を2行に分けないため、名前は flatName_（全角・半角の空白を取る）で突き合わせる。
  // 名簿は「吉田 恋花」、右の表は「吉田　恋花」のように空白がゆれるので、
  // そのまま比べると1人が2行に分かれ、引き算も合わなくなる。MARKER_COMPARE_FLATNAME
  var order = [], label = {}, plan = {}, wish = {}, made = {}, madeDow = {}, inRoster = {};

  // 予定の日（その人がその日に入るつもりだった日）。
  //  ・名簿に希望を出した日
  //  ・右の表に名前を入れた日（店長が頼んで入れた日。時間を消して削った日もふくむ）
  function planDay(key, dayIndex) {
    if (!plan[key]) plan[key] = {};
    plan[key][dayIndex] = true;
  }

  // 希望（名簿）
  var roster = sheet.getRange(ROSTER_START_ROW, 1, ROSTER_ROWS, REQUEST_INPUT_COLS + 1).getDisplayValues();
  roster.forEach(function(r) {
    var shown = String(r[0]).trim();
    var key = flatName_(shown);
    if (!key) return;
    if (order.indexOf(key) === -1) { order.push(key); label[key] = shown; }
    inRoster[key] = true;
    for (var d = 0; d < 7; d++) {
      if (String(r[1 + d * 2]).trim() === '' || String(r[2 + d * 2]).trim() === '') continue;
      planDay(key, d);
      wish[key] = (wish[key] || 0) + 1;
    }
  });

  // 今の表（曜日ブロック）。名前だけで時間が空の日は「一度は予定に入れて、削った日」
  DAY_BLOCKS.forEach(function(block, dayIndex) {
    var names = sheet.getRange(block.names).getValues();
    var hours = sheet.getRange(block.hours).getValues();
    for (var i = 0; i < names.length; i++) {
      var shown = String(names[i][0]).trim();
      var key = flatName_(shown);
      if (!key) continue;
      if (order.indexOf(key) === -1) { order.push(key); label[key] = shown; }   // 名簿に無い人も出す
      planDay(key, dayIndex);
      if (!(Number(hours[i][0]) > 0)) continue;
      made[key] = (made[key] || 0) + 1;
      madeDow[key] = (madeDow[key] || '') + WEEK_DOW_JP[dayIndex];
    }
  });

  // 締め月の数（社員＝公休が9日とれているか／スタッフ＝15日以上で交通費）
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var base = readRequestBase_(getRequestSettingsSheet_());
  var monday = Utilities.formatDate(weekMondayOf_(sheet), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  var empTerm = closingPeriodOf_(monday, base.closingDay);
  var castTerm = closingPeriodOf_(monday, base.castClosingDay);
  var empDays = attendanceDaysIn_(ss, empTerm.start, empTerm.end);
  var castDays = (empTerm.start === castTerm.start && empTerm.end === castTerm.end)
    ? empDays : attendanceDaysIn_(ss, castTerm.start, castTerm.end);
  var empLen = daysBetweenYmd_(empTerm.start, empTerm.end) + 1;
  var roles = roleOfNames_();

  var rows = [], totalPlan = 0, totalMade = 0;
  order.forEach(function(key) {
    var p = plan[key] ? Object.keys(plan[key]).length : 0, m = made[key] || 0;
    totalPlan += p;
    totalMade += m;
    var role = roles[key] || '';
    var emp = (role === ROLE_MANAGER || role === ROLE_EMPLOYEE);
    var done = emp ? (empLen - (empDays[key] || 0)) : (castDays[key] || 0);
    var target = emp ? COMPARE_EMP_OFF_TARGET : COMPARE_CAST_DAYS_TARGET;
    rows.push({ name: label[key], plan: p, wish: wish[key] || 0, made: m, diff: m - p, dows: madeDow[key] || '',
      inRoster: !!inRoster[key], emp: emp,
      // 「/9」「/15」まで出すと細かくなるので、数だけ出す（めやすは色で分かる）
      monthText: (emp ? '公休' : '出勤') + done,
      monthOk: emp ? (done >= target) : (done >= target) });
  });
  return { ok: true, week: sheet.getName(), label: weekRangeLabel_(sheet.getName()),
    rows: rows, totalPlan: totalPlan, totalMade: totalMade,
    terms: { emp: termInfo_(empTerm, monday), cast: termInfo_(castTerm, monday) } };
}

// 窓から呼ぶので、名前のうしろに _ を付けない（google.script.run は _ 付きを呼べない）
function requestCompareNow() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (!/^\d{6}$/.test(sheet.getName())) return { ok: false };
  return requestCompareData_(sheet);
}

function openRequestCompare() {
  var sheet = activeWeekSheet_();
  if (!sheet) return;
  SpreadsheetApp.getUi().showModelessDialog(
    HtmlService.createHtmlOutput(requestCompareHtml_(requestCompareData_(sheet))).setWidth(455).setHeight(COMPARE_HEIGHT),
    '希望と今の表をくらべる');
}

function requestCompareHtml_(data) {
  var json = JSON.stringify(data).replace(/</g, '\u003c');
  return [
    '<style>',
    '  body{ font-family:sans-serif; font-size:13px; line-height:1.5; margin:0; padding:8px 10px 12px; }',
    '  h2{ font-size:13.5px; margin:0 0 2px; }',
    '  .sub{ font-size:11px; color:#777; margin:0 0 8px; }',
    '  table{ border-collapse:collapse; width:100%; }',
    '  th,td{ padding:3px 5px; border-bottom:1px solid #eee; text-align:right; white-space:nowrap; }',
    '  th{ font-size:11px; color:#666; border-bottom:1px solid #bbb; }',
    '  td.name,th.name{ text-align:left; width:100%; }',
    '  td.dow{ font-size:10.5px; color:#999; }',
    '  td.month{ font-size:11.5px; color:#555; white-space:nowrap; }',
    '  td.month.ng{ color:#C00000; font-weight:bold; }',
    '  td.month.ok{ color:#1155CC; font-weight:bold; }',
    '  tbody td:nth-child(2), tbody td:nth-child(3), tbody td:nth-child(4){ font-size:16px; }',
    '  .minus{ color:#C00000; font-weight:bold; }',
    '  .plus{ color:#1155CC; font-weight:bold; }',
    '  .same{ color:#AAA; }',
    '  .none{ color:#BBB; }',
    '  tr.warn{ background:#FFF2F2; }',
    '  tr.extra{ background:#F2F7FF; }',
    '  .foot{ margin-top:9px; padding-top:7px; border-top:1px solid #bbb; font-size:14px; font-weight:bold; }',
    '  .stamp{ font-size:11px; color:#999; margin-top:5px; }',
    '  .term{ margin-top:7px; padding:6px 9px; border-radius:5px; font-size:12.5px; color:#222;',
    '         background:#EEF3FB; border-left:4px solid #1155CC; }',
    '  .term b{ font-weight:bold; font-size:13.5px; }',
    '  .term b.cross{ color:#C00000; }',
    '  .term .cm{ font-size:10.5px; color:#C00000; margin-left:3px; }',
    '  [hidden]{ display:none !important; }',
    '  .bar{ display:flex; align-items:flex-start; gap:8px; }',
    '  .bar h2{ flex:1; }',
    '  .bar button{ font-size:11px; padding:2px 9px; cursor:pointer; white-space:nowrap; }',
    '</style>',
    '<div class="bar"><h2 id="ttl">&nbsp;</h2><button id="btn" onclick="toggle()">小さく</button></div>',
    '<div id="main">',
    '<p class="sub">予定＝希望＋頼んで入れた日／今＝表に入っている日数<br>締め月＝社員は公休の数・スタッフは出勤の数（交通費の目安）</p>',
    '<table><thead><tr><th class="name">名前</th><th>予定</th><th>今</th><th>差</th><th>締め月</th><th></th></tr></thead>',
    '<tbody id="body"></tbody></table>',
    '</div>',
    '<div class="foot" id="foot">&nbsp;</div>',
    '<div class="term" id="term">&nbsp;</div>',
    '<div class="stamp" id="stamp">&nbsp;</div>',
    '<script>',
    'var busy = false;',
    'var small = false;',
    // 「小さく」＝名前の表だけ隠して窓の高さを縮める。数え直しは小さいままでも続く（合計は動く）
    'function toggle(){',
    '  small = !small;',
    '  document.getElementById("main").hidden = small;',
    '  document.getElementById("btn").textContent = small ? "もどす" : "小さく";',
    '  try { google.script.host.setHeight(small ? ' + COMPARE_MIN_HEIGHT + ' : ' + COMPARE_HEIGHT + '); } catch (e) {}',
    '}',
    'function esc(s){ return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;"); }',
    'function two(n){ return (n < 10 ? "0" : "") + n; }',
    '// 締め日がその週のうちに来る＝切り替わりの週。赤くして「この週で切替」と添える',
    'function term(t){',
    '  return "<b class=\\"" + (t.cross ? "cross" : "") + "\\">" + esc(t.text) + "</b>" +',
    '    (t.cross ? "<span class=\\"cm\\">この週で切替</span>" : "");',
    '}',
    'function draw(d){',
    '  var stamp = document.getElementById("stamp");',
    '  if (!d || !d.ok) { stamp.textContent = "週シートを開くと、その週にきりかわります"; return; }',
    '  document.getElementById("ttl").textContent = d.label + "（" + d.week + "）";',
    '  var html = "";',
    '  d.rows.forEach(function(r){',
    '    var cls = r.diff < 0 ? "minus" : (r.diff > 0 ? "plus" : "same");',
    '    var txt = r.diff === 0 ? "±0" : (r.diff > 0 ? "+" + r.diff : String(r.diff));',
    '    var tr = "";',
    '    if (r.plan > 0 && r.made === 0) tr = " class=\\"warn\\"";',
    '    else if (!r.inRoster || (r.wish === 0 && r.made > 0)) tr = " class=\\"extra\\"";',
    '    var none = "<span class=\\"none\\">—</span>";',
    '    html += "<tr" + tr + ">" +',
    '      "<td class=\\"name\\">" + esc(r.name) + "</td>" +',
    '      "<td>" + (r.plan ? r.plan : none) + "</td>" +',
    '      "<td>" + (r.made ? r.made : none) + "</td>" +',
    '      "<td class=\\"" + cls + "\\">" + (r.plan === 0 && r.made === 0 ? "" : txt) + "</td>" +',
    '      "<td class=\\"month " + (r.monthOk ? "ok" : (r.emp ? "ng" : "")) + "\\">" + esc(r.monthText) + "</td>" +',
    '      "<td class=\\"dow\\">" + esc(r.dows) + "</td></tr>";',
    '  });',
    '  document.getElementById("body").innerHTML = html;',
    '  document.getElementById("foot").textContent = "合計　予定 " + d.totalPlan + "日 → 今 " + d.totalMade + "日";',
    '  document.getElementById("term").innerHTML = "締め月　社員 " + term(d.terms.emp) +',
    '    "<br>　　　　スタッフ " + term(d.terms.cast);',
    '  var t = new Date();',
    '  stamp.textContent = "自動で数え直しています（" + two(t.getHours()) + ":" + two(t.getMinutes()) + ":" + two(t.getSeconds()) + " 時点）";',
    '}',
    'function tick(){',
    '  if (busy) return;',
    '  busy = true;',
    '  google.script.run',
    '    .withSuccessHandler(function(d){ busy = false; draw(d); })',
    '    .withFailureHandler(function(){ busy = false; }).requestCompareNow();',
    '}',
    'draw(' + json + ');',
    'setInterval(tick, ' + COMPARE_REFRESH_MS + ');',
    '</script>'
  ].join('\n');
}

var RECRUIT_SHEET_NAME = '募集と応募';
var RECRUIT_HEADERS = ['枠ID', '週シート', '日付', '曜日', '開始', '終了', '人数', '状態', '応募', '入れた人',
  '作った日時', '種類', '候補', '声をかける人'];
var RECRUIT_MIN_MINUTES = 90;        // これより短い不足は、画面で最初からチェックを外しておく
var RECRUIT_OPEN = '募集中';
var RECRUIT_KIND_NEW = '新しく';     // だれでも入れる枠（候補の時間から選んでもらう）
var RECRUIT_KIND_STRETCH = '延ばす'; // その日もう入っている人に、早く来る／遅くまでを頼む枠
var RECRUIT_KINDS = [RECRUIT_KIND_NEW, RECRUIT_KIND_STRETCH];
var RECRUIT_MAX_CHOICES = 4;         // 1つの枠に出せる候補の時間の数

// 「募集と応募」シート（アクセス記録の中。使うときに作る。スタッフには見せない）
function getRecruitSheet_() {
  var ss = SpreadsheetApp.openById(recordSpreadsheetId_());
  var sheet = ss.getSheetByName(RECRUIT_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(RECRUIT_SHEET_NAME);
    sheet.hideSheet();
  }
  if (sheet.getMaxColumns() < RECRUIT_HEADERS.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), RECRUIT_HEADERS.length - sheet.getMaxColumns());
  }
  var head = sheet.getRange(1, 1, 1, RECRUIT_HEADERS.length).getDisplayValues()[0];
  var same = RECRUIT_HEADERS.every(function(h, i) { return String(head[i]).trim() === h; });
  if (!same) {
    sheet.getRange(1, 1, 1, RECRUIT_HEADERS.length).setValues([RECRUIT_HEADERS])
      .setFontWeight('bold').setBackground('#EFF0ED');
    sheet.setFrozenRows(1);
    sheet.getRange('C:C').setNumberFormat('@');
    sheet.getRange('E:F').setNumberFormat('@');
    sheet.getRange('M:M').setNumberFormat('@');
  }
  return sheet;
}

// 分（540）を「9:00」にする
function minuteLabel_(m) {
  return Math.floor(m / 60) + ':' + ('0' + (m % 60)).slice(-2);
}

// 「9:00」「900」を分に直す。空なら -1
function hhmmToMin_(v) {
  var t = String(v == null ? '' : v).trim();
  if (!t) return -1;
  var m = /^(\d{1,2}):(\d{2})$/.exec(t);
  if (m) return Number(m[1]) * 60 + Number(m[2]);
  var n = Number(t.replace(/[^0-9]/g, ''));
  if (!(n > 0)) return -1;
  return Math.floor(n / 100) * 60 + (n % 100);
}

// 募集で選べる時間。1時間ごと（お店の開き時間から閉め時間まで）＋設定の端の時間（21:15 など）
function recruitTimeOptions_() {
  var base = readRequestBase_(getRequestSettingsSheet_());
  var first = hhmmToMin_(base.startFrom), last = hhmmToMin_(base.endTo);
  if (!(last > first)) { first = 9 * 60; last = 21 * 60 + 30; }
  var out = [];
  if (first % 60) out.push(minuteLabel_(first));
  for (var m = Math.ceil(first / 60) * 60; m <= last; m += 60) out.push(minuteLabel_(m));
  if (last % 60) out.push(minuteLabel_(last));
  return out;
}

// 選択肢に無い時間（19:30 など）を、いちばん近い選択肢に寄せる。up=true なら遅い側へ
function snapOption_(label, options, up) {
  var want = hhmmToMin_(label);
  var mins = options.map(hhmmToMin_);
  if (options.indexOf(label) >= 0) return label;
  var best = -1;
  for (var i = 0; i < mins.length; i++) {
    if (up ? mins[i] >= want : mins[i] <= want) { if (up) { best = i; break; } best = i; }
  }
  if (best < 0) best = up ? mins.length - 1 : 0;
  return options[best];
}

// 候補の時間。シートには「16:00-21:15」の行で持つ
function parseChoices_(text, fallbackFrom, fallbackTo) {
  var out = [];
  String(text == null ? '' : text).split('\n').forEach(function(line) {
    var m = /^\s*(\d{1,2}:\d{2})\s*[-–~〜]\s*(\d{1,2}:\d{2})\s*$/.exec(line);
    if (!m) return;
    var a = stepTime_(m[1], ''), b = stepTime_(m[2], '');
    if (a && b && hhmmToMin_(b) > hhmmToMin_(a)) out.push({ from: a, to: b });
  });
  if (!out.length && fallbackFrom && fallbackTo) out.push({ from: fallbackFrom, to: fallbackTo });
  return out;
}

function choicesText_(list) {
  return list.map(function(c) { return c.from + '-' + c.to; }).join('\n');
}

// 候補をまとめた言い方（「16:00〜21:15 か 18:00〜21:15」）
function choicesLabel_(list) {
  return list.map(function(c) { return c.from + '〜' + c.to; }).join(' か ');
}

// 名前を改行で積んだ列を一覧にする
function recruitListCell_(text) {
  return String(text == null ? '' : text).split('\n')
    .map(function(s) { return String(s).trim(); })
    .filter(function(s) { return s; });
}

// 週シートの開始日（T1）。無ければ空
function weekStartYmd_(sheet) {
  var v = sheet.getRange('T1').getValue();
  if (!(v instanceof Date)) return '';
  return Utilities.formatDate(v, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
}

// 1日ぶんの不足を、30分ごとの並びでそのまま返す（9:00 が先頭、足りていないマスは人数、足りていれば0）
function lackRowOfDay_(sheet, block) {
  var cols = TARGET_SLOTS * 2 - 1;
  var guide = sheet.getRange(block.row, block.col, 1, cols).getValues()[0];
  var actual = sheet.getRange(block.row + 1, block.col, 1, cols).getValues()[0];
  var out = [];
  for (var s = 0; s < TARGET_SLOTS; s++) {
    var lack = (Number(guide[s * 2]) || 0) - (Number(actual[s * 2]) || 0);
    out.push(lack > 0 ? lack : 0);
  }
  return out;
}

// 不足を見る終わりの時間から、色を付けるマスの数（9:00から30分ごと）
function lackCellCount_() {
  var n = Math.ceil((shiftEndLimitMin_() - TARGET_FIRST_HOUR * 60) / 30);
  return Math.max(1, Math.min(TARGET_SLOTS, n));
}

// 1日ぶんの不足。続いているところを1つの枠にする。
// 人数は、その枠でいちばん足りない時の人数。endLimitMin は退勤できる上限（22:00でなく21:15で止めるため）
function shortageOfDay_(sheet, block, endLimitMin) {
  var lacks = lackRowOfDay_(sheet, block);
  var out = [];
  var open = null;
  for (var s = 0; s < TARGET_SLOTS; s++) {
    var lack = lacks[s];
    if (lack > 0) {
      var from = (TARGET_FIRST_HOUR * 60) + s * 30;
      if (!open) open = { fromMin: from, toMin: from + 30, people: lack };
      else { open.toMin = from + 30; open.people = Math.max(open.people, lack); }
    } else if (open) { out.push(open); open = null; }
  }
  if (open) out.push(open);
  return out.map(function(x) {
    var to = endLimitMin > 0 && x.toMin > endLimitMin ? endLimitMin : x.toMin;
    return { fromMin: x.fromMin, toMin: to, people: x.people, minutes: to - x.fromMin };
  }).filter(function(x) { return x.minutes > 0; });
}

// 週シート1枚ぶんの不足。日付・曜日も付ける
function shortageOfWeek_(sheet) {
  var startYmd = weekStartYmd_(sheet);
  if (!startYmd) return [];
  var endLimit = shiftEndLimitMin_();
  var out = [];
  TARGET_BLOCKS.forEach(function(block, dayIndex) {
    var ymd = addDaysYmd_(startYmd, dayIndex);
    shortageOfDay_(sheet, block, endLimit).forEach(function(x) {
      out.push({
        week: sheet.getName(), dayIndex: dayIndex, ymd: ymd,
        label: ymdLabel_(ymd), dow: WEEK_DOW_JP[dayIndex],
        from: minuteLabel_(x.fromMin), to: minuteLabel_(x.toMin),
        people: x.people, minutes: x.minutes
      });
    });
  });
  return out;
}

// その日もう入っている人のうち「まだ延ばせる人」。
// 左の入力表（希望）と、右の日別ブロック（いまの時間）の差を見る
function stretchableOn_(sheet, dayIndex) {
  var last = sheet.getLastRow();
  if (last < ROSTER_START_ROW) return [];
  var wish = sheet.getRange(ROSTER_START_ROW, 1, last - ROSTER_START_ROW + 1, 15).getDisplayValues();
  var names = sheet.getRange(DAY_BLOCKS[dayIndex].names);
  var now = sheet.getRange(names.getRow(), names.getColumn(), names.getNumRows(), 3).getDisplayValues();
  var assigned = {};
  now.forEach(function(r) {
    var n = String(r[0]).trim();
    if (n) assigned[flatName_(n)] = { s: hhmmToMin_(r[1]), e: hhmmToMin_(r[2]) };
  });
  var out = [];
  wish.forEach(function(r) {
    var name = String(r[0]).trim();
    if (!name) return;
    var a = assigned[flatName_(name)];
    if (!a || !(a.e > a.s)) return;                         // その日に入っていない人は対象外
    var ws = hhmmToMin_(r[1 + dayIndex * 2]), we = hhmmToMin_(r[2 + dayIndex * 2]);
    var early = (we > ws) ? a.s - ws : 0, late = (we > ws) ? we - a.e : 0;
    out.push({ name: name, now: minuteLabel_(a.s) + '–' + minuteLabel_(a.e),
      wish: (we > ws) ? minuteLabel_(ws) + '–' + minuteLabel_(we) : '',
      nowFrom: minuteLabel_(a.s), nowTo: minuteLabel_(a.e),
      early: early > 0 ? early : 0, late: late > 0 ? late : 0 });
  });
  return out;
}

// 「募集と応募」シートの中身。週シート名を渡すとその週だけ
function readRecruits_(weekName) {
  var sheet = getRecruitSheet_();
  var last = sheet.getLastRow();
  if (last < 2) return [];
  var out = [];
  sheet.getRange(2, 1, last - 1, RECRUIT_HEADERS.length).getDisplayValues().forEach(function(r, i) {
    if (!String(r[0]).trim()) return;
    if (weekName && String(r[1]).trim() !== String(weekName)) return;
    var kind = String(r[11]).trim();
    out.push({ row: i + 2, id: String(r[0]).trim(), week: String(r[1]).trim(), ymd: String(r[2]).trim(),
      dow: String(r[3]).trim(), from: String(r[4]).trim(), to: String(r[5]).trim(),
      people: Number(r[6]) || 1, state: String(r[7]).trim(),
      applied: String(r[8]).trim(), filled: String(r[9]).trim(),
      kind: RECRUIT_KINDS.indexOf(kind) >= 0 ? kind : RECRUIT_KIND_NEW,
      choices: parseChoices_(r[12], String(r[4]).trim(), String(r[5]).trim()),
      targets: recruitListCell_(r[13]) });
  });
  return out;
}

// 画面から呼ばれる保存。選んだ枠だけを「募集中」にする（その週のぶんだけ入れ替える）。
// 同じ日・同じ種類・同じ候補の枠が前からあれば、応募の中身は引き継ぐ
// ★末尾に _ を付けると画面から呼べなくなるため付けない
function saveRecruits(values) {
  values = values || {};
  var week = String(values.week || '');
  if (!/^\d{6}$/.test(week)) throw new Error('週が正しくありません。');
  var list = Array.isArray(values.slots) ? values.slots : [];
  var sheet = getRecruitSheet_();
  var stamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
  var last = sheet.getLastRow();
  var keep = [];
  var before = {};
  if (last >= 2) {
    sheet.getRange(2, 1, last - 1, RECRUIT_HEADERS.length).getDisplayValues().forEach(function(r) {
      if (!String(r[0]).trim()) return;
      if (String(r[1]).trim() !== week) { keep.push(r); return; }    // ほかの週はそのまま残す
      before[r[2] + '|' + String(r[11]).trim() + '|' + String(r[12]).trim()] = r;  // 日付|種類|候補
    });
  }
  var added = [];
  list.forEach(function(x) {
    var ymd = String(x.ymd || '');
    if (!/^\d{4}-\d{2}-\d{2}$/.test(ymd)) throw new Error('日付が正しくありません。');
    var label = ymdLabel_(ymd);
    var kind = RECRUIT_KINDS.indexOf(String(x.kind)) >= 0 ? String(x.kind) : RECRUIT_KIND_NEW;
    // 候補をもらっていなければ、開始と終了を候補1つとして扱う
    var raw = Array.isArray(x.choices) && x.choices.length ? x.choices : [{ from: x.from, to: x.to }];
    var choices = [];
    raw.forEach(function(c) {
      var from = stepTime_(c && c.from, ''), to = stepTime_(c && c.to, '');
      if (!from || !to || hhmmToMin_(to) <= hhmmToMin_(from)) throw new Error(label + ' の時間が正しくありません。');
      var text = from + '-' + to;
      if (choices.indexOf(text) === -1) choices.push(text);
    });
    if (!choices.length) throw new Error(label + ' の時間を入れてください。');
    if (choices.length > RECRUIT_MAX_CHOICES) throw new Error(label + ' の候補は' + RECRUIT_MAX_CHOICES + 'つまでです。');
    var people = Number(x.people);
    if (!(people >= 1 && people <= 9 && Math.floor(people) === people)) throw new Error(label + ' の人数は1〜9で入れてください。');
    var targets = (Array.isArray(x.targets) ? x.targets : [])
      .map(function(n) { return String(n).trim(); }).filter(function(n) { return n; });
    if (kind === RECRUIT_KIND_STRETCH && !targets.length) throw new Error(label + ' は、声をかける人を選んでください。');
    if (kind === RECRUIT_KIND_NEW) targets = [];
    if (kind === RECRUIT_KIND_STRETCH && people > targets.length) people = targets.length;
    var mins = choices.map(function(t) { return t.split('-'); });
    var from = mins.map(function(p) { return p[0]; }).sort(function(a, b) { return hhmmToMin_(a) - hhmmToMin_(b); })[0];
    var to = mins.map(function(p) { return p[1]; }).sort(function(a, b) { return hhmmToMin_(b) - hhmmToMin_(a); })[0];
    var text = choices.join('\n');
    var old = before[ymd + '|' + kind + '|' + text];
    added.push(old
      ? [old[0], week, ymd, x.dow, from, to, people, RECRUIT_OPEN, old[8], old[9], old[10], kind, text, targets.join('\n')]
      : [week + '-' + ymd.replace(/-/g, '') + '-' + from.replace(':', '') + (kind === RECRUIT_KIND_STRETCH ? '-s' : ''),
         week, ymd, x.dow, from, to, people, RECRUIT_OPEN, '', '', stamp, kind, text, targets.join('\n')]);
  });
  var all = keep.concat(added);
  if (last >= 2) sheet.getRange(2, 1, last - 1, RECRUIT_HEADERS.length).clearContent();
  if (all.length) sheet.getRange(2, 1, all.length, RECRUIT_HEADERS.length).setValues(all);
  // 週シートの色をすぐ合わせる（募集をやめた分は色も消える）
  try {
    refreshRecruitMarks_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(week));
  } catch (err) { Logger.log('募集の色付けに失敗: ' + err.message); }
  return added.length;
}


/* --- 募集中の時間を、週シートの「目安」「実時間」の行に色で出す ---
   セルの背景を直接塗ると、もとの色が分からなくなって戻せない。
   条件付き書式（いつでも当たる「=TRUE」のルール）なら、ルールを消すだけで元どおりになる。   */
var RECRUIT_MARK_COLOR = '#F1948A';     // 募集する（赤）
var RECRUIT_MARK_COLOR2 = '#F5B041';    // 今いる人に頼む（橙）
var RECRUIT_MARK_FORMULA = '=TRUE';

// このルールは自分（募集）が入れたものか
function isRecruitMarkRule_(rule) {
  var bc = rule.getBooleanCondition();
  if (!bc) return false;
  if (String(bc.getCriteriaType()) !== 'CUSTOM_FORMULA') return false;
  var values = bc.getCriteriaValues();
  if (!values || String(values[0]) !== RECRUIT_MARK_FORMULA) return false;
  var bg = String(bc.getBackground() || '').toUpperCase();
  return bg === RECRUIT_MARK_COLOR.toUpperCase() || bg === RECRUIT_MARK_COLOR2.toUpperCase();
}

// 30分のマスの集まり（日 → マスの番号）を、続いているところごとに範囲にする。
// 1マスは2列ぶん。縦は「目安」と「実時間」の2行なので、帯のように見える
function slotsToRanges_(sheet, bag) {
  var out = [];
  Object.keys(bag).forEach(function(dayText) {
    var block = TARGET_BLOCKS[Number(dayText)];
    if (!block) return;
    var ks = Object.keys(bag[dayText]).map(Number).sort(function(a, b) { return a - b; });
    var i = 0;
    while (i < ks.length) {
      var from = ks[i], to = from;
      while (i + 1 < ks.length && ks[i + 1] === to + 1) { i++; to = ks[i]; }
      i++;
      out.push(sheet.getRange(block.row, block.col + from * 2, 2, (to - from + 1) * 2));
    }
  });
  return out;
}

// その週で募集している時間を、赤（新しく）と橙（今いる人に頼む）に分ける
function recruitMarkRanges_(sheet) {
  var startYmd = weekStartYmd_(sheet);
  if (!startYmd) return { red: [], orange: [] };
  var cells = lackCellCount_();
  var red = {}, orange = {};
  readRecruits_(sheet.getName()).forEach(function(r) {
    if (r.state !== RECRUIT_OPEN) return;
    var day = daysBetweenYmd_(startYmd, r.ymd);
    if (!(day >= 0 && day <= 6)) return;
    var bag = (r.kind === RECRUIT_KIND_STRETCH) ? orange : red;
    r.choices.forEach(function(c) {
      var from = hhmmToMin_(c.from), to = hhmmToMin_(c.to);
      for (var m = from; m < to; m += 30) {
        var k = Math.floor((m - TARGET_FIRST_HOUR * 60) / 30);
        if (k < 0 || k >= cells) continue;
        if (!bag[day]) bag[day] = {};
        bag[day][k] = true;
      }
    });
  });
  // 同じマスが両方に入っていたら赤を残す（二重に塗らない）
  Object.keys(orange).forEach(function(day) {
    Object.keys(orange[day]).forEach(function(k) {
      if (red[day] && red[day][k]) delete orange[day][k];
    });
    if (!Object.keys(orange[day]).length) delete orange[day];
  });
  return { red: slotsToRanges_(sheet, red), orange: slotsToRanges_(sheet, orange) };
}

// ===== 「不足」の欄（もとの「個人目標」「挨拶・清掃」の欄）=====
// 色だけだと印刷で分かりにくいので、日ごとに「時間と人数」を文字で書く。
// 募集が無くなれば空にする（書きっぱなしにしない）ため、ある日もない日も毎回書き直す。
// 9行を1つにつなげて（結合して）、その中で折り返す。1行ずつ書くと、列がせまくて字が切れるため
var LACK_TEXT_COL = 64;            // 升目の先頭から数えた列（月曜なら CK列）
var LACK_TEXT_ROWS = 9;            // つなげる行数（もとの「個人目標」〜「挨拶・清掃」の9行）
var LACK_TEXT_HEAD = '不足';
var LACK_TEXT_COLOR = '#C00000';   // 文字色
var LACK_TEXT_BG = '#FADBD8';      // うすい赤のしき色（ひと目で気づけるように）
var LACK_TEXT_SIZE = 12;           // 文字の大きさ。見出しも時間も人数も、1つのマスなので全部これ

// その日の募集を「16:00-21:15 1名」の形にする
function lackTextOf_(r) {
  var times = r.choices.map(function(c) { return c.from + '-' + c.to; }).join(' か ');
  return times + ' ' + (Number(r.people) || 1) + '名' +
    (r.kind === RECRUIT_KIND_STRETCH ? ' ※延長' : '');
}

// 曜日ごとの行（募集中のものだけ）
function lackTextLines_(sheet) {
  var out = {};
  var startYmd = weekStartYmd_(sheet);
  if (!startYmd) return out;
  readRecruits_(sheet.getName()).forEach(function(r) {
    if (r.state !== RECRUIT_OPEN) return;
    var day = daysBetweenYmd_(startYmd, r.ymd);
    if (!(day >= 0 && day <= 6)) return;
    if (!out[day]) out[day] = [];
    out[day].push(lackTextOf_(r));
  });
  return out;
}

// 「不足」の欄を1つのマスにまとめて返す。できなければ null
// ※ 結合は「その結合ぜんぶ」を選ばないとほどけない。もとからある結合（横に広い等）を
//    1つずつほどいてから、同じ幅で9行ぶんを結合し直す。
//    何かの理由でできなくても、シフトの保存そのものは止めない
function lackBox_(sheet, block) {
  var row = block.head + 1, col = block.col + LACK_TEXT_COL;
  try {
    var area = sheet.getRange(row, col, LACK_TEXT_ROWS, 1);
    var merges = area.getMergedRanges();
    // 上や下へはみ出している結合（売上の欄など）は、こわさずにそのままにする
    var outside = merges.some(function(m) {
      return m.getRow() < row || m.getLastRow() > row + LACK_TEXT_ROWS - 1;
    });
    if (outside) return null;
    var left = col, right = col;
    merges.forEach(function(m) {
      left = Math.min(left, m.getColumn());
      right = Math.max(right, m.getLastColumn());
    });
    merges.forEach(function(m) { m.breakApart(); });
    var wide = sheet.getRange(row, left, LACK_TEXT_ROWS, right - left + 1);
    wide.merge();
    return wide;
  } catch (e) {
    Logger.log('lackBox_: ' + e.message);
    return null;
  }
}

// 「不足」の欄を、いまの募集に合わせて書き直す
function refreshLackText_(sheet) {
  if (!sheet || !isWeekSheetName_(sheet.getName())) return 0;
  var byDay = lackTextLines_(sheet);
  var total = 0;
  ADJUST_BLOCKS.forEach(function(block, day) {
    var lines = (byDay[day] || []).slice();
    if (lines.length > 7) {                          // 入りきらない分はまとめる
      lines = lines.slice(0, 6).concat(['ほか ' + (lines.length - 6) + '件']);
    }
    var text = lines.length ? LACK_TEXT_HEAD + '\n' + lines.join('\n') : '';
    var box = lackBox_(sheet, block);
    if (box) {                                       // 1つのマスにまとめられた
      box.setValue(text);
      box.setWrap(true);                             // 中で折り返す（字が切れないように）
      box.setVerticalAlignment('top');
      box.setFontSize(LACK_TEXT_SIZE);
      box.setFontWeight(lines.length ? 'bold' : null);
      box.setFontColor(lines.length ? LACK_TEXT_COLOR : null);
      box.setBackground(lines.length ? LACK_TEXT_BG : null);
    } else {                                         // まとめられないときは1行ずつ入れる
      var rows = [];
      for (var i = 0; i < LACK_TEXT_ROWS; i++) rows.push([text.split('\n')[i] === undefined ? '' : text.split('\n')[i]]);
      var area = sheet.getRange(block.head + 1, block.col + LACK_TEXT_COL, LACK_TEXT_ROWS, 1);
      area.setValues(rows);
      area.setFontSize(LACK_TEXT_SIZE);
      area.setFontWeight(lines.length ? 'bold' : null);
      area.setFontColor(lines.length ? LACK_TEXT_COLOR : null);
      area.setBackground(lines.length ? LACK_TEXT_BG : null);
    }
    total += lines.length;
  });
  return total;
}

// 週シートの色を、いまの募集に合わせて引き直す。募集が無くなれば色も消える
function refreshRecruitMarks_(sheet) {
  if (!sheet || !isWeekSheetName_(sheet.getName())) return 0;
  var marks = recruitMarkRanges_(sheet);
  var mine = [];
  [[marks.red, RECRUIT_MARK_COLOR], [marks.orange, RECRUIT_MARK_COLOR2]].forEach(function(pair) {
    if (!pair[0].length) return;
    mine.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(RECRUIT_MARK_FORMULA)
      .setBackground(pair[1])
      .setRanges(pair[0])
      .build());
  });
  var keep = sheet.getConditionalFormatRules().filter(function(rule) { return !isRecruitMarkRule_(rule); });
  // 自分のルールを先に置く（先のルールが勝つので、募集の色がちゃんと出る）
  sheet.setConditionalFormatRules(mine.concat(keep));
  refreshLackText_(sheet);        // 「不足」の欄（時間と人数）も一緒に書き直す
  return marks.red.length + marks.orange.length;
}

// メニュー「足りないところを募集する」
// シフト表を見ながら直したいので、動かせる窓（モードレス）で出す。開いたままシートも触れる
function openRecruit() {
  var sheet = SpreadsheetApp.getActiveSheet();
  if (!isWeekSheetName_(sheet.getName())) {
    SpreadsheetApp.getUi().alert('週のシート（261019 のような名前）を開いてから押してください。');
    return;
  }
  refreshRecruitMarks_(sheet);        // スマホで承認したあとなど、色がずれていたら直す
  var opts = recruitTimeOptions_();
  var startYmd = weekStartYmd_(sheet);
  var stretchByDay = {};
  function stretchOf(dayIndex) {
    if (!(dayIndex >= 0 && dayIndex <= 6)) return [];
    if (!stretchByDay[dayIndex]) stretchByDay[dayIndex] = stretchableOn_(sheet, dayIndex);
    return stretchByDay[dayIndex];
  }
  var slots = [];
  var seen = {};
  // いま募集している枠は、直した中身のまま先に出す（チェックを外せば止まる）
  readRecruits_(sheet.getName()).forEach(function(r) {
    if (r.state !== RECRUIT_OPEN) return;
    seen[r.ymd + '|' + r.from + '|' + r.to] = true;
    slots.push({ ymd: r.ymd, dow: r.dow, label: ymdLabel_(r.ymd), kind: r.kind,
      choices: r.choices, people: r.people, targets: r.targets, minutes: 0, on: true,
      applied: r.applied, now: true,
      stretch: stretchOf(startYmd ? daysBetweenYmd_(startYmd, r.ymd) : -1) });
  });
  // 不足から出した枠（すでに募集している時間と同じものは出さない）
  shortageOfWeek_(sheet).forEach(function(x) {
    if (seen[x.ymd + '|' + x.from + '|' + x.to]) return;
    slots.push({ ymd: x.ymd, dow: x.dow, label: x.label, kind: RECRUIT_KIND_NEW,
      choices: [{ from: snapOption_(x.from, opts, false), to: snapOption_(x.to, opts, true) }],
      people: x.people, targets: [], minutes: x.minutes,
      on: x.minutes >= RECRUIT_MIN_MINUTES, applied: '', now: false,
      stretch: stretchOf(x.dayIndex) });
  });
  slots.sort(function(a, b) {
    if (a.ymd !== b.ymd) return a.ymd < b.ymd ? -1 : 1;
    return hhmmToMin_(a.choices[0].from) - hhmmToMin_(b.choices[0].from);
  });
  var data = { week: sheet.getName(), slots: slots, min: RECRUIT_MIN_MINUTES,
    options: opts, kinds: RECRUIT_KINDS, maxChoices: RECRUIT_MAX_CHOICES };
  SpreadsheetApp.getUi().showModelessDialog(
    HtmlService.createHtmlOutput(recruitDialogHtml_(data)).setWidth(470).setHeight(620),
    '足りないところを募集する');
}

function recruitDialogHtml_(data) {
  var json = JSON.stringify(data).replace(/</g, '\u003c');
  return [
    '<style>',
    '  body{ font-family:sans-serif; font-size:13px; line-height:1.6; margin:0; padding:6px 10px 10px; }',
    '  .lead{ margin:0 0 8px; font-size:11.5px; color:#555; }',
    '  .slot{ border:1px solid #ddd; border-left:4px solid #ddd; border-radius:8px; padding:7px 9px; margin-bottom:8px; }',
    '  .slot.on{ border-left-color:#C0392B; background:#FFFBFA; }',
    '  .head{ display:flex; align-items:center; gap:6px; flex-wrap:wrap; }',
    '  .head label{ font-weight:bold; }',
    '  .lack{ font-size:11px; color:#888; }',
    '  .sub{ font-size:11px; color:#555; margin:6px 0 2px; font-weight:bold; }',
    '  .row{ display:flex; align-items:center; gap:4px; margin-bottom:3px; }',
    '  .row select{ font-size:12px; padding:2px; }',
    '  .x{ border:0; background:none; color:#B03A2E; cursor:pointer; font-size:13px; padding:0 4px; }',
    '  .add{ border:1px dashed #bbb; background:#fff; border-radius:6px; font-size:11.5px; padding:3px 8px; cursor:pointer; }',
    '  .tgt{ display:block; font-size:11.5px; color:#333; }',
    '  .tgt small{ color:#777; }',
    '  .app{ font-size:11.5px; color:#B45309; margin-top:4px; }',
    '  .foot{ position:sticky; bottom:0; background:#fff; border-top:1px solid #eee; padding-top:8px;',
    '    display:flex; align-items:center; gap:8px; justify-content:flex-end; }',
    '  #msg{ flex:1; font-size:11.5px; color:#186A3B; margin:0; }',
    '  #saveBtn{ padding:6px 12px; font-size:13px; }',
    '</style>',
    '<p class="lead">「目安」と「実時間」の差から出しています。<b>この窓は動かせます</b>（上の帯をつかんで左右へ）。',
    '開いたままシートも見られます。募集にした時間は、週シートの「目安」「実時間」の行が',
    '<b style="color:#C0392B;">赤</b>（新しく）・<b style="color:#B9770E;">橙</b>（今いる人に頼む）になります。',
    data.min + '分より短いところは、最初から外してあります。</p>',
    '<div id="list"></div>',
    '<div class="foot"><p id="msg"></p>',
    '<button id="saveBtn">チェックした枠で募集する</button></div>',
    '<script>',
    'var DATA = ' + json + ';',
    'var NEW = DATA.kinds[0], STRETCH = DATA.kinds[1];',
    'function esc(s){ return String(s == null ? "" : s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"); }',
    'function hours(min){ return Math.round(min / 60 * 10) / 10; }',
    'function timeSel(cls, i, c, selected){',
    '  return "<select class=\\"" + cls + "\\" data-i=\\"" + i + "\\" data-c=\\"" + c + "\\">" +',
    '    DATA.options.map(function(t){ return "<option value=\\"" + t + "\\"" + (t === selected ? " selected" : "") + ">" + t + "</option>"; }).join("") +',
    '    "</select>";',
    '}',
    'function pplSel(i, people, max){',
    '  var out = "<select class=\\"ppl\\" data-i=\\"" + i + "\\">";',
    '  for (var n = 1; n <= max; n++) out += "<option value=\\"" + n + "\\"" + (n === people ? " selected" : "") + ">" + n + "人</option>";',
    '  return out + "</select>";',
    '}',
    'function card(s, i){',
    '  var stretch = s.kind === STRETCH;',
    '  var max = stretch ? Math.max(1, s.stretch.length) : 9;',
    '  if (s.people > max) s.people = max;',
    '  var head = "<div class=\\"head\\">" +',
    '    "<label><input type=\\"checkbox\\" class=\\"on\\" data-i=\\"" + i + "\\"" + (s.on ? " checked" : "") + "> " +',
    '    esc(s.label) + "</label>" +',
    '    "<select class=\\"kind\\" data-i=\\"" + i + "\\">" +',
    '      "<option value=\\"" + NEW + "\\"" + (stretch ? "" : " selected") + ">新しく入ってもらう</option>" +',
    '      "<option value=\\"" + STRETCH + "\\"" + (stretch ? " selected" : "") + ">今いる人に頼む</option>" +',
    '    "</select>" + pplSel(i, s.people, max) +',
    '    (s.minutes ? "<span class=\\"lack\\">" + hours(s.minutes) + "時間足りません</span>" : "") +',
    '    (s.now ? "<span class=\\"lack\\">募集中</span>" : "") + "</div>";',
    '  var tgt = "";',
    '  if (stretch) {',
    '    tgt = "<div class=\\"sub\\">声をかける人（その日に入っている人）</div>";',
    '    tgt += s.stretch.length ? s.stretch.map(function(p, t){',
    '      var on = s.targets.indexOf(p.name) >= 0;',
    '      return "<label class=\\"tgt\\"><input type=\\"checkbox\\" class=\\"t\\" data-i=\\"" + i + "\\" data-t=\\"" + t + "\\"" +',
    '        (on ? " checked" : "") + "> " + esc(p.name) +',
    '        " <small>いま " + esc(p.now) + (p.wish ? "／希望 " + esc(p.wish) : "") + "</small></label>";',
    '    }).join("") : "<div class=\\"tgt\\"><small>この日に入っている人がいません。</small></div>";',
    '  }',
    '  var head2 = stretch ? "入ってほしい時間（どれでもよい）" : "入れる時間（どれでもよい）";',
    '  var rows = "<div class=\\"sub\\">" + head2 + "</div>" + s.choices.map(function(c, k){',
    '    return "<div class=\\"row\\">" + timeSel("cf", i, k, c.from) + "〜" + timeSel("ct", i, k, c.to) +',
    '      (s.choices.length > 1 ? "<button type=\\"button\\" class=\\"x\\" data-del=\\"" + i + "\\" data-c=\\"" + k + "\\">✕</button>" : "") +',
    '      "</div>";',
    '  }).join("");',
    '  if (s.choices.length < DATA.maxChoices) {',
    '    rows += "<button type=\\"button\\" class=\\"add\\" data-add=\\"" + i + "\\">＋ この時間でもよい、を足す</button>";',
    '  }',
    '  return "<div class=\\"slot" + (s.on ? " on" : "") + "\\">" + head + tgt + rows +',
    '    (s.applied ? "<div class=\\"app\\">応募：" + esc(s.applied.split("\\n").map(function(l){ return l.split("|")[0]; }).join("、")) + "</div>" : "") +',
    '    "</div>";',
    '}',
    'function draw(){',
    '  var box = document.getElementById("list");',
    '  if (!DATA.slots.length) {',
    '    box.innerHTML = "<p style=\\"color:#186A3B;\\">足りないところはありません。</p>";',
    '    document.getElementById("saveBtn").disabled = true;',
    '    return;',
    '  }',
    '  box.innerHTML = DATA.slots.map(card).join("");',
    '}',
    'function show(t, err){ var m = document.getElementById("msg"); m.textContent = t; m.style.color = err ? "#B03A2E" : "#186A3B"; }',
    'document.getElementById("list").addEventListener("change", function(ev){',
    '  var el = ev.target, i = Number(el.getAttribute("data-i"));',
    '  if (!(i >= 0)) return;',
    '  var s = DATA.slots[i];',
    '  if (el.classList.contains("on")) { s.on = el.checked; draw(); return; }',
    '  if (el.classList.contains("kind")) {',
    '    s.kind = el.value;',
    '    if (s.kind === STRETCH && !s.targets.length) s.targets = s.stretch.map(function(p){ return p.name; });',
    '    draw(); return;',
    '  }',
    '  if (el.classList.contains("ppl")) { s.people = Number(el.value); return; }',
    '  if (el.classList.contains("t")) {',
    '    var name = s.stretch[Number(el.getAttribute("data-t"))].name;',
    '    var at = s.targets.indexOf(name);',
    '    if (el.checked && at === -1) s.targets.push(name);',
    '    if (!el.checked && at >= 0) s.targets.splice(at, 1);',
    '    return;',
    '  }',
    '  var c = Number(el.getAttribute("data-c"));',
    '  if (el.classList.contains("cf")) s.choices[c].from = el.value;',
    '  if (el.classList.contains("ct")) s.choices[c].to = el.value;',
    '});',
    'document.getElementById("list").addEventListener("click", function(ev){',
    '  var el = ev.target;',
    '  var add = el.getAttribute && el.getAttribute("data-add");',
    '  if (add !== null && add !== undefined) {',
    '    var s = DATA.slots[Number(add)];',
    '    var last = s.choices[s.choices.length - 1];',
    '    s.choices.push({ from: last.from, to: last.to });',
    '    draw(); return;',
    '  }',
    '  var del = el.getAttribute && el.getAttribute("data-del");',
    '  if (del !== null && del !== undefined) {',
    '    DATA.slots[Number(del)].choices.splice(Number(el.getAttribute("data-c")), 1);',
    '    draw();',
    '  }',
    '});',
    'document.getElementById("saveBtn").addEventListener("click", function(){',
    '  var slots = DATA.slots.filter(function(s){ return s.on; }).map(function(s){',
    '    return { ymd: s.ymd, dow: s.dow, kind: s.kind, people: s.people,',
    '      choices: s.choices, targets: s.kind === STRETCH ? s.targets : [] };',
    '  });',
    '  var btn = document.getElementById("saveBtn");',
    '  btn.disabled = true;',
    '  show("送っています…");',
    '  google.script.run',
    '    .withSuccessHandler(function(n){ show(n + "件を募集中にしました。スタッフの画面に出ます。"); btn.disabled = false; })',
    '    .withFailureHandler(function(e){ show(e.message, true); btn.disabled = false; })',
    '    .saveRecruits({ week: DATA.week, slots: slots });',
    '});',
    'draw();',
    '</script>'
  ].join('\n');
}

/* ================= 時給と交通費（スタッフのページに出す概算の給与で使う） ================= */
// 時給は契約の更新（3月21日／9月21日）で変わるので、「今の時給」と「前の時給」「いつから今の時給か」を持つ。
// 締め期間の初日が「いつから」より前なら前の時給で計算する。
var ROLE_WAGE_COL = 5;             // スタッフ区分シート：今の時給
var ROLE_PREV_WAGE_COL = 6;        // 前の時給（契約が更新される前の時給）
var ROLE_WAGE_FROM_COL = 7;        // 今の時給がいつからか
var FARE_KIND_CELL = 'E33';        // 交通費の出し方
var FARE_DAYS_CELL = 'E34';        // 「◯日以上で決まった額」のときの日数
var FARE_AMOUNT_CELL = 'E35';      // 金額
var FARE_KINDS = ['なし', '日数以上で決まった額', '出勤1日ごと'];
// 契約が更新される月と日。会社によって違うので設定で変えられる（飯塚店は3月21日と9月21日）
var CONTRACT_MONTHS_CELL = 'E36';
var CONTRACT_DAY_CELL = 'E37';
var CONTRACT_DEFAULT = { months: [3, 9], day: 21 };

// 契約の更新日の決まりを読む（空や書き間違いのときは初期値）
function readContractRule_(settings) {
  var months = splitDays_(settings.getRange(CONTRACT_MONTHS_CELL).getDisplayValue())
    .filter(function(n) { return n >= 1 && n <= 12 && Math.floor(n) === n; });
  months = months.filter(function(n, i) { return months.indexOf(n) === i; }).sort(function(a, b) { return a - b; });
  var day = Number(settings.getRange(CONTRACT_DAY_CELL).getDisplayValue());
  return {
    months: months.length ? months : CONTRACT_DEFAULT.months.slice(),
    day: day >= 1 && day <= 31 ? Math.floor(day) : CONTRACT_DEFAULT.day
  };
}

// 「4月21日」のような言い方にする（画面の説明に使う）
function contractRuleLabel_(rule) {
  return rule.months.map(function(m) { return m + '月' + rule.day + '日'; }).join('・');
}

// その月の更新日（月の日数を超える日は末日にする）
function contractDateOf_(year, month, day) {
  var lastDay = new Date(year, month, 0).getDate();
  return new Date(year, month - 1, Math.min(day, lastDay));
}

// today の直前にあった契約更新日
function contractStartOn_(today, rule) {
  rule = rule || readContractRule_(getRequestSettingsSheet_());
  var year = today.getFullYear();
  for (var i = rule.months.length - 1; i >= 0; i--) {
    var day = contractDateOf_(year, rule.months[i], rule.day);
    if (today >= day) return day;
  }
  return contractDateOf_(year - 1, rule.months[rule.months.length - 1], rule.day);
}

// スタッフ区分シートの時給を読む（名前でひく）
function readWageMap_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ROLE_SHEET_NAME);
  var map = {};
  if (!sheet || sheet.getLastRow() < 2) return map;
  sheet.getRange(2, 1, sheet.getLastRow() - 1, ROLE_WAGE_FROM_COL).getValues().forEach(function(row) {
    var key = flatName_(row[0]);
    if (!key) return;
    map[key] = {
      name: String(row[0]).trim(),
      kind: String(row[1]).trim(),
      wage: wageNumber_(row[ROLE_WAGE_COL - 1]),
      prevWage: wageNumber_(row[ROLE_PREV_WAGE_COL - 1]),
      from: (row[ROLE_WAGE_FROM_COL - 1] instanceof Date) ? row[ROLE_WAGE_FROM_COL - 1] : null
    };
  });
  return map;
}

// その締め期間に使う時給。期間の初日が「いつから」より前なら、前の時給を使う
function wageForPeriod_(info, startYmd) {
  if (!info) return 0;
  if (info.from && info.prevWage > 0) {
    var from = Utilities.formatDate(info.from, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (startYmd < from) return info.prevWage;
  }
  return info.wage;
}

// 交通費の決まりを読む
function readFareRule_(settings) {
  var kind = String(settings.getRange(FARE_KIND_CELL).getDisplayValue()).trim();
  return {
    kind: FARE_KINDS.indexOf(kind) >= 0 ? kind : FARE_KINDS[0],
    days: Number(settings.getRange(FARE_DAYS_CELL).getDisplayValue()) || 0,
    amount: Number(settings.getRange(FARE_AMOUNT_CELL).getDisplayValue()) || 0
  };
}

// 出勤日数から交通費を出す
function fareFor_(rule, workDays) {
  if (!rule || !workDays) return 0;
  if (rule.kind === '日数以上で決まった額') return workDays >= rule.days ? rule.amount : 0;
  if (rule.kind === '出勤1日ごと') return workDays * rule.amount;
  return 0;
}

// メニュー「時給と交通費の設定」
function openWageSettings() {
  syncRoleSheetBirths_();                    // 名簿から読めた分は先にシートへ入れておく
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settings = getRequestSettingsSheet_();
  var rule = readFareRule_(settings);
  var contract = readContractRule_(settings);
  var wages = readWageMap_();
  var newest = newestWeekSheet_(ss);
  var names = newest ? getRosterNames_(newest).filter(function(n) { return String(n).trim(); }) : [];
  var rows = names.map(function(name) {
    var info = wages[flatName_(name)] || {};
    return {
      name: String(name).trim(),
      wage: info.wage || '',
      prevWage: info.prevWage || '',
      from: info.from ? Utilities.formatDate(info.from, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd') : ''
    };
  });
  var data = JSON.stringify({ rule: rule, rows: rows, kinds: FARE_KINDS, contract: contract }).replace(/</g, '\\u003c');
  var box = 'width:100%;padding:6px;font-size:12px;';
  var html = '<div style="font-family:sans-serif;padding:4px 8px;font-size:14px;line-height:1.6;">' +
    '<p style="margin:0 0 8px;font-size:12px;color:#555;">スタッフのページに出す「概算の給与」で使います。社員には出しません。</p>' +
    '<label style="display:block;font-weight:bold;">交通費の出し方</label>' +
    '<select id="kind" style="' + box + '" onchange="build()"></select>' +
    '<div id="fareBox"></div>' +
    '<label style="display:block;font-weight:bold;margin-top:12px;">契約が更新される月（カンマ区切り）</label>' +
    '<p style="margin:2px 0;font-size:12px;color:#555;">この日に時給が変わります。会社によって違うので、お店に合わせて直してください。</p>' +
    '<input id="months" value="' + contract.months.join(',') + '" style="' + box + '">' +
    '<label style="display:block;font-weight:bold;margin-top:8px;">その月の何日から</label>' +
    '<select id="cday" style="' + box + '">' + numberOptions_(1, 31, contract.day, '#日から') + '</select>' +

    '<label style="display:block;font-weight:bold;margin-top:12px;">時給</label>' +
    '<p style="margin:2px 0;font-size:12px;color:#555;">「いつから」を入れると、それより前の締め期間は「前の時給」で計算します。' +
    '名簿から取り込むこともできます。</p>' +
    '<button onclick="doImport()" style="padding:5px 10px;margin-bottom:6px;">名簿から時給を取り込む</button>' +
    '<table style="width:100%;border-collapse:collapse;font-size:12px;"><thead><tr>' +
    '<th style="text-align:left;">名前</th><th>時給</th><th>前の時給</th><th>いつから</th></tr></thead>' +
    '<tbody id="rows"></tbody></table>' +
    '<p id="msg" style="margin:8px 0 0;font-size:12px;color:#186A3B;"></p>' +
    '<div style="margin-top:10px;text-align:right;">' +
    '<button id="saveBtn" onclick="doSave()" style="padding:6px 12px;">保存する</button></div></div>' +
    '<script>' +
    'var DATA = ' + data + ';' +
    'function show(t, err){ var m = document.getElementById("msg"); m.textContent = t; m.style.color = err ? "#B03A2E" : "#186A3B"; }' +
    'function build(){' +
    '  var sel = document.getElementById("kind");' +
    '  if (!sel.options.length) { DATA.kinds.forEach(function(k){ var o = document.createElement("option"); o.value = k; o.textContent = k; if (k === DATA.rule.kind) o.selected = true; sel.appendChild(o); }); }' +
    '  var box = document.getElementById("fareBox");' +
    '  var kind = sel.value;' +
    '  box.innerHTML = kind === "なし" ? "" :' +
    '    (kind === "日数以上で決まった額"' +
    '      ? \'<label style="display:block;margin-top:8px;">この日数以上はたらいたら</label><input id="days" type="number" value="\' + (DATA.rule.days || 15) + \'" style="' + box + '">\' +' +
    '        \'<label style="display:block;margin-top:8px;">この金額（円）</label><input id="amount" type="number" value="\' + (DATA.rule.amount || 8000) + \'" style="' + box + '">\'' +
    '      : \'<label style="display:block;margin-top:8px;">出勤1日あたり（円）</label><input id="amount" type="number" value="\' + (DATA.rule.amount || 200) + \'" style="' + box + '">\');' +
    '}' +
    'function drawRows(){' +
    '  var body = document.getElementById("rows");' +
    '  body.innerHTML = "";' +
    '  DATA.rows.forEach(function(r, i){' +
    '    var tr = document.createElement("tr");' +
    '    tr.innerHTML = \'<td style="padding:2px 0;">\' + r.name + \'</td>\' +' +
    '      \'<td><input id="w\' + i + \'" type="number" value="\' + r.wage + \'" style="width:70px;"></td>\' +' +
    '      \'<td><input id="p\' + i + \'" type="number" value="\' + r.prevWage + \'" style="width:70px;"></td>\' +' +
    '      \'<td><input id="f\' + i + \'" type="date" value="\' + r.from + \'" style="width:130px;"></td>\';' +
    '    body.appendChild(tr);' +
    '  });' +
    '}' +
    'function values(){' +
    '  var kind = document.getElementById("kind").value;' +
    '  var days = document.getElementById("days");' +
    '  var amount = document.getElementById("amount");' +
    '  return { kind: kind, days: days ? days.value : "", amount: amount ? amount.value : "",' +
    '    months: document.getElementById("months").value, contractDay: document.getElementById("cday").value,' +
    '    wages: DATA.rows.map(function(r, i){ return { name: r.name,' +
    '      wage: document.getElementById("w" + i).value,' +
    '      prevWage: document.getElementById("p" + i).value,' +
    '      from: document.getElementById("f" + i).value }; }) };' +
    '}' +
    'function doSave(){' +
    '  document.getElementById("saveBtn").disabled = true;' +
    '  google.script.run' +
    '    .withSuccessHandler(function(t){ show(t); document.getElementById("saveBtn").disabled = false; })' +
    '    .withFailureHandler(function(e){ show(e.message, true); document.getElementById("saveBtn").disabled = false; })' +
    '    .saveWageSettings(values());' +
    '}' +
    'function doImport(){' +
    '  show("名簿を見ています…");' +
    '  google.script.run' +
    '    .withSuccessHandler(function(res){ DATA.rows = res.rows; drawRows(); show(res.message); })' +
    '    .withFailureHandler(function(e){ show(e.message, true); })' +
    '    .importWagesFromRoster(values());' +
    '}' +
    'build(); drawRows();' +
    '</script>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(520).setHeight(640), '時給と交通費');
}

// 画面に入っている契約の決まりを使う（まだ保存していなくても、その場で効くように）。
// 空や書き間違いのときは null を返して、保存してある設定を使わせる
function contractRuleFromValues_(values) {
  if (!values || values.months === undefined) return null;
  var months = splitDays_(values.months).filter(function(n) { return n >= 1 && n <= 12 && Math.floor(n) === n; });
  var day = Number(values.contractDay);
  if (!months.length || !(day >= 1 && day <= 31)) return null;
  months = months.filter(function(n, i) { return months.indexOf(n) === i; }).sort(function(a, b) { return a - b; });
  return { months: months, day: Math.floor(day) };
}

// 画面の「名簿から時給を取り込む」。今の画面の内容を受け取り、名簿の時給で上書きしたものを返す
// （時給が変わっていたら、今の時給を「前の時給」に移して、いつからを直近の契約更新日にする）
function importWagesFromRoster(values) {
  var roster = readRosterBirths_();
  if (!Object.keys(roster).length) {
    throw new Error('名簿が見つかりませんでした。「シフト希望設定」シートのE28に名簿ファイルのURLを入れるか、' +
      'シフト表と同じフォルダに名簿を置いてください。');
  }
  var start = contractStartOn_(new Date(), contractRuleFromValues_(values));
  var startText = Utilities.formatDate(start, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var changed = 0, found = 0, births = 0;
  var rows = (values && values.wages ? values.wages : []).map(function(row) {
    var info = roster[flatName_(row.name)];
    if (!info || !info.wage) return row;
    found++;
    var now = wageNumber_(row.wage);
    // 生年月日も持ち帰る（スタッフのページで「自分かどうか」を確かめるのに使う）
    var birth = info.birth ? Utilities.formatDate(info.birth, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '';
    if (birth) births++;
    // まだ時給が入っていない人（はじめての取り込み）は、そのまま入れるだけ。前の時給はいじらない
    if (!now || now === info.wage) return { name: row.name, wage: info.wage, prevWage: row.prevWage, from: row.from, birth: birth };
    changed++;
    return { name: row.name, wage: info.wage, prevWage: now, from: startText, birth: birth };
  });
  return {
    rows: rows,
    message: found + '人の時給を名簿から読みました。' +
      (births ? births + '人は生年月日も読みました（スタッフが自分の給与を見るときの確認に使います）。' : '') +
      (changed ? changed + '人は名簿の時給が変わっていたので、これまでの時給を「前の時給」に移して、' +
                 '「いつから」を ' + startText + ' にしました。' : '') +
      '内容を確かめて「保存する」を押してください。'
  };
}

// 画面から来た時給と交通費を保存する
function saveWageSettings(values) {
  var settings = getRequestSettingsSheet_();
  var kind = FARE_KINDS.indexOf(String(values.kind)) >= 0 ? String(values.kind) : FARE_KINDS[0];
  var days = Number(values.days) || 0;
  var amount = Number(values.amount) || 0;
  if (kind === '日数以上で決まった額' && !(days >= 1 && days <= 31)) throw new Error('日数は1〜31で入れてください。');
  if (kind !== 'なし' && !(amount >= 0)) throw new Error('金額が正しくありません。');
  settings.getRange(FARE_KIND_CELL).setValue(kind);
  settings.getRange(FARE_DAYS_CELL).setValue(kind === '日数以上で決まった額' ? days : '');
  settings.getRange(FARE_AMOUNT_CELL).setValue(kind === 'なし' ? '' : amount);

  // 契約が更新される月と日
  if (values.months !== undefined) {
    var months = splitDays_(values.months).filter(function(n) { return n >= 1 && n <= 12 && Math.floor(n) === n; });
    if (!months.length) throw new Error('契約が更新される月を1つ以上入れてください（例 3,9）。');
    var contractDay = Number(values.contractDay);
    if (!(contractDay >= 1 && contractDay <= 31)) throw new Error('契約が更新される日が正しくありません。');
    months = months.filter(function(n, i) { return months.indexOf(n) === i; }).sort(function(a, b) { return a - b; });
    settings.getRange(CONTRACT_MONTHS_CELL).setValue(months.join(','));
    settings.getRange(CONTRACT_DAY_CELL).setValue(String(Math.floor(contractDay)));
  }

  var sheet = getRoleSheet_();
  sheet.getRange(1, ROLE_BIRTH_COL, 1, 2).setValues([['生年月日', '高校生']]);
  sheet.getRange(1, ROLE_WAGE_COL, 1, 3).setValues([['時給', '前の時給', '今の時給はいつから']]);
  var last = sheet.getLastRow();
  var names = last >= 2 ? sheet.getRange(2, 1, last - 1, 1).getValues().map(function(r) { return flatName_(r[0]); }) : [];
  (values.wages || []).forEach(function(row) {
    var key = flatName_(row.name);
    if (!key) return;
    var index = names.indexOf(key);
    var rowNo = index >= 0 ? index + 2 : sheet.getLastRow() + 1;
    if (index < 0) { sheet.getRange(rowNo, 1).setValue(String(row.name).trim()); names.push(key); }
    sheet.getRange(rowNo, ROLE_WAGE_COL).setValue(wageNumber_(row.wage) || '');
    sheet.getRange(rowNo, ROLE_PREV_WAGE_COL).setValue(wageNumber_(row.prevWage) || '');
    sheet.getRange(rowNo, ROLE_WAGE_FROM_COL).setValue(
      /^\d{4}-\d{2}-\d{2}$/.test(String(row.from)) ? new Date(String(row.from) + 'T00:00:00+09:00') : '');
    // 名簿から生年月日が読めていれば、空のときだけ入れる（手で直した内容は消さない）
    if (/^\d{4}-\d{2}-\d{2}$/.test(String(row.birth)) && !(sheet.getRange(rowNo, ROLE_BIRTH_COL).getValue() instanceof Date)) {
      sheet.getRange(rowNo, ROLE_BIRTH_COL).setValue(new Date(String(row.birth) + 'T00:00:00+09:00'));
    }
  });
  return '保存しました。';
}

/* ================= 人が少なめの日（お知らせの {人手不足}・{お礼}） ================= */
// 募集期間の日ごとに、希望を出している人数を数えて、欲しい人数より少ない日を知らせる。
// 「あと◯名」と言い切らずに「目安は◯名くらい」と書く（その日の売上や顔ぶれで変わるため）。
// 言い方は3段階。①そろっている＝お礼（{お礼}）②少し足りない＝軽めの声かけ
// ③ひとりあと1日以上足りない＝運営に支障が出ることを伝えて、具体的な日数をお願いする。
// ③にするのは締切まであと3日以内の回だけ（早い回から強く言うと効き目が薄れるため）。
var NEED_WEEKDAY_CELL = 'E29';       // 1日に欲しい人数の目安（平日）
var NEED_WEEKEND_CELL = 'E30';       // 1日に欲しい人数の目安（土日祝）
var NEED_DEFAULT = { weekday: 3, weekend: 5 };
var NEED_LIST_LIMIT = 3;             // 少なめの日がこれより多いときは、曜日でまとめて知らせる
var NEED_HARD_PER_PERSON = 1;        // ひとりあたりこれ以上足りないときは、強めの言い方にする
var NEED_HARD_WITHIN_DAYS = 3;       // 強めの言い方をするのは、締切まであとこの日数以内のときだけ
// 人がそろっている時のお礼（差し込み {お礼}）。足りない日がある時は出ない
var NOTICE_THANKS_TEXT = 'おかげさまで、どの日も人がそろいました。ご協力ありがとうございます。';
var NEED_URGENT_MAX = 3;             // 「とくに少ない日」に名前を出す日数

// 欲しい人数の目安を読む（空や変な数字なら初期値）
function readNeedCounts_(settings) {
  var weekday = Number(settings.getRange(NEED_WEEKDAY_CELL).getDisplayValue());
  var weekend = Number(settings.getRange(NEED_WEEKEND_CELL).getDisplayValue());
  return {
    weekday: weekday > 0 ? weekday : NEED_DEFAULT.weekday,
    weekend: weekend > 0 ? weekend : NEED_DEFAULT.weekend
  };
}

// 募集期間の日ごとに、希望を出している人数を数える（出勤の時間が入っていれば1人）。
// まだ作っていない週シートの日は0人（誰も入っていない日）として数える。
function dailyRequestCounts_(ss, period, keys, needs) {
  var out = [];
  for (var w = 0; w < period.weeks; w++) {
    var sheetName = addDaysYmd_(period.start, w * 7).slice(2).replace(/-/g, '');
    var sheet = ss.getSheetByName(sheetName);
    var rows = sheet ? sheet.getRange(ROSTER_START_ROW, 1, ROSTER_ROWS, REQUEST_INPUT_COLS + 1).getValues() : null;
    for (var d = 0; d < 7; d++) {
      var ymd = addDaysYmd_(period.start, w * 7 + d);
      var cat = categoryOfYmd_(ymd, keys);
      var count = 0;
      if (rows) {
        for (var i = 0; i < rows.length; i++) {
          if (!String(rows[i][0]).trim()) continue;
          if (String(rows[i][1 + d * 2]).trim() !== '') count++;
        }
      }
      out.push({ ymd: ymd, count: count, cat: cat, need: cat === '平日' ? needs.weekday : needs.weekend });
    }
  }
  return out;
}

// 日付の見せ方。祝日は「11/3(火・祝)」
function needDayLabel_(day) {
  var label = ymdLabel_(day.ymd);
  return day.cat === '祝日' ? label.replace(')', '・祝)') : label;
}

// 曜日（祝日は「祝日」）でまとめるときの名前
function needGroupOf_(day) {
  return day.cat === '祝日' ? '祝日' : DOW_JP[dowOfYmd_(day.ymd)] + '曜';
}

// 「ひとりあと◯日」の言い方。2.1なら「2〜3」、2ちょうどなら「2」
function perDaysLabel_(per) {
  var low = Math.max(1, Math.floor(per));
  var high = Math.ceil(per);
  return low === high ? String(low) : low + '〜' + high;
}

// さいごの一文。足りない量をみんなで分けると、ひとりあたり何日かで言い方を変える。
//   ひとり1日より少ない … 軽めに声をかける
//   ひとり1日以上      … 運営に支障が出ることを伝えて、具体的な日数をお願いする
function shortStaffAppeal_(few, people, allowHard) {
  var total = 0;
  few.forEach(function(d) { total += d.need - d.count; });
  var per = people > 0 ? total / people : 0;
  if (!allowHard || per < NEED_HARD_PER_PERSON) return '入れそうな方がいたら、お願いします。';
  return 'このままでは運営に支障が出ます。\n' +
    'ひとりあと' + perDaysLabel_(per) + '日程度でも、足りない日に出勤希望のご協力をお願いします。';
}

// 名簿に載っている人数（いちばん新しい週シートの名前の欄で数える）
function rosterCount_(ss) {
  var newest = newestWeekSheet_(ss);
  if (!newest) return 0;
  return getRosterNames_(newest).filter(function(n) { return String(n).trim(); }).length;
}

// お知らせに入れる「人が少なめの日」の文。少なめの日が無ければ空文字
// 日ごとの人数を数える（重いのはここだけ。文の組み立てとは分けてあるので、言い方違いを作り直しても速い）
function shortStaffInfo_(ss, period) {
  var days = dailyRequestCounts_(ss, period, holidayKeys_(), readNeedCounts_(getRequestSettingsSheet_()));
  return {
    days: days,
    few: days.filter(function(d) { return d.count < d.need; }),
    people: rosterCount_(ss)
  };
}

// 数えてから文を作る
function shortStaffText_(ss, period, allowHard) {
  return shortStaffTextFrom_(shortStaffInfo_(ss, period), allowHard);
}

// 数えた結果から文を作る。allowHard ＝ 強めの言い方をしてよい日か（締切が近い日だけ）
function shortStaffTextFrom_(info, allowHard) {
  var days = info.days;
  var few = info.few;
  if (!few.length) return '';

  var lines = [];
  if (few.length <= NEED_LIST_LIMIT) {
    lines.push('【人が少なめの日】');
    few.forEach(function(d) {
      lines.push('・' + needDayLabel_(d) + ' いま' + d.count + '名（目安は' + d.need + '名くらい）');
    });
  } else {
    lines.push('【人が少なめの日】' + few.length + '日あります');
    var all = {}, short = {};
    days.forEach(function(d) { var g = needGroupOf_(d); all[g] = (all[g] || 0) + 1; });
    few.forEach(function(d) { var g = needGroupOf_(d); short[g] = (short[g] || 0) + 1; });
    ['月曜', '火曜', '水曜', '木曜', '金曜', '土曜', '日曜', '祝日'].forEach(function(g) {
      if (!short[g]) return;
      lines.push('・' + g + ' ' + (short[g] === all[g] ? all[g] + '日ぜんぶ' : all[g] + '日のうち' + short[g] + '日'));
    });
    // まだ誰も入っていない日・1人だけの日は、日付を出す（いちばん困る日なので）
    var urgent = few.filter(function(d) { return d.count <= 1; })
      .sort(function(a, b) { return a.count - b.count || (a.ymd < b.ymd ? -1 : 1); })
      .slice(0, NEED_URGENT_MAX);
    if (urgent.length) {
      lines.push('とくに少ない日');
      urgent.forEach(function(d) {
        lines.push('・' + needDayLabel_(d) + ' ' + (d.count ? d.count + '名だけです' : 'まだ誰も入っていません'));
      });
    }
  }
  lines.push(shortStaffAppeal_(few, info.people, allowHard));
  return lines.join('\n');
}

// 文の中の {…} を、その時の中身に置き換える。
// ready ＝ 何通も続けて作るときに、同じ数え直しをくり返さないための入れ物
//（「今の内容で見てみる」は送る日の数だけ文を作るので、これが無いと同じ集計を何回もやることになる）
// card ＝ 渡すとリンクを文から外してボタンにまわす（長いURLで文が読みにくくなるため）
function fillNoticeText_(template, ss, period, left, ready, card) {
  var text = String(template);
  var last = addDaysYmd_(period.start, period.weeks * 7 - 1);
  var done = ready || {};
  // リンクを文から外してボタンにまわす。
  // 「パソコンから→{パソコン}」のようにリンクだけの行は、前置きごと1行まるごと消す
  if (card) {
    text = text.split('\n').map(function(line) {
      var rest = line;
      if (rest.indexOf('{リンク}') !== -1) { card.link = true; rest = rest.split('{リンク}').join(''); }
      if (rest.indexOf('{パソコン}') !== -1) { card.pc = true; rest = rest.split('{パソコン}').join(''); }
      if (rest === line) return line;                       // リンクの無い行はそのまま
      // 「パソコンから→」のような短い前置きだけが残る行は、リンクのための行なので丸ごと消す
      if (rest.replace(/^[\s\u3000]+|[\s\u3000]+$/g, '').length <= NOTICE_LINK_LINE_MAX) return null;
      return rest.replace(/[\s\u3000]+$/, '');
    }).filter(function(line) { return line !== null; }).join('\n');
  }
  var map = {
    '{期間}': ymdLabel_(period.start) + '〜' + ymdLabel_(last),
    '{締切}': ymdLabel_(period.deadline.slice(0, 10)) + ' ' + period.deadline.slice(11),
    '{締切日}': ymdLabel_(period.deadline.slice(0, 10)),       // 時刻なし（1行目を短くしたい時に使う）
    '{リンク}': staffPageUrl_(getRequestSettingsSheet_(), '?p=' + period.start.slice(2).replace(/-/g, '')),
    // パソコン用。LINEログインを通らずに直接開く（LINEに入っていない画面からでも使える）。
    // authuser=0 は、Googleに複数ログインしているパソコンで「ファイルを開くことができません」に
    // なるのを防ぐため（Googleが勝手に2つ目のアカウントで開こうとする）
    '{パソコン}': requestFormUrl_() + '?p=' + period.start.slice(2).replace(/-/g, '') + '&authuser=0',
    '{残り日数}': String(left),
    '{店舗名}': getStoreName_()
  };
  if (card) { card.linkUrl = map['{リンク}']; card.pcUrl = map['{パソコン}']; }
  if (/\{未提出\}|\{途中\}|\{入力状況\}/.test(text)) {
    if (left > NOTICE_STATUS_WITHIN_DAYS) {
      // まだ日があるうちは人数を出さない（早いうちから急かさないため）
      map['{未提出}'] = '';
      map['{途中}'] = '';
      map['{入力状況}'] = '';
    } else {
      if (!done.waiting) done.waiting = waitingCounts_(ss, period);
      var waiting = done.waiting;
      var lines = [];
      if (waiting.notSubmitted) lines.push('まだ出していない人：' + waiting.notSubmitted + '人');
      if (waiting.partial) lines.push('途中の人（1週でも空きがある）：' + waiting.partial + '人');
      map['{未提出}'] = String(waiting.notSubmitted);
      map['{途中}'] = String(waiting.partial);
      map['{入力状況}'] = lines.length ? lines.join('\n') : 'みなさん提出ありがとうございます。変える方はお早めに。';
    }
  }
  if (/\{人手不足\}|\{お礼\}/.test(text)) {
    if (left > NOTICE_STATUS_WITHIN_DAYS) {
      // まだ日があるうちは、ほとんど入力が無くて「少なめの日」だらけになるので出さない
      map['{人手不足}'] = '';
      map['{お礼}'] = '';
    } else {
      if (!done.info) done.info = shortStaffInfo_(ss, period);     // 数えるのは1回だけ
      var short = shortStaffTextFrom_(done.info, left <= NEED_HARD_WITHIN_DAYS);
      map['{人手不足}'] = short;
      // 人が足りている時だけお礼。足りない日がある時は {人手不足} のほうが出る
      map['{お礼}'] = short ? '' : NOTICE_THANKS_TEXT;
    }
  }
  Object.keys(map).forEach(function(key) { text = text.split(key).join(map[key]); });
  return text.replace(/\n{3,}/g, '\n\n').replace(/\n+$/, '');   // 中身が空だった差し込みの空行を詰める
}

// お知らせをLINEへ送る形にする。リンクは長くて読みにくいので、文には出さずにボタンにする
function noticeMessage_(text, card) {
  var c = card || {};
  if (!c.link && !c.pc) return text;                 // リンクを使わない文は、今までどおりの文だけ
  var buttons = [];
  if (c.link && c.linkUrl) {
    buttons.push({ type: 'button', style: 'primary', height: 'sm',
      action: { type: 'uri', label: 'シフト希望を入力する', uri: c.linkUrl } });
  }
  if (c.pc && c.pcUrl) {
    // パソコン用も枠のあるボタンにする（文字だけだと押せると気づかれにくい）
    buttons.push({ type: 'button', style: 'secondary', height: 'sm',
      action: { type: 'uri', label: 'パソコンで入力する', uri: c.pcUrl } });
  }
  if (!buttons.length) return text;
  return {
    type: 'flex',
    altText: String(text).slice(0, 300),             // 通知に出る短い文（400字まで）
    contents: {
      type: 'bubble',
      size: 'mega',
      body: { type: 'box', layout: 'vertical', contents: [
        { type: 'text', text: String(text), wrap: true, size: 'sm', color: '#111111' }] },
      footer: { type: 'box', layout: 'vertical', spacing: 'sm', contents: buttons }
    }
  };
}

// 「今の内容で見てみる」で、ボタンの見え方も分かるようにする
function noticeButtonLabels_(card) {
  var out = [];
  if (card.link && card.linkUrl) out.push('［ボタン］シフト希望を入力する');
  if (card.pc && card.pcUrl) out.push('［ボタン］パソコンで入力する');
  return out;
}

// メニュー「お知らせの文面・日程」
function openNoticeSettings() {
  var settings = getRequestSettingsSheet_();
  var notice = readNoticeSettings_(settings);
  var box = 'width:100%;padding:6px;font-size:12px;';
  var html = '<div style="font-family:sans-serif;padding:4px 8px;font-size:14px;line-height:1.6;">' +
    '<label style="display:block;font-weight:bold;">送る日（スタッフ入力締切の何日前か。カンマ区切り。0＝締切当日）</label>' +
    '<input id="days" value="' + notice.days.join(',') + '" style="' + box + '">' +
    '<label style="display:block;font-weight:bold;margin-top:10px;">送る時刻</label>' +
    '<select id="hour" style="' + box + '">' + numberOptions_(0, 23, notice.hour, '#時ごろ') + '</select>' +
    '<label style="display:block;font-weight:bold;margin-top:10px;">1回目（いちばん早い日）の文</label>' +
    '<textarea id="first" style="' + box + 'height:90px;">' + escapeHtml_(notice.first) + '</textarea>' +
    '<label style="display:block;font-weight:bold;margin-top:10px;">2回目以降の文</label>' +
    '<textarea id="later" style="' + box + 'height:110px;">' + escapeHtml_(notice.later) + '</textarea>' +
    '<label style="display:block;font-weight:bold;margin-top:10px;">締切当日の文（送る日に「0」を入れると送ります）</label>' +
    '<textarea id="last" style="' + box + 'height:110px;">' + escapeHtml_(notice.last) + '</textarea>' +
    '<p style="margin:8px 0 0;font-size:12px;color:#555;">使える差し込み：<br>' +
    '{期間}／{締切}／{リンク}／{残り日数}／{未提出}／{途中}／{入力状況}／{人手不足}／{お礼}／{店舗名}<br>' +
    '{締切日}＝締切の日にちだけ（時刻なし）。1行目を短くしたいときに使います。<br>' +
    '{パソコン}＝パソコン用のリンク（LINEログインを通らない）<br>' +
    '　{リンク}と{パソコン}は、その行がまるごと<b>下のボタン</b>になります（長いURLを文に出しません）。<br>' +
    '{入力状況}＝「まだ出していない人：◯人」などを自動で作ります（全員そろっていればお礼の文になります）。<br>' +
    '　出るのは<b>締切まであと' + NOTICE_STATUS_WITHIN_DAYS + '日以内</b>のときだけです。それより前の日は何も出ません。<br>' +
    '{人手不足}＝人が少なめの日を知らせます（少ない日が4日以上なら曜日でまとめます）。足りている時は何も出ません。<br>' +
    '　これも出るのは<b>締切まであと' + NOTICE_STATUS_WITHIN_DAYS + '日以内</b>のときだけです（それより前はまだ入力が少ないため）。<br>' +
    '　「このままでは運営に支障が出ます」という強めの言い方になるのは、<b>締切まであと3日以内</b>で、なおかつ<br>' +
    '　ひとりあと1日以上足りないときだけです。それより前の日は「入れそうな方がいたら」の軽い声かけになります。<br>' +
    '{お礼}＝人がそろっている時だけ出ます（足りない日がある時は{人手不足}のほうが出ます）。締切当日の文に向いています。<br>' +
    '　{入力状況}{人手不足}と同じで、<b>締切まであと' + NOTICE_STATUS_WITHIN_DAYS + '日以内</b>の回だけ出ます。<br>' +
    '　欲しい人数の目安は「シフト希望設定」シートのE29（平日）・E30（土日祝）で変えられます。</p>' +
    '<p id="msg" style="margin:8px 0 0;font-size:12px;color:#186A3B;"></p>' +
    '<div style="margin-top:10px;text-align:right;">' +
    '<button onclick="doPreview()" style="margin-right:8px;padding:6px 12px;">今の内容で見てみる</button>' +
    '<button id="testBtn" onclick="doTest()" style="margin-right:8px;padding:6px 12px;">お試しに送ってみる</button>' +
    '<button id="saveBtn" onclick="doSave()" style="padding:6px 12px;">保存する</button>' +
    '</div>' +
    '<pre id="preview" style="white-space:pre-wrap;background:#F4F4F2;border-radius:6px;padding:8px;font-size:12px;margin-top:8px;"></pre>' +
    '</div>' +
    '<script>' +
    'function show(t, err){ var m = document.getElementById("msg"); m.textContent = t; m.style.color = err ? "#B03A2E" : "#186A3B"; }' +
    'function values(){ return { days: document.getElementById("days").value, hour: document.getElementById("hour").value,' +
    '  first: document.getElementById("first").value, later: document.getElementById("later").value,' +
    '  last: document.getElementById("last").value }; }' +
    'function doSave(){' +
    '  document.getElementById("saveBtn").disabled = true;' +
    '  google.script.run' +
    '    .withSuccessHandler(function(t){ show(t); document.getElementById("saveBtn").disabled = false; })' +
    '    .withFailureHandler(function(e){ show(e.message, true); document.getElementById("saveBtn").disabled = false; })' +
    '    .saveNoticeSettings(values());' +
    '}' +
    'function doPreview(){' +
    '  google.script.run' +
    '    .withSuccessHandler(function(t){ document.getElementById("preview").textContent = t; show("この内容で送られます（まだ送っていません）"); })' +
    '    .withFailureHandler(function(e){ show(e.message, true); })' +
    '    .previewNoticeText(values());' +
    '}' +
    'function doTest(){' +
    '  document.getElementById("testBtn").disabled = true;' +
    '  show("送っています…");' +
    '  google.script.run' +
    '    .withSuccessHandler(function(t){ show(t); document.getElementById("testBtn").disabled = false; })' +
    '    .withFailureHandler(function(e){ show(e.message, true); document.getElementById("testBtn").disabled = false; })' +
    '    .sendNoticeTest(values());' +
    '}' +
    '</script>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(520).setHeight(620), 'お知らせの文面・日程');
}

function saveNoticeSettings(values) {
  var settings = getRequestSettingsSheet_();
  var days = splitDays_(values.days);
  if (!days.length) throw new Error('送る日を1つ以上入れてください（例 14,7,3）。');
  days.forEach(function(n) {
    if (!(n >= 0 && n <= 60 && Math.floor(n) === n)) throw new Error('送る日は0〜60の整数で入れてください。');
  });
  var hour = Number(values.hour);
  if (!(hour >= 0 && hour <= 23)) throw new Error('送る時刻が正しくありません。');
  ['first', 'later', 'last'].forEach(function(key) {
    if (String(values[key] || '').length > 1000) throw new Error('文が長すぎます（1000文字まで）。');
  });

  settings.getRange(NOTICE_DAYS_CELL).setValue(days.sort(function(a, b) { return b - a; }).join(','));
  settings.getRange(NOTICE_HOUR_CELL).setValue(String(hour));
  settings.getRange(NOTICE_FIRST_CELL).setValue(String(values.first || ''));
  settings.getRange(NOTICE_LATER_CELL).setValue(String(values.later || ''));
  settings.getRange(NOTICE_LAST_CELL).setValue(String(values.last || ''));

  // 送る時刻を変えたら、毎日の予約も入れ直す（自動が動いている時だけ）
  var running = ScriptApp.getProjectTriggers().filter(function(t) { return t.getHandlerFunction() === 'dailyShiftCheck'; });
  if (running.length) {
    running.forEach(function(t) { ScriptApp.deleteTrigger(t); });
    ScriptApp.newTrigger('dailyShiftCheck').timeBased().atHour(hour).everyDays(1).create();
    return '保存しました。毎日' + hour + '時ごろに動くよう入れ直しました。';
  }
  return '保存しました。';
}

// 「今の内容で見てみる」：保存せずに、今の募集期間で文面を作って見せる
function previewNoticeText(values) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settings = getRequestSettingsSheet_();
  var period = currentRequestPeriod_(settings, new Date());
  if (!period) return '今は募集期間がありません。';
  var days = splitDays_(values.days).filter(function(n) { return n >= 0; }).sort(function(a, b) { return b - a; });
  var notice = {
    days: days,
    first: String(values.first || NOTICE_DEFAULT_FIRST),
    later: String(values.later || NOTICE_DEFAULT_LATER),
    last: String(values.last || NOTICE_DEFAULT_LAST)
  };
  var out = [];
  var ready = {};                        // 何通ぶんも作るので、人数の数え直しは1回で済ませる
  days.forEach(function(left) {
    out.push('― ' + (left === 0 ? '締切当日' : left + '日前') +
      '（' + ymdLabel_(addDaysYmd_(period.deadline.slice(0, 10), -left)) + '） ―');
    var card = {};
    out.push(fillNoticeText_(noticeTemplateFor_(notice, left), ss, period, left, ready, card));
    noticeButtonLabels_(card).forEach(function(s) { out.push(s); });
    out.push('');
  });
  return out.join('\n');
}

// 「お試しに送ってみる」：今書いている内容のまま、お試しの送り先へ1通だけ送る。
// 送るのは「今日が送る日ならその日の文、そうでなければ次に送る日の文」
function sendNoticeTest(values) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settings = getRequestSettingsSheet_();
  var period = currentRequestPeriod_(settings, new Date());
  if (!period) throw new Error('今は募集期間がありません。');
  var line = lineConfig_(settings, true);
  if (!line.token) throw new Error('チャネルアクセストークンが入っていません。');
  if (!line.target) throw new Error('お試しの送り先が登録されていません。メニュー「LINE通知の設定」から登録してください。');
  var days = splitDays_(values.days).filter(function(n) { return n >= 0; }).sort(function(a, b) { return b - a; });
  if (!days.length) throw new Error('送る日を1つ以上入れてください（例 14,7,3）。');
  var today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  var rest = daysBetweenYmd_(today, period.deadline.slice(0, 10));
  var soon = days.filter(function(n) { return n <= rest; });
  var left = days.indexOf(rest) !== -1 ? rest : (soon.length ? soon[0] : days[days.length - 1]);
  var notice = {
    days: days,
    first: String(values.first || NOTICE_DEFAULT_FIRST),
    later: String(values.later || NOTICE_DEFAULT_LATER),
    last: String(values.last || NOTICE_DEFAULT_LAST)
  };
  var card = {};
  var text = fillNoticeText_(noticeTemplateFor_(notice, left), ss, period, left, null, card);
  linePush_(line, noticeMessage_(text, card));
  return 'お試しの送り先に送りました（' + (left === 0 ? '締切当日' : left + '日前') +
    'の文）。スタッフのグループには送っていません。';
}


// ===== 毎月のおしらせ（メニュー「毎月のおしらせ」） =====
// 毎月の決まった日（例：25日）を期日にして、その何日前かにLINEのグループへ送る。
// 送る日ごとに別の文を書ける（例：14日前は「締切り2週間前」、3日前は「締切りまで3日」）。
// 文面の {月} {日} {曜日} {期日} {残り} {残り日数} {店舗名} は、送るときに中身へ置き換わる。
// 送る時刻はこのおしらせ専用（シフト希望のお知らせとは別の時刻にできる）。
var MONTHLY_ON_CELL = 'E17';      // 入／切
var MONTHLY_DAY_CELL = 'E18';     // 期日（毎月この日。月の日数を超える時はその月の末日）
var MONTHLY_DAYS_CELL = 'E19';    // 送る日（期日の何日前か。例「14,7,3」）
var MONTHLY_TEXT_CELL = 'E20';    // 文面（送る日ごとに「----」の行で区切って並べる）
var MONTHLY_HOUR_CELL = 'E21';    // 送る時刻（0〜23時）
var MONTHLY_DEFAULT_DAY = 25;
var MONTHLY_DEFAULT_HOUR = 10;
var MONTHLY_DEFAULT_DAYS = [14, 7, 3];
var MONTHLY_TEXT_SEP = '\n----\n';
var MONTHLY_DEFAULT_FIRST = '【期日厳守】\n検体キットが届いています。締切り{残り}前です。\n{月}月{日}日{曜日}までに提出と記入をお願いします。';
var MONTHLY_DEFAULT_TEXT = '【期日厳守】\n検体キットが届いています。締切りまで{残り}です。\n{月}月{日}日{曜日}までに提出と記入をお願いします。';
// 期日の当日に送る文（「締切りまで当日です」では変なので、当日だけ別の言い方にする）
var MONTHLY_DEFAULT_LAST = '【本日が期日】\n検体キットの締切りは本日です。\n{月}月{日}日{曜日}までに提出と記入をお願いします。';
// 1回目だけ「締切り2週間前です」、2回目からは「締切りまで1週間です」になる
var MONTHLY_DEFAULT_TEXTS = [MONTHLY_DEFAULT_FIRST, MONTHLY_DEFAULT_TEXT];

// ----- 検体キットの提出チェック -----
// 誰が出した・出していないは健康にかかわるので、LINEには人数しか出さない。
// 名前つきの一覧は店長のシート「提出チェック」だけに置く。
var KIT_SHEET_NAME = '提出チェック';
var KIT_MODE_CELL = 'E38';                 // 人数／出さない／使わない
var KIT_MODES = ['人数', '出さない', '使わない'];
var KIT_VIEW_HEADERS = ['番号', '名前', 'LINE登録', '提出不要', '日にち', 'だれが入れたか'];
var KIT_FIRST_ROW = 3;                      // 1行目＝見出しの帯、2行目＝列の名前、3行目から中身
var KIT_LINE_OWNER_SHEET = 'LINEの持ち主';  // LINEログインの登録が済んだ人（アクセス記録）
var KIT_OLD_HEADERS = ['月', '名前', '提出した日', 'だれが入れたか'];   // 前の形。1回だけの移しかえで読む
var KIT_OLD_LOG_COL = 7;                    // 前の形でG列に記録を置いていた分

// LINEログイン（LIFF）。ここにIDを入れると、お知らせのボタンがLINEログインのページを開く。
// 空なら今までどおり、ふつうのページを開く（LINEログインが動かないときの逃げ道）。
var LIFF_ID_CELL = 'E39';
var LINE_LOGIN_ID_CELL = 'E40';             // LINEログイン：チャネルID
var LINE_LOGIN_SECRET_CELL = 'E41';         // LINEログイン：チャネルシークレット
var LINE_AUTH_URL = 'https://access.line.me/oauth2/v2.1/authorize';

function liffId_(settings) {
  var id = String(settings.getRange(LIFF_ID_CELL).getDisplayValue()).trim();
  return /^[0-9]{8,12}-[0-9a-zA-Z]{6,12}$/.test(id) ? id : '';
}

// スタッフに送るページのURL。path は '?kit=2026-09' のような形。
// チャネルIDが入っていれば、LINEログインを通ってからページを開く形にする。
// （GASのページは枠の中で表示されるため、LIFFでは本人が分からない。2026-09-26に実測で確認）
function staffPageUrl_(settings, path) {
  // 数字で入っていると「2,011,734,637」と表示されることがあるので、数字だけを取り出す
  var loginId = String(settings.getRange(LINE_LOGIN_ID_CELL).getDisplayValue()).replace(/[^0-9]/g, '');
  if (/^[0-9]{8,12}$/.test(loginId)) {
    return LINE_AUTH_URL + '?response_type=code'
      + '&client_id=' + loginId
      + '&redirect_uri=' + encodeURIComponent(requestFormUrl_())
      + '&state=' + encodeURIComponent(path || '-')
      + '&scope=' + encodeURIComponent('profile openid');
  }
  var id = liffId_(settings);
  return id ? 'https://liff.line.me/' + id + path : requestFormUrl_() + path;
}

// メニュー「LINEログインの設定」
// LINE Developers の「LINEログイン」チャネルにある2つの値を入れる。
// 先にそのチャネルの「コールバックURL」に、下のアドレスを登録しておくこと。
function openLiffSettings() {
  var settings = getRequestSettingsSheet_();
  var ui = SpreadsheetApp.getUi();
  var nowId = String(settings.getRange(LINE_LOGIN_ID_CELL).getDisplayValue()).trim();
  var res = ui.prompt('LINEログインの設定（1/2）',
    'LINE Developers の「LINEログイン」チャネルにある'+'\n'+
    '「チャネルID」（数字だけ）を貼り付けてください。'+'\n\n'+
    '今の設定：' + (nowId || 'まだ入っていません') + '\n\n'+
    '※先に、そのチャネルの「コールバックURL」に'+'\n'+
    requestFormUrl_() + '\n'+
    'を登録しておいてください。'+'\n\n'+
    '何も入れずにOKを押すと、LINEログインを使わない状態に戻ります。',
    ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  var id = String(res.getResponseText()).trim();
  if (!id) {
    settings.getRange(LINE_LOGIN_ID_CELL).setValue('');
    settings.getRange(LINE_LOGIN_SECRET_CELL).setValue('');
    ui.alert('LINEログインを使わない状態に戻しました。今までどおりのページを開きます。');
    return;
  }
  if (!/^[0-9]{8,12}$/.test(id)) {
    throw new Error('チャネルIDの形がちがいます。数字だけ（例 2011734637）で入れてください。');
  }
  var res2 = ui.prompt('LINEログインの設定（2/2）',
    '同じ画面にある「チャネルシークレット」を貼り付けてください。'+'\n'+
    '（英数字が並んだ長い文字）'+'\n\n'+
    '今の設定：' + (String(settings.getRange(LINE_LOGIN_SECRET_CELL).getDisplayValue()).trim() ? '入っています' : 'まだ入っていません'),
    ui.ButtonSet.OK_CANCEL);
  if (res2.getSelectedButton() !== ui.Button.OK) return;
  var secret = String(res2.getResponseText()).trim();
  if (!/^[0-9a-zA-Z]{20,64}$/.test(secret)) {
    throw new Error('チャネルシークレットの形がちがいます。英数字の長い文字を貼り付けてください。');
  }
  settings.getRange(LINE_LOGIN_ID_CELL).setNumberFormat('@');   // 数字にすると桁区切りが付くため文字で持つ
  settings.getRange(LINE_LOGIN_ID_CELL).setValue(id);
  settings.getRange(LINE_LOGIN_SECRET_CELL).setValue(secret);
  ui.alert('LINEログインを使います。'+'\n\n'+
    'これからのお知らせのボタンは、はじめに1回だけ「許可」の画面が出て、'+'\n'+
    'そのあとは自動で自分の画面になります。'+'\n\n'+
    'まず「毎月のおしらせをお試しで送る」で、自分のスマホから確かめてください。');
}

// 何も選んでいないうちは「使わない」＝今までどおりの文だけ。
// 店長が設定で選んだときに始まるようにする（急に見え方が変わらないように）。
function kitMode_(settings) {
  var v = String(settings.getRange(KIT_MODE_CELL).getDisplayValue()).trim();
  return KIT_MODES.indexOf(v) >= 0 ? v : '使わない';
}

// 「提出チェック」シート（無ければ作る）。
// 置き場所は「シフト希望_アクセス記録」。飯塚店シフトはスタッフ全員が開くので、
// 誰が出した・出していないという健康にかかわる記録は置かない。
function getKitSheet_() {
  var ss = SpreadsheetApp.openById(recordSpreadsheetId_());
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

// 期日の年月（2026-09）。提出は月に1回なので、この単位で数える
function kitMonthOf_(ymd) { return String(ymd).slice(0, 7); }

// 書いてある月を '2026-09' の形にそろえる（日付で入っていても読めるように）
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
    map[flatName_(name)] = { name: name, exempt: String(r[2]).trim(),
      day: sameMonth ? String(r[3]).trim() : '', by: sameMonth ? String(r[4]).trim() : '' };
  });
  return map;
}

// その月にもう出した人 { そろえた名前: {name, day, by} }。別の月の一覧なら空
function kitMonthMap_(ym) {
  var sheet = getKitSheet_();
  var map = {};
  if (kitSheetMonth_(sheet) !== String(ym)) return map;
  var table = kitTableMap_(sheet, true);
  for (var key in table) if (table[key].day) map[key] = table[key];
  return map;
}

// その月にもう出した人 { そろえた名前: 書いてある名前 }
function kitDoneMap_(ym) {
  var hits = kitMonthMap_(ym);
  var map = {};
  for (var key in hits) map[key] = hits[key].name;
  return map;
}

// LINEログインの登録が済んでいる人
function kitLineRegistered_() {
  var map = {};
  try {
    var sheet = SpreadsheetApp.openById(recordSpreadsheetId_()).getSheetByName(KIT_LINE_OWNER_SHEET);
    if (!sheet) return map;
    var last = sheet.getLastRow();
    if (last < 2) return map;
    sheet.getRange(2, 2, last - 1, 1).getValues().forEach(function(r) {
      if (String(r[0]).trim()) map[flatName_(r[0])] = true;
    });
  } catch (e) {}
  return map;
}

// 期日の見出し「9月25日(金)」。フォーム側の kitDueLabel_ と同じ形にそろえる
function kitDueLabel_(ymd) {
  var p = ymd.split('-').map(Number);
  return p[1] + '月' + p[2] + '日(' + DOW_JP[dowOfYmd_(ymd)] + ')';
}

// 提出した日の見出し「9/25(金)」
function kitDayLabel_(d) {
  if (!(d instanceof Date)) return String(d == null ? '' : d);
  return (d.getMonth() + 1) + '/' + d.getDate() + '(' + DOW_JP[d.getDay()] + ')';
}

// 店長が見る一覧を書く。mark を渡すと、その人を提出ずみにしてから書く。
// 飯塚店シフトのメニューには出さない（スタッフ全員が開くファイルなので、この話は置かない）。
function kitWriteSummary_(ym, dueYmd, mark) {
  var sheet = getKitSheet_();
  var table = kitTableMap_(sheet, kitSheetMonth_(sheet) === String(ym));
  var already = false;
  if (mark && mark.name) {
    var markKey = flatName_(mark.name);
    var was = table[markKey];
    already = !!(was && was.day);
    if (!already) {
      table[markKey] = { name: String(mark.name).trim(), exempt: was ? was.exempt : '',
        day: kitDayLabel_(new Date()), by: String(mark.by || '本人') };
    }
  }
  var names = kitTargetNames_();
  var line = kitLineRegistered_();
  var rows = [];
  var exemptRows = [];
  var target = 0;
  var done = 0;
  var lineDone = 0;
  var offCount = 0;
  for (var i = 0; i < names.length; i++) {
    var key = flatName_(names[i]);
    var hit = table[key] || {};
    if (line[key]) lineDone++;
    if (hit.exempt) { offCount++; exemptRows.push(KIT_FIRST_ROW + i); }
    else { target++; if (hit.day) done++; }
    rows.push([i + 1, names[i], line[key] ? '○' : '', hit.exempt || '', hit.day || '', hit.by || '']);
  }
  sheet.getRange(1, 1).setNumberFormat('@');
  sheet.getRange(1, 1, 1, KIT_VIEW_HEADERS.length).setValues([[
    String(ym),
    dueYmd ? '期日 ' + kitDueLabel_(dueYmd) : String(sheet.getRange(1, 2).getValue() || ''),
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

// 提出チェックシートを今の内容に書き直す（毎日の点検から呼ぶ）
function kitRefreshSummary_(settings, today) {
  try {
    var monthly = readMonthlyNotice_(settings);
    kitWriteSummary_(today.slice(0, 7), monthlyDueYmd_(today, monthly.day, 0));
  } catch (e) {}                              // 書けなくても、ほかの点検は止めない
}

// 出してもらう人（いちばん新しい週シートの名簿にいる人）
function kitTargetNames_() {
  var newest = newestWeekSheet_(SpreadsheetApp.getActiveSpreadsheet());
  if (!newest) return [];
  return getRosterNames_(newest).filter(function(n) { return String(n).trim() !== ''; });
}

// その月の残りを数える。「提出不要」の人は数に入れない
function kitCount_(ym) {
  var sheet = getKitSheet_();
  var table = kitTableMap_(sheet, kitSheetMonth_(sheet) === String(ym));
  var names = kitTargetNames_().filter(function(n) { var t = table[flatName_(n)]; return !(t && t.exempt); });
  var left = names.filter(function(n) { var t = table[flatName_(n)]; return !(t && t.day); });
  return { total: names.length, done: names.length - left.length, left: left.length, leftNames: left };
}

// 提出を記録する。同じ月に二度は書かない
function kitMark_(ym, name, by) {
  var clean = String(name == null ? '' : name).trim();
  if (!/^\d{4}-\d{2}$/.test(String(ym)) || !clean) return false;
  return !kitWriteSummary_(ym, '', { name: clean, by: by || '本人' }).already;
}


// 毎月のおしらせのLINE。文と、残りの人数と、「出したらここを押す」ボタン
function monthlyMessage_(text, count, pageUrl) {
  if (!pageUrl) return text;                // 提出チェックを使わないときは、今までどおりの文だけ
  var body = [{ type: 'text', text: String(text), wrap: true, size: 'sm', color: '#111111' }];
  if (count && count.left > 0) {
    body.push({ type: 'text', wrap: true, size: 'sm', color: '#9A5518', margin: 'md',
      text: 'まだの人：' + count.left + '人（名前は出しません）' });
  }
  return {
    type: 'flex',
    altText: String(text).split('\n')[0] || '毎月のおしらせ',
    contents: {
      type: 'bubble',
      body: { type: 'box', layout: 'vertical', contents: body },
      footer: { type: 'box', layout: 'vertical', contents: [
        { type: 'button', style: 'primary', height: 'sm',
          action: { type: 'uri', label: '出したらここを押す', uri: pageUrl } }] }
    }
  };
}

// 送る日に合わせた文を並べ直す。書いていない日のぶんは最後の文を使いまわす。
// ただし当日のぶんに「締切りまで◯です」の初期文が当たると「締切りまで当日です」と変になるので、
// 自分で書いた文でなければ当日用に差し替える
function monthlyTextsFor_(days, texts) {
  return days.map(function(d, i) {
    var text = i < texts.length ? texts[i] : texts[texts.length - 1];
    if (d !== 0) return text;
    var isDefault = i >= texts.length || text === MONTHLY_DEFAULT_FIRST || text === MONTHLY_DEFAULT_TEXT;
    return isDefault ? MONTHLY_DEFAULT_LAST : text;
  });
}
// 前の版（送る日ごとに分けられなかった頃）の文。そのまま残っていたら今の初期値に読み替える
var MONTHLY_TEXT_V1 = '【期日厳守】\n検体キットが届いています。{月}月{日}日{曜日}までに提出と記入をお願いします。';

// 「----」の行で区切って、送る日ごとの文に分ける
function splitMonthlyTexts_(raw) {
  return String(raw == null ? '' : raw).split(/\r?\n-{3,}\r?\n/)
    .map(function(t) { return String(t).replace(/^[\r\n]+|[\r\n\s]+$/g, ''); })
    .filter(function(t) { return t !== ''; });
}

// 残りの言い方。14→「2週間」、7→「1週間」、3→「3日」、0→「当日」
function leftLabel_(left) {
  if (left === 0) return '当日';
  if (left % 7 === 0) return (left / 7) + '週間';
  return left + '日';
}

// 設定を読む。空や書き間違いのときは初期値を使う
function readMonthlyNotice_(settings) {
  var day = Number(String(settings.getRange(MONTHLY_DAY_CELL).getDisplayValue()).trim());
  var days = splitDays_(settings.getRange(MONTHLY_DAYS_CELL).getDisplayValue())
    .filter(function(n) { return n >= 0 && n <= 60 && Math.floor(n) === n; });
  days = days.filter(function(n, i) { return days.indexOf(n) === i; }).sort(function(a, b) { return b - a; });
  if (!days.length) days = MONTHLY_DEFAULT_DAYS.slice();

  var raw = String(settings.getRange(MONTHLY_TEXT_CELL).getValue());
  var texts = splitMonthlyTexts_(raw === MONTHLY_TEXT_V1 ? '' : raw);
  if (!texts.length) texts = MONTHLY_DEFAULT_TEXTS.slice();
  texts = monthlyTextsFor_(days, texts);       // 送る日の数にそろえる

  var hour = Number(String(settings.getRange(MONTHLY_HOUR_CELL).getDisplayValue()).trim());
  return {
    kit: kitMode_(settings),
    on: String(settings.getRange(MONTHLY_ON_CELL).getDisplayValue()).trim() === '入',
    day: day >= 1 && day <= 31 ? Math.floor(day) : MONTHLY_DEFAULT_DAY,
    days: days,
    texts: texts,
    hour: hour >= 0 && hour <= 23 ? Math.floor(hour) : MONTHLY_DEFAULT_HOUR
  };
}

// 今日から数えて monthOffset か月先の「期日」。その月に無い日（2月31日など）はその月の末日にする
function monthlyDueYmd_(todayYmd, day, monthOffset) {
  var p = todayYmd.split('-').map(Number);
  var month = p[1] + (monthOffset || 0);
  var lastDay = new Date(Date.UTC(p[0], month, 0)).getUTCDate();
  var d = new Date(Date.UTC(p[0], month - 1, Math.min(day, lastDay)));
  return d.getUTCFullYear() + '-' + ('0' + (d.getUTCMonth() + 1)).slice(-2) + '-' + ('0' + d.getUTCDate()).slice(-2);
}

// 文の中の {…} を、その時の中身に置き換える
function fillMonthlyText_(template, dueYmd, left) {
  var p = dueYmd.split('-').map(Number);
  var map = {
    '{月}': String(p[1]),
    '{日}': String(p[2]),
    '{曜日}': DOW_JP[dowOfYmd_(dueYmd)] + '曜日',
    '{期日}': ymdLabel_(dueYmd),
    '{残り日数}': String(left),
    '{残り}': leftLabel_(left),
    '{店舗名}': getStoreName_()
  };
  var text = String(template);
  Object.keys(map).forEach(function(key) { text = text.split(key).join(map[key]); });
  return text;
}

// 決めた時刻の予約から動く。今日が「送る日」でなければ何もしない
function sendMonthlyNotice_(settings, today) {
  var monthly = readMonthlyNotice_(settings);
  if (!monthly.on) return '';
  for (var k = 0; k <= 1; k++) {                    // 今月の期日 → 過ぎていれば来月の期日
    var due = monthlyDueYmd_(today, monthly.day, k);
    var left = daysBetweenYmd_(today, due);
    var index = monthly.days.indexOf(left);
    if (left < 0 || index === -1) continue;
    var key = 'month/' + due + '/' + left;
    if (noticeAlreadySent_(settings, key)) return '';
    // 提出チェック。全員そろった月は、残りのおしらせを送らない
    var ym = kitMonthOf_(due);
    var mode = kitMode_(settings);
    var count = mode === '使わない' ? null : kitCount_(ym);
    if (count && count.total > 0 && count.left === 0) {
      writeNoticeLog_(settings, key);
      return '';
    }
    var what = '毎月のおしらせ（' + ymdLabel_(due) + (left === 0 ? ' 当日' : 'の' + left + '日前') + '）';
    var line = lineConfig_(settings);
    if (!line.token || !line.target) {
      noticeSendFailure_(settings, what, line.token ? '送り先のグループが登録されていません。' : 'チャネルアクセストークンが入っていません。');
      return '';
    }
    var text = fillMonthlyText_(monthly.texts[index], due, left);
    try {
      linePush_(line, monthlyMessage_(text, mode === '人数' ? count : null, count ? staffPageUrl_(settings, '?kit=' + ym) : ''));
    } catch (e) {
      noticeSendFailure_(settings, what, e.message);
      return '';
    }
    writeNoticeLog_(settings, key);
    return text;
  }
  return '';
}

// 「毎月のおしらせ」の予約から動く入り口（決めた時刻に1日1回）
function monthlyNoticeCheck() {
  return sendMonthlyNotice_(getRequestSettingsSheet_(), Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd'));
}

// 「毎月のおしらせ」の予約を入れ直す（切なら予約を消す）
function resetMonthlyTrigger_(monthly) {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'monthlyNoticeCheck') ScriptApp.deleteTrigger(t);
  });
  if (!monthly.on) return false;
  ScriptApp.newTrigger('monthlyNoticeCheck').timeBased().atHour(monthly.hour).everyDays(1).create();
  return true;
}

// メニュー「毎月のおしらせ」
function openMonthlyNotice() {
  var monthly = readMonthlyNotice_(getRequestSettingsSheet_());
  var box = 'width:100%;padding:6px;font-size:12px;';
  var data = JSON.stringify({ days: monthly.days, texts: monthly.texts,
    def: MONTHLY_DEFAULT_TEXT, lastDef: MONTHLY_DEFAULT_LAST }).replace(/</g, '\\u003c');
  var html = '<div style="font-family:sans-serif;padding:4px 8px;font-size:14px;line-height:1.6;">' +
    '<p style="margin:0 0 8px;font-size:12px;color:#555;">毎月の決まった日を期日にして、その前にLINEのグループへお知らせを送ります。</p>' +

    '<label style="display:block;font-weight:bold;">送る／送らない</label>' +
    '<select id="on" style="' + box + '">' +
    '<option value="入"' + (monthly.on ? ' selected' : '') + '>送る</option>' +
    '<option value="切"' + (monthly.on ? '' : ' selected') + '>送らない</option></select>' +

    '<label style="display:block;font-weight:bold;margin-top:10px;">期日（毎月この日）</label>' +
    '<select id="day" style="' + box + '">' + numberOptions_(1, 31, monthly.day, '毎月#日') + '</select>' +

    '<label style="display:block;font-weight:bold;margin-top:10px;">送る日（期日の何日前か。カンマ区切り。0＝期日の当日）</label>' +
    '<input id="days" value="' + monthly.days.join(',') + '" oninput="build()" style="' + box + '">' +

    '<label style="display:block;font-weight:bold;margin-top:10px;">送る時刻</label>' +
    '<select id="hour" style="' + box + '">' + numberOptions_(0, 23, monthly.hour, '#時ごろ') + '</select>' +

    '<label style="display:block;font-weight:bold;margin-top:10px;">提出チェック（出した人が押すボタン）</label>' +
    '<select id="kit" style="' + box + '">' +
    KIT_MODES.map(function(m) {
      var label = m === '人数' ? 'ボタンを付けて、まだの人の数も出す'
        : (m === '出さない' ? 'ボタンを付けるが、人数は出さない' : 'ボタンを付けない（前のまま）');
      return '<option value="' + m + '"' + (monthly.kit === m ? ' selected' : '') + '>' + label + '</option>';
    }).join('') + '</select>' +
    '<p style="margin:4px 0 0;font-size:12px;color:#555;">名前はLINEに出しません。出した人の名前は「提出チェック」シートだけに残ります。' +
    '全員そろった月は、残りのおしらせを送りません。</p>' +

    '<div id="texts"></div>' +

    '<p style="margin:8px 0 0;font-size:12px;color:#555;">使える差し込み：<br>' +
    '{月}／{日}／{曜日}（木曜日）／{期日}（9/25(木)）／{残り}（2週間・1週間・3日・当日）／{残り日数}（数字だけ）／{店舗名}</p>' +
    '<p id="msg" style="margin:8px 0 0;font-size:12px;color:#186A3B;"></p>' +
    '<div style="margin-top:10px;text-align:right;">' +
    '<button onclick="doPreview()" style="margin-right:8px;padding:6px 12px;">今の内容で見てみる</button>' +
    '<button id="saveBtn" onclick="doSave()" style="padding:6px 12px;">保存する</button>' +
    '</div>' +
    '<pre id="preview" style="white-space:pre-wrap;background:#F4F4F2;border-radius:6px;padding:8px;font-size:12px;margin-top:8px;"></pre>' +
    '</div>' +
    '<script>' +
    'var DATA = ' + data + ';' +
    'var TA_STYLE = "' + box + 'height:90px;";' +
    'function show(t, err){ var m = document.getElementById("msg"); m.textContent = t; m.style.color = err ? "#B03A2E" : "#186A3B"; }' +
    // 送る日を書き換えると、文の欄も同じ数だけ作り直す（書いた文はできるだけ残す）
    'function build(){' +
    '  var wrap = document.getElementById("texts");' +
    '  var olds = [], byDay = {};' +
    '  Array.prototype.forEach.call(wrap.getElementsByTagName("textarea"), function(t){ byDay[t.getAttribute("data-day")] = t.value; olds.push(t.value); });' +
    '  if (!olds.length) { DATA.days.forEach(function(d, i){ byDay[d] = DATA.texts[i]; olds.push(DATA.texts[i]); }); }' +
    '  var days = (document.getElementById("days").value.match(/\\d+/g) || []).map(Number);' +
    '  days = days.filter(function(n, i){ return days.indexOf(n) === i; }).sort(function(a, b){ return b - a; });' +
    '  wrap.innerHTML = "";' +
    '  days.forEach(function(d, i){' +
    '    var lab = document.createElement("label");' +
    '    lab.style.display = "block"; lab.style.fontWeight = "bold"; lab.style.marginTop = "10px";' +
    '    lab.textContent = (d === 0 ? "期日の当日に送る文" : d + "日前に送る文");' +
    '    var ta = document.createElement("textarea");' +
    '    ta.setAttribute("data-day", d); ta.style.cssText = TA_STYLE;' +
    '    ta.value = byDay[d] !== undefined ? byDay[d]' +
    '      : (d === 0 ? DATA.lastDef' +
    '      : (olds[i] !== undefined ? olds[i] : (olds.length ? olds[olds.length - 1] : DATA.def)));' +
    '    wrap.appendChild(lab); wrap.appendChild(ta);' +
    '  });' +
    '}' +
    'function values(){' +
    '  var texts = Array.prototype.map.call(document.getElementById("texts").getElementsByTagName("textarea"), function(t){ return t.value; });' +
    '  return { on: document.getElementById("on").value, day: document.getElementById("day").value,' +
    '    days: document.getElementById("days").value, hour: document.getElementById("hour").value,' +
    '    kit: document.getElementById("kit").value, texts: texts };' +
    '}' +
    'function doSave(){' +
    '  document.getElementById("saveBtn").disabled = true;' +
    '  google.script.run' +
    '    .withSuccessHandler(function(t){ show(t); document.getElementById("saveBtn").disabled = false; })' +
    '    .withFailureHandler(function(e){ show(e.message, true); document.getElementById("saveBtn").disabled = false; })' +
    '    .saveMonthlyNotice(values());' +
    '}' +
    'function doPreview(){' +
    '  google.script.run' +
    '    .withSuccessHandler(function(t){ document.getElementById("preview").textContent = t; show("この内容で送られます（まだ送っていません）"); })' +
    '    .withFailureHandler(function(e){ show(e.message, true); })' +
    '    .previewMonthlyText(values());' +
    '}' +
    'build();' +
    '</script>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(520).setHeight(680), '毎月のおしらせ');
}

// 画面から来た値を、送る日と文の組にそろえる
function monthlyValues_(values) {
  var days = splitDays_(values.days);
  if (!days.length) throw new Error('送る日を1つ以上入れてください（例 14,7,3）。');
  days.forEach(function(n) {
    if (!(n >= 0 && n <= 60 && Math.floor(n) === n)) throw new Error('送る日は0〜60の整数で入れてください。');
  });
  days = days.filter(function(n, i) { return days.indexOf(n) === i; }).sort(function(a, b) { return b - a; });

  var texts = values.texts && values.texts.length ? values.texts.map(String)
            : splitMonthlyTexts_(values.text);
  texts = texts.map(function(t) { return String(t).replace(/^[\r\n]+|[\r\n\s]+$/g, ''); }).filter(function(t) { return t !== ''; });
  if (!texts.length) texts = MONTHLY_DEFAULT_TEXTS.slice();
  texts = monthlyTextsFor_(days, texts);
  texts.forEach(function(t) {
    if (t.length > 1000) throw new Error('文が長すぎます（1つにつき1000文字まで）。');
    if (t.indexOf('----') !== -1) throw new Error('文の中に「----」だけの行は入れられません（区切りに使っています）。');
  });
  return { days: days, texts: texts };
}

function saveMonthlyNotice(values) {
  var settings = getRequestSettingsSheet_();
  var day = Number(values.day);
  if (!(day >= 1 && day <= 31 && Math.floor(day) === day)) throw new Error('期日は1〜31の日を選んでください。');
  var hour = Number(values.hour);
  if (!(hour >= 0 && hour <= 23)) throw new Error('送る時刻が正しくありません。');
  var set = monthlyValues_(values);
  var on = String(values.on) === '入';

  settings.getRange(MONTHLY_ON_CELL).setValue(on ? '入' : '切');
  settings.getRange(MONTHLY_DAY_CELL).setValue(String(day));
  settings.getRange(MONTHLY_DAYS_CELL).setValue(set.days.join(','));
  settings.getRange(MONTHLY_TEXT_CELL).setValue(set.texts.join(MONTHLY_TEXT_SEP));
  settings.getRange(MONTHLY_HOUR_CELL).setValue(String(hour));
  settings.getRange(KIT_MODE_CELL).setValue(KIT_MODES.indexOf(String(values.kit)) >= 0 ? String(values.kit) : KIT_MODES[0]);

  resetMonthlyTrigger_({ on: on, hour: hour });
  return on ? '保存しました。毎月' + day + '日の' + set.days.join('日前・') + '日前、' + hour + '時ごろに送ります。'
            : '保存しました。今は送りません（「送る」にすると始まります）。';
}

// 「今の内容で見てみる」：保存せずに、次の期日で文面を作って見せる
function previewMonthlyText(values) {
  var today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  var day = Number(values.day) >= 1 && Number(values.day) <= 31 ? Math.floor(Number(values.day)) : MONTHLY_DEFAULT_DAY;
  var set = monthlyValues_(values);
  var due = monthlyDueYmd_(today, day, 0);
  if (daysBetweenYmd_(today, due) < 0) due = monthlyDueYmd_(today, day, 1);   // 今月の期日を過ぎていれば来月
  var out = ['次の期日：' + ymdLabel_(due), ''];
  set.days.forEach(function(left, i) {
    out.push('― ' + (left === 0 ? '当日' : left + '日前') + '（' + ymdLabel_(addDaysYmd_(due, -left)) + '） ―');
    out.push(fillMonthlyText_(set.texts[i], due, left));
    out.push('');
  });
  return out.join('\n');
}

// ===================== 出勤履歴と月次集計 =====================
// 週シートの右側（転写したあとの日別シフト表）から、1人1日ぶんの出勤を「出勤履歴」に積む。
// 月次集計は、店長・社員とキャストで締め日が違うので、区分ごとの締め期間で合計する。
var HISTORY_SHEET_NAME = '出勤履歴';
var HISTORY_HEADERS = ['日付', '曜日', '名前', '出勤', '退勤', '休入', '休終', '実働時間', '元シート'];
var HISTORY_SOURCE_COL = 9;   // 「元シート」列
var HISTORY_HOURS_COL = 8;    // 「実働時間」列
var MONTHLY_SHEET_NAME = '月次集計';
var MONTHLY_HEADER_ROW = 4;
var MONTHLY_FIRST_ROW = 5;
var MONTHLY_HEADERS = ['名前', '区分', '締め期間', '出勤日数', '合計時間', '締め開始', '締め終了'];
var WEEK_DOW_JP = ['月', '火', '水', '木', '金', '土', '日'];

function getStoreName_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REQUEST_SETTINGS_SHEET);
  return sheet ? String(sheet.getRange(REQUEST_STORE_NAME_CELL).getDisplayValue()).trim() : '';
}

// 「出勤履歴」シートを取得（無ければ作る）
function getHistorySheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(HISTORY_SHEET_NAME);
  if (sheet) return sheet;

  var active = ss.getActiveSheet(); // シート作成で表示中のシートが切り替わるため先に覚えておく
  sheet = ss.insertSheet(HISTORY_SHEET_NAME);
  sheet.getRange(1, 1, 1, HISTORY_HEADERS.length).setValues([HISTORY_HEADERS]).setFontWeight('bold').setBackground('#EFF0ED');
  sheet.setFrozenRows(1);
  sheet.getRange('A:A').setNumberFormat('yyyy/mm/dd');
  sheet.getRange('D:G').setNumberFormat('0":"00');
  sheet.getRange('I:I').setNumberFormat('@');   // 元シート名が数字にならないように
  ss.setActiveSheet(active);
  return sheet;
}

// 週シートの右側から、その週の出勤を1行ずつ取り出す（実働時間が0の人は入れない）
function buildWeekHistoryRows_(sheet) {
  var weekStart = sheet.getRange('T1').getValue();
  if (!(weekStart instanceof Date)) return [];
  var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  var startYmd = Utilities.formatDate(weekStart, tz, 'yyyy-MM-dd');
  var weekName = sheet.getName();
  var rows = [];

  DAY_BLOCKS.forEach(function(block, dayIndex) {
    var namesRange = sheet.getRange(block.names);
    var firstRow = namesRange.getRow();
    var firstCol = namesRange.getColumn();
    var count = namesRange.getNumRows();
    var info = sheet.getRange(firstRow, firstCol, count, 5).getValues();   // 名前・出勤・退勤・休入・休終
    var hours = sheet.getRange(block.hours).getValues();                    // 実働時間（15分刻みの合計）
    // 日付は夏時間の影響を受けないよう、日付だけで計算してからその日の0時にする
    var date = Utilities.parseDate(addDaysYmd_(startYmd, dayIndex) + ' 00:00', tz, 'yyyy-MM-dd HH:mm');

    for (var i = 0; i < count; i++) {
      var person = String(info[i][0]).trim();
      var worked = Number(hours[i][0]);
      if (!person || !(worked > 0)) continue;
      rows.push([date, WEEK_DOW_JP[dayIndex], person, info[i][1], info[i][2], info[i][3], info[i][4], worked, weekName]);
    }
  });
  return rows;
}

function sortHistoryRows_(rows) {
  rows.sort(function(a, b) {
    var ta = a[0] instanceof Date ? a[0].getTime() : 0;
    var tb = b[0] instanceof Date ? b[0].getTime() : 0;
    return ta === tb ? String(a[2]).localeCompare(String(b[2])) : ta - tb;
  });
  return rows;
}

function writeHistoryRows_(historySheet, rows) {
  var last = historySheet.getLastRow();
  if (last >= 2) historySheet.getRange(2, 1, last - 1, HISTORY_HEADERS.length).clearContent();
  if (rows.length) historySheet.getRange(2, 1, rows.length, HISTORY_HEADERS.length).setValues(rows);
}

// 1週ぶんの行を入れ替える（他の週・書庫へ移した週の行はそのまま残す）
// ★実働時間（CI・FS列）は =SUM(升目) の数式なので、名前と時間を書き込んだ直後は計算がまだ終わっていない。
//   先に flush で確定させないと「実働時間が0」と見なされ、その日が丸ごと履歴に入らない
//   ＝日別勤務時間が0になる（転写ボタンで12/17〜20が落ちた原因）。
//   ここに置けば、転写・手直し・次週シート作成・作り直しのどこから呼ばれても安全。
//   MARKER_HISTORY_FLUSH_FIRST
function refreshWeekHistory_(sheet) {
  if (!isWeekSheetName_(sheet.getName())) return 0;
  SpreadsheetApp.flush();
  var historySheet = getHistorySheet_();
  var weekName = sheet.getName();
  var last = historySheet.getLastRow();
  var kept = [];
  if (last >= 2) {
    kept = historySheet.getRange(2, 1, last - 1, HISTORY_HEADERS.length).getValues().filter(function(r) {
      return r[0] !== '' && String(r[HISTORY_SOURCE_COL - 1]) !== weekName;
    });
  }
  var rows = buildWeekHistoryRows_(sheet);
  writeHistoryRows_(historySheet, sortHistoryRows_(kept.concat(rows)));
  return rows.length;
}

// メニュー「出勤履歴を作り直す」：このファイルにある週シートを全部読み直す（書庫へ移した週の行は残す）
function rebuildAttendanceHistory() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var historySheet = getHistorySheet_();
  var weekSheets = weekSheets_(ss);
  var present = {};
  weekSheets.forEach(function(s) { present[s.getName()] = true; });

  var last = historySheet.getLastRow();
  var kept = [];
  if (last >= 2) {
    kept = historySheet.getRange(2, 1, last - 1, HISTORY_HEADERS.length).getValues().filter(function(r) {
      return r[0] !== '' && !present[String(r[HISTORY_SOURCE_COL - 1])];
    });
  }
  var rows = [];
  weekSheets.forEach(function(s) { rows = rows.concat(buildWeekHistoryRows_(s)); });
  writeHistoryRows_(historySheet, sortHistoryRows_(kept.concat(rows)));

  refreshMonthlySummary_();
  if (ss.getSheetByName(DAILY_SHEET_NAME)) refreshDailyHours_();
  SpreadsheetApp.getUi().alert('出勤履歴を作り直しました。\n週シート ' + weekSheets.length + '枚から ' + rows.length + '行' +
    (kept.length ? '（書庫へ移した週の ' + kept.length + '行はそのまま残しています）' : ''));
}

// 対象年月の選択肢（1年前〜来月まで。13番目＝今月）
function monthChoices_() {
  var months = [];
  var ym = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM').split('-').map(Number);
  for (var back = 12; back >= -1; back--) {
    var d = new Date(Date.UTC(ym[0], ym[1] - 1 - back, 1));
    months.push(d.getUTCFullYear() + '-' + ('0' + (d.getUTCMonth() + 1)).slice(-2));
  }
  return months;
}

// 「月次集計」シートを取得（無ければ作る）
function getMonthlySheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(MONTHLY_SHEET_NAME);
  if (sheet) return sheet;
  var active = ss.getActiveSheet();
  sheet = ss.insertSheet(MONTHLY_SHEET_NAME);
  ss.setActiveSheet(active);
  return sheet;
}

// 月次集計の見出し・締め期間の説明・名前の一覧と数式を作り直す
function refreshMonthlySummary_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getMonthlySheet_();
  var set = "'" + REQUEST_SETTINGS_SHEET + "'!";
  var hist = "'" + HISTORY_SHEET_NAME + "'!";
  var roles = "'" + ROLE_SHEET_NAME + "'!";

  sheet.getRange('A1').setFormula('=IF(' + set + REQUEST_STORE_NAME_CELL + '="","月次出勤集計",' + set + REQUEST_STORE_NAME_CELL + '&" 月次出勤集計")')
    .setFontWeight('bold').setFontSize(14);
  sheet.getRange('A2').setValue('対象年月');
  sheet.getRange('A3').setFormula('="' + ROLE_MANAGER + '・' + ROLE_EMPLOYEE + '：前月"&(VALUE(' + set + REQUEST_BASE_CLOSING_CELL + ')+1)&"日〜当月"&' +
    set + REQUEST_BASE_CLOSING_CELL + '&"日　／　キャスト：前月"&(VALUE(' + set + REQUEST_CAST_CLOSING_CELL + ')+1)&"日〜当月"&' +
    set + REQUEST_CAST_CLOSING_CELL + '&"日"');
  sheet.getRange(MONTHLY_HEADER_ROW, 1, 1, MONTHLY_HEADERS.length).setValues([MONTHLY_HEADERS])
    .setFontWeight('bold').setBackground('#EFF0ED');

  // 対象年月は、今月の前後から選べるようにする
  var months = monthChoices_();
  var current = String(sheet.getRange('B2').getDisplayValue()).trim();
  sheet.getRange('B2').setNumberFormat('@')
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(months, true).setAllowInvalid(false).build());
  if (months.indexOf(current) === -1) sheet.getRange('B2').setValue(months[12]); // 既定は今月

  var newest = newestWeekSheet_(ss);
  var names = newest ? getRosterNames_(newest) : [];
  var lastRow = sheet.getLastRow();
  if (lastRow >= MONTHLY_FIRST_ROW) {
    sheet.getRange(MONTHLY_FIRST_ROW, 1, lastRow - MONTHLY_FIRST_ROW + 1, MONTHLY_HEADERS.length).clearContent();
  }
  if (!names.length) return;

  var formulas = names.map(function(name, i) {
    var r = MONTHLY_FIRST_ROW + i;
    var closingDay = 'VALUE(IF(OR($B' + r + '="' + ROLE_MANAGER + '",$B' + r + '="' + ROLE_EMPLOYEE + '"),' +
      set + REQUEST_BASE_CLOSING_CELL + ',' + set + REQUEST_CAST_CLOSING_CELL + '))';
    var year = 'VALUE(LEFT($B$2,4))';
    var month = 'VALUE(MID($B$2,6,2))';
    return [
      '=IFERROR(VLOOKUP($A' + r + ',' + roles + '$A:$B,2,FALSE),"キャスト")',
      '=TEXT($F' + r + ',"M/d")&"〜"&TEXT($G' + r + ',"M/d")',
      '=COUNTIFS(' + hist + '$C:$C,$A' + r + ',' + hist + '$A:$A,">="&$F' + r + ',' + hist + '$A:$A,"<="&$G' + r + ',' + hist + '$H:$H,">0")',
      '=SUMIFS(' + hist + '$H:$H,' + hist + '$C:$C,$A' + r + ',' + hist + '$A:$A,">="&$F' + r + ',' + hist + '$A:$A,"<="&$G' + r + ')',
      '=DATE(' + year + ',' + month + '-1,' + closingDay + ')+1',
      '=DATE(' + year + ',' + month + ',' + closingDay + ')'
    ];
  });

  sheet.getRange(MONTHLY_FIRST_ROW, 1, names.length, 1).setValues(names.map(function(n) { return [n]; }));
  sheet.getRange(MONTHLY_FIRST_ROW, 2, names.length, MONTHLY_HEADERS.length - 1).setFormulas(formulas);
  sheet.getRange(MONTHLY_FIRST_ROW, 6, names.length, 2).setNumberFormat('yyyy/mm/dd');
}

// メニュー「月次集計を開く」
function openMonthlySummary() {
  refreshMonthlySummary_();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setActiveSheet(getMonthlySheet_());
}

/* ================= 日別の勤務時間（1か月・店長社員とそれ以外に分ける） =================
   「出勤履歴」（1人1日ぶん）を、名前×日の表に並べ直して見るためのシート。
   月次集計は合計だけなので、どの日に誰が何時間入っているかを一覧で見たいときに使う。
   期間は給与の締めと同じ。店長・社員は「前月16日〜当月15日」、
   それ以外（キャスト）は「前月21日〜当月20日」のように、区分ごとに区切りが違う。
   中身はぜんぶ数式なので、対象年月を変えればその場で入れ替わる。              */
var DAILY_SHEET_NAME = '日別勤務時間';
var DAILY_DAYS = 31;                       // 1か月でいちばん長い日数ぶんの列を用意する
var DAILY_FIRST_COL = 2;                   // B列から日が並ぶ
var DAILY_COUNT_COL = DAILY_FIRST_COL + DAILY_DAYS;   // 出勤日数（表の見出しの行では「期間の開始」も置く）
var DAILY_OFF_COL = DAILY_COUNT_COL + 1;              // 休み日数（　　　〃　　　　「期間の終了」）
var DAILY_HOURS_COL = DAILY_COUNT_COL + 2;            // 合計時間
// 休み日数 ＝ 期間の日数 − 出勤日数。出勤日数と足すと必ず期間の日数になる。MARKER_DAILY_OFF_DAYS
// 店長・社員の表だけ、1人につき「その日の出勤（開始）の時間」の行を名前の下に足す（確認用）。
// 入れるのは「9」「9:30」といった文字なので、いちばん下の合計には入らない（SUMは文字をとばす）。
// MARKER_DAILY_START_TIME
var DAILY_START_LABEL = '　開始';
var DAILY_FIRST_BLOCK_ROW = 5;             // 1つめの表の見出し行
var DAILY_BLOCK_GAP = 2;                   // 表と表のあいだの空き行
var DAILY_RANGE_CELL = 'B3';               // 区切り方を選ぶところ
// 日別の表は「暦（1日〜末日）」が既定。店長社員もそれ以外も同じ列（1日〜末日）で並べる。
// 給与の締めで見たいときだけ、もう一方に切り替える
var DAILY_RANGE_CALENDAR = '暦（1日〜末日）';
var DAILY_RANGE_CLOSING = '給与の締め期間';
var DAILY_SUNDAY_BG = '#FDEDEC';
var DAILY_SATURDAY_BG = '#EAF2F8';
// 運営ジャーナル用の縦の表（左に1日〜末日・右にパートと社員の時間）
var DAILY_JOURNAL_TITLE = '【運営ジャーナル用（縦）】';
var DAILY_JOURNAL_DOW_COL = 2;          // B列＝曜日
var DAILY_JOURNAL_PART_COL = 3;         // C列＝パート（それ以外）→ 運営ジャーナルのR列
var DAILY_JOURNAL_EMP_COL = 4;          // D列＝社員 → 運営ジャーナルのS列
var DAILY_JOURNAL_SUM_COL = 5;          // E列＝合計（見るだけ）

function getDailySheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(DAILY_SHEET_NAME);
  if (sheet) return sheet;
  var active = ss.getActiveSheet();   // シート作成で表示中のシートが変わるので先に覚えておく
  sheet = ss.insertSheet(DAILY_SHEET_NAME);
  ss.setActiveSheet(active);
  return sheet;
}

// 日別の表を作り直す
function refreshDailyHours_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getDailySheet_();
  var set = "'" + REQUEST_SETTINGS_SHEET + "'!";
  var hist = "'" + HISTORY_SHEET_NAME + "'!";
  var keepYm = String(sheet.getRange('B2').getDisplayValue()).trim();
  var keepRange = String(sheet.getRange(DAILY_RANGE_CELL).getDisplayValue()).trim();

  // 手で書き込むところが無いシートなので、毎回まるごと作り直す。
  // 結合は、部分的に選んでほどこうとするとエラーになるので1つずつほどく
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).getMergedRanges()
    .forEach(function(m) { try { m.breakApart(); } catch (e) {} });
  sheet.clear();
  sheet.setConditionalFormatRules([]);
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearDataValidations();

  sheet.getRange('A1')
    .setFormula('=IF(' + set + REQUEST_STORE_NAME_CELL + '="","日別の勤務時間",' +
      set + REQUEST_STORE_NAME_CELL + '&" 日別の勤務時間")')
    .setFontWeight('bold').setFontSize(14);

  // 対象年月
  var months = monthChoices_();
  sheet.getRange('A2').setValue('対象年月').setFontWeight('bold');
  sheet.getRange('B2:F2').merge();
  sheet.getRange('B2').setNumberFormat('@')
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(months, true).setAllowInvalid(false).build())
    .setValue(months.indexOf(keepYm) >= 0 ? keepYm : months[12]);
  sheet.getRange('H2').setValue('← シフトができている分だけ出ます');

  // 区切り方（締め期間／暦の1日〜末日）
  sheet.getRange('A3').setValue('区切り方').setFontWeight('bold');
  sheet.getRange('B3:F3').merge();
  sheet.getRange(DAILY_RANGE_CELL).setNumberFormat('@')
    .setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList([DAILY_RANGE_CALENDAR, DAILY_RANGE_CLOSING], true).setAllowInvalid(false).build())
    .setValue(keepRange === DAILY_RANGE_CLOSING ? DAILY_RANGE_CLOSING : DAILY_RANGE_CALENDAR);

  var calRef = '$' + DAILY_RANGE_CELL.charAt(0) + '$' + DAILY_RANGE_CELL.slice(1) + '="' + DAILY_RANGE_CALENDAR + '"';
  sheet.getRange('H3').setFormula('=IF(' + calRef +
    ',"← その月の1日〜末日で出します（店長・社員もそれ以外も同じ列）",' +
    '"← 給与の締めで区切ります（店長・社員は16日〜15日のように、区分で始まりが変わります）")');

  // 名簿の名前を、区分で2つに分ける
  var newest = newestWeekSheet_(ss);
  var names = (newest ? getRosterNames_(newest) : [])
    .map(function(n) { return String(n).trim(); }).filter(function(n) { return n; });
  var roles = getStaffRoles_();
  function isEmployee(name) {
    var kind = String(roles[name] || '').trim();
    return kind === ROLE_MANAGER || kind === ROLE_EMPLOYEE;
  }
  var blocks = [
    { title: '【' + ROLE_MANAGER + '・' + ROLE_EMPLOYEE + '】', closing: set + REQUEST_BASE_CLOSING_CELL,
      names: names.filter(isEmployee), withStart: true },
    { title: '【それ以外（キャスト）】', closing: set + REQUEST_CAST_CLOSING_CELL,
      names: names.filter(function(n) { return !isEmployee(n); }) }
  ];

  var year = 'VALUE(LEFT($B$2,4))', month = 'VALUE(MID($B$2,6,2))';
  var rules = [];
  var row = DAILY_FIRST_BLOCK_ROW;
  blocks.forEach(function(b) {
    // 名前のマスが何行目になるかは writeDailyBlock_ が b.nameRows に入れる（縦の表がここを見る）。
    // 「開始」の行が間に入るので、上から順に1行ずつとは限らない
    row = writeDailyBlock_(sheet, b, row, hist, rules, year, month, calRef) + DAILY_BLOCK_GAP;
  });
  writeDailyJournalBlock_(sheet, blocks, row, hist, rules, year, month);
  sheet.setConditionalFormatRules(rules);

  sheet.setFrozenColumns(1);
  sheet.setColumnWidth(1, 110);
  sheet.setColumnWidths(DAILY_FIRST_COL, DAILY_DAYS, 34);
  sheet.setColumnWidth(DAILY_COUNT_COL, 62);
  sheet.setColumnWidth(DAILY_OFF_COL, 62);
  sheet.setColumnWidth(DAILY_HOURS_COL, 70);
  return sheet;
}

// 1人ぶんの「その日の実働時間」の行。右はしに 出勤日数・休み日数・合計時間が付く
function dailyHoursFormulas_(r, dateRow, hist, startRef, endRef) {
  var out = [];
  for (var i = 0; i < DAILY_DAYS; i++) {
    var c = colLetter_(DAILY_FIRST_COL + i);
    out.push('=IF(' + c + '$' + dateRow + '="","",SUMIFS(' + hist + '$H:$H,' +
      hist + '$C:$C,$A' + r + ',' + hist + '$A:$A,' + c + '$' + dateRow + '))');
  }
  out.push('=COUNTIFS(' + hist + '$C:$C,$A' + r + ',' + hist + '$A:$A,">="&' + startRef +
    ',' + hist + '$A:$A,"<="&' + endRef + ',' + hist + '$H:$H,">0")');
  // 休み日数＝期間の日数−出勤日数（出勤日数のマスをそのまま引くので、ズレることがない）
  out.push('=' + endRef + '-' + startRef + '+1-' + colLetter_(DAILY_COUNT_COL) + r);
  out.push('=SUMIFS(' + hist + '$H:$H,' + hist + '$C:$C,$A' + r + ',' +
    hist + '$A:$A,">="&' + startRef + ',' + hist + '$A:$A,"<="&' + endRef + ')');
  return out;
}

// 1人ぶんの「その日の出勤（開始）の時間」の行。名前の行のすぐ下に出す。
// 出勤履歴の「出勤」は 900（＝9:00）のような数字。9時なら「9」、9時半なら「9:30」と出す。
// 同じ日に2回入っているときは、いちばん早い出勤を取る。
// 数字ではなく文字で入れるので、いちばん下の合計には入らない。MARKER_DAILY_START_TIME
function dailyStartFormulas_(nameRow, dateRow, hist) {
  var out = [];
  for (var i = 0; i < DAILY_DAYS; i++) {
    var c = colLetter_(DAILY_FIRST_COL + i);
    var min = 'MINIFS(' + hist + '$D:$D,' + hist + '$C:$C,$A$' + nameRow + ',' +
      hist + '$A:$A,' + c + '$' + dateRow + ',' + hist + '$D:$D,">0")';
    out.push('=IF(' + c + '$' + dateRow + '="","",IFERROR(LET(t,' + min +
      ',IF(t=0,"",TEXT(INT(t/100),"0")&IF(MOD(t,100)=0,"",":"&TEXT(MOD(t,100),"00")))),""))');
  }
  out.push('=""');   // 出勤日数
  out.push('=""');   // 休み日数
  out.push('=""');   // 合計時間
  return out;
}

// 1つぶんの表を書く。返り値は、その表の最後の行
// 期間（開始・終了）は、その表の見出しの行の右はし2マスに入れる。
// 区切り方が「暦」なら その月の1日〜末日、「締め期間」なら 前月の締め日の翌日〜当月の締め日。
function writeDailyBlock_(sheet, block, top, hist, rules, year, month, calRef) {
  var startRef = '$' + colLetter_(DAILY_COUNT_COL) + '$' + top;
  var endRef = '$' + colLetter_(DAILY_OFF_COL) + '$' + top;
  var dateRow = top + 1, dowRow = top + 2, firstRow = top + 3;
  var count = block.names.length;
  var rowsPerName = block.withStart ? 2 : 1;      // 上＝その日の実働時間／下＝出勤した時間
  var totalRow = firstRow + count * rowsPerName;

  sheet.getRange(top, DAILY_COUNT_COL).setFormula('=IF(' + calRef +
    ',DATE(' + year + ',' + month + ',1),DATE(' + year + ',' + month + '-1,VALUE(' + block.closing + '))+1)');
  sheet.getRange(top, DAILY_OFF_COL).setFormula('=IF(' + calRef +
    ',EOMONTH(DATE(' + year + ',' + month + ',1),0),DATE(' + year + ',' + month + ',VALUE(' + block.closing + ')))');
  sheet.getRange(top, DAILY_COUNT_COL, 1, 2)
    .setNumberFormat('m/d').setFontColor('#777777').setHorizontalAlignment('center');

  // 見出しは「題名＋期間」を1つのマスにまとめる。前は A=題名／B=期間 と分けていたので、
  // 【それ以外（キャスト）】のような長い題名が、中身のある隣のマス(B)に切られて読めなかった。
  // 1つにまとめると、右の空いたマスへはみ出して全部見える（この行には罫線を引いていない）。
  // MARKER_DAILY_TITLE_ONE_CELL
  sheet.getRange(top, 1).setFormula('="' + String(block.title).replace(/"/g, '""') +
    '　期間 "&TEXT(' + startRef + ',"yyyy/m/d")&"〜"&TEXT(' + endRef + ',"yyyy/m/d")')
    .setFontWeight('bold').setFontSize(12);
  sheet.getRange(top, 2).clearContent();   // 前の作りで入っていた「期間…」を消す（残すと題名が切れる）

  // 日付の行（本物の日付を入れて、表示は「日」だけにする）
  var dateF = [], dowF = [];
  for (var i = 0; i < DAILY_DAYS; i++) {
    var here = colLetter_(DAILY_FIRST_COL + i);
    var prev = colLetter_(DAILY_FIRST_COL + i - 1);
    dateF.push(i === 0 ? '=' + startRef
      : '=IF(' + prev + dateRow + '="","",IF(' + prev + dateRow + '+1>' + endRef + ',"",' + prev + dateRow + '+1))');
    dowF.push('=IF(' + here + dateRow + '="","",CHOOSE(WEEKDAY(' + here + dateRow + '),"日","月","火","水","木","金","土"))');
  }
  sheet.getRange(dateRow, 1).setValue('名前');
  sheet.getRange(dowRow, 1).setValue('曜日');
  sheet.getRange(dateRow, DAILY_FIRST_COL, 1, DAILY_DAYS).setFormulas([dateF]).setNumberFormat('d');
  sheet.getRange(dowRow, DAILY_FIRST_COL, 1, DAILY_DAYS).setFormulas([dowF]);
  sheet.getRange(dateRow, DAILY_COUNT_COL).setValue('出勤日数');
  sheet.getRange(dateRow, DAILY_OFF_COL).setValue('休み日数');
  sheet.getRange(dateRow, DAILY_HOURS_COL).setValue('合計時間');
  sheet.getRange(dateRow, 1, 2, DAILY_HOURS_COL)
    .setFontWeight('bold').setBackground('#EFF0ED').setHorizontalAlignment('center');
  sheet.getRange(dateRow, 1, 2, 1).setHorizontalAlignment('left');

  block.nameRows = [];                            // 名前のマスが何行目か（縦の表がここを見る）
  if (count) {
    var labels = [], body = [], startRows = [];
    block.names.forEach(function(name, k) {
      var r = firstRow + k * rowsPerName;
      block.nameRows.push(r);
      labels.push([name]);
      body.push(dailyHoursFormulas_(r, dateRow, hist, startRef, endRef));
      if (block.withStart) {
        labels.push([DAILY_START_LABEL]);
        body.push(dailyStartFormulas_(r, dateRow, hist));
        startRows.push(r + 1);
      }
    });
    sheet.getRange(firstRow, 1, labels.length, 1).setValues(labels);
    sheet.getRange(firstRow, DAILY_FIRST_COL, body.length, DAILY_DAYS + 3).setFormulas(body);
    startRows.forEach(function(r) {                // 時間の行と見分けがつくように、小さい灰色の字にする
      sheet.getRange(r, 1, 1, DAILY_HOURS_COL).setFontColor('#777777').setFontSize(9);
      sheet.getRange(r, DAILY_FIRST_COL, 1, DAILY_DAYS).setHorizontalAlignment('center');
    });
  } else {
    sheet.getRange(firstRow, 1).setValue('（この区分の人は名簿にいません）').setFontColor('#888888');
  }

  // いちばん下の合計（その日に店全体で何時間入っているか）
  sheet.getRange(totalRow, 1).setValue('合計').setFontWeight('bold');
  if (count) {
    var totals = [];
    for (var j = 0; j < DAILY_DAYS + 3; j++) {
      var tc = colLetter_(DAILY_FIRST_COL + j);
      totals.push('=SUM(' + tc + firstRow + ':' + tc + (totalRow - 1) + ')');
    }
    sheet.getRange(totalRow, DAILY_FIRST_COL, 1, DAILY_DAYS + 3).setFormulas([totals]);
  }
  sheet.getRange(totalRow, 1, 1, DAILY_HOURS_COL).setFontWeight('bold').setBackground('#F7F7F7');

  // 0は出さない（空いている日が見てすぐ分かるように）
  var bodyRows = totalRow - firstRow + 1;
  sheet.getRange(firstRow, DAILY_FIRST_COL, bodyRows, DAILY_DAYS).setNumberFormat('0.0;-0.0;');
  sheet.getRange(firstRow, DAILY_COUNT_COL, bodyRows, 2).setNumberFormat('0;-0;');
  sheet.getRange(firstRow, DAILY_HOURS_COL, bodyRows, 1).setNumberFormat('0.0;-0.0;');
  sheet.getRange(dateRow, 1, totalRow - dateRow + 1, DAILY_HOURS_COL)
    .setBorder(true, true, true, true, true, true, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);

  // 土日に色を付ける（日付の行を見るので、対象年月や区切り方を変えても自動で付け替わる）
  var area = sheet.getRange(dateRow, DAILY_FIRST_COL, totalRow - dateRow + 1, DAILY_DAYS);
  var head = colLetter_(DAILY_FIRST_COL) + '$' + dateRow;
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(' + head + '<>"",WEEKDAY(' + head + ')=1)')
    .setBackground(DAILY_SUNDAY_BG).setRanges([area]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(' + head + '<>"",WEEKDAY(' + head + ')=7)')
    .setBackground(DAILY_SATURDAY_BG).setRanges([area]).build());

  return totalRow;
}

/* 運営ジャーナル用の縦の表。
   左に1日〜末日が並び、右に「パート（それ以外）」と「社員」の時間が出る。
   運営ジャーナルは R=パート・S=社員 の順なので、C列とD列を2列まとめてそのまま写せる。
   運営ジャーナルは暦の月で付けるので、区切り方（B3）に関係なく いつも1日〜末日で出す。 */
function writeDailyJournalBlock_(sheet, blocks, top, hist, rules, year, month) {
  var head = top + 1, firstRow = top + 2, lastRow = firstRow + DAILY_DAYS - 1;
  var first = 'DATE(' + year + ',' + month + ',1)';
  var last = 'EOMONTH(' + first + ',0)';
  var emp = blocks[0];

  sheet.getRange(top, 1).setValue(DAILY_JOURNAL_TITLE +
    '　← パート（C列）を運営ジャーナルのR列、社員（D列）をS列へ。いつも1日〜末日で出します')
    .setFontWeight('bold').setFontSize(12);

  sheet.getRange(head, 1).setValue('日');
  sheet.getRange(head, DAILY_JOURNAL_DOW_COL).setValue('曜日');
  sheet.getRange(head, DAILY_JOURNAL_PART_COL).setValue('パート');
  sheet.getRange(head, DAILY_JOURNAL_EMP_COL).setValue('社員');
  sheet.getRange(head, DAILY_JOURNAL_SUM_COL).setValue('合計');
  sheet.getRange(head, 1, 1, DAILY_JOURNAL_SUM_COL)
    .setFontWeight('bold').setBackground('#EFF0ED').setHorizontalAlignment('center').setWrap(true);

  // 社員（店長・社員）は人数が少ないので、名前ごとのSUMIFSを足していく
  function empFormula(r) {
    if (!emp.names.length) return '=IF($A' + r + '="","",0)';
    var terms = [];
    for (var k = 0; k < emp.names.length; k++) {
      terms.push('SUMIFS(' + hist + '$H:$H,' + hist + '$C:$C,$A$' + emp.nameRows[k] +
        ',' + hist + '$A:$A,$A' + r + ')');
    }
    return '=IF($A' + r + '="","",' + terms.join('+') + ')';
  }

  var pc = colLetter_(DAILY_JOURNAL_PART_COL), ec = colLetter_(DAILY_JOURNAL_EMP_COL);
  var sc = colLetter_(DAILY_JOURNAL_SUM_COL);
  var body = [];
  for (var i = 0; i < DAILY_DAYS; i++) {
    var r = firstRow + i;
    body.push([
      '=IF(' + first + '+' + i + '>' + last + ',"",' + first + '+' + i + ')',
      '=IF($A' + r + '="","",CHOOSE(WEEKDAY($A' + r + '),"日","月","火","水","木","金","土"))',
      '=IF($A' + r + '="","",' + sc + r + '-' + ec + r + ')',          // パート＝その日ぜんぶ−社員
      empFormula(r),
      '=IF($A' + r + '="","",SUMIFS(' + hist + '$H:$H,' + hist + '$A:$A,$A' + r + '))'   // その日ぜんぶ
    ]);
  }
  sheet.getRange(firstRow, 1, DAILY_DAYS, DAILY_JOURNAL_SUM_COL).setFormulas(body);
  sheet.getRange(firstRow, 1, DAILY_DAYS, 2).setHorizontalAlignment('center');
  sheet.getRange(firstRow, 1, DAILY_DAYS, 1).setNumberFormat('d');
  sheet.getRange(firstRow, DAILY_JOURNAL_PART_COL, DAILY_DAYS, 3).setNumberFormat('0.0;-0.0;');
  sheet.getRange(head, 1, DAILY_DAYS + 1, DAILY_JOURNAL_SUM_COL)
    .setBorder(true, true, true, true, true, true, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);

  // 月の合計は、マスが細いと ##### になってしまうので、文字で下に出す
  sheet.getRange(lastRow + 1, 1).setFormula('="月合計　パート "&TEXT(SUM(' + pc + firstRow + ':' + pc + lastRow +
    '),"0.0")&" 時間　／　社員 "&TEXT(SUM(' + ec + firstRow + ':' + ec + lastRow + '),"0.0")&" 時間"')
    .setFontWeight('bold');

  // 土日に色を付ける（日付のマスを見るので、対象年月を変えても自動で付け替わる）
  var area = sheet.getRange(firstRow, 1, DAILY_DAYS, DAILY_JOURNAL_SUM_COL);
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($A' + firstRow + '<>"",WEEKDAY($A' + firstRow + ')=1)')
    .setBackground(DAILY_SUNDAY_BG).setRanges([area]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($A' + firstRow + '<>"",WEEKDAY($A' + firstRow + ')=7)')
    .setBackground(DAILY_SATURDAY_BG).setRanges([area]).build());

  return lastRow + 1;
}

// メニュー「日別の勤務時間を開く」
function openDailyHours() {
  var sheet = refreshDailyHours_();
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
}

function newestWeekSheet_(ss) {
  var names = ss.getSheets().map(function(s) { return s.getName(); })
    .filter(function(n) { return /^\d{6}$/.test(n); }).sort();
  return names.length ? ss.getSheetByName(names[names.length - 1]) : null;
}

// 週シートの右側（転写した所）を手で直した時も、出勤履歴に自動で反映する
// ※承認のいらない簡易トリガー。左側の入力表(A〜S列)を直しただけの時は何もしない
function onEdit(e) {
  var sheet;
  try {
    if (!e || !e.range) return;
    sheet = e.range.getSheet();
    if (!isWeekSheetName_(sheet.getName())) return;
    if (e.range.getColumn() + e.range.getNumColumns() - 1 < 20) return;
    SpreadsheetApp.flush();            // 升目の計算を確定させてから色を見る
  } catch (err) { return; }
  // 3つを別々に守る。前は1つにまとめていたので、重い「出勤履歴」で時間切れになると
  // 色の塗り直しも決まりの確認も丸ごと動かず、休憩の水色や前の週の赤が残っていた。
  // 軽くて目に見えるものから先に動かす。MARKER_ONEDIT_STEPS
  autoStep_('時間の直し', function() {
    if (fixTypedTimes_(sheet, e.range)) SpreadsheetApp.flush();   // 時刻で入ったものを1600の形に戻す
  });
  autoStep_('休憩の終わりを入れる', function() {
    if (fillBreakEnd_(sheet, e.range)) SpreadsheetApp.flush();
  });
  autoStep_('色の塗り直し', function() { repaintEditedRows_(sheet, e.range); });
  autoStep_('決まりの確認', function() { checkEditedLaw_(sheet, e.range); });
  autoStep_('出勤履歴の更新', function() { refreshWeekHistory_(sheet); });
}

// ===================== 古い週を書庫へ移す =====================
// 直近 ARCHIVE_KEEP_WEEKS 週だけを今のファイルに残し、それより古い週シートは書庫ファイルへ移す。
// 移す前に出勤履歴へ取り込み、数式を値に変えるので、別ファイルになっても数字は壊れない。
var ARCHIVE_URL_CELL = 'E8';     // シフト希望設定シート：書庫ファイルのURL
var ARCHIVE_KEEP_WEEKS = 13;     // 今のファイルに残す直近の週数
// 1回に移す枚数の上限（GASは1回の実行が6分までのため、多い時は何回かに分ける）
var ARCHIVE_MAX_AUTO = 5;        // 次週シート作成のときに自動で移す枚数
var ARCHIVE_MAX_MANUAL = 20;     // メニューから移す枚数

// 週シート（6桁「261019」と、前の期の4桁「0929」の両方。T1に日付があるものだけ）
function weekSheets_(ss) {
  return ss.getSheets().filter(function(s) {
    return /^\d{4}$|^\d{6}$/.test(s.getName()) && s.getRange('T1').getValue() instanceof Date;
  });
}

function isWeekSheetName_(name) {
  return /^\d{4}$|^\d{6}$/.test(name);
}

// ----- 書庫は暦年ごとに分ける -----
// 週シートは「その週の年」の書庫へ入れる。年が変わって新しい書庫が要るときは、
// 自動で作って、メールと画面でお知らせする（作った書庫のURLは設定シートにも入れ直す）。
var ARCHIVE_YEAR_PROP = 'ARCHIVE_URL_';   // 文書プロパティ ARCHIVE_URL_2026

// URLから書庫を開く（開けなければ null）
function openArchiveByUrl_(ss, url) {
  var id = /[-\w]{25,}/.exec(String(url || ''));
  if (!id) return null;
  if (id[0] === ss.getId()) throw new Error('書庫ファイルに、このシフト表と同じURLが設定されています。');
  return SpreadsheetApp.openById(id[0]);
}

// 設定シートに入っている書庫のURL
function currentArchiveUrl_(ss) {
  var settings = ss.getSheetByName(REQUEST_SETTINGS_SHEET);
  return settings ? String(settings.getRange(ARCHIVE_URL_CELL).getDisplayValue()).trim() : '';
}

function archiveNameForYear_(ss, year) {
  return String(ss.getName()).replace(/\s*書庫\s*\d*\s*$/, '') + '書庫' + year;
}

// そのシートは何年の週か（T1の日付→名前の頭2桁→今年 の順で見る）
function archiveYearOf_(ss, sheet) {
  try {
    var t1 = sheet.getRange('T1').getValue();
    if (t1 instanceof Date) return Number(Utilities.formatDate(t1, ss.getSpreadsheetTimeZone(), 'yyyy'));
  } catch (e) {}
  var m = /^(\d{2})\d{4}/.exec(sheet.getName());
  if (m) return 2000 + Number(m[1]);
  return Number(Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy'));
}

// 今まで使っていた書庫を、その年のものとして覚えておく（名前に年が無ければ fallbackYear）
// ※これをしないと、年をまたいで移すときに、既にある書庫を見失って二重に作ってしまう
function adoptCurrentArchive_(ss, fallbackYear) {
  var url = currentArchiveUrl_(ss);
  if (!url) return;
  var book = openArchiveByUrl_(ss, url);
  if (!book) return;
  var m = /(20\d{2})/.exec(String(book.getName()));
  var year = m ? Number(m[1]) : fallbackYear;
  if (!(year > 2000)) return;
  var props = PropertiesService.getDocumentProperties();
  if (!props.getProperty(ARCHIVE_YEAR_PROP + year)) {
    props.setProperty(ARCHIVE_YEAR_PROP + year, url);
  }
}

// その年の書庫を返す。無ければ作る（作ったときは created:true）
function archiveForYear_(ss, year) {
  var props = PropertiesService.getDocumentProperties();
  var saved = props.getProperty(ARCHIVE_YEAR_PROP + year);
  if (saved) {
    var book = openArchiveByUrl_(ss, saved);
    if (book) return { book: book, created: false };
  }
  var made = SpreadsheetApp.create(archiveNameForYear_(ss, year));
  props.setProperty(ARCHIVE_YEAR_PROP + year, made.getUrl());
  var settings = ss.getSheetByName(REQUEST_SETTINGS_SHEET);
  if (settings) settings.getRange(ARCHIVE_URL_CELL).setValue(made.getUrl());
  noticeArchiveCreated_(year, made);
  return { book: made, created: true };
}

// 新しい書庫を作ったことを知らせる
function noticeArchiveCreated_(year, made) {
  var body = year + '年の週シートを入れる書庫を、新しく作りました。\n\n' +
    '名前：' + made.getName() + '\n' +
    'URL：' + made.getUrl() + '\n\n' +
    'マイドライブの中にできています。好きなフォルダへ移してかまいません' +
    '（場所ではなくURLで覚えているので、移しても大丈夫です）。\n\n' +
    'シフト希望設定シートの「書庫ファイルのURL」も、この新しい書庫に入れ替えました。\n' +
    '前の年の書庫はそのまま残っています。いつでも開けます。';
  try {
    MailApp.sendEmail(Session.getEffectiveUser().getEmail(),
      '【シフト表】' + year + '年の書庫を作りました', body);
  } catch (e) {
    Logger.log('書庫のお知らせメールに失敗: ' + e.message);
  }
}

// 書庫へ移す対象：古い週シートと、「〜のコピー」という残り物のシート
function archiveTargets_(ss) {
  var weeks = weekSheets_(ss).map(function(s) {
    return { sheet: s, start: Utilities.formatDate(s.getRange('T1').getValue(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd') };
  });
  weeks.sort(function(a, b) { return a.start < b.start ? 1 : (a.start > b.start ? -1 : 0); }); // 新しい順
  var old = weeks.slice(ARCHIVE_KEEP_WEEKS).map(function(w) { return w.sheet; });
  var leftovers = ss.getSheets().filter(function(s) { return /のコピー$/.test(s.getName()); });
  return old.concat(leftovers);
}

// 何枚かの週シートの出勤を、まとめて履歴に取り込む（1枚ずつ書き直すより速い）
function refreshHistoryForSheets_(sheets) {
  var historySheet = getHistorySheet_();
  var names = {};
  sheets.forEach(function(s) { names[s.getName()] = true; });
  var last = historySheet.getLastRow();
  var kept = [];
  if (last >= 2) {
    kept = historySheet.getRange(2, 1, last - 1, HISTORY_HEADERS.length).getValues().filter(function(r) {
      return r[0] !== '' && !names[String(r[HISTORY_SOURCE_COL - 1])];
    });
  }
  var rows = [];
  sheets.forEach(function(s) { rows = rows.concat(buildWeekHistoryRows_(s)); });
  writeHistoryRows_(historySheet, sortHistoryRows_(kept.concat(rows)));
  return rows.length;
}

// 1枚を書庫へ移す（コピーできたことを確かめてから元を消す）。履歴への取り込みは呼ぶ側で先に済ませる
function archiveSheet_(ss, archive, sheet) {
  var name = sheet.getName();

  // 数式を値に変える（別ファイルへ移ると、他のシートを見ている数式は壊れるため）
  var range = sheet.getDataRange();
  range.setValues(range.getValues());

  var copyName = name;
  for (var i = 2; archive.getSheetByName(copyName); i++) copyName = name + '(' + i + ')';
  var copied = sheet.copyTo(archive).setName(copyName);
  if (copied.getLastRow() !== sheet.getLastRow() || copied.getLastColumn() !== sheet.getLastColumn()) {
    throw new Error('コピーの大きさが元と違います');
  }
  ss.deleteSheet(sheet);
  return copyName;
}

// 古い週を書庫へ移す本体。silent=true なら画面に出さずに行う（次週シート作成から呼ぶ）
function archiveOldWeeks_(silent) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var all = archiveTargets_(ss);
  if (!all.length) return { moved: [], failed: [], message: '移す週はありません' };

  if (!currentArchiveUrl_(ss)) {
    return { moved: [], failed: [], message: '書庫ファイルが設定されていないため、' + all.length + '枚をそのまま残しました' };
  }

  var limit = silent ? ARCHIVE_MAX_AUTO : ARCHIVE_MAX_MANUAL;
  var targets = all.slice(0, limit);
  var rest = all.length - targets.length;

  // 移す週の出勤を、先にまとめて履歴へ取り込む
  var weekTargets = targets.filter(function(s) { return isWeekSheetName_(s.getName()); });
  if (weekTargets.length) refreshHistoryForSheets_(weekTargets);

  // 今まで使っていた書庫を、その年のものとして先に覚えておく（二重に作らないため）
  try {
    adoptCurrentArchive_(ss, archiveYearOf_(ss, targets[targets.length - 1]));
  } catch (e) {
    return { moved: [], failed: [], message: e.message };
  }

  var moved = [], failed = [], madeYears = [];
  targets.forEach(function(sheet) {
    var name = sheet.getName();
    try {
      var got = archiveForYear_(ss, archiveYearOf_(ss, sheet));
      if (got.created) madeYears.push(got.book.getName());
      moved.push(archiveSheet_(ss, got.book, sheet));
    } catch (e) {
      failed.push(name + '（' + e.message + '）');
    }
  });

  var message = moved.length ? moved.length + '枚を書庫へ移しました：' + moved.join('、') : '移した週はありません';
  if (madeYears.length) {
    message = '新しい書庫を作りました：' + madeYears.join('、') +
      '\n（マイドライブにあります。くわしくはメールを見てください）\n' + message;
  }
  if (rest > 0) message += '\n残り' + rest + '枚は次回に移します（1回で移す枚数を制限しています）。';
  if (failed.length) message += '\n移せなかったシート：' + failed.join('、');
  if (!silent) SpreadsheetApp.getUi().alert(message);
  else if (moved.length || failed.length) ss.toast(message, 'シフト表の整理', 8);
  return { moved: moved, failed: failed, message: message };
}

// メニュー「古い週を書庫へ移す」
function archiveOldWeeks() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var targets = archiveTargets_(ss);
  if (!targets.length) {
    ui.alert('今のファイルには直近' + ARCHIVE_KEEP_WEEKS + '週分しかないので、移すシートはありません。');
    return;
  }
  var names = targets.slice(0, ARCHIVE_MAX_MANUAL).map(function(s) { return s.getName(); }).join('、');
  var answer = ui.alert('古い週を書庫へ移す',
    '古いシートは全部で' + targets.length + '枚です。このうち' + Math.min(targets.length, ARCHIVE_MAX_MANUAL) +
    '枚を書庫ファイルへ移して、このファイルから消します。\n\n' + names +
    '\n\n出勤の数字は「出勤履歴」に残るので、月次集計は今までどおり出ます。移してよいですか？',
    ui.ButtonSet.OK_CANCEL);
  if (answer !== ui.Button.OK) return;
  archiveOldWeeks_(false);
}
