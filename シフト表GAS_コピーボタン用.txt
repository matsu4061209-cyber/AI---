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

  // 7-2. ポジションの色分けは新しい週に持ち越さない（水色の自動ルールを戻して、色を消す）
  resetPositionColorsOnCopy_(currentSheet, newSheet);

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
    return sheet.getRange(POSITION_NAME_ROW, 5, POSITION_COUNT, 1).getValues()
      .map(function(r, i) { return String(r[0]).trim() || POSITION_DEFAULTS[i].name; });
  } catch (e) {
    return POSITION_DEFAULTS.map(function(p) { return p.name; });
  }
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var posNames = positionNamesForMenu_();
  ui.createMenu('スタッフ管理')
    .addItem('スタッフを追加', 'addStaffMember')
    .addItem('スタッフを削除', 'removeStaffMember')
    .addItem('時間を固定するスタッフを設定', 'setProtectedStaff')
    .addItem('店長・社員を設定', 'setStaffRoles')
    .addSeparator()
    .addItem('シフト希望の募集設定', 'openShiftRequestSettings')
    .addSeparator()
    .addItem('月次集計を開く', 'openMonthlySummary')
    .addItem('売上予算を入れ直す', 'refreshWeekBudgets')
    .addItem('目安人数を開く', 'openTargetSheet')
    .addItem('目安人数を週シートに入れ直す', 'refreshWeekTargets')
    .addSeparator()
    .addItem('調整前として記録する（いまを基準にする）', 'recordBeforeAdjust')
    .addItem('確定する（薄い表示を消す）', 'confirmAdjust')
    .addItem('色を塗り直す（掃除）', 'refreshWeekColors')
    .addSubMenu(ui.createMenu('ポジション色分け')
      .addItem('色分けを始める（このシート）', 'unlockPositionColors')
      .addSeparator()
      .addItem(posNames[0] + 'に塗る', 'paintPosition1')
      .addItem(posNames[1] + 'に塗る', 'paintPosition2')
      .addItem(posNames[2] + 'に塗る', 'paintPosition3')
      .addItem(posNames[3] + 'に塗る', 'paintPosition4')
      .addItem('選んだところの色を戻す', 'clearPositionPaint')
      .addSeparator()
      .addItem('名前と色を変える', 'openPositionSettings')
      .addSeparator()
      .addItem('ボタンを画面に置く', 'placePositionButtons')
      .addItem('ボタンを消す', 'clearPositionButtons')
      .addItem('色分けをやめる（元に戻す）', 'relockPositionColors'))
    .addSeparator()
    .addItem('LINE通知の設定', 'openLineSettings')
    .addItem('お知らせの文面・日程', 'openNoticeSettings')
    .addItem('毎月のおしらせ', 'openMonthlyNotice')
    .addItem('自動でシート作成・お知らせを始める', 'startShiftAutomation')
    .addItem('自動を止める', 'stopShiftAutomation')
    .addItem('今すぐ点検する', 'runShiftCheckNow')
    .addItem('出勤履歴を作り直す', 'rebuildAttendanceHistory')
    .addItem('古い週を書庫へ移す', 'archiveOldWeeks')
    .addToUi();
}

// ===================== 店長・社員・キャストの時間数 =====================

// 「スタッフ区分」シートを取得（無ければ非表示で作る）
function getRoleSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(ROLE_SHEET_NAME);
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
  roleSheet.getRange('D1:D2').setValues([['数式の自動チェック'], [result]]);
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
function addStaffMember() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
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
function removeStaffMember() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
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
    '<p>削除する人を選んでください</p>' +
    '<select id="targetRow" style="width:100%;padding:6px;font-size:14px;">' + options + '</select>' +
    '<div style="margin-top:16px;text-align:right;">' +
    '<button onclick="google.script.host.close()" style="margin-right:8px;padding:6px 12px;">キャンセル</button>' +
    '<button onclick="doDelete()" style="padding:6px 12px;">削除する</button>' +
    '</div>' +
    '</div>' +
    '<script>' +
    'function doDelete(){' +
    '  var row = document.getElementById("targetRow").value;' +
    '  google.script.run.withSuccessHandler(function(){ google.script.host.close(); }).deleteStaffRow(Number(row));' +
    '}' +
    '</script>';

  var output = HtmlService.createHtmlOutput(html).setWidth(320).setHeight(180);
  ui.showModalDialog(output, 'スタッフを削除');
}

// removeStaffMember のプルダウンから呼ばれる削除の実処理
// ★末尾に _ を付けると google.script.run から呼べなくなるため付けない
// 削除した行より下の名簿(A〜P列)を1行ずつ上に詰め、最後の行を空にする
function deleteStaffRow(row) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRosterRow = ROSTER_START_ROW + ROSTER_ROWS - 1;

  if (row < lastRosterRow) {
    // copyTo は値だけでなく「時刻の表示形式」などの書式もまとめて移動するため、
    // getValues/setValues（値だけ）ではなくこちらを使う
    var sourceRange = sheet.getRange(row + 1, 1, lastRosterRow - row, 16);
    var destRange = sheet.getRange(row, 1, lastRosterRow - row, 16);
    sourceRange.copyTo(destRange);
  }
  sheet.getRange(lastRosterRow, 1, 1, 16).clearContent();
}

// ===================== スマホのシフト希望フォームとの連携 =====================
// フォーム（別のGAS）もこの「シフト希望設定」シートを同じ並びで読み書きする。並びを変える時は両方を直すこと
var REQUEST_SETTINGS_SHEET = 'シフト希望設定';
var REQUEST_PENDING_HEADER_ROW = 8;
var REQUEST_PENDING_FIRST_ROW = 9;   // まだ週シートが無い週の希望
var REQUEST_PENDING_COLS = 19;       // シート名・名前・時間14・備考・送信日時・端末ID
var REQUEST_INPUT_COLS = 15;         // 週シートのB〜P列
var REQUEST_FORM_URL = 'https://script.google.com/macros/s/AKfycbwFtgsWnq75UnttAODig9q3c89E1Iek0Oow-lV8tHzF_Zp5qBQe2DXPahAeksqkd4i2/exec';
var DOW_JP = ['日', '月', '火', '水', '木', '金', '土'];

// 募集期間の決まり。B3:B5＝今回だけ手で決めた期間（空なら自動）、E3:E5＝基本設定（店舗ごとの決まり）
var REQUEST_BASE_CLOSING_CELL = 'E3';   // 締め日（毎月この日。社員の締め日）
var REQUEST_BASE_WEEKS_CELL = 'E4';     // スタッフ入力締切＝募集が始まる何週間前の日曜日
var REQUEST_BASE_TIME_CELL = 'E5';      // 締切の時刻
var REQUEST_STORE_NAME_CELL = 'E6';     // 店舗名（見出しなどに使う）
var REQUEST_CAST_CLOSING_CELL = 'E7';   // キャストの締め日（月次集計に使う）
var REQUEST_DEFAULT_BASE = { closingDay: 15, deadlineWeeks: 2, deadlineTime: '23:59', castClosingDay: 20 };

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
    sheet.getRange('A7').setValue('まだ週シートが無い週の希望（「次週シート作成」でシートを作ると自動でシートに入り、ここから消えます）');
    var headers = ['シート名', '名前'];
    ['月', '火', '水', '木', '金', '土', '日'].forEach(function(d) { headers.push(d + ' 出勤', d + ' 退勤'); });
    headers.push('備考', '送信日時', '端末ID');
    sheet.getRange(REQUEST_PENDING_HEADER_ROW, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    ss.setActiveSheet(active);
    sheet.hideSheet();
  }
  // 基本設定の欄。空の項目だけ初期値を入れる（前の版で作ったシートにも足す）
  sheet.getRange('E3:E25').setNumberFormat('@');
  sheet.getRange('F22:F25').setNumberFormat('@');
  [[REQUEST_BASE_CLOSING_CELL, String(REQUEST_DEFAULT_BASE.closingDay)],
   [REQUEST_BASE_WEEKS_CELL, String(REQUEST_DEFAULT_BASE.deadlineWeeks)],
   [REQUEST_BASE_TIME_CELL, REQUEST_DEFAULT_BASE.deadlineTime],
   [REQUEST_CAST_CLOSING_CELL, String(REQUEST_DEFAULT_BASE.castClosingDay)],
   [MONTHLY_ON_CELL, '切'],
   [MONTHLY_DAY_CELL, String(MONTHLY_DEFAULT_DAY)],
   [MONTHLY_DAYS_CELL, MONTHLY_DEFAULT_DAYS.join(',')],
   [MONTHLY_TEXT_CELL, MONTHLY_DEFAULT_TEXT],
   [MONTHLY_HOUR_CELL, String(MONTHLY_DEFAULT_HOUR)]].forEach(function(pair) {
    if (String(sheet.getRange(pair[0]).getDisplayValue()).trim() === '') sheet.getRange(pair[0]).setValue(pair[1]);
  });
  // スタッフ入力締切の決め方を「期間の最終日の◯日前」から「募集が始まる◯週間前の日曜日」に変えたため、
  // 前の版で入っていた1〜4以外の数字（14など）は初期値に直す
  var weeksNow = Number(sheet.getRange(REQUEST_BASE_WEEKS_CELL).getDisplayValue());
  if (!(weeksNow >= 1 && weeksNow <= 4)) sheet.getRange(REQUEST_BASE_WEEKS_CELL).setValue(String(REQUEST_DEFAULT_BASE.deadlineWeeks));
  sheet.getRange('D3:D25').setValues([
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
    ['ポジション4：名前（E列）と色（F列）']
  ]);
  // ポジションの名前と色。空のところだけ初期値を入れる
  POSITION_DEFAULTS.forEach(function(pos, i) {
    var nameCell = sheet.getRange(POSITION_NAME_ROW + i, 5);
    var colorCell = sheet.getRange(POSITION_NAME_ROW + i, 6);
    if (String(nameCell.getDisplayValue()).trim() === '') nameCell.setValue(pos.name);
    if (String(colorCell.getDisplayValue()).trim() === '') colorCell.setValue(pos.color);
  });
  sheet.getRange('A3:A5').setValues([['今回だけの開始週（空なら自動）'], ['今回だけの週数'], ['今回だけのスタッフ入力締切']]);
  return sheet;
}

// 新しく作った週シートに、フォームから先に送られていたその週の希望を入れ、保留から消す
function applyPendingRequests_(sheet) {
  var settings = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REQUEST_SETTINGS_SHEET);
  if (!settings) return;
  var last = settings.getLastRow();
  if (last < REQUEST_PENDING_FIRST_ROW) return;

  var rows = settings.getRange(REQUEST_PENDING_FIRST_ROW, 1, last - REQUEST_PENDING_FIRST_ROW + 1, REQUEST_PENDING_COLS).getValues();
  var roster = sheet.getRange(ROSTER_START_ROW, 1, ROSTER_ROWS, 1).getValues()
    .map(function(r) { return String(r[0]).trim(); });

  for (var i = rows.length - 1; i >= 0; i--) {
    var name = String(rows[i][1]).trim();
    if (String(rows[i][0]) !== sheet.getName() || !name) continue;
    var index = roster.indexOf(name);
    if (index === -1) continue; // 名簿に名前が無い希望は消さずに残す
    sheet.getRange(ROSTER_START_ROW + index, 2, 1, REQUEST_INPUT_COLS).setValues([rows[i].slice(2, 2 + REQUEST_INPUT_COLS)]);
    settings.deleteRow(REQUEST_PENDING_FIRST_ROW + i);
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
    castClosingDay: cast >= 1 && cast <= 31 ? Math.floor(cast) : REQUEST_DEFAULT_BASE.castClosingDay
  };
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
  var last = sheet.getLastRow();
  if (last >= REQUEST_PENDING_FIRST_ROW) {
    sheet.getRange(REQUEST_PENDING_FIRST_ROW, 1, last - REQUEST_PENDING_FIRST_ROW + 1, 2).getValues().forEach(function(r) {
      if (r[0] && r[1]) pendingCounts[String(r[0])] = (pendingCounts[String(r[0])] || 0) + 1;
    });
  }
  var pendingText = Object.keys(pendingCounts).sort().map(function(n) { return n + '：' + pendingCounts[n] + '人'; }).join('、');

  // スタッフに送るリンクと、LINEにそのまま貼れる文（期間と締切が入る）。
  // リンクの末尾に期間を付けるのは、同じリンクだとLINEが前の見出しを覚えていて古い期間を出すため
  var linkUrl = REQUEST_FORM_URL + (shown ? '?p=' + shown.start.slice(2).replace(/-/g, '') : '');
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

  // 月次集計の見出し・締め期間も新しい設定に合わせる
  try {
    if (SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MONTHLY_SHEET_NAME)) refreshMonthlySummary_();
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
  sheet.getRange('A:A').setNumberFormat('yyyy/mm/dd');
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
  if (created) {
    ss.setActiveSheet(sheet);
    ui.alert('「売上予算」シートを作りました。\n\n曜日別の基準額（4行目）と、変える日だけ日別の予算（8行目から）を入れてください。\n次に週シートを作るときから、5行目の予算に自動で入ります。');
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
  ui.alert(changed ? sheets + '枚のシートで、' + changed + '日ぶんの予算を入れ直しました。' : '入れ直す予算はありませんでした（すでに同じ数字です）。');
}

// ===================== 目安人数 =====================
// 週シートの各曜日にある「目安」の行（9:00から30分ごとに30マス）を、
// 「目安人数」シートに登録した1時間ごとの人数から自動で入れる。
// 週シートの目安は30分ぶんの人時なので、人数の半分を入れる（3人 → 1.5）。
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
    sheet.getRange(TARGET_OPEN_ROW, 1).setValue('人を置く時間（開始〜終了。終了の時間には置きません）');
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

  // 表題と注意書き
  sheet.getRange(1, 1, 1, cols).breakApart().merge()
    .setBackground('#1C4587').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(13)
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 30);
  sheet.getRange(2, 1, 1, cols).breakApart().merge()
    .setFontColor('#666666').setFontSize(10).setWrap(true).setVerticalAlignment('middle');
  sheet.setRowHeight(2, 32);

  // 設定のところ（3〜6行目）
  sheet.getRange(TARGET_SPH_ROW, 1, 4, 1).setBackground(TARGET_LABEL_COLOR).setFontWeight('bold');
  sheet.getRange(TARGET_SPH_ROW, 2, 3, 1).setBackground(TARGET_INPUT_COLOR);
  sheet.getRange(TARGET_OPEN_ROW, 2, 1, 2).setBackground(TARGET_INPUT_COLOR);
  sheet.getRange(TARGET_SPH_ROW, 2, 2, 1).setNumberFormat('#,##0"円"');
  sheet.getRange(TARGET_OPEN_ROW, 2, 1, 2).setNumberFormat('0"時"');
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
      return [autoPeople_(avg, ratio, sph, minPeople, hour >= open.from && hour < open.to)];
    });
    sheet.getRange(TARGET_FIRST_ROW, c + 1, TARGET_HOURS, 1).setValues(autos);
  });
  return averages;
}

// 人を置く時間（開始〜終了。終了の時間には置かない）
function targetOpenHours_(sheet) {
  var from = Number(sheet.getRange(TARGET_OPEN_ROW, 2).getValue());
  var to = Number(sheet.getRange(TARGET_OPEN_ROW, 3).getValue());
  if (!(from >= 0 && from <= 23)) from = TARGET_DEFAULT.openFrom;
  if (!(to >= 1 && to <= 24) || to <= from) to = TARGET_DEFAULT.openTo;
  return { from: from, to: to };
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
    byCat: byCat
  };
}

// 1枚の週シートの「目安」を入れ直す。
// 目安のマスは結合されていることがあるので、変わるマスだけ1つずつ書く。
function applyWeekTargets_(sheet, targets, keys, tz) {
  var top = sheet.getRange('T1').getValue();
  if (!(top instanceof Date)) return 0;
  var start = Utilities.formatDate(top, tz, 'yyyy-MM-dd');
  var changed = 0;
  TARGET_BLOCKS.forEach(function(block, i) {
    var hours = targets.byCat[categoryOfYmd_(addDaysYmd_(start, i), keys)];
    if (!hours) return;
    var now = sheet.getRange(block.row, block.col, 1, TARGET_SLOTS * 2 - 1).getValues()[0];
    for (var s = 0; s < TARGET_SLOTS; s++) {
      var people = Number(hours[Math.floor(s / 2)]);
      var want = people > 0 ? people / 2 : '';          // 30分ぶんの人時
      if (String(now[s * 2]) === String(want)) continue;
      sheet.getRange(block.row, block.col + s * 2).setValue(want);
      changed++;
    }
  });
  return changed;
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
    '・自動人数：1日の平均売上×構成比÷目標人時売上高（人を置く時間は最低人数を下回りません）\n' +
    '・手打ち人数：実際に置く人数。週シートにはこちらが入ります（空なら自動人数）\n\n' +
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

// メニュー「調整前として記録する」
function recordBeforeAdjust() {
  var week = activeWeekSheet_();
  if (!week) return;
  var ui = SpreadsheetApp.getUi();
  if (ui.alert('「' + week.getName() + '」の今の時間を「調整前」として記録します。\n\n' +
      'このあと削った時間が、薄いグレーで残って見えるようになります。\nよろしいですか。',
      ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  clearAdjustMarks_(week);          // 前の薄い表示が残っていれば先に消す
  saveAdjustSnapshots_(week);
  repaintWeekColors_(week);
  ui.alert('いまの時間を「調整前」として記録しました。\nこれから削った時間は、その場で薄いグレーになります。');
}

// メニュー「確定する（薄い表示を消す）」
function confirmAdjust() {
  var week = activeWeekSheet_();
  if (!week) return;
  var ui = SpreadsheetApp.getUi();
  if (ui.alert('「' + week.getName() + '」を確定します。\n\n' +
      '薄いグレーの表示を消して、今の時間を新しい「調整前」として記録し直します。\nよろしいですか。',
      ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  setPen_(week, null);                 // 確定したらペンは置く
  var cleared = clearAdjustMarks_(week, ourColors_());
  saveAdjustSnapshots_(week);
  repaintWeekColors_(week);
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '確定しました。薄い表示を' + cleared + 'マス消して、今の時間を新しい「調整前」にしました。', 'シフト表', 6);
}


// ===================== ポジションの色分け =====================
// 時間帯バーの水色は「値が0.25なら#00B0F0」という条件付き書式で付いている。
// 条件付き書式は直接塗った色より強いので、そのままでは人がいるマスに色を塗っても見えない。
// そこで「色分けを始める」で、そのシートの水色ルールだけを外して控え、
// 代わりに人がいるマスを水色で直接塗る。以後はポジションごとに塗り分けられる。
// 「色分けをやめる」で、ルールを戻して元どおりにする（シート1枚ずつ・いつでも戻せる）。
// シートの条件付き書式に入っている水色。この色のルールを外して色分けを始める（変えない）
var POSITION_RULE_COLOR = '#00B0F0';
// 色分け中に「まだポジションを決めていない時間」を塗る色。ポジションの色と見分けやすい薄い水色
var POSITION_BASE_COLOR = '#CCEEFF';
var POSITION_COUNT = 4;
var POSITION_NAME_ROW = 22;                 // E22〜E25 名前、F22〜F25 色
var POSITION_DEFAULTS = [
  { name: 'デシャップ',   color: '#E06666' },
  { name: 'キッチン',     color: '#4A86E8' },
  { name: 'ウォッシャー', color: '#A9D18E' },
  { name: '店長',         color: '#B4A7D6' }
];

// ポジションの名前と色を読む
function positionList_() {
  var settings = getRequestSettingsSheet_();
  var values = settings.getRange(POSITION_NAME_ROW, 5, POSITION_COUNT, 2).getValues();
  return values.map(function(r, i) {
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

// 色分け中かどうか
// 印（A・B列）だけでなく、外した水色ルールの中身（D列）が入っていることまで確かめる。
// 印だけ残っていると「色分け中のつもりなのに水色が復活している」状態になり、
// 塗っても条件付き書式に隠れて何も起きないように見えるため。
function positionUnlockRow_(week) {
  var sheet = getAdjustSheet_();
  var row = adjustRowMap_(sheet)[week.getName() + '/cf'];
  if (!row) return 0;
  if (String(sheet.getRange(row, 4).getValue()).trim() === '') return 0;   // 中身が無い印は無視する
  return row;
}

// メニュー「色分けを始める（このシート）」
function unlockPositionColors() {
  var week = activeWeekSheet_();
  if (!week) return;
  var ui = SpreadsheetApp.getUi();
  if (positionUnlockRow_(week)) { ui.alert('「' + week.getName() + '」はすでに色分けを始めています。'); return; }
  if (ui.alert('「' + week.getName() + '」でポジションの色分けを始めます。\n\n' +
      '・時間帯バーの水色を、自動（条件付き書式）から手で塗る形に切り替えます\n' +
      '・削った時間の薄いグレーは、そのまま一緒に使えます\n' +
      '・「色分けをやめる」でいつでも元どおりに戻せます\n\nよろしいですか。',
      ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;

  var rules = week.getConditionalFormatRules();
  var keep = [], removed = [];
  rules.forEach(function(rule) {
    var bc = rule.getBooleanCondition();
    if (bc && String(bc.getCriteriaType()) === 'NUMBER_EQUAL_TO' &&
        Number(bc.getCriteriaValues()[0]) === 0.25 &&
        String(bc.getBackground()).toUpperCase() === POSITION_RULE_COLOR) {
      removed.push({
        ranges: rule.getRanges().map(function(r) { return r.getA1Notation(); }),
        value: Number(bc.getCriteriaValues()[0]),
        bg: String(bc.getBackground()),
        fc: bc.getFontColor() ? String(bc.getFontColor()) : ''
      });
    } else {
      keep.push(rule);
    }
  });
  if (!removed.length) { ui.alert('水色のルールが見つかりませんでした。何も変えていません。'); return; }
  week.setConditionalFormatRules(keep);

  saveUnlockRecord_(week, removed);
  repaintWeekColors_(week);
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '色分けを始めました。時間帯のマスを選んでから、ポジションのボタンを押してください。', 'シフト表', 6);
}

// 確認を出さずに色分けを始める（ポジションのボタンから呼ぶ）
function startPositionColors_(week) {
  var rules = week.getConditionalFormatRules();
  var keep = [], removed = [];
  rules.forEach(function(rule) {
    var bc = rule.getBooleanCondition();
    if (bc && String(bc.getCriteriaType()) === 'NUMBER_EQUAL_TO' &&
        Number(bc.getCriteriaValues()[0]) === 0.25 &&
        String(bc.getBackground()).toUpperCase() === POSITION_RULE_COLOR) {
      removed.push({
        ranges: rule.getRanges().map(function(r) { return r.getA1Notation(); }),
        value: Number(bc.getCriteriaValues()[0]),
        bg: String(bc.getBackground()),
        fc: bc.getFontColor() ? String(bc.getFontColor()) : ''
      });
    } else {
      keep.push(rule);
    }
  });
  if (!removed.length) return false;
  week.setConditionalFormatRules(keep);
  saveUnlockRecord_(week, removed);
  repaintWeekColors_(week);
  return true;
}

function saveUnlockRecord_(week, removed) {
  var sheet = getAdjustSheet_();
  var row = adjustRowMap_(sheet)[week.getName() + '/cf'];
  if (!row) {
    row = Math.max(sheet.getLastRow() + 1, 2);
    sheet.getRange(row, 1, 1, 2).setValues([[week.getName(), 'cf']]);
  }
  sheet.getRange(row, 3, 1, 2).setValues([
    [Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm'), JSON.stringify(removed)]
  ]);
}

// メニュー「色分けをやめる（元に戻す）」
function relockPositionColors() {
  var week = activeWeekSheet_();
  if (!week) return;
  var ui = SpreadsheetApp.getUi();
  var row = positionUnlockRow_(week);
  if (!row) { ui.alert('「' + week.getName() + '」は色分けを始めていません。'); return; }
  if (ui.alert('「' + week.getName() + '」の色分けをやめて、元の自動の水色に戻します。\n\n' +
      '塗ったポジションの色は消えます。（削った時間の薄いグレーは残ります）\nよろしいですか。',
      ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;

  var sheet = getAdjustSheet_();
  var removed = [];
  try { removed = JSON.parse(String(sheet.getRange(row, 4).getValue())); } catch (e) { removed = []; }
  var rules = week.getConditionalFormatRules();
  removed.forEach(function(item) {
    var builder = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberEqualTo(item.value)
      .setBackground(item.bg)
      .setRanges(item.ranges.map(function(a1) { return week.getRange(a1); }));
    if (item.fc) builder.setFontColor(item.fc);
    rules.push(builder.build());
  });
  week.setConditionalFormatRules(rules);
  // 控えの行ごと空にする（これで「色分け中」の印が消える）
  sheet.getRange(row, 1, 1, 5).setValues([['', '', '', '', '']]);
  setPen_(week, null);
  repaintWeekColors_(week);          // ポジションの色を白に戻す（薄いグレーは残す）
  ui.alert('元に戻しました。時間帯バーの水色は、また自動で付くようになります。');
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
function paintCells_(week, targets, color) {
  var painted = 0;
  ADJUST_BLOCKS.forEach(function(block) {
    var hits = [];
    for (var p = 0; p < ADJUST_ROWS; p++) {
      for (var c = 0; c < ADJUST_GRID_COLS; c++) {
        if (targets[(block.head + 1 + p) + ',' + (block.col + c)]) hits.push([p, c]);
      }
    }
    if (!hits.length) return;
    var range = week.getRange(block.head + 1, block.col, ADJUST_ROWS, ADJUST_GRID_COLS);
    var values = range.getValues();
    var colors = range.getBackgrounds();
    var changed = 0;
    hits.forEach(function(h) {
      if (!(Number(values[h[0]][h[1]]) > 0)) return;      // 人がいないマスは触らない
      colors[h[0]][h[1]] = color;
      changed++;
    });
    if (changed) { range.setBackgrounds(colors); painted += changed; }
  });
  return painted;
}

// ----- ポジションのペン -----
// ボタンを押すと「ペン」になり、そのあとドラッグして選んだところがその色で塗られる。
// 同じボタンをもう一度押すと置ける。
// ペンの中身は控えシート（非表示）の G1＝シート名・G2＝色・G3＝もとのタブの色 に持つ。
// 週シートのセルは結合されていることがあるので触らない。
// 持っているかどうかは「シートのタブの色」で分かるようにする。
var PEN_COL_IN_ADJUST = 7;      // 控えシートの G列

// 今のペン（持っていなければ null）。色と名前をまとめて返す
function readPenInfo_(week) {
  var adj = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ADJUST_SHEET_NAME);
  if (!adj) return null;
  var v = adj.getRange(1, PEN_COL_IN_ADJUST, 4, 1).getValues();
  if (String(v[0][0]) !== week.getName()) return null;
  var color = String(v[1][0]).trim();
  if (!/^#[0-9A-Fa-f]{6}$/.test(color)) return null;
  return { color: color, label: String(v[3][0]) || 'ポジション' };
}

// 今のペンの色（持っていなければ null）
function readPen_(week) {
  var info = readPenInfo_(week);
  return info ? info.color : null;
}

// 画面すみに「いまのペン」をずっと出しておく（消えないお知らせ）
function showPenBanner_(label, extra) {
  SpreadsheetApp.getActiveSpreadsheet().toast(
    (extra ? extra + '\n' : '') + '同じボタンをもう一度押すと、ペンを置きます。',
    '🖊 いまのペン：' + label, -1);
}

// ペンを持つ／置く（color を渡さなければ置く）
function setPen_(week, color, label) {
  var adj = getAdjustSheet_();
  if (!color) {
    if (readPen_(week)) {
      var old = String(adj.getRange(3, PEN_COL_IN_ADJUST).getValue()).trim();
      try { week.setTabColor(/^#[0-9A-Fa-f]{6}$/.test(old) ? old : null); } catch (e) {}
    }
    adj.getRange(1, PEN_COL_IN_ADJUST, 4, 1).setValues([[''], [''], [''], ['']]);
    return;
  }
  if (!readPen_(week)) {                       // 初めて持つとき、今のタブの色を控えておく
    var tab = '';
    try { tab = String(week.getTabColor() || ''); } catch (e) { tab = ''; }
    adj.getRange(3, PEN_COL_IN_ADJUST).setValue(tab);
  }
  adj.getRange(1, PEN_COL_IN_ADJUST, 4, 1)
     .setValues([[week.getName()], [color], [adj.getRange(3, PEN_COL_IN_ADJUST).getValue()], [label || '']]);
  try { week.setTabColor(color); } catch (e) {}   // タブの色＝いま持っているペン
}

// 選んだ範囲のうち、塗る場所を決める
// ・時間帯バーの升目 … そのマス
// ・名前や出勤の列   … その人の1日ぶん
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

// その範囲が、升目や名前の列に少しでもかかっているか（計算だけ・シートは読まない）
function rangeTouchesGrid_(r0, c0, nr, nc) {
  var r1 = r0 + nr - 1, c1 = c0 + nc - 1;
  for (var i = 0; i < ADJUST_BLOCKS.length; i++) {
    var b = ADJUST_BLOCKS[i];
    if (r1 <= b.head || r0 > b.head + ADJUST_ROWS) continue;
    if (c1 < b.col - 6 || c0 >= b.col + ADJUST_GRID_COLS) continue;
    return true;
  }
  return false;
}

// 今えらんでいる範囲（飛び地に対応）
function activeRanges_(week) {
  var list = week.getActiveRangeList();
  if (list) return list.getRanges();
  var one = week.getActiveRange();
  return one ? [one] : [];
}

// ポジションのボタンを押したとき
function paintSelectedPosition_(color, label) {
  var week = activeWeekSheet_();
  if (!week) return;
  var book = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  if (!positionUnlockRow_(week)) {
    if (ui.alert('「' + week.getName() + '」は、まだポジションの色分けを始めていません。\n\n' +
        '今すぐ始めますか。（あとで「色分けをやめる」で元に戻せます）',
        ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
    if (!startPositionColors_(week)) {
      ui.alert('水色のルールが見つかりませんでした。何も変えていません。');
      return;
    }
  }

  // 同じボタンをもう一度＝ペンを置く
  var now = readPen_(week);
  if (now && isSameColor_(now, color)) {
    setPen_(week, null);
    book.toast('ペンを置きました。タブの色も戻します。これで、選んでも色は変わりません。', 'ポジション', 5);
    return;
  }

  setPen_(week, color, label);
  var painted = paintCells_(week, targetsOfRanges_(activeRanges_(week)), color);
  showPenBanner_(label, painted
    ? '選んでいた ' + painted + 'マスを塗りました。このままドラッグすれば続けて塗れます。'
    : '塗りたい時間帯や、名前をドラッグしてください。');
}

// 選んだ瞬間に、ペンの色で塗る（承認のいらない簡易トリガー）
function onSelectionChange(e) {
  try {
    if (!e || !e.range) return;
    var sheet = e.range.getSheet();
    if (!isWeekSheetName_(sheet.getName())) return;
    var r0 = e.range.getRow(), c0 = e.range.getColumn();
    var nr = e.range.getNumRows(), nc = e.range.getNumColumns();
    if (nr * nc > 4000) return;                       // 行や列ごと選んだときは何もしない
    if (!rangeTouchesGrid_(r0, c0, nr, nc)) return;   // 升目にかかっていない
    var pen = readPenInfo_(sheet);
    if (!pen) return;                                 // ペンを持っていない
    var painted = paintCells_(sheet, targetsOfRanges_([e.range]), pen.color);
    if (painted) {
      showPenBanner_(pen.label,
        painted + 'マス（' + (Math.round(painted / 4 * 100) / 100) + '時間）塗りました。');
    }
  } catch (err) {
    Logger.log('onSelectionChange: ' + err.message);
  }
}

function paintPosition1() { var p = positionList_()[0]; paintSelectedPosition_(p.color, p.name); }
function paintPosition2() { var p = positionList_()[1]; paintSelectedPosition_(p.color, p.name); }
function paintPosition3() { var p = positionList_()[2]; paintSelectedPosition_(p.color, p.name); }
function paintPosition4() { var p = positionList_()[3]; paintSelectedPosition_(p.color, p.name); }
function clearPositionPaint() { paintSelectedPosition_(POSITION_BASE_COLOR, 'もとの水色'); }

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
  settings.getRange(POSITION_NAME_ROW, 5, POSITION_COUNT, 2).setValues(rows);
  return '保存しました。メニューの表示は、開き直すと変わります。';
}



// ----- 色の塗り直し（ここ1か所で色を決める） -----
// マスの色は、いつでも次のルールで決まる。
//   今その人がいる … ポジションの色があればそのまま／無ければ 水色（色分け中）・白（通常時は自動の水色に任せる）
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
    unlocked: !!map[week.getName() + '/cf'],
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
        if (isSameColor_(color, ADJUST_COLOR)) {
          colors[p][c] = ctx.unlocked ? POSITION_BASE_COLOR : '#ffffff'; changed = true;
        } else if (ctx.unlocked && isSameColor_(color, '#ffffff')) {
          colors[p][c] = POSITION_BASE_COLOR; changed = true;             // 増えた時間は薄い水色
        } else if (ctx.unlocked && isSameColor_(color, POSITION_RULE_COLOR)) {
          colors[p][c] = POSITION_BASE_COLOR; changed = true;             // 前に塗った濃い水色も薄くする
        } else if (!ctx.unlocked && isOurColor_(color, ctx.ours)) {
          colors[p][c] = '#ffffff'; changed = true;                        // 通常時は自動の水色に任せる
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
// 色分け中にコピーすると「水色の自動ルールが外れたまま・先週の色が残ったまま」になるので、
// 新しいシートだけを、色分けしていない元の状態に戻す（今のシートの色はそのまま残す）。
function resetPositionColorsOnCopy_(source, newSheet) {
  var sheet = getAdjustSheet_();
  var row = adjustRowMap_(sheet)[source.getName() + '/cf'];
  var saved = row ? String(sheet.getRange(row, 4).getValue()) : '';
  if (saved) {
    var removed = [];
    try { removed = JSON.parse(saved); } catch (e) { removed = []; }
    var rules = newSheet.getConditionalFormatRules();
    removed.forEach(function(item) {
      var builder = SpreadsheetApp.newConditionalFormatRule()
        .whenNumberEqualTo(item.value)
        .setBackground(item.bg)
        .setRanges(item.ranges.map(function(a1) { return newSheet.getRange(a1); }));
      if (item.fc) builder.setFontColor(item.fc);
      rules.push(builder.build());
    });
    newSheet.setConditionalFormatRules(rules);
  }
  clearOurGridColors_(newSheet);
  try { newSheet.setTabColor(null); } catch (e) {}     // ペンの色も持ち越さない
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
  setPen_(week, null);                 // 掃除したらペンは置く
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
var POSITION_BUTTON_PER_ROW = 4;              // 4個・3個の2段にする
var POSITION_BUTTON_ROW_GAP = 2;              // 2段目は2行下
var POSITION_BUTTON_FUNCS = ['paintPosition1', 'paintPosition2', 'paintPosition3', 'paintPosition4',
                             'clearPositionPaint', 'refreshWeekColors', 'confirmAdjust'];
var POSITION_BUTTON_IMAGES = [
  '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'
];

// すでに置いてあるボタンを外す（同じ命令が付いている絵だけ）
function removePositionButtons_(sheet) {
  var images;
  try { images = sheet.getImages(); } catch (e) { return 0; }
  var n = 0;
  images.forEach(function(img) {
    var fn = '';
    try { fn = String(img.getScript() || ''); } catch (e) { fn = ''; }
    if (POSITION_BUTTON_FUNCS.indexOf(fn) >= 0) { img.remove(); n++; }
  });
  return n;
}

// ボタンの絵の色は決め打ちなので、設定シートの色もそれに合わせる
// （設定とボタンで色が食い違うと、押した色と見た目が変わってしまうため）
function syncPositionColors_() {
  var settings = getRequestSettingsSheet_();
  var range = settings.getRange(POSITION_NAME_ROW, 6, POSITION_COUNT, 1);
  var now = range.getValues();
  var changed = false;
  var next = POSITION_DEFAULTS.map(function(pos, i) {
    if (String(now[i][0]).trim().toUpperCase() !== pos.color.toUpperCase()) changed = true;
    return [pos.color];
  });
  if (changed) range.setValues(next);
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

// メニュー「ボタンを画面に置く」
function placePositionButtons() {
  var week = activeWeekSheet_();
  if (!week) return;
  syncPositionColors_();
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
  var placed = countPositionButtons_(week);
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'ボタンを' + placed + '個置きました。' +
    (placed < POSITION_BUTTON_FUNCS.length ? '足りないときはもう一度押してください。'
                                           : '画面に出ないときは、ページを再読み込みしてください。'),
    'ポジション', 8);
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
var RECORD_SPREADSHEET_ID = '1tgFvH1k3WxsMb_fG2C_m6vMu_n9KTKGgdScSk7jNxGM';   // シフト希望_アクセス記録（飯塚店）
var RECORD_SUBMIT_SHEET = '送信の記録';
var RECORD_NAME_COL = 3;    // 名前の列
var RECORD_DATA_COL = 16;   // 送った中身（この期間のものかを見る）
var LINE_PUSH_URL = 'https://api.line.me/v2/bot/message/push';
var LINE_REGISTER_CELL = 'E12';     // 送り先の登録を受け付ける期限（他の人に書き換えられないように、10分間だけ受け付ける）
var LINE_REGISTER_MINUTES = 10;
// LINE Developers の「Webhook URL」に貼るアドレス（スマホのフォームと同じアドレスで受け取る）
var LINE_WEBHOOK_URL = REQUEST_FORM_URL;

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
  var period = currentRequestPeriod_(settings, now);
  if (!period) return { created: [], missingBudget: [], notice: '' };

  var created = ensurePeriodSheets_(ss, period);
  var missing = missingBudgetDays_(ss, period);
  var notice = sendDeadlineNotice_(ss, settings, period, today);

  // 予算が足りない日は、週シートを作った日とお知らせを送った日にだけメールする（毎日は送らない）
  if (missing.length && (created.length || notice)) noticeMissingBudget_(ss, period, missing, settings, today);
  return { created: created, missingBudget: missing, notice: notice };
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
    '（入っていない日は、曜日別の基準額がそのまま入ります）\n\n' +
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
  var what = 'シフト希望の催促（スタッフ入力締切の' + left + '日前）';
  var line = lineConfig_(settings);
  if (!line.token || !line.target) {
    noticeSendFailure_(settings, what, line.token ? '送り先のグループが登録されていません。' : 'チャネルアクセストークンが入っていません。');
    return '';
  }

  var text = fillNoticeText_(left === notice.days[0] ? notice.first : notice.later, ss, period, left);
  try {
    linePush_(line, text);
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
  var settings = getRequestSettingsSheet_();
  var pending = {};
  var last = settings.getLastRow();
  if (last >= REQUEST_PENDING_FIRST_ROW) {
    settings.getRange(REQUEST_PENDING_FIRST_ROW, 1, last - REQUEST_PENDING_FIRST_ROW + 1, 2).getValues().forEach(function(r) {
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
    var sheet = SpreadsheetApp.openById(RECORD_SPREADSHEET_ID).getSheetByName(RECORD_SUBMIT_SHEET);
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

function lineConfig_(settings) {
  return {
    token: String(settings.getRange(LINE_TOKEN_CELL).getDisplayValue()).trim(),
    target: String(settings.getRange(LINE_TARGET_CELL).getDisplayValue()).trim()
  };
}

// LINEへ送る。うまくいかない時は理由をそのまま出す（トークンの間違い・無料通数切れなどが分かるように）
function linePush_(line, text) {
  var res = UrlFetchApp.fetch(LINE_PUSH_URL, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + line.token },
    payload: JSON.stringify({ to: line.target, messages: [{ type: 'text', text: text }] }),
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
  var box = 'width:100%;padding:6px;font-size:12px;';
  var html = '<div style="font-family:sans-serif;padding:4px 8px;font-size:14px;line-height:1.6;">' +
    '<p style="margin:0 0 8px;font-size:12px;color:#555;">LINE公式アカウントから、スタッフのグループへお知らせを送るための設定です。</p>' +

    '<label style="display:block;font-weight:bold;">① LINE Developers に貼り付けるアドレス（Webhook URL）</label>' +
    '<input readonly value="' + LINE_WEBHOOK_URL + '" onclick="this.select()" style="' + box + '">' +

    '<label style="display:block;font-weight:bold;margin-top:12px;">② チャネルアクセストークン</label>' +
    '<textarea id="token" style="' + box + 'height:70px;">' + escapeHtml_(line.token) + '</textarea>' +

    '<label style="display:block;font-weight:bold;margin-top:12px;">③ お知らせを送る先</label>' +
    '<p style="margin:2px 0;font-size:12px;color:#555;">下のボタンを押してから' + LINE_REGISTER_MINUTES + '分の間に、グループで何か一言送ると登録されます。</p>' +
    '<button onclick="doOpen()" style="padding:5px 10px;">送り先の登録を受け付ける（' + LINE_REGISTER_MINUTES + '分）</button>' +
    '<p style="margin:2px 0;font-size:12px;color:' + (line.target ? '#186A3B' : '#B03A2E') + ';">' +
    (line.target ? '登録済み（' + escapeHtml_(line.target.slice(0, 8)) + '…）' :
      'まだ登録されていません。公式アカウントをグループに招待し、そのグループで何か一言送ってください。') + '</p>' +

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
    'function doTest(){' +
    '  show("送っています…");' +
    '  google.script.run' +
    '    .withSuccessHandler(function(t){ show(t); })' +
    '    .withFailureHandler(function(e){ show(e.message, true); })' +
    '    .sendLineTest();' +
    '}' +
    '</script>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(460).setHeight(420), 'LINE通知の設定');
}

function saveLineSettings(token) {
  getRequestSettingsSheet_().getRange(LINE_TOKEN_CELL).setValue(String(token == null ? '' : token).trim());
  return '保存しました。';
}

// 「送り先の登録を受け付ける」：これから10分の間だけ、LINEから来た連絡の送り先を登録する
function openLineRegister() {
  var until = new Date(Date.now() + LINE_REGISTER_MINUTES * 60000);
  getRequestSettingsSheet_().getRange(LINE_REGISTER_CELL).setValue(Utilities.formatDate(until, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm'));
  return LINE_REGISTER_MINUTES + '分の間、登録を受け付けます。グループで何か一言送ってください。';
}

function sendLineTest() {
  var settings = getRequestSettingsSheet_();
  var line = lineConfig_(settings);
  if (!line.token) throw new Error('先にトークンを入れて「保存する」を押してください。');
  if (!line.target) throw new Error('送り先がまだ登録されていません。公式アカウントをグループに招待し、そのグループで何か一言送ってください。');
  linePush_(line, 'テスト送信です。この形でシフト希望のお知らせが届きます。');
  return '送りました。LINEを見てください。';
}

// メニュー「今すぐ点検する」：毎日の自動点検を、その場で1回動かす（お知らせの日でなければ送りません）
function runShiftCheckNow() {
  var result = dailyShiftCheck();
  var monthly = monthlyNoticeCheck();
  SpreadsheetApp.getUi().alert(
    '作った週シート：' + (result.created.length ? result.created.join('、') : 'なし') + '\n' +
    '日別の予算が無い日：' + (result.missingBudget.length ? result.missingBudget.length + '日' : 'なし') + '\n' +
    'LINEに送ったお知らせ：' + (result.notice ? '\n' + result.notice : 'なし（今日は送る日ではありません）') + '\n' +
    '毎月のおしらせ：' + (monthly ? '\n' + monthly : 'なし（今日は送る日ではありません）'));
}


// ===== お知らせの文面・日程の設定（メニュー「お知らせの文面・日程」） =====
// 送る日・送る時刻・文面を、店長が画面から変えられるようにする。
// 文面の {期間} {締切} {リンク} {残り日数} {未提出} {途中} {入力状況} は、送るときに中身へ置き換わる。
var NOTICE_DAYS_CELL = 'E13';    // 送る日（例「14,7,5,3」＝締切の何日前）
var NOTICE_HOUR_CELL = 'E14';    // 送る時刻（0〜23時）
var NOTICE_FIRST_CELL = 'E15';   // 1回目（いちばん早い日）に送る文
var NOTICE_LATER_CELL = 'E16';   // 2回目以降に送る文
var NOTICE_DEFAULT_FIRST = 'シフト希望の入力をお願いします。\n期間 {期間}\nスタッフ入力締切 {締切}\n{リンク}';
var NOTICE_DEFAULT_LATER = '【あと{残り日数}日】シフト希望のスタッフ入力締切が近づいています。\n期間 {期間}\nスタッフ入力締切 {締切}\n{入力状況}\n{リンク}';

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
  return {
    days: days.length ? days : NOTICE_DAYS.slice(),
    hour: hour >= 0 && hour <= 23 ? Math.floor(hour) : NOTICE_HOUR,
    first: first.trim() ? first : NOTICE_DEFAULT_FIRST,
    later: later.trim() ? later : NOTICE_DEFAULT_LATER
  };
}

// 文の中の {…} を、その時の中身に置き換える
function fillNoticeText_(template, ss, period, left) {
  var text = String(template);
  var last = addDaysYmd_(period.start, period.weeks * 7 - 1);
  var map = {
    '{期間}': ymdLabel_(period.start) + '〜' + ymdLabel_(last),
    '{締切}': ymdLabel_(period.deadline.slice(0, 10)) + ' ' + period.deadline.slice(11),
    '{リンク}': REQUEST_FORM_URL + '?p=' + period.start.slice(2).replace(/-/g, ''),
    '{残り日数}': String(left),
    '{店舗名}': getStoreName_()
  };
  if (/\{未提出\}|\{途中\}|\{入力状況\}/.test(text)) {
    var waiting = waitingCounts_(ss, period);
    var lines = [];
    if (waiting.notSubmitted) lines.push('まだ出していない人：' + waiting.notSubmitted + '人');
    if (waiting.partial) lines.push('途中の人（1週でも空きがある）：' + waiting.partial + '人');
    map['{未提出}'] = String(waiting.notSubmitted);
    map['{途中}'] = String(waiting.partial);
    map['{入力状況}'] = lines.length ? lines.join('\n') : 'みなさん提出ありがとうございます。変える方はお早めに。';
  }
  Object.keys(map).forEach(function(key) { text = text.split(key).join(map[key]); });
  return text;
}

// メニュー「お知らせの文面・日程」
function openNoticeSettings() {
  var settings = getRequestSettingsSheet_();
  var notice = readNoticeSettings_(settings);
  var box = 'width:100%;padding:6px;font-size:12px;';
  var html = '<div style="font-family:sans-serif;padding:4px 8px;font-size:14px;line-height:1.6;">' +
    '<label style="display:block;font-weight:bold;">送る日（スタッフ入力締切の何日前か。カンマ区切り）</label>' +
    '<input id="days" value="' + notice.days.join(',') + '" style="' + box + '">' +
    '<label style="display:block;font-weight:bold;margin-top:10px;">送る時刻</label>' +
    '<select id="hour" style="' + box + '">' + numberOptions_(0, 23, notice.hour, '#時ごろ') + '</select>' +
    '<label style="display:block;font-weight:bold;margin-top:10px;">1回目（いちばん早い日）の文</label>' +
    '<textarea id="first" style="' + box + 'height:90px;">' + escapeHtml_(notice.first) + '</textarea>' +
    '<label style="display:block;font-weight:bold;margin-top:10px;">2回目以降の文</label>' +
    '<textarea id="later" style="' + box + 'height:110px;">' + escapeHtml_(notice.later) + '</textarea>' +
    '<p style="margin:8px 0 0;font-size:12px;color:#555;">使える差し込み：<br>' +
    '{期間}／{締切}／{リンク}／{残り日数}／{未提出}／{途中}／{入力状況}／{店舗名}<br>' +
    '{入力状況}＝「まだ出していない人：◯人」などを自動で作ります（全員そろっていればお礼の文になります）。</p>' +
    '<p id="msg" style="margin:8px 0 0;font-size:12px;color:#186A3B;"></p>' +
    '<div style="margin-top:10px;text-align:right;">' +
    '<button onclick="doPreview()" style="margin-right:8px;padding:6px 12px;">今の内容で見てみる</button>' +
    '<button id="saveBtn" onclick="doSave()" style="padding:6px 12px;">保存する</button>' +
    '</div>' +
    '<pre id="preview" style="white-space:pre-wrap;background:#F4F4F2;border-radius:6px;padding:8px;font-size:12px;margin-top:8px;"></pre>' +
    '</div>' +
    '<script>' +
    'function show(t, err){ var m = document.getElementById("msg"); m.textContent = t; m.style.color = err ? "#B03A2E" : "#186A3B"; }' +
    'function values(){ return { days: document.getElementById("days").value, hour: document.getElementById("hour").value,' +
    '  first: document.getElementById("first").value, later: document.getElementById("later").value }; }' +
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
  ['first', 'later'].forEach(function(key) {
    if (String(values[key] || '').length > 1000) throw new Error('文が長すぎます（1000文字まで）。');
  });

  settings.getRange(NOTICE_DAYS_CELL).setValue(days.sort(function(a, b) { return b - a; }).join(','));
  settings.getRange(NOTICE_HOUR_CELL).setValue(String(hour));
  settings.getRange(NOTICE_FIRST_CELL).setValue(String(values.first || ''));
  settings.getRange(NOTICE_LATER_CELL).setValue(String(values.later || ''));

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
  var first = String(values.first || NOTICE_DEFAULT_FIRST);
  var later = String(values.later || NOTICE_DEFAULT_LATER);
  var out = [];
  days.forEach(function(left, i) {
    out.push('― ' + left + '日前（' + ymdLabel_(addDaysYmd_(period.deadline.slice(0, 10), -left)) + '） ―');
    out.push(fillNoticeText_(i === 0 ? first : later, ss, period, left));
    out.push('');
  });
  return out.join('\n');
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
// 1回目だけ「締切り2週間前です」、2回目からは「締切りまで1週間です」になる
var MONTHLY_DEFAULT_TEXTS = [MONTHLY_DEFAULT_FIRST, MONTHLY_DEFAULT_TEXT];
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
  // 送る日の数にそろえる（足りないぶんは最後の文を使いまわす）
  texts = days.map(function(d, i) { return texts[Math.min(i, texts.length - 1)]; });

  var hour = Number(String(settings.getRange(MONTHLY_HOUR_CELL).getDisplayValue()).trim());
  return {
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
    var what = '毎月のおしらせ（' + ymdLabel_(due) + 'の' + left + '日前）';
    var line = lineConfig_(settings);
    if (!line.token || !line.target) {
      noticeSendFailure_(settings, what, line.token ? '送り先のグループが登録されていません。' : 'チャネルアクセストークンが入っていません。');
      return '';
    }
    var text = fillMonthlyText_(monthly.texts[index], due, left);
    try {
      linePush_(line, text);
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
  var data = JSON.stringify({ days: monthly.days, texts: monthly.texts, def: MONTHLY_DEFAULT_TEXT }).replace(/</g, '\\u003c');
  var html = '<div style="font-family:sans-serif;padding:4px 8px;font-size:14px;line-height:1.6;">' +
    '<p style="margin:0 0 8px;font-size:12px;color:#555;">毎月の決まった日を期日にして、その前にLINEのグループへお知らせを送ります。</p>' +

    '<label style="display:block;font-weight:bold;">送る／送らない</label>' +
    '<select id="on" style="' + box + '">' +
    '<option value="入"' + (monthly.on ? ' selected' : '') + '>送る</option>' +
    '<option value="切"' + (monthly.on ? '' : ' selected') + '>送らない</option></select>' +

    '<label style="display:block;font-weight:bold;margin-top:10px;">期日（毎月この日）</label>' +
    '<select id="day" style="' + box + '">' + numberOptions_(1, 31, monthly.day, '毎月#日') + '</select>' +

    '<label style="display:block;font-weight:bold;margin-top:10px;">送る日（期日の何日前か。カンマ区切り）</label>' +
    '<input id="days" value="' + monthly.days.join(',') + '" oninput="build()" style="' + box + '">' +

    '<label style="display:block;font-weight:bold;margin-top:10px;">送る時刻</label>' +
    '<select id="hour" style="' + box + '">' + numberOptions_(0, 23, monthly.hour, '#時ごろ') + '</select>' +

    '<div id="texts"></div>' +

    '<p style="margin:8px 0 0;font-size:12px;color:#555;">使える差し込み：<br>' +
    '{月}／{日}／{曜日}（木曜日）／{期日}（9/25(木)）／{残り}（2週間・1週間・3日）／{残り日数}（数字だけ）／{店舗名}</p>' +
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
    '    lab.textContent = d + "日前に送る文";' +
    '    var ta = document.createElement("textarea");' +
    '    ta.setAttribute("data-day", d); ta.style.cssText = TA_STYLE;' +
    '    ta.value = byDay[d] !== undefined ? byDay[d] : (olds[i] !== undefined ? olds[i] : (olds.length ? olds[olds.length - 1] : DATA.def));' +
    '    wrap.appendChild(lab); wrap.appendChild(ta);' +
    '  });' +
    '}' +
    'function values(){' +
    '  var texts = Array.prototype.map.call(document.getElementById("texts").getElementsByTagName("textarea"), function(t){ return t.value; });' +
    '  return { on: document.getElementById("on").value, day: document.getElementById("day").value,' +
    '    days: document.getElementById("days").value, hour: document.getElementById("hour").value, texts: texts };' +
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
  texts = days.map(function(d, i) { return texts[Math.min(i, texts.length - 1)]; });
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
    out.push('― ' + left + '日前（' + ymdLabel_(addDaysYmd_(due, -left)) + '） ―');
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
function refreshWeekHistory_(sheet) {
  if (!isWeekSheetName_(sheet.getName())) return 0;
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
  SpreadsheetApp.getUi().alert('出勤履歴を作り直しました。\n週シート ' + weekSheets.length + '枚から ' + rows.length + '行' +
    (kept.length ? '（書庫へ移した週の ' + kept.length + '行はそのまま残しています）' : ''));
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
  var tz = ss.getSpreadsheetTimeZone();
  var months = [];
  var todayYm = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM').split('-').map(Number);
  for (var back = 12; back >= -1; back--) {
    var d = new Date(Date.UTC(todayYm[0], todayYm[1] - 1 - back, 1));
    months.push(d.getUTCFullYear() + '-' + ('0' + (d.getUTCMonth() + 1)).slice(-2));
  }
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

function newestWeekSheet_(ss) {
  var names = ss.getSheets().map(function(s) { return s.getName(); })
    .filter(function(n) { return /^\d{6}$/.test(n); }).sort();
  return names.length ? ss.getSheetByName(names[names.length - 1]) : null;
}

// 週シートの右側（転写した所）を手で直した時も、出勤履歴に自動で反映する
// ※承認のいらない簡易トリガー。左側の入力表(A〜S列)を直しただけの時は何もしない
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    var sheet = e.range.getSheet();
    if (!isWeekSheetName_(sheet.getName())) return;
    if (e.range.getColumn() + e.range.getNumColumns() - 1 < 20) return;
    refreshWeekHistory_(sheet);
    // 出勤・退勤・休憩を直したら、その人の色をその場で塗り直す
    SpreadsheetApp.flush();
    repaintEditedRows_(sheet, e.range);
  } catch (err) {
    Logger.log('onEdit: ' + err.message);
  }
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
