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
    // 0. 複製する前に「今のシート」の余分な空行・空列を縮めておく
    //    （ワークブック全体のセル数が上限(1000万)に近づくと複製自体が失敗するため、
    //      次週シート作成のたびに毎回自動でこれを行い、二度と同じ問題が起きないようにする）
    trimSheetSize_(currentSheet);

    // 1. T1セルから日付を取得
    const currentDate = new Date(currentSheet.getRange("T1").getValue());

    // 2. 7日後を計算
    const nextDate = new Date(currentDate.getTime() + 7 * 24 * 60 * 60 * 1000);

    // 3. YYMMDD形式の名前を作成 (例: 260629)
    const nextSheetName = Utilities.formatDate(nextDate, "JST", "yyMMdd");

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

    // 8. 祝日の色付けを新しいシートにも反映する（「祝日」シートは既にあるものを使う）
    applyHolidayColoring_(newSheet);

    // 9. シフト時間入力表(B4:O25)の表示形式を、常に「9:00」のようにコロン付きになるよう整える
    fixTimeFormat_(newSheet);

    // 10. 一番左に移動してアクティブにする
    ss.setActiveSheet(newSheet);
    ss.moveActiveSheet(1);

    // 11. 終わった週の出勤を「出勤履歴」に取り込み、設定・集計のシートの余白も詰めておく
    refreshWeekHistory_(currentSheet);
    [REQUEST_SETTINGS_SHEET, ROLE_SHEET_NAME, HISTORY_SHEET_NAME, MONTHLY_SHEET_NAME].forEach(function(name) {
      var s = ss.getSheetByName(name);
      if (s) trimSheetSize_(s);
    });

    // 12. 直近の週だけを残し、古い週は書庫ファイルへ移す（ファイルが重くならないようにする）
    archiveOldWeeks_(true);

  } catch (e) {
    // ★エラーの本当の原因を突き止めるための一時的な処理。
    //   画面に詳しい内容をそのまま表示し、あとで確認できるようにも保存しておく。
    var detail = '次週シート作成でエラー: ' + e.message + '\n' + (e.stack || '');
    PropertiesService.getDocumentProperties().setProperty('lastError', detail);
    Logger.log(detail);
    SpreadsheetApp.getUi().alert(detail);
  }
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
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('スタッフ管理')
    .addItem('スタッフを追加', 'addStaffMember')
    .addItem('スタッフを削除', 'removeStaffMember')
    .addItem('時間を固定するスタッフを設定', 'setProtectedStaff')
    .addItem('店長・社員を設定', 'setStaffRoles')
    .addSeparator()
    .addItem('シフト希望の募集設定', 'openShiftRequestSettings')
    .addSeparator()
    .addItem('月次集計を開く', 'openMonthlySummary')
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
var REQUEST_BASE_DAYS_CELL = 'E4';      // 締切＝期間の最終日の何日前
var REQUEST_BASE_TIME_CELL = 'E5';      // 締切の時刻
var REQUEST_STORE_NAME_CELL = 'E6';     // 店舗名（見出しなどに使う）
var REQUEST_CAST_CLOSING_CELL = 'E7';   // キャストの締め日（月次集計に使う）
var REQUEST_DEFAULT_BASE = { closingDay: 15, deadlineDays: 14, deadlineTime: '23:59', castClosingDay: 20 };

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
  sheet.getRange('E3:E7').setNumberFormat('@');
  [[REQUEST_BASE_CLOSING_CELL, String(REQUEST_DEFAULT_BASE.closingDay)],
   [REQUEST_BASE_DAYS_CELL, String(REQUEST_DEFAULT_BASE.deadlineDays)],
   [REQUEST_BASE_TIME_CELL, REQUEST_DEFAULT_BASE.deadlineTime],
   [REQUEST_CAST_CLOSING_CELL, String(REQUEST_DEFAULT_BASE.castClosingDay)]].forEach(function(pair) {
    if (String(sheet.getRange(pair[0]).getDisplayValue()).trim() === '') sheet.getRange(pair[0]).setValue(pair[1]);
  });
  sheet.getRange('D3:D8').setValues([
    ['基本設定：社員の締め日（毎月この日）'],
    ['基本設定：スタッフ入力締切は期間の最終日の何日前'],
    ['基本設定：スタッフ入力締切の時刻'],
    ['基本設定：店舗名'],
    ['基本設定：キャストの締め日（毎月この日）'],
    ['基本設定：書庫ファイルのURL（古い週の移動先）']
  ]);
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
  var daysText = String(sheet.getRange(REQUEST_BASE_DAYS_CELL).getDisplayValue()).trim();
  var time = String(sheet.getRange(REQUEST_BASE_TIME_CELL).getDisplayValue()).trim();
  var cast = Number(sheet.getRange(REQUEST_CAST_CLOSING_CELL).getDisplayValue());
  return {
    closingDay: closing >= 1 && closing <= 31 ? Math.floor(closing) : REQUEST_DEFAULT_BASE.closingDay,
    deadlineDays: daysText !== '' && Number(daysText) >= 0 && Number(daysText) <= 28 ? Math.floor(Number(daysText)) : REQUEST_DEFAULT_BASE.deadlineDays,
    deadlineTime: /^\d{2}:\d{2}$/.test(time) ? time : REQUEST_DEFAULT_BASE.deadlineTime,
    castClosingDay: cast >= 1 && cast <= 31 ? Math.floor(cast) : REQUEST_DEFAULT_BASE.castClosingDay
  };
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
// 締切＝期間の最終日の◯日前。締切を過ぎた期間と、FIRST_ROLE_SHEET より前から始まる期間は飛ばす
function autoRequestPeriod_(base, now) {
  var ym = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM').split('-').map(Number);
  for (var i = -1; i <= 13; i++) {
    var end = requestPeriodEnd_(ym[0], ym[1] + i, base.closingDay);
    var start = addDaysYmd_(requestPeriodEnd_(ym[0], ym[1] + i - 1, base.closingDay), 1);
    if (start.slice(2).replace(/-/g, '') < FIRST_ROLE_SHEET) continue;
    var deadline = addDaysYmd_(end, -base.deadlineDays) + ' ' + base.deadlineTime;
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
    '<select id="deadlineDays" style="' + select + '">' + numberOptions_(0, 28, base.deadlineDays, '期間の最終日の#日前') + '</select>' +
    '<label style="' + label + '">スタッフ入力締切の時刻</label>' +
    '<select id="baseTime" style="' + select + '">' + requestTimeOptions_(base.deadlineTime) + '</select>' +
    '<label style="' + label + '">キャストの締め日（月次集計に使います）</label>' +
    '<select id="castClosingDay" style="' + select + '">' + numberOptions_(1, 31, base.castClosingDay, '毎月#日') + '</select>' +
    '<label style="' + label + '">店舗名</label>' +
    '<input id="storeName" value="' + escapeHtml_(storeName) + '" placeholder="例：ペッパーランチ飯塚店" style="' + select + '">' +
    '<label style="' + label + '">書庫ファイルのURL（' + ARCHIVE_KEEP_WEEKS + '週より前の週シートの移動先）</label>' +
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
    '      deadlineDays: document.getElementById("deadlineDays").value,' +
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
  var deadlineDays = Number(settings.deadlineDays);
  var baseTime = String(settings.baseTime || '');
  var castClosingDay = Number(settings.castClosingDay);
  var storeName = String(settings.storeName == null ? '' : settings.storeName).trim().slice(0, 40);
  if (!(closingDay >= 1 && closingDay <= 31 && Math.floor(closingDay) === closingDay)) throw new Error('社員の締め日は1〜31日で選んでください。');
  if (!(castClosingDay >= 1 && castClosingDay <= 31 && Math.floor(castClosingDay) === castClosingDay)) throw new Error('キャストの締め日は1〜31日で選んでください。');
  if (!(deadlineDays >= 0 && deadlineDays <= 28 && Math.floor(deadlineDays) === deadlineDays)) throw new Error('スタッフ入力締切は0〜28日前で選んでください。');
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
    .setValues([[String(closingDay)], [String(deadlineDays)], [baseTime], [storeName], [String(castClosingDay)], [archiveUrl]]);

  // 月次集計の見出し・締め期間も新しい設定に合わせる
  try {
    if (SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MONTHLY_SHEET_NAME)) refreshMonthlySummary_();
  } catch (e) {
    Logger.log('月次集計の更新に失敗: ' + e.message);
  }
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

function getArchiveSpreadsheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settings = ss.getSheetByName(REQUEST_SETTINGS_SHEET);
  if (!settings) return null;
  var url = String(settings.getRange(ARCHIVE_URL_CELL).getDisplayValue()).trim();
  var id = /[-\w]{25,}/.exec(url);
  if (!id) return null;
  if (id[0] === ss.getId()) throw new Error('書庫ファイルに、このシフト表と同じURLが設定されています。');
  return SpreadsheetApp.openById(id[0]);
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
  var archive;
  try {
    archive = getArchiveSpreadsheet_();
  } catch (e) {
    return { moved: [], failed: [], message: e.message };
  }
  var all = archiveTargets_(ss);
  if (!all.length) return { moved: [], failed: [], message: '移す週はありません' };
  if (!archive) return { moved: [], failed: [], message: '書庫ファイルが設定されていないため、' + all.length + '枚をそのまま残しました' };

  var limit = silent ? ARCHIVE_MAX_AUTO : ARCHIVE_MAX_MANUAL;
  var targets = all.slice(0, limit);
  var rest = all.length - targets.length;

  // 移す週の出勤を、先にまとめて履歴へ取り込む
  var weekTargets = targets.filter(function(s) { return isWeekSheetName_(s.getName()); });
  if (weekTargets.length) refreshHistoryForSheets_(weekTargets);

  var moved = [], failed = [];
  targets.forEach(function(sheet) {
    var name = sheet.getName();
    try {
      moved.push(archiveSheet_(ss, archive, sheet));
    } catch (e) {
      failed.push(name + '（' + e.message + '）');
    }
  });

  var message = moved.length ? moved.length + '枚を書庫へ移しました：' + moved.join('、') : '移した週はありません';
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
