function copyAttendanceData() {
  const sheet = SpreadsheetApp.getActiveSheet();

  // 1. 転記先エリアのデータをすべて綺麗に消去
  const clearRangesStr = ["T24:X39", "T43:X58", "T62:X77", "DD4:DH19", "DD24:DH39", "DD43:DH58", "DD62:DH77"];
  clearRangesStr.forEach(rangeStr => sheet.getRange(rangeStr).clearContent());

  // 2. データの最終行を取得
  const lastRow = sheet.getLastRow();
  if (lastRow < 4) return;

  // ★超重要修正： getDisplayValues() を使うことで、
  // プログラムが裏の「時差バグ」に巻き込まれず、人間が見ている画面そのままの文字を取得します。
  const data = sheet.getRange(4, 1, lastRow - 3, 15).getDisplayValues();

  // ★「9:00」でも「900」でも、確実に数式が反応する「900」の数字に整える関数
  function normalizeTime(val) {
    if (!val) return "";
    let strVal = val.toString().trim();
    if (strVal === "") return "";

    // コロンが含まれている場合（例："9:30" → 930 に変換）
    if (strVal.includes(":")) {
      let parts = strVal.split(":");
      return parseInt(parts[0], 10) * 100 + parseInt(parts[1], 10);
    }

    // コロンがない場合（例："900" → 900 のまま）
    let num = parseInt(strVal, 10);
    return isNaN(num) ? "" : num;
  }

  let weeklyData = {
    "月": [], "火": [], "水": [], "木": [], "金": [], "土": [], "日": []
  };

  // 3. 全行をスキャンして曜日ごとに仕分け
  for (let i = 0; i < data.length; i++) {
    const name = data[i][0];
    if (!name) continue;

    // 画面上の見た目文字をそのまま normalizeTime に渡す
    if (data[i][1] !== "") weeklyData["月"].push([name, normalizeTime(data[i][1]), normalizeTime(data[i][2])]);
    if (data[i][3] !== "") weeklyData["火"].push([name, normalizeTime(data[i][3]), normalizeTime(data[i][4])]);
    if (data[i][5] !== "") weeklyData["水"].push([name, normalizeTime(data[i][5]), normalizeTime(data[i][6])]);
    if (data[i][7] !== "") weeklyData["木"].push([name, normalizeTime(data[i][7]), normalizeTime(data[i][8])]);
    if (data[i][9] !== "") weeklyData["金"].push([name, normalizeTime(data[i][9]), normalizeTime(data[i][10])]);
    if (data[i][11] !== "") weeklyData["土"].push([name, normalizeTime(data[i][11]), normalizeTime(data[i][12])]);
    if (data[i][13] !== "") weeklyData["日"].push([name, normalizeTime(data[i][13]), normalizeTime(data[i][14])]);
  }

  // 4. 指定されたセルへ貼り付け、見た目を一括で整える関数
  function pasteDayData(dayKey, destRangeStr) {
    const output = weeklyData[dayKey];
    if (output.length > 0) {
      const destStartCell = sheet.getRange(destRangeStr);
      const startRow = destStartCell.getRow();
      const startCol = destStartCell.getColumn();

      // 値を書き込む (名前, 出勤, 退勤)
      sheet.getRange(startRow, startCol, output.length, 3).setValues(output);

      // 出勤・退勤セルの「見た目」だけを自動で "9:00" とコロン付きになるよう整える
      sheet.getRange(startRow, startCol + 1, output.length, 2).setNumberFormat('0":"00');
    }
  }

  // 5. 各曜日の貼り付けを実行
  pasteDayData("月", "T24");
  pasteDayData("火", "T43");
  pasteDayData("水", "T62");
  pasteDayData("木", "DD4");
  pasteDayData("金", "DD24");
  pasteDayData("土", "DD43");
  pasteDayData("日", "DD62");

  // 6. 転写した内容を「出勤履歴」にも取り込む（月次集計がすぐ最新になる）
  try {
    refreshWeekHistory_(sheet);
  } catch (e) {
    Logger.log('出勤履歴の更新に失敗: ' + e.message);
  }

  // 7. 転写した直後の時間を「調整前」として自動で記録する
  //    （希望どおりの状態が調整前。あとはメニューの「削った時間を薄く表示する」を押すだけでよい）
  try {
    if (/^\d{6}$/.test(sheet.getName())) {
      SpreadsheetApp.flush();                 // 升目の計算を確定させてから読む
      clearAdjustMarks_(sheet);               // 前回の薄いグレーが残っていれば先に消す
      clearOurGridColors_(sheet);             // 前のポジションの色も消す（希望からやり直すため）
      saveAdjustSnapshots_(sheet);
      repaintWeekColors_(sheet);              // いる時間を薄い水色に戻す
      SpreadsheetApp.getActiveSpreadsheet().toast(
        '転写しました。今の時間を「調整前」として記録しました。', 'シフト表', 6);
    }
  } catch (e) {
    Logger.log('調整前の記録に失敗: ' + e.message);
  }
}

function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var params = e.parameter;

  // ?listSheets=1 でタブ一覧だけを返す
  if (params.listSheets === '1') {
    var names = ss.getSheets().map(function(s) { return s.getName(); });
    return ContentService.createTextOutput(JSON.stringify({
      book: ss.getName(),
      bookId: ss.getId(),
      sheets: names
    })).setMimeType(ContentService.MimeType.JSON);
  }

  // ?lastError=1 で、直近に createNextWeekSheet 等で発生したエラー内容を返す（調査用）
  if (params.lastError === '1') {
    var lastError = PropertiesService.getDocumentProperties().getProperty('lastError') || '(エラーは記録されていません)';
    return ContentService.createTextOutput(JSON.stringify({
      lastError: lastError
    })).setMimeType(ContentService.MimeType.JSON);
  }

  // ?cellCount=1 で、ワークブック全体のセル数（1000万上限との比較）を調べる（調査用）
  // 実際に使っている範囲（lastRow/lastCol）も一緒に返し、縮めても安全な余白量を調べる
  if (params.cellCount === '1') {
    var sheetsInfo = ss.getSheets().map(function(s) {
      var r = s.getMaxRows();
      var c = s.getMaxColumns();
      return {
        name: s.getName(),
        rows: r, cols: c, cells: r * c,
        lastRow: s.getLastRow(), lastCol: s.getLastColumn()
      };
    });
    var total = sheetsInfo.reduce(function(sum, s) { return sum + s.cells; }, 0);
    return ContentService.createTextOutput(JSON.stringify({
      totalCells: total,
      limit: 10000000,
      sheets: sheetsInfo
    })).setMimeType(ContentService.MimeType.JSON);
  }

  var sheetName = params.sheet || ss.getSheets()[0].getName();
  var mode = params.mode || 'both'; // values / formulas / both
  var rows = parseInt(params.rows, 10) || 50;
  if (rows > 300) rows = 300;
  var startRow = parseInt(params.startRow, 10) || 1;
  var startCol = parseInt(params.startCol, 10) || 1;
  var cols = parseInt(params.cols, 10) || 0; // 0なら最終列まで

  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    return ContentService.createTextOutput(JSON.stringify({
      error: 'sheet not found: ' + sheetName,
      book: ss.getName(),
      bookId: ss.getId()
    })).setMimeType(ContentService.MimeType.JSON);
  }

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var numRows = Math.min(rows, Math.max(lastRow - startRow + 1, 1));
  var numCols = cols > 0 ? Math.min(cols, lastCol - startCol + 1) : (lastCol - startCol + 1);
  if (numCols < 1) numCols = 1;

  var range = sheet.getRange(startRow, startCol, numRows, numCols);

  var result = {
    book: ss.getName(),
    bookId: ss.getId(),
    sheet: sheetName,
    startRow: startRow,
    startCol: startCol,
    rows: numRows,
    columns: numCols
  };

  if (mode === 'values' || mode === 'both') {
    result.values = range.getValues();
  }
  if (mode === 'formulas' || mode === 'both') {
    result.formulas = range.getFormulas();
  }
  if (mode === 'both') {
    // 各セルの背景色も一緒に返す（ただし条件付き書式の色はここに出ない）
    result.backgrounds = range.getBackgrounds();
    // 条件付き書式のルール自体を返す（日曜が赤くなる仕組みを確認するため）
    result.conditionalFormats = serializeConditionalFormats(sheet);
  }

  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

function serializeConditionalFormats(sheet) {
  var rules = sheet.getConditionalFormatRules();
  return rules.map(function(rule) {
    var ranges = rule.getRanges().map(function(r) { return r.getA1Notation(); });
    var obj = { ranges: ranges };
    var bc = rule.getBooleanCondition();
    if (bc) {
      obj.criteriaType = bc.getCriteriaType().toString();
      obj.criteriaValues = bc.getCriteriaValues();
      obj.background = bc.getBackground();
      obj.fontColor = bc.getFontColor();
    } else {
      obj.gradient = true;
    }
    return obj;
  });
}

// 内閣府の祝日CSVを取得して「祝日」シート（非表示）に書き込む
function refreshHolidayList_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('祝日');
  if (!sheet) {
    sheet = ss.insertSheet('祝日');
  } else {
    sheet.clearContents();
  }

  var url = 'https://www8.cao.go.jp/chosei/shukujitsu/syukujitsu.csv';
  var response = UrlFetchApp.fetch(url);
  var text = response.getContentText('Shift_JIS');
  var lines = text.split('\n');
  var rows = [];

  for (var i = 1; i < lines.length; i++) {
    var line = lines[i].trim();
    if (!line) continue;
    var cols = line.split(',');
    var dateStr = cols[0];
    var name = cols[1] || '';
    var parts = dateStr.split('/');
    if (parts.length !== 3) continue;
    var y = parseInt(parts[0], 10), m = parseInt(parts[1], 10), d = parseInt(parts[2], 10);
    if (!y || !m || !d) continue;
    rows.push([new Date(y, m - 1, d), name]);
  }

  sheet.getRange(1, 1, 1, 2).setValues([['日付', '祝日名']]);
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 2).setValues(rows);
    sheet.getRange(2, 1, rows.length, 1).setNumberFormat('yyyy/mm/dd');
  }
  sheet.hideSheet();
}

// 「祝日」シートの日付一覧と照らし合わせて、対象セルの背景色を
// 直接赤/白で塗り分ける（条件付き書式は使わない）
function applyHolidayColoring_(sheet) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var holidaySheet = ss.getSheetByName('祝日');
  if (!holidaySheet) {
    throw new Error('「祝日」シートが見つかりません。先に refreshHolidayList_ を実行してください。');
  }

  var lastRow = holidaySheet.getLastRow();
  var holidayKeys = {};
  if (lastRow >= 2) {
    var dateVals = holidaySheet.getRange(2, 1, lastRow - 1, 1).getValues();
    dateVals.forEach(function(r) {
      if (r[0] instanceof Date) {
        holidayKeys[Utilities.formatDate(r[0], Session.getScriptTimeZone(), 'yyyy-MM-dd')] = true;
      }
    });
  }

  function isHoliday(val) {
    if (!(val instanceof Date)) return false;
    var key = Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    return !!holidayKeys[key];
  }

  // B2:K2（日付セルB,D,F,H,Jと、右隣の曜日セルをペアで塗る）
  [2, 4, 6, 8, 10].forEach(function(col) {
    var dateCell = sheet.getRange(2, col);
    var weekdayCell = sheet.getRange(2, col + 1);
    var color = isHoliday(dateCell.getValue()) ? '#ff0000' : '#ffffff';
    dateCell.setBackground(color);
    weekdayCell.setBackground(color);
  });

  // Y3:BV4（月〜金、10列ごとに結合された日付見出しブロック）
  [25, 35, 45, 55, 65].forEach(function(col) { // Y, AI, AS, BC, BM
    var dateCell = sheet.getRange(3, col);
    var color = isHoliday(dateCell.getValue()) ? '#ff0000' : '#ffffff';
    sheet.getRange(3, col, 2, 10).setBackground(color);
  });

  // 右側の曜日見出し（R23=月, R42=火, R61=水, DB3=木, DB23=金）も祝日なら赤くする
  // ★土曜(R42系の土・DB42)・日曜(DB61)は元から色が付いているため対象外
  [
    { col: 18, labelRow: 23 },  // R23 月曜
    { col: 18, labelRow: 42 },  // R42 火曜
    { col: 18, labelRow: 61 },  // R61 水曜
    { col: 106, labelRow: 3 },  // DB3 木曜
    { col: 106, labelRow: 23 }  // DB23 金曜
  ].forEach(function(h) {
    var dateCell = sheet.getRange(h.labelRow + 2, h.col); // 見出しの2行下に実際の日付が入っている
    var labelCell = sheet.getRange(h.labelRow, h.col);
    var color = isHoliday(dateCell.getValue()) ? '#ff0000' : '#ffffff';
    labelCell.setBackground(color);
  });
}

// ★このシフト表の「今使っている週のシート」を開いた状態で実行してください。
//   新しい週のシートを作った後も、同じ関数をもう一度実行すれば色が更新されます。
function setupHolidayColoring() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet(); // 「祝日」シート作成で切り替わる前に先に確定させる

  refreshHolidayList_();
  applyHolidayColoring_(sheet);

  Logger.log('祝日の色付けを更新しました。（対象シート: ' + sheet.getName() + '）');
}
