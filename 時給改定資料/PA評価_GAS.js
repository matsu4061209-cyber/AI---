// ============================================================
//  ペッパーランチ PA評価・時給変更管理 GASコード
// ============================================================

// ===== スプレッドシートID =====
const EVAL_SS_ID    = '1jv1lKCDfz_GL14cSSkH9OJEchgZxYOjzbrK5It3VHl0';   // 評価スプレッドシート（ペッパーランチPA評価表）
const ROSTER_SS_ID  = '16R13gwiAfU8hNq_hObXS3jgfbkMseodXvuS4JSUNBFQ';   // アルバイト名簿

// ===== シート名 =====
const EVAL_SHEET    = 'クラスターレベル・チェックシート';
const ROSTER_SHEET  = 'スタッフマスタ';
const HISTORY_SHEET = '評価履歴';
const DETAIL_SHEET  = '評価詳細履歴';

// ===== 評価入力シートのセル位置 =====
// ★ 実際のセル位置が違う場合はここだけ修正してください
const C_PERIOD    = 'A1';  // 評価年月
const C_NAME      = 'A3';  // 氏名
const C_GENDER    = 'G3';  // 性別（自動）
const C_AGE       = 'I3';  // 年齢（自動）
const C_TENURE    = 'K3';  // 経験年数（自動）
const C_OLD_CLS   = 'N3';  // 旧クラスター（自動）
const C_NEW_CLS   = 'R3';  // 新クラスター（自動 or 数式）
const C_OLD_WAGE  = 'V3';  // 旧時給（自動）
const C_NEW_WAGE  = 'Y3';  // 新時給（自動 or 数式）
const C_PREV_INFO = 'A4';  // 前回評価情報の表示

// ===== チェック欄の範囲 =====
const R_LEFT  = 'M5:M64';   // 左側60項目（規律17＋マインド14＋共通スキル29）
const R_RIGHT = 'AA5:AA44'; // 右側40項目（デシャップ15＋ウォッシャー6＋キッチン19）
const R_BONUS = 'U57:U63';  // 特別加算（7項目）

// ===== スタッフマスタの列番号（A=1） =====
const MC_NAME    = 4;   // D列: 名前
const MC_GENDER  = 3;   // C列: 男女
const MC_CLUSTER = 10;  // J列: クラスター
const MC_WAGE    = 14;  // N列: 時給
const MC_TENURE  = 24;  // X列: 勤続年数
const MC_AGE     = 25;  // Y列: 年齢
const MC_LEAVE   = 6;   // F列: 退社日（空＝在職者）

// ===== クラスター基準（点数→クラスター→時給） =====
const CLUSTER_TABLE = [
  { min: 90, name: 'マネージャー', wage: 1200 },
  { min: 70, name: '3トレーナー', wage: 1160 },
  { min: 60, name: '2トレーナー', wage: 1150 },
  { min: 50, name: '1トレーナー', wage: 1140 },
  { min: 35, name: '3クルー',    wage: 1130 },
  { min: 25, name: '2クルー',    wage: 1120 },
  { min: 15, name: '1クルー',    wage: 1110 },
  { min: 0,  name: 'ルーキー',   wage: 1100 },
];

// ------------------------------------------------------------
//  onOpen：スプレッドシートを開いたとき自動実行
// ------------------------------------------------------------
function onOpen() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(EVAL_SHEET);
  if (!sheet) return;

  // 評価年月プルダウン（既存コードを統合）
  const today = new Date();
  const year  = today.getFullYear();
  const month = today.getMonth() + 1;
  const periods = (month <= 6)
    ? [year + '年3月', year + '年9月']
    : [year + '年9月', (year + 1) + '年3月'];

  const periodRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(periods, true).build();
  // '@STRING@' でテキスト形式に固定し、日付への自動変換を防ぐ
  sheet.getRange(C_PERIOD).setNumberFormat('@STRING@').setDataValidation(periodRule).setValue(periods[0]);

  // 氏名プルダウン（スタッフマスタから取得）
  buildNameDropdown_(sheet);

  // メニューを追加
  ss.addMenu('📋 評価管理', [
    { name: '✅ 評価を保存',                      functionName: 'saveEvaluation'     },
    null,
    { name: '📊 時給変更集計を更新',              functionName: 'updateSummarySheet' },
    { name: '📝 名簿に時給を反映',               functionName: 'applyToRoster'      },
    null,
    { name: '⚙️ 初回セットアップ（初回のみ）',   functionName: 'setupTrigger'        },
    { name: '🔧 履歴シートを初期化（初回のみ）', functionName: 'initHistorySheets'   },
  ]);
}

// ------------------------------------------------------------
//  氏名プルダウンをスタッフマスタから構築
// ------------------------------------------------------------
function buildNameDropdown_(sheet) {
  try {
    const rosterSS    = SpreadsheetApp.openById(ROSTER_SS_ID);
    const masterSheet = rosterSS.getSheetByName(ROSTER_SHEET);
    const lastRow     = masterSheet.getLastRow();
    if (lastRow < 2) return;

    // 名前列と退社日列を取得し、在職者（退社日が空）のみ抽出
    const data  = masterSheet
      .getRange(2, 1, lastRow - 1, MC_LEAVE)
      .getValues();
    const names = data
      .filter(r => r[MC_NAME - 1] !== '' && r[MC_NAME - 1] !== null && r[MC_LEAVE - 1] === '')
      .map(r => r[MC_NAME - 1]);

    if (names.length === 0) return;

    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(names, true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(C_NAME).setDataValidation(rule);
  } catch (e) {
    console.log('名簿読み込みエラー: ' + e.message);
  }
}

// ------------------------------------------------------------
//  インストール型トリガー登録（初回のみメニューから実行）
// ------------------------------------------------------------
function setupTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // 既存の onEditInstalled トリガーを削除してから再登録（重複防止）
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'onEditInstalled') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onEditInstalled').forSpreadsheet(ss).onEdit().create();
  SpreadsheetApp.getUi().alert('セットアップ完了！\nonEditInstalledトリガーを登録しました。');
}

// ------------------------------------------------------------
//  onEditInstalled：編集時に自動実行（インストール型）
//  ※ openById を使うためシンプルトリガーではなくこちらを使用
// ------------------------------------------------------------
function onEditInstalled(e) {
  const range = e.range;
  const sheet = range.getSheet();
  if (sheet.getName() !== EVAL_SHEET) return;

  // ── チェックセルのトグル処理（スペース+Enter で 1 ↔ 空白）──
  if (range.getNumRows() === 1 && range.getNumColumns() === 1 && isCheckCell_(range, sheet)) {
    const raw = range.getValue();
    const cur = String(raw).trim();
    const old = e.oldValue;
    if (!isNaN(Number(cur)) && Number(cur) > 1) return;
    if ((!old || old === '0') && (cur === '' || cur === '0')) {
      range.setValue(1);
    } else if (old === '1' && cur === '1') {
      range.setValue('');
    }
    return;
  }

  // ── 氏名セル（A3）変更時：スタッフ情報・前回評価を自動入力 ──
  if (range.getA1Notation() !== C_NAME) return;
  const name = range.getValue();
  if (!name || name === '') return;
  loadStaffInfo_(sheet, name);
  loadPreviousInfo_(sheet, name);
}

// ------------------------------------------------------------
//  チェックセル判定（M列・AA列・U列）
// ------------------------------------------------------------
function isCheckCell_(range, sheet) {
  const row = range.getRow();
  const col = range.getColumn();
  if (col === 13 && row >= 5  && row <= 64) return true; // M列（左60項目）
  if (col === 27 && row >= 5  && row <= 44) return true; // AA列（右40項目）
  if (col === 21 && row >= 57 && row <= 63) return true; // U列（特別加算）
  return false;
}

// ------------------------------------------------------------
//  スタッフ情報をスタッフマスタから取得して自動入力
// ------------------------------------------------------------
function loadStaffInfo_(sheet, staffName) {
  try {
    const rosterSS    = SpreadsheetApp.openById(ROSTER_SS_ID);
    const masterSheet = rosterSS.getSheetByName(ROSTER_SHEET);
    const lastRow     = masterSheet.getLastRow();
    if (lastRow < 2) return;

    const data  = masterSheet.getRange(2, 1, lastRow - 1, 30).getValues();
    const found = data.find(row => row[MC_NAME - 1] === staffName);
    if (!found) return;

    sheet.getRange(C_GENDER).setValue(found[MC_GENDER  - 1]); // 性別
    sheet.getRange(C_AGE).setValue(found[MC_AGE     - 1]); // 年齢
    sheet.getRange(C_TENURE).setValue(found[MC_TENURE  - 1]); // 経験年数
    sheet.getRange(C_OLD_CLS).setValue(found[MC_CLUSTER - 1]); // 旧クラスター
    sheet.getRange(C_OLD_WAGE).setValue(found[MC_WAGE   - 1]); // 旧時給
  } catch (e) {
    console.log('スタッフ情報読み込みエラー: ' + e.message);
  }
}

// ------------------------------------------------------------
//  直近評価情報を評価詳細履歴から取得して累積ロード
//  - M5:M64（左60）・AA5:AA44（右40）・U57:U63（特別加算7）を前回値で埋める
//  - ユーザーはその上に新たな加点のみ手動で入力する
// ------------------------------------------------------------
function loadPreviousInfo_(sheet, staffName) {
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const detailSheet = ss.getSheetByName(DETAIL_SHEET);

  // まず全チェック列をクリア（名前切り替え時に前スタッフの値が残らないよう）
  sheet.getRange(R_LEFT).clearContent();
  sheet.getRange(R_RIGHT).clearContent();
  sheet.getRange(R_BONUS).clearContent();
  sheet.getRange(C_PREV_INFO).clearContent();

  if (!detailSheet) {
    sheet.getRange(C_PREV_INFO).setValue('前回評価：なし（初回）');
    return;
  }

  const lastRow = detailSheet.getLastRow();
  if (lastRow < 2) {
    sheet.getRange(C_PREV_INFO).setValue('前回評価：なし（初回）');
    return;
  }

  // 全列取得：名前(1)+期間(1)+合計(1)+左60+右40+特別7 = 110列
  const data = detailSheet.getRange(2, 1, lastRow - 1, 110).getValues();
  let foundRow = null;
  for (let i = data.length - 1; i >= 0; i--) {
    if (data[i][0] === staffName) { foundRow = data[i]; break; }
  }

  if (!foundRow) {
    sheet.getRange(C_PREV_INFO).setValue('前回評価：なし（初回）');
    return;
  }

  const prevPeriod = foundRow[1];
  const prevTotal  = foundRow[2];
  const leftVals   = foundRow.slice(3,   63);  // 60項目（index 3〜62）
  const rightVals  = foundRow.slice(63,  103); // 40項目（index 63〜102）
  const bonusVals  = foundRow.slice(103, 110); // 7項目（index 103〜109）

  // M5:M64 に前回左側チェック値を書き込む（累積ベース）
  sheet.getRange(R_LEFT).setValues(leftVals.map(v => [Number(v) || 0]));
  // AA5:AA44 に前回右側チェック値を書き込む（累積ベース）
  sheet.getRange(R_RIGHT).setValues(rightVals.map(v => [Number(v) || 0]));
  // U57:U63 に前回特別加算を書き込む
  sheet.getRange(R_BONUS).setValues(bonusVals.map(v => [Number(v) || 0]));

  // A4 に前回評価サマリを表示
  sheet.getRange(C_PREV_INFO).setValue(
    `前回評価：${prevPeriod}  合計 ${prevTotal}点（新たに習得した項目のみ追加してください）`
  );
}

// ------------------------------------------------------------
//  点数からクラスターと時給を返す
// ------------------------------------------------------------
function calcCluster_(score) {
  for (const row of CLUSTER_TABLE) {
    if (score >= row.min) return { name: row.name, wage: row.wage };
  }
  return { name: 'ルーキー', wage: 1100 };
}

// ------------------------------------------------------------
//  評価を保存（メニュー or ボタンから実行）
// ------------------------------------------------------------
function saveEvaluation() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(EVAL_SHEET);
  const ui    = SpreadsheetApp.getUi();

  const period = sheet.getRange(C_PERIOD).getDisplayValue();
  const name   = sheet.getRange(C_NAME).getValue();

  if (!name)   { ui.alert('氏名を選択してください。');     return; }
  if (!period) { ui.alert('評価年月を選択してください。'); return; }

  // チェック値を取得
  const leftVals  = sheet.getRange(R_LEFT).getValues().flat().map(v => Number(v) || 0);
  const rightVals = sheet.getRange(R_RIGHT).getValues().flat().map(v => Number(v) || 0);
  const bonusVals = sheet.getRange(R_BONUS).getValues().flat().map(v => Number(v) || 0);

  const leftSum  = leftVals.reduce((a, b) => a + b, 0);
  const rightSum = rightVals.reduce((a, b) => a + b, 0);
  const total    = leftSum + rightSum;
  const bonusSum = bonusVals.reduce((a, b) => a + b, 0);

  // カテゴリ別集計
  const scoreKiritsu    = leftVals.slice(0,  17).reduce((a,b) => a+b, 0); // 規律17点
  const scoreMind       = leftVals.slice(17, 31).reduce((a,b) => a+b, 0); // マインド14点
  const scoreSkill      = leftVals.slice(31, 60).reduce((a,b) => a+b, 0); // 共通スキル29点

  const oldCluster = sheet.getRange(C_OLD_CLS).getValue();
  const oldWage    = sheet.getRange(C_OLD_WAGE).getValue();
  const cluster    = calcCluster_(total);
  const newCluster = cluster.name;
  const newWage    = cluster.wage + bonusSum; // 特別加算を加算

  // 新クラスター・新時給を評価シートに表示
  sheet.getRange(C_NEW_CLS).setValue(newCluster);
  sheet.getRange(C_NEW_WAGE).setValue(newWage);

  // 確認ダイアログ
  const res = ui.alert(
    '評価保存確認',
    `${name}（${period}）\n` +
    `規律：${scoreKiritsu}点　マインド：${scoreMind}点　共通スキル：${scoreSkill}点\n` +
    `ポジション：${rightSum}点\n` +
    `合計：${total}点　特別加算：${bonusSum}円\n` +
    `新クラスター：${newCluster}　新時給：¥${newWage}\n\n` +
    '保存しますか？',
    ui.ButtonSet.YES_NO
  );
  if (res !== ui.Button.YES) return;

  // 履歴シートが無ければ作成
  let histSheet   = ss.getSheetByName(HISTORY_SHEET)   || initHistorySheet_(ss);
  let detailSheet = ss.getSheetByName(DETAIL_SHEET)    || initDetailSheet_(ss);

  // 評価履歴に追記
  histSheet.appendRow([
    name, period,
    scoreKiritsu, scoreMind, scoreSkill, rightSum,
    bonusSum, total,
    oldCluster, newCluster,
    oldWage, newWage,
    '未承認',
    new Date()
  ]);

  // 評価詳細履歴に追記（チェック項目100個＋特別加算7個）
  detailSheet.appendRow([
    name, period, total,
    ...leftVals,   // 60個
    ...rightVals,  // 40個
    ...bonusVals   // 7個
  ]);

  ui.alert(`保存しました。\n${name}（${period}）　合計${total}点`);
}

// ------------------------------------------------------------
//  時給変更集計シートを更新（原本書式で再現・上司提出用）
// ------------------------------------------------------------
function updateSummarySheet() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const evalSh  = ss.getSheetByName(EVAL_SHEET);
  const ui      = SpreadsheetApp.getUi();
  const period  = evalSh.getRange(C_PERIOD).getDisplayValue();

  if (!period) { ui.alert('評価年月を選択してください。'); return; }

  const histSh = ss.getSheetByName(HISTORY_SHEET);
  if (!histSh) { ui.alert('評価履歴がありません。先に評価を保存してください。'); return; }

  const lastRow = histSh.getLastRow();
  if (lastRow < 2) { ui.alert('評価データがありません。'); return; }

  const histData = histSh.getRange(2, 1, lastRow - 1, 14).getValues();
  const targets  = histData.filter(r => normalizePeriod_(r[1]) === period);
  if (targets.length === 0) { ui.alert(`${period}の評価データが見つかりません。`); return; }

  // スタッフマスタから性別・年齢・経験年数を取得
  let masterData = [];
  try {
    const rosterSS = SpreadsheetApp.openById(ROSTER_SS_ID);
    const masterSh = rosterSS.getSheetByName(ROSTER_SHEET);
    masterData = masterSh.getRange(2, 1, masterSh.getLastRow() - 1, 30).getValues();
  } catch (e) { console.log('スタッフマスタエラー: ' + e.message); }

  // シート準備
  let sh = ss.getSheetByName('時給変更集計');
  if (!sh) {
    sh = ss.insertSheet('時給変更集計');
  } else {
    sh.clearContents();
    sh.clearFormats();
    try { sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart(); } catch(e){}
  }

  // 列幅設定（A4縦向きに収まるよう調整）
  const colWidths = [35, 110, 40, 40, 72, 88, 25, 88, 65, 25, 65, 58];
  colWidths.forEach((w, i) => sh.setColumnWidth(i + 1, w));

  // 行高さ設定
  sh.setRowHeight(1, 28); // タイトル
  sh.setRowHeight(2, 22); // 平均時給
  sh.setRowHeight(3, 22); // ヘッダー
  for (let i = 4; i <= 37; i++) sh.setRowHeight(i, 18); // データ34行
  sh.setRowHeight(38, 8);  // 区切り
  for (let i = 39; i <= 41; i++) sh.setRowHeight(i, 18); // フッター

  // 平均時給計算
  const oldWages = targets.map(r => Number(r[10]) || 0).filter(v => v > 0);
  const newWages = targets.map(r => Number(r[11]) || 0).filter(v => v > 0);
  const avgOld   = oldWages.length ? Math.round(oldWages.reduce((a,b)=>a+b,0) / oldWages.length) : 0;
  const avgNew   = newWages.length ? Math.round(newWages.reduce((a,b)=>a+b,0) / newWages.length) : 0;

  // ===== 行1：タイトル =====
  sh.getRange('A1:L1').merge()
    .setValue(`${period}　時給変更届`)
    .setFontSize(13).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBackground('#1F4E79').setFontColor('#FFFFFF');

  // ===== 行2：平均時給 =====
  sh.getRange('A2:B2').merge()
    .setValue('平均時給').setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBackground('#D6E4F7').setFontColor('#000000')
    .setBorder(true, true, true, true, null, null, '#888888', SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange('C2')
    .setValue(`¥${avgOld.toLocaleString()}`).setHorizontalAlignment('center')
    .setBorder(true, true, true, true, null, null, '#888888', SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange('D2')
    .setValue('⇒').setHorizontalAlignment('center')
    .setBorder(true, true, true, true, null, null, '#888888', SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange('E2')
    .setValue(`¥${avgNew.toLocaleString()}`).setHorizontalAlignment('center')
    .setBorder(true, true, true, true, null, null, '#888888', SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange('F2:L2').merge().setBackground('#D6E4F7');

  // ===== 行3：ヘッダー =====
  const hdrs = ['No.','氏名','性別','年齢','経験年数','旧クラスター','→','新クラスター','旧時給','→','新時給','昇給額'];
  sh.getRange(3, 1, 1, 12).setValues([hdrs])
    .setFontWeight('bold').setFontSize(9)
    .setBackground('#4472C4').setFontColor('#FFFFFF')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, '#FFFFFF', SpreadsheetApp.BorderStyle.SOLID);

  // ===== 行4〜37：データ34行 =====
  const DATA_ROWS = 34;
  for (let i = 0; i < DATA_ROWS; i++) {
    const row = 4 + i;
    const t   = targets[i];
    const bg  = i % 2 === 0 ? '#FFFFFF' : '#EBF3FB';
    let vals;

    if (t) {
      const name    = t[0];
      const staff   = masterData.find(r => r[MC_NAME - 1] === name);
      const gender  = staff ? staff[MC_GENDER  - 1] : '';
      const age     = staff ? (staff[MC_AGE    - 1] || '') : '';
      const tenure  = staff ? (staff[MC_TENURE - 1] || '') : '';
      const oldCls  = t[8];
      const newCls  = t[9];
      const oldW    = Number(t[10]) || 0;
      const newW    = Number(t[11]) || 0;
      const diff    = newW - oldW;
      vals = [
        i + 1, name, gender, age, tenure,
        oldCls, '→', newCls,
        `¥${oldW.toLocaleString()}`, '→', `¥${newW.toLocaleString()}`,
        diff > 0 ? `+${diff}円` : diff === 0 ? '変更なし' : `${diff}円`
      ];
    } else {
      vals = [i + 1, '', '', '', '', '', '→', '', '', '→', '', ''];
    }

    const r = sh.getRange(row, 1, 1, 12);
    r.setValues([vals]).setBackground(bg).setFontSize(9).setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true, '#BBBBBB', SpreadsheetApp.BorderStyle.SOLID);

    // 列ごとの横揃え
    [1,3,4,5,6,7,8,10].forEach(c => sh.getRange(row, c).setHorizontalAlignment('center'));
    [9,11,12].forEach(c => sh.getRange(row, c).setHorizontalAlignment('right'));
    sh.getRange(row, 2).setHorizontalAlignment('left');
  }

  // ===== フッター（事業部・店舗・店長）=====
  const footerItems = [
    ['事業部名', '第2事業部'],
    ['店舗名',   'ペッパーランチ飯塚'],
    ['店長名',   '松山　貴志'],
  ];
  footerItems.forEach((fd, i) => {
    sh.getRange(39 + i, 1).setValue(fd[0]).setFontWeight('bold').setHorizontalAlignment('center')
      .setBorder(true, true, true, true, null, null, '#888888', SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(39 + i, 2, 1, 3).merge().setValue(fd[1])
      .setBorder(true, true, true, true, null, null, '#888888', SpreadsheetApp.BorderStyle.SOLID);
  });

  ss.setActiveSheet(sh);
  ui.alert(`${targets.length}件の集計を更新しました。\n承認後に「名簿に時給を反映」を実行してください。`);
}

// ------------------------------------------------------------
//  アルバイト名簿に時給・クラスターを反映（承認済みのみ）
// ------------------------------------------------------------
function applyToRoster() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(EVAL_SHEET);
  const ui    = SpreadsheetApp.getUi();
  const period = sheet.getRange(C_PERIOD).getDisplayValue();

  if (!period) { ui.alert('評価年月を選択してください。'); return; }

  const histSheet = ss.getSheetByName(HISTORY_SHEET);
  if (!histSheet) { ui.alert('評価履歴がありません。'); return; }

  // 承認済みデータの確認
  const lastHist = histSheet.getLastRow();
  if (lastHist < 2) { ui.alert('評価履歴が空です。'); return; }

  const histData = histSheet.getRange(2, 1, lastHist - 1, 13).getValues();
  const approved = histData.filter(r => normalizePeriod_(r[1]) === period && r[12] === '承認済み');

  if (approved.length === 0) {
    ui.alert(`${period}に「承認済み」のデータがありません。\n評価履歴シートのM列を「承認済み」に変更してから実行してください。`);
    return;
  }

  const res = ui.alert(
    '名簿反映確認',
    `${period}の承認済み ${approved.length}件を\nアルバイト名簿に反映します。よろしいですか？`,
    ui.ButtonSet.YES_NO
  );
  if (res !== ui.Button.YES) return;

  const rosterSS    = SpreadsheetApp.openById(ROSTER_SS_ID);
  const masterSheet = rosterSS.getSheetByName(ROSTER_SHEET);
  const lastMaster  = masterSheet.getLastRow();
  const masterData  = masterSheet.getRange(2, 1, lastMaster - 1, 15).getValues();

  let count = 0;
  approved.forEach(row => {
    const targetName = row[0];
    const newCluster = row[9];
    const newWage    = row[11];
    const idx        = masterData.findIndex(r => r[MC_NAME - 1] === targetName);
    if (idx >= 0) {
      masterSheet.getRange(idx + 2, MC_CLUSTER).setValue(newCluster);
      masterSheet.getRange(idx + 2, MC_WAGE).setValue(newWage);
      count++;
    }
  });

  ui.alert(`${count}件のスタッフ情報をアルバイト名簿に更新しました。`);
}

// ------------------------------------------------------------
//  履歴シートの初期化（初回のみ実行）
// ------------------------------------------------------------
function initHistorySheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(HISTORY_SHEET)) initHistorySheet_(ss);
  if (!ss.getSheetByName(DETAIL_SHEET))  initDetailSheet_(ss);
}

function initHistorySheet_(ss) {
  let s = ss.getSheetByName(HISTORY_SHEET);
  if (!s) s = ss.insertSheet(HISTORY_SHEET);
  s.clearContents();
  const h = ['スタッフ名','評価年月','規律点','マインド点','共通スキル点','ポジション点','特別加算','合計点','旧クラスター','新クラスター','旧時給','新時給','承認状況','保存日時'];
  s.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold').setBackground('#1565C0').setFontColor('#FFFFFF');
  s.setFrozenRows(1);
  return s;
}

// Date型でも文字列でも「2026年3月」形式に統一して返す
function normalizePeriod_(v) {
  if (!v) return '';
  if (v instanceof Date) {
    return `${v.getFullYear()}年${v.getMonth() + 1}月`;
  }
  // 「2026/03/01」→「2026年3月」に変換
  const m = String(v).match(/^(\d{4})[\/-]0?(\d{1,2})/);
  if (m) return `${m[1]}年${Number(m[2])}月`;
  return String(v);
}

function initDetailSheet_(ss) {
  let s = ss.getSheetByName(DETAIL_SHEET);
  if (!s) s = ss.insertSheet(DETAIL_SHEET);
  s.clearContents();
  const h = [
    'スタッフ名', '評価年月', '合計点',
    ...Array.from({length: 60}, (_, i) => `左${i+1}`),  // M5:M64
    ...Array.from({length: 40}, (_, i) => `右${i+1}`),  // AA5:AA44
    ...Array.from({length: 7},  (_, i) => `加算${i+1}`) // U57:U63
  ];
  s.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold').setBackground('#2E7D32').setFontColor('#FFFFFF');
  s.setFrozenRows(1);
  return s;
}
