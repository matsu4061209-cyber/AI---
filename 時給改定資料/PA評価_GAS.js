// ============================================================
//  ペッパーランチ PA評価・時給変更管理 GASコード　★MASU_GRID_3★
// ============================================================

// ===== スプレッドシートID =====
const EVAL_SS_ID    = '1jv1lKCDfz_GL14cSSkH9OJEchgZxYOjzbrK5It3VHl0';   // 評価スプレッドシート（ペッパーランチPA評価表）
const ROSTER_SS_ID  = '1AaK0BWbvooMvuFUKX_94TiFq6XkvhS6BclS_FUOcNgE';   // アルバイト名簿（会社ドライブ・本番）

// ===== 店舗情報（初期値）。メニュー「🏢 店舗情報を変更」で各店舗ごとに設定できます =====
const STORE_DIVISION = '第4事業部';
const STORE_NAME     = 'ペッパーランチ飯塚';
const STORE_MANAGER  = '松山　貴志';

// 押印欄の役職名（初期値）。メニュー「🖊️ 押印欄の役職名を変更」で変更できます
const STAMP_TITLES_DEFAULT = ['会長', '社長', '専務', '管理部', '部長', '次長'];

// ===== シート名 =====
const EVAL_SHEET    = 'チェックシート';
const ROSTER_SHEET  = 'スタッフマスタ';
const HISTORY_SHEET = '評価履歴';
const DETAIL_SHEET  = '評価詳細履歴';
const SUMMARY_SHEET = '時給変更届';   // 上司提出・会社回覧用（旧「時給変更集計」）

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
const MC_KUBUN   = 7;   // G列: 区分（社員/高校/大学/主婦/フリーター等）。'社員'は時給変更届の対象外
const MC_CLUSTER = 10;  // J列: クラスター
const MC_WAGE    = 14;  // N列: 時給
const MC_TENURE  = 24;  // X列: 勤続年数
const MC_AGE     = 25;  // Y列: 年齢
const MC_LEAVE   = 6;   // F列: 退社日（空＝在職者）

// ===== クラスター基準（点数→クラスター→基本時給からの加算額）★BASE_WAGE_U47★ =====
//  時給＝「基本時給（チェックシートのU47）」＋各クラスターの加算額。
//  基本時給が変わったら U47 を書き換えるだけで、シートの表もGASの計算も自動で合う。
const BASE_WAGE_CELL    = 'U47';   // 基本時給を入れるセル（チェックシート）
const BASE_WAGE_DEFAULT = 1100;    // U47 が読めないときの初期値
const CLUSTER_TABLE = [
  { min: 90, name: 'マネージャー', add: 100 },
  { min: 70, name: '3トレーナー', add: 60 },
  { min: 60, name: '2トレーナー', add: 50 },
  { min: 50, name: '1トレーナー', add: 40 },
  { min: 35, name: '3クルー',    add: 30 },
  { min: 25, name: '2クルー',    add: 20 },
  { min: 15, name: '1クルー',    add: 10 },
  { min: 0,  name: 'ルーキー',   add: 0  },
];

// ------------------------------------------------------------
//  アルバイト名簿を「同じフォルダの中」から自動で見つける
//  ※ フォルダごとコピーして他店舗に渡しても、その店舗の名簿を
//    自動で参照できるようにするための仕組み。
//    見つからない場合は従来の固定ID（飯塚店・保険）を使う。
// ------------------------------------------------------------
let _cachedRosterSS = null;   // ★ROSTER_IN_FOLDER★ 1回の実行内でキャッシュ（Drive検索を減らす）
function getRosterSS_() {
  if (_cachedRosterSS) return _cachedRosterSS;
  try {
    const ss      = SpreadsheetApp.getActiveSpreadsheet();
    const parents = DriveApp.getFileById(ss.getId()).getParents();
    if (parents.hasNext()) {
      const folder = parents.next();
      const files  = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
      let best = null;
      while (files.hasNext()) {
        const f = files.next();
        if (f.getId() === ss.getId()) continue;               // 自分自身（PA評価表）は除外
        const nm = f.getName();
        if (nm.indexOf('アルバイト名簿') !== -1) { best = f; break; }  // 最優先で確定
        if (!best && nm.indexOf('名簿') !== -1) best = f;              // 次点（あとで上書きされうる）
      }
      if (best) {
        _cachedRosterSS = SpreadsheetApp.openById(best.getId());
        return _cachedRosterSS;
      }
    }
  } catch (e) {
    console.log('名簿の自動検出に失敗（固定IDにフォールバック）: ' + e.message);
  }
  _cachedRosterSS = SpreadsheetApp.openById(ROSTER_SS_ID);   // 保険：飯塚店の固定ID
  return _cachedRosterSS;
}

// ------------------------------------------------------------
//  参照中の名簿を確認（渡した先で「正しい名簿を見ているか」を点検）
// ------------------------------------------------------------
function checkRosterSource() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let folderName   = '(不明)';
    let inFolderName = null;
    const parents = DriveApp.getFileById(ss.getId()).getParents();
    if (parents.hasNext()) {
      const folder = parents.next();
      folderName = folder.getName();
      const files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
      let best = null;
      while (files.hasNext()) {
        const f = files.next();
        if (f.getId() === ss.getId()) continue;
        const nm = f.getName();
        if (nm.indexOf('アルバイト名簿') !== -1) { best = f; break; }
        if (!best && nm.indexOf('名簿') !== -1) best = f;
      }
      if (best) inFolderName = best.getName();
    }
    const roster = getRosterSS_();
    const master = roster.getSheetByName(ROSTER_SHEET);
    const rows   = master ? Math.max(master.getLastRow() - 1, 0) : 0;

    let msg;
    if (inFolderName) {
      msg = '✅ 同じフォルダ内の名簿を参照しています。\n\n'
          + '　フォルダ：' + folderName + '\n'
          + '　名簿ファイル：' + inFolderName + '\n'
          + '　スタッフマスタ「' + ROSTER_SHEET + '」の行数：約' + rows + '人分\n\n'
          + 'このフォルダごとコピーすれば、コピー先でも自動でその店舗の名簿を参照します。';
    } else {
      msg = '⚠️ 同じフォルダ内に「名簿」を含むスプレッドシートが見つかりませんでした。\n'
          + '現在は保険の固定ID（飯塚店の名簿）を参照しています。\n\n'
          + '　名簿：' + roster.getName() + '\n'
          + '　スタッフマスタ「' + ROSTER_SHEET + '」の行数：約' + rows + '人分\n\n'
          + 'このPA評価表と同じフォルダに「アルバイト名簿」という名前のスプレッドシートを置いてください。';
    }
    ui.alert('参照中のアルバイト名簿', msg, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('確認できませんでした：' + e.message);
  }
}

// ------------------------------------------------------------
//  onOpen：スプレッドシートを開いたとき自動実行
// ------------------------------------------------------------
function onOpen() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  // 旧シート名「クラスターレベル・チェックシート」を「チェックシート」に自動リネーム（シート取得の前に実行）
  const oldEval = ss.getSheetByName('クラスターレベル・チェックシート');
  if (oldEval && !ss.getSheetByName(EVAL_SHEET)) oldEval.setName(EVAL_SHEET);
  const sheet = ss.getSheetByName(EVAL_SHEET);
  if (!sheet) return;

  // 評価年月プルダウン（直近2期＋評価履歴にある過去の期間を含める）
  const today = new Date();
  const year  = today.getFullYear();
  const month = today.getMonth() + 1;
  const recent = (month <= 6)
    ? [year + '年3月', year + '年9月']
    : [year + '年9月', (year + 1) + '年3月'];

  // 評価履歴にある過去の期間を追加
  const allPeriods = [...recent];
  try {
    const histSh = ss.getSheetByName(HISTORY_SHEET);
    if (histSh && histSh.getLastRow() >= 2) {
      histSh.getRange(2, 1, histSh.getLastRow() - 1, 2).getValues()
        .map(r => normalizePeriod_(r[1])).filter(Boolean)
        .forEach(p => { if (!allPeriods.includes(p)) allPeriods.push(p); });
    }
  } catch(e) {}

  const periodRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(allPeriods, true).build();
  // '@STRING@' でテキスト形式に固定し、日付への自動変換を防ぐ
  const currentPeriod = sheet.getRange(C_PERIOD).getDisplayValue();
  sheet.getRange(C_PERIOD).setNumberFormat('@STRING@').setDataValidation(periodRule);
  if (!currentPeriod) sheet.getRange(C_PERIOD).setValue(recent[0]);

  // 氏名プルダウン（スタッフマスタから取得）
  buildNameDropdown_(sheet);

  // メニューを追加（用途ごとにサブメニューで区分）
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📋 評価管理')
    .addSubMenu(ui.createMenu('📝 チェックシート')
      .addItem('👤 氏名を選ぶ（クラスター表示）', 'showNamePicker')
      .addItem('✅ 評価を保存', 'saveEvaluation'))
    .addSubMenu(ui.createMenu('📊 時給変更届')
      .addItem('📊 時給変更届を更新（何度でも可）', 'updateSummarySheet')
      .addItem('🏢 店舗情報を変更（事業部名・店舗名・店長名）', 'setStoreInfo')
      .addItem('🖊️ 押印欄の役職名を変更', 'setStampTitles'))
    .addSubMenu(ui.createMenu('🏢 全体')
      .addItem('📄 PDF一括出力（提出時用）', 'exportAllPDFs')
      .addItem('📝 名簿に新クラスターを反映（上長申請後）', 'applyToRoster')
      .addSeparator()
      .addItem('⚙️ 初回セットアップ／最新に更新', 'setupTrigger')
      .addSeparator()
      .addItem('🏪 他店舗渡し用：個人情報をクリア', 'clearForTransfer'))
    .addToUi();
}

// ------------------------------------------------------------
//  氏名プルダウンをスタッフマスタから構築
// ------------------------------------------------------------
function buildNameDropdown_(sheet) {
  const names = getRosterStaff_().map(s => s.name);
  // 0名のときは何もしない（既存プルダウンは触らない）。
  // ※onOpen（簡易トリガー）は名簿を読めず常に0名になるため、ここで消すと開くたびにプルダウンが消える。
  //   渡し用の氏名リスト削除は clearForTransfer 側（メニュー実行・権限あり）で行う。
  if (names.length === 0) return;

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(names, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(C_NAME).setDataValidation(rule);
}

// ------------------------------------------------------------
//  氏名選択画面を開く（表示は「氏名＋旧クラスター」、
//  クリックするとセルには氏名だけが入る）
// ------------------------------------------------------------
function showNamePicker() {
  const staff = getRosterStaff_();
  if (staff.length === 0) {
    SpreadsheetApp.getUi().alert('名簿からスタッフを読み込めませんでした。');
    return;
  }
  // 氏名プルダウンも最新の名簿で作り直す（開いている最中に追加した人もここで反映）★AUTO_DROPDOWN★
  const evalSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(EVAL_SHEET);
  if (evalSh) buildNameDropdown_(evalSh);

  const rows = staff.map(s =>
    `<div class="row" onclick="pick('${s.name.replace(/'/g, "\\'")}')">` +
    `<span class="nm">${s.name}</span>` +
    `<span class="cl">${s.cluster || '－'}</span></div>`
  ).join('');

  const html = HtmlService.createHtmlOutput(
    `<style>
       body   { font-family: sans-serif; margin: 0; padding: 4px; }
       .row   { display: flex; justify-content: space-between; align-items: center;
                padding: 9px 12px; border-bottom: 1px solid #eee; cursor: pointer; }
       .row:hover { background: #e8f0fe; }
       .nm    { font-size: 14px; }
       .cl    { font-size: 12px; color: #666; }
       #msg   { padding: 10px; text-align: center; color: #1a73e8; display: none; }
     </style>
     <div id="list">${rows}</div>
     <div id="msg">読み込み中…</div>
     <script>
       function pick(name) {
         document.getElementById('list').style.display = 'none';
         document.getElementById('msg').style.display  = 'block';
         google.script.run
           .withSuccessHandler(function() { google.script.host.close(); })
           .withFailureHandler(function(err) {
             document.getElementById('msg').textContent = 'エラー: ' + err.message;
           })
           .applyPickedName(name);
       }
     </script>`
  ).setWidth(320).setHeight(460);

  SpreadsheetApp.getUi().showModalDialog(html, '氏名を選ぶ');
}

// ------------------------------------------------------------
//  選択画面から呼ばれる（氏名だけをセルに入れて情報を読み込む）
//  ※ google.script.run から呼ぶため末尾に _ を付けないこと
// ------------------------------------------------------------
function applyPickedName(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(EVAL_SHEET);
  sheet.getRange(C_NAME).setValue(name);
  loadStaffInfo_(sheet, name);
  loadPreviousInfo_(sheet, name);
}

// ------------------------------------------------------------
//  名簿から在職スタッフ（氏名・クラスター）を取得
// ------------------------------------------------------------
function getRosterStaff_() {
  try {
    const rosterSS    = getRosterSS_();
    const masterSheet = rosterSS.getSheetByName(ROSTER_SHEET);
    const lastRow     = masterSheet.getLastRow();
    if (lastRow < 2) return [];

    return masterSheet
      .getRange(2, 1, lastRow - 1, Math.max(MC_LEAVE, MC_CLUSTER))
      .getValues()
      .filter(r => r[MC_NAME - 1] !== '' && r[MC_NAME - 1] !== null && r[MC_LEAVE - 1] === '')
      .map(r => ({ name: r[MC_NAME - 1], cluster: r[MC_CLUSTER - 1] }));
  } catch (e) {
    console.log('名簿読み込みエラー: ' + e.message);
    return [];
  }
}

// ------------------------------------------------------------
//  氏名プルダウンを手動更新（名簿に新スタッフ追加後に使用）
// ------------------------------------------------------------
function refreshNameDropdown() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(EVAL_SHEET);
  if (!sheet) { SpreadsheetApp.getUi().alert('評価シートが見つかりません'); return; }
  buildNameDropdown_(sheet);
  SpreadsheetApp.getUi().alert('✅ 氏名プルダウンを更新しました！\n新しいスタッフが追加されています。');
}

// ------------------------------------------------------------
//  インストール型トリガー登録（初回のみメニューから実行）
// ------------------------------------------------------------
function setupTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // ① 既存の onEditInstalled / onOpenInstalled トリガーを削除してから再登録（重複防止）
  ScriptApp.getProjectTriggers().forEach(t => {
    const fn = t.getHandlerFunction();
    if (fn === 'onEditInstalled' || fn === 'onOpenInstalled') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onEditInstalled').forSpreadsheet(ss).onEdit().create();
  // 開いたとき自動で氏名プルダウンを最新の名簿で作り直す（★AUTO_DROPDOWN★）
  ScriptApp.newTrigger('onOpenInstalled').forSpreadsheet(ss).onOpen().create();
  // ② 履歴シート（評価履歴・評価詳細履歴）が無ければ作成（既存は消さない）
  if (!ss.getSheetByName(HISTORY_SHEET)) initHistorySheet_(ss);
  if (!ss.getSheetByName(DETAIL_SHEET))  initDetailSheet_(ss);
  // ③ PDF出力に使う外部リクエスト権限を、このメニュー実行時に承認させる（初回は承認画面が出る）
  try { UrlFetchApp.fetch('https://www.googleapis.com/discovery/v1/apis', { muteHttpExceptions: true }); } catch(e) {}
  // ④ 古い版のコピー向け：チェックシートの数式（Z60・U56）を最新に修正（★COPY_UPDATE★）
  fixChecksheetFormulas_(ss);
  // ⑤ 保存済みの店舗情報があれば、時給変更届のフッターへ反映（★STORE_INFO_APPLY★）
  applyStoreInfoToSummary_(ss);
  // 登録直後に一度プルダウンを最新化（初回セットアップ時点で反映）
  const evalSh = ss.getSheetByName(EVAL_SHEET);
  if (evalSh) buildNameDropdown_(evalSh);
  SpreadsheetApp.getUi().alert('セットアップ／更新 完了！\n・自動動作トリガーを登録しました（編集時＋開いたとき）\n・開くたびに氏名プルダウンを名簿から自動更新します\n・履歴シートを用意しました\n・チェックシートの数式（総計など）を最新に修正しました\n・PDF出力の権限を承認しました');
}

// ------------------------------------------------------------
//  チェックシートの数式を最新版に修正（★COPY_UPDATE★）
//  ※ 配布済みの古いコピーは Z60（総計）に 1100 固定やクラスター境界のズレが
//    残っていることがあるため、飯塚の正しい数式で上書きして揃える。
//  ※ 何度実行しても同じ結果（安全）。飯塚で実データ確認済みの数式を使用。
// ------------------------------------------------------------
function fixChecksheetFormulas_(ss) {
  const sh = ss.getSheetByName(EVAL_SHEET);
  if (!sh) return;
  // 総計時給 Z60：低点数(ルーキー1〜14点)・勤続3ヶ月未満は基本給(U47)、それ以外は表引き＋特別加算
  const z60 = `=IF(K3="","",IF(OR(AA59<=14,VALUE(LEFT(K3,FIND("年",K3)-1))*12+VALUE(MID(K3,FIND("年",K3)+1,FIND("ヶ",K3)-FIND("年",K3)-1))<3),U47,VLOOKUP(AA59,R48:V54,4,TRUE)+SUM(U57:V63)))`;
  // U56：表引きが#N/Aなら基本給(U47)にフォールバック
  const u56 = `=IFERROR(VLOOKUP(AA59,R48:V54,4,TRUE),U47)`;
  try { sh.getRange('Z60').setFormula(z60); } catch (e) { console.log('Z60修正エラー: ' + e.message); }
  try { sh.getRange('U56').setFormula(u56); } catch (e) { console.log('U56修正エラー: ' + e.message); }
}

// ============================================================
//  権限承認用（GASエディタから1回だけ「実行」する）
//  ※ UrlFetchApp・DriveApp を使うので、エディタで実行すると
//    「承認が必要です」画面が出て、PDF出力に必要な権限を承認できる
// ============================================================
function authorizeNow() {
  UrlFetchApp.fetch('https://www.googleapis.com/discovery/v1/apis', { muteHttpExceptions: true });
  DriveApp.getRootFolder();  // Drive権限も確定
  Logger.log('✅ 権限の承認が完了しました。スプレッドシートに戻って「PDF一括出力」を実行してください。');
}

// ------------------------------------------------------------
//  シートを単独でPDF化（UrlFetchApp不要・Drive権限だけで動くフォールバック）
//  ※ 一時スプレッドシートに対象シートだけ複製してPDF化し、後で削除する
// ------------------------------------------------------------
function sheetToPdfBlob_(srcSheet, fileName) {
  const tempSS = SpreadsheetApp.create('_pa_pdf_tmp');
  const tempId = tempSS.getId();
  try {
    const copied = srcSheet.copyTo(tempSS);
    copied.setName('印刷');
    // 数式を現在値に固定（他シート参照が#REFになるのを防ぐ）
    const rng = copied.getDataRange();
    rng.copyTo(rng, { contentsOnly: true });
    // 既定の空シートを削除して対象シートだけ残す
    tempSS.getSheets().forEach(s => { if (s.getSheetId() !== copied.getSheetId()) tempSS.deleteSheet(s); });
    SpreadsheetApp.flush();
    return DriveApp.getFileById(tempId).getAs('application/pdf').setName(fileName);
  } finally {
    try { DriveApp.getFileById(tempId).setTrashed(true); } catch (e) {}
  }
}

// ------------------------------------------------------------
//  onEdit：シンプルトリガー（承認不要・素早く動作）
//  チェックセルのトグルを担当
// ------------------------------------------------------------
function onEdit(e) {
  // ★NO_TOGGLE★ 点数セルの自動トグル(setValue)は廃止。
  // スクリプトが編集直後のセルを書き換えるとカーソルが元セルに戻される
  // （Googleスプレッドシートの既知のクセ）ため、Enterで下に進めなかった。
  // 点数セルは普通のセルとして扱い、標準のEnter動作（下へ移動）を妨げない。
  // 点数は手入力（1を打つだけ）。氏名変更時の自動入力は onEditInstalled 側で処理。
}

// ------------------------------------------------------------
//  氏名セル判定（結合セルでも判定できるよう行・列で比較）
// ------------------------------------------------------------
function isNameCell_(range) {
  const nameCell = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(EVAL_SHEET).getRange(C_NAME);
  return range.getRow() === nameCell.getRow()
      && range.getColumn() === nameCell.getColumn();
}

// ------------------------------------------------------------
//  onEditInstalled：編集時に自動実行（インストール型）
//  ※ openById を使うためシンプルトリガーではなくこちらを使用
// ------------------------------------------------------------
function onEditInstalled(e) {
  const range = e.range;
  const sheet = range.getSheet();
  if (sheet.getName() !== EVAL_SHEET) return;

  // ── チェックセルは onEdit 側で処理済み ──
  if (range.getNumRows() === 1 && range.getNumColumns() === 1 && isCheckCell_(range, sheet)) return;

  // ── 氏名セル（A3）変更時：スタッフ情報・前回評価を自動入力 ──
  if (!isNameCell_(range)) return;
  const name = String(sheet.getRange(C_NAME).getValue()).trim();
  if (!name) return;
  loadStaffInfo_(sheet, name);
  loadPreviousInfo_(sheet, name);
}

// ------------------------------------------------------------
//  onOpenInstalled：開いたとき自動実行（インストール型）★AUTO_DROPDOWN★
//  ※ 権限ありで同フォルダのアルバイト名簿を読めるので、開くたびに
//    氏名プルダウンを最新の名簿で作り直す（手動更新ボタンを不要にする）。
//  ※ 簡易トリガー onOpen は名簿を読めない（0名）ため、そちらでは更新しない。
// ------------------------------------------------------------
function onOpenInstalled() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(EVAL_SHEET);
  if (sheet) buildNameDropdown_(sheet);
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
    const rosterSS    = getRosterSS_();
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
  sheet.getRange(C_NEW_CLS).clearContent();
  sheet.getRange(C_NEW_WAGE).clearContent();

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
  sheet.getRange(R_LEFT).setValues(leftVals.map(v => [Number(v) === 1 ? 1 : '']));
  // AA5:AA44 に前回右側チェック値を書き込む（累積ベース）
  sheet.getRange(R_RIGHT).setValues(rightVals.map(v => [Number(v) === 1 ? 1 : '']));
  // U57:U63 に前回特別加算を書き込む
  sheet.getRange(R_BONUS).setValues(bonusVals.map(v => [Number(v) === 1 ? 1 : '']));

  // A4 に前回評価サマリを表示
  sheet.getRange(C_PREV_INFO).setValue(
    `前回評価：${normalizePeriod_(prevPeriod)}  合計 ${prevTotal}点（新たに習得した項目のみ追加してください）`
  );

  // 前回評価の結果（新クラスター・新時給）も再現して表示する（★SHOW_PREV_RESULT★）
  // 評価は終わっているが、訂正時に見えるように、保存時と同じ計算で埋める。
  // ※ 復元済みのチェックセルから計算するので、saveEvaluation と必ず一致する。
  const lVals = sheet.getRange(R_LEFT).getValues().flat().map(v => Number(v) || 0);
  const rVals = sheet.getRange(R_RIGHT).getValues().flat().map(v => Number(v) || 0);
  const bVals = sheet.getRange(R_BONUS).getValues().flat().map(v => Number(v) || 0);
  const tot   = lVals.reduce((a, b) => a + b, 0) + rVals.reduce((a, b) => a + b, 0);
  const bSum  = bVals.reduce((a, b) => a + b, 0);
  const cl    = calcCluster_(tot);
  sheet.getRange(C_NEW_CLS).setValue(cl.name);
  sheet.getRange(C_NEW_WAGE).setValue(cl.wage + bSum);
}

// ------------------------------------------------------------
//  点数からクラスターと時給を返す
// ------------------------------------------------------------
function calcCluster_(score) {
  const base = getBaseWage_();
  for (const row of CLUSTER_TABLE) {
    if (score >= row.min) return { name: row.name, wage: base + row.add };
  }
  return { name: 'ルーキー', wage: base };
}

// ------------------------------------------------------------
//  基本時給を取得（チェックシートのU47。読めなければ初期値1100）
// ------------------------------------------------------------
function getBaseWage_() {
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(EVAL_SHEET);
    const v  = Number(sh.getRange(BASE_WAGE_CELL).getValue());
    if (v > 0) return v;
  } catch (e) { console.log('基本時給(U47)の取得エラー: ' + e.message); }
  return BASE_WAGE_DEFAULT;
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

  // 時給変更届も自動で最新に（その人の行だけ更新。できなければ全体を作り直す）
  const np = normalizePeriod_(period);
  let summaryMsg = '';
  try {
    const ok = refreshSummaryForStaff_(ss, np, String(name).trim());
    if (!ok) runUpdateSummary_(np);
    SpreadsheetApp.flush();
    summaryMsg = '\n時給変更届も更新しました。';
  } catch (e) {
    summaryMsg = '\n※時給変更届の自動更新に失敗しました（「時給変更届を更新」で手動更新してください）。';
    console.log('時給変更届の自動更新エラー: ' + e.message);
  }

  ui.alert(`保存しました。\n${name}（${period}）　合計${total}点${summaryMsg}`);
}

// ------------------------------------------------------------
//  時給変更集計シートを更新（原本書式で再現・上司提出用）
// ------------------------------------------------------------
function updateSummarySheet() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const ui  = SpreadsheetApp.getUi();

  const histSh = ss.getSheetByName(HISTORY_SHEET);
  if (!histSh) { ui.alert('評価履歴がありません。先に評価を保存してください。'); return; }

  const lastRow = histSh.getLastRow();
  if (lastRow < 2) { ui.alert('評価データがありません。'); return; }

  const histData = histSh.getRange(2, 1, lastRow - 1, 14).getValues();

  // 評価履歴にある期間を取得（新しい順）
  const availPeriods = [...new Set(histData.map(r => normalizePeriod_(r[1])).filter(Boolean))].sort().reverse();
  if (availPeriods.length === 0) { ui.alert('評価データがありません。'); return; }

  // 1件なら即実行、複数なら「ボタンで選ぶ」画面（入力しない）
  if (availPeriods.length === 1) { ui.alert(runUpdateSummary_(availPeriods[0])); return; }

  const btns = availPeriods.map(p =>
    `<button class="pd" onclick="pick('${p}')">${p}</button>`).join('');
  const html = HtmlService.createHtmlOutput(
    `<style>
       body { font-family: sans-serif; margin:0; padding:14px; }
       .h   { font-size:13px; color:#333; margin-bottom:12px; }
       .pd  { display:block; width:100%; font-size:15px; padding:11px; margin-bottom:8px;
              border:1px solid #1a73e8; background:#fff; color:#1a73e8; border-radius:6px; cursor:pointer; }
       .pd:hover { background:#e8f0fe; }
       #msg { display:none; text-align:center; color:#1a73e8; margin-top:6px; white-space:pre-wrap; font-size:13px; }
       .close { display:none; width:100%; font-size:14px; padding:9px; margin-top:12px;
                border:none; background:#1a73e8; color:#fff; border-radius:6px; cursor:pointer; }
     </style>
     <div class="h">時給変更届を作る評価年月を選んでください：</div>
     <div id="list">${btns}</div>
     <div id="msg"></div>
     <button id="closeBtn" class="close" onclick="google.script.host.close()">閉じる</button>
     <script>
       function pick(p) {
         document.getElementById('list').style.display='none';
         var m=document.getElementById('msg');
         m.style.display='block'; m.textContent='時給変更届を作成しています…\\nしばらくお待ちください。';
         google.script.run
           .withSuccessHandler(function(res){
             m.textContent = res;
             document.getElementById('closeBtn').style.display='block';
           })
           .withFailureHandler(function(err){
             m.style.color='#d33';
             m.textContent='エラー: '+err.message;
             document.getElementById('closeBtn').style.display='block';
           })
           .updateSummaryForPeriod(p);
       }
     </script>`
  ).setWidth(320).setHeight(150 + availPeriods.length * 52);

  ui.showModalDialog(html, '📊 時給変更届を更新');
}

// 選択画面から呼ばれる（google.script.run 用・末尾 _ なし）
function updateSummaryForPeriod(period) {
  return runUpdateSummary_(period);
}

// ===== 時給変更届グリッドの共通定義（renderSummarySheet_ と 行単位更新で共用）★SUMMARY_SHARED★ =====
const SUMMARY_GRID_COLS = 24;
const SUMMARY_SEG   = { no:1, name:3, gender:2, age:1, tenure:3, oldCls:3, arrow1:1,
                        newCls:3, oldWage:2, arrow2:1, newWage:2, diff:2 };
const SUMMARY_ORDER = ['no','name','gender','age','tenure','oldCls','arrow1','newCls','oldWage','arrow2','newWage','diff'];
const SUMMARY_START = (function(){ const s = {}; let c = 1; SUMMARY_ORDER.forEach(k => { s[k] = c; c += SUMMARY_SEG[k]; }); return s; })();

// 同じ氏名で複数行あれば、日付（列14＝index13）が最新の1件だけを残す
function latestPerName_(rows, nameIdx, dateIdx) {
  const best = {}, bestT = {};
  rows.forEach(r => {
    const nm  = String(r[nameIdx]).trim();
    const raw = r[dateIdx];
    const t   = (raw instanceof Date) ? raw.getTime() : (raw ? new Date(raw).getTime() : 0);
    if (!(nm in best) || t >= bestT[nm]) { best[nm] = r; bestT[nm] = t; }
  });
  return Object.keys(best).map(k => best[k]);
}

// 時給変更届の元データ（在籍アルバイト全員の旧→新）を作る。{rows, evalCount} を返す。
// onlyNames（配列）を渡すと、その氏名だけに絞る（PDF出力で選んだ人だけの時給変更届に使う）。
function buildSummaryRows_(ss, period, onlyNames) {
  const histSh = ss.getSheetByName(HISTORY_SHEET);
  if (!histSh || histSh.getLastRow() < 2) return { rows: [], evalCount: 0 };
  const histData = histSh.getRange(2, 1, histSh.getLastRow() - 1, 14).getValues();
  // 同一年月・同一人物は最新の1件だけを採用（日付＝index13）
  const targets = latestPerName_(histData.filter(r => normalizePeriod_(r[1]) === period), 0, 13);

  let masterData = [];
  try {
    const masterSh = getRosterSS_().getSheetByName(ROSTER_SHEET);
    masterData = masterSh.getRange(2, 1, masterSh.getLastRow() - 1, 30).getValues();
  } catch (e) { console.log('スタッフマスタエラー: ' + e.message); }

  // 評価履歴（今回の期間）を氏名で引けるようにする（r[9]=新クラスター, r[11]=新時給）
  const evalMap = {};
  targets.forEach(r => { evalMap[String(r[0]).trim()] = r; });

  // 在籍アルバイト全員（退社日が空・区分が「社員」以外）を対象。旧＝マスタ現在値、新＝評価があれば上書き（無ければ旧と同じ＝変更なし）
  const nameSet = (onlyNames && onlyNames.length)
    ? new Set(onlyNames.map(n => String(n).trim())) : null;
  const rows = masterData
    .filter(r => r[MC_NAME - 1] !== '' && r[MC_NAME - 1] !== null && r[MC_LEAVE - 1] === ''
              && String(r[MC_KUBUN - 1]).trim() !== '社員')
    .filter(r => !nameSet || nameSet.has(String(r[MC_NAME - 1]).trim()))
    .map(r => {
      const name   = String(r[MC_NAME - 1]).trim();
      const oldCls = r[MC_CLUSTER - 1];
      const oldW   = Number(r[MC_WAGE - 1]) || 0;
      const ev     = evalMap[name];
      return {
        name:   name,
        gender: r[MC_GENDER - 1] || '',
        age:    r[MC_AGE    - 1] || '',
        tenure: r[MC_TENURE - 1] || '',
        oldCls: oldCls,
        oldW:   oldW,
        newCls: ev ? ev[9] : oldCls,
        newW:   ev ? (Number(ev[11]) || 0) : oldW
      };
    });
  return { rows: rows, evalCount: targets.length };
}

// 実際の時給変更届の生成（全体を作り直す）。結果メッセージ（文字列）を返す。※ getUi は使わない
function runUpdateSummary_(period) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const built = buildSummaryRows_(ss, period);
  if (built.evalCount === 0)     return `${period}の評価データが見つかりません。`;
  if (built.rows.length === 0)   return '在籍スタッフが取得できませんでした。スタッフマスタを確認してください。';
  return renderSummarySheet_(ss, period, built.rows, false);
}

// 評価保存時：その人の行だけ時給変更届を更新（速い）。できなければ false（呼び出し側で全体を作り直す）
function refreshSummaryForStaff_(ss, period, name) {
  const sh = ss.getSheetByName(SUMMARY_SHEET);
  if (!sh) return false;                                        // シートが無い
  const title = String(sh.getRange(1, 1).getValue());
  if (!period || title.indexOf(period) === -1) return false;    // 別の期（タイトル不一致）
  const built = buildSummaryRows_(ss, period);
  if (!built.rows.length) return false;

  // 氏名の行を探す（データは4行目から。氏名は SUMMARY_START.name 列）
  const DATA_TOP = 4;
  const n = Math.max(sh.getLastRow() - DATA_TOP + 1, 0);
  if (n === 0) return false;
  const nameCol = sh.getRange(DATA_TOP, SUMMARY_START.name, n, 1).getValues();
  let rowIdx = -1;
  for (let i = 0; i < nameCol.length; i++) {
    if (String(nameCol[i][0]).trim() === name) { rowIdx = DATA_TOP + i; break; }
  }
  if (rowIdx === -1) return false;                              // 居ない（新人など）
  const me = built.rows.find(r => r.name === name);
  if (!me) return false;

  // 当人の行を丸ごと最新化（性別・年齢・経験年数・旧新クラスター・旧新時給・昇給額）
  //  ※旧クラスター等も名簿の最新に合わせる（更新漏れ防止）
  const diff = me.newW - me.oldW;
  sh.getRange(rowIdx, SUMMARY_START.gender).setValue(me.gender);
  sh.getRange(rowIdx, SUMMARY_START.age).setValue(me.age);
  sh.getRange(rowIdx, SUMMARY_START.tenure).setValue(me.tenure);
  sh.getRange(rowIdx, SUMMARY_START.oldCls).setValue(me.oldCls);
  sh.getRange(rowIdx, SUMMARY_START.newCls).setValue(me.newCls);
  sh.getRange(rowIdx, SUMMARY_START.oldWage).setValue(`¥${me.oldW.toLocaleString()}`);
  sh.getRange(rowIdx, SUMMARY_START.newWage).setValue(`¥${me.newW.toLocaleString()}`);
  sh.getRange(rowIdx, SUMMARY_START.diff).setValue(diff > 0 ? `+${diff}円` : diff < 0 ? `${diff}円` : '変更なし');

  // 平均時給（行2）を再計算（在籍全員の旧・新から）
  const oldWs = built.rows.map(r => r.oldW).filter(v => v > 0);
  const newWs = built.rows.map(r => r.newW).filter(v => v > 0);
  const avgOld = oldWs.length ? Math.round(oldWs.reduce((a,b)=>a+b,0) / oldWs.length) : 0;
  const avgNew = newWs.length ? Math.round(newWs.reduce((a,b)=>a+b,0) / newWs.length) : 0;
  const avgMid = Math.floor((5 + SUMMARY_GRID_COLS) / 2);
  sh.getRange(2, 5).setValue(`変更前　¥${avgOld.toLocaleString()}`);
  sh.getRange(2, avgMid + 1).setValue(`変更後　¥${avgNew.toLocaleString()}`);
  return true;
}

// ------------------------------------------------------------
//  渡し用テンプレート：スタッフ0名でも「見出し＋押印欄だけの空フォーム」を描く
//  （事業部名/店舗名/店長名の見出しと、会長〜次長の押印欄が残るようにする）
// ------------------------------------------------------------
function buildBlankSummaryForm_(ss) {
  return renderSummarySheet_(ss, '', [], true);
}

// ------------------------------------------------------------
//  時給変更届シートを実際に描画する（runUpdateSummary_ とテンプレートで共用）
//  templateMode=true のときは店舗情報（事業部名等の値）を空にして渡し用の空フォームにする
// ------------------------------------------------------------
function renderSummarySheet_(ss, period, rows, templateMode) {
  // シート準備
  // 新シート名で探し、無ければ旧シート名（時給変更集計）を引き継いで改名
  let sh = ss.getSheetByName(SUMMARY_SHEET) || ss.getSheetByName('時給変更集計');
  if (!sh) {
    sh = ss.insertSheet(SUMMARY_SHEET);
  } else {
    if (sh.getName() !== SUMMARY_SHEET) sh.setName(SUMMARY_SHEET);
    sh.clearContents();
    sh.clearFormats();
    try { sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart(); } catch(e){}
  }

  // 行レイアウトの位置（全員分で行数が変わるため動的に計算）
  const DATA_TOP  = 4;
  const DATA_ROWS = rows.length;
  const DATA_END  = DATA_TOP + DATA_ROWS - 1;
  const SEP_ROW   = DATA_END + 1;
  const FOOT_TOP  = SEP_ROW + 1;   // フッター3行
  const FOOT_END  = FOOT_TOP + 2;
  const SPACER    = FOOT_END + 1;  // A4下部へ押し下げる区切り
  const STAMP_LBL = SPACER + 1;    // 押印欄ラベル
  const STAMP_BOX = STAMP_LBL + 1; // 押印スペース

  // ===== 升目グリッド（24等分で押印欄を完全等幅）★MASU_GRID_3★ =====
  //  列位置は共通定義（★SUMMARY_SHARED★）を使用。行単位更新と必ず一致させる。
  const GRID_COLS = SUMMARY_GRID_COLS;
  const CELL_W    = 32;             // 各升の幅(px)。24×32=768px。印刷は「ページに合わせる」で横幅いっぱいに広がる
  // シートの列数が足りない場合は不足分を追加（範囲外エラー防止）
  if (sh.getMaxColumns() < GRID_COLS) sh.insertColumnsAfter(sh.getMaxColumns(), GRID_COLS - sh.getMaxColumns());
  for (let c = 1; c <= GRID_COLS; c++) sh.setColumnWidth(c, CELL_W);

  const SEG   = SUMMARY_SEG;
  const ORDER = SUMMARY_ORDER;
  const START = SUMMARY_START;

  // 指定項目の升を結合するヘルパー
  const seg = (row, key, val) => {
    const rng = sh.getRange(row, START[key], 1, SEG[key]);
    if (SEG[key] > 1) rng.merge();
    if (val !== undefined) rng.setValue(val);
    return rng;
  };

  const BLUE_D = '#1F4E79', BLUE_H = '#4472C4', BLUE_L = '#D6E4F7', BLUE_R = '#EBF3FB';
  const GRAY = '#888888', LINE = '#BBBBBB';
  const SOLID = SpreadsheetApp.BorderStyle.SOLID;

  // 行高さ設定
  const H_TITLE = 28, H_AVG = 22, H_HDR = 22, H_DATA = 18, H_SEP1 = 8,
        H_FOOT = 18, H_STAMP_LBL = 20, H_STAMP_BOX = 58;
  sh.setRowHeight(1, H_TITLE); // タイトル
  sh.setRowHeight(2, H_AVG);   // 平均時給
  sh.setRowHeight(3, H_HDR);   // ヘッダー
  for (let r = DATA_TOP; r <= DATA_END; r++) sh.setRowHeight(r, H_DATA); // データ（在籍全員）
  sh.setRowHeight(SEP_ROW, H_SEP1); // 区切り
  for (let r = FOOT_TOP; r <= FOOT_END; r++) sh.setRowHeight(r, H_FOOT); // フッター
  sh.setRowHeight(STAMP_LBL, H_STAMP_LBL); // 押印欄：役職ラベル
  sh.setRowHeight(STAMP_BOX, H_STAMP_BOX); // 押印欄：押印スペース

  // ★押印欄を用紙下部に置く。前回の高さで約5cm(≈190px)はみ出したので、1枚に収まる高さに調整
  const A4_ONE_PAGE_H = 1000;  // A4縦1枚に収まる高さ(px)の実測目安（これより上に押印欄を置く）
  const usedH = H_TITLE + H_AVG + H_HDR + DATA_ROWS * H_DATA + H_SEP1 + 3 * H_FOOT
              + H_STAMP_LBL + H_STAMP_BOX;
  sh.setRowHeight(SPACER, Math.max(24, A4_ONE_PAGE_H - usedH)); // 押印欄を下部に、かつ1枚に収める

  // 平均時給計算（在籍全員の旧・新）
  const oldWages = rows.map(r => r.oldW).filter(v => v > 0);
  const newWages = rows.map(r => r.newW).filter(v => v > 0);
  const avgOld   = oldWages.length ? Math.round(oldWages.reduce((a,b)=>a+b,0) / oldWages.length) : 0;
  const avgNew   = newWages.length ? Math.round(newWages.reduce((a,b)=>a+b,0) / newWages.length) : 0;

  // ===== 行1：タイトル =====
  sh.getRange(1, 1, 1, GRID_COLS).merge()
    .setValue(templateMode ? '時給変更届' : `${period}　時給変更届`)
    .setFontSize(13).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBackground(BLUE_D).setFontColor('#FFFFFF');

  // ===== 行2：平均時給（見出しはA〜D列、E列から変更前→矢印→変更後）=====
  const avgMid = Math.floor((5 + GRID_COLS) / 2); // 変更前と変更後の境目（矢印の位置）
  sh.getRange(2, 1, 1, 4).merge()   // A〜D：見出し
    .setValue('平均時給').setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBackground(BLUE_L).setFontColor('#000000')
    .setBorder(true, true, true, true, null, null, GRAY, SOLID);
  sh.getRange(2, 5, 1, avgMid - 5).merge()   // E〜：変更前
    .setValue(templateMode ? '変更前' : `変更前　¥${avgOld.toLocaleString()}`).setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, true, true, true, null, null, GRAY, SOLID);
  sh.getRange(2, avgMid, 1, 1)               // 矢印
    .setValue('⇒').setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, true, true, true, null, null, GRAY, SOLID);
  sh.getRange(2, avgMid + 1, 1, GRID_COLS - avgMid).merge()   // 〜末尾：変更後
    .setValue(templateMode ? '変更後' : `変更後　¥${avgNew.toLocaleString()}`).setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, true, true, true, null, null, GRAY, SOLID);

  // ===== 行3：ヘッダー =====
  const HDRS = { no:'No.', name:'氏名', gender:'性別', age:'年齢', tenure:'経験年数',
                 oldCls:'旧クラスター', arrow1:'→', newCls:'新クラスター',
                 oldWage:'旧時給', arrow2:'→', newWage:'新時給', diff:'昇給額' };
  ORDER.forEach(k => {
    seg(3, k, HDRS[k])
      .setFontWeight('bold').setFontSize(9)
      .setBackground(BLUE_H).setFontColor('#FFFFFF')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
  });
  sh.getRange(3, 1, 1, GRID_COLS)
    .setBorder(true, true, true, true, true, true, '#FFFFFF', SOLID);

  // ===== データ（在籍スタッフ全員）=====
  rows.forEach((rw, i) => {
    const row  = DATA_TOP + i;
    const bg   = i % 2 === 0 ? '#FFFFFF' : BLUE_R;
    const diff = rw.newW - rw.oldW;
    const vals = {
      no: i + 1, name: rw.name, gender: rw.gender, age: rw.age, tenure: rw.tenure,
      oldCls: rw.oldCls, arrow1: '→', newCls: rw.newCls,
      oldWage: `¥${rw.oldW.toLocaleString()}`, arrow2: '→', newWage: `¥${rw.newW.toLocaleString()}`,
      diff: diff > 0 ? `+${diff}円` : diff < 0 ? `${diff}円` : '変更なし'
    };
    ORDER.forEach(k => {
      const align = (k === 'name') ? 'left'
                  : (k === 'oldWage' || k === 'newWage' || k === 'diff') ? 'right' : 'center';
      seg(row, k, vals[k]).setHorizontalAlignment(align);
    });
    sh.getRange(row, 1, 1, GRID_COLS)
      .setBackground(bg).setFontSize(9).setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true, LINE, SOLID);
  });

  // ===== フッター（事業部・店舗・店長）=====
  const info = getStoreInfo_();
  const footerItems = [
    ['事業部名', templateMode ? '' : info.division],
    ['店舗名',   templateMode ? '' : info.name],
    ['店長名',   templateMode ? '' : info.manager],
  ];
  footerItems.forEach((fd, i) => {
    const r = FOOT_TOP + i;
    // 見出し：A〜D列（1〜4）左寄せ
    sh.getRange(r, 1, 1, 4).merge().setValue(fd[0]).setFontWeight('bold')
      .setHorizontalAlignment('left').setVerticalAlignment('middle').setBackground(BLUE_L)
      .setBorder(true, true, true, true, null, null, GRAY, SOLID);
    // 名称：E〜M列（5〜13）左寄せ
    sh.getRange(r, 5, 1, 9).merge().setValue(fd[1]).setVerticalAlignment('middle')
      .setHorizontalAlignment('left')
      .setBorder(true, true, true, true, null, null, GRAY, SOLID);
    // N列以降（14〜24）は空白のまま（枠線なし）
  });

  // ===== 会社回覧用 押印欄（最下部・完全等幅）=====
  //  24升を役職数Nで割り切って各枠に同じ升数を割り当てる（既定6枠なら各4升で完全等幅）
  const stampTitles = getStampTitles_();
  const N = stampTitles.length;
  let startCol = 1;
  stampTitles.forEach((title, i) => {
    const endBoundary = Math.round((GRID_COLS * (i + 1)) / N);
    const span = endBoundary - Math.round((GRID_COLS * i) / N);
    // 役職ラベル
    sh.getRange(STAMP_LBL, startCol, 1, span).merge()
      .setValue(title).setFontWeight('bold').setFontSize(9)
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBackground(BLUE_L)
      .setBorder(true, true, true, true, null, null, GRAY, SOLID);
    // 押印スペース
    sh.getRange(STAMP_BOX, startCol, 1, span).merge()
      .setBorder(true, true, true, true, null, null, GRAY, SOLID);
    startCol += span;
  });

  ss.setActiveSheet(sh);
  const changedCount = rows.filter(r => r.newW !== r.oldW || String(r.newCls) !== String(r.oldCls)).length;
  return `在籍スタッフ ${rows.length}名を出力しました（うち変更あり ${changedCount}名）。\n承認後に「名簿に時給を反映」を実行してください。`;
}

// ------------------------------------------------------------
//  押印欄の役職名を取得（保存されていなければ初期値）
// ------------------------------------------------------------
function getStampTitles_() {
  const saved = PropertiesService.getDocumentProperties().getProperty('STAMP_TITLES');
  if (saved) {
    const arr = saved.split(',').map(s => s.trim()).filter(s => s);
    if (arr.length) return arr;
  }
  return STAMP_TITLES_DEFAULT;
}

// ------------------------------------------------------------
//  押印欄の役職名を変更（6枠それぞれの入力欄・変えたい枠だけ直せる）
// ------------------------------------------------------------
function setStampTitles() {
  const cur = getStampTitles_();
  const inputs = cur.map((t, i) =>
    `<div class="row"><label>${i + 1}枠目</label>` +
    `<input type="text" id="t${i}" value="${String(t).replace(/"/g, '&quot;')}"></div>`
  ).join('');

  const html = HtmlService.createHtmlOutput(
    `<style>
       body  { font-family: sans-serif; margin: 0; padding: 12px; }
       .hint { font-size: 12px; color: #666; margin-bottom: 10px; line-height: 1.5; }
       .row  { display: flex; align-items: center; margin-bottom: 8px; }
       .row label { width: 56px; font-size: 13px; color: #333; }
       .row input { flex: 1; font-size: 14px; padding: 6px 8px;
                    border: 1px solid #bbb; border-radius: 4px; }
       .btns  { text-align: right; margin-top: 12px; }
       button { font-size: 14px; padding: 7px 16px; border: none; border-radius: 4px; cursor: pointer; }
       .save   { background: #1a73e8; color: #fff; }
       .cancel { background: #eee; color: #333; margin-right: 8px; }
       #msg   { text-align: center; color: #1a73e8; display: none; margin-top: 10px; }
     </style>
     <div class="hint">変えたい枠だけ書き換えて「保存」を押してください。<br>
       空欄のままにした枠は今の役職名がそのまま残ります。</div>
     <div id="form">${inputs}</div>
     <div class="btns">
       <button class="cancel" onclick="google.script.host.close()">キャンセル</button>
       <button class="save" onclick="save()">保存</button>
     </div>
     <div id="msg">保存中…</div>
     <script>
       var N = ${cur.length};
       function save() {
         var arr = [];
         for (var i = 0; i < N; i++) arr.push(document.getElementById('t' + i).value);
         document.getElementById('msg').style.display = 'block';
         google.script.run
           .withSuccessHandler(function() { google.script.host.close(); })
           .withFailureHandler(function(err) {
             document.getElementById('msg').style.color = '#d33';
             document.getElementById('msg').textContent = 'エラー: ' + err.message;
           })
           .saveStampTitles(arr);
       }
     </script>`
  ).setWidth(340).setHeight(150 + cur.length * 42);

  SpreadsheetApp.getUi().showModalDialog(html, '🖊️ 押印欄の役職名を変更');
}

// ------------------------------------------------------------
//  役職名の保存画面から呼ばれる（空欄の枠は今の役職名を維持）
//  ※ google.script.run から呼ぶため末尾に _ を付けないこと
// ------------------------------------------------------------
function saveStampTitles(arr) {
  const cur = getStampTitles_();
  const titles = (arr || []).map((v, i) => {
    const t = String(v == null ? '' : v).trim();
    return t !== '' ? t : (cur[i] || '');   // 空欄はいまの役職名を維持
  }).filter(t => t !== '');
  if (titles.length === 0) throw new Error('役職名がすべて空です。');
  PropertiesService.getDocumentProperties().setProperty('STAMP_TITLES', titles.join(','));
  SpreadsheetApp.getActiveSpreadsheet().toast(
    titles.join('、') + '　→「📊 時給変更届を更新」で押印欄に反映されます。', '役職名を保存しました', 6);
}

// ------------------------------------------------------------
//  店舗情報（事業部名・店舗名・店長名）を取得
//  この店舗で保存されていればそれを、無ければ初期値（定数）を使う。
//  ※ 保存済みなら「空欄」もそのまま返す（渡し用でクリアした店舗を空のままにするため）
// ------------------------------------------------------------
function getStoreInfo_() {
  const saved = PropertiesService.getDocumentProperties().getProperty('STORE_INFO');
  if (saved != null) {
    try {
      const o = JSON.parse(saved);
      return { division: o.division || '', name: o.name || '', manager: o.manager || '' };
    } catch (e) { console.log('店舗情報の解析エラー: ' + e.message); }
  }
  return { division: STORE_DIVISION, name: STORE_NAME, manager: STORE_MANAGER };
}

// ------------------------------------------------------------
//  店舗情報を変更（事業部名・店舗名・店長名。変えたい欄だけ直せる）
// ------------------------------------------------------------
function setStoreInfo() {
  const cur = getStoreInfo_();
  const fields = [
    ['division', '事業部名', cur.division],
    ['name',     '店舗名',   cur.name],
    ['manager',  '店長名',   cur.manager],
  ];
  const inputs = fields.map((f, i) =>
    `<div class="row"><label>${f[1]}</label>` +
    `<input type="text" id="f${i}" value="${String(f[2]).replace(/"/g, '&quot;')}"></div>`
  ).join('');

  const html = HtmlService.createHtmlOutput(
    `<style>
       body  { font-family: sans-serif; margin: 0; padding: 12px; }
       .hint { font-size: 12px; color: #666; margin-bottom: 10px; line-height: 1.5; }
       .row  { display: flex; align-items: center; margin-bottom: 8px; }
       .row label { width: 72px; font-size: 13px; color: #333; }
       .row input { flex: 1; font-size: 14px; padding: 6px 8px;
                    border: 1px solid #bbb; border-radius: 4px; }
       .btns  { text-align: right; margin-top: 12px; }
       button { font-size: 14px; padding: 7px 16px; border: none; border-radius: 4px; cursor: pointer; }
       .save   { background: #1a73e8; color: #fff; }
       .cancel { background: #eee; color: #333; margin-right: 8px; }
       #msg   { text-align: center; color: #1a73e8; display: none; margin-top: 10px; }
     </style>
     <div class="hint">この店舗の情報を入力して「保存」を押してください。<br>
       時給変更届のフッター（事業部名・店舗名・店長名）に反映されます。</div>
     <div id="form">${inputs}</div>
     <div class="btns">
       <button class="cancel" onclick="google.script.host.close()">キャンセル</button>
       <button class="save" onclick="save()">保存</button>
     </div>
     <div id="msg">保存中…</div>
     <script>
       function save() {
         var o = {
           division: document.getElementById('f0').value,
           name:     document.getElementById('f1').value,
           manager:  document.getElementById('f2').value
         };
         document.getElementById('msg').style.display = 'block';
         google.script.run
           .withSuccessHandler(function() { google.script.host.close(); })
           .withFailureHandler(function(err) {
             document.getElementById('msg').style.color = '#d33';
             document.getElementById('msg').textContent = 'エラー: ' + err.message;
           })
           .saveStoreInfo(o);
       }
     </script>`
  ).setWidth(360).setHeight(280);

  SpreadsheetApp.getUi().showModalDialog(html, '🏢 店舗情報を変更');
}

// ------------------------------------------------------------
//  店舗情報の保存画面から呼ばれる（※ google.script.run 用・末尾 _ なし）
// ------------------------------------------------------------
function saveStoreInfo(o) {
  o = o || {};
  const info = {
    division: String(o.division == null ? '' : o.division).trim(),
    name:     String(o.name     == null ? '' : o.name).trim(),
    manager:  String(o.manager  == null ? '' : o.manager).trim()
  };
  PropertiesService.getDocumentProperties().setProperty('STORE_INFO', JSON.stringify(info));

  // ★STORE_INFO_APPLY★ 保存と同時に、時給変更届シートのフッターへ即反映する
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const applied = applyStoreInfoToSummary_(ss);
  const shown = [info.division, info.name, info.manager].filter(Boolean).join('　');
  ss.toast(
    shown + (applied
      ? '　→ 時給変更届のフッターに反映しました。'
      : '　→ 保存しました（時給変更届はまだ無いので、作成時に反映されます）。'),
    '店舗情報を保存しました', 6);
}

// 時給変更届シートのフッター（事業部名・店舗名・店長名）に、保存済みの店舗情報を書き込む。
// 見出しラベルの行を探して名称セル（E列＝結合セルの先頭）に入れるので、行位置に依存しない。反映できたら true。
function applyStoreInfoToSummary_(ss) {
  const sh = ss.getSheetByName(SUMMARY_SHEET);
  if (!sh) return false;
  const info = getStoreInfo_();
  const map  = { '事業部名': info.division, '店舗名': info.name, '店長名': info.manager };
  const last = sh.getLastRow();
  if (last < 1) return false;
  const colA = sh.getRange(1, 1, last, 1).getValues();  // A列（結合セルの見出しはここに入る）
  let updated = 0;
  for (let r = 1; r <= last; r++) {
    const lab = String(colA[r - 1][0]).trim();
    if (Object.prototype.hasOwnProperty.call(map, lab)) {
      sh.getRange(r, 5).setValue(map[lab]);  // 名称セル（E〜M列の結合＝先頭はE列）
      updated++;
    }
  }
  return updated > 0;
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
    '⚠️【重要】名簿反映の最終確認',
    '・この操作は元に戻せません\n' +
    '・上長の承認が取れていることを確認してください\n' +
    '・反映後はアルバイト名簿のクラスター・時給が書き換わります\n\n' +
    `${period}　承認済み ${approved.length}件を名簿に反映します。\n\n` +
    '本当に実行しますか？',
    ui.ButtonSet.YES_NO
  );
  if (res !== ui.Button.YES) return;

  const rosterSS    = getRosterSS_();
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

  ui.alert(`✅ ${count}件のクラスター・時給をアルバイト名簿に反映しました。`);
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

// クラスター名を順位数値に変換（indexが小さいほど上位）
function clusterRank_(name) {
  const idx = CLUSTER_TABLE.findIndex(c => c.name === name);
  return idx === -1 ? 999 : idx;
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

// ------------------------------------------------------------
//  PDF一括出力（全スタッフのチェックシート＋時給変更集計）
// ------------------------------------------------------------
function exportAllPDFs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // 評価履歴から期間を取得
  const histSh = ss.getSheetByName(HISTORY_SHEET);
  if (!histSh || histSh.getLastRow() < 2) { ui.alert('評価データがありません。'); return; }
  const histData = histSh.getRange(2, 1, histSh.getLastRow() - 1, 14).getValues();
  const availPeriods = [...new Set(histData.map(r => normalizePeriod_(r[1])).filter(Boolean))].sort().reverse();
  if (availPeriods.length === 0) { ui.alert('評価データがありません。'); return; }

  const periodBtns  = availPeriods.map(p => `<button class="pd" onclick="pick('${p}')">${p}</button>`).join('');
  const periodsJson = JSON.stringify(availPeriods);

  // ★PDF_SELECT★ 期間 → 出力対象を選ぶ → 出力形式を選ぶ → 出力、の一画面ダイアログ
  const html = HtmlService.createHtmlOutput(
    `<style>
       body { font-family: sans-serif; margin:0; padding:14px; font-size:13px; color:#333; }
       .h   { font-size:13px; margin-bottom:10px; }
       .pd  { display:block; width:100%; font-size:15px; padding:11px; margin-bottom:8px;
              border:1px solid #1a73e8; background:#fff; color:#1a73e8; border-radius:6px; cursor:pointer; }
       .pd:hover { background:#e8f0fe; }
       #stafflist { max-height:230px; overflow-y:auto; border:1px solid #ddd; border-radius:6px; padding:4px 8px; margin:8px 0; }
       .item { display:flex; align-items:center; padding:5px 2px; border-bottom:1px solid #f0f0f0; }
       .item input { margin-right:8px; transform:scale(1.2); }
       .item em { font-style:normal; font-size:11px; margin-left:6px; font-weight:bold; }
       .item em.kaku { color:#c5221f; }
       .item em.kyu  { color:#188038; }
       .tools button { font-size:12px; padding:4px 8px; margin-right:6px; border:1px solid #999;
                       background:#fff; border-radius:5px; cursor:pointer; }
       .fmt { margin:10px 0; }
       .fmt label { display:block; padding:6px 0; cursor:pointer; }
       .go  { width:100%; font-size:15px; padding:11px; margin-top:6px; border:none;
              background:#1a73e8; color:#fff; border-radius:6px; cursor:pointer; }
       #msg { display:none; text-align:center; color:#1a73e8; margin-top:6px; white-space:pre-wrap; }
       .close { display:none; width:100%; font-size:14px; padding:9px; margin-top:12px;
                border:none; background:#1a73e8; color:#fff; border-radius:6px; cursor:pointer; }
     </style>
     <div id="phase1">
       <div class="h">PDFを出力する評価年月を選んでください：</div>
       ${periodBtns}
     </div>
     <div id="loading" style="display:none; text-align:center; color:#1a73e8;">読み込み中…</div>
     <div id="phase2" style="display:none;">
       <div class="h" id="phdr"></div>
       <div class="tools">
         <button onclick="selAll(true)">全選択</button>
         <button onclick="selAll(false)">全解除</button>
       </div>
       <div id="stafflist"></div>
       <div class="fmt">
         <label><input type="radio" name="mode" value="combined" checked> 1つのPDFにまとめる（1人1ページ）</label>
         <label><input type="radio" name="mode" value="individual"> 個別に作る（1人1ファイル）</label>
       </div>
       <button class="go" onclick="doExport()">出力する</button>
     </div>
     <div id="msg"></div>
     <button id="closeBtn" class="close" onclick="google.script.host.close()">閉じる</button>
     <script>
       var PERIODS = ${periodsJson};
       var PERIOD  = null;
       function show(id,v){ var el=document.getElementById(id); if(el) el.style.display = v ? '' : 'none'; }
       function pick(p){
         PERIOD = p;
         show('phase1', false);
         show('loading', true);
         google.script.run.withSuccessHandler(renderStaff).withFailureHandler(showErr).getExportCandidates(p);
       }
       function renderStaff(list){
         show('loading', false);
         var box = document.getElementById('stafflist');
         box.innerHTML = '';
         if(!list || !list.length){ box.textContent = '在籍スタッフが取得できませんでした。'; }
         (list || []).forEach(function(s){
           var lab = document.createElement('label'); lab.className = 'item';
           var cb = document.createElement('input'); cb.type = 'checkbox'; cb.value = s.name; cb.checked = !!s.promoted;
           var sp = document.createElement('span'); sp.textContent = s.name;
           if(s.label){ var em = document.createElement('em'); em.textContent = s.label;
             em.className = (s.label === '昇格') ? 'kaku' : 'kyu'; sp.appendChild(em); }
           lab.appendChild(cb); lab.appendChild(sp); box.appendChild(lab);
         });
         document.getElementById('phdr').textContent = PERIOD + ' の出力対象を選んでください（昇格者は選択済み）';
         show('phase2', true);
       }
       function selAll(v){ var cbs = document.querySelectorAll('#stafflist input'); for(var i=0;i<cbs.length;i++) cbs[i].checked = v; }
       function doExport(){
         var cbs = document.querySelectorAll('#stafflist input'); var names = [];
         for(var i=0;i<cbs.length;i++) if(cbs[i].checked) names.push(cbs[i].value);
         if(!names.length){ alert('1人以上選んでください。'); return; }
         var mode = document.querySelector('input[name=mode]:checked').value;
         show('phase2', false);
         var m = document.getElementById('msg'); show('msg', true);
         m.textContent = 'PDFを作成しています…\\nしばらくお待ちください（数十秒かかることがあります）。';
         google.script.run
           .withSuccessHandler(function(res){
             if(res && res.clean){ google.script.host.close(); return; }   // 問題なければ自動で閉じる
             m.textContent = (res && res.message) ? res.message : String(res);
             show('closeBtn', true);
           })
           .withFailureHandler(showErr)
           .runExportSelected(PERIOD, names, mode);
       }
       function showErr(err){
         show('loading', false); show('phase2', false);
         var m = document.getElementById('msg'); show('msg', true);
         m.style.color = '#d33'; m.textContent = 'エラー: ' + err.message; show('closeBtn', true);
       }
       window.onload = function(){ if(PERIODS.length === 1){ pick(PERIODS[0]); } };
     </script>`
  ).setWidth(370).setHeight(580);

  ui.showModalDialog(html, '📄 PDF出力');
}

// 選択画面用：その期間の在籍スタッフ一覧＋「昇格したか」を返す（google.script.run 用・末尾 _ なし）
function getExportCandidates(period) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 在籍スタッフ（退社日が空・区分が社員以外）
  let staff = [];
  try {
    const masterSh = getRosterSS_().getSheetByName(ROSTER_SHEET);
    const md = masterSh.getRange(2, 1, masterSh.getLastRow() - 1, 30).getValues();
    staff = md.filter(r => r[MC_NAME - 1] !== '' && r[MC_NAME - 1] !== null
                        && r[MC_LEAVE - 1] === '' && String(r[MC_KUBUN - 1]).trim() !== '社員')
              .map(r => String(r[MC_NAME - 1]).trim());
  } catch (e) { console.log('候補取得エラー: ' + e.message); }

  // 昇格（クラスターが上がった）／昇給（時給が上がった）を判定（どちらもデフォルトでチェック）
  const labelMap = {};
  const histSh = ss.getSheetByName(HISTORY_SHEET);
  if (histSh && histSh.getLastRow() >= 2) {
    const hist = histSh.getRange(2, 1, histSh.getLastRow() - 1, 14).getValues();
    const targets = latestPerName_(hist.filter(r => normalizePeriod_(r[1]) === period), 0, 13);
    targets.forEach(t => {
      const nm        = String(t[0]).trim();
      const clusterUp = clusterRank_(t[9]) < clusterRank_(t[8]);
      const wageUp    = (Number(t[11]) || 0) > (Number(t[10]) || 0);
      if (clusterUp)   labelMap[nm] = '昇格';   // クラスターが上がった
      else if (wageUp) labelMap[nm] = '昇給';   // クラスターは同じで時給だけ上がった
    });
  }
  return staff.map(n => ({ name: n, promoted: !!labelMap[n], label: labelMap[n] || '' }));
}

// 選択された人だけをPDF出力する（google.script.run 用・末尾 _ なし）。結果メッセージを返す。
//   mode='combined' … チェックシートを1人1ページで1つのPDFに / 'individual' … 1人1ファイル
//   時給変更届は選んだ人だけの一覧を1ファイルに。保存先は日時入りの新フォルダ（上書きしない）。
function runExportSelected(period, names, mode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  names = (names || []).map(n => String(n).trim()).filter(Boolean);
  if (!names.length) return '出力する人が選ばれていません。';

  // 評価履歴（最新1件／人）
  const histSh   = ss.getSheetByName(HISTORY_SHEET);
  const histData = (histSh && histSh.getLastRow() >= 2)
    ? histSh.getRange(2, 1, histSh.getLastRow() - 1, 14).getValues() : [];
  const targets  = latestPerName_(histData.filter(r => normalizePeriod_(r[1]) === period), 0, 13);
  const targetMap = {}; targets.forEach(t => { targetMap[String(t[0]).trim()] = t; });

  // 評価詳細履歴
  const detailSh = ss.getSheetByName(DETAIL_SHEET);
  const detailData = (detailSh && detailSh.getLastRow() >= 2)
    ? detailSh.getRange(2, 1, detailSh.getLastRow() - 1, 110).getValues() : [];

  // スタッフマスタ
  let masterData = [];
  try {
    const m = getRosterSS_().getSheetByName(ROSTER_SHEET);
    masterData = m.getRange(2, 1, m.getLastRow() - 1, 30).getValues();
  } catch (e) { console.log('スタッフマスタエラー: ' + e.message); }

  // チェックシート対象＝選んだ人のうち、この期間の評価詳細がある人
  const csNames = [], noData = [];
  names.forEach(n => {
    const has = detailData.some(r => String(r[0]).trim() === n && normalizePeriod_(r[1]) === period);
    if (has) csNames.push(n); else noData.push(n);
  });

  // 保存先：日時入りの新フォルダ（毎回別フォルダ＝上書きしない）
  const parentFolder = DriveApp.getFileById(ss.getId()).getParents().next();
  const tz    = ss.getSpreadsheetTimeZone() || 'Asia/Tokyo';
  const stamp = Utilities.formatDate(new Date(), tz, 'yyyyMMdd-HHmm');
  const folderName = `PA評価PDF_${period}_${stamp}`;
  const pdfFolder  = parentFolder.createFolder(folderName);

  const evalSh  = ss.getSheetByName(EVAL_SHEET);
  const evalGid = evalSh.getSheetId();
  const ssId    = ss.getId();
  const token   = ScriptApp.getOAuthToken();
  const pdfBase     = 'format=pdf&size=A4&fitw=true&sheetnames=false&printtitle=false&pagenumbers=false&gridlines=false&fzr=false';
  const pdfPortrait = pdfBase + '&portrait=true';
  const pdfChecksheet = pdfPortrait + '&top_margin=0.2&bottom_margin=0.2&left_margin=0.25&right_margin=0.25';

  const prevName = evalSh.getRange(C_NAME).getValue();  // 最後に元の人へ戻すため
  let usedFallback = false;
  const errorNames = [];

  // チェックシートに1人分のデータを流し込む
  function fillChecksheet_(name) {
    const dts = detailData.filter(r => String(r[0]).trim() === name && normalizePeriod_(r[1]) === period);
    const detail = dts.length ? dts[dts.length - 1] : null;
    const t     = targetMap[name];
    const staff = masterData.find(r => String(r[MC_NAME - 1]).trim() === name);
    evalSh.getRange(C_PERIOD).setValue(period);
    evalSh.getRange(C_NAME).setValue(name);
    evalSh.getRange(C_GENDER).setValue(staff ? staff[MC_GENDER - 1] : '');
    evalSh.getRange(C_AGE).setValue(staff    ? staff[MC_AGE    - 1] : '');
    evalSh.getRange(C_TENURE).setValue(staff ? staff[MC_TENURE - 1] : '');
    evalSh.getRange(C_OLD_CLS).setValue(t ? t[8]  : (staff ? staff[MC_CLUSTER - 1] : ''));
    evalSh.getRange(C_NEW_CLS).setValue(t ? t[9]  : '');
    evalSh.getRange(C_OLD_WAGE).setValue(t ? t[10] : (staff ? staff[MC_WAGE - 1] : ''));
    evalSh.getRange(C_NEW_WAGE).setValue(t ? t[11] : '');
    evalSh.getRange(C_PREV_INFO).setValue('');
    if (detail) {
      evalSh.getRange(R_LEFT).setValues( detail.slice(3,   63).map(v => [Number(v) === 1 ? 1 : '']));
      evalSh.getRange(R_RIGHT).setValues(detail.slice(63, 103).map(v => [Number(v) === 1 ? 1 : '']));
      evalSh.getRange(R_BONUS).setValues(detail.slice(103, 110).map(v => [Number(v) === 1 ? 1 : '']));
    }
  }

  let csCount = 0;

  if (csNames.length && mode === 'combined') {
    // 一時スプレッドシートに「1人1シート」で作り、丸ごと1PDFに（各シートが別ページになる）
    const tempSS = SpreadsheetApp.create('_pa_cs_tmp');
    try {
      csNames.forEach((name, idx) => {
        fillChecksheet_(name);
        SpreadsheetApp.flush();
        const copied = evalSh.copyTo(tempSS);
        copied.setName(safeSheetName_(name, idx));
        // 値のある範囲より下・右の空白を削って、余計な空白ページを防ぐ
        const dr = copied.getDataRange();
        const lastR = dr.getLastRow(), lastC = dr.getLastColumn();
        if (copied.getMaxRows()    > lastR) copied.deleteRows(lastR + 1, copied.getMaxRows() - lastR);
        if (copied.getMaxColumns() > lastC) copied.deleteColumns(lastC + 1, copied.getMaxColumns() - lastC);
        csCount++;
      });
      // 最初から在る空シート（コピーより前）を削除し、コピーだけ残す
      const sheets = tempSS.getSheets();
      for (let i = 0; i < sheets.length - csCount; i++) tempSS.deleteSheet(sheets[i]);
      SpreadsheetApp.flush();

      const fname = `チェックシート_${period}.pdf`;
      let blob;
      try {
        const url = `https://docs.google.com/spreadsheets/d/${tempSS.getId()}/export?${pdfChecksheet}`;
        const r = UrlFetchApp.fetch(url, { headers: { Authorization: `Bearer ${token}` }, muteHttpExceptions: true });
        if (r.getResponseCode() !== 200) throw new Error('code ' + r.getResponseCode());
        blob = r.getBlob().setName(fname);
      } catch (fe) {
        usedFallback = true;
        blob = DriveApp.getFileById(tempSS.getId()).getAs('application/pdf').setName(fname);
      }
      pdfFolder.createFile(blob);
    } catch (e) {
      console.log('まとめPDF作成エラー: ' + e.message);
      errorNames.push('（まとめPDF）' + e.message);
    } finally {
      try { DriveApp.getFileById(tempSS.getId()).setTrashed(true); } catch (e) {}
    }

  } else if (csNames.length) {
    // 個別：1人1ファイル
    csNames.forEach(name => {
      try {
        fillChecksheet_(name);
        SpreadsheetApp.flush();
        const fname = `チェックシート_${name}_${period}.pdf`;
        let blob;
        try {
          const range = evalSh.getDataRange().getA1Notation();
          const url = `https://docs.google.com/spreadsheets/d/${ssId}/export?${pdfChecksheet}&gid=${evalGid}&range=${range}`;
          const r = UrlFetchApp.fetch(url, { headers: { Authorization: `Bearer ${token}` }, muteHttpExceptions: true });
          if (r.getResponseCode() !== 200) throw new Error('code ' + r.getResponseCode());
          blob = r.getBlob().setName(fname);
        } catch (fe) {
          usedFallback = true;
          blob = sheetToPdfBlob_(evalSh, fname);
        }
        pdfFolder.createFile(blob);
        csCount++;
      } catch (e) { errorNames.push(name); console.log(`${name}のPDF作成エラー: ${e.message}`); }
    });
  }

  // 時給変更届（選んだ人だけの一覧）を作ってPDF化
  let sumOk = false, sumErr = '';
  try {
    const built = buildSummaryRows_(ss, period, names);
    if (built.rows.length) {
      renderSummarySheet_(ss, period, built.rows, false);
      SpreadsheetApp.flush();
      const sumSh = ss.getSheetByName(SUMMARY_SHEET);
      const fname = `時給変更届_${period}.pdf`;
      let blob;
      try {
        const url = `https://docs.google.com/spreadsheets/d/${ssId}/export?${pdfPortrait}&gid=${sumSh.getSheetId()}`;
        const r = UrlFetchApp.fetch(url, { headers: { Authorization: `Bearer ${token}` }, muteHttpExceptions: true });
        if (r.getResponseCode() !== 200) throw new Error('code ' + r.getResponseCode());
        blob = r.getBlob().setName(fname);
      } catch (fe) {
        usedFallback = true;
        blob = sheetToPdfBlob_(sumSh, fname);
      }
      pdfFolder.createFile(blob);
      sumOk = true;
    } else {
      sumErr = '対象者の行がありませんでした';
    }
  } catch (e) { sumErr = e.message; console.log('時給変更届PDFエラー: ' + e.message); }

  // 画面を元に戻す：時給変更届は全員表示に、チェックシートは元の人に
  try { runUpdateSummary_(period); } catch (e) {}
  try { if (prevName) applyPickedName(prevName); } catch (e) {}
  SpreadsheetApp.flush();

  // 結果メッセージ
  const csText = (mode === 'combined')
    ? (csCount > 0 ? `1ファイル（${csCount}名分）` : 'なし')
    : `${csCount}ファイル`;
  let msg = `✅ PDF出力完了（${period}）\n\n・チェックシート：${csText}\n・時給変更届：${sumOk ? '1ファイル' : '未作成'}\n\n保存先フォルダ：\n「${folderName}」`;
  if (noData.length)        msg += `\n\n※ 評価データが無いためチェックシートを作れなかった人：\n${noData.join('、')}`;
  if (errorNames.length)    msg += `\n\n⚠️ 出力できなかった項目：\n${errorNames.join('、')}`;
  if (!sumOk && sumErr)     msg += `\n\n⚠️ 時給変更届が作れませんでした：\n${sumErr}`;
  if (usedFallback)         msg += `\n\n※ 簡易レイアウトで作成した書類があります。よりきれいなA4版にするには一度だけ権限承認が必要です（拡張機能→Apps Script→authorizeNow を実行）。`;

  // 完了トースト（画面右下に数秒だけ表示）。警告が無ければダイアログは自動で閉じる。
  try { ss.toast('保存先フォルダ：' + folderName, '✅ PDF出力完了', 8); } catch (e) {}
  const clean = (noData.length === 0 && errorNames.length === 0 && sumOk && !usedFallback);
  return { message: msg, clean: clean };
}

// 一時スプレッドシートのシート名に使えるよう整える（重複回避に番号を付ける）
function safeSheetName_(name, idx) {
  const s = (idx + 1) + '_' + String(name).replace(/[\[\]\*\?\/\\:]/g, ' ');
  return s.substring(0, 90);
}

// 実際のPDF生成。結果メッセージ（文字列）を返す。※ getUi は使わない（選択画面から呼ばれるため）
function runPdfExport_(period) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const histSh = ss.getSheetByName(HISTORY_SHEET);
  if (!histSh || histSh.getLastRow() < 2) return '評価データがありません。';
  const histData = histSh.getRange(2, 1, histSh.getLastRow() - 1, 14).getValues();

  // 同一年月・同一人物は最新の1件だけ（重複したチェックシートを二重出力しない）
  const targets = latestPerName_(histData.filter(r => normalizePeriod_(r[1]) === period), 0, 13);
  if (targets.length === 0) return `${period}の評価データが見つかりません。`;

  // 昇給した or 昇格した（クラスターが上がった）スタッフを対象（チェックシートPDF対象）
  const pdfTargets = targets.filter(t =>
    (Number(t[11]) || 0) > (Number(t[10]) || 0) || clusterRank_(t[9]) < clusterRank_(t[8]));

  // 評価詳細履歴を取得
  const detailSh = ss.getSheetByName(DETAIL_SHEET);
  let detailData = [];
  if (detailSh && detailSh.getLastRow() >= 2) {
    detailData = detailSh.getRange(2, 1, detailSh.getLastRow() - 1, 110).getValues();
  }

  // スタッフマスタを取得
  let masterData = [];
  try {
    const rosterSS = getRosterSS_();
    const masterSh = rosterSS.getSheetByName(ROSTER_SHEET);
    masterData = masterSh.getRange(2, 1, masterSh.getLastRow() - 1, 30).getValues();
  } catch(e) { console.log('スタッフマスタエラー: ' + e.message); }

  // 保存先フォルダを取得・作成
  const parentFolder = DriveApp.getFileById(ss.getId()).getParents().next();
  const folderName   = `PA評価PDF_${period}`;
  const folderIter   = parentFolder.getFoldersByName(folderName);
  const pdfFolder    = folderIter.hasNext() ? folderIter.next() : parentFolder.createFolder(folderName);

  const evalSh  = ss.getSheetByName(EVAL_SHEET);
  const evalGid = evalSh.getSheetId();
  const ssId    = ss.getId();
  const token   = ScriptApp.getOAuthToken();
  const pdfBase     = 'format=pdf&size=A4&fitw=true&sheetnames=false&printtitle=false&pagenumbers=false&gridlines=false&fzr=false';
  const pdfPortrait = pdfBase + '&portrait=true';  // 時給変更届（縦向き）
  // チェック表は縦に長いので、余白を詰めて1枚に収める（単位=インチ）
  const pdfChecksheet = pdfPortrait + '&top_margin=0.2&bottom_margin=0.2&left_margin=0.25&right_margin=0.25';

  let successCount = 0;
  let errorNames   = [];
  let usedFallback = false;

  // クラスターが上がったスタッフのみチェックシートをPDF化
  pdfTargets.forEach(t => {
    const name = t[0];
    try {
      // 評価詳細履歴から該当スタッフのチェックデータを取得
      const details = detailData.filter(r => r[0] === name && normalizePeriod_(r[1]) === period);
      const detail  = details.length > 0 ? details[details.length - 1] : null;

      // スタッフマスタから個人情報を取得
      const staff = masterData.find(r => r[MC_NAME - 1] === name);

      // チェックシートにデータを書き込む
      evalSh.getRange(C_PERIOD).setValue(period);
      evalSh.getRange(C_NAME).setValue(name);
      evalSh.getRange(C_GENDER).setValue(staff ? staff[MC_GENDER  - 1] : '');
      evalSh.getRange(C_AGE).setValue(staff    ? staff[MC_AGE     - 1] : '');
      evalSh.getRange(C_TENURE).setValue(staff ? staff[MC_TENURE  - 1] : '');
      evalSh.getRange(C_OLD_CLS).setValue(t[8]);
      evalSh.getRange(C_NEW_CLS).setValue(t[9]);
      evalSh.getRange(C_OLD_WAGE).setValue(t[10]);
      evalSh.getRange(C_NEW_WAGE).setValue(t[11]);
      evalSh.getRange(C_PREV_INFO).setValue('');

      if (detail) {
        evalSh.getRange(R_LEFT).setValues( detail.slice(3,   63).map(v => [Number(v) === 1 ? 1 : '']));
        evalSh.getRange(R_RIGHT).setValues(detail.slice(63,  103).map(v => [Number(v) === 1 ? 1 : '']));
        evalSh.getRange(R_BONUS).setValues(detail.slice(103, 110).map(v => [Number(v) === 1 ? 1 : '']));
      }

      SpreadsheetApp.flush();

      const fname = `チェックシート_${name}_${period}.pdf`;
      let blob;
      try {
        // 実際に値が入っている範囲だけを出力（下の空の書式行による白紙ページを防ぐ）
        const range = evalSh.getDataRange().getA1Notation();
        const url = `https://docs.google.com/spreadsheets/d/${ssId}/export?${pdfChecksheet}&gid=${evalGid}&range=${range}`;
        const r = UrlFetchApp.fetch(url, { headers: { Authorization: `Bearer ${token}` }, muteHttpExceptions: true });
        if (r.getResponseCode() !== 200) throw new Error('code ' + r.getResponseCode());
        blob = r.getBlob().setName(fname);
      } catch (fe) {
        usedFallback = true;                       // 権限が無い等→Driveで作成
        blob = sheetToPdfBlob_(evalSh, fname);
      }
      pdfFolder.createFile(blob);
      successCount++;
    } catch(e) {
      errorNames.push(name);
      console.log(`${name}のPDF作成エラー: ${e.message}`);
    }
  });

  // 時給変更届を最新レイアウトで作り直してからPDF化（「更新し忘れ」で古いまま出るのを防ぐ）
  let sumUpdateErr = '';
  try { runUpdateSummary_(period); } catch (e) { sumUpdateErr = e.message; console.log('時給変更届の再作成エラー: ' + e.message); }
  SpreadsheetApp.flush();   // ★書き込みを確定してからPDF化（これが無いと古い内容が出力される）

  // 時給変更届シートをPDF化
  const sumSh = ss.getSheetByName(SUMMARY_SHEET);
  let sumOk = false, sumErr = '';
  if (sumSh) {
    try {
      const fname = `時給変更届_${period}.pdf`;
      let blob;
      try {
        const url = `https://docs.google.com/spreadsheets/d/${ssId}/export?${pdfPortrait}&gid=${sumSh.getSheetId()}`;
        const r = UrlFetchApp.fetch(url, { headers: { Authorization: `Bearer ${token}` }, muteHttpExceptions: true });
        if (r.getResponseCode() !== 200) throw new Error('code ' + r.getResponseCode());
        blob = r.getBlob().setName(fname);
      } catch (fe) {
        usedFallback = true;                       // 権限が無い等→Driveで作成
        blob = sheetToPdfBlob_(sumSh, fname);
      }
      pdfFolder.createFile(blob);
      sumOk = true;
    } catch(e) { sumErr = e.message; console.log('時給変更届PDF作成エラー: ' + e.message); }
  } else {
    sumErr = '「時給変更届」シートが見つかりません。先に「時給変更届を更新」を実行してください';
  }

  let msg = `✅ PDF出力完了（${period}）\n\n・チェックシート：${successCount}名分\n・時給変更届：${sumOk ? '1件' : '未作成'}\n\n保存先：「${folderName}」フォルダ`;
  if (usedFallback) msg += `\n\n※ 簡易レイアウトで作成しました。よりきれいなA4版にするには一度だけ権限承認が必要です（拡張機能→Apps Script→authorizeNow を実行）。`;
  if (errorNames.length > 0) msg += `\n\n⚠️ 出力できなかったスタッフ：\n${errorNames.join('、')}`;
  if (!sumOk && sumErr)     msg += `\n\n⚠️ 時給変更届が作れませんでした：\n${sumErr}`;
  if (sumUpdateErr)         msg += `\n\n⚠️ 時給変更届の再作成でエラーが出たため、古い内容のままかもしれません：\n${sumUpdateErr}`;
  return msg;
}

// ------------------------------------------------------------
//  他店舗渡し用：個人情報をクリア
// ------------------------------------------------------------
function clearForTransfer() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const res = ui.alert(
    '⚠️ 個人情報クリア確認',
    '以下のデータを全て削除します。元に戻せません。\n\n' +
    '・チェックシートの氏名・チェック欄・個人情報\n' +
    '・時給変更届シートの氏名・個人情報・事業部名等\n' +
    '・評価履歴・評価詳細履歴の全データ\n\n' +
    '実行しますか？',
    ui.ButtonSet.YES_NO
  );
  if (res !== ui.Button.YES) return;

  // ── チェックシートをクリア ──
  const evalSh = ss.getSheetByName(EVAL_SHEET);
  if (evalSh) {
    evalSh.getRange(C_NAME).clearContent();
    evalSh.getRange(C_NAME).clearDataValidations();  // 氏名プルダウンの名前リストも削除（名前は個人情報）
    evalSh.getRange(C_GENDER).clearContent();
    evalSh.getRange(C_AGE).clearContent();
    evalSh.getRange(C_TENURE).clearContent();
    evalSh.getRange(C_OLD_CLS).clearContent();
    evalSh.getRange(C_NEW_CLS).clearContent();
    evalSh.getRange(C_OLD_WAGE).clearContent();
    evalSh.getRange(C_NEW_WAGE).clearContent();
    evalSh.getRange(C_PREV_INFO).clearContent();
    evalSh.getRange(R_LEFT).clearContent();
    evalSh.getRange(R_RIGHT).clearContent();
    evalSh.getRange(R_BONUS).clearContent();
  }

  // ── 店舗情報を空にする（コピーは元店舗の情報を引き継ぐため、明示的に空へ）──
  //  渡した先で「🏢 店舗情報を変更」を実行して自店の情報を入れてもらう
  PropertiesService.getDocumentProperties().setProperty(
    'STORE_INFO', JSON.stringify({ division: '', name: '', manager: '' }));

  // ── 時給変更届シートを「空フォーム」に作り直す ──
  //  個人データは消しつつ、見出し（事業部名/店舗名/店長名）と押印欄（会長〜次長）は残す。
  //  店舗情報の値は空にして、渡した先で記入・更新できるようにする。
  try {
    buildBlankSummaryForm_(ss);
  } catch (e) {
    // 万一描画に失敗しても最低限データは消す
    const sumSh = ss.getSheetByName(SUMMARY_SHEET);
    if (sumSh) {
      const lastR = sumSh.getLastRow();
      const lastC = Math.max(sumSh.getLastColumn(), 24);
      if (lastR >= 4) sumSh.getRange(4, 1, lastR - 3, lastC).clearContent();
    }
  }

  // ── 評価履歴をクリア（ヘッダーだけ残す）──
  const histSh = ss.getSheetByName(HISTORY_SHEET);
  if (histSh && histSh.getLastRow() >= 2) {
    histSh.getRange(2, 1, histSh.getLastRow() - 1, histSh.getLastColumn()).clearContent();
  }

  // ── 評価詳細履歴をクリア（ヘッダーだけ残す）──
  const detailSh = ss.getSheetByName(DETAIL_SHEET);
  if (detailSh && detailSh.getLastRow() >= 2) {
    detailSh.getRange(2, 1, detailSh.getLastRow() - 1, detailSh.getLastColumn()).clearContent();
  }

  ui.alert('✅ 個人情報をクリアしました。\n他店舗へ渡す準備ができました。');
}

// ============================================================
//  【ClaudeAI読み取り用】ウェブAPIエンドポイント
//  ※ GASエディタで「デプロイ」→「新しいデプロイ」→
//    「ウェブアプリ」→「アクセス：全員」でデプロイして使用
// ============================================================

function doGet(e) {
  const params   = e ? (e.parameter || {}) : {};
  const sheetName = params.sheet    || EVAL_SHEET;
  const mode      = params.mode     || 'both';  // values / formulas / both
  const rows      = parseInt(params.rows     || '50', 10);
  const startRow  = parseInt(params.startRow || '1',  10);

  try {
    const ss    = SpreadsheetApp.openById(EVAL_SS_ID);
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      const names = ss.getSheets().map(s => s.getName());
      return jsonResponse_({ error: `シート "${sheetName}" が見つかりません`, availableSheets: names });
    }

    const lastCol = sheet.getLastColumn();
    const lastRow = sheet.getLastRow();
    const endRow  = Math.min(startRow + rows - 1, lastRow);
    const numRows = Math.max(endRow - startRow + 1, 0);

    if (numRows === 0) {
      return jsonResponse_({ sheetName, startRow, rows: 0, cols: lastCol, values: [], formulas: [] });
    }

    const range  = sheet.getRange(startRow, 1, numRows, lastCol);
    const result = { sheetName, startRow, rows: numRows, cols: lastCol };

    if (mode === 'values' || mode === 'both') {
      result.values   = range.getValues();
    }
    if (mode === 'formulas' || mode === 'both') {
      result.formulas = range.getFormulas();
    }

    return jsonResponse_(result);

  } catch (err) {
    return jsonResponse_({ error: err.message });
  }
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
