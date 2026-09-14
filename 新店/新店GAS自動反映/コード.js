/**
 * 新店 レシピ台帳 作成スクリプト
 *
 * 【新しく作る場合】
 *  1. 新しい空のスプレッドシートを作る
 *  2. 拡張機能 → Apps Script
 *  3. このコードを全部貼り付けて保存
 *  4. 関数「新店レシピ台帳を作成」を実行  ← 8シート全部できる
 *
 * 【既にある台帳に試作シートだけ足す場合】
 *  関数「試作シートを追加」を実行
 *
 * 【ふだん使うもの】
 *  メニュー「🧪 試作」から
 *   ・次の回の試作をつくる      … 前回の分量をコピーして次の回を始める
 *   ・採用した試作をレシピへ反映 … 決まった配合を本番レシピに移す
 */

var SUPPLY_SHEET = '仕入先マスタ';
var BASE_SHEETS  = ['食材マスタ', '仕込み品マスタ', '商品マスタ', 'レシピ明細', '調理手順'];
var TRIAL_SHEETS = ['試作記録', '試作レシピ明細', '試作比較'];
var SHEET_NAMES  = [SUPPLY_SHEET].concat(BASE_SHEETS).concat(TRIAL_SHEETS);

// 発注可能曜日の書き方。仕入先マスタで使う
var OPT_曜日 = ['月', '火', '水', '木', '金', '土', '日'];

var MAX_TRIALS = 12;   // 試作比較に横並びで出せる回数

// 総合評価＝甘味・塩味・酸味・旨味(F〜I)が 10(ちょうど良い) にどれだけ近いか。
// ずれが0なら20点、平均10ずれたら1点。苦味・辛味(J,K)は狙いが料理ごとに違うので入れない。
var OVERALL_FORMULA =
  '=ARRAYFORMULA(IF(A2:A="","",IF(N(F2:F)*N(G2:G)*N(H2:H)*N(I2:I)=0,"",' +
  'ROUND(20-(ABS(F2:F-10)+ABS(G2:G-10)+ABS(H2:H-10)+ABS(I2:I-10))/4*1.9,1))))';

var C_TASTE_CORE = '#9C3F22';   // 評価に入る味の枠（濃い）
var C_TASTE_REF  = '#B8A98F';   // 参考の味の枠（薄い・破線）

// 入力候補（プルダウン）。業態に合わせて後から自由に変更できます
var OPT_分類     = ['肉', '魚介', '野菜', '米・粉', '調味料', '乳製品', 'ドリンク', '資材', 'その他'];
var OPT_在庫単位 = ['g', 'ml', '個', '枚', '本'];
var OPT_保存区分 = ['冷凍', '冷蔵', '常温'];
var OPT_発注単位 = ['ケース', '箱', '袋', '本', '丁', '個'];

var C_HEADER_BG = '#F3EADB';
var C_AUTO_BG   = '#FBF6EE';   // 自動計算列は薄い色にして「触らない列」と分かるようにする
var LAST_ROW    = 1000;        // プルダウン・書式を適用する行数


function 新店レシピ台帳を作成() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var existing = SHEET_NAMES.filter(function (n) { return ss.getSheetByName(n) !== null; });
  if (existing.length > 0) {
    SpreadsheetApp.getUi().alert(
      '中止しました。\n\n既に次のシートが存在します:\n  ' + existing.join('\n  ') +
      '\n\n空のスプレッドシートで実行してください。'
    );
    return;
  }

  build仕入先マスタ(ss);
  build食材マスタ(ss);
  build仕込み品マスタ(ss);
  build商品マスタ(ss);
  buildレシピ明細(ss);
  build調理手順(ss);
  build試作記録(ss);
  build試作レシピ明細(ss);
  build試作比較(ss);

  // シートの並び順を整える
  SHEET_NAMES.forEach(function (name, i) {
    ss.setActiveSheet(ss.getSheetByName(name));
    ss.moveActiveSheet(i + 1);
  });
  ss.setActiveSheet(ss.getSheetByName('商品マスタ'));

  var name = ss.getName();
  if (name.indexOf('無題') === 0 || name.indexOf('Untitled') === 0) {
    ss.rename('新店 レシピ台帳');
  }

  SpreadsheetApp.getUi().alert(
    '作成しました。\n\n' +
    '8つのシートと、確認用のサンプルが入っています。\n' +
    '商品マスタ … 原価 ¥290 / 原価率 29.6% が自動計算されています。\n' +
    '試作比較   … B1に PRE-0001 を選ぶと、生姜だれの試作1回目と2回目が横並びで比較できます。\n\n' +
    'サンプルが不要になったら、各シートのサンプル行を削除してください。'
  );
}


/* ========== ⓪ 仕入先マスタ（発注可能曜日を「コード」ではなく「表」で持つ） ==========
 *
 * ペッパーランチの発注書は発注可能日がコードに直接書かれているため、
 * 店舗が増えるたびにコードを配り直す必要が出ている。
 * 新店では最初からここを表にしておく＝仕入先が増えても店舗が増えても、
 * 行を足すだけで済む。
 */
function build仕入先マスタ(ss) {
  var sh = newSheet_(ss, SUPPLY_SHEET);

  var headers = ['仕入先ID', '仕入先名', '発注可能曜日', '発注締め時間', '最低発注金額', '担当・連絡先', '備考', '登録食材数(自動)'];
  var widths  = [90, 200, 160, 110, 120, 200, 240, 120];
  setHeader_(sh, headers, widths);

  sh.getRange(1, 3).setNote('その仕入先に発注できる曜日を、カンマ区切りで入れます。\n例: 火,水,金,日\n\nここが将来の発注計算の土台になります。仕入先ごとに違っていて構いません。');
  sh.getRange(1, 4).setNote('その日の何時までに出せばその日の発注として扱われるか。例: 15:00');

  var rows = [
    ['SUP-001', '（例）まだ決まっていない仕入先', '火,水,金', '15:00', 10000, '', '仕入先が決まったらこの行を書き換えてください']
  ];
  sh.getRange(2, 1, rows.length, 7).setValues(rows);

  sh.getRange('H2').setFormula('=ARRAYFORMULA(IF(A2:A="","",COUNTIF(食材マスタ!$H:$H,A2:A)))');

  sh.getRange(2, 5, LAST_ROW, 1).setNumberFormat('¥#,##0');
  markAuto_(sh, 8, 1);
}

/** 食材マスタの「仕入先ID」を仕入先マスタから選べるようにする */
function linkSupplier_(ss, ingSheet) {
  var sup = ss.getSheetByName(SUPPLY_SHEET);
  if (!sup) return;
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(sup.getRange('A2:A'), true)
    .setAllowInvalid(true)
    .build();
  ingSheet.getRange(2, 8, LAST_ROW, 1).setDataValidation(rule);
}


/* ========== ① 食材マスタ（買うもの） ========== */
function build食材マスタ(ss) {
  var sh = newSheet_(ss, '食材マスタ');

  var headers = [
    '食材ID', '食材名', '分類', '在庫単位', '保存区分', '日持ち日数', '歩留まり率',
    '仕入先ID', '仕入先商品コード', '発注単位', '入数', '仕入価格', 'kg単価(不定貫)', 'リードタイム(日)',
    '安全係数', '実測補正',
    '単価(自動)', '実効単価(自動)'
  ];
  var widths = [90, 160, 90, 80, 80, 90, 90, 100, 120, 90, 80, 90, 120, 110, 90, 90, 110, 120];
  setHeader_(sh, headers, widths);

  sh.getRange(1, 11).setNote('定貫 … 1発注単位に入っている量（2kgケースなら 2000）。\n不定貫 … 1箱の「目安」重量（約17kgの箱なら 17000）。発注数の計算にだけ使います。');
  sh.getRange(1, 13).setNote('不定貫（重さが箱ごとに違うもの／ラウンドオイスターなど）はここに 1000gあたりの価格 を入れます。\n例: 1000gが2100円 → 2100\n\nここに数字が入っていると、単価は「仕入価格÷入数」ではなく こちらから計算されます。\n定貫のものは空のままにしてください。');
  sh.getRange(1, 15).setNote('欠品を防ぐための余裕。空なら 1.0（余裕なし）。\n例: 1.2 = 必要量の1.2倍を目標在庫にする。\n\n将来の発注計算で使います。今は空のままで構いません。');
  sh.getRange(1, 16).setNote('実際に使った量 ÷ レシピ上の理論量。空なら 1.0（補正なし）。\n米のようにレシピ通りに減らない食材を、実態に合わせるための数字です。\n\n開店して実績が貯まってから入れます。今は空のままで構いません。');
  sh.getRange(1, 14).setNote('発注してから届くまでの日数。仕入先が決まったら入れます。');

  // H〜N列は発注システム連結用。仕入先が決まるまで空でOK（サンプルは動作確認のため仮の値）
  // ING-0010 は不定貫の例。仕入価格は空で、kg単価から単価が出る
  var rows = [
    ['ING-0001', '豚ロース',         '肉',     'g', '冷蔵',   7, 0.98, '', '', 'ケース', 2000, 3200,   '', 2],
    ['ING-0002', '米(精白)',         '米・粉', 'g', '常温', 365, 1.00, '', '', '袋',     5000, 2500,   '', 2],
    ['ING-0003', 'キャベツ',         '野菜',   'g', '冷蔵',   7, 0.80, '', '', 'ケース', 1000,  200,   '', 1],
    ['ING-0004', '醤油',             '調味料', 'g', '常温', 365, 1.00, '', '', '本',     1800,  600,   '', 3],
    ['ING-0005', 'みりん',           '調味料', 'g', '常温', 365, 1.00, '', '', '本',     1000,  400,   '', 3],
    ['ING-0006', '料理酒',           '調味料', 'g', '常温', 365, 1.00, '', '', '本',     1800,  500,   '', 3],
    ['ING-0007', '生姜',             '野菜',   'g', '冷蔵',  14, 0.85, '', '', '袋',     1000, 1200,   '', 2],
    ['ING-0008', '味噌',             '調味料', 'g', '冷蔵', 180, 1.00, '', '', '袋',     1000,  500,   '', 3],
    ['ING-0009', '豆腐',             'その他', 'g', '冷蔵',   5, 1.00, '', '', '丁',      300,   60,   '', 1],
    ['ING-0010', 'ラウンドオイスター', '肉',     'g', '冷凍',  90, 0.92, '', '', '箱',    17000,   '', 2100, 3]
  ];
  sh.getRange(2, 1, rows.length, 14).setValues(rows);

  // 単価 … kg単価が入っていれば不定貫として ÷1000、空なら定貫として 仕入価格÷入数
  // 実効単価 … 単価 ÷ 歩留まり率
  sh.getRange('Q2').setFormula('=ARRAYFORMULA(IF(A2:A="","",IFERROR(IF(M2:M<>"",M2:M/1000,L2:L/K2:K),"")))');
  sh.getRange('R2').setFormula('=ARRAYFORMULA(IF(A2:A="","",IFERROR(Q2:Q/G2:G,"")))');

  addList_(sh, 3,  OPT_分類);
  addList_(sh, 4,  OPT_在庫単位);
  addList_(sh, 5,  OPT_保存区分);
  addList_(sh, 10, OPT_発注単位);
  linkSupplier_(ss, sh);

  sh.getRange(2, 7,  LAST_ROW, 1).setNumberFormat('0%');
  sh.getRange(2, 12, LAST_ROW, 2).setNumberFormat('¥#,##0');
  sh.getRange(2, 17, LAST_ROW, 2).setNumberFormat('¥#,##0.0000');
  markAuto_(sh, 17, 2);

  // 発注連結用ゾーンの目印（仕入先〜実測補正）
  sh.getRange(1, 8, 1, 9).setBackground('#EADFCB');
}


/* ========== ② 仕込み品マスタ（自店で作るもの） ========== */
function build仕込み品マスタ(ss) {
  var sh = newSheet_(ss, '仕込み品マスタ');

  var headers = ['仕込みID', '仕込み品名', '出来上がり量', '単位', '保存区分', '日持ち日数', '原価合計(自動)', '単位原価(自動)', 'チェック(自動)'];
  var widths  = [90, 180, 110, 70, 80, 90, 120, 130, 220];
  setHeader_(sh, headers, widths);

  var rows = [
    ['PRE-0001', '生姜だれ', 1000, 'g', '冷蔵', 7]
  ];
  sh.getRange(2, 1, rows.length, 6).setValues(rows);

  // 原価合計 = レシピ明細の「食材分原価」を親IDで合計
  sh.getRange('G2').setFormula('=ARRAYFORMULA(IF(A2:A="","",SUMIF(レシピ明細!$A:$A,A2:A,レシピ明細!$G:$G)))');
  sh.getRange('H2').setFormula('=ARRAYFORMULA(IF(A2:A="","",IFERROR(G2:G/C2:C,"")))');
  // 仕込み品の中に仕込み品（3階層）が入っていないか見張る
  sh.getRange('I2').setFormula('=ARRAYFORMULA(IF(A2:A="","",IF(COUNTIFS(レシピ明細!$A:$A,A2:A,レシピ明細!$C:$C,"PRE*")>0,"⚠ 仕込み品の中の仕込み品は原価に含まれません","")))');

  addList_(sh, 4, OPT_在庫単位);
  addList_(sh, 5, OPT_保存区分);

  sh.getRange(2, 7, LAST_ROW, 1).setNumberFormat('¥#,##0.0');
  sh.getRange(2, 8, LAST_ROW, 1).setNumberFormat('¥#,##0.0000');
  markAuto_(sh, 7, 3);
}


/* ========== ③ 商品マスタ（売るもの） ========== */
function build商品マスタ(ss) {
  var sh = newSheet_(ss, '商品マスタ');

  var headers = ['商品ID', '商品名', 'カテゴリ', '売価(税抜)', '目標原価率', '提供開始日', '提供終了日', '原価(自動)', '原価率(自動)', '目標差(自動)'];
  var widths  = [100, 220, 100, 110, 110, 110, 110, 110, 110, 110];
  setHeader_(sh, headers, widths);

  var rows = [
    ['MENU-0001', 'サンプル 生姜焼き定食', '定食', 980, 0.35, '', '']
  ];
  sh.getRange(2, 1, rows.length, 7).setValues(rows);

  // 原価 = レシピ明細の「原価計」を商品IDで合計
  sh.getRange('H2').setFormula('=ARRAYFORMULA(IF(A2:A="","",SUMIF(レシピ明細!$A:$A,A2:A,レシピ明細!$I:$I)))');
  sh.getRange('I2').setFormula('=ARRAYFORMULA(IF(A2:A="","",IFERROR(H2:H/D2:D,"")))');
  sh.getRange('J2').setFormula('=ARRAYFORMULA(IF(A2:A="","",IFERROR(I2:I-E2:E,"")))');

  sh.getRange(2, 4,  LAST_ROW, 1).setNumberFormat('¥#,##0');
  sh.getRange(2, 5,  LAST_ROW, 1).setNumberFormat('0.0%');
  sh.getRange(2, 6,  LAST_ROW, 2).setNumberFormat('yyyy/mm/dd');
  sh.getRange(2, 8,  LAST_ROW, 1).setNumberFormat('¥#,##0.0');
  sh.getRange(2, 9,  LAST_ROW, 1).setNumberFormat('0.0%');
  sh.getRange(2, 10, LAST_ROW, 1).setNumberFormat('+0.0%;-0.0%;0.0%');
  markAuto_(sh, 8, 3);

  // 目標を超えたら赤、下回ったら緑
  var over = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0)
    .setBackground('#F7E4E1').setFontColor('#B23B3B')
    .setRanges([sh.getRange(2, 10, LAST_ROW, 1)]).build();
  var under = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThanOrEqualTo(0)
    .setBackground('#E5EFE3').setFontColor('#4B7B4E')
    .setRanges([sh.getRange(2, 10, LAST_ROW, 1)]).build();
  sh.setConditionalFormatRules([over, under]);
}


/* ========== ④ レシピ明細（台帳の心臓部） ========== */
function buildレシピ明細(ss) {
  var sh = newSheet_(ss, 'レシピ明細');

  var headers = ['親ID', '親名(自動)', '子ID', '子名(自動)', '使用量(正味)', '単位(自動)', '食材分原価(自動)', '仕込み分原価(自動)', '原価計(自動)', '備考'];
  var widths  = [100, 200, 100, 180, 110, 80, 130, 140, 110, 200];
  setHeader_(sh, headers, widths);

  // [親ID, 子ID, 使用量, 備考]
  var rows = [
    ['MENU-0001', 'ING-0001', 120, '生姜焼き用'],
    ['MENU-0001', 'ING-0002', 110, '炊く前の重量'],
    ['MENU-0001', 'ING-0003',  50, '千切り'],
    ['MENU-0001', 'PRE-0001',  30, '仕込み品を使う例'],
    ['MENU-0001', 'ING-0008',  15, '味噌汁'],
    ['MENU-0001', 'ING-0009',  30, '味噌汁'],
    ['PRE-0001',  'ING-0004', 400, ''],
    ['PRE-0001',  'ING-0005', 300, ''],
    ['PRE-0001',  'ING-0006', 200, ''],
    ['PRE-0001',  'ING-0007', 100, 'すりおろし']
  ];
  var n = rows.length;
  sh.getRange(2,  1, n, 1).setValues(rows.map(function (r) { return [r[0]]; }));
  sh.getRange(2,  3, n, 1).setValues(rows.map(function (r) { return [r[1]]; }));
  sh.getRange(2,  5, n, 1).setValues(rows.map(function (r) { return [r[2]]; }));
  sh.getRange(2, 10, n, 1).setValues(rows.map(function (r) { return [r[3]]; }));

  // 親名：MENUなら商品マスタ、それ以外は仕込み品マスタを引く
  sh.getRange('B2').setFormula('=ARRAYFORMULA(IF(A2:A="","",IFERROR(IF(LEFT(A2:A,4)="MENU",VLOOKUP(A2:A,商品マスタ!$A:$B,2,FALSE),VLOOKUP(A2:A,仕込み品マスタ!$A:$B,2,FALSE)),"⚠ID不明")))');
  // 子名：INGなら食材マスタ、PREなら仕込み品マスタを引く
  sh.getRange('D2').setFormula('=ARRAYFORMULA(IF(C2:C="","",IFERROR(IF(LEFT(C2:C,3)="ING",VLOOKUP(C2:C,食材マスタ!$A:$B,2,FALSE),VLOOKUP(C2:C,仕込み品マスタ!$A:$B,2,FALSE)),"⚠ID不明")))');
  sh.getRange('F2').setFormula('=ARRAYFORMULA(IF(C2:C="","",IFERROR(IF(LEFT(C2:C,3)="ING",VLOOKUP(C2:C,食材マスタ!$A:$D,4,FALSE),VLOOKUP(C2:C,仕込み品マスタ!$A:$D,4,FALSE)),"")))');
  // 食材分と仕込み分を別の列にすることで循環参照を避けている（構造上とても重要）
  sh.getRange('G2').setFormula('=ARRAYFORMULA(IF(C2:C="","",IF(LEFT(C2:C,3)="ING",IFERROR(E2:E*VLOOKUP(C2:C,食材マスタ!$A:$R,18,FALSE),0),0)))');
  sh.getRange('H2').setFormula('=ARRAYFORMULA(IF(C2:C="","",IF(LEFT(C2:C,3)="PRE",IFERROR(E2:E*VLOOKUP(C2:C,仕込み品マスタ!$A:$H,8,FALSE),0),0)))');
  sh.getRange('I2').setFormula('=ARRAYFORMULA(IF(C2:C="","",G2:G+H2:H))');

  sh.getRange(2, 7, LAST_ROW, 3).setNumberFormat('¥#,##0.0');
  markAuto_(sh, 2, 1);
  markAuto_(sh, 4, 1);
  markAuto_(sh, 6, 4);
}


/* ========== ⑤ 調理手順 ========== */
function build調理手順(ss) {
  var sh = newSheet_(ss, '調理手順');

  var headers = ['商品ID', '商品名(自動)', '手順番号', '手順内容', '所要時間', '温度・火加減', 'ポイント'];
  var widths  = [100, 200, 90, 380, 100, 110, 280];
  setHeader_(sh, headers, widths);

  // [商品ID, 手順番号, 手順内容, 所要時間, 温度・火加減, ポイント]
  var rows = [
    ['MENU-0001', 1, '豚ロースを常温に戻し、筋切りをする',              '3分', '',     '冷たいまま焼くと硬くなる'],
    ['MENU-0001', 2, 'フライパンを中火で熱し、豚ロースを両面焼く',      '4分', '中火', '焼き色がついたら返す'],
    ['MENU-0001', 3, '生姜だれ30gを加えて絡める',                       '1分', '弱火', '煮詰めすぎない'],
    ['MENU-0001', 4, 'ご飯・千切りキャベツと共に盛り付ける',            '1分', '',     'キャベツは冷やしておく'],
    ['MENU-0001', 5, '味噌汁を仕上げて提供',                            '1分', '',     '提供直前に味噌を溶く']
  ];
  var n = rows.length;
  sh.getRange(2, 1, n, 1).setValues(rows.map(function (r) { return [r[0]]; }));
  sh.getRange(2, 3, n, 5).setValues(rows.map(function (r) { return r.slice(1); }));

  sh.getRange('B2').setFormula('=ARRAYFORMULA(IF(A2:A="","",IFERROR(VLOOKUP(A2:A,商品マスタ!$A:$B,2,FALSE),"⚠ID不明")))');

  sh.getRange(2, 4, LAST_ROW, 1).setWrap(true);
  sh.getRange(2, 7, LAST_ROW, 1).setWrap(true);
  markAuto_(sh, 2, 1);
}


/* ========== 共通ヘルパー ========== */

function newSheet_(ss, name) {
  var all = ss.getSheets();
  // 新規作成直後の空シートがあれば、それを使い回す（シートを削除しないため）
  if (all.length === 1 && all[0].getLastRow() === 0 &&
      (all[0].getName() === 'シート1' || all[0].getName() === 'Sheet1')) {
    return all[0].setName(name);
  }
  return ss.insertSheet(name);
}

function setHeader_(sh, headers, widths) {
  sh.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground(C_HEADER_BG)
    .setFontColor('#2A2019')
    .setVerticalAlignment('middle');
  sh.setFrozenRows(1);
  sh.setRowHeight(1, 34);
  widths.forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
}

// 自動計算列に色をつけて「手入力しない列」と分かるようにする
function markAuto_(sh, startCol, numCols) {
  sh.getRange(2, startCol, LAST_ROW, numCols).setBackground(C_AUTO_BG);
}

/**
 * 「評価に入る味」と「参考の味」を枠で見分けられるようにする。
 * coreCol から4列＝甘塩酸旨（実線の濃い枠）、その右2列＝苦辛（破線の薄い枠）。
 */
function tasteBorders_(sh, coreCol, refCol) {
  sh.getRange(1, coreCol, LAST_ROW + 1, 4)
    .setBorder(true, true, true, true, false, false, C_TASTE_CORE, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sh.getRange(1, refCol, LAST_ROW + 1, 2)
    .setBorder(true, true, true, true, false, false, C_TASTE_REF, SpreadsheetApp.BorderStyle.DASHED);
}

function addList_(sh, col, options) {
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(options, true)
    .setAllowInvalid(true)
    .build();
  sh.getRange(2, col, LAST_ROW, 1).setDataValidation(rule);
}


/* =====================================================================
 *  ここから下は「試作」機能
 * ===================================================================== */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🧪 試作')
    .addItem('次の回の試作をつくる', '次の回の試作をつくる')
    .addItem('採用した試作をレシピへ反映', '採用試作をレシピへ反映')
    .addSeparator()
    .addItem('試作シートを追加(初回のみ)', '試作シートを追加')
    .addItem('不定貫の列を追加(初回のみ)', '不定貫の列を追加')
    .addItem('発注準備の列を追加(初回のみ)', '発注準備の列を追加')
    .addItem('辛味の列を追加(初回のみ)', '辛味の列を追加')
    .addItem('総合評価を自動計算にする(初回のみ)', '総合評価を自動計算にする')
    .addSeparator()
    .addItem('試作比較を作り直す', '試作比較を作り直す')
    .addToUi();
}


/**
 * 苦味と旨味を入れ替え（→ 甘塩酸旨・苦辛 の並び）、総合評価を手入力から自動計算に変える。
 * 入力済みの味の数字はそのまま持ち越す。
 */
function 総合評価を自動計算にする() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var rec = ss.getSheetByName('試作記録');
  if (!rec) { ui.alert('試作記録が見つかりません。\n\n先に「試作シートを追加」を実行してください。'); return; }
  if (String(rec.getRange(1, 11).getValue()).trim() !== '辛味') {
    ui.alert('中止しました。\n\n先に「辛味の列を追加」を実行してください。');
    return;
  }
  if (String(rec.getRange(1, 9).getValue()).trim() === '旨味') { ui.alert('すでに変更済みです。'); return; }
  if (String(rec.getRange(1, 9).getValue()).trim() !== '苦味') {
    ui.alert('中止しました。\n\n試作記録のI列が想定(苦味)と違います。\n列を手で動かしている場合は、先に元に戻してください。');
    return;
  }

  if (ui.alert('確認',
      '・苦味と旨味の列を入れ替えます（並びが 甘・塩・酸・旨・苦・辛 になります）\n' +
      '・総合評価を手入力から自動計算に変えます\n' +
      '　（甘味・塩味・酸味・旨味が10にどれだけ近いかで算出。苦味・辛味は入れません）\n\n' +
      '※ 今まで手で入れた総合評価の点数は、自動計算に置き換わります。\n' +
      '　 味の数字はそのまま残ります。よろしいですか?',
      ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;

  // 苦味(I)と旨味(J)の中身を入れ替える
  var last = rec.getLastRow();
  if (last >= 2) {
    var v = rec.getRange(2, 9, last - 1, 2).getValues();
    rec.getRange(2, 9, last - 1, 2).setValues(v.map(function (r) { return [r[1], r[0]]; }));
  }
  rec.getRange(1, 9, 1, 2).setValues([['旨味', '苦味']]);

  rec.getRange(1, 6, 1, 4).setNote('味の強さを 1〜20 で入れます。\n  10前後 = ちょうど良い\n  20に近い = 強すぎ\n  1に近い = 足りない\n\nこの4つ（甘味・塩味・酸味・旨味）が総合評価の土台になります。');
  rec.getRange(1, 10, 1, 2).setNote('苦味・辛味も 1〜20 で記録しますが、総合評価の計算には入れません。\n料理によって狙いが大きく変わるためです。');

  // 総合評価を自動計算に
  rec.getRange(1, 12).setValue('総合評価(自動)');
  rec.setColumnWidth(12, 110);
  rec.getRange(2, 12, LAST_ROW, 1).clearContent().clearDataValidations();
  rec.getRange('L2').setFormula(OVERALL_FORMULA);
  rec.getRange(1, 12).setNote('甘味・塩味・酸味・旨味が 10（ちょうど良い）にどれだけ近いかを自動計算した点数です。\n20に近いほど狙い通り。手入力はできません。\n4つのうち1つでも空欄だと計算されません。');

  rec.getRange(2,  6, LAST_ROW, 6).setNumberFormat('0').setHorizontalAlignment('center');
  rec.getRange(2, 12, LAST_ROW, 1).setNumberFormat('0.0').setHorizontalAlignment('center');
  rec.getRange(2, 13, LAST_ROW, 1).setNumberFormat('¥#,##0.0');
  markAuto_(rec, 12, 2);
  tasteBorders_(rec, 6, 10);

  build試作比較(ss);

  ss.setActiveSheet(rec);
  ui.alert('変更しました。\n\n' +
           '・並び: 甘味・塩味・酸味・旨味（実線の枠）／ 苦味・辛味（破線の枠）\n' +
           '・総合評価は自動計算になりました（薄い色の列＝手入力しない）\n\n' +
           '試作比較も作り直してあります。');
}


/** 試作比較は中身が全部数式＝いつ作り直しても安全。表示が崩れたときや設定を入れ直したいとき用 */
function 試作比較を作り直す() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName('試作記録')) {
    SpreadsheetApp.getUi().alert('試作記録が見つかりません。\n\n先に「試作シートを追加」を実行してください。');
    return;
  }
  build試作比較(ss);
  ss.setActiveSheet(ss.getSheetByName('試作比較'));
  SpreadsheetApp.getUi().alert('試作比較を作り直しました。\n\nB1で対象を選んでください（B1だけは警告なしで触れます）。');
}


/**
 * 既にある試作記録に「辛味」列を足す（旨味の右）。入力済みのデータは消さない。
 * 試作比較は中身が全部数式なので、行が1本増えるぶん作り直す。
 */
function 辛味の列を追加() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var rec = ss.getSheetByName('試作記録');
  if (!rec) { ui.alert('試作記録が見つかりません。\n\n先に「試作シートを追加」を実行してください。'); return; }

  var k = String(rec.getRange(1, 11).getValue()).trim();
  if (k === '辛味') { ui.alert('すでに追加済みです。'); return; }
  if (k !== '総合評価') {
    ui.alert('中止しました。\n\n試作記録のK列が想定(総合評価)と違います。\n列を手で動かしている場合は、先に元に戻してください。');
    return;
  }

  if (ui.alert('確認',
      '試作記録の「旨味」の右に「辛味」列を追加します。\n' +
      '試作比較は行が1本増えるので作り直します（入力データは持っていないシートです）。\n\n' +
      '入力済みのデータは消えません。よろしいですか?',
      ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;

  rec.insertColumnAfter(10);                      // J(旨味) の右
  rec.setColumnWidth(11, 60);
  rec.getRange(1, 11).setValue('辛味')
    .setFontWeight('bold').setBackground(C_HEADER_BG).setFontColor('#2A2019')
    .setVerticalAlignment('middle');
  rec.getRange(1, 6, 1, 6).setNote('味の強さを 1〜20 で入れます。\n  10前後 = ちょうど良い\n  20に近い = 強すぎ\n  1に近い = 足りない\n※ 使わない味の見出しは「コク」などに書き換えて構いません（列の位置は変えないこと）。');
  rec.getRange(1, 12).setNote('総合評価も 1〜20。こちらは大きいほど良い、という意味です。');

  var score = SpreadsheetApp.newDataValidation().requireNumberBetween(1, 20).setAllowInvalid(false)
    .setHelpText('1〜20 で入れてください（10前後 = ちょうど良い）').build();
  rec.getRange(2, 6, LAST_ROW, 6).setDataValidation(score);
  rec.getRange(2, 6, LAST_ROW, 7).setHorizontalAlignment('center');

  // 原価(自動)が L→M にずれるので入れ直す
  rec.getRange(2, 13, LAST_ROW, 1).clearContent();
  rec.getRange('M2').setFormula('=ARRAYFORMULA(IF(A2:A="","",SUMIF(試作レシピ明細!$A:$A,A2:A,試作レシピ明細!$G:$G)))');
  rec.getRange(2, 13, LAST_ROW, 1).setNumberFormat('¥#,##0.0');
  rec.getRange(2, 14, LAST_ROW, 1).setWrap(true);
  markAuto_(rec, 13, 1);

  build試作比較(ss);                                // 行が増えるので作り直す

  ss.setActiveSheet(ss.getSheetByName('試作比較'));
  ui.alert('辛味を追加しました。\n\n試作記録に「辛味」列が増え、試作比較にも辛味の行が出ます。\n過去の回の辛味は空欄なので、必要なら後から入れてください。');
}


/**
 * 将来の発注システムに備えて、既にある台帳に次を足す。
 *   ・食材マスタ … 「安全係数」「実測補正」の2列（リードタイムの右）
 *   ・仕入先マスタ … 新規シート。発注可能曜日をコードでなく表で持つ
 * 入力済みのデータは消さない。
 */
function 発注準備の列を追加() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var ing = ss.getSheetByName('食材マスタ');
  if (!ing) { ui.alert('食材マスタが見つかりません。'); return; }

  var needCols = String(ing.getRange(1, 15).getValue()).indexOf('安全係数') !== 0;
  var needSup  = ss.getSheetByName(SUPPLY_SHEET) === null;

  if (!needCols && !needSup) { ui.alert('すでに追加済みです。'); return; }

  if (needCols) {
    if (String(ing.getRange(1, 13).getValue()).indexOf('kg単価') !== 0) {
      ui.alert('中止しました。\n\n先に「不定貫の列を追加」を実行してください。\n(食材マスタのM列が kg単価 になっている必要があります)');
      return;
    }
    if (String(ing.getRange(1, 15).getValue()).indexOf('単価') !== 0) {
      ui.alert('中止しました。\n\n食材マスタのO列が想定(単価(自動))と違います。\n列を手で動かしている場合は、先に元に戻してください。');
      return;
    }
  }

  var plan = [];
  if (needSup)  plan.push('・「仕入先マスタ」シートを新しく作る');
  if (needCols) plan.push('・食材マスタに「安全係数」「実測補正」の2列を足す（リードタイムの右）');
  if (ui.alert('確認', plan.join('\n') + '\n\n入力済みのデータは消えません。\nよろしいですか?', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;

  if (needSup) {
    build仕入先マスタ(ss);
    ss.setActiveSheet(ss.getSheetByName(SUPPLY_SHEET));
    ss.moveActiveSheet(1);
  }

  if (needCols) {
    ing.insertColumnsAfter(14, 2);                  // N(リードタイム) の右に2本
    ing.setColumnWidth(15, 90);
    ing.setColumnWidth(16, 90);
    ing.getRange(1, 15, 1, 2).setValues([['安全係数', '実測補正']])
      .setFontWeight('bold').setBackground(C_HEADER_BG).setFontColor('#2A2019')
      .setVerticalAlignment('middle');
    ing.getRange(1, 8).setValue('仕入先ID');
    ing.getRange(1, 9).setValue('仕入先商品コード');
    ing.getRange(1, 8, 1, 9).setBackground('#EADFCB');

    ing.getRange(1, 14).setNote('発注してから届くまでの日数。仕入先が決まったら入れます。');
    ing.getRange(1, 15).setNote('欠品を防ぐための余裕。空なら 1.0（余裕なし）。\n例: 1.2 = 必要量の1.2倍を目標在庫にする。\n\n将来の発注計算で使います。今は空のままで構いません。');
    ing.getRange(1, 16).setNote('実際に使った量 ÷ レシピ上の理論量。空なら 1.0（補正なし）。\n米のようにレシピ通りに減らない食材を、実態に合わせるための数字です。\n\n開店して実績が貯まってから入れます。今は空のままで構いません。');

    // 単価・実効単価を正しい位置で引き直す
    ing.getRange(2, 17, LAST_ROW, 2).clearContent();
    ing.getRange('Q2').setFormula('=ARRAYFORMULA(IF(A2:A="","",IFERROR(IF(M2:M<>"",M2:M/1000,L2:L/K2:K),"")))');
    ing.getRange('R2').setFormula('=ARRAYFORMULA(IF(A2:A="","",IFERROR(Q2:Q/G2:G,"")))');
    ing.getRange(2, 17, LAST_ROW, 2).setNumberFormat('¥#,##0.0000');
    markAuto_(ing, 17, 2);

    // 実効単価を見に行っている2か所は 16→18 にずれるので入れ直す
    var rep = ss.getSheetByName('レシピ明細');
    if (rep) rep.getRange('G2').setFormula('=ARRAYFORMULA(IF(C2:C="","",IF(LEFT(C2:C,3)="ING",IFERROR(E2:E*VLOOKUP(C2:C,食材マスタ!$A:$R,18,FALSE),0),0)))');
    var tri = ss.getSheetByName('試作レシピ明細');
    if (tri) tri.getRange('F2').setFormula('=ARRAYFORMULA(IF(B2:B="","",IFERROR(IF(LEFT(B2:B,3)="ING",VLOOKUP(B2:B,食材マスタ!$A:$R,18,FALSE),VLOOKUP(B2:B,仕込み品マスタ!$A:$H,8,FALSE)),"")))');
  }

  linkSupplier_(ss, ing);

  ss.setActiveSheet(ing);
  ui.alert('追加しました。\n\n' +
           '・仕入先マスタ … 仕入先が決まったら、発注可能曜日を「火,水,金」の形で入れます\n' +
           '・安全係数 / 実測補正 … 今は空のままでOK（空 = 1.0 として扱います）\n\n' +
           '商品マスタの原価が変わっていないか確認してください。');
}


/**
 * 既にある台帳の食材マスタに「kg単価(不定貫)」列を足す。
 * 入力済みのデータは消さずに、列を1本差し込んで関係する数式だけ引き直す。
 */
function 不定貫の列を追加() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var ing = ss.getSheetByName('食材マスタ');
  if (!ing) { ui.alert('食材マスタが見つかりません。'); return; }

  if (String(ing.getRange(1, 13).getValue()).indexOf('kg単価') === 0) {
    ui.alert('すでに追加済みです。\n\n食材マスタのM列が「kg単価(不定貫)」になっています。');
    return;
  }
  if (String(ing.getRange(1, 13).getValue()).indexOf('リードタイム') !== 0) {
    ui.alert('中止しました。\n\n食材マスタのM列が想定(リードタイム)と違います。\n列を手で動かしている場合は、先に元に戻してください。');
    return;
  }

  if (ui.alert('確認',
      '食材マスタの「仕入価格」の右に「kg単価(不定貫)」列を追加します。\n' +
      '入力済みのデータは消えません。\n\nよろしいですか?',
      ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;

  ing.insertColumnAfter(12);                       // L(仕入価格) の右に差し込む
  ing.setColumnWidth(13, 120);
  ing.getRange(1, 13).setValue('kg単価(不定貫)')
    .setFontWeight('bold').setBackground(C_HEADER_BG).setFontColor('#2A2019')
    .setVerticalAlignment('middle');
  ing.getRange(1, 8, 1, 7).setBackground('#EADFCB');

  ing.getRange(1, 11).setNote('定貫 … 1発注単位に入っている量（2kgケースなら 2000）。\n不定貫 … 1箱の「目安」重量（約17kgの箱なら 17000）。発注数の計算にだけ使います。');
  ing.getRange(1, 13).setNote('不定貫（重さが箱ごとに違うもの／ラウンドオイスターなど）はここに 1000gあたりの価格 を入れます。\n例: 1000gが2100円 → 2100\n\nここに数字が入っていると、単価は「仕入価格÷入数」ではなく こちらから計算されます。\n定貫のものは空のままにしてください。');

  // 単価・実効単価を引き直す（列がずれた分を正しい形で入れ直す）
  ing.getRange(2, 15, LAST_ROW, 2).clearContent();
  ing.getRange('O2').setFormula('=ARRAYFORMULA(IF(A2:A="","",IFERROR(IF(M2:M<>"",M2:M/1000,L2:L/K2:K),"")))');
  ing.getRange('P2').setFormula('=ARRAYFORMULA(IF(A2:A="","",IFERROR(O2:O/G2:G,"")))');
  ing.getRange(2, 12, LAST_ROW, 2).setNumberFormat('¥#,##0');
  ing.getRange(2, 15, LAST_ROW, 2).setNumberFormat('¥#,##0.0000');
  markAuto_(ing, 15, 2);

  // 実効単価を参照している2か所は、列番号が15→16になるので入れ直す
  var rep = ss.getSheetByName('レシピ明細');
  if (rep) {
    rep.getRange('G2').setFormula('=ARRAYFORMULA(IF(C2:C="","",IF(LEFT(C2:C,3)="ING",IFERROR(E2:E*VLOOKUP(C2:C,食材マスタ!$A:$P,16,FALSE),0),0)))');
  }
  var trial = ss.getSheetByName('試作レシピ明細');
  if (trial) {
    trial.getRange('F2').setFormula('=ARRAYFORMULA(IF(B2:B="","",IFERROR(IF(LEFT(B2:B,3)="ING",VLOOKUP(B2:B,食材マスタ!$A:$P,16,FALSE),VLOOKUP(B2:B,仕込み品マスタ!$A:$H,8,FALSE)),"")))');
  }

  ss.setActiveSheet(ing);
  ui.alert('追加しました。\n\n' +
           '不定貫のものは次のように入れてください。\n' +
           '  入数     … 1箱の目安重量（例: 17000）\n' +
           '  仕入価格 … 空のまま（毎回変わるため）\n' +
           '  kg単価   … 1000gあたりの価格（例: 2100）\n\n' +
           '商品マスタの原価が変わっていないか確認してください。');
}


/** 既にある台帳に、試作シート3枚だけを足す */
function 試作シートを追加() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var missing = BASE_SHEETS.filter(function (n) { return ss.getSheetByName(n) === null; });
  if (missing.length > 0) {
    SpreadsheetApp.getUi().alert('中止しました。\n\n先に台帳本体が必要です。無いシート:\n  ' + missing.join('\n  '));
    return;
  }
  var existing = TRIAL_SHEETS.filter(function (n) { return ss.getSheetByName(n) !== null; });
  if (existing.length > 0) {
    SpreadsheetApp.getUi().alert('中止しました。\n\n既に次のシートがあります:\n  ' + existing.join('\n  '));
    return;
  }

  build試作記録(ss);
  build試作レシピ明細(ss);
  build試作比較(ss);

  TRIAL_SHEETS.forEach(function (name) {
    ss.setActiveSheet(ss.getSheetByName(name));
    ss.moveActiveSheet(ss.getNumSheets());        // 末尾へ（仕入先マスタの有無に左右されない）
  });
  ss.setActiveSheet(ss.getSheetByName('試作比較'));

  SpreadsheetApp.getUi().alert(
    '試作シートを追加しました。\n\n' +
    'サンプルとして、生姜だれ(PRE-0001)の試作1回目・2回目が入っています。\n' +
    '試作比較シートのB1で PRE-0001 を選ぶと、分量と味の変化が横並びで見られます。\n\n' +
    'この画面を一度閉じて開き直すと、上部に「🧪 試作」メニューが出ます。'
  );
}


/* ========== ⑥ 試作記録（1行 = 1回の試作） ========== */
function build試作記録(ss) {
  var sh = newSheet_(ss, '試作記録');

  var headers = ['試作ID', '対象ID', '対象名(自動)', '回数', '試作日',
                 '甘味', '塩味', '酸味', '旨味', '苦味', '辛味',
                 '総合評価(自動)', '原価(自動)', '所感・次回の調整', '採用'];
  var widths  = [90, 100, 180, 60, 100, 60, 60, 60, 60, 60, 60, 110, 100, 380, 140];
  setHeader_(sh, headers, widths);

  sh.getRange(1, 6, 1, 4).setNote('味の強さを 1〜20 で入れます。\n  10前後 = ちょうど良い\n  20に近い = 強すぎ\n  1に近い = 足りない\n\nこの4つ（甘味・塩味・酸味・旨味）が総合評価の土台になります。');
  sh.getRange(1, 10, 1, 2).setNote('苦味・辛味も 1〜20 で記録しますが、総合評価の計算には入れません。\n料理によって狙いが大きく変わるためです。');
  sh.getRange(1, 12).setNote('甘味・塩味・酸味・旨味が 10（ちょうど良い）にどれだけ近いかを自動計算した点数です。\n20に近いほど狙い通り。手入力はできません。\n4つのうち1つでも空欄だと計算されません。');

  // [試作ID, 対象ID, 回数, 試作日, 甘, 塩, 酸, 旨, 苦, 辛, 所感, 採用]  ※総合評価と原価は自動
  var rows = [
    ['TRY-0001', 'PRE-0001', 1, new Date(2026, 8, 8),  17,  6,  9,  8, 3, 5, '甘さが立ちすぎ。醤油を増やしてみりんを減らす', ''],
    ['TRY-0002', 'PRE-0001', 2, new Date(2026, 8, 11), 11, 10, 10, 12, 3, 7, 'バランス改善。生姜をもう少し効かせたい', '']
  ];
  var n = rows.length;
  sh.getRange(2,  1, n, 2).setValues(rows.map(function (r) { return [r[0], r[1]]; }));
  sh.getRange(2,  4, n, 8).setValues(rows.map(function (r) { return r.slice(2, 10); }));
  sh.getRange(2, 14, n, 2).setValues(rows.map(function (r) { return [r[10], r[11]]; }));

  sh.getRange('C2').setFormula('=ARRAYFORMULA(IF(B2:B="","",IFERROR(IF(LEFT(B2:B,4)="MENU",VLOOKUP(B2:B,商品マスタ!$A:$B,2,FALSE),VLOOKUP(B2:B,仕込み品マスタ!$A:$B,2,FALSE)),"⚠ID不明")))');
  sh.getRange('L2').setFormula(OVERALL_FORMULA);
  sh.getRange('M2').setFormula('=ARRAYFORMULA(IF(A2:A="","",SUMIF(試作レシピ明細!$A:$A,A2:A,試作レシピ明細!$G:$G)))');

  // 20段階。プルダウンだと選ぶのが遅いので、数字を直接打てる形にしている
  var score = SpreadsheetApp.newDataValidation().requireNumberBetween(1, 20).setAllowInvalid(false)
    .setHelpText('1〜20 で入れてください（10前後 = ちょうど良い）').build();
  sh.getRange(2, 6, LAST_ROW, 6).setDataValidation(score);
  addList_(sh, 15, ['○']);

  sh.getRange(2,  5, LAST_ROW, 1).setNumberFormat('yyyy/mm/dd');
  sh.getRange(2,  6, LAST_ROW, 6).setNumberFormat('0').setHorizontalAlignment('center');
  sh.getRange(2, 12, LAST_ROW, 1).setNumberFormat('0.0').setHorizontalAlignment('center');
  sh.getRange(2, 13, LAST_ROW, 1).setNumberFormat('¥#,##0.0');
  sh.getRange(2, 14, LAST_ROW, 1).setWrap(true);
  markAuto_(sh, 3, 1);
  markAuto_(sh, 12, 2);
  tasteBorders_(sh, 6, 10);
}


/* ========== ⑦ 試作レシピ明細（その回の分量） ========== */
function build試作レシピ明細(ss) {
  var sh = newSheet_(ss, '試作レシピ明細');

  var headers = ['試作ID', '子ID', '子名(自動)', '使用量', '単位(自動)', '単価(自動)', '原価(自動)', '備考'];
  var widths  = [90, 100, 180, 90, 80, 110, 100, 260];
  setHeader_(sh, headers, widths);

  // [試作ID, 子ID, 使用量, 備考]
  var rows = [
    ['TRY-0001', 'ING-0004', 400, ''],
    ['TRY-0001', 'ING-0005', 300, ''],
    ['TRY-0001', 'ING-0006', 200, ''],
    ['TRY-0001', 'ING-0007', 100, 'すりおろし'],
    ['TRY-0002', 'ING-0004', 420, '醤油を増やした'],
    ['TRY-0002', 'ING-0005', 250, 'みりんを減らした'],
    ['TRY-0002', 'ING-0006', 200, ''],
    ['TRY-0002', 'ING-0007', 110, 'すりおろし']
  ];
  var n = rows.length;
  sh.getRange(2, 1, n, 2).setValues(rows.map(function (r) { return [r[0], r[1]]; }));
  sh.getRange(2, 4, n, 1).setValues(rows.map(function (r) { return [r[2]]; }));
  sh.getRange(2, 8, n, 1).setValues(rows.map(function (r) { return [r[3]]; }));

  sh.getRange('C2').setFormula('=ARRAYFORMULA(IF(B2:B="","",IFERROR(IF(LEFT(B2:B,3)="ING",VLOOKUP(B2:B,食材マスタ!$A:$B,2,FALSE),VLOOKUP(B2:B,仕込み品マスタ!$A:$B,2,FALSE)),"⚠ID不明")))');
  sh.getRange('E2').setFormula('=ARRAYFORMULA(IF(B2:B="","",IFERROR(IF(LEFT(B2:B,3)="ING",VLOOKUP(B2:B,食材マスタ!$A:$D,4,FALSE),VLOOKUP(B2:B,仕込み品マスタ!$A:$D,4,FALSE)),"")))');
  sh.getRange('F2').setFormula('=ARRAYFORMULA(IF(B2:B="","",IFERROR(IF(LEFT(B2:B,3)="ING",VLOOKUP(B2:B,食材マスタ!$A:$R,18,FALSE),VLOOKUP(B2:B,仕込み品マスタ!$A:$H,8,FALSE)),"")))');
  sh.getRange('G2').setFormula('=ARRAYFORMULA(IF(B2:B="","",IFERROR(D2:D*F2:F,"")))');

  sh.getRange(2, 6, LAST_ROW, 1).setNumberFormat('¥#,##0.0000');
  sh.getRange(2, 7, LAST_ROW, 1).setNumberFormat('¥#,##0.0');
  markAuto_(sh, 3, 1);
  markAuto_(sh, 5, 3);
}


/* ========== ⑧ 試作比較（回を横に並べて見る画面） ========== */
function build試作比較(ss) {
  var sh = ss.getSheetByName('試作比較') || newSheet_(ss, '試作比較');

  // このシートは中身が全部数式＝作り直して問題ない。何度実行しても同じ状態になるようにする
  resetSheet_(sh);

  sh.setColumnWidth(1, 130);
  sh.setColumnWidth(2, 190);
  for (var c = 0; c < MAX_TRIALS; c++) sh.setColumnWidth(3 + c, 86);

  sh.getRange('A1').setValue('対象を選ぶ →').setFontWeight('bold');
  sh.getRange('B1').setBackground('#FFFFFF').setBorder(true, true, true, true, false, false)
    .setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange('C1').setFormula('=IFERROR(IF(LEFT($B$1,4)="MENU",VLOOKUP($B$1,商品マスタ!$A:$B,2,FALSE),VLOOKUP($B$1,仕込み品マスタ!$A:$B,2,FALSE)),"")')
    .setFontWeight('bold');

  // 選択候補（商品IDと仕込みIDをまとめた裏方リスト）
  sh.getRange('Z1').setFormula('=IFERROR(FILTER({商品マスタ!A2:A;仕込み品マスタ!A2:A},{商品マスタ!A2:A;仕込み品マスタ!A2:A}<>""),"")');
  sh.hideColumns(26);
  sh.getRange('B1').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInRange(sh.getRange('Z1:Z'), true).setAllowInvalid(true).build()
  );

  var labels = ['回数', '試作日', '甘味', '塩味', '酸味', '旨味', '苦味', '辛味', '総合評価', '原価', '所感'];
  sh.getRange(3, 1, labels.length, 1).setValues(labels.map(function (l) { return [l]; })).setFontWeight('bold');
  sh.getRange(3, 1, labels.length, 1).setBackground(C_HEADER_BG);
  sh.getRange('A15').setValue('【分量】').setFontWeight('bold');
  sh.getRange('A15:B15').setBackground(C_HEADER_BG);

  var nums = [], idRow = [], dateRow = [], total = [], cost = [], memo = [], amount = [];
  var taste = [[], [], [], [], [], []];                 // 甘 塩 酸 苦 旨 辛 ＝ 試作記録 F〜K
  for (var i = 0; i < MAX_TRIALS; i++) {
    var col = columnLetter_(3 + i);
    nums.push(i + 1);
    idRow.push('=IFERROR(INDEX(FILTER(試作記録!$A:$A,試作記録!$B:$B=$B$1,試作記録!$D:$D=' + col + '$3),1),"")');
    dateRow.push(vl_(col, 5));
    for (var t = 0; t < 6; t++) taste[t].push(vl_(col, 6 + t));
    total.push(vl_(col, 12));
    cost.push(vl_(col, 13));
    memo.push(vl_(col, 14));
    amount.push('=ARRAYFORMULA(IF($A16:$A="","",IF(' + col + '$2="","",IFERROR(SUMIFS(試作レシピ明細!$D:$D,試作レシピ明細!$A:$A,' + col + '$2,試作レシピ明細!$B:$B,$A16:$A),""))))');
  }

  sh.getRange(2, 3, 1, MAX_TRIALS).setFormulas([idRow]);
  sh.hideRows(2);                                   // 裏方の行（試作IDの対応表）
  sh.getRange(3, 3, 1, MAX_TRIALS).setValues([nums]).setFontWeight('bold').setHorizontalAlignment('center').setBackground(C_HEADER_BG);
  sh.getRange(4,  3, 1, MAX_TRIALS).setFormulas([dateRow]).setNumberFormat('m/d');
  for (var t2 = 0; t2 < 6; t2++) sh.getRange(5 + t2, 3, 1, MAX_TRIALS).setFormulas([taste[t2]]);
  sh.getRange(5,  3, 6, MAX_TRIALS).setNumberFormat('0');
  sh.getRange(11, 3, 1, MAX_TRIALS).setFormulas([total]).setNumberFormat('0.0');
  sh.getRange(12, 3, 1, MAX_TRIALS).setFormulas([cost]).setNumberFormat('¥#,##0.0');
  sh.getRange(13, 3, 1, MAX_TRIALS).setFormulas([memo]);
  sh.getRange(4, 3, 9, MAX_TRIALS).setHorizontalAlignment('center');
  sh.getRange(13, 3, 1, MAX_TRIALS).setHorizontalAlignment('left').setWrap(true);

  // 評価に入る4つ（甘塩酸旨）と、参考の2つ（苦辛）を枠で見分けられるようにする
  var w = MAX_TRIALS + 2;
  sh.getRange(5, 1, 4, w).setBorder(true, true, true, true, false, false, C_TASTE_CORE, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sh.getRange(9, 1, 2, w).setBorder(true, true, true, true, false, false, C_TASTE_REF, SpreadsheetApp.BorderStyle.DASHED);
  sh.getRange('A5').setNote('この4つ（甘味・塩味・酸味・旨味）が総合評価の計算に入ります。');
  sh.getRange('A9').setNote('苦味・辛味は記録するだけで、総合評価の計算には入りません。');

  // 分量ブロック：食材の一覧は選んだ対象の全試作回から自動で集める
  sh.getRange('A16').setFormula('=IFERROR(SORT(UNIQUE(FILTER(試作レシピ明細!$B$2:$B,ISNUMBER(MATCH(試作レシピ明細!$A$2:$A,FILTER(試作記録!$A$2:$A,試作記録!$B$2:$B=$B$1),0))))),"")');
  sh.getRange('B16').setFormula('=ARRAYFORMULA(IF($A16:$A="","",IFERROR(IF(LEFT($A16:$A,3)="ING",VLOOKUP($A16:$A,食材マスタ!$A:$B,2,FALSE),VLOOKUP($A16:$A,仕込み品マスタ!$A:$B,2,FALSE)),"")))');
  sh.getRange(16, 3, 1, MAX_TRIALS).setFormulas([amount]);
  sh.getRange(16, 3, LAST_ROW, MAX_TRIALS).setHorizontalAlignment('center');

  // 味：10前後(ちょうど良い)が白。足りないと青、強すぎると赤に寄る
  var tasteRule = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpointWithValue('#8AAFCC', SpreadsheetApp.InterpolationType.NUMBER, '1')
    .setGradientMidpointWithValue('#FFFFFF', SpreadsheetApp.InterpolationType.NUMBER, '10')
    .setGradientMaxpointWithValue('#D98E6A', SpreadsheetApp.InterpolationType.NUMBER, '20')
    .setRanges([sh.getRange(5, 3, 6, MAX_TRIALS)]).build();
  // 総合評価：こちらは大きいほど良いので、高いほど緑が濃くなる
  var overall = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpointWithValue('#FFFFFF', SpreadsheetApp.InterpolationType.NUMBER, '1')
    .setGradientMaxpointWithValue('#8FB98C', SpreadsheetApp.InterpolationType.NUMBER, '20')
    .setRanges([sh.getRange(11, 3, 1, MAX_TRIALS)]).build();
  sh.setConditionalFormatRules([tasteRule, overall]);

  sh.setFrozenRows(3);
  sh.setFrozenColumns(2);
  sh.getRange('A1').setNote('B1で対象を選ぶと、その試作の全回が横に並びます。上が味、下が分量です。\n\n※ この画面は「見るだけ」です。数字を直接打ち込むと自動計算が壊れます。\n　 分量は「試作レシピ明細」、味は「試作記録」に入れてください。');

  // 誤って打ち込むと数式が壊れるので警告を出す。
  // ただし B1（対象を選ぶセル）は毎回使うので、そこだけは警告なしで触れるようにする。
  var msg = '試作比較は見るだけの画面です（入力は 試作レシピ明細 と 試作記録 へ）';
  [sh.getRange('A1'), sh.getRange('C1:Z1'), sh.getRange(2, 1, LAST_ROW, 26)].forEach(function (r) {
    r.protect().setDescription(msg).setWarningOnly(true);
  });
}

function vl_(col, index) {
  return '=IF(' + col + '$2="","",IFERROR(VLOOKUP(' + col + '$2,試作記録!$A:$O,' + index + ',FALSE),""))';
}

/** 中身が全部数式のシートを、作り直せる状態に戻す */
function resetSheet_(sh) {
  sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function (p) { p.remove(); });
  sh.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function (p) { p.remove(); });
  sh.setConditionalFormatRules([]);
  sh.showRows(1, sh.getMaxRows());
  sh.showColumns(1, sh.getMaxColumns());
  sh.setFrozenRows(0);
  sh.setFrozenColumns(0);
  var all = sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns());
  all.clearDataValidations();
  all.clearNote();
  sh.clear();
}

function columnLetter_(n) {
  var s = '';
  while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; }
  return s;
}


/* ---------- 次の回の試作をつくる（前回の分量をコピー） ---------- */
function 次の回の試作をつくる() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var cmp = ss.getSheetByName('試作比較');
  var target = cmp ? String(cmp.getRange('B1').getValue()).trim() : '';
  if (!target) {
    var res = ui.prompt('次の回の試作をつくる', '対象のID(MENU-0001 や PRE-0001)を入力してください。', ui.ButtonSet.OK_CANCEL);
    if (res.getSelectedButton() !== ui.Button.OK) return;
    target = res.getResponseText().trim();
  }
  if (!target) return;

  var recSh = ss.getSheetByName('試作記録');
  var detSh = ss.getSheetByName('試作レシピ明細');
  var rec = recSh.getDataRange().getValues();
  var det = detSh.getDataRange().getValues();

  // 対象の最新回を探す
  var lastNo = 0, lastId = '';
  for (var i = 1; i < rec.length; i++) {
    if (String(rec[i][1]).trim() === target && Number(rec[i][3]) >= lastNo) {
      lastNo = Number(rec[i][3]);
      lastId = String(rec[i][0]).trim();
    }
  }

  var newNo = lastNo + 1;
  var newId = 次の試作ID_(rec);

  // 前回の分量をコピー
  var copied = [];
  for (var j = 1; j < det.length; j++) {
    if (lastId && String(det[j][0]).trim() === lastId) {
      copied.push([newId, det[j][1], det[j][3]]);
    }
  }

  var msg = '対象: ' + target + '\n' +
            '新しく作る回: 第' + newNo + '回 (' + newId + ')\n\n' +
            (copied.length > 0
              ? '第' + lastNo + '回の分量 ' + copied.length + '件をコピーします。'
              : '前の回が無いので、分量は空で作ります。') +
            '\n\nよろしいですか?';
  if (ui.alert('確認', msg, ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;

  var r = recSh.getLastRow() + 1;
  recSh.getRange(r, 1, 1, 2).setValues([[newId, target]]);
  recSh.getRange(r, 4, 1, 2).setValues([[newNo, new Date()]]);

  if (copied.length > 0) {
    var d = detSh.getLastRow() + 1;
    detSh.getRange(d, 1, copied.length, 2).setValues(copied.map(function (x) { return [x[0], x[1]]; }));
    detSh.getRange(d, 4, copied.length, 1).setValues(copied.map(function (x) { return [x[2]]; }));
  }

  if (cmp) { cmp.getRange('B1').setValue(target); ss.setActiveSheet(cmp); }

  ui.alert('第' + newNo + '回 (' + newId + ') を作りました。\n\n' +
           '試作レシピ明細で分量を調整し、作って食べたら試作記録に5味を入れてください。');
}

function 次の試作ID_(rec) {
  var max = 0;
  for (var i = 1; i < rec.length; i++) {
    var m = String(rec[i][0]).match(/^TRY-(\d+)$/);
    if (m) max = Math.max(max, Number(m[1]));
  }
  return 'TRY-' + ('0000' + (max + 1)).slice(-4);
}


/* ---------- 採用した試作を本番レシピへ反映 ---------- */
function 採用試作をレシピへ反映() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var recSh = ss.getSheetByName('試作記録');
  var detSh = ss.getSheetByName('試作レシピ明細');
  var repSh = ss.getSheetByName('レシピ明細');
  if (!recSh || !detSh || !repSh) { ui.alert('シートが見つかりません。'); return; }

  var rec = recSh.getDataRange().getValues();
  var picks = [];
  for (var i = 1; i < rec.length; i++) {
    if (String(rec[i][14]).trim() === '○') {        // O列＝採用
      picks.push({ row: i + 1, tryId: String(rec[i][0]).trim(), target: String(rec[i][1]).trim(), no: rec[i][3] });
    }
  }
  if (picks.length === 0) {
    ui.alert('採用(○)が付いた試作がありません。\n\n試作記録の「採用」列に ○ を付けてから実行してください。');
    return;
  }

  // 同じ対象に○が2つ以上ある場合は、どちらを採るか決められないので中止
  var seen = {};
  for (var p = 0; p < picks.length; p++) {
    if (seen[picks[p].target]) {
      ui.alert('中止しました。\n\n対象 ' + picks[p].target + ' に採用(○)が複数付いています。\n1つだけ残してください。');
      return;
    }
    seen[picks[p].target] = true;
  }

  // 反映先のIDが実在するか確認
  var ids = {};
  [['商品マスタ'], ['仕込み品マスタ']].forEach(function (x) {
    var s = ss.getSheetByName(x[0]);
    s.getRange(2, 1, Math.max(s.getLastRow() - 1, 1), 1).getValues().forEach(function (v) {
      if (v[0]) ids[String(v[0]).trim()] = true;
    });
  });
  var bad = picks.filter(function (q) { return !ids[q.target]; });
  if (bad.length > 0) {
    ui.alert('中止しました。\n\n次の対象IDが商品マスタ・仕込み品マスタに見つかりません:\n  ' +
             bad.map(function (q) { return q.target; }).join('\n  ') +
             '\n\n先にマスタへ登録してください。');
    return;
  }

  var det = detSh.getDataRange().getValues();
  var lines = {};
  picks.forEach(function (q) { lines[q.tryId] = []; });
  for (var j = 1; j < det.length; j++) {
    var t = String(det[j][0]).trim();
    if (lines[t]) lines[t].push([q_(det[j][1]), det[j][3], q_(det[j][7])]);
  }

  var summary = picks.map(function (q) {
    return '  ' + q.target + ' ← 第' + q.no + '回 (' + q.tryId + ') / ' + lines[q.tryId].length + '件';
  }).join('\n');
  if (ui.alert('本番レシピへ反映します',
      summary + '\n\n対象の既存レシピは、この内容で置き換わります。\nよろしいですか?',
      ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;

  // 行は削除せず、入力列(A/C/E/J)だけ組み直す。
  // ※ B/D/F/G/H/I 列の2行目には自動計算の数式が入っているため、行削除は厳禁。
  var last = Math.max(repSh.getLastRow(), 2);
  var cur = repSh.getRange(2, 1, last - 1, 10).getValues()
    .map(function (r) { return [q_(r[0]), q_(r[2]), r[4], q_(r[9])]; })
    .filter(function (r) { return r[0] !== ''; });

  picks.forEach(function (q) {
    var fresh = lines[q.tryId].map(function (L) { return [q.target, L[0], L[1], L[2]]; });
    var out = [], done = false;
    cur.forEach(function (r) {
      if (r[0] === q.target) {
        if (!done) { out = out.concat(fresh); done = true; }
      } else { out.push(r); }
    });
    if (!done) out = out.concat(fresh);
    cur = out;
  });

  repSh.getRange(2, 1, last - 1, 1).clearContent();
  repSh.getRange(2, 3, last - 1, 1).clearContent();
  repSh.getRange(2, 5, last - 1, 1).clearContent();
  repSh.getRange(2, 10, last - 1, 1).clearContent();

  if (cur.length > 0) {
    var need = cur.length + 1 - repSh.getMaxRows();
    if (need > 0) repSh.insertRowsAfter(repSh.getMaxRows(), need);
    repSh.getRange(2, 1,  cur.length, 1).setValues(cur.map(function (r) { return [r[0]]; }));
    repSh.getRange(2, 3,  cur.length, 1).setValues(cur.map(function (r) { return [r[1]]; }));
    repSh.getRange(2, 5,  cur.length, 1).setValues(cur.map(function (r) { return [r[2]]; }));
    repSh.getRange(2, 10, cur.length, 1).setValues(cur.map(function (r) { return [r[3]]; }));
  }

  var stamp = Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd');
  picks.forEach(function (q) { recSh.getRange(q.row, 15).setValue('反映済 ' + stamp); });

  ss.setActiveSheet(repSh);
  ui.alert('反映しました。\n\n' + summary + '\n\n商品マスタの原価が更新されているか確認してください。');
}

function q_(v) { return v === null || v === undefined ? '' : String(v).trim(); }

