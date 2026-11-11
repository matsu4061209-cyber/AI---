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
var TRIAL_SHEETS = ['試作記録', '試作レシピ明細', '試作手順', '試作比較'];
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

/* =====================================================================
 *  食材の「味の強さ」＝ 1g がその味をどれだけ出すか（甘・塩・酸・旨・苦・辛）
 *
 *  ・塩味／甘味は「塩分%／糖分%」が目安（計算で比較的正確に出せる）
 *  ・酸味・旨味・苦味・辛味は定量化が難しいので「目安」の経験値
 *  ・この値から試作の味の初期値（スライダー）を推定します。あくまでたたき台で、
 *    食べた感覚と違うところは人がスライダーで直す前提です。
 *  ・名前で引くので、店長が同じ名前の食材を足せば自動で初期値が入ります。
 * ===================================================================== */
var TASTE_KEYS = ['甘味度', '塩味度', '酸味度', '旨味度', '苦味度', '辛味度'];

// 「味の強さ」を推定するときの基準の濃さ（この濃さで ちょうど良い=10 になる）
//   メーター値 = 10 × (料理全体のその味の平均% ÷ 基準)
var TASTE_BASE = { sweet: 6, salt: 0.9, sour: 1.2, umami: 5, bitter: 1.5, hot: 1.2 };

// 食材名 → [甘, 塩, 酸, 旨, 苦, 辛]（1gあたりの味ポイント）
var TASTE_DEFAULTS = {
  // --- 甘味料 ---
  '砂糖': [100, 0, 0, 0, 0, 0], '上白糖': [100, 0, 0, 0, 0, 0], 'グラニュー糖': [100, 0, 0, 0, 0, 0],
  'みりん': [43, 0, 0, 2, 0, 0], '本みりん': [43, 0, 0, 2, 0, 0], 'はちみつ': [80, 0, 1, 1, 0, 0],
  // --- 塩・しょうゆ・みそ ---
  '塩': [0, 99, 0, 0, 0, 0], '食塩': [0, 99, 0, 0, 0, 0],
  '醤油': [10, 16, 1, 15, 1, 0], 'しょうゆ': [10, 16, 1, 15, 1, 0], '濃口醤油': [10, 16, 1, 15, 1, 0],
  '薄口醤油': [8, 18, 1, 13, 1, 0], '味噌': [10, 12, 1, 14, 1, 0], 'みそ': [10, 12, 1, 14, 1, 0],
  // --- 酒・酢 ---
  '料理酒': [4, 2, 1, 3, 0, 0], '酒': [2, 0, 1, 2, 0, 0], '日本酒': [2, 0, 1, 2, 0, 0],
  '酢': [1, 0, 4.5, 1, 0, 0], '米酢': [1, 0, 4.5, 1, 0, 0], '穀物酢': [1, 0, 4, 1, 0, 0],
  'レモン汁': [2, 0, 6, 0, 1, 0], 'レモン': [2, 0, 6, 0, 1, 0],
  // --- だし・うま味 ---
  'だし': [2, 1, 0, 10, 0, 0], 'だしの素': [5, 40, 0, 25, 0, 0], '和風だし': [5, 40, 0, 25, 0, 0],
  '鶏がらスープの素': [3, 45, 0, 22, 0, 0], 'コンソメ': [4, 43, 0, 20, 0, 0],
  'オイスターソース': [18, 10, 1, 18, 1, 0],
  // --- ソース・たれ類 ---
  'ケチャップ': [25, 3, 8, 6, 0, 0], 'ソース': [20, 6, 5, 6, 1, 1], '中濃ソース': [20, 6, 5, 6, 1, 1],
  'ウスターソース': [18, 8, 6, 6, 1, 2], 'マヨネーズ': [2, 2, 6, 3, 0, 1],
  // --- 香味・辛味 ---
  'にんにく': [2, 0, 0, 5, 1, 4], 'おろしにんにく': [2, 0, 0, 5, 1, 4],
  '生姜': [1, 0, 0, 1, 2, 6], 'しょうが': [1, 0, 0, 1, 2, 6], 'おろし生姜': [1, 0, 0, 1, 2, 6],
  '唐辛子': [1, 0, 0, 1, 2, 18], '一味唐辛子': [1, 0, 0, 1, 2, 18], '七味唐辛子': [1, 1, 0, 1, 2, 14],
  '豆板醤': [2, 10, 2, 6, 1, 12], 'コチュジャン': [20, 8, 2, 8, 1, 10],
  'こしょう': [0, 0, 0, 1, 3, 7], '黒こしょう': [0, 0, 0, 1, 3, 7], 'ごま油': [1, 0, 0, 3, 1, 0],
  'バター': [1, 2, 0, 3, 0, 0], 'サラダ油': [0, 0, 0, 0, 0, 0],
  // --- 素材 ---
  '豚ロース': [0, 0.1, 0, 8, 0, 0], '豚バラ': [0, 0.1, 0, 8, 0, 0], '鶏もも': [0, 0.1, 0, 8, 0, 0],
  '牛肉': [0, 0.1, 0, 9, 0, 0], '米': [0.5, 0, 0, 1, 0, 0], '米(精白)': [0.5, 0, 0, 1, 0, 0],
  'キャベツ': [3, 0, 0, 1, 0.5, 0], '玉ねぎ': [4, 0, 0, 2, 0, 1], 'トマト': [3, 0, 1, 4, 0, 0],
  '豆腐': [0.5, 0, 0, 2, 0, 0],
  'ラウンドオイスター': [2, 1, 0, 12, 1, 0], '牡蠣': [2, 1, 0, 12, 1, 0]
};

/** 食材名から味の初期値 [甘,塩,酸,旨,苦,辛] を引く。無ければ空6個（表記ゆれに強い照合）。 */
function tasteDefaultFor_(name) {
  var hit = TASTE_DEFAULTS[String(name || '').trim()];
  if (hit) return hit.slice();
  var n = normName_(name);
  for (var k in TASTE_DEFAULTS) {
    if (TASTE_DEFAULTS.hasOwnProperty(k) && normName_(k) === n) return TASTE_DEFAULTS[k].slice();
  }
  return ['', '', '', '', '', ''];
}

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
  build試作手順(ss);
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

  applyAutoFormulas_(sh);   // 数式は AUTO_FORMULAS にまとめてあります

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
    '単価(自動)', '実効単価(自動)',
    TASTE_KEYS[0], TASTE_KEYS[1], TASTE_KEYS[2], TASTE_KEYS[3], TASTE_KEYS[4], TASTE_KEYS[5]
  ];
  var widths = [90, 160, 90, 80, 80, 90, 90, 100, 120, 90, 80, 90, 120, 110, 90, 90, 110, 120,
                64, 64, 64, 64, 64, 64];
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

  // 味の初期値（甘塩酸旨苦辛）＝ 食材名から引く。たたき台なので後から直せる
  sh.getRange(2, 19, rows.length, 6).setValues(rows.map(function (r) { return tasteDefaultFor_(r[1]); }));

  applyAutoFormulas_(sh);   // 数式は AUTO_FORMULAS にまとめてあります

  addList_(sh, 3,  OPT_分類);
  addList_(sh, 4,  OPT_在庫単位);
  addList_(sh, 5,  OPT_保存区分);
  addList_(sh, 10, OPT_発注単位);
  linkSupplier_(ss, sh);

  sh.getRange(2, 7,  LAST_ROW, 1).setNumberFormat('0%');
  sh.getRange(2, 12, LAST_ROW, 2).setNumberFormat('¥#,##0');
  sh.getRange(2, 17, LAST_ROW, 2).setNumberFormat('¥#,##0.0000');
  sh.getRange(2, 19, LAST_ROW, 6).setNumberFormat('0.#').setHorizontalAlignment('center');
  markAuto_(sh, 17, 2);
  tasteBorders_(sh, 19, 23);   // 甘塩酸旨＝実線／苦辛＝破線

  // 発注連結用ゾーンの目印（仕入先〜実測補正）
  sh.getRange(1, 8, 1, 9).setBackground('#EADFCB');
}


/* ========== ② 仕込み品マスタ（自店で作るもの） ========== */
function build仕込み品マスタ(ss) {
  var sh = newSheet_(ss, '仕込み品マスタ');

  // 原価合計・単位原価・チェック（自動）は位置を変えない。昔の基準表の項目はその右に足す。
  // ※「出来上がり量」が昔の「仕上がり総量」にあたるので、重複して足さない。
  var headers = ['仕込みID', '仕込み品名', '出来上がり量', '単位', '保存区分', '日持ち日数',
                 '原価合計(自動)', '単位原価(自動)', 'チェック(自動)',
                 '仕込み時間', '製作者', '制作年月日'];
  var widths  = [90, 180, 110, 70, 80, 90, 120, 130, 220,
                 100, 100, 110];
  setHeader_(sh, headers, widths);

  var rows = [
    ['PRE-0001', '生姜だれ', 1000, 'g', '冷蔵', 7]
  ];
  sh.getRange(2, 1, rows.length, 6).setValues(rows);

  applyAutoFormulas_(sh);   // 数式は AUTO_FORMULAS にまとめてあります

  addList_(sh, 4, OPT_在庫単位);
  addList_(sh, 5, OPT_保存区分);

  sh.getRange(2, 7, LAST_ROW, 1).setNumberFormat('¥#,##0.0');
  sh.getRange(2, 8, LAST_ROW, 1).setNumberFormat('¥#,##0.0000');
  sh.getRange(2, 12, LAST_ROW, 1).setNumberFormat('yyyy/mm/dd');   // 制作年月日
  markAuto_(sh, 7, 3);
}


/* ========== ③ 商品マスタ（売るもの） ========== */
function build商品マスタ(ss) {
  var sh = newSheet_(ss, '商品マスタ');

  // 原価・原価率・目標差（自動）は位置を変えない＝原価計算・採用・登録フォームを壊さない。
  // 昔の「メニュー基準表」にあった項目（商品コード〜制作年月日）は、その右に足す。
  var headers = ['商品ID', '商品名', 'カテゴリ', '売価(税抜)', '目標原価率', '提供開始日', '提供終了日',
                 '原価(自動)', '原価率(自動)', '目標差(自動)',
                 '商品コード', '仕上がり量', '仕込み時間', '製作者', '制作年月日'];
  var widths  = [100, 220, 100, 110, 110, 110, 110, 110, 110, 110,
                 120, 100, 100, 100, 110];
  setHeader_(sh, headers, widths);

  var rows = [
    ['MENU-0001', 'サンプル 生姜焼き定食', '定食', 980, 0.35, '', '']
  ];
  sh.getRange(2, 1, rows.length, 7).setValues(rows);

  applyAutoFormulas_(sh);   // 数式は AUTO_FORMULAS にまとめてあります

  sh.getRange(2, 4,  LAST_ROW, 1).setNumberFormat('¥#,##0');
  sh.getRange(2, 5,  LAST_ROW, 1).setNumberFormat('0.0%');
  sh.getRange(2, 6,  LAST_ROW, 2).setNumberFormat('yyyy/mm/dd');
  sh.getRange(2, 8,  LAST_ROW, 1).setNumberFormat('¥#,##0.0');
  sh.getRange(2, 9,  LAST_ROW, 1).setNumberFormat('0.0%');
  sh.getRange(2, 10, LAST_ROW, 1).setNumberFormat('+0.0%;-0.0%;0.0%');
  sh.getRange(2, 15, LAST_ROW, 1).setNumberFormat('yyyy/mm/dd');   // 制作年月日
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

  applyAutoFormulas_(sh);   // 数式は AUTO_FORMULAS にまとめてあります

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

  applyAutoFormulas_(sh);   // 数式は AUTO_FORMULAS にまとめてあります

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
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('🍳 レシピ')
    .addItem('メニューを登録・修正する', 'メニュー登録フォームを開く')
    .addItem('レシピを見る（カード）', 'レシピを見る')
    .addSeparator()
    .addItem('台帳を整える（色分け・保護・入力ガイド）', '台帳を整える')
    .addSeparator()
    .addItem('サンプルのご飯を仕込み品にする（1回だけ）', 'ご飯を仕込み品にする')
    .addItem('白だしを取り込む（1回だけ）', '白だしを取り込む')
    .addItem('柚子胡椒・柚子ポン酢を取り込む（1回だけ）', '柚子胡椒と柚子ポン酢を取り込む')
    .addToUi();

  ui.createMenu('🧪 試作')
    .addItem('試作をはじめる / 次の回をつくる', '次の回の試作をつくる')
    .addItem('味を記録・調整する', '味を記録調整する')
    .addItem('採用した試作をレシピへ反映', '採用試作をレシピへ反映')
    .addItem('試作を取り消す（間違えて作った回を消す）', '試作を取り消す')
    .addSeparator()
    .addItem('試作シートを追加(初回のみ)', '試作シートを追加')
    .addItem('不定貫の列を追加(初回のみ)', '不定貫の列を追加')
    .addItem('発注準備の列を追加(初回のみ)', '発注準備の列を追加')
    .addItem('辛味の列を追加(初回のみ)', '辛味の列を追加')
    .addItem('味の列を追加(初回のみ)', '味の列を追加')
    .addItem('基準表の項目を追加(初回のみ)', '基準表の項目を追加')
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
 * 既にある台帳の食材マスタに「味の列（甘・塩・酸・旨・苦・辛）」を6本足す。
 *   ・実効単価の右に足すので、原価のVLOOKUP（$A:$R の18列目）は動かない＝原価は変わらない
 *   ・よく使う調味料には名前から初期値を入れる。入力済みの味はそのまま残す
 */
function 味の列を追加() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var ing = ss.getSheetByName('食材マスタ');
  if (!ing) { ui.alert('食材マスタが見つかりません。'); return; }

  if (String(ing.getRange(1, 19).getValue()).indexOf('甘味度') === 0) {
    ui.alert('すでに追加済みです。\n\n食材マスタのS列が「甘味度」になっています。');
    return;
  }
  // 実効単価(R=18)の右に足す前提。ここがズレていると原価計算に関わるので止める。
  if (String(ing.getRange(1, 18).getValue()).indexOf('実効単価') !== 0) {
    ui.alert('中止しました。\n\n食材マスタのR列が想定(実効単価(自動))と違います。\n' +
             '先に「発注準備の列を追加」を実行してください。\n' +
             '（列を手で動かしている場合は、先に元に戻してください）');
    return;
  }

  if (ui.alert('確認',
      '食材マスタの「実効単価」の右に、味の列を6本足します。\n' +
      '（甘味度・塩味度・酸味度・旨味度・苦味度・辛味度）\n\n' +
      'よく使う調味料には初期値を入れます。\n' +
      '入力済みのデータ・原価は変わりません。よろしいですか?',
      ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;

  ing.insertColumnsAfter(18, 6);                          // R(実効単価) の右に6本
  for (var i = 0; i < 6; i++) ing.setColumnWidth(19 + i, 64);
  ing.getRange(1, 19, 1, 6).setValues([TASTE_KEYS])
    .setFontWeight('bold').setBackground(C_HEADER_BG).setFontColor('#2A2019')
    .setVerticalAlignment('middle');

  // 既存の食材に、名前から味の初期値を入れる（すでに値がある行はそのまま残す）
  var last = lastDataRow_(ing);
  if (last >= 2) {
    var names = ing.getRange(2, 2,  last - 1, 1).getValues();   // B列 = 食材名
    var cur   = ing.getRange(2, 19, last - 1, 6).getValues();   // 今の味（ふつう空）
    var out = names.map(function (nm, r) {
      var def = tasteDefaultFor_(nm[0]);
      return cur[r].map(function (v, c) { return (v === '' || v == null) ? def[c] : v; });
    });
    ing.getRange(2, 19, last - 1, 6).setValues(out);
  }

  ing.getRange(2, 19, LAST_ROW, 6).setNumberFormat('0.#').setHorizontalAlignment('center');
  tasteBorders_(ing, 19, 23);                             // 甘塩酸旨＝実線／苦辛＝破線

  ss.setActiveSheet(ing);
  ui.alert('味の列を追加しました。\n\n' +
           '・食材マスタの右に 甘味度〜辛味度 の6列が増えました\n' +
           '・よく使う調味料には初期値が入っています（空欄＝0 として扱います）\n' +
           '・原価は変わっていないはずです（実効単価の右に足したため）\n\n' +
           '仕上げに「🍳 レシピ → 台帳を整える」を1回押すと、色分けと説明が付きます。');
}


/* =====================================================================
 *  昔の「メニュー基準表」にあって台帳に無かった項目を足す。
 *   ・商品マスタ   … 商品コード / 仕上がり量 / 仕込み時間 / 製作者 / 制作年月日（11〜15列）
 *   ・仕込み品マスタ … 仕込み時間 / 製作者 / 制作年月日（10〜12列。仕上がり総量は「出来上がり量」が兼ねる）
 *
 *  ★計算に関わる列（原価・原価率・目標差・単位原価…）は一切動かさない＝原価は変わらない。
 *   これらは自動計算列の「右の空き列」に名前を付けるだけ（列の挿入はしない）。
 *   新規に作った台帳では build〜 が最初から入れるので、この関数は昔の台帳用。
 * ===================================================================== */

/** 商品マスタ・仕込み品マスタに基準表の項目があることを保証する（何度呼んでも安全）。足した数を返す */
function ensure基準表Cols_(ss) {
  var added = 0;

  var pm = ss.getSheetByName('商品マスタ');
  if (pm) {
    ensureCols_(pm, 15);
    // 8列目が「原価(自動)」＝想定どおりの並びのときだけ、その右(11〜15)に名前を付ける
    if (String(pm.getRange(1, 8).getValue()).indexOf('原価') === 0) {
      var pmHeads = [['商品コード', 120], ['仕上がり量', 100], ['仕込み時間', 100], ['製作者', 100], ['制作年月日', 110]];
      pmHeads.forEach(function (h, i) {
        var col = 11 + i;
        if (stripMark_(pm.getRange(1, col).getValue()) !== h[0]) {
          pm.getRange(1, col).setValue(h[0])
            .setFontWeight('bold').setBackground(C_HEADER_BG).setFontColor('#2A2019').setVerticalAlignment('middle');
          pm.setColumnWidth(col, h[1]);
          added++;
        }
      });
      pm.getRange(2, 15, LAST_ROW, 1).setNumberFormat('yyyy/mm/dd');   // 制作年月日
    }
  }

  var prep = ss.getSheetByName('仕込み品マスタ');
  if (prep) {
    ensureCols_(prep, 12);
    // 9列目が「チェック(自動)」＝想定どおりのときだけ、その右(10〜12)に名前を付ける
    if (String(prep.getRange(1, 9).getValue()).indexOf('チェック') === 0) {
      var prHeads = [['仕込み時間', 100], ['製作者', 100], ['制作年月日', 110]];
      prHeads.forEach(function (h, i) {
        var col = 10 + i;
        if (stripMark_(prep.getRange(1, col).getValue()) !== h[0]) {
          prep.getRange(1, col).setValue(h[0])
            .setFontWeight('bold').setBackground(C_HEADER_BG).setFontColor('#2A2019').setVerticalAlignment('middle');
          prep.setColumnWidth(col, h[1]);
          added++;
        }
      });
      prep.getRange(2, 12, LAST_ROW, 1).setNumberFormat('yyyy/mm/dd');   // 制作年月日
    }
  }

  return added;
}

/** メニュー：昔の基準表の項目（商品コード・仕上がり量・仕込み時間・製作者・制作年月日）を足す */
function 基準表の項目を追加() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  if (!ss.getSheetByName('商品マスタ') || !ss.getSheetByName('仕込み品マスタ')) {
    ui.alert('商品マスタ・仕込み品マスタが見つかりません。\n\n先に台帳を作ってください。');
    return;
  }

  // 並びが想定どおりか（計算列を手で動かしていないか）を確認
  if (String(ss.getSheetByName('商品マスタ').getRange(1, 8).getValue()).indexOf('原価') !== 0 ||
      String(ss.getSheetByName('仕込み品マスタ').getRange(1, 9).getValue()).indexOf('チェック') !== 0) {
    ui.alert('中止しました。\n\n商品マスタ／仕込み品マスタの並びが想定と違います。\n' +
             '（商品マスタの8列目＝原価、仕込み品マスタの9列目＝チェック のはずです）\n' +
             '列を手で動かしている場合は、先に元に戻してください。');
    return;
  }

  if (ui.alert('確認',
      '昔の「メニュー基準表」にあった項目を足します（覚え書き欄）。\n\n' +
      '・商品マスタ … 商品コード / 仕上がり量 / 仕込み時間 / 製作者 / 制作年月日\n' +
      '・仕込み品マスタ … 仕込み時間 / 製作者 / 制作年月日\n' +
      '　（仕上がり総量は、今ある「出来上がり量」がそれにあたります）\n\n' +
      '★自動計算の列（原価など）の右の空き列に名前を付けるだけです。\n' +
      '　原価・売価・原価率は変わりません。入力済みのデータも消えません。\n\n' +
      'よろしいですか?',
      ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;

  var added = ensure基準表Cols_(ss);

  ss.setActiveSheet(ss.getSheetByName('商品マスタ'));
  ui.alert((added > 0 ? '追加しました（' + added + '項目）。\n\n' : 'すでに揃っていました。\n\n') +
           '・商品マスタの右に 商品コード〜制作年月日\n' +
           '・仕込み品マスタの右に 仕込み時間〜制作年月日\n\n' +
           '仕上げに「🍳 レシピ → 台帳を整える」を1回押すと、色分けと説明が付きます。');
}


/* =====================================================================
 *  昔の基準表の「白だし」を、仕込み品として取り込む（1回だけ）。
 *   ・足りない食材（水・昆布・鰹節・日本酒・塩・薄口醤油）を食材マスタに追加
 *     （仕入価格は“だいたいの相場”の仮入力。あとで実際の値に直してください）
 *   ・仕込み品マスタに「白だし」を追加（出来上がり1000cc・冷蔵・21日・作者/日付/時間つき）
 *   ・レシピ明細に材料7つ、調理手順に作り方9手順を入れる（rebuildBlock_ なので何度実行しても二重にならない）
 * ===================================================================== */
/**
 * 食材マスタに食材が無ければ1つ追加して、そのIDを返す（あればそのID）。
 *   spec = [名前, 分類, 在庫単位, 保存区分, 日持ち日数, 発注単位, 入数, 仕入価格(相場・仮)]
 *   idByName … 名前→IDの辞書。追加したら更新する（呼び出し側と共有）。
 */
function addIngredientIfMissing_(ing, idByName, spec, hasTaste) {
  var name = String(spec[0]).trim();
  if (idByName[name]) return idByName[name];
  var id = 'ING-' + zero4_(nextIdNumber_(ing, ID_RULES['食材マスタ']));
  var row = lastDataRow_(ing) + 1;
  // 1:ID 2:名 3:分類 4:在庫単位 5:保存区分 6:日持ち 7:歩留まり率 8:仕入先ID 9:商品コード 10:発注単位 11:入数 12:仕入価格
  ing.getRange(row, 1, 1, 12).setValues([[id, name, spec[1], spec[2], spec[3], spec[4], 1, '', '', spec[5], spec[6], spec[7]]]);
  if (hasTaste) ing.getRange(row, 19, 1, 6).setValues([tasteDefaultFor_(name)]);
  idByName[name] = id;
  SpreadsheetApp.flush();                  // 次のID採番がかぶらないように確定
  return id;
}

/**
 * 仕込み品を1つ台帳に取り込む共通処理（白だし・柚子だれ 等はこれを呼ぶだけ）。
 * 何度呼んでも二重にならない（名前で探して無ければ作る／材料・作り方は rebuildBlock_）。
 *   spec = { name, yield_(空='材料合計'), unit, save, keep, time, author, made(Date|''),
 *            ingredients:[ [名,分類,単位,保存,日持ち,発注単位,入数,仕入価格], ... ],
 *            mats:[ {name, qty, memo} ... ],  steps:[ '手順文', ... ] }
 * 返り値 = その仕込み品のID（PRE-xxxx）。
 */
function importPrep_(ss, spec) {
  var ing  = ss.getSheetByName('食材マスタ');
  var prep = ss.getSheetByName('仕込み品マスタ');
  var det  = ss.getSheetByName('レシピ明細');
  var proc = ss.getSheetByName('調理手順');
  if (!ing || !prep || !det || !proc) throw new Error('必要なシート（食材マスタ・仕込み品マスタ・レシピ明細・調理手順）が見つかりません。');

  ensure基準表Cols_(ss);   // 製作者・制作年月日・仕込み時間 の置き場所を用意

  var hasTaste = (String(ing.getRange(1, 19).getValue()).indexOf('甘味度') === 0);
  var idByName = {};
  readRows_(ss, '食材マスタ', 2).forEach(function (r) { idByName[String(r[1]).trim()] = String(r[0]).trim(); });

  // ① 足りない食材を用意
  (spec.ingredients || []).forEach(function (s) { addIngredientIfMissing_(ing, idByName, s, hasTaste); });

  // ② 仕込み品本体（名前で探して無ければ追加）
  var preId = '';
  readRows_(ss, '仕込み品マスタ', 2).forEach(function (r) { if (String(r[1]).trim() === spec.name && !preId) preId = String(r[0]).trim(); });
  var yieldVal = (spec.yield_ === '' || spec.yield_ == null) ? '' : spec.yield_;
  if (!preId) {
    preId = 'PRE-' + zero4_(nextIdNumber_(prep, ID_RULES['仕込み品マスタ']));
    var prow = lastDataRow_(prep) + 1;
    prep.getRange(prow, 1, 1, 6).setValues([[preId, spec.name, yieldVal, spec.unit, spec.save, spec.keep]]);
  } else {
    var frow = findRowById_(prep, preId);
    if (frow) prep.getRange(frow, 3, 1, 4).setValues([[yieldVal, spec.unit, spec.save, spec.keep]]);
  }
  var mrow = findRowById_(prep, preId);   // 基準表の覚え書き（仕込み時間=10 / 製作者=11 / 制作年月日=12）
  if (mrow) prep.getRange(mrow, 10, 1, 3).setValues([[spec.time || '', spec.author || '', spec.made || '']]);

  // ③ 材料（レシピ明細）。入力列 = 親ID(1)/子ID(3)/使用量(5)/備考(10)
  var mats = (spec.mats || []).map(function (m) { return { id: idByName[String(m.name).trim()] || '', qty: m.qty, memo: m.memo || '' }; })
                              .filter(function (m) { return m.id; });
  rebuildBlock_(det, [1, 3, 5, 10], [{ key: preId, rows: mats.map(function (m) { return [preId, m.id, m.qty, m.memo]; }) }]);

  // ④ 作り方（調理手順）。入力列 = 商品ID(1)/手順番号(3)/手順内容(4)/所要時間(5)/温度火加減(6)/ポイント(7)
  rebuildBlock_(proc, [1, 3, 4, 5, 6, 7], [{ key: preId, rows: (spec.steps || []).map(function (t, i) { return [preId, i + 1, t, '', '', '']; }) }]);

  SpreadsheetApp.flush();
  return preId;
}


function 白だしを取り込む() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  if (!ss.getSheetByName('食材マスタ')) { ui.alert('台帳が見つかりません。'); return; }

  if (ui.alert('白だしを取り込みますか?',
      '昔の基準表の「白だし」を、仕込み品として取り込みます。\n\n' +
      '・足りない食材（水・昆布・鰹節・日本酒・塩・薄口醤油）を食材マスタに足します\n' +
      '　※仕入価格は “だいたいの相場” の仮入力です。あとで実際の値に直してください\n' +
      '・仕込み品「白だし」＋材料7つ＋作り方9手順 を入れます\n\n' +
      '何度押しても二重にはなりません。よろしいですか?',
      ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;

  var preId = importPrep_(ss, {
    name: '白だし', yield_: 1000, unit: 'ml', save: '冷蔵', keep: 21,
    time: '40分', author: '津原', made: new Date(2013, 7, 31),
    ingredients: [
      ['水',       'ドリンク', 'ml', '常温', 999, '—', 1000,    0],
      ['昆布',     'その他',   'g',  '常温', 365, '袋',  1000, 2500],
      ['鰹節',     'その他',   'g',  '常温', 365, '袋',  1000, 3000],
      ['日本酒',   '調味料',   'ml', '常温', 365, '本',  1800, 1200],
      ['塩',       '調味料',   'g',  '常温', 365, '袋',  1000,  150],
      ['薄口醤油', '調味料',   'ml', '常温', 365, '本',  1800,  700]
    ],
    mats: [
      { name: '水',       qty: 400 },
      { name: '昆布',     qty: 10, memo: '水に浸けて一晩' },
      { name: '鰹節',     qty: 15 },
      { name: '日本酒',   qty: 100 },
      { name: 'みりん',   qty: 100 },
      { name: '塩',       qty: 20 },
      { name: '薄口醤油', qty: 60 }
    ],
    steps: [
      '昆布を分量の水に浸け、冷蔵庫で一晩（約1日）おく',
      '火にかけ、沸騰する直前に昆布を取り出す',
      '鰹節を入れる',
      '1〜2分おいて漉す',
      '日本酒・みりんを鍋に入れ、1〜2分煮る（煮切る）',
      '沸騰したまま塩を加えて混ぜる',
      '出汁と薄口醤油を加える',
      '軽く混ぜ、粗熱が取れたら完成',
      '冷蔵庫で2〜3週間ほど保存できる'
    ]
  });

  ss.setActiveSheet(ss.getSheetByName('仕込み品マスタ'));
  ui.alert('白だしを取り込みました（' + preId + '）。\n\n' +
           '・仕込み品マスタ … 白だし（出来上がり1000cc・冷蔵・約21日・製作者 津原）\n' +
           '・材料7つ、作り方9手順 を入れました\n\n' +
           '【要確認】\n' +
           '・昆布・鰹節・日本酒・塩・薄口醤油の仕入価格は “相場の仮” です。実際の値に直すと原価が正しくなります。\n' +
           '・出来上がり量は基準表どおり1000ccにしています（材料の合計は約705cc）。\n' +
           '　煮詰め方や水の量で変わるので、実際に作って量ったら直してください。\n\n' +
           '「🍳 レシピ → レシピを見る（カード）」で白だしを選ぶと、まとめて見られます。');
}


/**
 * 柚子胡椒・柚子ポン酢を、画像（料理人のレシピ）の分量で仕込み品として取り込む（1回だけ）。
 *   ・作り方は画像に無いので“一般的な作り方”を入れています（店長が食べて調整する前提）。
 *   ・新しい食材（柚子の皮・青唐辛子・柚子果汁・米酢）は相場の仮価格。みりん/醤油/昆布/鰹節/塩は既存を使う。
 */
function 柚子胡椒と柚子ポン酢を取り込む() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  if (!ss.getSheetByName('食材マスタ')) { ui.alert('台帳が見つかりません。'); return; }

  if (ui.alert('柚子胡椒・柚子ポン酢を取り込みますか?',
      '画像の分量で、2つの仕込み品を取り込みます。\n\n' +
      '・柚子胡椒（柚子の皮130g・青唐辛子130g・塩52g）\n' +
      '・柚子ポン酢（柚子果汁200・米酢70・みりん120・醤油300・鰹節4・昆布4）\n\n' +
      '※足りない食材（柚子の皮・青唐辛子・柚子果汁・米酢 ほか）は相場の仮価格で足します。\n' +
      '※作り方は画像に無いため“一般的な手順”を入れます（あとで直せます）。\n\n' +
      '何度押しても二重にはなりません。よろしいですか?',
      ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;

  var id1 = importPrep_(ss, {
    name: '柚子胡椒', yield_: '', unit: 'g', save: '冷蔵', keep: 180,
    time: '', author: '', made: '',
    ingredients: [
      ['柚子の皮', '野菜',   'g', '冷蔵', 14, '袋', 1000, 1500],
      ['青唐辛子', '野菜',   'g', '冷蔵', 14, '袋', 1000, 1500],
      ['塩',       '調味料', 'g', '常温', 365, '袋', 1000, 150]
    ],
    mats: [
      { name: '柚子の皮', qty: 130, memo: 'ゆず約15個の皮' },
      { name: '青唐辛子', qty: 130, memo: '丸170g→実130g' },
      { name: '塩',       qty: 52,  memo: '海塩・全体の約20%' }
    ],
    steps: [
      '柚子の皮を、白いワタを避けてすりおろす（または細かく刻む）',
      '青唐辛子はヘタと種を取り、みじん切りにする（手袋をする）',
      '皮・青唐辛子・塩を、すり鉢かフードプロセッサーでペースト状にする',
      '清潔な容器に入れ、冷蔵で数日なじませて完成'
    ]
  });

  var id2 = importPrep_(ss, {
    name: '柚子ポン酢', yield_: '', unit: 'ml', save: '冷蔵', keep: 90,
    time: '', author: '', made: '',
    ingredients: [
      ['柚子果汁', '調味料', 'ml', '冷蔵', 14,  '本', 1000, 2000],
      ['米酢',     '調味料', 'ml', '常温', 365, '本', 1000, 500],
      ['昆布',     'その他', 'g',  '常温', 365, '袋', 1000, 2500],
      ['鰹節',     'その他', 'g',  '常温', 365, '袋', 1000, 3000]
    ],
    mats: [
      { name: '柚子果汁', qty: 200, memo: 'ゆず15個ぶん' },
      { name: '米酢',     qty: 70 },
      { name: 'みりん',   qty: 120, memo: '本みりん' },
      { name: '醤油',     qty: 300, memo: '丸大豆醤油' },
      { name: '鰹節',     qty: 4 },
      { name: '昆布',     qty: 4,   memo: '5cm×2本' }
    ],
    steps: [
      '柚子を搾って果汁をとる（約200cc）',
      'すべての材料を合わせる（本みりんは軽く煮切るとよい）',
      '冷蔵で1〜2週間おき、昆布とかつお節を漉す',
      '漉したら冷蔵で保存して完成'
    ]
  });

  ss.setActiveSheet(ss.getSheetByName('仕込み品マスタ'));
  ui.alert('取り込みました。\n\n' +
           '・柚子胡椒（' + id1 + '）… 材料3つ・作り方4手順\n' +
           '・柚子ポン酢（' + id2 + '）… 材料6つ・作り方4手順\n\n' +
           '【要確認】\n' +
           '・柚子の皮・青唐辛子・柚子果汁・米酢 ほかの仕入価格は “相場の仮” です。実際の値に直すと原価が正しくなります。\n' +
           '・みりん/醤油/昆布/鰹節/塩 は既にある食材を使っています（備考に正式名）。\n' +
           '・作り方は画像に無いので一般的な手順です。食べてみて直してください。\n\n' +
           '「🍳 レシピ → レシピを見る（カード）」で選ぶと、まとめて見られます。');
}

/** IDから行番号を探す（1列目を上から照合）。無ければ 0 */
function findRowById_(sh, id) {
  var last = lastDataRow_(sh);
  if (last < 2) return 0;
  var vals = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === String(id).trim()) return i + 2;
  }
  return 0;
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
  build試作手順(ss);
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

  applyAutoFormulas_(sh);   // 数式は AUTO_FORMULAS にまとめてあります

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

  applyAutoFormulas_(sh);   // 数式は AUTO_FORMULAS にまとめてあります

  sh.getRange(2, 6, LAST_ROW, 1).setNumberFormat('¥#,##0.0000');
  sh.getRange(2, 7, LAST_ROW, 1).setNumberFormat('¥#,##0.0');
  markAuto_(sh, 3, 1);
  markAuto_(sh, 5, 3);
}


/* ========== ⑧ 試作比較（回を横に並べて見る画面） ========== */
function build試作比較(ss) {
  var sh = ss.getSheetByName('試作比較') || newSheet_(ss, '試作比較');

  // 作り直しても「いま見ている対象」が消えないように覚えておく
  var keepTarget = '';
  try { keepTarget = String(sh.getRange('B1').getValue()).trim(); } catch (e) {}

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

  // 選択候補（「PRE-0001 生姜だれ」の形。選ぶと onEdit がIDだけに直す）
  sh.getRange('Z1').setFormula(PICK_SRC_FORMULA.parent);
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
    // ★SUMIFS は ARRAYFORMULA の中だと1行目しか計算されず、全部の食材に同じ数字が入ってしまう。
    //  「試作ID|子ID」の合体キーにして SUMIF で引くと、行ごとに正しく計算される。ここを SUMIFS に戻さないこと。
    amount.push('=ARRAYFORMULA(IF($A16:$A="","",IF(' + col + '$2="","",IFERROR(' +
      'SUMIF(試作レシピ明細!$A$2:$A&"|"&試作レシピ明細!$B$2:$B,' + col + '$2&"|"&$A16:$A,試作レシピ明細!$D$2:$D),""))))');
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

  // 覚えておいた対象を戻す（何も選んでいなければ、試作記録の一番新しい対象を出す）
  if (!keepTarget) {
    try {
      var recSh = ss.getSheetByName('試作記録');
      var recLast = recSh ? lastDataRow_(recSh) : 1;
      if (recLast >= 2) {
        var col = recSh.getRange(2, 2, recLast - 1, 1).getValues();
        for (var z = col.length - 1; z >= 0; z--) {
          if (String(col[z][0]).trim()) { keepTarget = String(col[z][0]).trim(); break; }
        }
      }
    } catch (e2) {}
  }
  if (keepTarget) sh.getRange('B1').setValue(keepTarget);

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
/**
 * 試作をはじめる（3つの画面をひとつづきで）
 *
 *   1. 何を試作するか  … 一覧から選ぶ / 名前を入れて新しく作る
 *   2. 材料と分量      … 無い材料はその場で食材マスタに登録できる
 *   3. できました      … どこに何が入ったか、次に何をするか
 *
 * ★人にIDを打たせない・画面を行き来させない・作った結果を必ず見せる。
 */
function 次の回の試作をつくる() {
  var html = HtmlService.createHtmlOutput(trialStartHtml_())
    .setWidth(900).setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, '🧪 試作をはじめる');
}

/** 材料の候補（食材＋仕込み品）。登録フォームと同じ形で返す */
function trialItems() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var items = [];
  readRows_(ss, '食材マスタ', 18).forEach(function (r) {
    if (r[0]) items.push({ id: String(r[0]), name: String(r[1]), unit: String(r[3]), cost: Number(r[17]) || 0, kind: '食材' });
  });
  readRows_(ss, '仕込み品マスタ', 8).forEach(function (r) {
    if (r[0]) items.push({ id: String(r[0]), name: String(r[1]), unit: String(r[3]), cost: Number(r[7]) || 0, kind: '仕込み品' });
  });
  return items;
}

/** 画面が最初に読み込むデータ */
function trialStartData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var recSh = ss.getSheetByName('試作記録');
  var maxNo = {};
  var recLast = recSh ? lastDataRow_(recSh) : 1;
  if (recLast >= 2) {
    recSh.getRange(2, 1, recLast - 1, 5).getValues().forEach(function (r) {
      var t = String(r[1]).trim();
      if (!t) return;
      var n = Number(r[3]) || 0;
      if (!maxNo[t] || n > maxNo[t]) maxNo[t] = n;
    });
  }

  var targets = [];
  readRows_(ss, '仕込み品マスタ', 8).forEach(function (r) {
    if (r[0]) targets.push({ id: String(r[0]), name: String(r[1]), kind: 'prep', no: maxNo[String(r[0])] || 0 });
  });
  readRows_(ss, '商品マスタ', 10).forEach(function (r) {
    if (r[0]) targets.push({ id: String(r[0]), name: String(r[1]), kind: 'menu', no: maxNo[String(r[0])] || 0 });
  });
  targets.sort(function (a, b) { return b.no - a.no; });

  return {
    targets: targets, items: trialItems(),
    cats: OPT_分類, units: OPT_在庫単位, stores: OPT_保存区分, packs: OPT_発注単位
  };
}

/**
 * 試作の回を作る（画面1 → 画面2 へ進むときに呼ばれる）
 * p = { mode:'exist'|'new', id, kind:'prep'|'menu', name }
 */
function trialStartCreate(p) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var recSh = ss.getSheetByName('試作記録');
  var detSh = ss.getSheetByName('試作レシピ明細');
  if (!recSh || !detSh) throw new Error('試作シートが見つかりません。先に「試作シートを追加」を実行してください。');

  var target = '', targetName = '', created = '';

  if (p.mode === 'new') {
    var name = String(p.name || '').trim();
    if (!name) throw new Error('名前を入れてください。');

    var isPrep = (p.kind === 'prep');
    var sh   = ss.getSheetByName(isPrep ? '仕込み品マスタ' : '商品マスタ');
    var rule = ID_RULES[isPrep ? '仕込み品マスタ' : '商品マスタ'];

    // 同じ名前が既にあれば、それを使う（二重登録を防ぐ）
    var exist = '';
    readRows_(ss, isPrep ? '仕込み品マスタ' : '商品マスタ', 8).forEach(function (r) {
      if (String(r[1]).trim() === name && !exist) exist = String(r[0]).trim();
    });

    if (exist) {
      target = exist; targetName = name;
    } else {
      target = rule.prefix + zero4_(nextIdNumber_(sh, rule));
      targetName = name;
      var row = lastDataRow_(sh) + 1;
      if (isPrep) {
        // 出来上がり量・保存区分・日持ちは、この時点ではまだ決まらないので空のまま。
        // 出来上がり量が空の間は「材料の合計」が自動で使われる。
        sh.getRange(row, 1, 1, 4).setValues([[target, name, '', 'g']]);
      } else {
        sh.getRange(row, 1, 1, 2).setValues([[target, name]]);
      }
      created = target;
    }
  } else {
    target = String(p.id || '').trim();
    if (!target) throw new Error('どれを試作するか選んでください。');
    readRows_(ss, target.indexOf('PRE-') === 0 ? '仕込み品マスタ' : '商品マスタ', 8).forEach(function (r) {
      if (String(r[0]).trim() === target) targetName = String(r[1]);
    });
  }

  // これまでの最新回と、その分量を探す
  var rec = recSh.getDataRange().getValues();
  var lastNo = 0, lastId = '';
  for (var i = 1; i < rec.length; i++) {
    if (String(rec[i][1]).trim() === target && (Number(rec[i][3]) || 0) >= lastNo) {
      lastNo = Number(rec[i][3]) || 0;
      lastId = String(rec[i][0]).trim();
    }
  }

  var newNo = lastNo + 1;
  var newId = 次の試作ID_(rec);

  // 前回の分量をそのまま引き継ぐ（1回目なら空）
  var lines = [];
  if (lastId) {
    detSh.getDataRange().getValues().forEach(function (r, j) {
      if (j > 0 && String(r[0]).trim() === lastId) {
        lines.push({ id: String(r[1]).trim(), qty: Number(r[3]) || 0, memo: String(r[7] || '') });
      }
    });
  }

  // 前回の作り方（手順）も引き継ぐ（1回目や手順シートが無ければ空）
  var steps = [];
  var stepSh = ss.getSheetByName('試作手順');
  if (lastId && stepSh) {
    stepSh.getDataRange().getValues().forEach(function (r, j) {
      if (j > 0 && String(r[0]).trim() === lastId) {
        steps.push({ text: String(r[2] || ''), time: String(r[3] || ''), heat: String(r[4] || ''), point: String(r[5] || '') });
      }
    });
  }

  var r2 = lastDataRow_(recSh) + 1;
  recSh.getRange(r2, 1, 1, 2).setValues([[newId, target]]);
  recSh.getRange(r2, 4, 1, 2).setValues([[newNo, new Date()]]);

  if (lines.length > 0) {
    rebuildBlock_(detSh, [1, 2, 4, 8], [{ key: newId, rows: lines.map(function (L) {
      return [newId, L.id, L.qty, L.memo];
    }) }]);
  }

  var cmp = ss.getSheetByName('試作比較');
  if (cmp) cmp.getRange('B1').setValue(target);
  SpreadsheetApp.flush();

  return {
    tryId: newId, no: newNo, target: target, targetName: targetName,
    created: created, lastNo: lastNo, lines: lines, steps: steps
  };
}

/** 材料の名前から、似ている食材・仕込み品を探す */
function ingredientSuggest(name) {
  var n = normName_(name);
  if (!n) return [];
  var out = [];
  trialItems().forEach(function (it) {
    var m = normName_(it.name);
    var score = 0;
    if (m === n) score = 100;
    else if (m.indexOf(n) >= 0 || n.indexOf(m) >= 0) score = 80;
    else {
      var d = editDistance_(m, n);
      if (d <= 2 && Math.min(m.length, n.length) >= 3) score = 60 - d * 10;
    }
    if (score > 0) out.push({ id: it.id, name: it.name, unit: it.unit, cost: it.cost, kind: it.kind, score: score });
  });
  out.sort(function (a, b) { return b.score - a.score; });
  return out.slice(0, 5);
}

/** 名前を比べやすい形にそろえる（カタカナ→ひらがな・全角→半角・空白と記号を除く） */
function normName_(s) {
  var t = String(s == null ? '' : s).trim().toLowerCase();
  t = t.replace(/[\s　・（）()／\/\-ー_]/g, '');
  t = t.replace(/[\u30A1-\u30F6]/g, function (c) { return String.fromCharCode(c.charCodeAt(0) - 0x60); });
  t = t.replace(/[Ａ-Ｚａ-ｚ０-９]/g, function (c) { return String.fromCharCode(c.charCodeAt(0) - 0xFEE0); });
  return t;
}

/** 2つの文字列がどれくらい違うか（1文字の打ち間違いなら1） */
function editDistance_(a, b) {
  if (a === b) return 0;
  if (!a.length) return b.length;
  if (!b.length) return a.length;
  var prev = [], cur = [], i, j;
  for (j = 0; j <= b.length; j++) prev[j] = j;
  for (i = 1; i <= a.length; i++) {
    cur[0] = i;
    for (j = 1; j <= b.length; j++) {
      cur[j] = Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (a.charAt(i - 1) === b.charAt(j - 1) ? 0 : 1));
    }
    for (j = 0; j <= b.length; j++) prev[j] = cur[j];
  }
  return prev[b.length];
}

/** 新しい食材を、その場で食材マスタに登録する */
function ingredientCreate(p) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('食材マスタ');
  if (!sh) throw new Error('食材マスタが見つかりません。');

  var name = String(p.name || '').trim();
  if (!name) throw new Error('食材名を入れてください。');

  // 同じ名前が既にあればそれを返す（二重登録を防ぐ）
  var exist = null;
  trialItems().forEach(function (it) {
    if (!exist && it.kind === '食材' && it.name === name) exist = it;
  });
  if (exist) return { id: exist.id, name: exist.name, unit: exist.unit, cost: exist.cost, already: true };

  var id  = 'ING-' + zero4_(nextIdNumber_(sh, ID_RULES['食材マスタ']));
  var row = lastDataRow_(sh) + 1;

  // A食材ID B食材名 C分類 D在庫単位 E保存区分 F日持ち日数 G歩留まり率
  sh.getRange(row, 1, 1, 7).setValues([[
    id, name, String(p.cat || ''), String(p.unit || 'g'), String(p.store || ''),
    Number(p.days) || '', Number(p.yieldRate) > 0 ? Number(p.yieldRate) / 100 : 1
  ]]);
  // J発注単位 K入数 L仕入価格（原価を出すのに要る。空でも登録はできる）
  sh.getRange(row, 10, 1, 3).setValues([[
    String(p.pack || ''), Number(p.packQty) || '', Number(p.price) || ''
  ]]);

  SpreadsheetApp.flush();

  var cost = 0;
  if (Number(p.packQty) > 0 && Number(p.price) > 0) {
    cost = Number(p.price) / Number(p.packQty) / (Number(p.yieldRate) > 0 ? Number(p.yieldRate) / 100 : 1);
  }
  return { id: id, name: name, unit: String(p.unit || 'g'), cost: cost, already: false };
}

/** 材料と分量を試作レシピ明細に書き込む */
function trialLinesSave(p) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var detSh = ss.getSheetByName('試作レシピ明細');
  if (!detSh) throw new Error('試作レシピ明細が見つかりません。');

  var tryId = String(p.tryId || '').trim();
  if (!tryId) throw new Error('試作IDがありません。');

  var lines = (p.lines || [])
    .filter(function (L) { return String(L.id || '').trim() !== ''; })
    .map(function (L) { return [tryId, String(L.id).trim(), Number(L.qty) || 0, String(L.memo || '')]; });

  rebuildBlock_(detSh, [1, 2, 4, 8], [{ key: tryId, rows: lines }]);
  SpreadsheetApp.flush();

  var cost = 0, total = 0, unit = 'g';
  var byId = {};
  trialItems().forEach(function (it) { byId[it.id] = it; });
  lines.forEach(function (L) {
    var it = byId[L[1]];
    if (it) { cost += it.cost * L[2]; unit = it.unit || unit; }
    total += L[2];
  });

  return { count: lines.length, cost: Math.round(cost * 10) / 10, total: total, unit: unit };
}

/* ========== 試作手順（作り方＝試作の回ごと） ========== */

/** 試作手順シートが無ければ作って返す（ヘッダー・体裁・役割まで整える） */
function ensureTrialSteps_(ss) {
  var sh = ss.getSheetByName('試作手順');
  if (sh) return sh;
  sh = newSheet_(ss, '試作手順');
  var headers = ['試作ID', '手順番号', '手順内容', '所要時間', '温度・火加減', 'ポイント'];
  var widths  = [90, 70, 380, 90, 110, 240];
  setHeader_(sh, headers, widths);
  sh.getRange(2, 2, LAST_ROW, 1).setNumberFormat('0').setHorizontalAlignment('center');
  sh.getRange(1, 3).setNote('作り方を1行ずつ。番号は自動で振られます。\n試作の回ごとに残るので、前の回とくらべられます。');
  if (COLUMN_ROLES['試作手順']) { applyRoles_(sh); applyProtection_(sh); }
  return sh;
}

/** 試作手順シート（作り方）を作る。サンプル付き（新規台帳・試作シート追加で使う） */
function build試作手順(ss) {
  var sh = ensureTrialSteps_(ss);
  if (lastDataRow_(sh) < 2) {
    sh.getRange(2, 1, 3, 6).setValues([
      ['TRY-0002', 1, '醤油・みりん・酒を鍋に合わせる',       '1分', '',     ''],
      ['TRY-0002', 2, 'ひと煮立ちさせてアルコールを飛ばす',   '3分', '中火', '沸かしすぎない'],
      ['TRY-0002', 3, 'すりおろした生姜を加えて火を止める',   '1分', '',     '香りを残す']
    ]);
  }
}

/** フォームから作り方を保存する（試作記録の回＝tryId ごとに試作手順へ書く） */
function trialStepsSave(p) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ensureTrialSteps_(ss);

  var tryId = String(p && p.tryId || '').trim();
  if (!tryId) throw new Error('試作IDがありません。');

  var steps = (p.steps || [])
    .filter(function (s) { return String(s.text || '').trim() !== ''; })
    .map(function (s, i) {
      return [tryId, i + 1, String(s.text).trim(), String(s.time || ''), String(s.heat || ''), String(s.point || '')];
    });

  rebuildBlock_(sh, [1, 2, 3, 4, 5, 6], [{ key: tryId, rows: steps }]);
  SpreadsheetApp.flush();
  return { count: steps.length };
}

/** 指定のシートを開く（できました画面の「見に行く」ボタン用） */
function trialOpenSheet(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (sh) ss.setActiveSheet(sh);
  return true;
}


/** 試作の画面 */
function trialStartHtml_() {
  return [
'<!DOCTYPE html><html><head><meta charset="utf-8"><base target="_top">',
'<style>',
'  body { font-family: "Meiryo", "Hiragino Kaku Gothic ProN", sans-serif; font-size: 13px;',
'         color: #2A2019; background: #FBF8F3; margin: 0; padding: 14px 18px 78px; }',
'  h2 { font-size: 15px; margin: 0 0 4px; }',
'  .lead { color: #7A7168; font-size: 12px; margin: 0 0 12px; }',
'  .steps { display: flex; gap: 6px; margin-bottom: 14px; font-size: 11px; }',
'  .steps div { flex: 1; padding: 7px 8px; background: #F0EBE2; color: #9A9189; border-radius: 4px; text-align: center; }',
'  .steps div.on { background: #8A6A2A; color: #fff; }',
'  .steps div.done { background: #DCCFB4; color: #5A4A2A; }',
'  .page { display: none; } .page.on { display: block; }',
'  .card { background: #fff; border: 1px solid #E4DED3; border-radius: 8px; margin-bottom: 12px; overflow: hidden; }',
'  .card.pick { cursor: pointer; }',
'  .card.on { border-color: #8A6A2A; box-shadow: 0 0 0 2px rgba(138,106,42,.15); }',
'  .head { display: flex; align-items: center; gap: 10px; padding: 12px 16px; background: #F7F4EE; }',
'  .card.on .head { background: #F3EADB; }',
'  .head b { font-size: 13px; } .head span { font-size: 11px; color: #7A7168; }',
'  .body { padding: 14px 16px; display: none; } .card.on .body { display: block; }',
'  .card.plain .body { display: block; }',
'  label { display: block; font-size: 11px; color: #7A7168; margin-bottom: 3px; }',
'  input, select { font-family: inherit; font-size: 13px; padding: 7px 8px; width: 100%;',
'         border: 1px solid #D8D0C4; border-radius: 4px; background: #fff; box-sizing: border-box; }',
'  input:focus, select:focus { outline: 2px solid #C8A96A; outline-offset: -1px; }',
'  .row { display: flex; gap: 10px; flex-wrap: wrap; margin-bottom: 10px; align-items: flex-end; }',
'  .row > div { flex: 1 1 120px; display: flex; flex-direction: column; justify-content: flex-end; }',
'  .note { font-size: 11px; color: #7A7168; margin: 8px 0 0; line-height: 1.6; }',
'  .kindbtn { display: flex; margin-bottom: 12px; }',
'  .kindbtn button { flex: 1; padding: 9px; font-size: 12px; font-family: inherit; cursor: pointer;',
'          border: 1px solid #D8D0C4; background: #F7F4EE; color: #7A7168; }',
'  .kindbtn button:first-child { border-radius: 5px 0 0 5px; }',
'  .kindbtn button:last-child { border-radius: 0 5px 5px 0; border-left: none; }',
'  .kindbtn button.on { background: #8A6A2A; color: #fff; border-color: #8A6A2A; }',
'  table { width: 100%; border-collapse: collapse; }',
'  th { font-size: 11px; color: #7A7168; text-align: left; font-weight: normal; padding: 0 6px 4px; }',
'  td { padding: 2px 6px 2px 0; vertical-align: middle; }',
'  td input { width: 100%; }',
'  .qty { text-align: right; }',
'  .u { font-size: 11px; color: #7A7168; white-space: nowrap; }',
'  .cost { font-size: 12px; text-align: right; white-space: nowrap; color: #4A423A; }',
'  .del { border: none; background: none; color: #B23B3B; cursor: pointer; font-size: 15px; padding: 2px 6px; }',
'  .add { border: 1px dashed #C8A96A; background: #FEFCF8; color: #8A6A2A; cursor: pointer;',
'         border-radius: 4px; padding: 6px 12px; font-size: 12px; margin-top: 6px; font-family: inherit; }',
'  .sum { display: flex; gap: 18px; align-items: baseline; background: #F3EADB;',
'         border-radius: 6px; padding: 10px 14px; margin-top: 12px; flex-wrap: wrap; }',
'  .sum b { font-size: 17px; }',
'  .newbox { background: #FDF6F4; border: 1px solid #E8C9C2; border-radius: 6px; padding: 12px 14px; margin: 6px 0 12px; }',
'  .newbox h4 { margin: 0 0 8px; font-size: 12px; color: #A0453A; }',
'  .sug { display: flex; gap: 8px; flex-wrap: wrap; margin-bottom: 10px; }',
'  .sug button { border: 1px solid #C8A96A; background: #fff; color: #6B5320; cursor: pointer;',
'         border-radius: 14px; padding: 5px 12px; font-size: 12px; font-family: inherit; }',
'  .sug button:hover { background: #F3EADB; }',
'  .mini { background: #8A6A2A; color: #fff; border: none; border-radius: 4px; padding: 7px 16px;',
'          font-size: 12px; cursor: pointer; font-family: inherit; }',
'  .done { background: #EEF4EC; border: 1px solid #C6DCC0; border-radius: 8px; padding: 16px 18px; margin-bottom: 14px; }',
'  .done h3 { margin: 0 0 10px; font-size: 14px; color: #3E6B3F; }',
'  .done table { margin-bottom: 4px; } .done td { padding: 3px 10px 3px 0; font-size: 12px; }',
'  .next { background: #fff; border: 1px solid #E4DED3; border-radius: 8px; padding: 14px 18px; }',
'  .next h3 { margin: 0 0 8px; font-size: 13px; }',
'  .next ol { margin: 0; padding-left: 20px; line-height: 1.9; font-size: 12px; }',
'  .link { background: none; border: none; color: #8A6A2A; text-decoration: underline; cursor: pointer;',
'          font-family: inherit; font-size: 12px; padding: 0; }',
'  .bar { position: fixed; left: 0; right: 0; bottom: 0; background: #fff; border-top: 1px solid #E4DED3;',
'         padding: 12px 18px; display: flex; gap: 12px; align-items: center; }',
'  .go { background: #8A6A2A; color: #fff; border: none; border-radius: 5px; padding: 10px 26px;',
'          font-size: 14px; cursor: pointer; font-family: inherit; }',
'  .go:disabled { background: #C4BCB0; cursor: default; }',
'  .back { background: #fff; border: 1px solid #D8D0C4; color: #7A7168; border-radius: 5px;',
'          padding: 10px 18px; font-size: 13px; cursor: pointer; font-family: inherit; }',
'  .msg { font-size: 12px; } .ng { color: #B23B3B; } .ok { color: #4B7B4E; }',
'  .dot { width: 16px; height: 16px; border-radius: 50%; border: 2px solid #C4BCB0; flex: none; }',
'  .card.on .dot { border-color: #8A6A2A; background: radial-gradient(#8A6A2A 0 5px, #fff 5px); }',
'  h4 { font-size: 12px; margin: 0 0 6px; color: #5A4A2A; }',
'  .stepcard { border: 1px solid #E4DED3; border-radius: 6px; padding: 10px 12px; margin-bottom: 10px; background: #fff; }',
'  .ings { display: flex; flex-wrap: wrap; gap: 6px; margin-bottom: 8px; }',
'  .ings label { display: inline-flex; align-items: center; gap: 4px; background: #F7F4EE; border: 1px solid #E0D8CC;',
'                border-radius: 14px; padding: 3px 10px; font-size: 12px; color: #4A423A; cursor: pointer; margin: 0; }',
'  .ings input { width: auto; margin: 0; }',
'  .stepline { display: flex; gap: 8px; align-items: center; flex-wrap: wrap; }',
'  .stepline select, .stepline input { width: auto; flex: 0 0 auto; }',
'  .taste { display: flex; align-items: center; gap: 8px; margin: 6px 0; }',
'  .taste .nm { width: 54px; font-size: 12px; font-weight: bold; }',
'  .taste.ref .nm { font-weight: normal; color: #9A9189; }',
'  .taste input[type=range] { flex: 1; accent-color: #8A6A2A; }',
'  .taste .val { width: 22px; text-align: right; }',
'  .taste .st { width: 58px; font-size: 11px; text-align: right; }',
'  .sep { border-top: 1px dashed #D8D0C4; margin: 8px 0 4px; font-size: 11px; color: #9A9189; padding-top: 6px; }',
'  .score { background: #F3EADB; border-radius: 6px; padding: 8px 12px; margin-top: 8px;',
'           display: flex; justify-content: space-between; align-items: baseline; }',
'  .score b { font-size: 20px; color: #8A6A2A; }',
'  .pricebox { background: #EEF4EC; border: 1px solid #C6DCC0; border-radius: 6px; padding: 8px 12px; margin-top: 10px; font-size: 13px; }',
'</style></head><body>',

'<div class="steps">',
'  <div id="s1" class="on">1. 何を試作するか</div>',
'  <div id="s2">2. 材料と分量</div>',
'  <div id="s3">3. 作り方</div>',
'  <div id="s4">4. できました</div>',
'</div>',

// ===== 画面1 =====
'<div class="page on" id="page1">',
'  <h2>試作をはじめる</h2>',
'  <p class="lead">IDを覚える必要はありません。一覧から選ぶか、名前を入れるだけです。</p>',
'  <div class="card pick on" id="cardExist" onclick="setMode(\'exist\')">',
'    <div class="head"><div class="dot"></div><b>いまあるものの続きをつくる</b><span id="existHint"></span></div>',
'    <div class="body">',
'      <label>どれの試作をしますか（選ぶとすぐ材料の画面に進みます）</label>',
'      <select id="target" onchange="onPick()"></select>',
'      <p class="note" id="nextNote"></p>',
'    </div>',
'  </div>',
'  <div class="card pick" id="cardNew" onclick="setMode(\'new\')">',
'    <div class="head"><div class="dot"></div><b>まだ台帳に無いものを新しく試す</b><span>IDは自動で入ります</span></div>',
'    <div class="body">',
'      <div class="kindbtn">',
'        <button type="button" id="kPrep" class="on" onclick="setKind(\'prep\',event)">仕込み品（タレ・ご飯など）</button>',
'        <button type="button" id="kMenu" onclick="setKind(\'menu\',event)">売る商品（メニュー）</button>',
'      </div>',
'      <label>名前</label><input id="newName" placeholder="例: 新しい甘辛だれ" oninput="preview()">',
'      <p class="note" id="newNote">名前を入れると、台帳への登録とIDの採番までまとめて行います。</p>',
'      <p class="note">出来上がり量や日持ちは、いまはまだ決めなくて構いません。<br>',
'         出来上がり量は、入れなければ「材料の合計」がそのまま使われます。</p>',
'    </div>',
'  </div>',
'</div>',

// ===== 画面2 =====
'<div class="page" id="page2">',
'  <h2 id="p2title">材料と分量を入れる</h2>',
'  <p class="lead" id="p2lead"></p>',
'  <div class="card plain">',
'    <div class="body">',
'      <table><thead><tr>',
'        <th style="width:38%">材料（打ちながら探せます）</th><th style="width:13%">分量</th>',
'        <th style="width:5%"></th><th style="width:13%">原価</th><th>ひとこと</th><th style="width:28px"></th>',
'      </tr></thead><tbody id="lines"></tbody></table>',
'      <p class="note" style="margin-top:4px">※一番下の行に入れると、自動で次の行が出ます。間違えた行は × で消せます。</p>',
'      <datalist id="itemlist"></datalist>',
'      <div id="newArea"></div>',
'      <div class="sum">',
'        <span>材料の合計 <b id="sumQty">0</b> <span id="sumUnit">g</span></span>',
'        <span>原価 <b id="sumCost">0</b> 円</span>',
'      </div>',
'    </div>',
'  </div>',
'</div>',

// ===== 画面3：作り方と味 =====
'<div class="page" id="page3">',
'  <h2 id="p3title">作り方と味</h2>',
'  <p class="lead">材料をえらんで動作をえらぶと、作り方の文ができます。味は材料から自動で出しています。</p>',
'  <div class="card plain"><div class="body">',
'    <h4>味のバランス（材料からの目安）</h4>',
'    <div id="tastebox"></div>',
'    <div class="score"><span>総合評価の目安</span><b id="tscore">—</b></div>',
'    <p class="note" id="tnote"></p>',
'    <div class="pricebox" id="pricebox"></div>',
'  </div></div>',
'  <div class="card plain"><div class="body">',
'    <h4>作り方</h4>',
'    <div id="steps"></div>',
'    <button class="add" onclick="addStep()">＋ 手順を追加</button>',
'    <p class="note">材料にチェック→動作をえらぶと文ができます（文は手で直してもOK）。<br>火を使わない工程は火加減を空に。作り方が未定なら空のまま進めます。</p>',
'  </div></div>',
'</div>',
'',
// ===== 画面4：できました =====
'<div class="page" id="page4">',
'  <h2>できました</h2>',
'  <p class="lead">どこに何が入ったかは、次のとおりです。</p>',
'  <div class="done" id="doneBox"></div>',
'  <div class="next">',
'    <h3>次にやること</h3>',
'    <ol>',
'      <li>書いた作り方のとおりに、実際に作ってみる</li>',
'      <li>食べて、<b>🧪 試作 → 味を記録・調整する</b>を押して味をスライダーで記録する',
'          <br><span class="note">材料から初期値が出るので、違うところだけ直せます。</span></li>',
'      <li>分量や作り方を変えて次の回を作るときは、もう一度この「試作をはじめる」を押す',
'          <br><span class="note">前の回の分量と作り方がコピーされるので、変えたいところだけ直せます。</span></li>',
'      <li>味が決まったら、試作記録の「採用」に ○ を付けて',
'          <br>「🧪 試作 → 採用した試作をレシピへ反映」を押す</li>',
'    </ol>',
'  </div>',
'</div>',

'<div class="bar">',
'  <button class="back" id="backBtn" onclick="back()" style="display:none">もどる</button>',
'  <button class="go" id="goBtn" onclick="go()">次へ（材料を入れる）</button>',
'  <span class="msg" id="msg">読み込んでいます…</span>',
'</div>',

'<script>',
'var MODE = "exist", KIND = "prep", STEP = 1;',
'var TARGETS = [], ITEMS = [], BYID = {}, OPT = {};',
'var TRY = null, LAST = null, STEPMATS = [], TASTEV = [10,10,10,10,10,10];',
'',
'google.script.run.withSuccessHandler(function (d) {',
'  TARGETS = d.targets; OPT = d;',
'  setItems(d.items);',
'  var sel = document.getElementById("target");',
'  sel.innerHTML = "<option value=\\"\\">▼ 選んでください</option>";',
'  TARGETS.forEach(function (t) {',
'    var o = document.createElement("option");',
'    o.value = t.id;',
'    o.textContent = t.name + "　" + (t.no > 0 ? "（第" + t.no + "回まで試作ずみ）" : "（まだ試作していません）");',
'    sel.appendChild(o);',
'  });',
'  document.getElementById("existHint").textContent =',
'    TARGETS.length ? "台帳に " + TARGETS.length + " 件" : "台帳がまだ空です";',
'  if (TARGETS.length === 0) setMode("new");',
'  showNext();',
'  say("");',
'}).withFailureHandler(fail).trialStartData();',
'',
'function setItems(items) {',
'  ITEMS = items; BYID = {};',
'  var dl = document.getElementById("itemlist");',
'  dl.innerHTML = "";',
'  ITEMS.forEach(function (it) {',
'    BYID[it.id] = it;',
'    var o = document.createElement("option");',
'    o.value = it.id + " " + it.name;',
'    o.label = it.kind + " / " + it.unit;',
'    dl.appendChild(o);',
'  });',
'}',
'',
'function say(t, bad) {',
'  var m = document.getElementById("msg");',
'  m.textContent = t; m.className = "msg " + (bad ? "ng" : "ok");',
'}',
'function fail(err) { say("エラー: " + err.message, true); document.getElementById("goBtn").disabled = false; }',
'',
'function setMode(m) {',
'  MODE = m;',
'  document.getElementById("cardExist").className = "card pick" + (m === "exist" ? " on" : "");',
'  document.getElementById("cardNew").className   = "card pick" + (m === "new"   ? " on" : "");',
'  say("");',
'}',
'function setKind(k, ev) {',
'  if (ev) ev.stopPropagation();',
'  KIND = k;',
'  document.getElementById("kPrep").className = (k === "prep" ? "on" : "");',
'  document.getElementById("kMenu").className = (k === "menu" ? "on" : "");',
'  preview();',
'}',
'function showNext() {',
'  var id = document.getElementById("target").value;',
'  var t = TARGETS.filter(function (x) { return x.id === id; })[0];',
'  document.getElementById("nextNote").textContent = t',
'    ? "→ " + t.name + " の第" + (t.no + 1) + "回をつくります。" +',
'      (t.no > 0 ? "前の回の分量がコピーされるので、変えたいところだけ直せます。" : "1回目です。")',
'    : "";',
'}',
// 一覧から選んだら、そのまま材料の画面へ進む（選んだ時点で②へ）
'function onPick() {',
'  setMode("exist"); showNext();',
'  if (document.getElementById("target").value) go();',
'}',
'function preview() {',
'  var n = document.getElementById("newName").value.trim();',
'  document.getElementById("newNote").textContent = n',
'    ? "→ " + (KIND === "prep" ? "仕込み品" : "商品") + "「" + n + "」を台帳に登録し、IDを自動で振って第1回をつくります。"',
'    : "名前を入れると、台帳への登録とIDの採番までまとめて行います。";',
'}',
'',
'function setStep(n) {',
'  STEP = n;',
'  [1,2,3,4].forEach(function (i) {',
'    document.getElementById("page" + i).className = "page" + (i === n ? " on" : "");',
'    document.getElementById("s" + i).className = (i === n ? "on" : (i < n ? "done" : ""));',
'  });',
'  document.getElementById("backBtn").style.display = (n === 2 || n === 3) ? "inline-block" : "none";',
'  var go = document.getElementById("goBtn");',
'  if (n === 1) go.textContent = "次へ（材料を入れる）";',
'  if (n === 2) go.textContent = "次へ（作り方を入れる）";',
'  if (n === 3) go.textContent = "保存する";',
'  if (n === 4) { go.textContent = "閉じる"; document.getElementById("backBtn").style.display = "none"; }',
'  go.disabled = false;',
'  window.scrollTo(0, 0);',
'}',
'function back() {',
'  if (STEP === 2) { setStep(1); document.getElementById("target").value = ""; document.getElementById("nextNote").textContent = ""; }',
'  else if (STEP === 3) setStep(2);',
'}',
'',
'/* ---------- 画面3：作り方（材料をえらぶ＋動作をえらぶ） ---------- */',
'var DOACTS = ["合わせる","混ぜる","炒める","焼く","煮る","煮詰める","ひと煮立ちさせる","揚げる","ゆでる","蒸す","冷やす","漬ける","和える","味をととのえる","盛り付ける"];',
'function addStep(v) {',
'  v = v || {};',
'  var card = document.createElement("div"); card.className = "stepcard";',
'  var ings = STEPMATS.map(function (m) {',
'    return "<label><input type=\\"checkbox\\" class=\\"ing\\" value=\\"" + esc(m.name) + "\\" onchange=\\"mkText(this)\\"> " + esc(m.name) + "</label>";',
'  }).join("");',
'  card.innerHTML =',
'    "<div class=\\"ings\\">" + ings + "<label><input type=\\"checkbox\\" class=\\"ingall\\" onchange=\\"mkText(this)\\"> ぜんぶ</label></div>" +',
'    "<div class=\\"stepline\\">" +',
'      "<select class=\\"act\\" onchange=\\"mkText(this)\\"><option value=\\"\\">動作をえらぶ</option>" +',
'         DOACTS.map(function (a) { return "<option>" + a + "</option>"; }).join("") + "</select>" +',
'      "<input class=\\"stime\\" placeholder=\\"時間 例:3分\\" style=\\"width:92px\\">" +',
'      "<select class=\\"sheat\\"><option value=\\"\\">火加減</option><option>弱火</option><option>中火</option><option>強火</option></select>" +',
'      "<button class=\\"del\\" onclick=\\"this.closest(\'.stepcard\').remove()\\">×</button>" +',
'    "</div>" +',
'    "<input class=\\"stext\\" placeholder=\\"ここに作り方の文ができます（手で直してもOK）\\" style=\\"margin-top:6px\\">";',
'  document.getElementById("steps").appendChild(card);',
'  if (v.text !== undefined) {',
'    card.querySelector(".stext").value = v.text || "";',
'    card.querySelector(".stime").value = v.time || "";',
'    if (v.heat) card.querySelector(".sheat").value = v.heat;',
'  }',
'  return card;',
'}',
'function mkText(el) {',
'  var card = el.closest(".stepcard");',
'  var all = card.querySelector(".ingall").checked;',
'  var names = [];',
'  card.querySelectorAll(".ing").forEach(function (c) { if (all || c.checked) names.push(c.value); });',
'  var act = card.querySelector(".act").value;',
'  card.querySelector(".stext").value = names.length ? (names.join("・") + "を" + act) : act;',
'}',
'/* ---------- 画面3：味バランス（材料からの目安） ---------- */',
'function buildTaste() {',
'  var core = ["甘味","塩味","酸味","旨味"], ref = ["苦味","辛味"];',
'  document.getElementById("tastebox").innerHTML =',
'    core.map(function (nm, i) { return trow(nm, i, false); }).join("") +',
'    "<div class=\\"sep\\">参考（総合評価には入れません）</div>" +',
'    ref.map(function (nm, i) { return trow(nm, i + 4, true); }).join("");',
'}',
'function trow(nm, i, ref) {',
'  return "<div class=\\"taste" + (ref ? " ref" : "") + "\\"><span class=\\"nm\\">" + nm + "</span>" +',
'    "<input type=\\"range\\" min=\\"1\\" max=\\"20\\" value=\\"10\\" id=\\"tr" + i + "\\" oninput=\\"tslide(" + i + ")\\">" +',
'    "<span class=\\"val\\" id=\\"tv" + i + "\\">10</span><span class=\\"st\\" id=\\"ts" + i + "\\"></span></div>";',
'}',
'function setTaste(v) { TASTEV = v.slice(); for (var i = 0; i < 6; i++) { document.getElementById("tr" + i).value = v[i]; tpaint(i); } tsc(); }',
'function tslide(i) { TASTEV[i] = Number(document.getElementById("tr" + i).value); tpaint(i); tsc(); }',
'function tpaint(i) {',
'  var v = TASTEV[i], core = i < 4, st = document.getElementById("ts" + i), t, c;',
'  document.getElementById("tv" + i).textContent = v;',
'  if (core) { if (v <= 7) { t = "足りない"; c = "#2E6DA4"; } else if (v >= 13) { t = "強すぎ"; c = "#B5482E"; } else { t = "ちょうど"; c = "#3E8E5A"; } }',
'  else      { if (v <= 7) { t = "弱い";     c = "#8A9199"; } else if (v >= 13) { t = "強い";   c = "#B5482E"; } else { t = "ふつう";   c = "#7A7168"; } }',
'  st.textContent = t; st.style.color = c;',
'}',
'function tsc() {',
'  var d = Math.abs(TASTEV[0]-10) + Math.abs(TASTEV[1]-10) + Math.abs(TASTEV[2]-10) + Math.abs(TASTEV[3]-10);',
'  document.getElementById("tscore").textContent = (Math.round((20 - d/4*1.9) * 10) / 10).toFixed(1);',
'}',
'function showPrice() {',
'  var box = document.getElementById("pricebox");',
'  var cost = LAST ? LAST.cost : 0;',
'  if (TRY && TRY.target && TRY.target.indexOf("PRE") === 0) {',
'    box.innerHTML = "これは仕込み品（タレ・ご飯など）なので、売価の目安は出しません。";',
'    return;',
'  }',
'  var price = cost > 0 ? Math.round(cost / 0.30 / 10) * 10 : 0;',
'  box.innerHTML = "原価 <b>" + cost + "</b> 円　→　売価の目安 <b>" + price + "</b> 円　<span class=\\"note\\">（原価率30%の目安。本番にするとき決められます）</span>";',
'}',
'',
'/* ---------- 画面2：材料 ---------- */',
'function addLine(v) {',
'  v = v || {};',
'  var tr = document.createElement("tr");',
'  tr.innerHTML =',
'    "<td><input class=\\"item\\" list=\\"itemlist\\" oninput=\\"calc()\\" placeholder=\\"例: 醤油\\"></td>" +',
'    "<td><input class=\\"qty amount\\" type=\\"number\\" step=\\"any\\" oninput=\\"calc()\\"></td>" +',
'    "<td class=\\"u\\"></td><td class=\\"cost\\">-</td>" +',
'    "<td><input class=\\"memo\\"></td>" +',
'    "<td><button class=\\"del\\" onclick=\\"this.closest(\'tr\').remove(); calc();\\">×</button></td>";',
'  document.getElementById("lines").appendChild(tr);',
'  if (v.id) {',
'    var it = BYID[v.id];',
'    tr.querySelector(".item").value = v.id + (it ? " " + it.name : "");',
'    tr.querySelector(".amount").value = v.qty;',
'    tr.querySelector(".memo").value = v.memo || "";',
'  }',
'  return tr;',
'}',
'',
'function resolve(text) {',
'  var t = (text || "").trim();',
'  if (!t) return null;',
'  var m = t.match(/^((?:ING|PRE)-\\d+)/);',
'  if (m && BYID[m[1]]) return BYID[m[1]];',
'  var hit = ITEMS.filter(function (it) { return it.name === t; });',
'  if (hit.length === 1) return hit[0];',
'  return null;',
'}',
'',
'function calc() {',
'  var cost = 0, qty = 0, unit = "g", unknown = [];',
'  var rows = document.querySelectorAll("#lines tr");',
'  for (var i = 0; i < rows.length; i++) {',
'    var tr = rows[i];',
'    var text = tr.querySelector(".item").value.trim();',
'    var it = resolve(text);',
'    var q = parseFloat(tr.querySelector(".amount").value) || 0;',
'    if (it) {',
'      cost += it.cost * q; qty += q; unit = it.unit || unit;',
'      tr.querySelector(".u").textContent = it.unit;',
'      tr.querySelector(".cost").textContent = "\\u00A5" + (Math.round(it.cost * q * 10) / 10);',
'      tr.querySelector(".item").style.background = "#fff";',
'    } else {',
'      tr.querySelector(".u").textContent = "";',
'      tr.querySelector(".cost").textContent = "-";',
'      tr.querySelector(".item").style.background = text ? "#FBEDEA" : "#fff";',
'      if (text) unknown.push(text);',
'    }',
'  }',
'  document.getElementById("sumQty").textContent = Math.round(qty * 10) / 10;',
'  document.getElementById("sumUnit").textContent = unit;',
'  document.getElementById("sumCost").textContent = Math.round(cost * 10) / 10;',
'  showUnknown(unknown);',
'  var rs = document.querySelectorAll("#lines tr");',
'  var lastRow = rs[rs.length - 1];',
'  if (lastRow && lastRow.querySelector(".item").value.trim() !== "") addLine();',
'}',
'',
'var askedFor = "";',
'function showUnknown(list) {',
'  var area = document.getElementById("newArea");',
'  if (list.length === 0) { area.innerHTML = ""; askedFor = ""; return; }',
'  var name = list[0];',
'  if (name === askedFor) return;          // 同じ名前で作り直さない（入力中に消えないように）',
'  askedFor = name;',
'  area.innerHTML =',
'    "<div class=\\"newbox\\">" +',
'    "  <h4>「" + esc(name) + "」は台帳にありません</h4>" +',
'    "  <div class=\\"sug\\" id=\\"sug\\"><span class=\\"note\\">似ているものを探しています…</span></div>" +',
'    "  <div id=\\"newForm\\"></div>" +',
'    "</div>";',
'  google.script.run.withSuccessHandler(function (s) {',
'    if (askedFor !== name) return;',
'    var sug = document.getElementById("sug");',
'    if (!sug) return;',
'    if (s.length > 0) {',
'      sug.innerHTML = "<span class=\\"note\\">もしかして、これですか？</span>";',
'      s.forEach(function (x) {',
'        var b = document.createElement("button");',
'        b.textContent = x.name + "（" + x.kind + "）";',
'        b.onclick = function () { useSuggest(name, x); };',
'        sug.appendChild(b);',
'      });',
'    } else { sug.innerHTML = "<span class=\\"note\\">似たものは見つかりませんでした。</span>"; }',
'    buildNewForm(name);',
'  }).withFailureHandler(function () { buildNewForm(name); }).ingredientSuggest(name);',
'}',
'',
'function esc(s) { return String(s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/\\"/g,"&quot;"); }',
'',
'function useSuggest(typed, item) {',
'  var rows = document.querySelectorAll("#lines tr");',
'  for (var i = 0; i < rows.length; i++) {',
'    if (rows[i].querySelector(".item").value.trim() === typed) {',
'      rows[i].querySelector(".item").value = item.id + " " + item.name;',
'    }',
'  }',
'  askedFor = "";',
'  calc();',
'}',
'',
'function buildNewForm(name) {',
'  var f = document.getElementById("newForm");',
'  if (!f) return;',
'  f.innerHTML =',
'    "<div class=\\"note\\" style=\\"margin-bottom:8px\\">新しい食材として、いま登録できます。</div>" +',
'    "<div class=\\"row\\">" +',
'      "<div><label>分類</label>"    + sel("nCat",   OPT.cats)   + "</div>" +',
'      "<div><label>単位</label>"    + sel("nUnit",  OPT.units, "g")  + "</div>" +',
'      "<div><label>保存</label>"    + sel("nStore", OPT.stores, "冷蔵") + "</div>" +',
'      "<div><label>日持ち日数</label><input id=\\"nDays\\" type=\\"number\\" placeholder=\\"7\\"></div>" +',
'      "<div><label>歩留まり率(%)</label><input id=\\"nYield\\" type=\\"number\\" value=\\"100\\"></div>" +',
'    "</div>" +',
'    "<div class=\\"row\\">" +',
'      "<div><label>仕入の単位</label>" + sel("nPack", OPT.packs) + "</div>" +',
'      "<div><label>1つ何" + "gか（入数）</label><input id=\\"nPackQty\\" type=\\"number\\" step=\\"any\\" placeholder=\\"1000\\" oninput=\\"newCost()\\"></div>" +',
'      "<div><label>仕入価格（円）</label><input id=\\"nPrice\\" type=\\"number\\" step=\\"any\\" placeholder=\\"500\\" oninput=\\"newCost()\\"></div>" +',
'      "<div><label>&nbsp;</label><button class=\\"mini\\" onclick=\\"createItem()\\">食材マスタに登録する</button></div>" +',
'    "</div>" +',
'    "<p class=\\"note\\" id=\\"nCostNote\\">入数と仕入価格を入れると、1あたりの原価が出ます。空のままでも登録はできます（原価は0円）。</p>";',
'  f.dataset.name = name;',
'}',
'',
'function sel(id, arr, def) {',
'  var s = "<select id=\\"" + id + "\\"><option value=\\"\\"></option>";',
'  (arr || []).forEach(function (v) {',
'    s += "<option value=\\"" + esc(v) + "\\"" + (v === def ? " selected" : "") + ">" + esc(v) + "</option>";',
'  });',
'  return s + "</select>";',
'}',
'',
'function newCost() {',
'  var q = parseFloat((document.getElementById("nPackQty")||{}).value) || 0;',
'  var p = parseFloat((document.getElementById("nPrice")||{}).value) || 0;',
'  var y = parseFloat((document.getElementById("nYield")||{}).value) || 100;',
'  var n = document.getElementById("nCostNote");',
'  if (!n) return;',
'  n.textContent = (q > 0 && p > 0)',
'    ? "→ 1" + (document.getElementById("nUnit").value || "g") + "あたり " + (Math.round(p / q / (y/100) * 10000) / 10000) + " 円"',
'    : "入数と仕入価格を入れると、1あたりの原価が出ます。空のままでも登録はできます（原価は0円）。";',
'}',
'',
'function createItem() {',
'  var f = document.getElementById("newForm");',
'  var name = f.dataset.name;',
'  say("登録しています…");',
'  google.script.run.withSuccessHandler(function (res) {',
'    google.script.run.withSuccessHandler(function (items) {',
'      setItems(items);',
'      var rows = document.querySelectorAll("#lines tr");',
'      for (var i = 0; i < rows.length; i++) {',
'        if (rows[i].querySelector(".item").value.trim() === name) {',
'          rows[i].querySelector(".item").value = res.id + " " + res.name;',
'        }',
'      }',
'      askedFor = "";',
'      document.getElementById("newArea").innerHTML = "";',
'      calc();',
'      say(res.already ? res.id + "「" + res.name + "」は既にありました。そちらを使います。"',
'                      : res.id + "「" + res.name + "」を食材マスタに登録しました。");',
'    }).trialItems();',
'  }).withFailureHandler(fail).ingredientCreate({',
'    name: name,',
'    cat:   (document.getElementById("nCat")||{}).value,',
'    unit:  (document.getElementById("nUnit")||{}).value,',
'    store: (document.getElementById("nStore")||{}).value,',
'    days:  parseFloat((document.getElementById("nDays")||{}).value) || 0,',
'    yieldRate: parseFloat((document.getElementById("nYield")||{}).value) || 100,',
'    pack:  (document.getElementById("nPack")||{}).value,',
'    packQty: parseFloat((document.getElementById("nPackQty")||{}).value) || 0,',
'    price: parseFloat((document.getElementById("nPrice")||{}).value) || 0',
'  });',
'}',
'',
'function openSheet(n) { google.script.run.trialOpenSheet(n); }',
'',
'/* ---------- 進む ---------- */',
'function go() {',
'  if (STEP === 4) { google.script.host.close(); return; }',
'',
'  if (STEP === 1) {',
'    var p = { mode: MODE };',
'    if (MODE === "exist") {',
'      p.id = document.getElementById("target").value;',
'      if (!p.id) { say("どれを試作するか選んでください。", true); return; }',
'    } else {',
'      p.kind = KIND;',
'      p.name = document.getElementById("newName").value.trim();',
'      if (!p.name) { say("名前を入れてください。", true); return; }',
'    }',
'    document.getElementById("goBtn").disabled = true;',
'    say("つくっています…");',
'    google.script.run.withSuccessHandler(function (res) {',
'      TRY = res;',
'      document.getElementById("p2title").textContent = res.targetName + "　第" + res.no + "回の材料";',
'      document.getElementById("p2lead").innerHTML =',
'        (res.created ? "<b>" + res.created + "「" + res.targetName + "」を台帳に登録しました。</b><br>" : "") +',
'        (res.lines.length > 0',
'          ? "第" + res.lastNo + "回の分量をコピーしました。変えたいところだけ直してください。"',
'          : "1回目です。使う材料を入れてください。");',
'      document.getElementById("lines").innerHTML = "";',
'      if (res.lines.length > 0) res.lines.forEach(addLine);',
'      else { addLine(); addLine(); addLine(); }',
'      calc();',
'      setStep(2);',
'      say("材料を打ち始めると候補が出ます。無いものはその場で登録できます。");',
'    }).withFailureHandler(fail).trialStartCreate(p);',
'    return;',
'  }',
'',
'  if (STEP === 2) {',
'    var lines = [], bad = [];',
'    var rows = document.querySelectorAll("#lines tr");',
'    for (var i = 0; i < rows.length; i++) {',
'      var text = rows[i].querySelector(".item").value.trim();',
'      if (!text) continue;',
'      var it = resolve(text);',
'      if (!it) { bad.push(text); continue; }',
'      lines.push({ id: it.id, qty: parseFloat(rows[i].querySelector(".amount").value) || 0,',
'                   memo: rows[i].querySelector(".memo").value });',
'    }',
'    if (bad.length > 0) { say("「" + bad.join("」「") + "」がまだ登録されていません。上の赤い枠から登録するか、行を消してください。", true); return; }',
'    if (lines.length === 0) { say("材料を1つ以上入れてください。", true); return; }',
'    document.getElementById("goBtn").disabled = true;',
'    say("材料を保存しています…");',
'    google.script.run.withSuccessHandler(function (r) {',
'      LAST = r;',
'      STEPMATS = lines.map(function (L) { var it = BYID[L.id]; return { id: L.id, name: it ? it.name : L.id }; });',
'      document.getElementById("p3title").textContent = TRY.targetName + "　第" + TRY.no + "回の作り方と味";',
'      buildTaste(); setTaste([10,10,10,10,10,10]);',
'      document.getElementById("tnote").textContent = "味を計算しています…";',
'      google.script.run.withSuccessHandler(function (res) {',
'        setTaste(res.values);',
'        document.getElementById("tnote").textContent = res.count > 0 ? "材料から出した目安です。食べて違うところはスライダーで直せます。" : "材料が読めないので、ぜんぶ10から始めます。";',
'      }).tasteEstimate(TRY.tryId);',
'      showPrice();',
'      var sb = document.getElementById("steps"); sb.innerHTML = "";',
'      if (TRY.steps && TRY.steps.length > 0) TRY.steps.forEach(addStep);',
'      else { addStep(); addStep(); }',
'      setStep(3);',
'      say("");',
'    }).withFailureHandler(fail).trialLinesSave({ tryId: TRY.tryId, lines: lines });',
'    return;',
'  }',
'',
'  if (STEP === 3) {',
'    var steps = [];',
'    var srows = document.querySelectorAll("#steps .stepcard");',
'    for (var i = 0; i < srows.length; i++) {',
'      var t = srows[i].querySelector(".stext").value.trim();',
'      if (!t) continue;',
'      steps.push({ text: t, time: srows[i].querySelector(".stime").value, heat: srows[i].querySelector(".sheat").value });',
'    }',
'    document.getElementById("goBtn").disabled = true;',
'    say("保存しています…");',
'    google.script.run.withSuccessHandler(function (sres) {',
'      google.script.run.withSuccessHandler(function () {',
'        var r = LAST || { count: 0, total: 0, unit: "g", cost: 0 };',
'        var isMenu = !(TRY.target && TRY.target.indexOf("PRE") === 0);',
'        var price = (isMenu && r.cost > 0) ? Math.round(r.cost / 0.30 / 10) * 10 : 0;',
'        document.getElementById("doneBox").innerHTML =',
'          "<h3>" + esc(TRY.targetName) + "　第" + TRY.no + "回（" + TRY.tryId + "）ができました</h3>" +',
'          "<table>" +',
'          (TRY.created ? "<tr><td>" + (TRY.created.indexOf("PRE-") === 0 ? "仕込み品マスタ" : "商品マスタ") +',
'             "</td><td>" + TRY.created + "「" + esc(TRY.targetName) + "」を登録</td></tr>" : "") +',
'          "<tr><td>試作記録</td><td>" + TRY.tryId + "（第" + TRY.no + "回）／味も記録</td></tr>" +',
'          "<tr><td>試作レシピ明細</td><td>材料 " + r.count + " 件</td></tr>" +',
'          "<tr><td>試作手順</td><td>作り方 " + sres.count + " 手順</td></tr>" +',
'          "<tr><td>材料の合計</td><td>" + r.total + " " + r.unit + "</td></tr>" +',
'          "<tr><td>この回の原価</td><td>" + r.cost + " 円</td></tr>" +',
'          (isMenu ? "<tr><td>売価の目安</td><td>" + price + " 円（原価率30%）</td></tr>" : "") +',
'          "</table>" +',
'          "<button class=\\"link\\" onclick=\\"openSheet(\'試作レシピ明細\')\\">材料を見る</button>　" +',
'          "<button class=\\"link\\" onclick=\\"openSheet(\'試作手順\')\\">作り方を見る</button>　" +',
'          "<button class=\\"link\\" onclick=\\"openSheet(\'試作比較\')\\">試作比較を開く</button>";',
'        setStep(4);',
'        say("");',
'      }).withFailureHandler(fail).tasteSave({ tryId: TRY.tryId, values: TASTEV });',
'    }).withFailureHandler(fail).trialStepsSave({ tryId: TRY.tryId, steps: steps });',
'    return;',
'  }',
'}',
'<\/script></body></html>'
  ].join('\n');
}

function 次の試作ID_(rec) {
  var max = 0;
  for (var i = 1; i < rec.length; i++) {
    var m = String(rec[i][0]).match(/^TRY-(\d+)$/);
    if (m) max = Math.max(max, Number(m[1]));
  }
  return 'TRY-' + ('0000' + (max + 1)).slice(-4);
}


/* =====================================================================
 *  味の記録・調整
 *   材料の「味の強さ」から初期値を出し、食べた感覚で違うところだけ直す。
 *   保存先は試作記録の F〜K（甘・塩・酸・旨・苦・辛）。総合評価は自動で出る。
 * ===================================================================== */

/**
 * 試作の材料（lines=[{id,qty}]）から 6味の初期値 [甘,塩,酸,旨,苦,辛]（1〜20）を推定する。
 *   料理全体の「その味の平均%」を、基準の濃さ(TASTE_BASE)で割って 10前後になるよう換算。
 *   仕込み品(PRE)は構成食材から 1gあたりの味を出す（2階層まで。原価の約束と同じ）。
 */
function estimateTaste_(lines) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var ingTaste = {};   // 食材ID → [甘,塩,酸,旨,苦,辛]
  readRows_(ss, '食材マスタ', 24).forEach(function (r) {
    if (!r[0]) return;
    ingTaste[String(r[0]).trim()] = [18, 19, 20, 21, 22, 23].map(function (c) { return Number(r[c]) || 0; });
  });

  var prepQty = {};    // 仕込み品ID → 出来上がり量（空なら 0）
  readRows_(ss, '仕込み品マスタ', 8).forEach(function (r) {
    if (r[0]) prepQty[String(r[0]).trim()] = Number(r[2]) || 0;
  });

  var recipe = {};     // 親ID → [{id, qty}]
  readRows_(ss, 'レシピ明細', 10).forEach(function (r) {
    var p = String(r[0] || '').trim(); if (!p) return;
    (recipe[p] = recipe[p] || []).push({ id: String(r[2] || '').trim(), qty: Number(r[4]) || 0 });
  });

  // 仕込み品 1gあたりの味（中の仕込み品は味に含めない＝2階層）
  function prepPerG(prepId) {
    var kids = recipe[prepId] || [];
    var sum = [0, 0, 0, 0, 0, 0], w = 0;
    kids.forEach(function (k) {
      if (k.id.indexOf('ING') === 0 && ingTaste[k.id]) {
        for (var i = 0; i < 6; i++) sum[i] += ingTaste[k.id][i] * k.qty;
      }
      w += k.qty;   // PREの子も重さには数える（味だけ含めない）
    });
    var denom = prepQty[prepId] > 0 ? prepQty[prepId] : w;
    if (denom <= 0) return [0, 0, 0, 0, 0, 0];
    return sum.map(function (x) { return x / denom; });
  }

  var sum = [0, 0, 0, 0, 0, 0], total = 0;
  (lines || []).forEach(function (L) {
    var id = String(L.id || '').trim(), qty = Number(L.qty) || 0;
    if (!id || qty <= 0) return;
    var per = (id.indexOf('PRE') === 0) ? prepPerG(id) : (ingTaste[id] || [0, 0, 0, 0, 0, 0]);
    for (var i = 0; i < 6; i++) sum[i] += per[i] * qty;
    total += qty;
  });
  if (total <= 0) return [10, 10, 10, 10, 10, 10];

  var base = [TASTE_BASE.sweet, TASTE_BASE.salt, TASTE_BASE.sour, TASTE_BASE.umami, TASTE_BASE.bitter, TASTE_BASE.hot];
  return sum.map(function (s, i) {
    var v = Math.round(10 * (s / total) / base[i]);
    return Math.max(1, Math.min(20, v));
  });
}

/** 味フォームが最初に読み込むデータ（回の一覧・新しい順） */
function tasteFormData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var out = [];
  readRows_(ss, '試作記録', 15).forEach(function (r) {
    var id = String(r[0] || '').trim(); if (!id) return;
    var saved = [5, 6, 7, 8, 9, 10].map(function (c) { var v = r[c]; return (v === '' || v == null) ? '' : Number(v); });
    out.push({
      tryId: id,
      label: (String(r[2] || '').trim() || id) + '（' + (Number(r[3]) || 0) + '回目）',
      saved: saved,
      hasSaved: saved.some(function (v) { return v !== ''; })
    });
  });
  out.reverse();
  return out;
}

/** ある回の味の初期値を、材料から推定して返す */
function tasteEstimate(tryId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lines = [];
  readRows_(ss, '試作レシピ明細', 8).forEach(function (r) {
    if (String(r[0] || '').trim() === String(tryId).trim()) {
      lines.push({ id: String(r[1] || '').trim(), qty: Number(r[3]) || 0 });
    }
  });
  return { values: estimateTaste_(lines), count: lines.length };
}

/** スライダーの値を試作記録の F〜K（甘〜辛）に保存する */
function tasteSave(p) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('試作記録');
  if (!sh) throw new Error('試作記録が見つかりません。');

  var tryId = String(p && p.tryId || '').trim();
  if (!tryId) throw new Error('どの回か分かりません。');

  var vals = (p.values || []).map(function (v) { return Math.max(1, Math.min(20, Math.round(Number(v) || 10))); });
  if (vals.length !== 6) throw new Error('味の数がそろっていません。');

  var last = lastDataRow_(sh);
  if (last >= 2) {
    var ids = sh.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]).trim() === tryId) {
        sh.getRange(i + 2, 6, 1, 6).setValues([vals]);   // F〜K
        SpreadsheetApp.flush();
        return { ok: true };
      }
    }
  }
  throw new Error('試作記録に ' + tryId + ' が見つかりません。');
}

/** メニュー：味を記録・調整する */
function 味を記録調整する() {
  if (!SpreadsheetApp.getActiveSpreadsheet().getSheetByName('試作記録')) {
    SpreadsheetApp.getUi().alert('試作記録が見つかりません。\n\n先に「試作シートを追加」を実行してください。');
    return;
  }
  var html = HtmlService.createHtmlOutput(tasteFormHtml_()).setWidth(660).setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, '🍯 味を記録・調整する');
}

function tasteFormHtml_() {
  return [
'<!DOCTYPE html><html><head><meta charset="utf-8"><base target="_top">',
'<style>',
'  body { font-family: "Meiryo","Hiragino Kaku Gothic ProN",sans-serif; font-size: 13px;',
'         color: #2A2019; background: #FBF8F3; margin: 0; padding: 14px 18px 80px; }',
'  h2 { font-size: 15px; margin: 0 0 4px; }',
'  .lead { color: #7A7168; font-size: 12px; margin: 0 0 12px; line-height: 1.6; }',
'  label { display: block; font-size: 11px; color: #7A7168; margin-bottom: 3px; }',
'  select { font-family: inherit; font-size: 13px; padding: 7px 8px; width: 100%;',
'           border: 1px solid #D8D0C4; border-radius: 4px; background: #fff; box-sizing: border-box; }',
'  .card { background: #fff; border: 1px solid #E4DED3; border-radius: 8px; padding: 12px 16px; margin: 12px 0; }',
'  .taste { display: flex; align-items: center; gap: 10px; margin: 10px 0; }',
'  .taste .nm { width: 62px; font-size: 13px; font-weight: bold; }',
'  .taste.ref .nm { font-weight: normal; color: #9A9189; }',
'  .taste input[type=range] { flex: 1; accent-color: #8A6A2A; }',
'  .taste .val { width: 26px; text-align: right; font-variant-numeric: tabular-nums; }',
'  .taste .st { width: 64px; font-size: 11px; text-align: right; }',
'  .sep { border-top: 1px dashed #D8D0C4; margin: 12px 0 6px; font-size: 11px; color: #9A9189; padding-top: 8px; }',
'  .score { background: #F3EADB; border-radius: 6px; padding: 10px 14px; margin-top: 10px;',
'           display: flex; align-items: baseline; justify-content: space-between; }',
'  .score b { font-size: 22px; color: #8A6A2A; } .score span { font-size: 11px; color: #7A7168; }',
'  .bar { position: fixed; left: 0; right: 0; bottom: 0; background: #fff; border-top: 1px solid #E4DED3;',
'         padding: 10px 18px; display: flex; gap: 10px; }',
'  button { font-family: inherit; font-size: 13px; padding: 10px 14px; border-radius: 5px; cursor: pointer; border: 1px solid #D8D0C4; }',
'  .primary { background: #8A6A2A; color: #fff; border-color: #8A6A2A; flex: 1; }',
'  .ghost { background: #F7F4EE; color: #5A4A2A; }',
'  .hint { font-size: 11px; color: #9A9189; margin-top: 6px; }',
'  .msg { font-size: 12px; margin-left: 4px; align-self: center; }',
'</style></head><body>',
'<h2>🍯 味を記録・調整する</h2>',
'<p class="lead">材料から味の初期値を自動で出しています。<b>食べてみて違うところだけ</b>つまみを動かして直してください。<br>',
'  まん中(10)＝ちょうど良い／左＝足りない／右＝強すぎ。甘・塩・酸・旨が総合評価に入ります。</p>',
'<label>どの回の味を入れますか</label>',
'<select id="pick" onchange="loadRow()"></select>',
'<div class="card" id="panel">',
'  <div id="sliders"></div>',
'  <div class="sep">参考（総合評価には入れません）</div>',
'  <div id="sliders2"></div>',
'  <div class="score"><span id="scLabel">総合評価の目安</span><b id="score">—</b></div>',
'  <div class="hint" id="src"></div>',
'</div>',
'<div class="bar">',
'  <button class="ghost" onclick="reEstimate()">材料から推定し直す</button>',
'  <button class="primary" onclick="save()">保存する</button>',
'  <span class="msg" id="msg"></span>',
'</div>',
'<script>',
'var CORE = ["甘味","塩味","酸味","旨味"];',
'var REF  = ["苦味","辛味"];',
'var rows = [], cur = null, vals = [10,10,10,10,10,10];',
'function esc(s){ return String(s).replace(/[<>&]/g,function(c){return {"<":"&lt;",">":"&gt;","&":"&amp;"}[c];}); }',
'window.onload = function(){',
'  google.script.run.withSuccessHandler(function(d){',
'    rows = d || [];',
'    var sel = document.getElementById("pick");',
'    if(!rows.length){ sel.innerHTML = "<option>（試作の回がありません）</option>"; return; }',
'    sel.innerHTML = rows.map(function(r,i){ return "<option value=\\""+i+"\\">"+esc(r.label)+(r.hasSaved?"　✓記録済み":"")+"</option>"; }).join("");',
'    buildSliders(); loadRow();',
'  }).tasteFormData();',
'};',
'function buildSliders(){',
'  document.getElementById("sliders").innerHTML = CORE.map(function(nm,i){ return sliderHtml(nm,i,false); }).join("");',
'  document.getElementById("sliders2").innerHTML = REF.map(function(nm,i){ return sliderHtml(nm,i+4,true); }).join("");',
'}',
'function sliderHtml(nm,idx,ref){',
'  return "<div class=\\"taste"+(ref?" ref":"")+"\\">"+',
'         "<span class=\\"nm\\">"+nm+"</span>"+',
'         "<input type=\\"range\\" min=\\"1\\" max=\\"20\\" value=\\"10\\" id=\\"r"+idx+"\\" oninput=\\"onSlide("+idx+")\\">"+',
'         "<span class=\\"val\\" id=\\"v"+idx+"\\">10</span>"+',
'         "<span class=\\"st\\" id=\\"s"+idx+"\\"></span></div>";',
'}',
'function loadRow(){',
'  var i = Number(document.getElementById("pick").value) || 0;',
'  cur = rows[i]; if(!cur) return;',
'  document.getElementById("msg").textContent = "";',
'  if(cur.hasSaved){',
'    vals = cur.saved.map(function(v){ return v===""?10:v; });',
'    setAll(); setSource("前に記録した値を出しています。");',
'  } else {',
'    reEstimate();',
'  }',
'}',
'function reEstimate(){',
'  if(!cur) return;',
'  setSource("推定中…");',
'  google.script.run.withSuccessHandler(function(res){',
'    vals = res.values.slice(); setAll();',
'    setSource(res.count>0 ? "材料"+res.count+"種から推定した初期値です。食べて違うところを直してください。" : "材料がまだ入っていないので、ぜんぶ 10 から始めます。");',
'  }).tasteEstimate(cur.tryId);',
'}',
'function setAll(){ for(var i=0;i<6;i++){ document.getElementById("r"+i).value = vals[i]; paint(i); } score(); }',
'function onSlide(i){ vals[i] = Number(document.getElementById("r"+i).value); paint(i); score(); }',
'function paint(i){',
'  var v = vals[i], core = i<4;',
'  document.getElementById("v"+i).textContent = v;',
'  var st = document.getElementById("s"+i), txt, col;',
'  if(core){ if(v<=7){txt="足りない";col="#2E6DA4";} else if(v>=13){txt="強すぎ";col="#B5482E";} else {txt="ちょうど";col="#3E8E5A";} }',
'  else    { if(v<=7){txt="弱い";col="#8A9199";} else if(v>=13){txt="強い";col="#B5482E";} else {txt="ふつう";col="#7A7168";} }',
'  st.textContent = txt; st.style.color = col;',
'}',
'function score(){',
'  var d = Math.abs(vals[0]-10)+Math.abs(vals[1]-10)+Math.abs(vals[2]-10)+Math.abs(vals[3]-10);',
'  var s = Math.round((20 - d/4*1.9)*10)/10;',
'  document.getElementById("score").textContent = s.toFixed(1);',
'}',
'function setSource(t){ document.getElementById("src").textContent = t; }',
'function save(){',
'  if(!cur){ return; }',
'  document.getElementById("msg").textContent = "保存中…";',
'  google.script.run.withSuccessHandler(function(){',
'    document.getElementById("msg").textContent = "✓ 保存しました";',
'    cur.hasSaved = true; cur.saved = vals.slice();',
'    var sel = document.getElementById("pick"); var i = Number(sel.value)||0;',
'    if(sel.options[i] && sel.options[i].textContent.indexOf("✓")<0) sel.options[i].textContent += "　✓記録済み";',
'  }).withFailureHandler(function(e){ document.getElementById("msg").textContent = "エラー: "+(e && e.message || e); }).tasteSave({ tryId: cur.tryId, values: vals });',
'}',
'</script></body></html>'
  ].join('\n');
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
    s.getRange(2, 1, Math.max(lastDataRow_(s) - 1, 1), 1).getValues().forEach(function (v) {
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

  // 商品(MENU)が対象なら、本番に移すこのタイミングで売価を「原価率◯%」で付ける
  var menuPicks = picks.filter(function (q) { return q.target.indexOf('MENU') === 0; });
  var ratePct = 0;
  if (menuPicks.length > 0) {
    var res = ui.prompt('売価を決めます（本番の原価率）',
      '商品の売価を「原価率◯%」から自動で付けます。\n数字を入れてOK（空のままOKで 30%）。\n売価を付けない場合はキャンセル。',
      ui.ButtonSet.OK_CANCEL);
    if (res.getSelectedButton() === ui.Button.OK) {
      var txt = String(res.getResponseText()).trim();
      ratePct = (txt === '') ? 30 : Number(txt);
      if (!(ratePct > 0 && ratePct <= 100)) {
        ui.alert('原価率は 1〜100 の数字で入れてください。\n今回は売価の設定をとばして、レシピだけ反映します。');
        ratePct = 0;
      }
    }
  }

  // 行は削除せず、入力列(A/C/E/J)だけ組み直す。
  // ※ B/D/F/G/H/I 列の2行目には自動計算の数式が入っているため、行削除は厳禁。
  //   この組み直しは rebuildBlock_ が一手に引き受けている（メニュー登録フォームも同じ関数を通る）。
  rebuildBlock_(repSh, [1, 3, 5, 10], picks.map(function (q) {
    return {
      key: q.target,
      rows: lines[q.tryId].map(function (L) { return [q.target, L[0], L[1], L[2]]; })
    };
  }));

  var stamp = Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd');
  picks.forEach(function (q) { recSh.getRange(q.row, 15).setValue('反映済 ' + stamp); });

  // 原価率から売価を付ける（原価はレシピ反映で再計算されるので flush してから読む）
  var priceMsg = '';
  if (ratePct > 0) {
    SpreadsheetApp.flush();
    var pm = ss.getSheetByName('商品マスタ');
    var pmLast = Math.max(lastDataRow_(pm) - 1, 1);
    var pmVals = pm.getRange(2, 1, pmLast, 8).getValues();   // A〜H（H=原価(自動)）
    var map = {};
    pmVals.forEach(function (r, idx) { if (r[0]) map[String(r[0]).trim()] = { row: idx + 2, cost: Number(r[7]) || 0 }; });
    menuPicks.forEach(function (q) {
      var m = map[q.target];
      if (m && m.cost > 0) {
        var price = Math.round(m.cost / (ratePct / 100) / 10) * 10;
        pm.getRange(m.row, 4).setValue(price);             // D列=売価
        pm.getRange(m.row, 5).setValue(ratePct / 100);     // E列=目標原価率
        priceMsg += '\n  ' + q.target + ' 原価' + (Math.round(m.cost * 10) / 10) + '円 → 売価' + price + '円（原価率' + ratePct + '%）';
      }
    });
  }

  ss.setActiveSheet(repSh);
  ui.alert('反映しました。\n\n' + summary +
           (priceMsg ? '\n\n【売価】' + priceMsg : '') +
           '\n\n商品マスタの原価・売価・原価率を確認してください。');
}

function q_(v) { return v === null || v === undefined ? '' : String(v).trim(); }


/* =====================================================================
 *  レシピを見る（1品を1枚のカードにまとめて表示・印刷できる）
 * ===================================================================== */

/** メニュー：レシピを見る（カード） */
function レシピを見る() {
  if (!SpreadsheetApp.getActiveSpreadsheet().getSheetByName('試作記録')) {
    SpreadsheetApp.getUi().alert('試作記録が見つかりません。\n\n先に「試作シートを追加」を実行してください。');
    return;
  }
  var html = HtmlService.createHtmlOutput(recipeViewHtml_()).setWidth(720).setHeight(760);
  SpreadsheetApp.getUi().showModalDialog(html, '🍳 レシピを見る');
}

/** 品と回の一覧（新しい順）。試作の回ごとにカードにできる */
function recipeList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var map = {}, order = [];
  function ent(target, name, kind) {
    if (!map[target]) { map[target] = { target: target, name: name || target, kind: kind, trials: [] }; order.push(target); }
    if (name && map[target].name === target) map[target].name = name;
    return map[target];
  }

  // 登録済みレシピ（本番）＝ 商品マスタ・仕込み品マスタ。no=1e9 で先頭に並べる
  readRows_(ss, '商品マスタ', 2).forEach(function (r) {
    var id = String(r[0] || '').trim(); if (!id) return;
    ent(id, String(r[1] || '').trim(), '商品').trials.push({ tryId: id, no: 1e9, label: '登録レシピ（本番）' });
  });
  readRows_(ss, '仕込み品マスタ', 2).forEach(function (r) {
    var id = String(r[0] || '').trim(); if (!id) return;
    ent(id, String(r[1] || '').trim(), '仕込み品').trials.push({ tryId: id, no: 1e9, label: '登録レシピ（本番）' });
  });

  // 試作の回
  readRows_(ss, '試作記録', 15).forEach(function (r) {
    var id = String(r[0] || '').trim(); if (!id) return;
    var tgt = String(r[1] || '').trim();
    var no = Number(r[3]) || 0;
    ent(tgt, String(r[2] || '').trim(), (tgt.indexOf('PRE') === 0 ? '仕込み品' : '商品'))
      .trials.push({ tryId: id, no: no, label: '第' + no + '回' });
  });

  return order.map(function (t) {
    var o = map[t];
    o.trials.sort(function (a, b) { return b.no - a.no; });   // 登録レシピ→新しい回の順
    return o;
  });
}

/** 日付の値を yyyy/MM/dd に。文字列(2013.8.31)などはそのまま返す */
function fmtDate_(v) {
  try { return Utilities.formatDate(new Date(v), 'JST', 'yyyy/MM/dd'); }
  catch (e) { return String(v || ''); }
}

/** カードの入口：TRY-… は試作カード、MENU-/PRE-… は登録レシピ（本番）カード */
function recipeCard(key) {
  key = String(key || '').trim();
  if (key.indexOf('TRY') === 0) return recipeCardTrial_(key);
  return recipeCardMaster_(key);
}

/** 登録レシピ（本番＝商品マスタ・仕込み品マスタ）を、昔の基準表の形でまとめて返す */
function recipeCardMaster_(id) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var isPrep = (id.indexOf('PRE') === 0);
  var card = { tryId: id, target: id, isMaster: true, kind: isPrep ? '仕込み品' : '商品',
               name: '', no: 0, date: '', tastes: ['', '', '', '', '', ''], hasTaste: false,
               estimated: false, overall: '', cost: 0, price: 0, salePrice: '', margin: '', rate: '',
               unitCost: '', lines: [], steps: [], meta: {} };

  if (isPrep) {
    readRows_(ss, '仕込み品マスタ', 12).forEach(function (r) {
      if (String(r[0]).trim() !== id) return;
      card.name = String(r[1] || '').trim() || id;
      card.cost = Math.round((Number(r[6]) || 0) * 10) / 10;                 // 原価合計(自動)
      card.unitCost = (r[7] === '' || r[7] == null) ? '' : (Math.round(Number(r[7]) * 10000) / 10000);
      card.meta = { code: '', cat: '', yield_: (r[2] === '' || r[2] == null) ? '' : (r[2] + ' ' + String(r[3] || '')),
                    time: String(r[9] || ''), save: String(r[4] || ''),
                    keep: (r[5] === '' || r[5] == null) ? '' : (r[5] + '日'),
                    author: String(r[10] || ''), made: r[11] ? fmtDate_(r[11]) : '' };
    });
  } else {
    readRows_(ss, '商品マスタ', 15).forEach(function (r) {
      if (String(r[0]).trim() !== id) return;
      card.name = String(r[1] || '').trim() || id;
      card.cost = Math.round((Number(r[7]) || 0) * 10) / 10;                 // 原価(自動)
      card.salePrice = (r[3] === '' || r[3] == null) ? '' : Number(r[3]);
      card.rate = (r[8] === '' || r[8] == null) ? '' : (Math.round(Number(r[8]) * 1000) / 10);  // %
      if (card.salePrice !== '') card.margin = Math.round((card.salePrice - card.cost) * 10) / 10;
      card.meta = { code: String(r[10] || ''), cat: String(r[2] || ''), yield_: String(r[11] || ''),
                    time: String(r[12] || ''), save: '', keep: '',
                    author: String(r[13] || ''), made: r[14] ? fmtDate_(r[14]) : '' };
    });
  }

  // 材料（レシピ明細 親ID=id）。味の推定のために子IDも集める
  var matLines = [];
  readRows_(ss, 'レシピ明細', 10).forEach(function (r) {
    if (String(r[0]).trim() === id && r[2]) {
      card.lines.push({ name: String(r[3] || '').trim() || String(r[2]).trim(),
                        qty: Number(r[4]) || 0, unit: String(r[5] || ''),
                        cost: Math.round((Number(r[8]) || 0) * 10) / 10, memo: String(r[9] || '') });
      matLines.push({ id: String(r[2]).trim(), qty: Number(r[4]) || 0 });
    }
  });

  // 作り方（調理手順 商品ID/仕込みID=id）
  readRows_(ss, '調理手順', 7).forEach(function (r) {
    if (String(r[0]).trim() === id && r[3]) {
      card.steps.push({ no: Number(r[2]) || 0, text: String(r[3] || ''), time: String(r[4] || ''), heat: String(r[5] || '') });
    }
  });
  card.steps.sort(function (a, b) { return a.no - b.no; });

  // 味は材料から推定（目安）。登録レシピには味の記録そのものが無いため
  try {
    var est = estimateTaste_(matLines);
    if (est && est.length === 6 && matLines.length) { card.tastes = est; card.hasTaste = true; card.estimated = true; }
  } catch (e) {}

  return card;
}

/** 1つの試作の回（tryId）を、材料・作り方・味・原価・売価をまとめて返す */
function recipeCardTrial_(tryId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  tryId = String(tryId || '').trim();
  var card = { tryId: tryId, name: '', no: 0, kind: '', date: '', target: '',
               tastes: ['', '', '', '', '', ''], hasTaste: false, overall: '', isMaster: false,
               estimated: false, cost: 0, price: 0, salePrice: '', margin: '', rate: '', unitCost: '',
               lines: [], steps: [], meta: {} };

  readRows_(ss, '試作記録', 15).forEach(function (r) {
    if (String(r[0]).trim() !== tryId) return;
    card.target = String(r[1] || '').trim();
    card.name = String(r[2] || '').trim() || card.target;
    card.no = Number(r[3]) || 0;
    card.kind = (card.target.indexOf('PRE') === 0) ? '仕込み品' : '商品';
    card.date = r[4] ? Utilities.formatDate(new Date(r[4]), 'JST', 'yyyy/MM/dd') : '';
    card.tastes = [5, 6, 7, 8, 9, 10].map(function (c) { var v = r[c]; return (v === '' || v == null) ? '' : Number(v); });
    card.hasTaste = card.tastes.some(function (v) { return v !== ''; });
    card.overall = (r[11] === '' || r[11] == null) ? '' : Number(r[11]);
    card.cost = Math.round((Number(r[12]) || 0) * 10) / 10;
  });

  readRows_(ss, '試作レシピ明細', 8).forEach(function (r) {
    if (String(r[0]).trim() === tryId && r[1]) {
      card.lines.push({ name: String(r[2] || '').trim() || String(r[1]).trim(),
                        qty: Number(r[3]) || 0, unit: String(r[4] || ''),
                        cost: Math.round((Number(r[6]) || 0) * 10) / 10, memo: String(r[7] || '') });
    }
  });

  var stepSh = ss.getSheetByName('試作手順');
  if (stepSh) {
    readRows_(ss, '試作手順', 6).forEach(function (r) {
      if (String(r[0]).trim() === tryId && r[2]) {
        card.steps.push({ no: Number(r[1]) || 0, text: String(r[2] || ''), time: String(r[3] || ''), heat: String(r[4] || '') });
      }
    });
    card.steps.sort(function (a, b) { return a.no - b.no; });
  }

  if (card.kind === '商品' && card.cost > 0) card.price = Math.round(card.cost / 0.30 / 10) * 10;
  return card;
}

function recipeViewHtml_() {
  return [
'<!DOCTYPE html><html><head><meta charset="utf-8"><base target="_top">',
'<style>',
'  body { font-family: "Meiryo","Hiragino Kaku Gothic ProN",sans-serif; font-size: 13px;',
'         color: #2A2019; background: #FBF8F3; margin: 0; padding: 14px 18px 24px; }',
'  h2 { font-size: 15px; margin: 0 0 10px; }',
'  .pick { display: flex; gap: 10px; margin-bottom: 12px; }',
'  select { font-family: inherit; font-size: 13px; padding: 7px 8px; border: 1px solid #D8D0C4; border-radius: 4px; background: #fff; }',
'  #target { flex: 2; } #no { flex: 1; }',
'  .card { background: #fff; border: 1px solid #E4DED3; border-radius: 8px; padding: 18px 20px; }',
'  .ttl { font-size: 18px; font-weight: bold; margin: 0 0 2px; }',
'  .sub { font-size: 12px; color: #7A7168; margin-bottom: 14px; }',
'  .sec { margin: 14px 0 6px; font-size: 12px; font-weight: bold; color: #8A6A2A; border-bottom: 1px solid #EEE6D8; padding-bottom: 3px; }',
'  table { width: 100%; border-collapse: collapse; }',
'  td, th { padding: 4px 6px; font-size: 13px; text-align: left; }',
'  th { color: #7A7168; font-weight: normal; font-size: 11px; border-bottom: 1px solid #EEE6D8; }',
'  td.r, th.r { text-align: right; white-space: nowrap; }',
'  ol { margin: 4px 0; padding-left: 22px; line-height: 1.9; }',
'  .tag { font-size: 11px; color: #7A7168; }',
'  table.info td { padding: 3px 6px; vertical-align: top; }',
'  table.info td.il { color: #7A7168; width: 96px; white-space: nowrap; }',
'  .tr { display: flex; align-items: center; gap: 8px; margin: 3px 0; }',
'  .tr .nm { width: 48px; font-size: 12px; }',
'  .bar { flex: 1; height: 12px; background: #EFEAE0; border-radius: 6px; position: relative; }',
'  .bar i { position: absolute; top: -2px; width: 8px; height: 16px; border-radius: 3px; background: #8A6A2A; transform: translateX(-50%); }',
'  .bar u { position: absolute; top: 0; bottom: 0; left: 50%; width: 1px; background: #C4BCB0; }',
'  .tr .vv { width: 24px; text-align: right; font-size: 12px; }',
'  .price { background: #EEF4EC; border: 1px solid #C6DCC0; border-radius: 6px; padding: 10px 14px; margin-top: 8px; }',
'  .price b { font-size: 16px; }',
'  .empty { color: #9A9189; font-size: 12px; }',
'  .bar2 { margin-top: 14px; display: flex; gap: 10px; }',
'  button { font-family: inherit; font-size: 13px; padding: 9px 16px; border-radius: 5px; cursor: pointer; border: 1px solid #D8D0C4; background: #F7F4EE; color: #5A4A2A; }',
'  .msg { font-size: 12px; color: #7A7168; align-self: center; }',
'  @media print { .pick, .bar2, h2 { display: none; } body { background: #fff; padding: 0; } .card { border: none; } }',
'</style></head><body>',
'<h2>🍳 レシピを見る</h2>',
'<div class="pick">',
'  <select id="target" onchange="onTarget()"></select>',
'  <select id="no" onchange="show()"></select>',
'</div>',
'<div class="card" id="card"><p class="empty">品をえらぶと、ここにレシピが出ます。</p></div>',
'<div class="bar2">',
'  <button onclick="window.print()">印刷する</button>',
'  <span class="msg" id="msg">読み込んでいます…</span>',
'</div>',
'<script>',
'var LIST = [], CUR = null;',
'function esc(s){ return String(s==null?"":s).replace(/[<>&]/g,function(c){return {"<":"&lt;",">":"&gt;","&":"&amp;"}[c];}); }',
'window.onload = function(){',
'  google.script.run.withSuccessHandler(function(d){',
'    LIST = d || [];',
'    var sel = document.getElementById("target");',
'    if(!LIST.length){ sel.innerHTML = "<option>（まだ試作がありません）</option>"; document.getElementById("msg").textContent=""; return; }',
'    sel.innerHTML = LIST.map(function(o,i){ return "<option value=\\""+i+"\\">"+esc(o.name)+"（"+o.kind+"）</option>"; }).join("");',
'    onTarget();',
'    document.getElementById("msg").textContent = "";',
'  }).recipeList();',
'};',
'function onTarget(){',
'  var o = LIST[Number(document.getElementById("target").value)||0]; if(!o) return;',
'  var sel = document.getElementById("no");',
'  sel.innerHTML = o.trials.map(function(t){ return "<option value=\\""+t.tryId+"\\">"+esc(t.label||("第"+t.no+"回"))+"</option>"; }).join("");',
'  show();',
'}',
'function show(){',
'  var tryId = document.getElementById("no").value; if(!tryId) return;',
'  document.getElementById("msg").textContent = "読み込んでいます…";',
'  google.script.run.withSuccessHandler(draw).recipeCard(tryId);',
'}',
'function tasterow(nm, v){',
'  if(v==="") return "<div class=\\"tr\\"><span class=\\"nm\\">"+nm+"</span><div class=\\"bar\\"><u></u></div><span class=\\"vv\\">-</span></div>";',
'  var pos = Math.max(0, Math.min(100, (v-1)/19*100));',
'  return "<div class=\\"tr\\"><span class=\\"nm\\">"+nm+"</span><div class=\\"bar\\"><u></u><i style=\\"left:"+pos+"%\\"></i></div><span class=\\"vv\\">"+v+"</span></div>";',
'}',
'function draw(c){',
'  CUR = c;',
'  document.getElementById("msg").textContent = "";',
'  var names = ["甘味","塩味","酸味","旨味","苦味","辛味"];',
'  var taste = c.hasTaste ? names.map(function(nm,i){ return tasterow(nm, c.tastes[i]); }).join("")',
'                        : "<p class=\\"empty\\">まだ味は記録されていません。</p>";',
'  var tasteNote = c.estimated ? " <span class=\\"tag\\">（材料からの目安）</span>" : "";',
'  var overall = (c.overall==="") ? "-" : c.overall;',
'  var lines = c.lines.length ? ("<table><tr><th>材料</th><th class=\\"r\\">分量</th><th class=\\"r\\">原価</th></tr>" +',
'       c.lines.map(function(L){ return "<tr><td>"+esc(L.name)+(L.memo?" <span class=\\"tag\\">"+esc(L.memo)+"</span>":"")+"</td><td class=\\"r\\">"+L.qty+" "+esc(L.unit)+"</td><td class=\\"r\\">"+L.cost+"円</td></tr>"; }).join("") +',
'       "</table>") : "<p class=\\"empty\\">材料がまだありません。</p>";',
'  var steps = c.steps.length ? ("<ol>" + c.steps.map(function(s){ var ex=[]; if(s.time)ex.push(s.time); if(s.heat)ex.push(s.heat); return "<li>"+esc(s.text)+(ex.length?" <span class=\\"tag\\">（"+esc(ex.join("・"))+"）</span>":"")+"</li>"; }).join("") + "</ol>") : "<p class=\\"empty\\">作り方がまだありません。</p>";',
'  var m = c.meta || {};',
'  function inf(label, v){ return v ? ("<tr><td class=\\"il\\">"+label+"</td><td>"+esc(v)+"</td></tr>") : ""; }',
'  var info = "";',
'  if (c.isMaster) {',
'    var irows = inf("商品コード", m.code) + inf("カテゴリ", m.cat) + inf((c.kind==="商品"?"仕上がり量":"出来上がり量"), m.yield_) + inf("仕込み時間", m.time) + inf("保存", m.save) + inf("日持ち", m.keep) + inf("製作者", m.author) + inf("制作年月日", m.made);',
'    if (irows) info = "<div class=\\"sec\\">基本情報</div><table class=\\"info\\">" + irows + "</table>";',
'  }',
'  var price;',
'  if (c.isMaster && c.kind==="商品") {',
'    price = "<div class=\\"price\\">売価 <b>"+(c.salePrice===""?"-":c.salePrice)+"</b> 円　原価 <b>"+c.cost+"</b> 円　粗利 <b>"+(c.margin===""?"-":c.margin)+"</b> 円"+(c.rate===""?"":"　<span class=\\"tag\\">原価率 "+c.rate+"%</span>")+"</div>";',
'  } else if (c.isMaster) {',
'    price = "<div class=\\"price\\">原価 <b>"+c.cost+"</b> 円"+(c.unitCost===""?"":"　単位原価 <b>"+c.unitCost+"</b> 円")+" <span class=\\"tag\\">（仕込み品なので売価は出しません）</span></div>";',
'  } else if (c.kind==="商品") {',
'    price = "<div class=\\"price\\">原価 <b>"+c.cost+"</b> 円　売価の目安 <b>"+(c.price||"-")+"</b> 円 <span class=\\"tag\\">（原価率30%の目安）</span></div>";',
'  } else {',
'    price = "<div class=\\"price\\">原価 <b>"+c.cost+"</b> 円 <span class=\\"tag\\">（仕込み品なので売価は出しません）</span></div>";',
'  }',
'  var sub = c.isMaster ? (esc(c.kind)+"　登録レシピ（本番）"+(m.made?"　"+esc(m.made):"")) : (esc(c.kind)+"　第"+c.no+"回"+(c.date?"　"+esc(c.date):"")+"　総合評価 "+overall);',
'  document.getElementById("card").innerHTML =',
'    "<p class=\\"ttl\\">"+esc(c.name)+"</p>" +',
'    "<p class=\\"sub\\">"+sub+"</p>" +',
'    info +',
'    "<div class=\\"sec\\">味のバランス"+tasteNote+"</div>" + taste +',
'    "<div class=\\"sec\\">材料</div>" + lines +',
'    "<div class=\\"sec\\">作り方</div>" + steps +',
'    price;',
'}',
'</script></body></html>'
  ].join('\n');
}


/* =====================================================================
 *  ここから下は「入力してよい場所 / 触ってはいけない場所」を守る仕組み
 *
 *  三段構えで守っています:
 *   ① 見て分かる      … 列の色とヘッダーの印（applyRoles_）
 *   ② 触ると警告が出る … 自動計算の列に保護をかける（applyProtection_）
 *   ③ それでも入力されたら自動で元に戻す（onEdit → restoreAuto_）
 *
 *  ★どの列が何なのかは COLUMN_ROLES だけに書いてあります。
 *   列を足す・動かすときは、必ずここも直してください。ここが全ての土台です。
 * ===================================================================== */

var ROLE_IN   = 'in';     // ✋ 手で入力する列
var ROLE_PICK = 'pick';   // ▼ プルダウンから選ぶ列
var ROLE_AUTO = 'auto';   // 🔒 自動計算の列。触ると元に戻る
var ROLE_ID   = 'id';     // 🔑 自動で振られるID。ふつうは触らない

var COLUMN_ROLES = {
  // 仕入先ID / 仕入先名 / 発注可能曜日 / 発注締め時間 / 最低発注金額 / 担当・連絡先 / 備考 / 登録食材数
  '仕入先マスタ': [ROLE_ID, ROLE_IN, ROLE_IN, ROLE_IN, ROLE_IN, ROLE_IN, ROLE_IN, ROLE_AUTO],

  // 食材ID / 食材名 / 分類 / 在庫単位 / 保存区分 / 日持ち日数 / 歩留まり率 / 仕入先ID / 仕入先商品コード /
  // 発注単位 / 入数 / 仕入価格 / kg単価 / リードタイム / 安全係数 / 実測補正 / 単価 / 実効単価 /
  // 甘味度 / 塩味度 / 酸味度 / 旨味度 / 苦味度 / 辛味度（味は手入力＝初期値を人が直せる）
  '食材マスタ': [ROLE_ID, ROLE_IN, ROLE_PICK, ROLE_PICK, ROLE_PICK, ROLE_IN, ROLE_IN, ROLE_PICK, ROLE_IN,
                 ROLE_PICK, ROLE_IN, ROLE_IN, ROLE_IN, ROLE_IN, ROLE_IN, ROLE_IN, ROLE_AUTO, ROLE_AUTO,
                 ROLE_IN, ROLE_IN, ROLE_IN, ROLE_IN, ROLE_IN, ROLE_IN],

  // 仕込みID / 仕込み品名 / 出来上がり量 / 単位 / 保存区分 / 日持ち日数 / 原価合計 / 単位原価 / チェック /
  // 仕込み時間 / 製作者 / 制作年月日（基準表の項目を後ろに追加）
  '仕込み品マスタ': [ROLE_ID, ROLE_IN, ROLE_IN, ROLE_PICK, ROLE_PICK, ROLE_IN, ROLE_AUTO, ROLE_AUTO, ROLE_AUTO,
                     ROLE_IN, ROLE_IN, ROLE_IN],

  // 商品ID / 商品名 / カテゴリ / 売価 / 目標原価率 / 提供開始日 / 提供終了日 / 原価 / 原価率 / 目標差 /
  // 商品コード / 仕上がり量 / 仕込み時間 / 製作者 / 制作年月日（基準表の項目を後ろに追加）
  '商品マスタ': [ROLE_ID, ROLE_IN, ROLE_IN, ROLE_IN, ROLE_IN, ROLE_IN, ROLE_IN, ROLE_AUTO, ROLE_AUTO, ROLE_AUTO,
                 ROLE_IN, ROLE_IN, ROLE_IN, ROLE_IN, ROLE_IN],

  // 親ID / 親名 / 子ID / 子名 / 使用量 / 単位 / 食材分原価 / 仕込み分原価 / 原価計 / 備考
  'レシピ明細': [ROLE_PICK, ROLE_AUTO, ROLE_PICK, ROLE_AUTO, ROLE_IN, ROLE_AUTO, ROLE_AUTO, ROLE_AUTO, ROLE_AUTO, ROLE_IN],

  // 商品ID / 商品名 / 手順番号 / 手順内容 / 所要時間 / 温度・火加減 / ポイント
  '調理手順': [ROLE_PICK, ROLE_AUTO, ROLE_IN, ROLE_IN, ROLE_IN, ROLE_IN, ROLE_IN],

  // 試作ID / 対象ID / 対象名 / 回数 / 試作日 / 甘 / 塩 / 酸 / 旨 / 苦 / 辛 / 総合評価 / 原価 / 所感 / 採用
  '試作記録': [ROLE_ID, ROLE_PICK, ROLE_AUTO, ROLE_ID, ROLE_IN, ROLE_IN, ROLE_IN, ROLE_IN, ROLE_IN, ROLE_IN,
               ROLE_IN, ROLE_AUTO, ROLE_AUTO, ROLE_IN, ROLE_PICK],

  // 試作ID / 子ID / 子名 / 使用量 / 単位 / 単価 / 原価 / 備考
  '試作レシピ明細': [ROLE_PICK, ROLE_PICK, ROLE_AUTO, ROLE_IN, ROLE_AUTO, ROLE_AUTO, ROLE_AUTO, ROLE_IN],

  // 試作ID / 手順番号 / 手順内容 / 所要時間 / 温度・火加減 / ポイント（作り方＝試作の回ごと）
  '試作手順': [ROLE_IN, ROLE_IN, ROLE_IN, ROLE_IN, ROLE_IN, ROLE_IN]
};

/**
 * 自動計算の列に入る数式（2行目に1つ入れて、下は自動でスピルする）。
 *
 * ★数式はここが唯一の置き場所です。各シートを作る関数もここを読みます。
 *  同じ数式を2か所に書かない＝片方だけ直して壊れる事故を防ぐためです。
 */
var AUTO_FORMULAS = {
  '仕入先マスタ': {
    8: '=ARRAYFORMULA(IF(A2:A="","",COUNTIF(食材マスタ!$H:$H,A2:A)))'
  },
  '食材マスタ': {
    // 単価 … kg単価が入っていれば不定貫として ÷1000、空なら定貫として 仕入価格÷入数
    17: '=ARRAYFORMULA(IF(A2:A="","",IFERROR(IF(M2:M<>"",M2:M/1000,L2:L/K2:K),"")))',
    // 実効単価 … 単価 ÷ 歩留まり率
    18: '=ARRAYFORMULA(IF(A2:A="","",IFERROR(Q2:Q/G2:G,"")))'
  },
  '仕込み品マスタ': {
    // 原価合計 = レシピ明細の「食材分原価」を親IDで合計
    7: '=ARRAYFORMULA(IF(A2:A="","",SUMIF(レシピ明細!$A:$A,A2:A,レシピ明細!$G:$G)))',
    // 出来上がり量(C列)が空なら、材料の合計をそのまま使う。
    // タレ類は「材料を混ぜただけ＝合計」なので、ふつうは入れなくてよい。
    // 炊く・煮詰めるなど量が変わるものだけ C列に実際の量を入れる。
    8: '=ARRAYFORMULA(IF(A2:A="","",IFERROR(G2:G/IF(C2:C="",SUMIF(レシピ明細!$A:$A,A2:A,レシピ明細!$E:$E),C2:C),"")))',
    // 仕込み品の中に仕込み品（3階層）が入っていないか見張る
    // ①3階層の警告 ②材料の入れ忘れ ③出来上がり量が空＝合計を使用中 ④入っていれば倍率
    9: '=ARRAYFORMULA(IF(A2:A="","",' +
       'IF(COUNTIFS(レシピ明細!$A:$A,A2:A,レシピ明細!$C:$C,"PRE*")>0,"⚠ 仕込み品の中の仕込み品は原価に含まれません",' +
       'IFERROR(IF(SUMIF(レシピ明細!$A:$A,A2:A,レシピ明細!$E:$E)=0,"レシピ明細に材料を入れてください",' +
       'IF(C2:C="","材料の合計 "&SUMIF(レシピ明細!$A:$A,A2:A,レシピ明細!$E:$E)&" をそのまま使っています",' +
       '"材料 "&SUMIF(レシピ明細!$A:$A,A2:A,レシピ明細!$E:$E)&" → できあがり "&C2:C&"（"&ROUND(C2:C/MAX(SUMIF(レシピ明細!$A:$A,A2:A,レシピ明細!$E:$E),1),2)&"倍）")),""))))'
  },
  '商品マスタ': {
    // 原価 = レシピ明細の「原価計」を商品IDで合計
    8:  '=ARRAYFORMULA(IF(A2:A="","",SUMIF(レシピ明細!$A:$A,A2:A,レシピ明細!$I:$I)))',
    9:  '=ARRAYFORMULA(IF(A2:A="","",IFERROR(H2:H/D2:D,"")))',
    10: '=ARRAYFORMULA(IF(A2:A="","",IFERROR(I2:I-E2:E,"")))'
  },
  'レシピ明細': {
    // 親名：MENUなら商品マスタ、それ以外は仕込み品マスタを引く
    2: '=ARRAYFORMULA(IF(A2:A="","",IFERROR(IF(LEFT(A2:A,4)="MENU",VLOOKUP(A2:A,商品マスタ!$A:$B,2,FALSE),VLOOKUP(A2:A,仕込み品マスタ!$A:$B,2,FALSE)),"⚠ID不明")))',
    // 子名：INGなら食材マスタ、PREなら仕込み品マスタを引く
    4: '=ARRAYFORMULA(IF(C2:C="","",IFERROR(IF(LEFT(C2:C,3)="ING",VLOOKUP(C2:C,食材マスタ!$A:$B,2,FALSE),VLOOKUP(C2:C,仕込み品マスタ!$A:$B,2,FALSE)),"⚠ID不明")))',
    6: '=ARRAYFORMULA(IF(C2:C="","",IFERROR(IF(LEFT(C2:C,3)="ING",VLOOKUP(C2:C,食材マスタ!$A:$D,4,FALSE),VLOOKUP(C2:C,仕込み品マスタ!$A:$D,4,FALSE)),"")))',
    // ★食材分(G)と仕込み分(H)を別の列にすることで循環参照を避けている。1列にまとめてはいけない
    7: '=ARRAYFORMULA(IF(C2:C="","",IF(LEFT(C2:C,3)="ING",IFERROR(E2:E*VLOOKUP(C2:C,食材マスタ!$A:$R,18,FALSE),0),0)))',
    8: '=ARRAYFORMULA(IF(C2:C="","",IF(LEFT(C2:C,3)="PRE",IFERROR(E2:E*VLOOKUP(C2:C,仕込み品マスタ!$A:$H,8,FALSE),0),0)))',
    9: '=ARRAYFORMULA(IF(C2:C="","",G2:G+H2:H))'
  },
  '調理手順': {
    // 仕込み品(PRE)の作り方もここに書けるよう、商品と仕込み品の両方を引く
    2: '=ARRAYFORMULA(IF(A2:A="","",IFERROR(IF(LEFT(A2:A,4)="MENU",VLOOKUP(A2:A,商品マスタ!$A:$B,2,FALSE),VLOOKUP(A2:A,仕込み品マスタ!$A:$B,2,FALSE)),"⚠ID不明")))'
  },
  '試作記録': {
    3:  '=ARRAYFORMULA(IF(B2:B="","",IFERROR(IF(LEFT(B2:B,4)="MENU",VLOOKUP(B2:B,商品マスタ!$A:$B,2,FALSE),VLOOKUP(B2:B,仕込み品マスタ!$A:$B,2,FALSE)),"⚠ID不明")))',
    12: OVERALL_FORMULA,
    13: '=ARRAYFORMULA(IF(A2:A="","",SUMIF(試作レシピ明細!$A:$A,A2:A,試作レシピ明細!$G:$G)))'
  },
  '試作レシピ明細': {
    3: '=ARRAYFORMULA(IF(B2:B="","",IFERROR(IF(LEFT(B2:B,3)="ING",VLOOKUP(B2:B,食材マスタ!$A:$B,2,FALSE),VLOOKUP(B2:B,仕込み品マスタ!$A:$B,2,FALSE)),"⚠ID不明")))',
    5: '=ARRAYFORMULA(IF(B2:B="","",IFERROR(IF(LEFT(B2:B,3)="ING",VLOOKUP(B2:B,食材マスタ!$A:$D,4,FALSE),VLOOKUP(B2:B,仕込み品マスタ!$A:$D,4,FALSE)),"")))',
    6: '=ARRAYFORMULA(IF(B2:B="","",IFERROR(IF(LEFT(B2:B,3)="ING",VLOOKUP(B2:B,食材マスタ!$A:$R,18,FALSE),VLOOKUP(B2:B,仕込み品マスタ!$A:$H,8,FALSE)),"")))',
    7: '=ARRAYFORMULA(IF(B2:B="","",IFERROR(D2:D*F2:F,"")))'
  }
};

/** 役割ごとの見た目と説明 */
var ROLE_STYLE = {};
ROLE_STYLE[ROLE_IN]   = { mark: '',    head: '#F3EADB', headFont: '#2A2019', body: '#FFFFFF',
                          label: '✋ 入力する',      note: '✋ ここは手で入力する列です。' };
ROLE_STYLE[ROLE_PICK] = { mark: ' ▼',  head: '#DCE9F3', headFont: '#22333F', body: '#F6FAFD',
                          label: '▼ 選ぶ',          note: '▼ ここはプルダウンから選ぶ列です。\n手で打たずに、▼を押して選んでください。' };
ROLE_STYLE[ROLE_AUTO] = { mark: ' 🔒', head: '#E4DED3', headFont: '#6E655C', body: C_AUTO_BG,
                          label: '🔒 自動計算',     note: '🔒 ここは自動で計算される列です。\n入力しても自動で元に戻ります（消えたように見えても壊れていません）。' };
ROLE_STYLE[ROLE_ID]   = { mark: ' 🔑', head: '#EBE4D9', headFont: '#6E655C', body: '#FCFAF6',
                          label: '🔑 自動で振る',   note: '🔑 ここは自動で振られるIDです。\n名前を入れると自動で入ります。手で書き換えることもできます。' };

var ROLE_MARKS   = [' ▼', ' 🔒', ' 🔑'];

/** 「台帳を整える」で直したい見出し（列の意味が変わったもの） */
var FIX_HEADERS = {
  '調理手順': { 1: '商品ID / 仕込みID' }
};

/**
 * 間違えやすい列に付ける説明。ヘッダーにマウスを乗せると出ます。
 * ★炊く・茹でるなど「重さが増えるもの」の扱いを、ここで一度きちんと説明しています。
 */
var EXTRA_NOTES = {
  '食材マスタ': {
    7: '「捨てる部分」を引いたあとに残る割合です。\n例：キャベツの芯を落として80%残るなら 80%。\n\n' +
       '★ここに入るのは必ず100%以下です。\n' +
       '　「炊くと2倍になる」「茹でると増える」ものは、ここでは表せません。\n' +
       '　ここに 50% などと入れると、原価が2倍になってしまいます。\n\n' +
       '　ご飯・茹でた麺・戻した乾物は「仕込み品マスタ」に登録してください。',
    19: '★味の強さ＝この食材 1g がその味をどれだけ出すか（甘味）。\n' +
        '　試作の味の初期値（スライダー）を自動で出すために使います。\n' +
        '　甘味・塩味は「糖分%・塩分%」が目安（例：砂糖100／醤油の塩味16）。\n' +
        '　酸味・旨味・苦味・辛味は“おおよその目安”で構いません。\n\n' +
        '　空欄＝0 として扱います。よく使う調味料は最初から入れてあります。\n' +
        '　実際に作って味が違えば、ここを直すと次からの初期値が賢くなります。',
    20: '塩味の強さ（1gあたり）。塩分% が目安。例：塩99／醤油16／味噌12。空欄＝0。',
    21: '酸味の強さ（1gあたり）。酢・レモンなど。空欄＝0。',
    22: '旨味の強さ（1gあたり）。だし・醤油・肉・貝など。空欄＝0。',
    23: '苦味の強さ（1gあたり）。※総合評価の計算には入りません。空欄＝0。',
    24: '辛味の強さ（1gあたり）。唐辛子・こしょうなど。※総合評価には入りません。空欄＝0。'
  },
  '仕込み品マスタ': {
    3: '★ふつうは空のままで構いません。\n' +
       '　空にしておくと「材料の合計」がそのまま使われます。\n' +
       '　（例：醤油400＋みりん300＋酒200＋生姜100 → 1000g）\n\n' +
       '入れるのは、作ると量が変わるものだけです。\n' +
       '　炊く … 生米1000g → ご飯2200g なら 2200\n' +
       '　煮詰める … 材料1000g → 700g なら 700\n\n' +
       'ここが「炊き上がり倍率」を決める唯一の場所です。\n' +
       '実際に作って量ってみて違っていたら、ここだけ直せば\n' +
       'それを使っている全部の商品に反映されます。\n\n' +
       '※いまどうなっているかは、右の「チェック」列に出ています。',
    10: '仕込みにかかる時間の目安（例：40分）。覚え書きです。計算には使いません。',
    11: 'この仕込み品のレシピを作った人の名前（例：津原）。覚え書きです。',
    12: 'このレシピを作った（決めた）日。覚え書きです。'
  },
  'レシピ明細': {
    5: '1食ぶんに使う量です（捨てる部分を除いた正味）。\n\n' +
       '★ご飯のような仕込み品を使うときは、盛り付ける量（炊いた後の量）で書きます。\n' +
       '　原価は自動で生米に戻して計算されるので、生米の量を書く必要はありません。'
  },
  '商品マスタ': {
    11: '本部の商品コード・POSコードなど（任意）。台帳の中で使う「商品ID」とは別に、\nお店で使っている番号を控えておく欄です。空のままでも構いません。',
    12: '1食ぶんの仕上がり量の目安（例：250g、1杯）。覚え書きです。計算には使いません。',
    13: '作る（仕込む）のにかかる時間の目安（例：15分）。覚え書きです。',
    14: 'このレシピを作った人の名前（例：津原）。覚え書きです。',
    15: 'このレシピを作った（決めた）日。覚え書きです。'
  }
};
var PROTECT_TAG  = '【自動計算】';
var GUIDE_SHEET  = '入力ガイド';

/**
 * ★本当にデータが入っている最後の行を返す。
 *
 * getLastRow() は使ってはいけない。
 * 自動計算列の ARRAYFORMULA が全行に ""(空文字) を返すため、
 * getLastRow() は 1001行目あたりを「最終行」と答えてしまう。
 * そこに +1 して書き込むと、画面のはるか下に行が飛んでしまう
 * （2026-09-19 実際に発生。仕込み品が見えない場所に登録された）。
 *
 * col … 見張る列。ふつうはID列(1)。下から順に探すので、
 *       離れた場所に飛んでしまった行もちゃんと拾える。
 */
function lastDataRow_(sh, col) {
  col = col || 1;
  var last = sh.getLastRow();
  if (last < 2) return 1;
  var vals = sh.getRange(2, col, last - 1, 1).getValues();
  for (var i = vals.length - 1; i >= 0; i--) {
    if (String(vals[i][0]).trim() !== '') return i + 2;
  }
  return 1;
}

/** 色や保護をかける範囲(LAST_ROW行)が足りるように、行数を確保しておく */
function ensureRows_(sh) {
  var need = LAST_ROW + 1 - sh.getMaxRows();
  if (need > 0) sh.insertRowsAfter(sh.getMaxRows(), need);
}

/** 列が足りなければ右に足す（既存データは動かさない） */
function ensureCols_(sh, need) {
  var have = sh.getMaxColumns();
  if (need > have) sh.insertColumnsAfter(have, need - have);
}

/** ヘッダー名から役割の印を取り除く（何度実行しても印が増えないようにするため） */
function stripMark_(s) {
  var t = String(s);
  for (var i = 0; i < ROLE_MARKS.length; i++) {
    while (t.slice(-ROLE_MARKS[i].length) === ROLE_MARKS[i]) t = t.slice(0, -ROLE_MARKS[i].length);
  }
  return t.replace(/\s+$/, '');
}

/**
 * 空行をはさんで下に飛んでしまった行を、上に詰め直す。
 *
 * 「台帳を整える」を押すたびに走る自己修復。
 * すでに上から詰まっていれば何もしない（書き込みもしない）。
 */
function compactSheet_(sh) {
  var roles = COLUMN_ROLES[sh.getName()];
  if (!roles) return 0;

  var cols = inputColsOf_(sh.getName());
  if (cols.length === 0) return 0;

  var last = sh.getLastRow();
  if (last < 3) return 0;

  var grid = sh.getRange(2, 1, last - 1, roles.length).getValues();
  var keep = [], firstEmpty = -1;
  for (var r = 0; r < grid.length; r++) {
    var row = cols.map(function (c) { return grid[r][c - 1]; });
    var has = row.some(function (v) { return String(v).trim() !== ''; });
    if (has) keep.push(row);
    else if (firstEmpty < 0) firstEmpty = r;
  }

  // 空行より下にデータが無い＝すでに詰まっている
  if (firstEmpty < 0 || keep.length === firstEmpty) return 0;

  cols.forEach(function (c) { sh.getRange(2, c, grid.length, 1).clearContent(); });
  cols.forEach(function (c, i) {
    sh.getRange(2, c, keep.length, 1).setValues(keep.map(function (x) { return [x[i]]; }));
  });
  return keep.length - firstEmpty;      // 何行ぶん引き上げたか
}

/** AUTO_FORMULAS に書いてある数式を、そのシートの2行目に入れ直す */
function applyAutoFormulas_(sh) {
  var map = AUTO_FORMULAS[sh.getName()];
  if (!map) return;
  Object.keys(map).forEach(function (col) {
    sh.getRange(2, Number(col)).setFormula(map[col]);
  });
}


/* ========== ① 見て分かるようにする（色とヘッダーの印） ========== */

function applyRoles_(sh) {
  var roles = COLUMN_ROLES[sh.getName()];
  if (!roles) return;
  ensureRows_(sh);
  ensureCols_(sh, roles.length);   // 列が足りないと getRange で落ちるので先に確保

  var heads = sh.getRange(1, 1, 1, roles.length).getValues()[0];
  var newHeads = [], bgs = [], fonts = [];

  var fix   = FIX_HEADERS[sh.getName()] || {};
  var notes = EXTRA_NOTES[sh.getName()] || {};

  for (var i = 0; i < roles.length; i++) {
    var st = ROLE_STYLE[roles[i]];
    newHeads.push((fix[i + 1] || stripMark_(heads[i])) + st.mark);
    bgs.push(st.head);
    fonts.push(st.headFont);
    // ヘッダーのメモ … 既に別の説明が入っている列は、その説明を残して先頭に役割を足す
    var cell = sh.getRange(1, i + 1);
    var old  = cell.getNote() || '';
    var body = notes[i + 1] || old.replace(/^[✋▼🔒🔑][\s\S]*?(\n\n|$)/, '');   // 前回付けた役割メモを外す
    cell.setNote(body ? st.note + '\n\n' + body : st.note);
    // データ範囲の背景
    sh.getRange(2, i + 1, LAST_ROW, 1).setBackground(st.body);
  }

  sh.getRange(1, 1, 1, roles.length)
    .setValues([newHeads]).setBackgrounds([bgs]).setFontColors([fonts])
    .setFontWeight('bold').setVerticalAlignment('middle').setWrap(true);
  sh.setRowHeight(1, 40);
  sh.setFrozenRows(1);
}


/* ========== ② 触ると警告が出るようにする ========== */

function applyProtection_(sh) {
  var roles = COLUMN_ROLES[sh.getName()];
  if (!roles) return;
  ensureRows_(sh);
  ensureCols_(sh, roles.length);

  // 前回このスクリプトが付けた保護を外してから付け直す（何度実行しても増えない）
  sh.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function (p) {
    var d = p.getDescription() || '';
    if (d.indexOf(PROTECT_TAG) === 0) { try { p.remove(); } catch (e) {} }
  });

  // 自動計算の列を、続きの固まりごとにまとめて保護する
  var start = -1;
  for (var i = 0; i <= roles.length; i++) {
    var isAuto = (i < roles.length && roles[i] === ROLE_AUTO);
    if (isAuto && start < 0) start = i;
    if (!isAuto && start >= 0) {
      sh.getRange(1, start + 1, LAST_ROW + 1, i - start)
        .protect()
        .setDescription(PROTECT_TAG + sh.getName() + ' は自動で計算される列です。入力しても元に戻ります。')
        .setWarningOnly(true);
      start = -1;
    }
  }
}


/* ========== ③ それでも入力されたら自動で元に戻す ========== */

function onEdit(e) {
  if (!e || !e.range) return;
  var sh, name;
  try { sh = e.range.getSheet(); name = sh.getName(); } catch (err) { return; }

  try {
    if (name === '試作比較') { normalizeCompareTarget_(e, sh); return; }
    if (!COLUMN_ROLES[name]) return;
    normalizePickedId_(e, sh, name);   // 「ING-0001 豚ロース」→「ING-0001」に直す
    restoreAuto_(e, sh, name);         // 自動計算の列に入ったものを元に戻す
    autoNumberId_(e, sh, name);        // 名前を入れたらIDを自動で振る
    autoFillTrial_(e, sh, name);       // 試作記録の回数・日付を自動で入れる
    guardYieldRate_(e, sh, name);      // 歩留まり率の入れ間違いを知らせる
  } catch (err) {
    try {
      SpreadsheetApp.getActiveSpreadsheet()
        .toast('自動処理でエラーが出ました: ' + err.message, '⚠ 松山さんへ', 10);
    } catch (e2) {}
  }
}

/** 自動計算の列に入力されたら、数式を戻し、入力されたものを消す */
function restoreAuto_(e, sh, name) {
  var roles = COLUMN_ROLES[name];
  var forms = AUTO_FORMULAS[name] || {};
  var rng = e.range;
  var c1 = rng.getColumn(), c2 = rng.getLastColumn();
  var r1 = rng.getRow(),    r2 = rng.getLastRow();
  if (r2 > LAST_ROW + 200) r2 = LAST_ROW + 200;          // 行全体の貼り付け対策

  var fixed = [];
  for (var c = c1; c <= c2 && c <= roles.length; c++) {
    if (roles[c - 1] !== ROLE_AUTO) continue;
    var f = forms[c];
    if (!f) continue;                                     // 数式ではない自動列は触らない
    if (r1 <= 2 && r2 >= 2) sh.getRange(2, c).setFormula(f);   // 数式そのものを壊した場合は戻す
    var from = Math.max(r1, 3);
    if (r2 >= from) sh.getRange(from, c, r2 - from + 1, 1).clearContent();
    fixed.push(stripMark_(sh.getRange(1, c).getValue()));
  }

  if (fixed.length > 0) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      '「' + fixed.join('」「') + '」は自動で計算される列です。\n入力は取り消して元に戻しました。',
      '🔒 ここは入力する場所ではありません', 8);
  }
}

/**
 * 歩留まり率の入れ間違いを、その場で知らせる。
 *
 * ★よくある間違い：「米は炊くと2倍になるから歩留まり率50%」
 *  歩留まり率は「捨てる部分」を表す列で、実効単価＝単価÷歩留まり率。
 *  50%を入れると原価が2倍になります。増えるものは仕込み品で扱ってください。
 */
function guardYieldRate_(e, sh, name) {
  if (name !== '食材マスタ') return;
  var rng = e.range;
  if (rng.getColumn() > 7 || rng.getLastColumn() < 7) return;

  var r1 = Math.max(rng.getRow(), 2), r2 = rng.getLastRow();
  if (r2 < r1 || r2 - r1 > 200) return;

  var vals = sh.getRange(r1, 7, r2 - r1 + 1, 1).getValues();
  for (var i = 0; i < vals.length; i++) {
    var v = Number(vals[i][0]);
    if (!v) continue;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (v > 1) {
      ss.toast('歩留まり率は100%までです（捨てる部分を引いて残る割合）。\n' +
               '炊く・茹でるなど「増えるもの」は、仕込み品マスタで登録してください。',
               '⚠ 歩留まり率', 12);
      return;
    }
    if (v < 0.65) {
      ss.toast('捨てる部分が ' + Math.round((1 - v) * 100) + '% になります。合っていますか?\n\n' +
               '「炊くと2倍になる」ようなものをここに入れると、原価が狂います。\n' +
               'ご飯・茹で麺・戻した乾物は、仕込み品マスタで登録してください。',
               '⚠ 歩留まり率の確認', 12);
      return;
    }
  }
}

/** 試作比較の「対象を選ぶ」欄で名前つきを選んだとき、IDだけ残す */
function normalizeCompareTarget_(e, sh) {
  var rng = e.range;
  if (rng.getRow() !== 1 || rng.getColumn() !== 2) return;
  var m = String(rng.getValue()).match(/^((?:MENU|PRE)-\d+)[\s　]/);
  if (m) rng.setValue(m[1]);
}

/** プルダウンで「ING-0001 豚ロース」を選んだとき、IDだけ残す */
function normalizePickedId_(e, sh, name) {
  var roles = COLUMN_ROLES[name];
  var rng = e.range;
  var c1 = rng.getColumn(), c2 = rng.getLastColumn();
  var r1 = rng.getRow(),    r2 = rng.getLastRow();
  if (r1 < 2) r1 = 2;
  if (r2 < r1) return;
  if (r2 - r1 > 500) return;                              // 大量貼り付けのときは何もしない

  for (var c = c1; c <= c2 && c <= roles.length; c++) {
    if (roles[c - 1] !== ROLE_PICK && roles[c - 1] !== ROLE_ID) continue;
    var cells = sh.getRange(r1, c, r2 - r1 + 1, 1);
    var vals = cells.getValues();
    var hit = false;
    for (var i = 0; i < vals.length; i++) {
      var m = String(vals[i][0]).match(/^((?:ING|PRE|MENU|TRY|SUP)-\d+)[\s　].*/);
      if (m) { vals[i][0] = m[1]; hit = true; }
    }
    if (hit) cells.setValues(vals);
  }
}

/** 名前を入れたらIDを自動で振る（一度振ったIDは変えない） */
var ID_RULES = {
  '食材マスタ':     { nameCol: 2, idCol: 1, prefix: 'ING-' },
  '仕込み品マスタ': { nameCol: 2, idCol: 1, prefix: 'PRE-' },
  '商品マスタ':     { nameCol: 2, idCol: 1, prefix: 'MENU-' },
  '仕入先マスタ':   { nameCol: 2, idCol: 1, prefix: 'SUP-' }
};

function autoNumberId_(e, sh, name) {
  var rule = ID_RULES[name];
  if (!rule) return;
  var rng = e.range;
  if (rng.getColumn() > rule.nameCol || rng.getLastColumn() < rule.nameCol) return;

  var r1 = Math.max(rng.getRow(), 2), r2 = rng.getLastRow();
  if (r2 < r1 || r2 - r1 > 500) return;

  var names = sh.getRange(r1, rule.nameCol, r2 - r1 + 1, 1).getValues();
  var idsRg = sh.getRange(r1, rule.idCol,   r2 - r1 + 1, 1);
  var ids   = idsRg.getValues();

  var next = nextIdNumber_(sh, rule);
  var changed = false;
  for (var i = 0; i < names.length; i++) {
    if (String(names[i][0]).trim() === '') continue;
    if (String(ids[i][0]).trim() !== '') continue;         // 既にIDがあるものは触らない
    ids[i][0] = rule.prefix + zero4_(next);
    next++;
    changed = true;
  }
  if (changed) {
    idsRg.setValues(ids);
    SpreadsheetApp.getActiveSpreadsheet().toast('IDを自動で振りました。', '🔑 ' + name, 4);
  }
}

function nextIdNumber_(sh, rule) {
  var last = Math.max(lastDataRow_(sh, rule.idCol), 1);
  var vals = sh.getRange(1, rule.idCol, last, 1).getValues();
  var re = new RegExp('^' + rule.prefix.replace('-', '\\-') + '(\\d+)$');
  var max = 0;
  for (var i = 0; i < vals.length; i++) {
    var m = String(vals[i][0]).trim().match(re);
    if (m) max = Math.max(max, Number(m[1]));
  }
  return max + 1;
}

function zero4_(n) { return ('0000' + n).slice(-4); }

/** 試作記録に対象IDを入れたら、試作ID・回数・試作日を自動で入れる */
function autoFillTrial_(e, sh, name) {
  if (name !== '試作記録') return;
  var rng = e.range;
  if (rng.getColumn() > 2 || rng.getLastColumn() < 2) return;

  var r1 = Math.max(rng.getRow(), 2), r2 = rng.getLastRow();
  if (r2 < r1 || r2 - r1 > 200) return;

  var all = sh.getDataRange().getValues();
  var maxTry = 次の試作ID_(all).match(/(\d+)$/)[1] * 1;

  for (var r = r1; r <= r2; r++) {
    var target = String(sh.getRange(r, 2).getValue()).trim();
    if (!target) continue;
    if (String(sh.getRange(r, 1).getValue()).trim() !== '') continue;   // 既にIDがあるなら触らない

    // その対象の今までの最大回数を数える
    var no = 0;
    for (var i = 1; i < all.length; i++) {
      if (String(all[i][1]).trim() === target) no = Math.max(no, Number(all[i][3]) || 0);
    }
    sh.getRange(r, 1).setValue('TRY-' + zero4_(maxTry));
    sh.getRange(r, 4).setValue(no + 1);
    if (String(sh.getRange(r, 5).getValue()).trim() === '') sh.getRange(r, 5).setValue(new Date());
    all.push(['TRY-' + zero4_(maxTry), target, '', no + 1]);
    maxTry++;
  }
}


/* ========== IDのプルダウン（食材名で選べるようにする） ========== */

/**
 * 「ING-0001 豚ロース」の形の選択肢を裏方の列に作り、ID列のプルダウンにする。
 * 選んだ瞬間に onEdit が「ING-0001」だけに直すので、数式は今まで通り動く。
 */
var PICK_LISTS = {
  'レシピ明細':     { 1: 'parent', 3: 'child' },
  '調理手順':       { 1: 'parent' },   // 商品だけでなく仕込み品の作り方も書ける
  '試作記録':       { 2: 'parent' },
  '試作レシピ明細': { 2: 'child' },
  '食材マスタ':     { 8: 'supplier' }
};

// 裏方の選択肢リストを置く列（表示しているどの列よりも右）
var PICK_SRC_COL = {
  'レシピ明細': 20, '調理手順': 20, '試作記録': 20, '試作レシピ明細': 20, '食材マスタ': 22
};

var PICK_SRC_FORMULA = {
  // 商品と仕込み品（レシピの「親」になれるもの）
  parent: '=IFERROR(SORT(FILTER({商品マスタ!$A$2:$A&" "&商品マスタ!$B$2:$B;仕込み品マスタ!$A$2:$A&" "&仕込み品マスタ!$B$2:$B},' +
          '{商品マスタ!$A$2:$A;仕込み品マスタ!$A$2:$A}<>"")),"")',
  // 食材と仕込み品（レシピの「子」になれるもの）
  child:  '=IFERROR(SORT(FILTER({食材マスタ!$A$2:$A&" "&食材マスタ!$B$2:$B;仕込み品マスタ!$A$2:$A&" "&仕込み品マスタ!$B$2:$B},' +
          '{食材マスタ!$A$2:$A;仕込み品マスタ!$A$2:$A}<>"")),"")',
  // 商品だけ
  menu:   '=IFERROR(SORT(FILTER(商品マスタ!$A$2:$A&" "&商品マスタ!$B$2:$B,商品マスタ!$A$2:$A<>"")),"")',
  // 仕入先
  supplier: '=IFERROR(SORT(FILTER(仕入先マスタ!$A$2:$A&" "&仕入先マスタ!$B$2:$B,仕入先マスタ!$A$2:$A<>"")),"")'
};

function buildPickers_(sh) {
  var name = sh.getName();
  var lists = PICK_LISTS[name];
  if (!lists) return;

  ensureRows_(sh);
  var srcCol = PICK_SRC_COL[name];
  var kinds = [];
  Object.keys(lists).forEach(function (col) {
    if (kinds.indexOf(lists[col]) < 0) kinds.push(lists[col]);
  });

  // 裏方の列を用意して隠す
  var need = srcCol + kinds.length - 1;
  if (sh.getMaxColumns() < need) sh.insertColumnsAfter(sh.getMaxColumns(), need - sh.getMaxColumns());

  kinds.forEach(function (kind, i) {
    var c = srcCol + i;
    sh.getRange(1, c).setFormula(PICK_SRC_FORMULA[kind]);
    sh.getRange(1, c).setNote('プルダウンの選択肢を作っている裏方の列です。消さないでください。');
    sh.hideColumns(c);
  });

  Object.keys(lists).forEach(function (col) {
    var c = srcCol + kinds.indexOf(lists[col]);
    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(sh.getRange(1, c, LAST_ROW, 1), true)
      .setAllowInvalid(true)          // 選んだあとIDだけに直すので、これは true でないといけない
      .build();
    sh.getRange(2, Number(col), LAST_ROW, 1).setDataValidation(rule);
  });
}


/* ========== 入力ガイド（どこに入力してよいかの一覧表） ========== */

function build入力ガイド(ss) {
  var sh = ss.getSheetByName(GUIDE_SHEET) || ss.insertSheet(GUIDE_SHEET);
  resetSheet_(sh);

  sh.setColumnWidth(1, 150);
  sh.setColumnWidth(2, 50);
  sh.setColumnWidth(3, 190);
  sh.setColumnWidth(4, 130);
  sh.setColumnWidth(5, 460);

  var rows = [];
  rows.push(['入力ガイド', '', '', '', '']);
  rows.push(['このシートは見るだけです。どこに入力してよいか迷ったらここを見てください。', '', '', '', '']);
  rows.push(['', '', '', '', '']);
  rows.push(['【印の意味】', '', '', '', '']);
  rows.push(['', '', '✋ 入力する',    '白いセル',         '自分で打ち込む場所です。']);
  rows.push(['', '', '▼ 選ぶ',        '薄い青のセル',     'セルを選ぶと▼が出ます。押して一覧から選んでください。手で打たなくて構いません。']);
  rows.push(['', '', '🔒 自動計算',   '薄いベージュのセル', '触らないでください。入力しても自動で消えて元に戻ります（壊れてはいません）。']);
  rows.push(['', '', '🔑 自動で振る', '薄いベージュのセル', 'IDです。名前を入れると自動で入ります。']);
  rows.push(['', '', '', '', '']);
  rows.push(['シート', '列', '見出し', '区分', '中身']);

  var headerRow = rows.length;

  SHEET_NAMES.forEach(function (name) {
    var s = ss.getSheetByName(name);
    var roles = COLUMN_ROLES[name];
    if (!s) return;

    if (!roles) {
      // 試作比較のように、シート全部が見るだけのもの
      rows.push([name, '', '（シート全体）', '🔒 見るだけ',
                 '入力する場所ではありません。B1で対象を選ぶと結果が出ます。']);
      return;
    }
    var heads = s.getRange(1, 1, 1, roles.length).getValues()[0];
    for (var i = 0; i < roles.length; i++) {
      var st = ROLE_STYLE[roles[i]];
      rows.push([
        i === 0 ? name : '',
        columnLetter_(i + 1),
        stripMark_(heads[i]),
        st.label,
        st.note.replace(/^[✋▼🔒🔑]\s*/, '').split('\n')[0]
      ]);
    }
    rows.push(['', '', '', '', '']);
  });

  sh.getRange(1, 1, rows.length, 5).setValues(rows);

  sh.getRange('A1').setFontSize(16).setFontWeight('bold');
  sh.getRange('A2').setFontColor('#6E655C');
  sh.getRange('A4').setFontWeight('bold');
  sh.getRange(5, 3, 4, 1).setFontWeight('bold');
  sh.getRange(headerRow, 1, 1, 5).setFontWeight('bold').setBackground(C_HEADER_BG);
  sh.getRange(headerRow + 1, 1, rows.length - headerRow, 1).setFontWeight('bold');
  sh.getRange(1, 1, rows.length, 5).setVerticalAlignment('top').setWrap(true);
  sh.setFrozenRows(headerRow);

  // 区分ごとに色を付けて、一覧の上でも見分けられるようにする
  var rules = [];
  [[ROLE_IN, '✋ 入力する'], [ROLE_PICK, '▼ 選ぶ'], [ROLE_AUTO, '🔒 自動計算'], [ROLE_ID, '🔑 自動で振る']]
    .forEach(function (x) {
      rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(x[1])
        .setBackground(ROLE_STYLE[x[0]].body)
        .setRanges([sh.getRange(headerRow + 1, 4, rows.length, 1)]).build());
    });
  sh.setConditionalFormatRules(rules);

  sh.protect().setDescription(PROTECT_TAG + '入力ガイドは見るだけのシートです').setWarningOnly(true);

  var pos = ss.getNumSheets();
  ss.setActiveSheet(sh);
  ss.moveActiveSheet(pos);
  return sh;
}


/* ========== 台帳を整える（色分け・保護・プルダウン・ガイドをまとめて適用） ========== */

function 台帳を整える() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var missing = SHEET_NAMES.filter(function (n) { return ss.getSheetByName(n) === null; });
  if (missing.length > 0) {
    ui.alert('中止しました。\n\n次のシートが見つかりません:\n  ' + missing.join('\n  '));
    return;
  }

  if (ui.alert('台帳を整えます',
      '次のことを行います。入力済みのデータは消えません。\n\n' +
      '・入力する列／選ぶ列／自動計算の列を色分けする\n' +
      '・自動計算の列に保護をかける（触ると警告が出ます）\n' +
      '・IDを食材名で選べるプルダウンにする\n' +
      '・「入力ガイド」シートを作り直す\n' +
      '・試作比較の分量の計算を直す\n\n' +
      'よろしいですか?', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;

  var moved = [];
  SHEET_NAMES.forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (!sh || !COLUMN_ROLES[name]) return;
    var n = compactSheet_(sh);          // 下に飛んでしまった行があれば上に詰め直す
    if (n > 0) moved.push('  ' + name + '：' + n + '行');
    applyAutoFormulas_(sh);
    buildPickers_(sh);
    applyRoles_(sh);
    applyProtection_(sh);
  });

  build試作比較(ss);        // 分量の計算を直した版で作り直す
  build入力ガイド(ss);

  ss.setActiveSheet(ss.getSheetByName(GUIDE_SHEET));
  ui.alert('整えました。\n\n' +
           (moved.length > 0
             ? '離れた場所にあった行を、上に詰め直しました:\n' + moved.join('\n') + '\n\n'
             : '') +
           '「入力ガイド」シートに、どの列が入力してよい場所かの一覧を作りました。\n\n' +
           '・白いセル … 入力する\n' +
           '・薄い青  … ▼から選ぶ\n' +
           '・ベージュ … 自動計算（入力しても元に戻ります）');
}


/* =====================================================================
 *  登録フォーム（商品と仕込み品、どちらもこの1画面で作れる）
 *
 *  ★「炊くと増えるもの」の扱い：
 *   ご飯・茹で麺・戻した乾物のように、作ると重さが増えるものは
 *   食材マスタの歩留まり率では表せません（歩留まり率は「捨てる部分」＝1.0以下のため）。
 *   そういうものは仕込み品として登録します。
 *     例）炊いたご飯 … 材料に 米(精白) 1000g、出来上がり量 2200g
 *         → 単位原価が「ご飯1gあたり」になり、商品側は盛る量(240g)で書ける。
 *   倍率を直したくなったら、仕込み品マスタの出来上がり量を1か所だけ直せば
 *   その仕込み品を使う全部の商品に効きます。
 * ===================================================================== */

function メニュー登録フォームを開く() {
  var html = HtmlService.createHtmlOutput(menuFormHtml_())
    .setWidth(980).setHeight(780);
  SpreadsheetApp.getUi().showModalDialog(html, '🍳 商品・仕込み品を登録する');
}

/** フォームが最初に読み込むデータ（google.script.run から呼ぶので末尾に _ を付けない） */
function menuFormData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var items = [];
  readRows_(ss, '食材マスタ', 18).forEach(function (r) {
    if (r[0]) items.push({ id: String(r[0]), name: String(r[1]), unit: String(r[3]), cost: Number(r[17]) || 0, kind: '食材' });
  });
  readRows_(ss, '仕込み品マスタ', 8).forEach(function (r) {
    if (r[0]) items.push({ id: String(r[0]), name: String(r[1]), unit: String(r[3]), cost: Number(r[7]) || 0, kind: '仕込み品' });
  });

  var targets = [];
  readRows_(ss, '商品マスタ', 10).forEach(function (r) {
    if (r[0]) targets.push({ id: String(r[0]), name: String(r[1]), kind: 'menu' });
  });
  readRows_(ss, '仕込み品マスタ', 8).forEach(function (r) {
    if (r[0]) targets.push({ id: String(r[0]), name: String(r[1]), kind: 'prep' });
  });

  return { items: items, targets: targets, units: OPT_在庫単位, stores: OPT_保存区分 };
}

/** 既存の商品／仕込み品を読み込む */
function menuFormLoad(id) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var isPrep = (String(id).indexOf('PRE-') === 0);
  var head = null;

  if (isPrep) {
    readRows_(ss, '仕込み品マスタ', 8).forEach(function (r) {
      if (String(r[0]) === id) {
        head = { kind: 'prep', id: String(r[0]), name: String(r[1]),
                 yieldQty: Number(r[2]) || 0, unit: String(r[3]),
                 store: String(r[4]), days: Number(r[5]) || 0 };
      }
    });
  } else {
    readRows_(ss, '商品マスタ', 10).forEach(function (r) {
      if (String(r[0]) === id) {
        head = { kind: 'menu', id: String(r[0]), name: String(r[1]), category: String(r[2]),
                 price: Number(r[3]) || 0, rate: Number(r[4]) || 0 };
      }
    });
  }
  if (!head) throw new Error(id + ' が見つかりません。');

  var lines = [];
  readRows_(ss, 'レシピ明細', 10).forEach(function (r) {
    if (String(r[0]) === id && String(r[2])) {
      lines.push({ id: String(r[2]), qty: Number(r[4]) || 0, memo: String(r[9]) });
    }
  });

  var steps = [];
  readRows_(ss, '調理手順', 7).forEach(function (r) {
    if (String(r[0]) === id) {
      steps.push({ no: Number(r[2]) || 0, text: String(r[3]), time: String(r[4]), heat: String(r[5]), point: String(r[6]) });
    }
  });
  steps.sort(function (a, b) { return a.no - b.no; });

  return { head: head, lines: lines, steps: steps };
}

/** フォームの内容をシートに書き込む */
function menuFormSave(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var repSh  = ss.getSheetByName('レシピ明細');
  var stepSh = ss.getSheetByName('調理手順');
  var isPrep = (payload.kind === 'prep');

  var name = String(payload.name || '').trim();
  if (!name) throw new Error((isPrep ? '仕込み品名' : '商品名') + 'を入れてください。');
  if (!payload.lines || payload.lines.length === 0) throw new Error('材料を1つ以上入れてください。');

  // ① 見出しのシート（商品マスタ or 仕込み品マスタ）
  var headSh = ss.getSheetByName(isPrep ? '仕込み品マスタ' : '商品マスタ');
  var rule   = ID_RULES[isPrep ? '仕込み品マスタ' : '商品マスタ'];
  var id = String(payload.id || '').trim();
  var rowNo = 0;
  var ids = headSh.getRange(2, 1, Math.max(lastDataRow_(headSh) - 1, 1), 1).getValues();

  if (id) {
    for (var i = 0; i < ids.length; i++) if (String(ids[i][0]).trim() === id) rowNo = i + 2;
    if (!rowNo) throw new Error(id + ' が見つかりません。');
  } else {
    id = rule.prefix + zero4_(nextIdNumber_(headSh, rule));
    rowNo = lastDataRow_(headSh) + 1;
    headSh.getRange(rowNo, 1).setValue(id);
  }

  if (isPrep) {
    // 仕込み品マスタ … B仕込み品名 C出来上がり量 D単位 E保存区分 F日持ち日数
    headSh.getRange(rowNo, 2, 1, 5).setValues([[
      name, Number(payload.yieldQty) || 0, String(payload.unit || 'g'),
      String(payload.store || ''), Number(payload.days) || 0
    ]]);
  } else {
    // 商品マスタ … B商品名 Cカテゴリ D売価 E目標原価率
    headSh.getRange(rowNo, 2, 1, 4).setValues([[
      name, String(payload.category || ''), Number(payload.price) || 0, Number(payload.rate) || 0
    ]]);
  }

  // ② レシピ明細 … 行は消さず、入力列(A/C/E/J)だけ組み直す
  var lines = payload.lines
    .filter(function (L) { return String(L.id || '').trim() !== ''; })
    .map(function (L) { return [id, String(L.id).trim(), Number(L.qty) || 0, String(L.memo || '')]; });

  // 自分自身を材料にすると循環参照でシートが壊れるので止める
  for (var k = 0; k < lines.length; k++) {
    if (lines[k][1] === id) throw new Error('自分自身を材料にはできません（' + id + '）。');
  }
  // 仕込み品の中の仕込み品（3階層）は原価に入らないので、作る前に知らせる
  var deep = isPrep && lines.some(function (L) { return L[1].indexOf('PRE-') === 0; });

  rebuildBlock_(repSh, [1, 3, 5, 10], [{ key: id, rows: lines }]);

  // ③ 調理手順
  var steps = (payload.steps || [])
    .filter(function (s) { return String(s.text || '').trim() !== ''; })
    .map(function (s, i) {
      return [id, i + 1, String(s.text).trim(), String(s.time || ''), String(s.heat || ''), String(s.point || '')];
    });
  rebuildBlock_(stepSh, [1, 3, 4, 5, 6, 7], [{ key: id, rows: steps }]);

  SpreadsheetApp.flush();

  // その場で原価を出して返す（シートを見に行かなくても分かるように）
  var unitCost = {};
  menuFormData().items.forEach(function (it) { unitCost[it.id] = it.cost; });
  var cost = 0;
  lines.forEach(function (L) { cost += (unitCost[L[1]] || 0) * L[2]; });

  var msg = id + '「' + name + '」を保存しました。材料' + lines.length + '件 / 手順' + steps.length + '件 / 原価 ' +
            (Math.round(cost * 10) / 10) + '円';
  if (isPrep) {
    var matSum = 0;
    lines.forEach(function (L) { matSum += Number(L[2]) || 0; });
    var made = Number(payload.yieldQty) > 0 ? Number(payload.yieldQty) : matSum;
    var u = String(payload.unit || 'g');
    msg += '（' + (Number(payload.yieldQty) > 0 ? '' : '出来上がり量が空なので材料の合計 ') + made + u +
           'ぶん ＝ 1' + u + 'あたり ' + (made > 0 ? Math.round(cost / made * 10000) / 10000 : 0) + '円）';
  }
  if (deep) msg += ' ※仕込み品の中の仕込み品は原価に入りません';

  return { id: id, kind: payload.kind, name: name, message: msg };
}

/** シートの2行目以降を配列で読む */
function readRows_(ss, name, cols) {
  var sh = ss.getSheetByName(name);
  if (!sh) return [];
  var last = lastDataRow_(sh);
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, cols).getValues();
}

/**
 * 「行を消さずに、手入力の列だけ組み直す」共通処理。
 *
 * ★行を削除すると2行目のARRAYFORMULAが壊れるため、削除は絶対にしない。
 *  レシピ明細・調理手順・試作からの反映が、すべてこの1つの関数を通ります。
 *
 * inputCols    … 手入力の列番号。1つ目の列の値で「どの商品の行か」を判定する
 * replacements … [{key:'MENU-0001', rows:[[1列目, 2列目, ...], ...]}]
 */
function rebuildBlock_(sh, inputCols, replacements) {
  // ★getLastRow() ではなく lastDataRow_ を使う。
  //   下から探すので、離れた行に飛んでしまったデータもここで拾って詰め直せる。
  var last = Math.max(lastDataRow_(sh, inputCols[0]), 2);
  var width = Math.max(sh.getLastColumn(), Math.max.apply(null, inputCols));

  var cur = [];
  if (last >= 2) {
    sh.getRange(2, 1, last - 1, width).getValues().forEach(function (r) {
      var row = inputCols.map(function (c) { return r[c - 1]; });
      if (String(row[0]).trim() !== '') cur.push(row);
    });
  }

  replacements.forEach(function (rep) {
    var out = [], done = false;
    cur.forEach(function (r) {
      if (String(r[0]).trim() === rep.key) {
        if (!done) { out = out.concat(rep.rows); done = true; }
      } else {
        out.push(r);
      }
    });
    if (!done) out = out.concat(rep.rows);       // その対象が1行も無かった場合は末尾に足す
    cur = out;
  });

  inputCols.forEach(function (c) {
    sh.getRange(2, c, last - 1, 1).clearContent();
  });

  if (cur.length === 0) return;

  var need = cur.length + 1 - sh.getMaxRows();
  if (need > 0) sh.insertRowsAfter(sh.getMaxRows(), need);

  inputCols.forEach(function (c, i) {
    sh.getRange(2, c, cur.length, 1).setValues(cur.map(function (r) { return [r[i]]; }));
  });
}


/**
 * 登録フォームの画面。
 * HTMLファイルを別に作らず、ここで文字列として持っています
 * （GAS側のファイルを コード.js の1つだけに保つため）。
 */
function menuFormHtml_() {
  return [
'<!DOCTYPE html><html><head><meta charset="utf-8"><base target="_top">',
'<style>',
'  body { font-family: "Meiryo", "Hiragino Kaku Gothic ProN", sans-serif; font-size: 13px;',
'         color: #2A2019; background: #FBF8F3; margin: 0; padding: 16px 18px 90px; }',
'  h2 { font-size: 15px; margin: 0 0 4px; }',
'  .lead { color: #7A7168; font-size: 12px; margin: 0 0 14px; }',
'  .card { background: #fff; border: 1px solid #E4DED3; border-radius: 8px; padding: 14px 16px; margin-bottom: 14px; }',
'  .card h3 { font-size: 13px; margin: 0 0 10px; padding-left: 8px; border-left: 4px solid #C8A96A; }',
'  label { display: block; font-size: 11px; color: #7A7168; margin-bottom: 3px; }',
'  input, select, textarea { font-family: inherit; font-size: 13px; padding: 6px 8px;',
'         border: 1px solid #D8D0C4; border-radius: 4px; background: #fff; box-sizing: border-box; }',
'  input:focus, select:focus { outline: 2px solid #C8A96A; outline-offset: -1px; }',
'  .row { display: flex; gap: 12px; flex-wrap: wrap; margin-bottom: 10px; align-items: flex-end; }',
'  .row > div { flex: 1 1 140px; display: flex; flex-direction: column; justify-content: flex-end; }',
'  .row input, .row select { width: 100%; }',
'  table { width: 100%; border-collapse: collapse; }',
'  th { font-size: 11px; color: #7A7168; text-align: left; font-weight: normal; padding: 0 6px 4px; }',
'  td { padding: 2px 6px 2px 0; vertical-align: middle; }',
'  td input { width: 100%; }',
'  .qty { text-align: right; }',
'  .unit { font-size: 11px; color: #7A7168; white-space: nowrap; }',
'  .cost { font-size: 12px; text-align: right; white-space: nowrap; color: #4A423A; }',
'  .del { border: none; background: none; color: #B23B3B; cursor: pointer; font-size: 15px; padding: 2px 6px; }',
'  .add { border: 1px dashed #C8A96A; background: #FEFCF8; color: #8A6A2A; cursor: pointer;',
'         border-radius: 4px; padding: 6px 12px; font-size: 12px; margin-top: 6px; font-family: inherit; }',
'  .sum { display: flex; gap: 18px; align-items: baseline; background: #F3EADB;',
'         border-radius: 6px; padding: 10px 14px; margin-top: 12px; flex-wrap: wrap; }',
'  .sum b { font-size: 18px; }',
'  .ok { color: #4B7B4E; } .ng { color: #B23B3B; }',
'  .bar { position: fixed; left: 0; right: 0; bottom: 0; background: #fff; border-top: 1px solid #E4DED3;',
'         padding: 12px 18px; display: flex; gap: 12px; align-items: center; }',
'  .save { background: #8A6A2A; color: #fff; border: none; border-radius: 5px; padding: 10px 26px;',
'          font-size: 14px; cursor: pointer; font-family: inherit; }',
'  .save:disabled { background: #C4BCB0; cursor: default; }',
'  .msg { font-size: 12px; }',
'  .hint { background: #F7F2E8; border-left: 3px solid #C8A96A; padding: 8px 12px; font-size: 11px;',
'          color: #6B6058; border-radius: 0 4px 4px 0; margin-bottom: 10px; line-height: 1.6; }',
'  .kind { display: flex; gap: 0; margin-bottom: 12px; }',
'  .kind button { flex: 1; padding: 10px; font-size: 13px; font-family: inherit; cursor: pointer;',
'          border: 1px solid #D8D0C4; background: #F7F4EE; color: #7A7168; }',
'  .kind button:first-child { border-radius: 5px 0 0 5px; }',
'  .kind button:last-child { border-radius: 0 5px 5px 0; border-left: none; }',
'  .kind button.on { background: #8A6A2A; color: #fff; border-color: #8A6A2A; }',
'</style></head><body>',

'<h2>商品・仕込み品を登録する</h2>',
'<p class="lead">ここで入れると、マスタ・レシピ明細・調理手順に自動で書き込まれます。IDは自動で振られます。</p>',

'<div class="card">',
'  <h3>1. 何を作りますか</h3>',
'  <div class="kind">',
'    <button id="kindMenu" class="on" onclick="setKind(\'menu\')">売る商品（メニュー）</button>',
'    <button id="kindPrep" onclick="setKind(\'prep\')">仕込み品（自分で作るもの）</button>',
'  </div>',
'  <div class="hint" id="prepHint" style="display:none">',
'    <b>出来上がり量は、ふつう空のままで構いません。</b>空だと「材料の合計」がそのまま使われます。<br>',
'    （例：醤油400＋みりん300＋酒200＋生姜100 → 1000g）<br><br>',
'    <b>入れるのは、作ると量が変わるものだけです。</b><br>',
'    ・炊く … 材料「米 1000g」→ 出来上がり量「2200」<br>',
'    ・煮詰める … 材料 1000g → 出来上がり量「700」<br><br>',
'    こうすると商品側では「炊いたご飯 240g」と盛る量で書けて、原価は自動で生米に戻ります。<br>',
'    倍率を直したくなったら、この出来上がり量を1か所直すだけで、使っている全商品に効きます。',
'  </div>',
'  <div class="row">',
'    <div style="flex:2 1 320px">',
'      <label>新しく作る / いまあるものを直す</label>',
'      <select id="pick" onchange="pickTarget()"></select>',
'    </div>',
'    <div style="flex:2 1 260px"><label id="nameLabel">商品名</label><input id="name" placeholder="例: 生姜焼き定食"></div>',
'  </div>',
'  <div class="row" id="menuFields">',
'    <div><label>カテゴリ</label><input id="category" placeholder="例: 定食"></div>',
'    <div><label>売価（税抜・円）</label><input id="price" class="qty" type="number" oninput="calc()"></div>',
'    <div><label>目標原価率（％）</label><input id="rate" class="qty" type="number" value="35" oninput="calc()"></div>',
'  </div>',
'  <div class="row" id="prepFields" style="display:none">',
'    <div><label>出来上がり量</label><input id="yieldQty" class="qty" type="number" step="any" oninput="calc()" placeholder="空＝材料の合計"></div>',
'    <div><label>単位</label><select id="unit"></select></div>',
'    <div><label>保存区分</label><select id="store"></select></div>',
'    <div><label>日持ち日数</label><input id="days" class="qty" type="number" placeholder="1"></div>',
'  </div>',
'</div>',

'<div class="card">',
'  <h3>2. 使う材料</h3>',
'  <table><thead><tr>',
'    <th style="width:44%">材料（打ちながら探せます）</th><th style="width:14%">分量</th>',
'    <th style="width:6%"></th><th style="width:14%">原価</th><th>ひとこと</th><th style="width:28px"></th>',
'  </tr></thead><tbody id="lines"></tbody></table>',
'  <button class="add" onclick="addLine()">＋ 材料を追加</button>',
'  <datalist id="itemlist"></datalist>',
'  <div class="sum">',
'    <span>材料の原価 <b id="cost">0</b> 円</span>',
'    <span id="extra"></span>',
'    <span id="diff"></span>',
'  </div>',
'</div>',

'<div class="card">',
'  <h3>3. 作り方（1行に1手順。空の行は無視されます）</h3>',
'  <table><thead><tr>',
'    <th style="width:30px">#</th><th style="width:46%">手順</th><th style="width:12%">時間</th>',
'    <th style="width:12%">火加減</th><th>ポイント</th><th style="width:28px"></th>',
'  </tr></thead><tbody id="steps"></tbody></table>',
'  <button class="add" onclick="addStep()">＋ 手順を追加</button>',
'</div>',

'<div class="bar">',
'  <button class="save" id="saveBtn" onclick="save()">保存する</button>',
'  <span class="msg" id="msg">読み込んでいます…</span>',
'</div>',

'<script>',
'var ITEMS = [], BYID = {}, TARGETS = [], KIND = "menu";',
'',
'google.script.run.withSuccessHandler(function (d) {',
'  ITEMS = d.items; TARGETS = d.targets; BYID = {};',
'  var dl = document.getElementById("itemlist");',
'  ITEMS.forEach(function (it) {',
'    BYID[it.id] = it;',
'    var o = document.createElement("option");',
'    o.value = it.id + " " + it.name;',
'    o.label = it.kind + " / " + it.unit;',
'    dl.appendChild(o);',
'  });',
'  fillSelect("unit", d.units, "g");',
'  fillSelect("store", d.stores, "常温");',
'  fillTargets();',
'  addLine(); addLine(); addLine(); addStep(); addStep();',
'  say("材料を打ち始めると候補が出ます。");',
'}).withFailureHandler(fail).menuFormData();',
'',
'function fillSelect(id, arr, def) {',
'  var s = document.getElementById(id);',
'  arr.forEach(function (v) {',
'    var o = document.createElement("option");',
'    o.value = v; o.textContent = v;',
'    if (v === def) o.selected = true;',
'    s.appendChild(o);',
'  });',
'}',
'',
'function fillTargets() {',
'  var sel = document.getElementById("pick");',
'  sel.innerHTML = "";',
'  var head = document.createElement("option");',
'  head.value = "";',
'  head.textContent = KIND === "prep" ? "＋ 新しい仕込み品を作る" : "＋ 新しい商品を作る";',
'  sel.appendChild(head);',
'  TARGETS.filter(function (t) { return t.kind === KIND; }).forEach(function (t) {',
'    var o = document.createElement("option");',
'    o.value = t.id; o.textContent = t.id + "  " + t.name;',
'    sel.appendChild(o);',
'  });',
'}',
'',
'function setKind(k) {',
'  KIND = k;',
'  document.getElementById("kindMenu").className = (k === "menu" ? "on" : "");',
'  document.getElementById("kindPrep").className = (k === "prep" ? "on" : "");',
'  document.getElementById("menuFields").style.display = (k === "menu" ? "flex" : "none");',
'  document.getElementById("prepFields").style.display = (k === "prep" ? "flex" : "none");',
'  document.getElementById("prepHint").style.display  = (k === "prep" ? "block" : "none");',
'  document.getElementById("nameLabel").textContent = (k === "prep" ? "仕込み品名" : "商品名");',
'  document.getElementById("name").placeholder = (k === "prep" ? "例: 炊いたご飯" : "例: 生姜焼き定食");',
'  fillTargets();',
'  clearForm();',
'  say(k === "prep" ? "作ると重さが変わるものは、ここで登録します。" : "新しい商品を作ります。");',
'}',
'',
'function clearForm() {',
'  document.getElementById("name").value = "";',
'  document.getElementById("category").value = "";',
'  document.getElementById("price").value = "";',
'  document.getElementById("rate").value = 35;',
'  document.getElementById("yieldQty").value = "";',
'  document.getElementById("days").value = "";',
'  document.getElementById("lines").innerHTML = "";',
'  document.getElementById("steps").innerHTML = "";',
'  addLine(); addLine(); addLine(); addStep(); addStep();',
'  calc();',
'}',
'',
'function say(t, bad) {',
'  var m = document.getElementById("msg");',
'  m.textContent = t; m.className = "msg " + (bad ? "ng" : "");',
'}',
'function fail(err) { say("エラー: " + err.message, true); document.getElementById("saveBtn").disabled = false; }',
'',
'function addLine(v) {',
'  v = v || {};',
'  var tr = document.createElement("tr");',
'  tr.innerHTML =',
'    "<td><input class=\\"item\\" list=\\"itemlist\\" oninput=\\"calc()\\" placeholder=\\"例: 米\\"></td>" +',
'    "<td><input class=\\"qty amount\\" type=\\"number\\" step=\\"any\\" oninput=\\"calc()\\"></td>" +',
'    "<td class=\\"unit\\"></td><td class=\\"cost\\">-</td>" +',
'    "<td><input class=\\"memo\\"></td>" +',
'    "<td><button class=\\"del\\" onclick=\\"this.closest(\'tr\').remove(); calc();\\">×</button></td>";',
'  document.getElementById("lines").appendChild(tr);',
'  if (v.id) {',
'    var it = BYID[v.id];',
'    tr.querySelector(".item").value = v.id + (it ? " " + it.name : "");',
'    tr.querySelector(".amount").value = v.qty;',
'    tr.querySelector(".memo").value = v.memo || "";',
'  }',
'  return tr;',
'}',
'',
'function addStep(v) {',
'  v = v || {};',
'  var tr = document.createElement("tr");',
'  tr.innerHTML =',
'    "<td class=\\"no\\"></td>" +',
'    "<td><input class=\\"text\\"></td><td><input class=\\"time\\" placeholder=\\"3分\\"></td>" +',
'    "<td><input class=\\"heat\\" placeholder=\\"中火\\"></td><td><input class=\\"point\\"></td>" +',
'    "<td><button class=\\"del\\" onclick=\\"this.closest(\'tr\').remove(); renumber();\\">×</button></td>";',
'  document.getElementById("steps").appendChild(tr);',
'  if (v.text) {',
'    tr.querySelector(".text").value = v.text;',
'    tr.querySelector(".time").value = v.time || "";',
'    tr.querySelector(".heat").value = v.heat || "";',
'    tr.querySelector(".point").value = v.point || "";',
'  }',
'  renumber();',
'  return tr;',
'}',
'',
'function renumber() {',
'  var rows = document.querySelectorAll("#steps tr");',
'  for (var i = 0; i < rows.length; i++) rows[i].querySelector(".no").textContent = (i + 1);',
'}',
'',
'/** 入力欄の文字から材料を1つに決める。"ING-0001 豚ロース" でも "豚ロース" でも通す */',
'function resolve(text) {',
'  var t = (text || "").trim();',
'  if (!t) return null;',
'  var m = t.match(/^((?:ING|PRE)-\\d+)/);',
'  if (m && BYID[m[1]]) return BYID[m[1]];',
'  var hit = ITEMS.filter(function (it) { return it.name === t; });',
'  if (hit.length === 1) return hit[0];',
'  return null;',
'}',
'',
'function calc() {',
'  var total = 0;',
'  var rows = document.querySelectorAll("#lines tr");',
'  for (var i = 0; i < rows.length; i++) {',
'    var tr = rows[i];',
'    var it = resolve(tr.querySelector(".item").value);',
'    var q = parseFloat(tr.querySelector(".amount").value) || 0;',
'    if (it) {',
'      var c = it.cost * q;',
'      total += c;',
'      tr.querySelector(".unit").textContent = it.unit;',
'      tr.querySelector(".cost").textContent = "\\u00A5" + (Math.round(c * 10) / 10);',
'      tr.querySelector(".item").style.background = "#fff";',
'    } else {',
'      tr.querySelector(".unit").textContent = "";',
'      tr.querySelector(".cost").textContent = "-";',
'      tr.querySelector(".item").style.background =',
'        tr.querySelector(".item").value.trim() ? "#FBEDEA" : "#fff";',
'    }',
'  }',
'  document.getElementById("cost").textContent = Math.round(total * 10) / 10;',
'',
'  var extra = document.getElementById("extra");',
'  var d = document.getElementById("diff");',
'',
'  if (KIND === "prep") {',
'    var y = parseFloat(document.getElementById("yieldQty").value) || 0;',
'    var u = document.getElementById("unit").value;',
'    var inQty = 0;',
'    for (var j = 0; j < rows.length; j++) inQty += parseFloat(rows[j].querySelector(".amount").value) || 0;',
'    var use = y > 0 ? y : inQty;',
'    if (use > 0) {',
'      extra.innerHTML = "1" + u + "あたり <b>" + (Math.round(total / use * 10000) / 10000) + "</b> 円";',
'      d.textContent = y > 0',
'        ? "材料 " + inQty + u + " → できあがり " + y + u + "（" + (Math.round(y / inQty * 100) / 100) + "倍）"',
'        : "出来上がり量が空なので、材料の合計 " + inQty + u + " をそのまま使います";',
'      d.className = "";',
'    } else { extra.textContent = "材料を入れると1" + u + "あたりの原価が出ます"; d.textContent = ""; }',
'    return;',
'  }',
'',
'  var price = parseFloat(document.getElementById("price").value) || 0;',
'  var target = parseFloat(document.getElementById("rate").value) || 0;',
'  if (price > 0) {',
'    var r = total / price * 100;',
'    extra.innerHTML = "原価率 <b>" + (Math.round(r * 10) / 10) + "%</b>";',
'    if (target > 0) {',
'      var gap = r - target;',
'      d.textContent = "目標との差 " + (gap > 0 ? "+" : "") + (Math.round(gap * 10) / 10) + "%";',
'      d.className = gap > 0 ? "ng" : "ok";',
'    } else { d.textContent = ""; }',
'  } else { extra.textContent = ""; d.textContent = ""; }',
'}',
'',
'function pickTarget() {',
'  var id = document.getElementById("pick").value;',
'  if (!id) { clearForm(); say("新しく作ります。"); return; }',
'  say("読み込んでいます…");',
'  google.script.run.withSuccessHandler(function (d) {',
'    document.getElementById("lines").innerHTML = "";',
'    document.getElementById("steps").innerHTML = "";',
'    document.getElementById("name").value = d.head.name;',
'    if (d.head.kind === "prep") {',
'      document.getElementById("yieldQty").value = d.head.yieldQty;',
'      document.getElementById("unit").value = d.head.unit || "g";',
'      document.getElementById("store").value = d.head.store || "常温";',
'      document.getElementById("days").value = d.head.days;',
'    } else {',
'      document.getElementById("category").value = d.head.category;',
'      document.getElementById("price").value = d.head.price;',
'      document.getElementById("rate").value = Math.round(d.head.rate * 1000) / 10;',
'    }',
'    d.lines.forEach(addLine);',
'    d.steps.forEach(addStep);',
'    if (d.lines.length === 0) addLine();',
'    if (d.steps.length === 0) addStep();',
'    calc();',
'    say(d.head.id + " を読み込みました。直したら保存してください。");',
'  }).withFailureHandler(fail).menuFormLoad(id);',
'}',
'',
'function save() {',
'  var lines = [], bad = [];',
'  var rows = document.querySelectorAll("#lines tr");',
'  for (var i = 0; i < rows.length; i++) {',
'    var text = rows[i].querySelector(".item").value.trim();',
'    if (!text) continue;',
'    var it = resolve(text);',
'    if (!it) { bad.push(text); continue; }',
'    lines.push({ id: it.id, qty: parseFloat(rows[i].querySelector(".amount").value) || 0,',
'                 memo: rows[i].querySelector(".memo").value });',
'  }',
'  if (bad.length > 0) { say("この材料が見つかりません: " + bad.join(" / ") + "（先に食材マスタに登録してください）", true); return; }',
'  if (lines.length === 0) { say("材料を1つ以上入れてください。", true); return; }',
'  if (!document.getElementById("name").value.trim()) { say("名前を入れてください。", true); return; }',
'',
'  var steps = [];',
'  var srows = document.querySelectorAll("#steps tr");',
'  for (var j = 0; j < srows.length; j++) {',
'    var t = srows[j].querySelector(".text").value.trim();',
'    if (!t) continue;',
'    steps.push({ text: t, time: srows[j].querySelector(".time").value,',
'                 heat: srows[j].querySelector(".heat").value, point: srows[j].querySelector(".point").value });',
'  }',
'',
'  var payload = {',
'    kind: KIND,',
'    id: document.getElementById("pick").value,',
'    name: document.getElementById("name").value.trim(),',
'    category: document.getElementById("category").value.trim(),',
'    price: parseFloat(document.getElementById("price").value) || 0,',
'    rate: (parseFloat(document.getElementById("rate").value) || 0) / 100,',
'    yieldQty: parseFloat(document.getElementById("yieldQty").value) || 0,',
'    unit: document.getElementById("unit").value,',
'    store: document.getElementById("store").value,',
'    days: parseFloat(document.getElementById("days").value) || 0,',
'    lines: lines, steps: steps',
'  };',
'',
'  document.getElementById("saveBtn").disabled = true;',
'  say("保存しています…");',
'  google.script.run.withSuccessHandler(function (res) {',
'    document.getElementById("saveBtn").disabled = false;',
'    say(res.message);',
'    if (!payload.id) {',
'      TARGETS.push({ id: res.id, name: res.name, kind: res.kind });',
'      fillTargets();',
'      document.getElementById("pick").value = res.id;',
'      // 新しく作った仕込み品を、すぐ材料としても使えるようにする',
'      google.script.run.withSuccessHandler(function (d) {',
'        ITEMS = d.items; BYID = {};',
'        var dl = document.getElementById("itemlist"); dl.innerHTML = "";',
'        ITEMS.forEach(function (it) {',
'          BYID[it.id] = it;',
'          var o = document.createElement("option");',
'          o.value = it.id + " " + it.name; o.label = it.kind + " / " + it.unit;',
'          dl.appendChild(o);',
'        });',
'      }).menuFormData();',
'    }',
'  }).withFailureHandler(fail).menuFormSave(payload);',
'}',
'<\/script></body></html>'
  ].join('\n');
}


/* =====================================================================
 *  【1回だけ】サンプルの米を「炊いたご飯」の仕込み品に置き換える
 *
 *  ペッパーランチ発注書は商品マスタに「生米g」と「盛りg」の2列を持って
 *  炊き上がりを表しているが、あちらは 商品→食材 の1階層しか無いための形。
 *  この台帳は 商品→仕込み品→食材 の2階層があるので、
 *  「炊いたご飯」という仕込み品を1つ作れば、倍率は1か所で決まる。
 *
 *  実行すると：
 *   ① 仕込み品マスタに「炊いたご飯」ができる（生米1000g → 2200g）
 *   ② レシピ明細に その材料行（米 1000g）が入る
 *   ③ 米を直接使っていた商品が「炊いたご飯」に置き換わる（量は倍率ぶん増える）
 *  原価は置き換えの前後で変わりません。
 *
 *  ★実際に炊いて量ってみて倍率が違っていたら、
 *   仕込み品マスタの「出来上がり量」だけを直してください。
 *   その1か所で、ご飯を使う全部の商品に反映されます。
 * ===================================================================== */

var RICE_BATCH_RAW   = 1000;    // 一度に炊く生米の量(g)
var RICE_COOKED_QTY  = 2200;    // それを炊いた後の量(g)。実測したらここでなく仕込み品マスタを直す
var RICE_PREP_NAME   = '炊いたご飯';

function ご飯を仕込み品にする() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var ingSh  = ss.getSheetByName('食材マスタ');
  var prepSh = ss.getSheetByName('仕込み品マスタ');
  var repSh  = ss.getSheetByName('レシピ明細');
  if (!ingSh || !prepSh || !repSh) { ui.alert('シートが見つかりません。'); return; }

  // すでに作ってあれば何もしない
  var prepLast = lastDataRow_(prepSh);
  var prepRows = prepLast >= 2 ? prepSh.getRange(2, 1, prepLast - 1, 2).getValues() : [];
  for (var p = 0; p < prepRows.length; p++) {
    if (String(prepRows[p][1]).trim() === RICE_PREP_NAME) {
      ui.alert('すでに「' + RICE_PREP_NAME + '」があります。\n\n' +
               '炊き上がりの量を変えたいときは、仕込み品マスタの「出来上がり量」を直してください。');
      return;
    }
  }

  // 米を探す（名前に「米」が付く食材）
  var ingLast = lastDataRow_(ingSh);
  var ingRows = ingLast >= 2 ? ingSh.getRange(2, 1, ingLast - 1, 2).getValues() : [];
  var riceId = '', riceName = '';
  for (var i = 0; i < ingRows.length; i++) {
    var nm = String(ingRows[i][1]);
    if (nm.indexOf('米') >= 0 && nm.indexOf('米粉') < 0) {
      riceId = String(ingRows[i][0]).trim(); riceName = nm; break;
    }
  }
  if (!riceId) { ui.alert('食材マスタに米が見つかりませんでした。\n先に米を登録してください。'); return; }

  // 米を直接使っている商品の行を数える
  var last = Math.max(lastDataRow_(repSh), 2);
  var rep = repSh.getRange(2, 1, last - 1, 10).getValues();
  var hits = [];
  for (var r = 0; r < rep.length; r++) {
    if (String(rep[r][2]).trim() === riceId && String(rep[r][0]).indexOf('MENU-') === 0) {
      hits.push({ row: r + 2, parent: String(rep[r][0]).trim(), qty: Number(rep[r][4]) || 0 });
    }
  }

  var ratio = RICE_COOKED_QTY / RICE_BATCH_RAW;
  var msg = '材料の米：' + riceId + ' ' + riceName + '\n' +
            '作るもの：' + RICE_PREP_NAME + '（生米' + RICE_BATCH_RAW + 'g → ' + RICE_COOKED_QTY + 'g／' + ratio + '倍）\n\n';
  if (hits.length > 0) {
    msg += '米を直接使っている ' + hits.length + '件を「' + RICE_PREP_NAME + '」に置き換えます:\n';
    hits.forEach(function (h) {
      msg += '  ' + h.parent + '　米 ' + h.qty + 'g → ご飯 ' + Math.round(h.qty * ratio * 10) / 10 + 'g\n';
    });
    msg += '\n原価は置き換えの前後で変わりません。\n\n';
  } else {
    msg += '米を直接使っている商品はありませんでした。仕込み品だけ作ります。\n\n';
  }
  msg += 'よろしいですか?';
  if (ui.alert('ご飯を仕込み品にする', msg, ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;

  // ① 仕込み品マスタに追加
  var prepId = 'PRE-' + zero4_(nextIdNumber_(prepSh, ID_RULES['仕込み品マスタ']));
  var prepRow = lastDataRow_(prepSh) + 1;
  prepSh.getRange(prepRow, 1, 1, 6).setValues([[
    prepId, RICE_PREP_NAME, RICE_COOKED_QTY, 'g', '常温', 1
  ]]);

  // ② レシピ明細に「炊いたご飯 ← 米」の材料行を足す。
  //    ③の置き換えも同時に行う（行を消さずに入力列だけ組み直す）
  var lines = [];
  for (var j = 0; j < rep.length; j++) {
    var parent = String(rep[j][0]).trim();
    if (!parent) continue;
    var child = String(rep[j][2]).trim();
    var qty   = rep[j][4];
    var memo  = String(rep[j][9]);
    if (child === riceId && parent.indexOf('MENU-') === 0) {
      child = prepId;
      qty   = Math.round((Number(qty) || 0) * ratio * 10) / 10;
      memo  = memo.replace(/炊く前の重量/g, '').trim();
      memo  = memo ? memo + '（炊き上がりの量）' : '炊き上がりの量';
    }
    lines.push([parent, child, qty, memo]);
  }
  lines.push([prepId, riceId, RICE_BATCH_RAW, '炊く前の生米']);

  var inputCols = [1, 3, 5, 10];
  inputCols.forEach(function (c) { repSh.getRange(2, c, last - 1, 1).clearContent(); });
  var need = lines.length + 1 - repSh.getMaxRows();
  if (need > 0) repSh.insertRowsAfter(repSh.getMaxRows(), need);
  inputCols.forEach(function (c, k) {
    repSh.getRange(2, c, lines.length, 1).setValues(lines.map(function (x) { return [x[k]]; }));
  });

  SpreadsheetApp.flush();
  ss.setActiveSheet(prepSh);

  ui.alert('できました。\n\n' +
    prepId + '「' + RICE_PREP_NAME + '」\n' +
    '　生米 ' + RICE_BATCH_RAW + 'g → ' + RICE_COOKED_QTY + 'g（' + ratio + '倍）\n\n' +
    (hits.length > 0 ? '商品 ' + hits.length + '件を置き換えました。\n\n' : '') +
    '★実際に炊いて量ってみて倍率が違っていたら、\n' +
    '　仕込み品マスタの「出来上がり量」だけを直してください。\n' +
    '　その1か所で、ご飯を使う全部の商品に反映されます。');
}


/* =====================================================================
 *  試作を取り消す（間違えて作った回を消す）
 *
 *  ★消す前に「何が消えるか」を必ず見せる。黙って消さない。
 *   対象そのもの（マヨネーズなど）は、どこからも使われていないときだけ消せる。
 * ===================================================================== */

function 試作を取り消す() {
  var html = HtmlService.createHtmlOutput(trialDeleteHtml_())
    .setWidth(720).setHeight(560);
  SpreadsheetApp.getUi().showModalDialog(html, '🗑 試作を取り消す');
}

/** 消せる試作の一覧。それぞれ「何が入っているか」も一緒に返す */
function trialDeleteData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var det = readRows_(ss, '試作レシピ明細', 8);
  var lineCount = {};
  det.forEach(function (r) {
    var t = String(r[0]).trim();
    if (t) lineCount[t] = (lineCount[t] || 0) + 1;
  });

  var rep = readRows_(ss, 'レシピ明細', 10);
  var usedAsParent = {}, usedAsChild = {};
  rep.forEach(function (r) {
    var p = String(r[0]).trim(), c = String(r[2]).trim();
    if (p) usedAsParent[p] = (usedAsParent[p] || 0) + 1;
    if (c) usedAsChild[c] = (usedAsChild[c] || 0) + 1;
  });

  var rec = readRows_(ss, '試作記録', 15);
  var countByTarget = {};
  rec.forEach(function (r) {
    var t = String(r[1]).trim();
    if (t) countByTarget[t] = (countByTarget[t] || 0) + 1;
  });

  var names = {};
  readRows_(ss, '仕込み品マスタ', 8).forEach(function (r) { if (r[0]) names[String(r[0])] = String(r[1]); });
  readRows_(ss, '商品マスタ', 10).forEach(function (r) { if (r[0]) names[String(r[0])] = String(r[1]); });

  var list = [];
  rec.forEach(function (r) {
    var tryId = String(r[0]).trim();
    if (!tryId) return;
    var target = String(r[1]).trim();
    var taste = [5, 6, 7, 8, 9, 10].filter(function (i) { return String(r[i]).trim() !== ''; }).length;

    // 対象そのものを消せるか（この回が最後の試作で、レシピでも使われていないこと）
    var why = [];
    if ((countByTarget[target] || 0) > 1) why.push('ほかにも試作の回があります');
    if (usedAsParent[target]) why.push('レシピ明細で材料が登録されています');
    if (usedAsChild[target]) why.push('ほかの商品の材料として使われています');

    list.push({
      tryId: tryId, target: target, targetName: names[target] || target,
      no: Number(r[3]) || 0,
      date: r[4] ? Utilities.formatDate(new Date(r[4]), 'JST', 'yyyy/MM/dd') : '',
      lines: lineCount[tryId] || 0,
      taste: taste,
      memo: String(r[13] || ''),
      canDeleteTarget: why.length === 0,
      why: why.join('・')
    });
  });

  list.reverse();          // 新しいものを上に
  return { list: list };
}

/** 試作の回を消す。p = { tryId, alsoTarget } */
function trialDelete(p) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var recSh = ss.getSheetByName('試作記録');
  var detSh = ss.getSheetByName('試作レシピ明細');
  if (!recSh || !detSh) throw new Error('試作シートが見つかりません。');

  var tryId = String(p.tryId || '').trim();
  if (!tryId) throw new Error('どの回を消すか選んでください。');

  // 消してよいか、もう一度サーバー側でも確かめる（画面の情報が古い場合に備える）
  var info = null;
  trialDeleteData().list.forEach(function (x) { if (x.tryId === tryId) info = x; });
  if (!info) throw new Error(tryId + ' が見つかりません。すでに消えている可能性があります。');
  if (p.alsoTarget && !info.canDeleteTarget) {
    throw new Error('「' + info.targetName + '」は消せません（' + info.why + '）。');
  }

  var done = [];

  // 試作レシピ明細から、その回の材料を消す（rows を空にする＝消える）
  rebuildBlock_(detSh, [1, 2, 4, 8], [{ key: tryId, rows: [] }]);
  if (info.lines > 0) done.push('試作レシピ明細から材料 ' + info.lines + ' 件を消しました');

  // 試作記録から、その回を消す
  var recCols = inputColsOf_('試作記録');
  rebuildBlock_(recSh, recCols, [{ key: tryId, rows: [] }]);
  done.push('試作記録から ' + tryId + '（' + info.targetName + ' 第' + info.no + '回）を消しました');

  // 対象そのものも消す場合
  if (p.alsoTarget) {
    var isPrep = (info.target.indexOf('PRE-') === 0);
    var sh = ss.getSheetByName(isPrep ? '仕込み品マスタ' : '商品マスタ');
    rebuildBlock_(sh, inputColsOf_(isPrep ? '仕込み品マスタ' : '商品マスタ'),
                  [{ key: info.target, rows: [] }]);
    done.push((isPrep ? '仕込み品マスタ' : '商品マスタ') + 'から ' + info.target +
              '「' + info.targetName + '」を消しました');

    // 調理手順に残っていれば一緒に消す
    var stepSh = ss.getSheetByName('調理手順');
    if (stepSh) rebuildBlock_(stepSh, inputColsOf_('調理手順'), [{ key: info.target, rows: [] }]);
  }

  // 試作比較が消した対象を指していたら外す
  var cmp = ss.getSheetByName('試作比較');
  if (cmp && String(cmp.getRange('B1').getValue()).trim() === info.target && p.alsoTarget) {
    cmp.getRange('B1').clearContent();
  }

  SpreadsheetApp.flush();
  return { done: done };
}

/** そのシートの「手で入れる列」の番号（自動計算の列を除いたもの） */
function inputColsOf_(name) {
  var roles = COLUMN_ROLES[name] || [];
  var cols = [];
  for (var i = 0; i < roles.length; i++) if (roles[i] !== ROLE_AUTO) cols.push(i + 1);
  return cols;
}


/** 取り消しの画面 */
function trialDeleteHtml_() {
  return [
'<!DOCTYPE html><html><head><meta charset="utf-8"><base target="_top">',
'<style>',
'  body { font-family: "Meiryo", "Hiragino Kaku Gothic ProN", sans-serif; font-size: 13px;',
'         color: #2A2019; background: #FBF8F3; margin: 0; padding: 16px 18px 78px; }',
'  h2 { font-size: 15px; margin: 0 0 4px; }',
'  .lead { color: #7A7168; font-size: 12px; margin: 0 0 14px; }',
'  .card { background: #fff; border: 1px solid #E4DED3; border-radius: 8px; padding: 14px 16px; margin-bottom: 12px; }',
'  label { display: block; font-size: 11px; color: #7A7168; margin-bottom: 3px; }',
'  select { font-family: inherit; font-size: 13px; padding: 7px 8px; width: 100%;',
'         border: 1px solid #D8D0C4; border-radius: 4px; background: #fff; box-sizing: border-box; }',
'  .box { background: #FDF6F4; border: 1px solid #E8C9C2; border-radius: 6px; padding: 12px 14px; margin-top: 12px; }',
'  .box h4 { margin: 0 0 8px; font-size: 12px; color: #A0453A; }',
'  .box ul { margin: 0; padding-left: 20px; line-height: 1.8; font-size: 12px; }',
'  .opt { margin-top: 12px; padding: 10px 12px; background: #F7F4EE; border-radius: 6px; font-size: 12px; }',
'  .opt input { margin-right: 6px; }',
'  .opt .why { color: #9A9189; font-size: 11px; margin: 4px 0 0 22px; }',
'  .done { background: #EEF4EC; border: 1px solid #C6DCC0; border-radius: 8px; padding: 14px 16px; }',
'  .done ul { margin: 6px 0 0; padding-left: 20px; line-height: 1.9; font-size: 12px; }',
'  .bar { position: fixed; left: 0; right: 0; bottom: 0; background: #fff; border-top: 1px solid #E4DED3;',
'         padding: 12px 18px; display: flex; gap: 12px; align-items: center; }',
'  .go { background: #B23B3B; color: #fff; border: none; border-radius: 5px; padding: 10px 26px;',
'          font-size: 14px; cursor: pointer; font-family: inherit; }',
'  .go:disabled { background: #C4BCB0; cursor: default; }',
'  .msg { font-size: 12px; } .ng { color: #B23B3B; } .ok { color: #4B7B4E; }',
'</style></head><body>',

'<div id="page1">',
'  <h2>試作を取り消す</h2>',
'  <p class="lead">間違えて作ってしまった回を消せます。消す前に、何が消えるかを必ず確認してください。</p>',
'  <div class="card">',
'    <label>どの回を消しますか</label>',
'    <select id="pick" onchange="showInfo()"></select>',
'    <div id="info"></div>',
'    <div class="opt" id="optBox" style="display:none">',
'      <label style="display:inline"><input type="checkbox" id="alsoTarget" onchange="showInfo()">',
'        <span id="optLabel"></span></label>',
'      <div class="why" id="optWhy"></div>',
'    </div>',
'  </div>',
'</div>',

'<div id="page2" style="display:none">',
'  <h2>消しました</h2>',
'  <div class="done" id="doneBox"></div>',
'</div>',

'<div class="bar">',
'  <button class="go" id="goBtn" onclick="go()">この回を消す</button>',
'  <span class="msg" id="msg">読み込んでいます…</span>',
'</div>',

'<script>',
'var LIST = [], STEP = 1;',
'',
'google.script.run.withSuccessHandler(function (d) {',
'  LIST = d.list;',
'  var sel = document.getElementById("pick");',
'  if (LIST.length === 0) {',
'    sel.innerHTML = "<option>（試作の記録がありません）</option>";',
'    document.getElementById("goBtn").disabled = true;',
'    say("消せる試作がありません。");',
'    return;',
'  }',
'  LIST.forEach(function (x) {',
'    var o = document.createElement("option");',
'    o.value = x.tryId;',
'    o.textContent = x.targetName + "　第" + x.no + "回（" + x.tryId + "）" + (x.date ? "　" + x.date : "");',
'    sel.appendChild(o);',
'  });',
'  showInfo();',
'  say("");',
'}).withFailureHandler(fail).trialDeleteData();',
'',
'function say(t, bad) {',
'  var m = document.getElementById("msg");',
'  m.textContent = t; m.className = "msg " + (bad ? "ng" : "ok");',
'}',
'function fail(e) { say("エラー: " + e.message, true); document.getElementById("goBtn").disabled = false; }',
'function esc(s) { return String(s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;"); }',
'',
'function cur() {',
'  var id = document.getElementById("pick").value;',
'  return LIST.filter(function (x) { return x.tryId === id; })[0];',
'}',
'',
'function showInfo() {',
'  var x = cur();',
'  if (!x) return;',
'  var also = document.getElementById("alsoTarget").checked && x.canDeleteTarget;',
'  var items = [];',
'  items.push("試作記録から <b>" + x.tryId + "</b>（" + esc(x.targetName) + " 第" + x.no + "回）");',
'  items.push("試作レシピ明細から 材料 <b>" + x.lines + "</b> 件");',
'  if (x.taste > 0) items.push("<b style=\\"color:#B23B3B\\">入力ずみの味 " + x.taste + " 項目も一緒に消えます</b>");',
'  if (x.memo) items.push("所感「" + esc(x.memo) + "」");',
'  if (also) items.push("<b>" + esc(x.targetName) + "（" + x.target + "）そのもの</b>も台帳から消えます");',
'',
'  document.getElementById("info").innerHTML =',
'    "<div class=\\"box\\"><h4>次のものが消えます</h4><ul><li>" + items.join("</li><li>") + "</li></ul></div>";',
'',
'  var box = document.getElementById("optBox");',
'  box.style.display = "block";',
'  document.getElementById("optLabel").innerHTML =',
'    "「" + esc(x.targetName) + "」自体も台帳から消す（" + x.target + "）";',
'  var cb = document.getElementById("alsoTarget");',
'  cb.disabled = !x.canDeleteTarget;',
'  if (!x.canDeleteTarget) cb.checked = false;',
'  document.getElementById("optWhy").textContent = x.canDeleteTarget',
'    ? "どこからも使われていないので、一緒に消せます。"',
'    : "消せません：" + x.why;',
'}',
'',
'function go() {',
'  if (STEP === 2) { google.script.host.close(); return; }',
'  var x = cur();',
'  if (!x) return;',
'  document.getElementById("goBtn").disabled = true;',
'  say("消しています…");',
'  google.script.run.withSuccessHandler(function (r) {',
'    document.getElementById("page1").style.display = "none";',
'    document.getElementById("page2").style.display = "block";',
'    document.getElementById("doneBox").innerHTML =',
'      "<b>消しました</b><ul><li>" + r.done.join("</li><li>") + "</li></ul>";',
'    document.getElementById("goBtn").textContent = "閉じる";',
'    document.getElementById("goBtn").disabled = false;',
'    STEP = 2;',
'    say("");',
'  }).withFailureHandler(fail).trialDelete({',
'    tryId: x.tryId,',
'    alsoTarget: document.getElementById("alsoTarget").checked',
'  });',
'}',
'<\/script></body></html>'
  ].join('\n');
}
