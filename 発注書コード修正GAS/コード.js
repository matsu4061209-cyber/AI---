function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🔧 コード修正')
    .addItem('★最新CSVから消耗品を同期', 'syncFromLatestCatalogCsv')
    .addSeparator()
    .addItem('商品コード一括修正を実行', 'fixConsumableCodes')
    .addItem('商品名をサイト名に統一', 'syncNamesToCatalog')
    .addItem('単価・発注単位をサイトからコピー', 'syncPriceAndUnitFromCatalog')
    .addItem('消耗品をサイト順に並び替え', 'reorderConsumables')
    .addSeparator()
    .addItem('🩺 列構成を診断（書き込みなし）', 'diagnoseColumns')
    .addItem('🛠 単位規格列を復元', 'repairSpecColumn')
    .addItem('🗑 F列(単位規格)を削除', 'deleteSpecColumn')
    .addItem('🧹 過去の同期バグを一括修正', 'repairSyncBugs')
    .addToUi();
}

// 「発注」フォルダ内の my_catalog_trade_prod*.csv のうち、最終更新が最も新しいものを返す。
function findLatestCatalogFile_() {
  const folders = DriveApp.getFoldersByName('発注');
  if (!folders.hasNext()) return null;
  const folder = folders.next();
  const it = folder.getFiles();
  let latest = null;
  let latestDate = null;
  while (it.hasNext()) {
    const f = it.next();
    const m = f.getName().match(/^my_catalog_trade_prod(\d{8})/i);
    if (!m) continue;
    const dateStr = m[1]; // ファイル名の日付(YYYYMMDD)を正とする
    if (
      !latest ||
      dateStr > latestDate ||
      (dateStr === latestDate && f.getLastUpdated() > latest.getLastUpdated())
    ) {
      latest = f;
      latestDate = dateStr;
    }
  }
  return latest;
}

// Shift-JISのCSV Blobをパースして、[ヘッダー行を除いた] 2次元配列を返す。
function parseCatalogCsv_(file) {
  const bytes = file.getBlob().getBytes();
  const text = Utilities.newBlob(bytes).getDataAsString('Shift_JIS');
  const lines = text.split(/\r\n|\n/);
  // 1行目はタイトル行「[FOODS INFOMART]」なので除外し、2行目をヘッダーとして扱う
  const body = lines.slice(1).join('\n');
  const rows = Utilities.parseCsv(body);
  return rows.slice(1); // ヘッダー行を除いたデータ行
}

// 最新CSVから 自社管理コード→{name, price, spec} のマップを作る。
function buildCatalogMapFromCsv_(rows) {
  const map = {};
  const order = [];
  rows.forEach(function (r) {
    const code = (r[2] || '').trim();
    if (!code) return;
    map[code] = {
      name: (r[12] || '').trim(),
      price: r[17] !== '' && r[17] !== undefined ? Number(r[17]) : null,
      spec: (r[13] || '').trim()
    };
    order.push(code);
  });
  return { map: map, order: order };
}

// 最新CSVを自動で見つけて、消耗品シートの商品名・単価・発注単位(規格の目安)・並び順をまとめて同期する。
function syncFromLatestCatalogCsv() {
  const ui = SpreadsheetApp.getUi();
  const file = findLatestCatalogFile_();
  if (!file) {
    ui.alert('「発注」フォルダ内に my_catalog_trade_prod から始まるCSVが見つかりませんでした。');
    return;
  }

  const found = findConsumableHeader_();
  if (!found) {
    ui.alert('消耗品セクションの見出し行が見つかりませんでした。');
    return;
  }
  const sheet = found.sheet;
  const dataStartRow = found.headerRow + 1;
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const numRows = lastRow - dataStartRow + 1;
  if (numRows <= 0 || numRows % 2 !== 0) {
    ui.alert('データ行数が想定外です（' + numRows + '行）。中止しました。');
    return;
  }

  const range = sheet.getRange(dataStartRow, 1, numRows, lastCol);
  const values = range.getValues();
  const formulas = range.getFormulas();
  for (let r = 0; r < formulas.length; r++) {
    for (let c = 0; c < formulas[r].length; c++) {
      if (formulas[r][c]) {
        ui.alert(
          '数式が見つかりました',
          '行' + (dataStartRow + r) + ' 列' + (c + 1) + ' に数式があります。安全のため同期を中止しました。',
          ui.ButtonSet.OK
        );
        return;
      }
    }
  }

  // 「基準量/在庫」マーカー列と、日付が始まる列を見出し行から特定する（列構成の変化に自動対応）。
  const headerVals = sheet.getRange(found.headerRow, 1, 1, lastCol).getDisplayValues()[0];
  let datesStartCol = -1;
  for (let c = 5; c < lastCol; c++) { // F列(6列目)以降を探す
    if (/^\d+月\d+日$/.test(String(headerVals[c]).trim())) { datesStartCol = c + 1; break; }
  }
  if (datesStartCol === -1) {
    ui.alert('日付列が見つかりませんでした。列構成を診断してから実行してください。');
    return;
  }
  const markerCol = datesStartCol - 2; // 0-indexed配列上のマーカー列位置（datesStartColは1始まりの列番号なので-2）

  const parsedRows = parseCatalogCsv_(file);
  const catalog = buildCatalogMapFromCsv_(parsedRows);

  // 2行1組（基準量/在庫）で扱う
  const pairs = [];
  for (let i = 0; i < numRows; i += 2) {
    pairs.push({ code: String(values[i][1]).trim(), rows: [values[i], values[i + 1]] });
  }

  let nameCount = 0, priceCount = 0, specCount = 0;
  const notFound = [];
  pairs.forEach(function (p) {
    const info = catalog.map[p.code];
    if (!info) {
      if (p.code) notFound.push(p.code);
      return;
    }
    if (info.name) { p.rows[0][2] = info.name; nameCount++; }
    if (info.price !== null && !isNaN(info.price)) { p.rows[0][3] = info.price; priceCount++; }
    if (info.spec) { p.rows[0][4] = info.spec; specCount++; }
  });

  // CSVの並び順を優先し、CSVに無いものは元の相対順のまま後ろに続ける。
  // シートにまだ無いコードは、新商品として新しい行ペアを作って追加する。
  const used = new Array(pairs.length).fill(false);
  const ordered = [];
  const added = [];
  catalog.order.forEach(function (code) {
    for (let i = 0; i < pairs.length; i++) {
      if (!used[i] && pairs[i].code === code) {
        ordered.push(pairs[i]);
        used[i] = true;
        return;
      }
    }
    // シートに無い新商品 → 新規行ペアを作成
    const info = catalog.map[code];
    const row0 = new Array(lastCol).fill('');
    const row1 = new Array(lastCol).fill('');
    row0[0] = '消耗品';
    row0[1] = code;
    row0[2] = info.name || '';
    row0[3] = info.price !== null && !isNaN(info.price) ? info.price : '';
    row0[4] = info.spec || '';
    row0[markerCol] = '基準量';
    row1[1] = code;
    row1[markerCol] = '在庫';
    ordered.push({ code: code, rows: [row0, row1], isNew: true });
    added.push(code + (info.name ? '(' + info.name + ')' : ''));
  });
  pairs.forEach(function (p, i) { if (!used[i]) ordered.push(p); });

  const newNumRows = ordered.length * 2;
  if (newNumRows > numRows) {
    sheet.insertRowsAfter(lastRow, newNumRows - numRows);
  }
  const writeRange = sheet.getRange(dataStartRow, 1, newNumRows, lastCol);

  const newValues = [];
  ordered.forEach(function (p) { newValues.push(p.rows[0]); newValues.push(p.rows[1]); });
  writeRange.setValues(newValues);

  // 罫線・A〜C列とL列(最終列)の2行結合が無い行があれば、既存行と同じ書式を適用する。
  // 新規追加行だけでなく、以前から書式が崩れていた行（過去バグの取りこぼし等）も毎回チェックして直す。
  let formatFixed = 0;
  const templateRange = sheet.getRange(dataStartRow, 1, 2, lastCol);
  for (let i = 0; i < ordered.length; i++) {
    const r = dataStartRow + i * 2;
    const merged = sheet.getRange(r, 3).getMergedRanges();
    const isMerged = merged.length > 0 && merged[0].getNumRows() === 2;
    if (!isMerged) {
      templateRange.copyTo(sheet.getRange(r, 1, 2, lastCol), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      sheet.getRange(r, 1, 2, 1).merge(); // A
      sheet.getRange(r, 2, 2, 1).merge(); // B
      sheet.getRange(r, 3, 2, 1).merge(); // C
      sheet.getRange(r, lastCol, 2, 1).merge(); // L(最終列)
      formatFixed++;
    }
  }

  let message = 'CSV: ' + file.getName() + '\n更新日時: ' + file.getLastUpdated() + '\n\n'
    + '商品名更新: ' + nameCount + '件\n単価更新: ' + priceCount + '件\n発注単位(目安)更新: ' + specCount + '件\n罫線・結合の修正: ' + formatFixed + '件';
  if (added.length > 0) {
    message += '\n\n🆕 新商品として追加（発注数量は未設定・要入力）:\n' + added.join('\n');
  }
  if (notFound.length > 0) {
    message += '\n\n⚠️ 最新CSVに見つからなかったコード（要確認・自動変更なし）:\n' + notFound.join(', ');
  }
  ui.alert('最新CSVから同期完了', message, ui.ButtonSet.OK);
}

// サイト側の単価。キーは現在のシート上のコード。
const CATALOG_PRICE = {
  '636082': 240, '636266': 781, '636312': 650, '636313': 650, '636314': 650,
  '636410': 4077, '636054': 460, '636055': 836, '636024': 420, '636276': 800,
  '636092': 598, '636095': 598, '636096': 598, '636098': 598, '636525': 635,
  '638015': 7546, '636015': 542, '636020': 365, '638016': 13529, '636050': 1438,
  '638017': 10286, '636088': 608, '638018': 5349, '638019': 2808, '638020': 6319,
  '638021': 2028, '638023': 12489, '636113': 1235, '636186': 768, '636118': 251,
  '636143': 323, '636270': 1217, '636271': 904, '636272': 1136, '636273': 780,
  '636275': 430, '636112': 320, '636196': 361, '636265': 80, '636529': 267,
  '636005': 183, '636269': 1701, '360459': 337, '636280': 1096
};

// サイト側の「規格」欄の文字列（目安として発注単位欄にそのまま転記する）。空欄はスキップ。
const CATALOG_SPEC = {
  '636082': '100枚', '636312': '１００枚／袋×３０袋／ケース', '636313': '１００枚／袋×３０袋／ケース',
  '636314': '１００枚／袋×３０袋／ケース', '636410': '30個入', '636054': '50枚入', '636055': '50枚入',
  '636276': '250枚', '638015': '10ＫＧ', '636015': '100本', '636020': '50個入', '638016': '4L',
  '636050': '100枚入', '638017': '４Ｌ', '638018': '4kg', '638019': '4.8Ｌ', '638020': '4kg',
  '638021': '5kg', '638023': '2L', '636113': 'ｻｰﾏﾙ 58X80X20 ｺｱﾚｽ 5Rｸﾛﾎﾟﾘ', '636186': 'ｻｰﾏﾙ 80X65mX12 4ｶﾝｸﾛﾎﾟﾘ',
  '636118': '100膳', '636270': '50枚入り', '636271': '50枚入り', '636272': '50枚入り1', '636273': '50枚入り',
  '636275': '１００枚', '636112': '50枚入', '636196': '1個', '636265': '150枚入', '636529': '200本入',
  '636269': '５００本', '360459': '２２オンス', '636280': '１４オンス'
};

// D列(単価)にCATALOG_PRICE、E列(発注単位)にCATALOG_SPECを転記する。
function syncPriceAndUnitFromCatalog() {
  const found = findConsumableHeader_();
  const ui = SpreadsheetApp.getUi();
  if (!found) {
    ui.alert('消耗品セクションの見出し行が見つかりませんでした。');
    return;
  }
  const sheet = found.sheet;
  const lastRow = sheet.getLastRow();
  const dataStartRow = found.headerRow + 1;
  const numRows = lastRow - dataStartRow + 1;
  if (numRows <= 0) {
    ui.alert('データ行が見つかりませんでした。');
    return;
  }
  const codes = sheet.getRange(dataStartRow, 2, numRows, 1).getValues(); // B列
  let priceCount = 0;
  let specCount = 0;
  for (let i = 0; i < codes.length; i++) {
    const code = String(codes[i][0]).trim();
    const row = dataStartRow + i;
    if (CATALOG_PRICE[code] !== undefined) {
      sheet.getRange(row, 4).setValue(CATALOG_PRICE[code]); // D列
      priceCount++;
    }
    if (CATALOG_SPEC[code]) {
      sheet.getRange(row, 5).setValue(CATALOG_SPEC[code]); // E列
      specCount++;
    }
  }
  ui.alert('単価・発注単位の転記完了', '単価: ' + priceCount + '件 / 発注単位: ' + specCount + '件 を転記しました。', ui.ButtonSet.OK);
}

// サイト(マイカタログ)側の正式名称。キーは現在のシート上のコード（コード変更後の新コード）。
const CATALOG_NAMES = {
  '636082': 'エンボス手袋ﾌﾞﾙｰMｻｲｽﾞ【100枚/袋】',
  '636266': 'キッチンペーパーイイナ２６５　９０枚巻（２本入）',
  '636312': '【外部】ニトリル手袋６００１粉無　Ｓ　濃紺（１００枚）',
  '636313': '【外部】ニトリル手袋６００１粉無　Ｍ　濃紺（１００枚）',
  '636314': '【外部】ニトリル手袋６００１粉無　Ｌ　濃紺（１００枚）',
  '636410': 'ソフネットハンドタオル２００　中判',
  '636054': 'ＨＤゴミ袋　４５Ｌ　スタンダード　透明',
  '636055': 'ＨＤゴミ袋　７０Ｌ　スタンダード　透明',
  '636024': 'ポリマラップ　３００×１００　プラＰＥＴ',
  '636276': 'PL油よけはちまき＿A【250枚/1束】',
  '636092': '曜日シール　火曜日　1000枚',
  '636095': '曜日シール　日曜日　1000枚',
  '636096': '曜日シール　金曜日　1000枚',
  '636098': '曜日シール　木曜日　1000枚',
  '636525': '米炊ネット（ＫＮ－０７５－Ｍ）',
  '638015': 'リキッド輝跡　10kg-1（洗浄機用液体洗剤）',
  '636015': 'プラスプーン16cmアイボリー【100本/袋】',
  '636020': 'タレ入　角大（Ｄ）',
  '638016': 'ﾆｭｰﾘﾝｽﾄﾞﾗｲ　4L-1（食器洗浄機用乾燥促進剤）',
  '636050': 'ＣＦカップ９５－２７０　セット',
  '638017': 'ｵｱｼｽｺﾝﾊﾟｯｸ ｷｯﾁﾝｸﾘｰﾅｰ ﾌﾟﾗｽ 4L-1',
  '636088': '解凍シール 12枚/シート50',
  '638018': '無香料 ﾊﾝﾄﾞｿｰﾌﾟ 濃縮 4KG-1',
  '638019': 'エコラールＥ　4.8Ｌ-1（除菌用アルコール）',
  '638020': 'ｸﾞﾘｰｽｶｯﾀｰ EX 4KG-1',
  '638021': 'エコサンプラス　5ｋｇ-1 （塩素系漂白・殺菌剤）',
  '638023': 'HB ｸﾜｯﾄ 2L-1',
  '636113': 'レジロール紙　５巻入り',
  '636186': 'デリバリー用ロール紙　４巻入り',
  '636118': '中国元禄8寸（上質）100膳/袋',
  '636143': '弁当シール40枚（2時間以内）',
  '636270': 'ＤＬＶ角１９－１４Ｒ（５５）ＭＳＤ',
  '636271': 'Ｔ－ＤＬＶ角１９－１４内外嵌合Ｋ蓋穴有',
  '636272': 'Ｔ－ＤＬＶ角１９－１４Ｒ側面嵌合Ｋ蓋',
  '636273': 'ＤＬＶ角１９－１４Ｒ中皿－１　ＰＰＮＴ',
  '636275': 'PLﾊﾞｲｵﾏｽﾚｼﾞ袋 100枚',
  '636112': '無地エプロン　３９．５×６０',
  '636196': 'PLロゴ入りネームプレート＿A',
  '636265': 'Ｃ折　２／３ナフキン　無地　２４８×１６５',
  '636529': '旗楊枝（５ヶ国）',
  '636005': '妻楊枝１７０号　ポリ容器入',
  '636269': 'サステナペーパーストロー6X200白',
  '360459': 'ｶｯﾌﾟ TK 22OZCOKE 25入20依',
  '636280': 'ｶｯﾌﾟ TK_14OZCOKE 100入20依'
};

// 消耗品セクションの商品名を、CATALOG_NAMESの値に完全一致させる（コードは変更しない）。
function syncNamesToCatalog() {
  const found = findConsumableHeader_();
  const ui = SpreadsheetApp.getUi();
  if (!found) {
    ui.alert('消耗品セクションの見出し行が見つかりませんでした。');
    return;
  }
  const sheet = found.sheet;
  const lastRow = sheet.getLastRow();
  const dataStartRow = found.headerRow + 1;
  const numRows = lastRow - dataStartRow + 1;
  if (numRows <= 0) {
    ui.alert('データ行が見つかりませんでした。');
    return;
  }
  const codes = sheet.getRange(dataStartRow, 2, numRows, 1).getValues(); // B列
  let count = 0;
  for (let i = 0; i < codes.length; i++) {
    const code = String(codes[i][0]).trim();
    if (CATALOG_NAMES[code]) {
      sheet.getRange(dataStartRow + i, 3).setValue(CATALOG_NAMES[code]);
      count++;
    }
  }
  ui.alert('商品名の統一完了', count + '件の商品名をサイトの正式名称に統一しました。', ui.ButtonSet.OK);
}

// サイト(マイカタログ)のダウンロード順。存在しないコードは無視される。
const CATALOG_ORDER = [
  '636082','636266','636312','636313','636314','636410','636054','636055','636024',
  '636276','636092','636095','636096','636098','636525','638015','636015','636020',
  '638016','636050','638017','636088','638018','638019','638020','638021','638023',
  '636113','636186','636118','636143','636270','636271','636272','636273','636275',
  '636112','636196','636265','636529','636005','636269','360459','636280'
];

// 消耗品セクションの見出し行(「発注分類名称」/「商品コード」)を全シートから探す。
// シート名に「消耗品」を含むシートを優先する。
function findConsumableHeader_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets().sort(function (a, b) {
    const aMatch = a.getName().indexOf('消耗品') !== -1 ? 0 : 1;
    const bMatch = b.getName().indexOf('消耗品') !== -1 ? 0 : 1;
    return aMatch - bMatch;
  });
  for (const sheet of sheets) {
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 1 || lastCol < 2) continue;
    const values = sheet.getRange(1, 1, lastRow, 2).getValues();
    for (let r = 0; r < values.length; r++) {
      if (String(values[r][0]).trim() === '発注分類名称' && String(values[r][1]).trim() === '商品コード') {
        return { sheet: sheet, headerRow: r + 1 };
      }
    }
  }
  return null;
}

function reorderConsumables() {
  const ui = SpreadsheetApp.getUi();
  const found = findConsumableHeader_();
  if (!found) {
    ui.alert('消耗品セクションの見出し行が見つかりませんでした。');
    return;
  }
  const sheet = found.sheet;
  const dataStartRow = found.headerRow + 1;
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const numRows = lastRow - dataStartRow + 1;

  if (numRows <= 0 || numRows % 2 !== 0) {
    ui.alert('データ行数が想定外です（' + numRows + '行）。並び替えを中止しました。');
    return;
  }

  const range = sheet.getRange(dataStartRow, 1, numRows, lastCol);
  const values = range.getValues();
  const formulas = range.getFormulas();

  // 数式が含まれる場合は参照ズレの危険があるため中止する。
  for (let r = 0; r < formulas.length; r++) {
    for (let c = 0; c < formulas[r].length; c++) {
      if (formulas[r][c]) {
        ui.alert(
          '数式が見つかりました',
          '行' + (dataStartRow + r) + ' 列' + (c + 1) + ' に数式があります。安全のため並び替えを中止しました。数式を含む並び替えは手動確認が必要です。',
          ui.ButtonSet.OK
        );
        return;
      }
    }
  }

  // 2行1組（基準量/在庫）として現在の並びを取得
  const pairs = [];
  for (let i = 0; i < numRows; i += 2) {
    const code = String(values[i][1]).trim(); // B列
    pairs.push({ code: code, rows: [values[i], values[i + 1]] });
  }

  // サイト順で見つかったものを先に、残りは元の相対順のまま後ろに続ける
  const used = new Array(pairs.length).fill(false);
  const ordered = [];
  CATALOG_ORDER.forEach(function (code) {
    for (let i = 0; i < pairs.length; i++) {
      if (!used[i] && pairs[i].code === code) {
        ordered.push(pairs[i]);
        used[i] = true;
        break;
      }
    }
  });
  pairs.forEach(function (p, i) {
    if (!used[i]) ordered.push(p);
  });

  if (ordered.length !== pairs.length) {
    ui.alert('並び替えの件数が一致しませんでした。中止しました。');
    return;
  }

  const newValues = [];
  ordered.forEach(function (p) {
    newValues.push(p.rows[0]);
    newValues.push(p.rows[1]);
  });

  range.setValues(newValues);
  ui.alert('並び替え完了', pairs.length + '品目をサイトのダウンロード順に並び替えました。', ui.ButtonSet.OK);
}

// 消耗品セクションの商品コード・商品名を修正する。
// B列(商品コード)を検索し、一致した行のB列とC列を書き換える。
function fixConsumableCodes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const fixes = {
    '636302': { newCode: '636276', newName: 'PL油よけはちまき＿A' },
    '636515': { newCode: '636005', newName: '妻楊枝１７０号　ポリ容器入' },
    '636269': { newCode: '636269', newName: 'サステナペーパーストロー6X200白' },
    '636525': { newCode: '636525', newName: '米炊ネット（ＫＮ－０７５－Ｍ）' },
    '636529': { newCode: '636529', newName: '旗楊枝（５ヶ国）' }
  };

  const results = [];
  ss.getSheets().forEach(function (sheet) {
    const lastRow = sheet.getLastRow();
    if (lastRow < 1) return;
    const codes = sheet.getRange(1, 2, lastRow, 1).getValues(); // B列
    for (let i = 0; i < codes.length; i++) {
      const cellValue = String(codes[i][0]).trim();
      if (fixes[cellValue]) {
        const row = i + 1;
        const fix = fixes[cellValue];
        sheet.getRange(row, 2).setValue(fix.newCode);
        sheet.getRange(row, 3).setValue(fix.newName);
        results.push(sheet.getName() + ' 行' + row + ': ' + cellValue + ' → ' + fix.newCode + ' / ' + fix.newName);
      }
    }
  });

  const message = results.length > 0
    ? results.join('\n')
    : '対象コードが見つかりませんでした。';
  SpreadsheetApp.getUi().alert('修正結果', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

// 書き込みは一切行わず、現在の列構成をそのまま表示する（原因調査用）。
function diagnoseColumns() {
  const ui = SpreadsheetApp.getUi();
  const found = findConsumableHeader_();
  if (!found) {
    ui.alert('消耗品セクションの見出し行が見つかりませんでした。');
    return;
  }
  const sheet = found.sheet;
  const headerRow = found.headerRow;
  const lastCol = sheet.getLastColumn();
  const header = sheet.getRange(headerRow, 1, 1, lastCol).getDisplayValues()[0];
  const row1 = sheet.getRange(headerRow + 1, 1, 1, lastCol).getDisplayValues()[0];
  const row2 = sheet.getRange(headerRow + 2, 1, 1, lastCol).getDisplayValues()[0];

  const colLetter = function (n) {
    let s = '';
    while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
    return s;
  };

  let msg = 'シート: ' + sheet.getName() + ' / 見出し行: ' + headerRow + ' / 最終列: ' + colLetter(lastCol) + '(' + lastCol + ')\n\n';
  for (let c = 0; c < lastCol; c++) {
    msg += colLetter(c + 1) + ': [見出し]' + header[c] + ' / [行' + (headerRow + 1) + ']' + row1[c] + ' / [行' + (headerRow + 2) + ']' + row2[c] + '\n';
  }
  ui.alert('列構成の診断結果', msg, ui.ButtonSet.OK);
  Logger.log(msg);
}

// 元の「単位規格」列(F列)の値。キーは現在のシート上のコード。
const ORIGINAL_SPEC = {
  '360459': '1個', '360460': '', '636280': '本', '370070': '1個', '636012': '1個',
  '636015': '本', '636020': '個', '636024': '50枚/袋', '636050': '12', '636054': '100枚/袋',
  '636055': '100枚/袋', '636082': '枚', '636088': '枚', '636092': '枚', '636095': '枚',
  '636096': '枚', '636098': '枚', '636108': '枚', '636109': '枚', '636110': '枚', '636111': '枚',
  '636112': '1000枚/巻', '636113': '1000枚/巻', '636118': '1000枚/巻', '636186': '4巻',
  '636196': '12枚/袋', '636265': '枚', '636266': '個', '636269': '個', '636270': '個',
  '636271': '個', '636272': '個', '636273': '個', '636275': '個', '636276': '枚', '636312': '箱',
  '636313': '箱', '636314': '箱', '636410': '40P', '636005': '枚', '636525': '', '636529': '本',
  '636549': '枚', '638015': '本', '638016': '本', '638017': '本', '638018': '本', '638019': '本',
  '638020': '本', '638021': '本', '638023': '本', '636143': '枚'
};

// 消えた「単位規格」列(元F列)を、現在のF列(基準量/在庫マーカー)の手前に挿入して復元する。
function repairSpecColumn() {
  const ui = SpreadsheetApp.getUi();
  const found = findConsumableHeader_();
  if (!found) {
    ui.alert('消耗品セクションの見出し行が見つかりませんでした。');
    return;
  }
  const sheet = found.sheet;
  const headerRow = found.headerRow;
  const lastRow = sheet.getLastRow();

  // 現在のF列(6列目)が「基準量」であることを確認してから挿入する（誤爆防止）。
  const f2 = sheet.getRange(headerRow + 1, 6).getDisplayValue().trim();
  if (f2 !== '基準量') {
    ui.alert('想定と違う状態です（F列2行目=' + f2 + '）。安全のため中止しました。診断機能で確認してください。');
    return;
  }

  sheet.insertColumnBefore(6);
  sheet.getRange(headerRow, 6).setValue('単位規格');

  const dataStartRow = headerRow + 1;
  const numRows = lastRow - dataStartRow + 1; // 挿入により行数自体は変わらない
  const codes = sheet.getRange(dataStartRow, 2, numRows, 1).getValues();
  let count = 0;
  for (let i = 0; i < numRows; i += 2) {
    const code = String(codes[i][0]).trim();
    if (ORIGINAL_SPEC[code]) {
      sheet.getRange(dataStartRow + i, 6).setValue(ORIGINAL_SPEC[code]);
      count++;
    }
  }
  ui.alert('復元完了', '単位規格列を復元し、' + count + '件の値を書き戻しました。', ui.ButtonSet.OK);
}

// E列(発注単位)と内容が重複しているF列(単位規格)を削除する。
function deleteSpecColumn() {
  const ui = SpreadsheetApp.getUi();
  const found = findConsumableHeader_();
  if (!found) {
    ui.alert('消耗品セクションの見出し行が見つかりませんでした。');
    return;
  }
  const sheet = found.sheet;
  const headerRow = found.headerRow;

  const f1 = sheet.getRange(headerRow, 6).getDisplayValue().trim();
  if (f1 !== '単位規格') {
    ui.alert('想定と違う状態です（F列見出し=' + f1 + '）。安全のため中止しました。診断機能で確認してください。');
    return;
  }

  sheet.deleteColumn(6);
  ui.alert('削除完了', 'F列（単位規格）を削除しました。', ui.ButtonSet.OK);
}

// 過去の同期バグで壊れた行を一括修正する。
// (1)新規追加時にマーカー(基準量/在庫)が1列右にズレた行 → 左に詰め直す
// (2)在庫行にも商品名が重複して書かれた行 → 在庫行の商品名を消す
function repairSyncBugs() {
  const ui = SpreadsheetApp.getUi();
  const found = findConsumableHeader_();
  if (!found) {
    ui.alert('消耗品セクションの見出し行が見つかりませんでした。');
    return;
  }
  const sheet = found.sheet;
  const dataStartRow = found.headerRow + 1;
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const numRows = lastRow - dataStartRow + 1;

  const headerVals = sheet.getRange(found.headerRow, 1, 1, lastCol).getDisplayValues()[0];
  let datesStartCol = -1;
  for (let c = 5; c < lastCol; c++) {
    if (/^\d+月\d+日$/.test(String(headerVals[c]).trim())) { datesStartCol = c + 1; break; }
  }
  if (datesStartCol === -1) {
    ui.alert('日付列が見つかりませんでした。');
    return;
  }
  const markerCol = datesStartCol - 2; // 0-indexed

  const range = sheet.getRange(dataStartRow, 1, numRows, lastCol);
  const values = range.getValues();

  let shiftFixed = 0;
  let nameDupFixed = 0;

  for (let i = 0; i < numRows; i += 2) {
    const row0 = values[i];
    const row1 = values[i + 1];

    // (1) マーカーが1列右にズレている場合、そこから右を1列左に詰める
    if (row0[markerCol] !== '基準量' && markerCol + 1 < lastCol && row0[markerCol + 1] === '基準量') {
      for (let c = markerCol; c < lastCol - 1; c++) {
        row0[c] = row0[c + 1];
        row1[c] = row1[c + 1];
      }
      row0[lastCol - 1] = '';
      row1[lastCol - 1] = '';
      shiftFixed++;
    }

    // (2) 在庫行に商品名が残っている場合は消す（B列が空欄なのにC列だけ埋まっている行）
    if (String(row1[1]).trim() === '' && String(row1[2]).trim() !== '') {
      row1[2] = '';
      nameDupFixed++;
    }
  }

  range.setValues(values);

  // (3) 罫線・A〜C列/L列の2行結合が無い行に、テンプレート行と同じ書式を適用する
  let formatFixed = 0;
  const templateRange = sheet.getRange(dataStartRow, 1, 2, lastCol);
  for (let i = 0; i < numRows; i += 2) {
    const row = dataStartRow + i;
    const code = String(sheet.getRange(row, 2).getValue()).trim();
    if (!code) continue;
    const merged = sheet.getRange(row, 3).getMergedRanges();
    const isMerged = merged.length > 0 && merged[0].getNumRows() === 2;
    if (!isMerged) {
      templateRange.copyTo(sheet.getRange(row, 1, 2, lastCol), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      sheet.getRange(row, 1, 2, 1).merge();
      sheet.getRange(row, 2, 2, 1).merge();
      sheet.getRange(row, 3, 2, 1).merge();
      sheet.getRange(row, lastCol, 2, 1).merge();
      formatFixed++;
    }
  }

  ui.alert(
    '修正完了',
    'マーカー列ズレの修正: ' + shiftFixed + '件\n在庫行の商品名重複の修正: ' + nameDupFixed + '件\n罫線・結合の修正: ' + formatFixed + '件',
    ui.ButtonSet.OK
  );
}
