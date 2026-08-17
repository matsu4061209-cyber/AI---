function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🔧 コード修正')
    .addItem('商品コード一括修正を実行', 'fixConsumableCodes')
    .addItem('商品名をサイト名に統一', 'syncNamesToCatalog')
    .addItem('単価・発注単位をサイトからコピー', 'syncPriceAndUnitFromCatalog')
    .addItem('消耗品をサイト順に並び替え', 'reorderConsumables')
    .addToUi();
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
