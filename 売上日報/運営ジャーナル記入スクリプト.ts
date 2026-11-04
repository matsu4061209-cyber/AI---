/*************************************************************
 * 運営ジャーナル 自動記入スクリプト（v2・Power Automate対応）
 *
 * 2通りの使い方に対応：
 *  ① Power Automate から：数字の文字列を「data（入力）」として受け取って記入
 *  ② 手動（今まで通り）：PL飯塚シートの BZ1 に貼り付けてから実行
 *
 * 文字列の並び：
 *  日 | 天候 | 税抜売上 | 客数 | 現金過不足 | UberEats | 出前館 | ゆめ | 8%対象額
 *
 * 書き込み先（PL飯塚シート・行=日付+5／1日=6行目）：
 *  D=天候  L=売上  N=客数  P=現金過不足  AR=UberEats  AS=出前館  AO=ゆめか  AQ=テイクアウト8%
 *  ※累計・前年・客単価などの数式セルには一切触れません。
 *************************************************************/
function main(workbook: ExcelScript.Workbook, data?: string) {
  // 書き込み先シート＝PL飯塚（無ければ今開いているシート）
  const sheet = workbook.getWorksheet("PL飯塚") || workbook.getActiveWorksheet();

  // データの取得元：Power Automateから渡されたら それを使う。
  // 渡されなければ、手動用に BZ1 を読む。
  let raw = "";
  let fromBZ1 = false;
  if (data !== undefined && data !== null && String(data).trim() !== "") {
    raw = String(data);
  } else {
    raw = String(sheet.getRange("BZ1").getValue());
    fromBZ1 = true;
  }

  if (raw.trim() === "") {
    console.log("データが空です。文字列を渡すか、BZ1に貼り付けてから実行してください。");
    return;
  }

  const parts = raw.split("|");
  if (parts.length < 9) {
    console.log("データの形がちがいます（9項目必要）。中身: " + raw);
    return;
  }

  const day = parseInt(parts[0], 10);
  if (isNaN(day) || day < 1 || day > 31) {
    console.log("日付が読めません: " + parts[0]);
    return;
  }
  const row = day + 5;   // 1日→6行目

  // 文字列を数値に（カンマや空白は除去。空なら空のまま）
  const num = (s: string): number | string => {
    const cleaned = String(s).replace(/,/g, "").trim();
    if (cleaned === "") return "";
    const n = parseFloat(cleaned);
    return isNaN(n) ? "" : n;
  };

  const tenko = String(parts[1]).trim();

  sheet.getRange("D" + row).setValue(tenko);          // 天候
  sheet.getRange("L" + row).setValue(num(parts[2]));  // 売上（税抜）
  sheet.getRange("N" + row).setValue(num(parts[3]));  // 客数
  sheet.getRange("P" + row).setValue(num(parts[4]));  // 現金過不足
  sheet.getRange("AR" + row).setValue(num(parts[5])); // UberEats
  sheet.getRange("AS" + row).setValue(num(parts[6])); // 出前館
  sheet.getRange("AO" + row).setValue(num(parts[7])); // ゆめか
  sheet.getRange("AQ" + row).setValue(num(parts[8])); // テイクアウト8%

  // 手動（BZ1）で読んだ時だけ、後片付けでBZ1を消す
  if (fromBZ1) {
    sheet.getRange("BZ1").clear(ExcelScript.ClearApplyTo.ApplyToContents);
  }

  console.log(day + "日（" + row + "行目）に記入しました。");
}
