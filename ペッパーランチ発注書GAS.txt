// ─────────────────────────────────────────────────────────────────────
// デバッグ：商品マスタの登録状態を確認
// ─────────────────────────────────────────────────────────────────────
// コードを比較用に正規化：前後空白除去・全角数字→半角・末尾の .0 除去
// （数値でも文字列でも同じ結果になり、型や書式の違いで照合が外れるのを防ぐ）
function normalizeCode_(v) {
  if (v === null || v === undefined) return '';
  let s = String(v).trim();
  if (s === '') return '';
  s = s.replace(/[０-９]/g, function(c){ return String.fromCharCode(c.charCodeAt(0) - 0xFEE0); }); // 全角→半角
  s = s.replace(/\.0+$/, ''); // 4982.0 → 4982
  return s;
}

// 未登録コードの原因を診断してメール送信（メニューから手動実行）
function debugMenuCodes_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const menuSheet = ss.getSheetByName('商品マスタ');
  const jSheet = ss.getSheetByName('実績データ');
  const targets = ['4982', '4986', '4987', '5038'];
  const L = [];

  if (!menuSheet) { MailApp.sendEmail('matsu4061209@gmail.com', '【デバッグ】エラー', '商品マスタが見つかりません'); return; }
  const mLast = menuSheet.getLastRow();
  const mCol = menuSheet.getRange(2, 1, mLast - 1, 1).getValues();
  const exactSet = new Set(), normSet = new Set();
  mCol.forEach(r => { if (r[0] !== '' && r[0] !== null) { exactSet.add(String(r[0])); normSet.add(normalizeCode_(r[0])); } });

  L.push('■商品マスタ 総行数:' + mLast + ' / 登録コード数(正規化):' + normSet.size);
  L.push('A列 先頭3件 (型:値) → ' + mCol.slice(0, 3).map(r => typeof r[0] + ':[' + r[0] + ']').join(' / '));
  L.push('');
  L.push('【判定】 コード : 完全一致 / 正規化一致 / 該当セルの(型:値)');
  targets.forEach(c => {
    const ex = exactSet.has(c);
    const nm = normSet.has(c);
    const hit = mCol.find(r => normalizeCode_(r[0]) === c);
    const detail = hit ? (typeof hit[0] + ':[' + hit[0] + ']') : 'なし';
    L.push('  ' + c + ' : 完全' + (ex ? '○' : '×') + ' 正規化' + (nm ? '○' : '×') + ' → ' + detail);
  });

  L.push('');
  L.push('■実績データ E列（直近から探す） コード : (型:値) 名前');
  if (jSheet) {
    const jLast = jSheet.getLastRow();
    const from = Math.max(2, jLast - 800);
    const jRows = jSheet.getRange(from, 1, jLast - from + 1, 6).getValues();
    targets.forEach(c => {
      let f = null;
      for (let i = jRows.length - 1; i >= 0; i--) { if (normalizeCode_(jRows[i][4]) === c) { f = jRows[i]; break; } }
      L.push('  ' + c + ' : ' + (f ? (typeof f[4] + ':[' + f[4] + '] ' + f[5]) : '直近に無し'));
    });
  }

  const msg = L.join('\n');
  MailApp.sendEmail('matsu4061209@gmail.com', '【デバッグ】未登録コード診断', msg);
  console.log(msg);
}

// ─────────────────────────────────────────────────────────────────────
// 実績の訂正を後工程へ反映する（実績データ=元 を正として、日次売上実績・売上予測を揃え直す）
//   CSV/実績データを手で直した後に1回押せば、日次売上実績と売上予測の該当日が正しい値に揃う。
//   ・実績データの各日の総売上・総客数（メニュー行に同じ値が入っている）を正とする。
//   ・日次売上実績で食い違う日を修正（無い日は追加）。
//   ・売上予測の該当日の 実績売上(K)・誤差率(L) を正しい値に直す（誤差率=(予測−実績)÷実績）。
//   ※実績確定行は再実行では上書きされない（凍結）ため、この関数で直接直す必要がある。
// ─────────────────────────────────────────────────────────────────────
function resyncActualsFromRaw() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const raw = ss.getSheetByName('実績データ');
  const daily = ss.getSheetByName('日次売上実績');
  const fc = ss.getSheetByName('売上予測');
  if (!raw || !daily) { try { SpreadsheetApp.getUi().alert('「実績データ」または「日次売上実績」シートが見つかりません'); } catch (e) {} return; }

  // 1. 実績データ(A日付 B曜日名 C総売上 D総客数)から日別の正しい値を集計（同一日はメニュー行に同値が入る）
  const rawVals = raw.getRange(2, 1, raw.getLastRow() - 1, 4).getValues();
  const truth = {};
  rawVals.forEach(r => {
    if (!r[0]) return;
    const ds = toDateStr_(r[0]);
    const s = Number(r[2]) || 0, c = Number(r[3]) || 0;
    if (!truth[ds]) truth[ds] = { sales: s, customers: c };
    else { if (s > truth[ds].sales) truth[ds].sales = s; if (c > truth[ds].customers) truth[ds].customers = c; }
  });

  const changed = [];
  const DOWNAME = ['日', '月', '火', '水', '木', '金', '土'];

  // 2. 日次売上実績(A日付 B曜日コード C曜日名 D総売上 E総客数)を突き合わせて修正
  const dLast = daily.getLastRow();
  const seen = {};
  if (dLast >= 2) {
    const dVals = daily.getRange(2, 1, dLast - 1, 5).getValues();
    for (let i = 0; i < dVals.length; i++) {
      if (!dVals[i][0]) continue;
      const ds = toDateStr_(dVals[i][0]); seen[ds] = true;
      const t = truth[ds]; if (!t) continue;
      const curS = Number(dVals[i][3]) || 0, curC = Number(dVals[i][4]) || 0;
      if (curS !== t.sales || curC !== t.customers) {
        daily.getRange(i + 2, 4, 1, 2).setValues([[t.sales, t.customers]]);
        changed.push({ ds, oldSales: curS, newSales: t.sales });
      }
    }
  }
  // 実績データにあって日次売上実績に無い日は追加
  const toAppend = [];
  Object.keys(truth).forEach(ds => {
    if (seen[ds]) return;
    const d = new Date(ds), code = d.getDay() + 1;
    toAppend.push([d, code, DOWNAME[d.getDay()], truth[ds].sales, truth[ds].customers]);
    changed.push({ ds, oldSales: 0, newSales: truth[ds].sales });
  });
  if (toAppend.length > 0) daily.getRange(Math.max(dLast, 1) + 1, 1, toAppend.length, 5).setValues(toAppend);

  // 3. 変わった日付について、売上予測の実績(K=11列)・誤差(L=12列)を直す
  const changedSet = {}; changed.forEach(c => changedSet[c.ds] = true);
  let fcFixed = 0;
  const todayMid = new Date(); todayMid.setHours(0, 0, 0, 0);
  if (fc && fc.getLastRow() >= 2) {
    const fVals = fc.getRange(2, 1, fc.getLastRow() - 1, 12).getValues();
    for (let i = 0; i < fVals.length; i++) {
      if (!fVals[i][0]) continue;
      const ds = toDateStr_(fVals[i][0]);
      if (!changedSet[ds] || !truth[ds]) continue;
      const forecast = Number(fVals[i][8]) || 0, actual = truth[ds].sales;
      const err = (actual > 0 && forecast > 0) ? Math.round((forecast - actual) / actual * 1000) / 1000 : '';
      fc.getRange(i + 2, 11, 1, 2).setValues([[actual, err]]);
      if ((fVals[i][9] || '').toString() !== '実績確定' && new Date(ds) < todayMid && actual > 0) fc.getRange(i + 2, 10).setValue('実績確定');
      fcFixed++;
    }
  }

  const msg = changed.length === 0
    ? '実績データ・日次売上実績・売上予測は既に一致しています（修正なし）。'
    : '実績データ（元）を正として ' + changed.length + '日分を直しました。\n\n'
      + changed.map(c => '・' + c.ds + '：' + Number(c.oldSales).toLocaleString() + ' → ' + Number(c.newSales).toLocaleString()).join('\n')
      + '\n\n売上予測の実績・誤差も ' + fcFixed + ' 行 直しました。';
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { console.log(msg); }
}

// ─────────────────────────────────────────────────────────────────────
// カスタムメニュー
// ─────────────────────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getActiveSpreadsheet().addMenu('📅 イベント管理', [
    { name: 'イベント追加フォーム',            functionName: 'showEventForm_' },
    { name: 'キャンペーン商品の登録フォーム',   functionName: 'showCpForm_' },
    { name: '暫定日イベントの日付確認メールを今すぐ送る', functionName: 'remindProvisionalEventsNow_' },
    { name: 'デバッグ：未登録コード診断メール', functionName: 'debugMenuCodes_' },
    null,
    { name: '売上予測を更新（buildSalesForecast）', functionName: 'buildSalesForecast' },
    { name: '実績データから日次売上・売上予測を直す（訂正の反映）', functionName: 'resyncActualsFromRaw' },
    { name: '予測精度を分析（日別＋週別合計）', functionName: 'analyzeForecastAccuracy' }
  ]);
  SpreadsheetApp.getActiveSpreadsheet().addMenu('💰 原価管理', [
    { name: '理論原価率を計算（原価分析シートを更新）', functionName: 'buildCostAnalysis' },
    { name: '理論原価を今すぐ月次保存（先月分）', functionName: 'saveLastMonthCostSnapshotNow' },
    { name: '過去の全月を今すぐ保存（コード削除の前に推奨）', functionName: 'saveAllMonthlyCostSnapshots' },
    { name: '履歴の理論原価を今の状態で再計算・上書き（食材復元後の修正用／実原価は保持）', functionName: 'refreshTheoreticalCostHistory' },
    { name: '毎月1日に理論原価を自動保存をONにする', functionName: 'installMonthlyCostSnapshotTrigger' },
    null,
    { name: '納品データCSVを取り込む（商品別棚卸高→納品実績）', functionName: 'importDeliveryInventoryCsv' },
    { name: '実際原価率を反映（納品実績→原価月次履歴）', functionName: 'updateActualCostFromDelivery' },
    { name: '実損分析（食材別に実際vs理論の使用金額＝ロスを洗い出す）', functionName: 'buildLossAnalysis' },
    { name: '米使用分析（1食あたり生米gを理論vs実際で確認）', functionName: 'analyzeRiceUsage' },
    { name: 'ドリンク実測分析（実際使用÷出食数で1杯単価を実測）', functionName: 'analyzeDrinkUsage' },
    { name: 'ドリンクの単位使用量を実測に合わせて更新（旧レシピは履歴保存）', functionName: 'adjustDrinkUsageToActual' },
    null,
    { name: '原価抜けチェック（商品マスタが使う食材で未登録を検出）', functionName: 'checkMissingIngredients' },
    { name: 'チェックした食材を食材マスタに復元', functionName: 'applyIngredientRestore' },
    null,
    { name: '原価調整設定を作成・旧コードの読み替え候補を追加', functionName: 'setupCostAdjustSheet' },
    null,
    { name: '食材コードの切替を今すぐ実行', functionName: 'applyFoodCodeSwitches' },
    { name: '商品マスタの#REF!を直す（切替のスピル復活）', functionName: 'repairProductMasterSpill' },
    null,
    { name: 'コードの稼働チェック（旧コード・再利用・切替の検出）', functionName: 'buildCodeUsageCheck' },
    { name: 'チェックした商品コードを削除待ちに移す（元に戻せます）', functionName: 'archiveCheckedProductCodes' },
    { name: '削除待ちから商品コードを元に戻す', functionName: 'restoreArchivedProductCodes' },
    null,
    { name: '【1回だけ】ブロッコリー(616661)を登録', functionName: 'registerBroccoliSwitch_202608' }
  ]);
  SpreadsheetApp.getActiveSpreadsheet().addMenu('🖨 発注ワークシート', [
    { name: '印刷用に空行を整理する', functionName: 'trimOrderSheetRows' },
    { name: '「在庫」と「基準」の列を入れ替える（もう一度押すと元に戻る）', functionName: 'swapOrderSheetStockPar' },
    null,
    { name: 'トレンド履歴の曜日別平均を再計算', functionName: 'buildTrendHistoryAverages' },
    { name: '食材マスタを区分順・コード順に並べ替える', functionName: 'sortIngredientMaster' },
    { name: '基準を細かい数字にする（1以上は整数）', functionName: 'setOrderSheetParFormula' },
    { name: '横型の発注書を作り直す（食材を変えた時）', functionName: 'buildOrderSheetWide' },
    { name: '今週分で作り直す（臨時・予測変更や商品追加を今すぐ反映）', functionName: '発注ワークシート_今週分で作り直す' },
    { name: '仕込み・解凍表を作る（来週分・自動と同じ）', functionName: 'buildPrepThawSheet' },
    { name: '仕込み・解凍表を作る（今週分）', functionName: 'buildPrepThawSheetThisWeek' },
    { name: '仕込み・解凍表の入力フォームを開く（在庫・調整・品目）', functionName: 'showPrepForm' },
    { name: '【1回だけ】主要9品目の安全係数を最適化（欠品対策）', functionName: 'setPrepSafetyFactors' },
    { name: 'コーンを「缶」単位に直す', functionName: 'fixCornCanUnit' },
    { name: '仕込みマスタの重複を整理（ブロッコリーも300gに）', functionName: 'fixPrepMasterDedupe' }
  ]);
  SpreadsheetApp.getActiveSpreadsheet().addMenu('🗂 シート整理', [
    { name: '使っていないシートを「_削除待ち_」にする（削除はしません）', functionName: 'archiveUnusedSheets' },
    { name: '「_削除待ち_」のシートを元に戻す', functionName: 'restoreArchivedSheets' }
  ]);
  SpreadsheetApp.getActiveSpreadsheet().addMenu('🔧 メンテナンス', [
    { name: '⚠️ 今すぐデータ復旧（天候・売上・予測が止まった時）', functionName: 'recoverDailyJobsNow' },
    { name: 'データの新しさを今すぐチェック', functionName: 'checkDataFreshnessAlert' },
    { name: '毎日の自動処理（トリガー）を入れ直す', functionName: 'installDailyTriggers' }
  ]);
  try { warnIfDataStale_(); } catch(e) {}          // 天候・売上・予測が止まっていたら開いた時に警告
  // ※起動時は売上予測に飛ばさない。売上予測シートを選んだ時に onSelectionChange で「予報反映中」を先頭表示する。
}

// 売上予測シートを「選んだ時だけ」、予報反映中の行を画面の一番上に表示する。
//   （起動時に毎回飛ばされるのを避け、ユーザーが売上予測タブを開いた瞬間にだけ動かす）
//   onSelectionChange は簡易トリガー。同じシート内のセル移動では動かないよう、直前のシート名を記録して切替時だけ実行。
function onSelectionChange(e) {
  try {
    const ss = (e && e.source) ? e.source : SpreadsheetApp.getActiveSpreadsheet();
    const cur = ss.getActiveSheet().getName();
    const props = PropertiesService.getDocumentProperties();
    const last = props.getProperty('lastActiveSheetName') || '';
    if (cur === last) return;                       // 同じシート内のクリックでは何もしない（＝飛ばない）
    props.setProperty('lastActiveSheetName', cur);
    if (cur === '売上予測') jumpToForecastPending_();
  } catch (err) {}
}

// 起動時に売上予測を開き、「予報反映中」の最初の行が画面の一番上に来るようにする。
function jumpToForecastPending_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('売上予測');
  if (!sh || sh.getLastRow() < 2) return;
  const vals = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  let target = -1;
  for (let r = 1; r < vals.length; r++) {
    if (vals[r].some(c => String(c).indexOf('予報反映中') >= 0)) { target = r + 1; break; }
  }
  if (target < 0) return;
  sh.activate();
  // activateは「対象セルを画面の下端寄りに表示」する性質がある。
  // なので予報反映中の約30行下のセルを1回だけactivateすると、予報反映中がちょうど画面の一番上に来る。
  // （そのあと予報反映中をactivateし直すと下端に戻ってしまうので、あえて1回だけにする）
  const bottomRow = Math.min(sh.getLastRow(), target + 29);
  sh.getRange(bottomRow, 1).activate();
}

/**gsgs
 * 統合_発注システム.gs  ── 完成版
 * ═══════════════════════════════════════════════════
 * 【統合元】
 *   データ読込.gs / 売上予測.gs / 発注処理.gs
 *   異常値処理.gs / 祝日マスタ.gs / 祝日倍率マスタ.gs
 *
 * 【主な修正・改善点】
 *   ① buildIngredientDailyUsage() の二重実行を解消
 *      → CacheService に保存し次工程はキャッシュから読む
 *   ② 天気情報は23時間以内なら再取得しない（1日1回制限）
 *   ③ 売上予測.gs 末尾の発注計算自動呼び出しを削除（タイムアウト解消）
 *   ④ errAdj（誤差率補正）を廃止（実績ベースの平均に二重補正になるため）
 *   ⑤ buildOrderBufferO_ に28日フィルターを追加（期間を統一）
 *   ⑥ サンプル2件未満の曜日は全曜日平均で補完
 *   ⑦ 祝日マスタに登録された日はIQR判定から除外
 *      → 曜日別平均には影響しないよう別途除外
 *   ⑧ 取込停止時にメール通知＋わかりやすいメッセージ
 *
 * 【トリガー設定（初回のみ setupTriggersSplit() を実行）】
 *   06:00 importFooditCSV                  CSV取込
 *   06:15 fetchWeather_OpenMeteo_UpdateDB  天候更新
 *   07:00 buildSalesForecast              売上予測（importFooditCSV完了後に確実に実行するため7:00に設定）
 *   07:20 発注バッファ_ステップ1             発注計算①
 *   07:40 発注バッファ_ステップ2             発注計算②
 *   ※毎月1日 05:00  updateSmartHolidayMaster（土日平均更新）
 *   ※毎週月曜 06:40 発注バッファ_トレンド更新
 *
 * 【トリガー設定履歴】
 *   2026-07-13  buildSalesForecast を 06:30 → 07:00 に変更
 *               発注バッファ_ステップ1を 06:50 → 07:20、ステップ2を 07:10 → 07:40 に変更
 *               ← importFooditCSV（最大6分かかる場合あり）が必ず完了した後に売上予測が動くよう修正
 *               ← setupTriggersSplit() を再実行して反映すること
 *
 *   2026-07-07  setupTriggersSplit() 実行
 *               ← 水曜予測が60,516円に固着する問題を受け、毎日自動実行を設定
 *               ← buildSalesForecast が手動実行のみだったため古い曜日別平均が残っていた
 *
 *   2026-07-07  売上予測ロジック大幅改善
 *               【変更A】avgExcludedDates から cpEventDates（日付指定CP/イベント期間）を除外しない
 *                 → 期間指定CPの実績もサンプルとして使い、曜日別平均のサンプル不足を解消
 *                 → イベント効果は eventActualRates / repeatEventCoeffMap の倍率学習で反映
 *               【変更B】毎年固定日イベント（ワールドペッパーライスデー7/3等）は avgExcludedDates に追加
 *                 → 売上が通常と大きく異なる特定日のみ除外（毎月系・肉の日29日等は除外しない）
 *               【変更C】祝日マスタI列（売上実績平均）をタグ別に自動計算・書き込み
 *                 → GW・振替休日等の予測をタグ別実績平均から算出（従来は通常曜日平均を誤使用）
 *               ★バックアップ: ペッパーランチ発注書GAS_backup_20260707.txt
 *
 * 【単独実行できる関数】
 *   importFooditCSV()
 *   fetchWeather_OpenMeteo_UpdateDB()
 *   buildIngredientDailyUsage()
 *   buildSalesForecast()
 *   発注計算バッファ_完全生成()
 *   buildOutlierLog()
 *   updateSmartHolidayMaster()
 *   buildHolidayRateMaster()
 *   recoverFromDate()        ← 日付飛び復旧用
 *   setupTriggers()          ← 初回1回だけ
 */


// ═══════════════════════════════════════════════════════════════════════
// 【A】 定数
// ═══════════════════════════════════════════════════════════════════════

const LEAD_TIME_DAYS     = 2;
const ORDER_BUFFER_DAYS  = 3;
const WRITE_START_DAYS   = LEAD_TIME_DAYS + ORDER_BUFFER_DAYS;
const NEW_PRODUCT_RATE   = 0.1;
const NEW_PRODUCT_SAFETY = 1.5;
const WEEK_WEIGHTS       = [5, 3, 1];
const OUTLIER_IQR_MULTIPLIER = 1.5;
const OUTLIER_MIN_SAMPLES    = 5;


// ═══════════════════════════════════════════════════════════════════════
// 【B】 キャッシュ管理
// ═══════════════════════════════════════════════════════════════════════

function saveCpSubtractMap_(map) {
  try {
    const cache = CacheService.getScriptCache();
    const json  = JSON.stringify(map || {});
    if (json.length < 90000) {
      cache.put('cpSubtractMap', json, 21600);
    } else {
      PropertiesService.getScriptProperties().setProperty('cpSubtractMap_large', json);
    }
    PropertiesService.getScriptProperties()
      .setProperty('cpSubtractMap_time', new Date().toISOString());
  } catch(e) {
    console.log('cpSubtractMapキャッシュ保存スキップ: ' + e.message);
  }
}

function loadCpSubtractMap_() {
  try {
    const cache = CacheService.getScriptCache();
    const json  = cache.get('cpSubtractMap');
    if (json) return JSON.parse(json);
    const props   = PropertiesService.getScriptProperties();
    const savedAt = props.getProperty('cpSubtractMap_time');
    if (!savedAt) return null;
    const age = (new Date() - new Date(savedAt)) / 3600000;
    if (age > 6) return null;
    const large = props.getProperty('cpSubtractMap_large');
    return large ? JSON.parse(large) : null;
  } catch(e) {
    return null;
  }
}

function isWeatherFetchedToday_() {
  try {
    const last  = PropertiesService.getScriptProperties().getProperty('lastWeatherFetch');
    if (!last) return false;
    return (new Date() - new Date(last)) / 3600000 < 23;
  } catch(e) { return false; }
}

function markWeatherFetched_() {
  PropertiesService.getScriptProperties()
    .setProperty('lastWeatherFetch', new Date().toISOString());
}


// ═══════════════════════════════════════════════════════════════════════
// 【C】 トリガー設定
// ═══════════════════════════════════════════════════════════════════════

function setupTriggers() {
  const targets = [
    'importAndFetchWeather', 'buildSalesForecast',
    '発注計算バッファ_完全生成', '食材マスタ重複チェック'
  ];
  ScriptApp.getProjectTriggers().forEach(t => {
    if (targets.includes(t.getHandlerFunction())) ScriptApp.deleteTrigger(t);
  });
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.newTrigger('importAndFetchWeather')
    .timeBased().atHour(6).everyDays(1).create();
  ScriptApp.newTrigger('buildSalesForecast')
    .timeBased().atHour(6).nearMinute(20).everyDays(1).create();
  ScriptApp.newTrigger('発注計算バッファ_完全生成')
    .timeBased().atHour(6).nearMinute(40).everyDays(1).create();
  ScriptApp.newTrigger('食材マスタ重複チェック')
    .forSpreadsheet(ss).onEdit().create();
  console.log('トリガー設定完了\n  06:00 importAndFetchWeather\n  06:20 buildSalesForecast\n  06:40 発注計算バッファ_完全生成');
}

function setupTriggersSplit() {
  const targets = [
    'importAndFetchWeather', 'importFooditCSV', 'fetchWeather_OpenMeteo_UpdateDB',
    'buildSalesForecast', '発注計算バッファ_完全生成', '発注バッファ_ステップ1', '発注バッファ_ステップ2',
    '発注バッファ_トレンド更新', '食材マスタ重複チェック', 'updateSmartHolidayMaster', 'dailyDataWatchdog'
  ];
  ScriptApp.getProjectTriggers().forEach(t => {
    if (targets.includes(t.getHandlerFunction())) ScriptApp.deleteTrigger(t);
  });
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // 毎日トリガー
  ScriptApp.newTrigger('importFooditCSV').timeBased().atHour(6).everyDays(1).create();
  ScriptApp.newTrigger('fetchWeather_OpenMeteo_UpdateDB').timeBased().atHour(6).nearMinute(15).everyDays(1).create();
  ScriptApp.newTrigger('buildSalesForecast').timeBased().atHour(7).everyDays(1).create();
  ScriptApp.newTrigger('発注バッファ_ステップ1').timeBased().atHour(7).nearMinute(20).everyDays(1).create();
  ScriptApp.newTrigger('発注バッファ_ステップ2').timeBased().atHour(7).nearMinute(40).everyDays(1).create();
  // 見張り（毎朝9時）：天候・売上・予測が古ければ自己修復＋失敗ならメール通知
  ScriptApp.newTrigger('dailyDataWatchdog').timeBased().atHour(9).everyDays(1).create();
  // 週1回トリガー（月曜06:40）
  ScriptApp.newTrigger('発注バッファ_トレンド更新').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(6).nearMinute(40).create();
  // 月1回トリガー（毎月1日 05:00）土日の売上平均を最新化
  ScriptApp.newTrigger('updateSmartHolidayMaster').timeBased().onMonthDay(1).atHour(5).create();
  // スプレッドシートトリガー
  ScriptApp.newTrigger('食材マスタ重複チェック').forSpreadsheet(ss).onEdit().create();
  console.log('トリガー設定完了\n【毎日】\n  06:00 importFooditCSV\n  06:15 fetchWeather\n  07:00 buildSalesForecast\n  07:20 発注バッファ_ステップ1\n  07:40 発注バッファ_ステップ2\n  09:00 dailyDataWatchdog（見張り）\n【週1・月曜】\n  06:40 発注バッファ_トレンド更新\n【月1・1日】\n  05:00 updateSmartHolidayMaster（土日平均更新）');
}

// ─────────────────────────────────────────────────────────────────────
// データの新しさ監視＆自己修復（トリガーが止まっても気づける安全網）
//   天候(実測)・売上・売上予測の最終日をチェックし、古ければ警告・自動復旧する。
//   ・onOpen(warnIfDataStale_) … 開いた時に古ければトーストで警告（トリガー不要で必ず動く）
//   ・dailyDataWatchdog … 毎朝9時に古ければ自動復旧を試み、直らなければメール通知
//   ・recoverDailyJobsNow … メニュー1発でトリガー再設定＋取込・天候・予測を再実行（権限も承認できる）
// ─────────────────────────────────────────────────────────────────────
function fmtD_(d) { return d ? Utilities.formatDate(d, 'Asia/Tokyo', 'M/d') : 'なし'; }

function checkDataFreshness_(ss) {
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const behind = d => d ? Math.round((today - d) / 86400000) : 9999;
  const maxDateWhere = (name, cols, pred) => {
    const sh = ss.getSheetByName(name);
    if (!sh || sh.getLastRow() < 2) return null;
    let m = null;
    sh.getRange(2, 1, sh.getLastRow() - 1, cols).getValues().forEach(r => {
      if (r[0] instanceof Date && (!pred || pred(r)) && (!m || r[0] > m)) m = r[0];
    });
    return m;
  };
  const wLast = maxDateWhere('天候生データ', 5, r => r[4] === 'actual');   // 天候の実測最終日
  const sLast = maxDateWhere('実績データ', 1, null);                        // 売上の最終日
  const fLast = maxDateWhere('売上予測', 1, null);                          // 売上予測の最終日（本来は未来日まである）
  return { wLast, sLast, fLast, wBehind: behind(wLast), sBehind: behind(sLast), fLast_before_today: !!(fLast && fLast < today) };
}

// 開いた時：古ければトーストで警告（トリガーに依存せず必ず動く）
function warnIfDataStale_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const f = checkDataFreshness_(ss);
  const probs = [];
  if (f.wBehind >= 3) probs.push('天候' + f.wBehind + '日遅れ');
  if (f.sBehind >= 3) probs.push('売上' + f.sBehind + '日遅れ');
  if (f.fLast_before_today) probs.push('売上予測が本日分なし');
  if (probs.length) {
    ss.toast('⚠️ ' + probs.join('・') + '。メニュー「🔧 メンテナンス → 今すぐデータ復旧」を押してください。', 'データ更新が止まっている可能性', 60);
  }
}

// メニュー：データの新しさを表示
function checkDataFreshnessAlert() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); const ui = SpreadsheetApp.getUi();
  const f = checkDataFreshness_(ss);
  const ok = f.wBehind < 3 && f.sBehind < 3 && !f.fLast_before_today;
  ui.alert('データの新しさ',
    '天候の実測：' + fmtD_(f.wLast) + '（' + f.wBehind + '日前）\n' +
    '売上：' + fmtD_(f.sLast) + '（' + f.sBehind + '日前）\n' +
    '売上予測の最終：' + fmtD_(f.fLast) + '\n\n' +
    (ok ? '正常です。' : '→ 古いものがあります。「今すぐデータ復旧」を押してください。'), ui.ButtonSet.OK);
}

// メニュー：1発で復旧（トリガー再設定＋取込・天候・予測を再実行）。UrlFetchの権限承認もここで通る
function recoverDailyJobsNow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); const ui = SpreadsheetApp.getUi();
  const steps = [];
  const run = (label, fn) => { try { fn(); steps.push('✓ ' + label); } catch (e) { steps.push('✗ ' + label + '：' + (e.message || e)); } };
  run('毎日の自動処理（トリガー）を入れ直し', () => setupTriggersSplit());
  run('売上（実績データ）を取込', () => importFooditCSV());
  run('天候データを更新', () => fetchWeather_OpenMeteo_UpdateDB());
  run('売上予測を再計算', () => buildSalesForecast());
  const f = checkDataFreshness_(ss);
  ui.alert('データ復旧の結果',
    steps.join('\n') + '\n\n【現在の最新日】\n' +
    '天候の実測：' + fmtD_(f.wLast) + '（' + f.wBehind + '日前）\n' +
    '売上：' + fmtD_(f.sLast) + '（' + f.sBehind + '日前）\n' +
    '売上予測の最終：' + fmtD_(f.fLast), ui.ButtonSet.OK);
}

// メニュー：トリガーだけ入れ直す
function installDailyTriggers() {
  const ui = SpreadsheetApp.getUi();
  try { setupTriggersSplit(); ui.alert('毎日の自動処理（トリガー）を入れ直しました。\n（初回は権限の許可ダイアログが出たら「許可」してください）'); }
  catch (e) { ui.alert('トリガー設定でエラー：' + (e.message || e)); }
}

// 見張り本体（毎朝9時トリガー）：古ければ自己修復し、直らなければメール通知
function dailyDataWatchdog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const isStale = f => f.wBehind >= 2 || f.sBehind >= 2 || f.fLast_before_today;
  let f = checkDataFreshness_(ss);
  if (!isStale(f)) return;   // 正常

  const log = [];
  const run = (label, fn) => { try { fn(); log.push('OK ' + label); } catch (e) { log.push('NG ' + label + '：' + (e.message || e)); } };
  run('トリガー再設定', () => setupTriggersSplit());   // 無効化されたトリガーを復活
  run('importFooditCSV', () => importFooditCSV());
  run('fetchWeather', () => fetchWeather_OpenMeteo_UpdateDB());
  run('buildSalesForecast', () => buildSalesForecast());

  f = checkDataFreshness_(ss);
  if (isStale(f)) {
    // 自動では直らなかった（多くは外部通信の権限切れ）→ メールで手動復旧を促す
    try {
      MailApp.sendEmail(Session.getEffectiveUser().getEmail(),
        '【発注システム】データ更新が止まっています（要・手動復旧）',
        '自動復旧を試みましたが、まだ古いままです。\n' +
        '天候の実測：' + fmtD_(f.wLast) + '（' + f.wBehind + '日前）\n売上：' + fmtD_(f.sLast) + '（' + f.sBehind + '日前）\n売上予測の最終：' + fmtD_(f.fLast) + '\n\n' +
        '【対処】スプレッドシートを開き、メニュー「🔧 メンテナンス → 今すぐデータ復旧」を押してください。\n' +
        '（外部通信の権限が切れている場合は、その時に許可ダイアログが出るので「許可」してください）\n\n' +
        '実行ログ：\n' + log.join('\n'));
    } catch (e) { console.log('watchdog メール失敗: ' + (e.message || e)); }
  } else {
    console.log('dailyDataWatchdog: 自己修復しました / ' + log.join(' / '));
  }
}


// ═══════════════════════════════════════════════════════════════════════
// 【D】 データ読込セクション
// ═══════════════════════════════════════════════════════════════════════

function importAndFetchWeather() {
  try {
    importFooditCSV();
  } catch(e) {
    // CSV取込が失敗しても天候更新は必ず続行する
    console.error('CSV取込エラー（天候更新は継続）: ' + e.message);
  }
  if (isWeatherFetchedToday_()) {
    console.log('天気データは本日取得済みのためスキップ');
  } else {
    fetchWeather_OpenMeteo_UpdateDB();
    markWeatherFetched_();
  }
}

// ─────────────────────────────────────────────────────────────────────
// メイン：POSからCSVを読み込む
// ─────────────────────────────────────────────────────────────────────

function importFooditCSV() {

  const FOLDER_ID = '1UdQiP4XYGKZ89qGmq4yMOAcC0AGujYH8';
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('実績データ') || ss.insertSheet('実績データ');

  // ── 除外日セットを取得 ───────────────────────────────────────────
  const excludedDates = new Set();
  const evSheet = ss.getSheetByName('イベント登録');
  if (evSheet && evSheet.getLastRow() > 1) {
    evSheet.getRange(2, 1, evSheet.getLastRow() - 1, 5).getValues().forEach(row => {
      if (!row[0] || (row[4] || '').toString().trim() !== '除外') return;
      const start = row[1] instanceof Date ? row[1] : new Date(row[1]);
      const end   = row[2] instanceof Date ? row[2] : new Date(row[2]);
      if (!start || !end || isNaN(start) || isNaN(end)) return;
      const cur = new Date(start); cur.setHours(0, 0, 0, 0);
      const endD = new Date(end);  endD.setHours(0, 0, 0, 0);
      while (cur <= endD) {
        excludedDates.add(Utilities.formatDate(cur, 'Asia/Tokyo', 'yyyy/MM/dd'));
        cur.setDate(cur.getDate() + 1);
      }
    });
  }

  // ── 実績データの既存日付セットと最終日付を取得 ───────────────────
  // ★変更点：readFromDate の代わりに existingDates で重複チェックする
  //   → 順番通りでなくても、まだ取込まれていない日付を全て読める
  const existingDates = new Set();
  let latestDate = null;
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().forEach(r => {
      if (!r[0]) return;
      const d = r[0] instanceof Date ? new Date(r[0]) : new Date(r[0]);
      d.setHours(0, 0, 0, 0);
      const ds = Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy/MM/dd');
      existingDates.add(ds);
      if (!latestDate || d > latestDate) latestDate = d;
    });
  }

  // ── 直近14日の連続性チェック（警告のみ・止まらない） ─────────────
  // ★変更点：ギャップを検出しても return せず、通知だけして続行する
  //   → 後から届いたデータも自動で取込めるようにするため
  if (latestDate) {
    const checkStart = new Date(latestDate);
    checkStart.setDate(latestDate.getDate() - 13);
    checkStart.setHours(0, 0, 0, 0);
    const yesterday = new Date(); yesterday.setHours(0, 0, 0, 0);
    yesterday.setDate(yesterday.getDate() - 1);

    const missingDates = [];
    const cur = new Date(checkStart);
    while (cur <= latestDate && cur <= yesterday) {
      const ds = Utilities.formatDate(cur, 'Asia/Tokyo', 'yyyy/MM/dd');
      if (!existingDates.has(ds) && !excludedDates.has(ds)) missingDates.push(ds);
      cur.setDate(cur.getDate() + 1);
    }

    if (missingDates.length > 0) {
      // 警告ログ
      console.log('【警告】以下の日付のデータが不足しています: ' + missingDates.join(', '));

      // メール通知（止まらずに続行）
      try {
        MailApp.sendEmail(
          Session.getActiveUser().getEmail(),
          '【発注システム】売上データに欠落日があります（取込は継続中）',
          '以下の日付のデータが見つかりません。\n' +
          '取込処理は継続していますが、発注精度に影響する可能性があります。\n\n' +
          missingDates.map(d => '・' + d).join('\n') +
          '\n\n【データがフォルダにある場合】\n' +
          '  → 次回の自動取込で自動的に取込まれます。\n' +
          '     手動で確認したい場合は importFooditCSV() を実行してください。\n\n' +
          '【停電・休業など売上ゼロの日の場合】\n' +
          '  → イベント登録シートに「除外」区分で登録してください。\n' +
          '     以降この警告は表示されなくなります。'
        );
      } catch(e) {
        console.log('メール送信失敗: ' + e.message);
      }
      // ★ return しない → そのまま取込を続ける
    }
  }

  // ── CSVファイルを読み込む ──────────────────────────────────────────
  // ★変更点：existingDates にない日付のCSVを全て読む（順番不問）
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const files  = folder.getFiles();
  let dailyStats  = {};
  let menuDetails = [];

  while (files.hasNext()) {
    const file = files.next();
    const name = file.getName();
    if (name.length < 32) continue;

    const typeFlag = name.charAt(16);
    const fileDate = name.substring(24, 32);
    const formattedDate =
      fileDate.substring(0, 4) + '/' +
      fileDate.substring(4, 6) + '/' +
      fileDate.substring(6, 8);

    // ★変更点：既に取込済みの日付はスキップ（latestDate比較をやめた）
    if (existingDates.has(formattedDate)) continue;

    let csvText;
    if      (typeFlag === '1') csvText = file.getBlob().getDataAsString('UTF-8');
    else if (typeFlag === '3') csvText = file.getBlob().getDataAsString('Shift_JIS');
    else continue;

    const csv = Utilities.parseCsv(csvText);

    if (typeFlag === '3') {
      let sales = 0, customers = 0;
      for (let i = 0; i < csv.length; i++) {
        const label = csv[i][0] ? csv[i][0].toString() : '';
        if (sales === 0 && (label.includes('売上') || label.includes('実績'))) {
          if (csv[i][1]) sales = Number(csv[i][1].toString().replace(/[^0-9-]/g, ''));
        }
        if (customers === 0 && label.includes('客数')) {
          if (csv[i][1]) customers = Number(csv[i][1].toString().replace(/[^0-9-]/g, ''));
        }
      }
      if (sales === 0     && csv[64] && csv[64][1]) sales     = Number(csv[64][1].toString().replace(/[^0-9-]/g, ''));
      if (customers === 0 && csv[55] && csv[55][1]) customers = Number(csv[55][1].toString().replace(/[^0-9-]/g, ''));
      dailyStats[formattedDate] = { sales, customers };
    }

    if (typeFlag === '1') {
      for (let i = 11; i < csv.length; i++) {
        const code     = csv[i][1];
        const nameItem = csv[i][2];
        const qty      = Number(csv[i][9]);
        if (!code || qty <= 0 || nameItem === '合計') continue;
        menuDetails.push({ date: formattedDate, code, name: nameItem, qty });
      }
    }
  }

  if (menuDetails.length === 0 && Object.keys(dailyStats).length === 0) {
    console.log('新規データなし');
    return;
  }

  // ── 実績データシートに追記 ─────────────────────────────────────────
  if (menuDetails.length > 0) {
    // ★既存行の売上0を修正する処理を追加
  if (sheet.getLastRow() > 1) {
    const existingData = sheet.getRange(2, 1, sheet.getLastRow()-1, 4).getValues();
    existingData.forEach((row, i) => {
      const ds = row[0] instanceof Date
        ? Utilities.formatDate(row[0], 'Asia/Tokyo', 'yyyy/MM/dd')
        : row[0].toString();
      const stats = dailyStats[ds];
      if (stats && stats.sales > 0 && Number(row[2]) === 0) {
        sheet.getRange(i+2, 3, 1, 2).setValues([[stats.sales, stats.customers]]);
        console.log('実績データ売上0を修正: ' + ds + ' → ' + stats.sales);
      }
    });
  }

  const rows = menuDetails.map(m => {
    const stats = dailyStats[m.date] || { sales: 0, customers: 0 };
    return [m.date, getDayOfWeek_(m.date), stats.sales, stats.customers, m.code, m.name, m.qty];
  });
    rows.sort((a, b) => new Date(a[0]) - new Date(b[0]));
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['日付', '曜日', '総売上', '総客数', 'メニューCD', 'メニュー名', '出食数']);
    }
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 7).setValues(rows);
    sheet.getRange(2, 3, sheet.getLastRow() - 1, 2).setNumberFormat('#,##0');
  }

  writeDailySalesSheet_(ss, dailyStats);

  const cpSubtractMap = buildIngredientDailyUsage();
  saveCpSubtractMap_(cpSubtractMap);

  try {
    buildOutlierLog();
    console.log('異常値ログを自動更新しました');
  } catch (e) {
    console.log('異常値ログ更新をスキップ: ' + e.message);
  }

  // 商品マスタ未登録メニューチェック
  try {
    checkUnregisteredMenus_();
  } catch(e) {
    console.log('未登録メニューチェックをスキップ: ' + e.message);
  }
　  // ★売上0の自動再チェック処理
  recheckZeroSales_(sheet, excludedDates, dailyStats);
  console.log('取込完了');
}
function recheckZeroSales_(jissekiSheet, excludedDates, newDailyStats) {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const props   = PropertiesService.getScriptProperties();
  const tz      = 'Asia/Tokyo';
  const today   = new Date(); today.setHours(0,0,0,0);

  // ── 既存の要確認リストを読み込む ──────────────────────────
  let watchList = {};  // { 'yyyy/MM/dd': { firstDetected: ISO文字列, notified: bool } }
  try {
    const saved = props.getProperty('zeroSalesWatchList');
    if (saved) watchList = JSON.parse(saved);
  } catch(e) { watchList = {}; }

  // ── 実績データから売上0の日付を検出（過去30日以内） ────────
  const limitDate = new Date(today); limitDate.setDate(today.getDate() - 30);
  if (jissekiSheet && jissekiSheet.getLastRow() > 1) {
    const seenDates = new Set();
    jissekiSheet.getRange(2, 1, jissekiSheet.getLastRow()-1, 4)
      .getValues()
      .forEach(row => {
        if (!row[0]) return;
        const d = row[0] instanceof Date ? row[0] : new Date(row[0]);
        d.setHours(0,0,0,0);
        if (d >= today || d < limitDate) return;
        const ds = Utilities.formatDate(d, tz, 'yyyy/MM/dd');
        if (seenDates.has(ds)) return;
        seenDates.add(ds);
        if (Number(row[2]) === 0 && !excludedDates.has(ds)) {
          // 要確認リストに追加（まだなければ）
          if (!watchList[ds]) {
            watchList[ds] = {
              firstDetected: new Date().toISOString(),
              notified: false
            };
            console.log('売上0を検出、監視リストに追加: ' + ds);
          }
        }
      });
  }

  // ── 今回取り込んだCSVで売上が入ったものがあれば上書き ────
  const fixedDates = [];
  Object.keys(newDailyStats).forEach(dateStr => {
    const stats = newDailyStats[dateStr];
    if (!watchList[dateStr]) return;
    if (stats.sales <= 0) return;

    // 実績データの該当日を上書き
    if (jissekiSheet && jissekiSheet.getLastRow() > 1) {
      const data = jissekiSheet.getRange(2, 1, jissekiSheet.getLastRow()-1, 4).getValues();
      data.forEach((row, i) => {
        if (!row[0]) return;
        const ds = row[0] instanceof Date
          ? Utilities.formatDate(row[0], tz, 'yyyy/MM/dd')
          : row[0].toString();
        if (ds !== dateStr) return;
        if (Number(row[2]) === 0) {
          jissekiSheet.getRange(i+2, 3, 1, 2).setValues([[stats.sales, stats.customers]]);
        }
      });
    }

    // 日次売上実績シートも上書き
    const salesSheet = ss.getSheetByName('日次売上実績');
    if (salesSheet && salesSheet.getLastRow() > 1) {
      const sData = salesSheet.getRange(2, 1, salesSheet.getLastRow()-1, 5).getValues();
      sData.forEach((row, i) => {
        if (!row[0]) return;
        const ds = row[0] instanceof Date
          ? Utilities.formatDate(row[0], tz, 'yyyy/MM/dd')
          : row[0].toString();
        if (ds !== dateStr) return;
        if (Number(row[3]) === 0) {
          salesSheet.getRange(i+2, 4, 1, 2).setValues([[stats.sales, stats.customers]]);
        }
      });
    }

    fixedDates.push(dateStr + '（売上: ' + stats.sales.toLocaleString() + '円）');
    delete watchList[dateStr];
    console.log('売上0を自動修正: ' + dateStr + ' → ' + stats.sales);
  });

  // ── 除外登録済みの日付は監視リストから削除 ────────────────
  Object.keys(watchList).forEach(ds => {
    if (excludedDates.has(ds)) {
      delete watchList[ds];
      console.log('除外登録済みのため監視リストから削除: ' + ds);
    }
  });

  // ── 3日以上経っても0のままの日付をメール通知 ──────────────
  const alertDates = [];
  Object.keys(watchList).forEach(ds => {
    const entry     = watchList[ds];
    const detected  = new Date(entry.firstDetected);
    const daysPassed = Math.floor((today - detected) / 86400000);
    if (daysPassed >= 3 && !entry.notified) {
      alertDates.push(ds + '（' + daysPassed + '日間売上0）');
      watchList[ds].notified = true;
    }
  });

  // ── メール送信 ────────────────────────────────────────────
  try {
    const email = Session.getActiveUser().getEmail();

    if (fixedDates.length > 0) {
      MailApp.sendEmail(
        email,
        '【発注システム】売上0を自動修正しました',
        '以下の日付の売上データが自動的に修正されました。\n\n' +
        fixedDates.map(d => '・' + d).join('\n') +
        '\n\n売上予測を更新するには buildSalesForecast() を実行してください。'
      );
      console.log('自動修正メールを送信: ' + fixedDates.join(', '));
    }

    if (alertDates.length > 0) {
      MailApp.sendEmail(
        email,
        '【発注システム】売上0が3日以上続いています（要確認）',
        '以下の日付の売上が3日以上0のままです。\n\n' +
        alertDates.map(d => '・' + d).join('\n') +
        '\n\n【対処方法】\n' +
        '① fooditから正しいCSVを再ダウンロードしてフォルダに上書き\n' +
        '   → 翌朝の自動実行で自動修正されます\n\n' +
        '② 停電・休業など売上ゼロが正当な場合\n' +
        '   → イベント登録シートに「除外」区分で登録してください\n\n' +
        '③ 手動で直接修正する場合\n' +
        '   → 日次売上実績シートの該当行を直接編集後\n' +
        '      buildSalesForecast() を実行してください'
      );
      console.log('売上0アラートメールを送信: ' + alertDates.join(', '));
    }
  } catch(e) {
    console.log('メール送信失敗: ' + e.message);
  }

  // ── 監視リストを保存 ─────────────────────────────────────
  try {
    props.setProperty('zeroSalesWatchList', JSON.stringify(watchList));
  } catch(e) {
    console.log('監視リスト保存失敗: ' + e.message);
  }

  if (Object.keys(watchList).length > 0) {
    console.log('売上0監視中: ' + Object.keys(watchList).join(', '));
  }
}
/**
 * 商品マスタに登録されていないメニューが実績データにあれば
 * ログ出力＋メール通知する。importFooditCSV() から自動で呼ばれる。
 *
 * 【誤通知を防ぐ仕組み（2026-07-23 全面改修）】
 *  1回の読み取りで即メールを送ると、取込中の一瞬のブレや、コードを登録する
 *  直前のタイミングで走ったときに「登録済みなのに未登録」と誤通知してしまう。
 *  そこで監視リスト（ScriptProperties）に貯め、
 *   ・2日連続で未登録だったコードだけをメールする（1日限りの誤検知を除外）
 *   ・商品マスタに登録された瞬間、監視リストから自動で消える（自己解決）
 *   ・同じ商品名で登録済みのコードがあれば「コード切替の可能性」として明記する
 */
function checkUnregisteredMenus_() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const normName_ = s => String(s === null || s === undefined ? '' : s).replace(/[\s　]/g, '').toUpperCase();

  // ── 商品マスタの登録済みコードと「商品名→コード」対応を取得 ───────────
  const menuSheet = ss.getSheetByName('商品マスタ');
  const registeredCodes = new Set();
  const regNameToCode = {};   // 正規化した商品名 → 登録済みコード（コード切替のヒント用）
  if (menuSheet && menuSheet.getLastRow() > 1) {
    menuSheet.getRange(2, 1, menuSheet.getLastRow() - 1, 2).getValues().forEach(r => {
      const c = normalizeCode_(r[0]);
      if (!c) return;
      registeredCodes.add(c);
      const nk = normName_(r[1]);
      if (nk && !regNameToCode[nk]) regNameToCode[nk] = c;
    });
  }

  // 安全装置：商品マスタが読めていない（極端に少ない）ときは誤通知を防ぐため中止
  if (registeredCodes.size < 50) {
    console.log('商品マスタの読み込みが不十分（コード数' + registeredCodes.size + '）のため未登録チェックを中止');
    return;
  }

  // ── 直近3日の実績データから未登録メニューを集計 ───────────────────────
  const jissekiSheet = ss.getSheetByName('実績データ');
  if (!jissekiSheet || jissekiSheet.getLastRow() < 2) return;

  const today = new Date(); today.setHours(0, 0, 0, 0);
  const limit = new Date(today); limit.setDate(today.getDate() - 3);
  const todayStr = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy/MM/dd');

  const candidates = {};  // code -> { name, qty, lastDate }
  jissekiSheet.getRange(2, 1, jissekiSheet.getLastRow() - 1, 7).getValues().forEach(row => {
    if (!row[0] || !row[4]) return;
    const d = row[0] instanceof Date ? row[0] : new Date(row[0]);
    d.setHours(0, 0, 0, 0);
    if (d < limit) return;  // 直近3日以内のみ
    const code = normalizeCode_(row[4]);
    if (!code || registeredCodes.has(code)) return;
    const ds = Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy/MM/dd');
    if (!candidates[code]) candidates[code] = { name: (row[5] || '名称不明').toString(), qty: 0, lastDate: ds };
    candidates[code].qty += Number(row[6]) || 0;
    if (ds > candidates[code].lastDate) candidates[code].lastDate = ds;
  });

  // ── 監視リストを読み込み・更新 ───────────────────────────────────────
  const props = PropertiesService.getScriptProperties();
  let watch = {};
  try { watch = JSON.parse(props.getProperty('unregisteredMenuWatch') || '{}'); } catch (e) { watch = {}; }

  // 登録された／しばらく売れていないコードを監視リストから除去（自己解決）
  Object.keys(watch).forEach(code => {
    if (registeredCodes.has(code)) { delete watch[code]; return; }          // 登録された → 解決
    const last = watch[code].lastStr || '2000/01/01';
    const lastMs = new Date(last.replace(/\//g, '/')).getTime();
    if (!isNaN(lastMs) && (today.getTime() - lastMs) > 7 * 86400000) delete watch[code]; // 7日売れず → 監視終了
  });

  // 今回の未登録コードを監視リストに反映（1日1回だけ日数カウントを進める）
  Object.keys(candidates).forEach(code => {
    const c = candidates[code];
    if (!watch[code]) {
      watch[code] = { name: c.name, qty: c.qty, firstStr: todayStr, lastStr: todayStr, lastCountedDay: todayStr, dayCount: 1, notified: false };
    } else {
      const w = watch[code];
      w.name = c.name; w.qty = c.qty; w.lastStr = todayStr;
      if (w.lastCountedDay !== todayStr) { w.dayCount = (w.dayCount || 1) + 1; w.lastCountedDay = todayStr; }
    }
  });

  // ── 2日以上連続で未登録＝本物、かつ未通知のものだけをメール ─────────────
  const toNotify = Object.keys(watch)
    .filter(code => (watch[code].dayCount || 0) >= 2 && !watch[code].notified)
    .sort((a, b) => (watch[b].qty || 0) - (watch[a].qty || 0));

  if (toNotify.length > 0) {
    const lines = toNotify.map(code => {
      const w = watch[code];
      const hintCode = regNameToCode[normName_(w.name)];
      const hint = (hintCode && hintCode !== code)
        ? '\n      ↑同名の登録済み商品あり（コード ' + hintCode + '）。本部がコードを切り替えた可能性。'
        : '';
      return '  コード:' + code + '　' + w.name +
             '（' + Math.round(w.qty || 0) + '食 / 最終 ' + w.lastStr + '）' + hint;
    });

    console.log('【警告】商品マスタ未登録メニュー（2日以上連続）\n' + lines.join('\n'));

    try {
      const email   = Session.getActiveUser().getEmail();
      const nowStr  = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
      const subject = '【発注システム】商品マスタ未登録のメニューが売れています';
      const body =
        '検出日時: ' + nowStr + '\n\n' +
        '以下のメニューが2日以上続けて実績に上がっていますが、商品マスタに登録されていません。\n' +
        '発注計算に含まれないため、食材が不足する可能性があります。\n\n' +
        lines.join('\n') +
        '\n\n【対処方法】\n' +
        '・普通の新メニューの場合 → 商品マスタにレシピを登録\n' +
        '・「同名の登録済み商品あり」と出ている場合 → 本部がコードを切り替えた可能性。\n' +
        '　でかもりなど別商品かもしれないので、旧コードのレシピを流用せず中身を確認してから登録。\n' +
        '登録後は 発注計算バッファ_完全生成() を実行してください。\n\n' +
        '※このメールは同じコードにつき1回だけ送られます（登録すると自動で監視解除されます）。';
      MailApp.sendEmail(email, subject, body);
      toNotify.forEach(code => { watch[code].notified = true; });  // 二重送信を防ぐ
      console.log('未登録メニュー通知メールを送信しました（' + toNotify.length + '件）');
    } catch (e) {
      console.log('メール送信失敗: ' + e.message);
    }
  }

  // ── 監視リストを保存 ────────────────────────────────────────────────
  try { props.setProperty('unregisteredMenuWatch', JSON.stringify(watch)); }
  catch (e) { console.log('監視リスト保存失敗: ' + e.message); }
}

// ─────────────────────────────────────────────────────────────────────
// 日次売上実績シートに書き込む
// ─────────────────────────────────────────────────────────────────────

// ─────────────────────────────────────────────────────────────────────
// 暫定日イベントの日付確認リマインド
//   売上予測(buildSalesForecast)実行時に自動チェック。開始日が21日以内に迫った
//   暫定日イベントがあれば「正確な日付を確認して」とメール通知（同じ日付は1回だけ）。
//   force=true でメニューから手動実行（通知済みでも再送）。
// ─────────────────────────────────────────────────────────────────────
function checkProvisionalEvents_(force) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const evSheet = ss.getSheetByName('イベント登録');
  if (!evSheet || evSheet.getLastRow() < 2) { if (force) SpreadsheetApp.getUi().alert('イベント登録シートにデータがありません'); return; }

  const DOW = ['日','月','火','水','木','金','土'];
  const today = new Date(); today.setHours(0,0,0,0);
  const horizon = new Date(today); horizon.setDate(today.getDate() + 21); // 21日先まで見る

  const evCols = Math.max(10, evSheet.getLastColumn());
  const rows = evSheet.getRange(2, 1, evSheet.getLastRow()-1, evCols).getValues();
  const props = PropertiesService.getScriptProperties();
  const due = [];
  rows.forEach(row => {
    const name = (row[0]||'').toString().trim(); if (!name) return;
    const note = (row[9]||'').toString();
    const repeatPat = (evCols >= 11 ? (row[10]||'') : '').toString().trim();
    if (note.indexOf('暫定') < 0 && !/暫定/.test(repeatPat)) return; // 暫定日イベントのみ対象
    // 開始日：明示日付があればそれ、無ければ「繰り返し」列を展開して直近の該当日を使う
    let start = toDate_(row[1]);
    if (!start && repeatPat) {
      const occ = expandEventRepeat_(repeatPat, today, horizon);
      const future = occ.map(function(o){ return o.start; }).filter(function(d){ return d >= today; }).sort(function(a,b){ return a - b; });
      start = future.length ? future[0] : null;
    }
    if (!start) return;
    start.setHours(0,0,0,0);
    if (start < today || start > horizon) return;  // 21日以内に迫ったものだけ
    const key = 'provNotified_' + name + '_' + Utilities.formatDate(start, 'Asia/Tokyo', 'yyyyMMdd');
    if (!force && props.getProperty(key)) return;  // 既に通知済み（自動時のみスキップ）
    due.push({ name: name, start: start, key: key });
  });

  if (due.length === 0) { if (force) SpreadsheetApp.getUi().alert('21日以内に迫った暫定日イベントはありません'); return; }

  const lines = due.map(function(e){
    return '  ' + e.name + '　→ 現在の暫定日：' + Utilities.formatDate(e.start, 'Asia/Tokyo', 'M月d日') + '(' + DOW[e.start.getDay()] + ')';
  });
  try {
    const email = Session.getActiveUser().getEmail();
    const subject = '【発注システム】暫定日イベントの日付確認をお願いします';
    const body =
      '売上予測に使っている「暫定日イベント」が近づいています。\n' +
      '今年の正確な日付をご確認のうえ、ずれていれば「イベント登録」シートで\n' +
      '日付を修正してください（繰り返し設定の行は「繰り返し」列の日付を、\n' +
      '日付を直接入れた行は開始日・終了日を直します）。\n\n' +
      lines.join('\n') +
      '\n\n日付が正しければ、そのままで問題ありません。';
    MailApp.sendEmail(email, subject, body);
    due.forEach(function(e){ props.setProperty(e.key, '1'); }); // 通知済みフラグを保存（同じ日付の再送を防ぐ）
    console.log('暫定日イベント確認リマインドを送信: ' + due.length + '件');
    if (force) SpreadsheetApp.getUi().alert('確認メールを送信しました（' + due.length + '件）');
  } catch(e) {
    console.log('暫定日リマインド送信失敗: ' + e.message);
    if (force) SpreadsheetApp.getUi().alert('メール送信に失敗しました: ' + e.message);
  }
}

function remindProvisionalEventsNow_() { checkProvisionalEvents_(true); }

// ─────────────────────────────────────────────────────────────────────
// 日次売上実績シートに書き込む
// ─────────────────────────────────────────────────────────────────────

function writeDailySalesSheet_(ss, dailyStats) {
  const SHEET_NAME = '日次売上実績';
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['日付', '曜日コード', '曜日名', '総売上', '総客数']);
    sheet.setFrozenRows(1);
  }
  const existingDates = new Set();
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().forEach(r => {
      if (r[0]) {
        const ds = r[0] instanceof Date
          ? Utilities.formatDate(r[0], 'Asia/Tokyo', 'yyyy/MM/dd') : r[0].toString();
        existingDates.add(ds);
      }
    });
  }
  const newRows = [];
  Object.keys(dailyStats).sort().forEach(dateStr => {
    // ★売上0で登録済みの場合は上書きする
    if (existingDates.has(dateStr)) {
      const stats = dailyStats[dateStr];
      if (stats.sales > 0) {
        // 既存行を探して上書き
        const existingData = sheet.getRange(2, 1, sheet.getLastRow()-1, 5).getValues();
        existingData.forEach((row, i) => {
          const ds = row[0] instanceof Date
            ? Utilities.formatDate(row[0], 'Asia/Tokyo', 'yyyy/MM/dd')
            : row[0].toString();
          if (ds === dateStr && Number(row[3]) === 0) {
            sheet.getRange(i+2, 4, 1, 2).setValues([[stats.sales, stats.customers]]);
            console.log('売上0を修正: ' + dateStr + ' → ' + stats.sales);
          }
        });
      }
      return;
    }
    const stats   = dailyStats[dateStr];
    const d       = new Date(dateStr);
    const dow     = d.getDay() + 1;
    const dowName = getDayOfWeek_(dateStr);
    newRows.push([dateStr, dow, dowName, stats.sales, stats.customers]);
  });
  if (newRows.length === 0) return;
  sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, 5).setValues(newRows);
  sheet.getRange(2, 4, sheet.getLastRow() - 1, 2).setNumberFormat('#,##0');
}

// ─────────────────────────────────────────────────────────────────────
// 食材日時実績シートを再計算する
// ─────────────────────────────────────────────────────────────────────

function buildIngredientDailyUsage() {

  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet  = ss.getSheetByName('実績データ');
  const menuSheet = ss.getSheetByName('商品マスタ');
  const ingSheet  = ss.getSheetByName('食材マスタ');
  const outSheet  = ss.getSheetByName('食材日時実績') || ss.insertSheet('食材日時実績');

  // 食材マスタ読込
  const ingMap = {};
  ingSheet.getDataRange().getValues().forEach((row, i) => {
    if (i === 0 || !row[0]) return;
    ingMap[row[0]] = { 
      name: row[1] || '',
      spec:  Number(row[5]) > 0 ? Number(row[5]) : 1,  // F列 仕入規格量
      yield: Number(row[7]) > 0 ? Number(row[7]) : 1   // H列 歩留率
  };
  });

  // 食材コード切替履歴
  const ingCodeMap = {};
  const switchSheet2 = ss.getSheetByName('食材コード切替履歴');
  if (switchSheet2 && switchSheet2.getLastRow() > 1) {
    switchSheet2.getRange(2, 1, switchSheet2.getLastRow() - 1, 3).getValues().forEach(row => {
      const oldCd = Number(row[0]); const newCd = Number(row[1]);
      if (oldCd && newCd) ingCodeMap[oldCd] = newCd;
    });
  }

  // 商品マスタ読込
  const menuMap = {};
  menuSheet.getDataRange().getValues().forEach((row, i) => {
    if (i === 0 || !row[0] || !row[4]) return;
    const qty = Number(row[12]); if (qty === 0) return;
    const ingCd = ingCodeMap[row[4]] || row[4];
    if (!menuMap[row[0]]) menuMap[row[0]] = [];
    menuMap[row[0]].push({ ingCd, qty });
  });

  // CP設定シート読込（日付指定CP と 繰り返しCP を区別して収集）
  const cpList = [], endedProductMap = {}, repeatCpList = [];
  const cpSettingSheet = ss.getSheetByName('CP設定');
  if (cpSettingSheet && cpSettingSheet.getLastRow() > 1) {
    cpSettingSheet.getRange(2, 1, cpSettingSheet.getLastRow() - 1, 9).getValues().forEach(row => {
      const cpName = (row[0] || '').toString().trim(); if (!cpName) return;
      if ((row[8] || '').toString().trim() === '休止') return; // CP休止行は別処理（buildCpPauseSubMap_）。通常CPとして扱わない
      const note = (row[7] || '').toString().trim();
      if (note === '除外') return; // 発注除外フラグ行はスキップ
      let codes = [];
      if (typeof row[3] === 'number') { codes = [Math.round(row[3])]; }
      else { codes = String(row[3] || '').split(',').flatMap(c => { c = c.trim(); const m = c.match(/^(\d+)-(\d+)$/); if (m) { const arr = []; for (let n = Number(m[1]); n <= Number(m[2]); n++) arr.push(n); return arr; } const n = Number(c); return n > 0 ? [n] : []; }); }
      if (codes.length === 0) return;
      const setServingRate = Number(row[4]) || 0.10; // E列: 参考値（実績なし時のフォールバック）
      const repeat = (row[6] || '').toString().trim();
      const startDate = row[1] instanceof Date ? row[1] : (row[1] ? new Date(String(row[1]).replace(/^(\d{4})(\d{2})(\d{2})$/, '$1/$2/$3')) : null);
      const endDate   = row[2] instanceof Date ? row[2] : (row[2] ? new Date(String(row[2]).replace(/^(\d{4})(\d{2})(\d{2})$/, '$1/$2/$3')) : null);
      const hasValidDates = startDate && endDate && !isNaN(startDate.getTime()) && !isNaN(endDate.getTime());
      if (!hasValidDates && repeat && /毎年\d+月\d+日〜次の日曜/.test(repeat)) {
        // 毎年X月Y日〜次の日曜 → 昨年・今年・来年の期間を計算してcpListへ直接追加（実績スキャン対象）
        const mAS2 = repeat.match(/毎年(\d+)月(\d+)日〜次の日曜/);
        const tMonth2 = Number(mAS2[1]) - 1, tDay2 = Number(mAS2[2]);
        const thisYear2 = new Date().getFullYear();
        const cpFlag2 = (row[8] || '').toString().trim();
        const isNc2 = cpFlag2 !== '' && cpFlag2 !== '切替', isSw2 = cpFlag2 === '切替';
        const cutoffAS2 = new Date(); cutoffAS2.setFullYear(cutoffAS2.getFullYear() - 1); cutoffAS2.setHours(0,0,0,0);
        [thisYear2 - 1, thisYear2, thisYear2 + 1].forEach(yr => {
          const sd2 = new Date(yr, tMonth2, tDay2); sd2.setHours(0,0,0,0);
          const ed2 = new Date(sd2);
          while (ed2.getDay() !== 0) ed2.setDate(ed2.getDate() + 1);
          ed2.setHours(0,0,0,0);
          if (ed2 < cutoffAS2) return; // 1年以上前はスキップ
          cpList.push({ name: cpName, codes, startDate: sd2, endDate: ed2, rate: setServingRate, setServingRate, isNormalCp: isNc2, isSwitchCp: isSw2 });
        });
        return;
      }
      if (!hasValidDates && repeat && (/毎月\d+/.test(repeat) || /毎年\d+月\d+/.test(repeat))) {
        // 日付なし＋繰り返しパターン → repeatCpListへ（出食率は実績から自動計算）
        repeatCpList.push({ name: cpName, codes, setServingRate, rate: setServingRate, repeat });
        return;
      }
      if (!hasValidDates) return;
      startDate.setHours(0,0,0,0); endDate.setHours(0,0,0,0);
      const cpCutoff_ = new Date(); cpCutoff_.setDate(cpCutoff_.getDate() - 28); cpCutoff_.setHours(0,0,0,0);
      if (endDate < cpCutoff_) return; // 28日以上前に終了したCPはスキップ
      const rate = setServingRate; // 実績スキャン後に更新
      const cpFlag = (row[8] || '').toString().trim();
      const isNormalCp = cpFlag !== '' && cpFlag !== '切替';
      const isSwitchCp = cpFlag === '切替';
      if (isNormalCp) {
        cpList.push({ name: cpName, codes, startDate, endDate, rate, setServingRate, isNormalCp: true, isSwitchCp: false });
      } else {
        if (!isSwitchCp) {
          codes.forEach(cd => {
            if (!endedProductMap[cd] || endDate > endedProductMap[cd]) {
              endedProductMap[cd] = endDate;
            }
          });
        }
        cpList.push({ name: cpName, codes, startDate, endDate, rate, setServingRate, isNormalCp: false, isSwitchCp });
      }
    });
  }

  // 実績データを全期間スキャンして出食率を自動計算
  // 実績2日以上ある場合は実績値を採用、なければ設定値（E列）をフォールバック
  {
    const allCpItems = [...cpList, ...repeatCpList];
    const cpDayData = {};
    allCpItems.forEach(cp => { cpDayData[cp.name] = { totalDishes: 0, cpDishes: 0, days: new Set() }; });
    logSheet.getDataRange().getValues().forEach((row, idx) => {
      if (idx === 0 || !row[0]) return;
      const d = row[0] instanceof Date ? row[0] : new Date(row[0]); d.setHours(0,0,0,0);
      const ds = Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy/MM/dd');
      const menuCd = Number(row[4]);
      const count  = Number(row[6]) || 0;
      if (!menuCd || count === 0) return;
      allCpItems.forEach(cp => {
        let hit = false;
        if (cp.startDate && cp.endDate) {
          hit = d >= cp.startDate && d <= cp.endDate;
        } else if (cp.repeat) {
          const mAnn2 = cp.repeat.match(/毎年(\d+)月(\d+)/);
          if (mAnn2) { hit = d.getMonth() === Number(mAnn2[1]) - 1 && d.getDate() === Number(mAnn2[2]); }
          else { const m2 = cp.repeat.match(/毎月(\d+)/); hit = m2 ? d.getDate() === Number(m2[1]) : false; }
        }
        if (!hit) return;
        const dd = cpDayData[cp.name];
        dd.totalDishes += count;
        dd.days.add(ds);
        if (cp.codes.includes(menuCd)) dd.cpDishes += count;
      });
    });
    allCpItems.forEach(cp => {
      const dd = cpDayData[cp.name];
      if (dd.days.size >= 2 && dd.totalDishes > 0) {
        cp.rate = Math.round(dd.cpDishes / dd.totalDishes * 100) / 100;
        console.log('CP [' + cp.name + '] 出食率（実績）: ' + cp.rate + ' (' + dd.days.size + '日, 設定値:' + cp.setServingRate + ')');
      } else {
        cp.rate = cp.setServingRate;
        if (dd.days.size > 0) console.log('CP [' + cp.name + '] 出食率（設定値）: ' + cp.rate + ' ※実績' + dd.days.size + '日のみ');
      }
    });
  }

  // 繰り返しCPを直近5日の対象日に展開して cpList へ追加（食材発注計算に組み込む）
  {
    const todayR = new Date(); todayR.setHours(0,0,0,0);
    const fEndR  = new Date(todayR); fEndR.setDate(todayR.getDate() + 5);
    repeatCpList.forEach(rcp => {
      const mAnn3 = rcp.repeat.match(/毎年(\d+)月(\d+)/);
      const m3    = !mAnn3 && rcp.repeat.match(/毎月(\d+)/);
      if (!mAnn3 && !m3) return;
      const cur3 = new Date(todayR);
      while (cur3 <= fEndR) {
        const matchDate = mAnn3
          ? (cur3.getMonth() === Number(mAnn3[1]) - 1 && cur3.getDate() === Number(mAnn3[2]))
          : cur3.getDate() === Number(m3[1]);
        if (matchDate) {
          const sd = new Date(cur3), ed = new Date(cur3);
          sd.setHours(0,0,0,0); ed.setHours(0,0,0,0);
          cpList.push({ name: rcp.name, codes: rcp.codes, startDate: sd, endDate: ed, rate: rcp.rate, setServingRate: rcp.setServingRate, isNormalCp: true, isSwitchCp: false });
        }
        cur3.setDate(cur3.getDate() + 1);
      }
    });
  }

  // 異常値ログ＋（真の）祝日を平均計算から除外
  //   祝日は普段の曜日平均を膨らませる（検証：平日の祝日を除外で通常日WAPE 31.6→30.8%改善）。
  //   真の祝日・連休のみ除外し、通常の土日は実績として使う（getHolidayTaggedDates_(false)）。
  const outlierExcludedDates = new Set();
  try {
    getOutlierExcludedDates_().forEach(d => outlierExcludedDates.add(d));
    getHolidayTaggedDates_(false).forEach(d => outlierExcludedDates.add(d));
    if (outlierExcludedDates.size > 0) {
      console.log('平均計算除外日: ' + outlierExcludedDates.size + '日（異常値＋祝日）');
    }
  } catch (e) {
    console.log('除外日の取得をスキップ: ' + e.message);
  }

  // 実績データ集計（直近56日＝8週。バックテストで4週より安定かつ正確）
  const today     = new Date(); today.setHours(0, 0, 0, 0);
  const limitDate = new Date(today); limitDate.setDate(today.getDate() - 56);
  const resultMap = {}, cpActualDaysMap = {}, cpSubtractMap = {};
  cpList.forEach(cp => { cpActualDaysMap[cp.name] = new Set(); });
  const cpActualIngKeys = new Set(), jissekiSeen = new Set();

  // パス1：各食材の実績日を事前収集（CP開始前30日以内の実績有無で「通常食材 vs CP専用新食材」を自動判定）
  const PRE_CP_JISSEKI_DAYS = 30;
  const ingJissekiDates = {};
  logSheet.getDataRange().getValues().forEach((row, i) => {
    if (i === 0) return;
    const date = row[0], menuCd = row[4], count = Number(row[6]);
    if (!date || !menuMap[menuCd] || count === 0) return;
    const dateObj = date instanceof Date ? date : new Date(date);
    dateObj.setHours(0, 0, 0, 0);
    if (dateObj < limitDate) return;
    menuMap[menuCd].forEach(r => {
      if (!ingJissekiDates[r.ingCd]) ingJissekiDates[r.ingCd] = [];
      ingJissekiDates[r.ingCd].push(dateObj.getTime());
    });
  });

  logSheet.getDataRange().getValues().forEach((row, i) => {
    if (i === 0) return;
    const date = row[0], menuCd = row[4], count = Number(row[6]);
    if (!date || !menuMap[menuCd] || count === 0) return;
    const dateObj = date instanceof Date ? date : new Date(date);
    dateObj.setHours(0, 0, 0, 0);
    if (dateObj < limitDate) return;
    const dateStr = Utilities.formatDate(dateObj, 'Asia/Tokyo', 'yyyy/MM/dd');
    const jissekiKey = dateStr + '_' + menuCd;
    if (jissekiSeen.has(jissekiKey)) return;
    jissekiSeen.add(jissekiKey);
    if (outlierExcludedDates.has(dateStr)) return;
    if (endedProductMap[menuCd] && dateObj > endedProductMap[menuCd]) return;
    const isInNormalCpPeriod = cpList.some(cp =>
      cp.isNormalCp && cp.codes.includes(menuCd) && dateObj >= cp.startDate && dateObj <= cp.endDate);
    if (isInNormalCpPeriod) return;
    const dow = dateObj.getDay() + 1;
    menuMap[menuCd].forEach(r => {
      const ingInfo = ingMap[r.ingCd] || { name: '', yield: 1 };
      const usedQty = (count * r.qty) / ingInfo.spec / ingInfo.yield;
      const key = dateObj.getTime() + '_' + r.ingCd;
      if (!resultMap[key]) resultMap[key] = { date: dateObj, dow, ingCd: r.ingCd, ingName: ingInfo.name, qty: 0 };
      resultMap[key].qty += usedQty;
cpList.forEach(cp => {
        if (!cp.isNormalCp && cp.codes.includes(menuCd) &&
            dateObj >= cp.startDate && dateObj <= cp.endDate) {
          cpActualDaysMap[cp.name].add(dateStr);
          cpActualIngKeys.add(key);

          // ★切替CPは差し引きしない
          // ★CP開始前30日以内に実績がない食材は「CP専用新食材」と自動判定し差し引きしない
          if (!cp.isSwitchCp) {
            const preStart = new Date(cp.startDate);
            preStart.setDate(preStart.getDate() - PRE_CP_JISSEKI_DAYS);
            const dates = ingJissekiDates[r.ingCd] || [];
            const hasPreCpJisseki = dates.some(t => t >= preStart.getTime() && t < cp.startDate.getTime());
            if (hasPreCpJisseki) {
              const subtractKey = String(r.ingCd) + '_' + dow;
              if (!cpSubtractMap[subtractKey]) cpSubtractMap[subtractKey] = 0;
              cpSubtractMap[subtractKey] += usedQty;
            }
          }
        }
      });
    });
  });

  // CP推定行を追加（実績3日未満のみ）
  const CP_SWITCH_DAYS = 3;
  const salesSheet2 = ss.getSheetByName('日次売上実績');
  const customersMap = {};
  if (salesSheet2 && salesSheet2.getLastRow() > 1) {
    salesSheet2.getRange(2, 1, salesSheet2.getLastRow() - 1, 5).getValues().forEach(row => {
      if (row[0] && row[4] > 0) {
        const ds = Utilities.formatDate(row[0] instanceof Date ? row[0] : new Date(row[0]), 'Asia/Tokyo', 'yyyy/MM/dd');
        customersMap[ds] = Number(row[4]);
      }
    });
  }
  // 日次売上実績に未登録の日は実績データシートの総客数（D列）で補完
  // （日次売上実績の更新が遅れていてもCP推定行が正しく計算されるようにする）
  if (logSheet && logSheet.getLastRow() > 1) {
    logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 4).getValues().forEach(row => {
      if (!row[0] || row[3] <= 0) return;
      const d = row[0] instanceof Date ? row[0] : new Date(row[0]);
      const ds = Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy/MM/dd');
      if (!customersMap[ds]) customersMap[ds] = Number(row[3]);
    });
  }
  cpList.forEach(cp => {
    const actualDays = cpActualDaysMap[cp.name].size;
    if (actualDays >= CP_SWITCH_DAYS) { console.log('CP「' + cp.name + '」: 実績' + actualDays + '日 → 実績ベースに切替済み'); return; }
    console.log('CP「' + cp.name + '」: 実績' + actualDays + '日 → 推定行を追加中');
    const cur = new Date(Math.max(cp.startDate.getTime(), limitDate.getTime()));
    while (cur <= cp.endDate && cur < today) {
      const dateStr = Utilities.formatDate(cur, 'Asia/Tokyo', 'yyyy/MM/dd');
      const dow = cur.getDay() + 1;
      const customers = customersMap[dateStr] || 0;
      if (customers > 0) {
        const estDishes = Math.ceil(customers * cp.rate);
        cp.codes.forEach(menuCd => {
          if (!menuMap[menuCd]) return;
          menuMap[menuCd].forEach(r => {
            const ingInfo = ingMap[r.ingCd] || { name: '', spec: 1, yield: 1 };
            const key = cur.getTime() + '_' + r.ingCd;
            if (cpActualIngKeys.has(key)) return;
            const estQty = (estDishes * r.qty) / (ingInfo.spec || 1) / ingInfo.yield;
            if (!resultMap[key]) resultMap[key] = { date: new Date(cur), dow, ingCd: r.ingCd, ingName: ingInfo.name, qty: 0 };
            resultMap[key].qty += estQty;
          });
        });
      }
      cur.setDate(cur.getDate() + 1);
    }
  });
  if (cpList.length > 0) console.log('CP処理完了: ' + cpList.length + '件');

  // 食材日時実績シートの処理
  // 今回 F='' で実績・推定を書き直す食材コードのセット
  // これらの食材の古い F='推定' 行は引き継がない（旧コードで書かれた誤値を除去するため）
  const refreshedIngCodes = new Set(Object.values(resultMap).map(r => String(r.ingCd)));
  const estimationRows = [];
  if (outSheet.getLastRow() > 1) {
    outSheet.getRange(2, 1, outSheet.getLastRow() - 1, 6).getValues().forEach(row => {
      if ((row[5] || '').toString().trim() !== '推定') return;
      const d = row[0] instanceof Date ? row[0] : new Date(row[0]);
      d.setHours(0,0,0,0);
      if (d < limitDate) return; // 28日以上前の推定行は削除してシートを軽量化
      if (refreshedIngCodes.has(String(row[2]))) return; // 今回書き直す食材の推定行は除去
      estimationRows.push(row);
    });
  }
  const actualRows = Object.values(resultMap)
    .sort((a, b) => a.date - b.date || String(a.ingCd).localeCompare(String(b.ingCd)))
    .map(r => [r.date, r.dow, r.ingCd, r.ingName, r.qty, '']);
  outSheet.clearContents();
  outSheet.getRange(1, 1, 1, 6).setValues([['日付', '曜日', '食材コード', '食材名', '使用量', 'フラグ']]);
  const allRows = [...actualRows, ...estimationRows].sort((a, b) => {
    const da = a[0] instanceof Date ? a[0] : new Date(a[0]);
    const db = b[0] instanceof Date ? b[0] : new Date(b[0]);
    return da - db;
  });
  if (allRows.length > 0) {
    outSheet.getRange(2, 1, allRows.length, 6).setValues(allRows);
    outSheet.getRange(2, 1, allRows.length, 1).setNumberFormat('yyyy/mm/dd');
    allRows.forEach((row, i) => {
      if ((row[5] || '').toString().trim() === '推定')
        outSheet.getRange(i + 2, 1, 1, 6).setBackground('#fff9e6');
    });
  }

  // 食材コード統合
  if (Object.keys(ingCodeMap).length > 0) {
    const lastRowOut = outSheet.getLastRow();
    if (lastRowOut > 1) {
      const allData = outSheet.getRange(2, 1, lastRowOut - 1, 6).getValues();
      let mergeCount = 0;
      const newCodeRowMap = {};
      allData.forEach((row, i) => {
        const code = Number(row[2]);
        if (!ingCodeMap[code] && row[0] && row[2]) {
          const ts = (row[0] instanceof Date ? row[0] : new Date(row[0])).getTime();
          newCodeRowMap[ts + '_' + code] = i;
        }
      });
      allData.forEach((row, i) => {
        const code = Number(row[2]); const newCode = ingCodeMap[code]; if (!newCode) return;
        const ts = (row[0] instanceof Date ? row[0] : new Date(row[0])).getTime();
        const existingIdx = newCodeRowMap[ts + '_' + newCode];
        if (existingIdx !== undefined) { allData[existingIdx][4] = Number(allData[existingIdx][4]) + Number(row[4]); allData[i][2] = ''; }
        else { allData[i][2] = newCode; const ni = ingMap[newCode]; if (ni) allData[i][3] = ni.name; newCodeRowMap[ts + '_' + newCode] = i; }
        mergeCount++;
      });
      if (mergeCount > 0) {
        const mergedRows = allData.filter(row => row[2] !== '');
        outSheet.clearContents();
        outSheet.getRange(1, 1, 1, 6).setValues([['日付', '曜日', '食材コード', '食材名', '使用量', 'フラグ']]);
        if (mergedRows.length > 0) outSheet.getRange(2, 1, mergedRows.length, 6).setValues(mergedRows);
      }
    }
  }

  console.log('食材日時実績を更新しました');
  console.log('  実績行: ' + actualRows.length + '行');
  console.log('  推定行（保持）: ' + estimationRows.length + '行');
  return cpSubtractMap;
}

// ─────────────────────────────────────────────────────────────────────
// 天気取得
// ─────────────────────────────────────────────────────────────────────

function fetchWeather_OpenMeteo_UpdateDB() {
  const SHEET_NAME = '天候生データ';
  const ss    = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('「' + SHEET_NAME + '」シートが見つかりません。');

  const lat = 33.645, lon = 130.691;
  const tz  = 'Asia/Tokyo';

  // ── 既存データをMapに読み込む ──────────────────────────
  const dataMap = {};
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    sheet.getRange(2, 1, lastRow - 1, 8).getValues().forEach(row => {
      if (row[0] instanceof Date) {
        const key = Utilities.formatDate(row[0], tz, 'yyyy-MM-dd');
        dataMap[key] = row;
      }
    });
  }

  const todayStart = new Date(); todayStart.setHours(0,0,0,0);
  const nowStr     = Utilities.formatDate(new Date(), tz, 'yyyy/MM/dd');

  // ── ① archive API：過去データで空白期間を補完 ──────────
  const existingDates = Object.keys(dataMap).sort();
  const missingRanges = [];
  const yesterday = new Date(todayStart); yesterday.setDate(yesterday.getDate() - 1);

  if (existingDates.length > 0) {
    // (A) 既存行でB列(天気コード)が空の場合 → 先頭日〜昨日を再取得
    const hasEmpty = existingDates.some(k => dataMap[k][1] === null || dataMap[k][1] === '');
    if (hasEmpty) {
      const firstDate = new Date(existingDates[0] + 'T00:00:00+09:00');
      missingRanges.push({ start: firstDate, end: yesterday });
    }
    // (B) 直近の日付が丸ごと欠損している場合 → 最終日翌日〜昨日を取得
    const lastDateStr = existingDates[existingDates.length - 1];
    const lastDate    = new Date(lastDateStr + 'T00:00:00+09:00');
    if (lastDate < yesterday) {
      const gapStart = new Date(lastDate); gapStart.setDate(gapStart.getDate() + 1);
      console.log('直近データ欠損を検出: ' + Utilities.formatDate(gapStart, tz, 'yyyy-MM-dd')
        + ' 〜 ' + Utilities.formatDate(yesterday, tz, 'yyyy-MM-dd'));
      missingRanges.push({ start: gapStart, end: yesterday });
    }
  } else {
    // データが1件もない場合：実績データの最初の日付から取得
    const jissekiSheet = ss.getSheetByName('実績データ');
    let archiveStart = new Date(todayStart); archiveStart.setDate(archiveStart.getDate() - 180);
    if (jissekiSheet && jissekiSheet.getLastRow() > 1) {
      const firstRow = jissekiSheet.getRange(2, 1, 1, 1).getValue();
      if (firstRow instanceof Date) { archiveStart = new Date(firstRow); archiveStart.setHours(0,0,0,0); }
    }
    missingRanges.push({ start: archiveStart, end: yesterday });
  }

  // archive APIで補完
  missingRanges.forEach(range => {
    const startStr = Utilities.formatDate(range.start, tz, 'yyyy-MM-dd');
    const endStr   = Utilities.formatDate(range.end,   tz, 'yyyy-MM-dd');
    if (startStr >= endStr) return;

    console.log('archive API取得: ' + startStr + ' 〜 ' + endStr);
    const archiveUrl =
      'https://archive-api.open-meteo.com/v1/archive'
      + '?latitude=' + lat + '&longitude=' + lon
      + '&start_date=' + startStr + '&end_date=' + endStr
      + '&daily=weathercode,temperature_2m_max,temperature_2m_min,precipitation_sum'
      + '&timezone=' + tz;
    try {
      const res  = UrlFetchApp.fetch(archiveUrl, { muteHttpExceptions: true });
      if (res.getResponseCode() !== 200) {
        console.log('archive API エラー: ' + res.getResponseCode() + ' ' + res.getContentText().substring(0, 200));
        return;
      }
      const json   = JSON.parse(res.getContentText());
      const dates  = json.daily.time, codes = json.daily.weathercode;
      const maxT   = json.daily.temperature_2m_max, minT = json.daily.temperature_2m_min;
      const precip = json.daily.precipitation_sum || [];
      dates.forEach((d, i) => {
        // archive APIは過去データのみ → 全て actual
        if (!dataMap[d] || dataMap[d][1] === null || dataMap[d][1] === '') {
          const dateObj = new Date(d + 'T00:00:00+09:00');
          dataMap[d] = [dateObj, codes[i], maxT[i], minT[i], 'actual', nowStr, convertWeatherCode_(codes[i]), precip[i] != null ? precip[i] : ''];
        }
      });
      console.log('archive API取得完了: ' + dates.length + '件');
    } catch(e) {
      console.log('archive API例外: ' + e.message);
    }
  });

  // ── ② forecast API：直近92日＋今後15日 ────────────────
  const forecastUrl =
    'https://api.open-meteo.com/v1/forecast'
    + '?latitude=' + lat + '&longitude=' + lon
    + '&daily=weathercode,temperature_2m_max,temperature_2m_min,precipitation_sum'
    + '&past_days=92&forecast_days=15&timezone=' + tz;
  try {
    const res  = UrlFetchApp.fetch(forecastUrl, { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) {
      console.log('forecast API エラー: ' + res.getResponseCode());
    } else {
      const json   = JSON.parse(res.getContentText());
      const dates  = json.daily.time, codes = json.daily.weathercode;
      const maxT   = json.daily.temperature_2m_max, minT = json.daily.temperature_2m_min;
      const precip = json.daily.precipitation_sum || [];
      dates.forEach((d, i) => {
        const dateObj = new Date(d + 'T00:00:00+09:00');
        const status  = dateObj < todayStart ? 'actual' : 'forecast';
        // nullデータで既存の有効なデータを上書きしない
        if (codes[i] === null || codes[i] === undefined || maxT[i] === null || minT[i] === null) {
          if (dataMap[d]) { dataMap[d][4] = status; dataMap[d][5] = nowStr; }
          return;
        }
        dataMap[d] = [dateObj, codes[i], maxT[i], minT[i], status, nowStr, convertWeatherCode_(codes[i]), precip[i] != null ? precip[i] : ''];
      });
      console.log('forecast API取得完了: ' + dates.length + '件');
    }
  } catch(e) {
    console.log('forecast API例外: ' + e.message);
  }

  // ── シートに書き戻す ──────────────────────────────────
  const finalData = Object.values(dataMap).sort((a, b) => a[0].getTime() - b[0].getTime());
  if (sheet.getMaxRows() > 1) sheet.getRange(2, 1, sheet.getMaxRows() - 1, 8).clearContent();
  if (finalData.length > 0)   sheet.getRange(2, 1, finalData.length, 8).setValues(finalData);

  // ヘッダーにH列「降水量(mm)」を追加（初回のみ）
  if (sheet.getRange(1, 8).getValue() === '') {
    sheet.getRange(1, 8).setValue('降水量(mm)');
  }

  // B〜D列・H列に数値書式を設定
  if (finalData.length > 0) {
    sheet.getRange(2, 2, finalData.length, 1).setNumberFormat('0');   // 天気コード
    sheet.getRange(2, 3, finalData.length, 2).setNumberFormat('0.0'); // 気温
    sheet.getRange(2, 8, finalData.length, 1).setNumberFormat('0.0'); // 降水量
  }

  console.log('天候データ更新完了: ' + finalData.length + '件（うちactual: '
    + finalData.filter(r => r[4] === 'actual').length + '件、forecast: '
    + finalData.filter(r => r[4] === 'forecast').length + '件）');
}

// 天候データの実測が古ければ自動で再取得する（定期実行トリガーが止まった場合の自己修復）
// buildSalesForecast の冒頭から呼ばれる。実測の最新日が「昨日」に達していなければ再取得し、
// 直っていなければ1日1回だけアラートメールを送る（トリガー再設定を促す）。
function ensureFreshWeather_(ss) {
  const tz = 'Asia/Tokyo';
  const wx = ss.getSheetByName('天候生データ');
  const today = new Date(); today.setHours(0,0,0,0);
  const yesterday = new Date(today); yesterday.setDate(yesterday.getDate() - 1);

  const latestActual_ = () => {
    if (!wx || wx.getLastRow() < 2) return null;
    let latest = null;
    wx.getRange(2, 1, wx.getLastRow()-1, 5).getValues().forEach(r => {
      if (r[0] instanceof Date && r[4] === 'actual' && (!latest || r[0] > latest)) latest = r[0];
    });
    return latest;
  };

  let latest = latestActual_();
  if (latest && latest >= yesterday) return; // 実測が最新（昨日以降）＝正常

  console.log('天候の実測が古い（最新actual: ' + (latest ? Utilities.formatDate(latest, tz, 'yyyy-MM-dd') : 'なし') + '）→ 自動再取得します');
  let refetchErr = '';
  try { fetchWeather_OpenMeteo_UpdateDB(); } catch(e) { refetchErr = e.message || String(e); console.log('自動再取得エラー: ' + refetchErr); }

  // 再取得後もまだ古い場合のみ、1日1回アラートメールを送る（Open-Meteoの反映遅延なら翌日直るので過剰通知を防ぐ）
  latest = latestActual_();
  if (latest && latest >= yesterday) { console.log('天候データを自動復旧しました'); return; }

  // 権限エラー（新しいブックで外部通信が未承認）かどうかを判定して、対処法を出し分ける
  const looksAuth = /authoriz|permission|権限|scope|OAuth|承認/i.test(refetchErr);
  try {
    const props = PropertiesService.getScriptProperties();
    const flagKey = 'weatherStaleAlerted_' + Utilities.formatDate(today, tz, 'yyyyMMdd');
    if (props.getProperty(flagKey)) return; // 本日通知済み
    const email = Session.getActiveUser().getEmail();
    const cause = refetchErr
      ? ('自動再取得でエラーが出ました：\n  ' + refetchErr + '\n\n')
      : '自動再取得でも直りませんでした。\n\n';
    const howto = looksAuth
      ? ('【原因】このブックで外部通信（天候API）の権限がまだ承認されていません。\n'
         + '【対処】Apps Scriptエディタで関数 fetchWeather_OpenMeteo_UpdateDB を1回手動実行し、\n'
         + '　　　　表示される許可ダイアログで「許可」を押してください。1回承認すれば以降は自動で動きます。\n')
      : ('【対処】Apps Scriptで setupTriggersSplit を1回実行して定期実行を再設定するか、\n'
         + '　　　　fetchWeather_OpenMeteo_UpdateDB を手動実行してください。\n'
         + '（Open-Meteo側の反映遅延の場合は翌日には自動で解消します）\n');
    MailApp.sendEmail(email, '【発注システム】天候データが更新されていません',
      '天候の実測データが「' + (latest ? Utilities.formatDate(latest, tz, 'yyyy-MM-dd') : 'なし') + '」で止まっています。\n\n'
      + cause + howto);
    props.setProperty(flagKey, '1');
    console.log('天候更新停止のアラートメールを送信しました（authエラー疑い: ' + looksAuth + '）');
  } catch(e) { console.log('アラートメール送信失敗: ' + e.message); }
}

function convertWeatherCode_(code) {
  if (code === 0) return '快晴'; if (code === 1) return '晴れ';
  if (code === 2 || code === 3) return '曇り'; if (code === 45 || code === 48) return '曇り';
  if (code >= 51 && code <= 57) return '小雨'; if (code === 61) return '小雨';
  if (code === 63 || code === 81) return '雨'; if (code === 65 || code === 82) return '豪雨';
  if (code === 66 || code === 67) return '豪雨';
  if (code === 71 || code === 73 || code === 77) return '雪'; if (code === 85) return '雪';
  if (code === 75 || code === 86) return '大雪';
  if (code === 95 || code === 96 || code === 99) return '雷雨';
  return '曇り';
}

// ─────────────────────────────────────────────────────────────────────
// 過去データの降水量を一括補完（初回1回だけ実行する復旧用）
// ─────────────────────────────────────────────────────────────────────

function backfillPrecipitation() {
  const SHEET_NAME = '天候生データ';
  const ss    = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('「' + SHEET_NAME + '」シートが見つかりません。');

  const lat = 33.645, lon = 130.691;
  const tz  = 'Asia/Tokyo';

  // ヘッダーにH列「降水量(mm)」を追加
  if (sheet.getRange(1, 8).getValue() === '') {
    sheet.getRange(1, 8).setValue('降水量(mm)');
  }

  // 既存データを読み込む（8列）
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { console.log('データなし'); return; }
  const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();

  // 降水量が空の実績行を抽出して日付範囲を特定
  const emptyDates = [];
  data.forEach((row, i) => {
    if (!row[0] || row[4] !== 'actual') return;
    if (row[7] === '' || row[7] === null) emptyDates.push({ rowIndex: i + 2, date: Utilities.formatDate(row[0], tz, 'yyyy-MM-dd') });
  });

  if (emptyDates.length === 0) { console.log('補完対象なし（全行に降水量あり）'); return; }

  const startStr = emptyDates[0].date;
  const endStr   = emptyDates[emptyDates.length - 1].date;
  console.log('補完対象: ' + emptyDates.length + '日 (' + startStr + ' 〜 ' + endStr + ')');

  // archive APIで一括取得（最大1年分ずつ分割）
  const precipMap = {};
  const MS_PER_DAY = 86400000;
  const start = new Date(startStr + 'T00:00:00+09:00');
  const end   = new Date(endStr   + 'T00:00:00+09:00');

  let cur = new Date(start);
  while (cur <= end) {
    const chunkEnd = new Date(Math.min(cur.getTime() + 364 * MS_PER_DAY, end.getTime()));
    const s = Utilities.formatDate(cur,      tz, 'yyyy-MM-dd');
    const e = Utilities.formatDate(chunkEnd, tz, 'yyyy-MM-dd');
    const url = 'https://archive-api.open-meteo.com/v1/archive'
      + '?latitude=' + lat + '&longitude=' + lon
      + '&start_date=' + s + '&end_date=' + e
      + '&daily=precipitation_sum&timezone=' + tz;
    try {
      const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      if (res.getResponseCode() === 200) {
        const json = JSON.parse(res.getContentText());
        (json.daily.time || []).forEach((d, i) => {
          const v = json.daily.precipitation_sum[i];
          if (v != null) precipMap[d] = v;
        });
        console.log('取得: ' + s + ' 〜 ' + e + ' (' + (json.daily.time || []).length + '件)');
      } else {
        console.log('APIエラー ' + s + '〜' + e + ': ' + res.getResponseCode());
      }
    } catch(e2) {
      console.log('例外 ' + s + '〜' + e + ': ' + e2.message);
    }
    cur = new Date(chunkEnd.getTime() + MS_PER_DAY);
    Utilities.sleep(500); // APIレート制限対策
  }

  // シートのH列に書き込む
  let updated = 0;
  emptyDates.forEach(item => {
    const v = precipMap[item.date];
    if (v != null) {
      sheet.getRange(item.rowIndex, 8).setValue(v);
      updated++;
    }
  });

  sheet.getRange(2, 8, lastRow - 1, 1).setNumberFormat('0.0');
  console.log('補完完了: ' + updated + '/' + emptyDates.length + '日');
  SpreadsheetApp.getUi().alert('降水量補完完了\n' + updated + '/' + emptyDates.length + '日分を更新しました。\n\n次に buildSalesForecast() を実行してください。');
}

// ─────────────────────────────────────────────────────────────────────
// 過去データ一括取込
// ─────────────────────────────────────────────────────────────────────

function importBulkHistoricalData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('実績データ');
  const importSheet = ss.getSheetByName('インポート用');
  if (!importSheet || importSheet.getLastRow() < 1) {
    showMessage_('「インポート用」シートが見つかりません。'); return;
  }
  const existingKeys = new Set();
  if (sheet && sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues().forEach(r => {
      if (r[0] && r[4]) {
        const ds = r[0] instanceof Date ? Utilities.formatDate(r[0], 'Asia/Tokyo', 'yyyy/MM/dd') : r[0].toString();
        existingKeys.add(ds + '_' + String(r[4]));
      }
    });
  }
  const importData = importSheet.getRange(1, 1, importSheet.getLastRow(), 7).getValues().filter(r => r[0]);
  const dailyStats = {}, toAppend = [];
  importData.forEach(row => {
    const dateStr = row[0] instanceof Date ? Utilities.formatDate(row[0], 'Asia/Tokyo', 'yyyy/MM/dd') : row[0].toString();
    const menuCd = row[4] ? String(row[4]) : '';
    const key = dateStr + '_' + menuCd;
    if (!existingKeys.has(key) && menuCd) { toAppend.push([dateStr, row[1], row[2], row[3], row[4], row[5], row[6]]); existingKeys.add(key); }
    if (!dailyStats[dateStr] && (Number(row[2]) > 0 || Number(row[3]) > 0))
      dailyStats[dateStr] = { sales: Number(row[2]), customers: Number(row[3]) };
  });
  if (toAppend.length === 0) { showMessage_('追加すべき新しいデータはありませんでした'); return; }
  if (!sheet) { const ns = ss.insertSheet('実績データ'); ns.appendRow(['日付', '曜日', '総売上', '総客数', 'メニューCD', 'メニュー名', '出食数']); }
  const targetSheet = ss.getSheetByName('実績データ');
  toAppend.sort((a, b) => new Date(a[0]) - new Date(b[0]));
  targetSheet.getRange(targetSheet.getLastRow() + 1, 1, toAppend.length, 7).setValues(toAppend);
  if (Object.keys(dailyStats).length > 0) writeDailySalesSheet_(ss, dailyStats);
  const cpSubtractMap = buildIngredientDailyUsage();
  saveCpSubtractMap_(cpSubtractMap);
  showMessage_(`一括取込完了\n追記: ${toAppend.length}行\n重複スキップ: ${importData.length - toAppend.length}行`);
}

// ─────────────────────────────────────────────────────────────────────
// 復旧用：指定日以降を削除して再取込
// ─────────────────────────────────────────────────────────────────────

function recoverFromDate() {
  // ★UIダイアログで復旧開始日を入力（ハードコード廃止）
  let FROM_DATE;
  try {
    const ui       = SpreadsheetApp.getUi();
    const response = ui.prompt(
      '復旧開始日の入力',
      '復旧開始日を入力してください（例: 2026/05/12）\n\n※この日付以降のデータを削除してCSVから再取込します。',
      ui.ButtonSet.OK_CANCEL
    );
    if (response.getSelectedButton() !== ui.Button.OK) {
      console.log('復旧をキャンセルしました'); return;
    }
    FROM_DATE = response.getResponseText().trim();
    // 入力値の検証
    const testDate = new Date(FROM_DATE);
    if (isNaN(testDate.getTime())) {
      ui.alert('日付の形式が正しくありません。\n例: 2026/05/12\nもう一度実行してください。');
      return;
    }
    // 確認ダイアログ
    const confirm = ui.alert(
      '確認',
      FROM_DATE + ' 以降のデータを削除して再取込します。\nよろしいですか？',
      ui.ButtonSet.OK_CANCEL
    );
    if (confirm !== ui.Button.OK) { console.log('復旧をキャンセルしました'); return; }
  } catch(e) {
    // トリガー実行など UI が使えない環境では従来通りコードに書く
    FROM_DATE = '2026/05/12';
    console.log('UI非対応環境のためハードコード日付を使用: ' + FROM_DATE);
  }

  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const fromDateObj = new Date(FROM_DATE); fromDateObj.setHours(0,0,0,0);

  const deleteFrom = (sheetName, colCount) => {
    const s = ss.getSheetByName(sheetName);
    if (!s || s.getLastRow() < 2) return 0;
    const data = s.getRange(2, 1, s.getLastRow()-1, colCount).getValues();
    const keep = data.filter(r => {
      if (!r[0]) return false;
      const d = r[0] instanceof Date ? r[0] : new Date(r[0]); d.setHours(0,0,0,0);
      return d < fromDateObj;
    });
    s.getRange(2, 1, s.getMaxRows()-1, colCount).clearContent();
    if (keep.length > 0) s.getRange(2, 1, keep.length, colCount).setValues(keep);
    return data.length - keep.length;
  };

  const r1 = deleteFrom('実績データ', 7);
  const r2 = deleteFrom('日次売上実績', 5);
  console.log(FROM_DATE + '以降を削除しました');
  console.log('  実績データ: ' + r1 + '行削除');
  console.log('  日次売上実績: ' + r2 + '行削除');
  console.log('次に importFooditCSV() を実行してください');

  try {
    SpreadsheetApp.getUi().alert(
      '削除完了',
      FROM_DATE + ' 以降を削除しました。\n\n実績データ: ' + r1 + '行\n日次売上実績: ' + r2 + '行\n\n次に importFooditCSV() を実行してください。',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch(e) {}
}

// ─────────────────────────────────────────────────────────────────────
// 重複除去
// ─────────────────────────────────────────────────────────────────────

function deduplicateJissekiData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('実績データ');
  if (!sheet || sheet.getLastRow() < 2) return;
  const data = sheet.getRange(2, 1, sheet.getLastRow()-1, 7).getValues().filter(r => r[0]);
  const seen = new Set(), cleaned = [];
  data.forEach(row => {
    const ds = row[0] instanceof Date ? Utilities.formatDate(row[0], 'Asia/Tokyo', 'yyyy/MM/dd') : row[0].toString();
    const key = ds + '_' + String(row[4]);
    if (seen.has(key)) return; seen.add(key); cleaned.push(row);
  });
  cleaned.sort((a, b) => new Date(a[0]) - new Date(b[0]));
  sheet.getRange(2, 1, sheet.getMaxRows()-1, 7).clearContent();
  if (cleaned.length > 0) sheet.getRange(2, 1, cleaned.length, 7).setValues(cleaned);
  showMessage_('実績データ重複除去完了\n削除: ' + (data.length - cleaned.length) + '行');
}

function deduplicateDailySales() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('日次売上実績');
  if (!sheet || sheet.getLastRow() < 2) return;
  const data = sheet.getRange(2, 1, sheet.getLastRow()-1, 5).getValues().filter(r => r[0]);
  const seen = new Set(), cleaned = [];
  data.forEach(row => { const key = row[0].toString(); if (seen.has(key)) return; seen.add(key); cleaned.push(row); });
  cleaned.sort((a, b) => new Date(a[0]) - new Date(b[0]));
  sheet.getRange(2, 1, sheet.getMaxRows()-1, 5).clearContent();
  if (cleaned.length > 0) sheet.getRange(2, 1, cleaned.length, 5).setValues(cleaned);
  console.log('日次売上実績重複除去完了: ' + (data.length - cleaned.length) + '行削除');
}


// ═══════════════════════════════════════════════════════════════════════
// 【E】 売上予測セクション
// ═══════════════════════════════════════════════════════════════════════

function buildSalesForecast() {

  const LONG_CP_DAYS = 21; // 長期CPの定義：21日（3週間）以上

  try { checkProvisionalEvents_(false); } catch(e) { console.log('暫定日チェックskip: ' + e.message); } // 暫定日イベントの日付確認リマインド

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 天候データが古ければ自動で再取得（毎朝の天候トリガーが止まっても、毎日動くこの予測処理が天候を自己修復する二重の安全網）
  try { ensureFreshWeather_(ss); } catch(e) { console.log('天候自己修復skip: ' + e.message); }

  const salesSheet = ss.getSheetByName('日次売上実績');
  if (!salesSheet || salesSheet.getLastRow() < 2) { console.log('日次売上実績にデータがありません'); return; }

  const salesData = salesSheet.getRange(2, 1, salesSheet.getLastRow()-1, 5).getValues().filter(r => r[0]);
  const salesMap = {};
  salesData.forEach(row => {
    salesMap[toDateStr_(row[0])] = { dow: Number(row[1]), sales: Number(row[3]), customers: Number(row[4]) };
  });

  const today = new Date(); today.setHours(0, 0, 0, 0);
  const allActualDates = Object.keys(salesMap).map(s => new Date(s)).filter(d => !isNaN(d)).sort((a,b) => a-b);
  const startDate = allActualDates.length > 0 ? allActualDates[0] : new Date(today);
  startDate.setHours(0, 0, 0, 0);
  const endDate = new Date(today.getFullYear(), today.getMonth() + 2, 0);

  // 除外日・イベント・新商品
  const excludedDates = new Set(), eventList = [], newProductList = [], repeatEventList = [];

  try { getOutlierExcludedDates_().forEach(d => excludedDates.add(d)); } catch(e) {}

  const cpSettingSheet = ss.getSheetByName('CP設定');
  if (cpSettingSheet && cpSettingSheet.getLastRow() > 1) {
    cpSettingSheet.getRange(2, 1, cpSettingSheet.getLastRow()-1, 9).getValues().forEach(row => {
      const cpName = (row[0] || '').toString().trim();
      if (!cpName) return;
      if ((row[8] || '').toString().trim() === '休止') return; // CP休止行は売上予測に影響させない（別処理）
      const note   = (row[7] || '').toString().trim();
      if (note === '除外') return; // 発注計算用除外フラグ行はスキップ
      const rate   = Number(row[5]) || 0;  // 補正倍率（0 or 空 = 実績から自動計算）
      const repeat = (row[6] || '').toString().trim();
      const startRaw = row[1], endRaw = row[2];
      const start = startRaw instanceof Date ? startRaw : (startRaw ? new Date(String(startRaw).replace(/^(\d{4})(\d{2})(\d{2})$/, '$1/$2/$3')) : null);
      const end   = endRaw   instanceof Date ? endRaw   : (endRaw   ? new Date(String(endRaw  ).replace(/^(\d{4})(\d{2})(\d{2})$/, '$1/$2/$3')) : null);
      const hasValidDates = start && end && !isNaN(start.getTime()) && !isNaN(end.getTime());
      if (hasValidDates) {
        // 日付指定CP → eventListへ（繰り返し列は無視。実績から倍率を自動計算）
        start.setHours(0,0,0,0); end.setHours(0,0,0,0);
        const cpCutoff = new Date(); cpCutoff.setDate(cpCutoff.getDate() - 28); cpCutoff.setHours(0,0,0,0);
        if (end < cpCutoff) return;
        eventList.push({ name: cpName, start, end, rate: rate > 0 ? rate : 1.0, priority: 5, isCP: true });
      } else if (repeat && /毎年\d+月\d+日〜次の日曜/.test(repeat)) {
        // 毎年X月Y日〜次の日曜 → 今年・来年分の期間を計算してeventListへ追加
        const mAS = repeat.match(/毎年(\d+)月(\d+)日〜次の日曜/);
        const tMonth = Number(mAS[1]) - 1, tDay = Number(mAS[2]);
        const cpCutoffAS = new Date(); cpCutoffAS.setDate(cpCutoffAS.getDate() - 28); cpCutoffAS.setHours(0,0,0,0);
        [today.getFullYear(), today.getFullYear() + 1].forEach(yr => {
          const sd = new Date(yr, tMonth, tDay); sd.setHours(0,0,0,0);
          const ed = new Date(sd);
          while (ed.getDay() !== 0) ed.setDate(ed.getDate() + 1); // 次の日曜まで
          ed.setHours(0,0,0,0);
          if (ed < cpCutoffAS) return; // 28日以上前に終了したものはスキップ
          if (sd > endDate) return;    // 予測期間外はスキップ
          eventList.push({ name: cpName, start: sd, end: ed, rate: rate > 0 ? rate : 1.0, priority: 5, isCP: true });
        });
      } else if (repeat && (/毎月\d+/.test(repeat) || /毎年\d+月\d+/.test(repeat))) {
        // 日付なし＋繰り返しパターン → repeatEventListへ（倍率は実績から自動計算）
        repeatEventList.push({ name: cpName, rate, priority: 5, repeat });
      }
    });
  }

  const evSheet = ss.getSheetByName('イベント登録');
  if (evSheet && evSheet.getLastRow() > 1) {
    const evCols = Math.max(10, evSheet.getLastColumn()); // K列(繰り返し)が無くても安全に読む
    evSheet.getRange(2, 1, evSheet.getLastRow()-1, evCols).getValues().forEach(row => {
      if (!row[0]) return;
      const start = toDate_(row[1]), end = toDate_(row[2]);
      const type = (row[4] || '').toString().trim();
      const rate = Number(row[6]) || 1.0, priority = Number(row[7]) || 0;
      const flag = (row[8] || '').toString().trim();
      const repeatPat = (evCols >= 11 ? (row[10] || '') : '').toString().trim();
      const learnByName = (flag === '自動') || (repeatPat !== ''); // 繰り返し行は同名実績で倍率を自動学習する

      // 日付が空で「繰り返し」列にパターンがある行 → 今年・来年へ自動展開（繰り返しルールシート＋生成メニューの代替）
      if ((!start || !end) && repeatPat) {
        expandEventRepeat_(repeatPat, today, endDate).forEach(oc => {
          if (type === '除外') {
            const cur = new Date(oc.start);
            while (cur <= oc.end) { excludedDates.add(toDateStr_(cur)); cur.setDate(cur.getDate() + 1); }
          } else {
            eventList.push({ name: row[0], start: oc.start, end: oc.end, rate, priority, learnByName, provisional: oc.provisional, isCP: false });
          }
        });
        return;
      }

      if (!start || !end) return;
      if (type === '除外') {
        const cur = new Date(start);
        while (cur <= end) { excludedDates.add(toDateStr_(cur)); cur.setDate(cur.getDate() + 1); }
      } else if (type === '新商品') {
        const productCode = row[5] ? Number(row[5]) : null;
        newProductList.push({ name: row[0], start, end, productCode, priority });
        eventList.push({ name: row[0], start, end, rate: 1.0, priority, isCP: false });
      } else {
        eventList.push({ name: row[0], start, end, rate, priority, learnByName, isCP: false });
      }
    });
  }

  // CP設定から「実績ゼロ食材を含む近日開始CP商品」をnewProductListに自動追加
  // イベント登録に「新商品」と登録しなくてもCP設定だけで推定発注量が生成される
  {
    const cpSheetFcst = ss.getSheetByName('CP設定');
    const ingJissekiSh = ss.getSheetByName('食材日時実績');
    if (cpSheetFcst && cpSheetFcst.getLastRow() > 1) {
      const existIngSet = new Set();
      if (ingJissekiSh && ingJissekiSh.getLastRow() > 1) {
        ingJissekiSh.getRange(2, 1, ingJissekiSh.getLastRow()-1, 6).getValues()
          .filter(r => r[2] && (!r[5] || r[5].toString().trim() !== '推定'))
          .forEach(r => existIngSet.add(String(r[2])));
      }
      const pmShCp = ss.getSheetByName('商品マスタ');
      const cpProdIngMap = {};
      if (pmShCp && pmShCp.getLastRow() > 1) {
        pmShCp.getRange(2, 1, pmShCp.getLastRow()-1, 5).getValues()
          .filter(r => r[0] && r[4])
          .forEach(r => {
            const mc = String(Math.round(Number(r[0])));
            if (!cpProdIngMap[mc]) cpProdIngMap[mc] = [];
            cpProdIngMap[mc].push(String(r[4]));
          });
      }
      cpSheetFcst.getRange(2, 1, cpSheetFcst.getLastRow()-1, 9).getValues().forEach(row => {
        const cpName = (row[0]||'').toString().trim(); if (!cpName) return;
        if ((row[8] || '').toString().trim() === '休止') return; // CP休止行は新商品推定の対象外
        const start = toDate_(row[1]), end = toDate_(row[2]);
        if (!start || !end) return;
        const msTillStart = start.getTime() - today.getTime();
        if (msTillStart < -86400000 || msTillStart > 14 * 86400000) return; // 前日〜14日先のCPのみ
        const rawVal = row[3];
        const menuCodes = typeof rawVal === 'number'
          ? [String(Math.round(rawVal))]
          : String(rawVal||'').split(',').map(c=>c.trim()).filter(c=>c);
        menuCodes.forEach(mc => {
          const ingCodes = cpProdIngMap[mc] || [];
          if (ingCodes.length === 0 || !ingCodes.some(ic => !existIngSet.has(ic))) return;
          const entryName = cpName + '（商品' + mc + '）';
          if (newProductList.some(np => np.name === entryName)) return;
          newProductList.push({ name: entryName, start, end, productCode: Number(mc), priority: 0 });
          console.log('CP新商品検出: ' + entryName + ' → 発注推定値を自動生成します');
        });
      });
    }
  }

  // データ異常チェック
  const anomalies = [];
  Object.keys(salesMap).forEach(dateStr => {
    if (salesMap[dateStr].sales === 0 && !excludedDates.has(dateStr) && new Date(dateStr) < today)
      anomalies.push('【売上0・除外未登録】' + dateStr);
  });
  if (allActualDates.length > 1) {
    const lastDate = allActualDates[allActualDates.length - 1];
    const cur = new Date(allActualDates[0]); cur.setDate(cur.getDate() + 1);
    while (cur <= lastDate) {
      const ds = toDateStr_(cur);
      if (!salesMap[ds] && !excludedDates.has(ds)) anomalies.push('【日付抜け】' + ds);
      cur.setDate(cur.getDate() + 1);
    }
  }
  if (anomalies.length > 0) {
    const msg = 'データ異常が検出されました。\n\n' + anomalies.join('\n') + '\n\n※除外すべき日はイベント登録シートに「除外」区分で追加してください。';
    try { const ui = SpreadsheetApp.getUi(); if (ui.alert('データ異常', msg, ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return; }
    catch (e) { console.log(msg); }
  }

  // 天候・気温・降水量帯係数（長期CP期間を除外して逆算）
  const cpDatesForWx = new Set();
  eventList.forEach(ev => {
    const days = (ev.end - ev.start) / 86400000;
    if (days < LONG_CP_DAYS) return;
    const c = new Date(ev.start);
    while (c <= ev.end) { cpDatesForWx.add(toDateStr_(c)); c.setDate(c.getDate()+1); }
  });
  const { coeffs: weatherCoeffMap, sampleInfo: wxSampleInfo } = calcWeatherCoeff_(ss, salesMap, excludedDates, cpDatesForWx);
  const { coeffs: tempCoeffMap }   = calcTempCoeff_(ss, salesMap, excludedDates, cpDatesForWx);
  const { coeffs: precipCoeffMap } = calcPrecipCoeff_(ss, salesMap, excludedDates, cpDatesForWx);
  const wxSheetForPrecip = ss.getSheetByName('天候生データ');
  console.log('天候係数（実績・CP除外）: ' + JSON.stringify(weatherCoeffMap));
  console.log('気温係数（実績・CP除外）: ' + JSON.stringify(tempCoeffMap));
  console.log('降水量帯係数（実績・CP除外）: ' + JSON.stringify(precipCoeffMap));
  writeWeatherCoeffAnalysis_(ss, wxSampleInfo, precipCoeffMap, tempCoeffMap);

  const monthPatternCoeff = calcMonthPatternCoeff_(salesMap, excludedDates);


  // ★祝日マスタの日付も曜日平均から除外
  const holidayTaggedDates = getHolidayTaggedDates_();          // 全タグ（土日含む）：トレンド・bias・月係数の計算で除外
  const realHolidayTaggedDates = getHolidayTaggedDates_(false); // 真の祝日のみ（土日タグは含めない）：曜日別平均の計算で除外
  // ★cpEventDates（日付指定CP/イベント期間）は通常実績として使う（除外しない）
  //   → 期間指定CPの実績もサンプルとして活用し、サンプル不足を防ぐ
  //   → イベント効果は eventActualRates / repeatEventCoeffMap で倍率学習して反映
  const cpEventDates = new Set(); // 後段（イベント実績計算）で参照するため保持
  eventList.forEach(ev => {
    const cur = new Date(ev.start);
    while (cur <= ev.end) { cpEventDates.add(toDateStr_(cur)); cur.setDate(cur.getDate() + 1); }
  });
  // ★毎年固定日イベント（例：ワールドペッパーライスデー7/3）は除外
  //   → 売上が大きく通常と異なるため曜日平均を歪めない
  //   → 毎月系（肉の日29日等）は17%増程度のため除外せず倍率学習で対応
  const repeatAnnualExcludedDates = new Set();
  repeatEventList.forEach(ev => {
    if (!/毎年\d+月\d+/.test(ev.repeat)) return; // 毎月系はスキップ
    const mAnn = ev.repeat.match(/毎年(\d+)月(\d+)/);
    if (!mAnn) return;
    const month = Number(mAnn[1]) - 1, day = Number(mAnn[2]);
    Object.keys(salesMap).forEach(ds => {
      const d = new Date(ds);
      if (d.getMonth() === month && d.getDate() === day) repeatAnnualExcludedDates.add(ds);
    });
    [today.getFullYear(), today.getFullYear() + 1].forEach(yr => {
      repeatAnnualExcludedDates.add(toDateStr_(new Date(yr, month, day)));
    });
  });
  if (repeatAnnualExcludedDates.size > 0) console.log('毎年固定日イベント除外: ' + [...repeatAnnualExcludedDates].sort().join(', '));
  const avgExcludedDates   = new Set([...excludedDates, ...holidayTaggedDates, ...repeatAnnualExcludedDates]);
  // 曜日別平均は通常の土日も実績として使い、直近10件の指数平滑で「今の水準」を反映する。真の祝日のみ除外
  const dowAvgExcludedDates = new Set([...excludedDates, ...realHolidayTaggedDates, ...repeatAnnualExcludedDates]);
  const weightedDowAvg     = calcWeightedDowAvg_(salesMap, dowAvgExcludedDates, today);
  console.log('加重平均（曜日別）: ' + JSON.stringify(Object.fromEntries(Object.entries(weightedDowAvg).map(([k,v]) => [k, Math.round(v)]))));

  const longTermTrend = calcLongTermTrend_(salesMap, avgExcludedDates, today);
  // バイアス補正はCP・イベント期間を除いた通常日の誤差のみで計算する
  // （CP期間の誤差はeventActualRatesで別学習しているため二重計上を防ぐ）
  const biasExcludedDates = new Set([...avgExcludedDates, ...cpEventDates]);
  const recentBiasCoeff = calcRecentBiasCoeff_(salesMap, biasExcludedDates, ss);
  const monthCoeffMap   = calcMonthCoeff_(salesMap, avgExcludedDates);

  const cpPostEndCoeff = calcCpPostEndCoeff_(eventList, salesMap, excludedDates, today);
  if (Object.keys(cpPostEndCoeff).length > 0) console.log('CP終了後補正: ' + JSON.stringify(cpPostEndCoeff));

  // 祝日マスタ
  const holidayInfoMap = {};
  const hmSheet = ss.getSheetByName('祝日マスタ');
  if (hmSheet && hmSheet.getLastRow() > 1) {
    hmSheet.getRange(2, 1, hmSheet.getLastRow()-1, 9).getValues().forEach(row => {
      if (!row[0]) return;
      holidayInfoMap[toDateStr_(row[0])] = { isOff: row[2], tag: row[6] || '通常営業', salesAvg: Number(row[8]) || 0, chainId: row[3] || '', chainLen: Number(row[4]) || 0, chainDay: Number(row[5]) || 0 };
    });
  }

  // ★祝日マスタI列（売上実績平均）をタグ別に自動計算して書き込む
  // → GW・振替休日等の予測基準値を自動更新（手動入力不要）
  {
    const tagSalesMap = {}; // タグ → 売上リスト
    Object.keys(salesMap).forEach(ds => {
      const info = holidayInfoMap[ds];
      if (!info || !info.isOff) return;
      const tag = info.tag;
      const sales = salesMap[ds] ? salesMap[ds].sales : 0;
      if (!sales || sales <= 0) return;
      if (!tagSalesMap[tag]) tagSalesMap[tag] = [];
      tagSalesMap[tag].push(sales);
    });
    const tagAvgMap = {};
    Object.entries(tagSalesMap).forEach(([tag, arr]) => {
      tagAvgMap[tag] = Math.round(arr.reduce((a, b) => a + b, 0) / arr.length);
    });
    console.log('祝日タグ別実績平均: ' + JSON.stringify(tagAvgMap));
    if (hmSheet && hmSheet.getLastRow() > 1) {
      const hmData = hmSheet.getRange(2, 1, hmSheet.getLastRow()-1, 7).getValues();
      const updates = [];
      hmData.forEach((row, i) => {
        if (!row[0] || !row[2]) return; // isOff でない行はスキップ
        const tag = (row[6] || '').toString().trim();
        const newAvg = tagAvgMap[tag];
        if (newAvg && newAvg > 0) updates.push([i + 2, newAvg]);
      });
      updates.forEach(([rowIdx, avg]) => hmSheet.getRange(rowIdx, 9).setValue(avg));
      // holidayInfoMap のsalesAvgも更新し、実績から算出できたタグに hasRealData 印を付ける
      // （＝過去に同じタグの祝日が実際にあった。これが無い祝日＝お盆/山の日等はセルの手入力値を信用せず倍率で見積もる）
      Object.keys(holidayInfoMap).forEach(ds => {
        const info = holidayInfoMap[ds];
        if (!info.isOff) return;
        const newAvg = tagAvgMap[info.tag];
        if (newAvg && newAvg > 0) { info.salesAvg = newAvg; info.hasRealData = true; }
      });
      if (updates.length > 0) console.log('祝日マスタI列を自動更新しました（' + updates.length + '行）');
    }
  }

  // ★実績のない祝日（お盆など）の見積り用：ティア×平日週末別の中央値倍率を過去実績から算出
  const holidayTierMult = calcHolidayTierMultipliers_(salesMap, dowAvgExcludedDates, holidayInfoMap);

  // 【連休の形を学習】過去の連休(≥3連休)から「連休内の各日が、その連休平均に対して何倍だったか」を学習し、
  //   実績のない連休（お盆等）の各日に相対位置(0=初日,1=最終日)で当てはめる。
  //   → 過去の正月・GWで最終日が落ちていれば、お盆の最終日も自動で落ちる（"連休最終日は下がる"を実績から反映）。
  //   売上そのものを連休平均で正規化するので、連休全体の大きさ(magnitude)は変えず、日ごとの"形"だけを与える。
  const chainShapeSamples = []; // {relpos, shape}
  (() => {
    const byChain = {};
    Object.keys(holidayInfoMap).forEach(ds => {
      const hi = holidayInfoMap[ds];
      if (!hi.isOff || !(hi.chainLen >= 3) || !hi.chainDay) return;
      if (new Date(ds) >= today) return;                                  // 過去の連休のみ
      if (excludedDates.has(ds) || !salesMap[ds] || !(salesMap[ds].sales > 0)) return;
      const cid = hi.chainId || ('L' + hi.chainLen + '_' + ds.slice(0, 7));
      (byChain[cid] = byChain[cid] || []).push({ day: hi.chainDay, len: hi.chainLen, sales: salesMap[ds].sales });
    });
    Object.values(byChain).forEach(days => {
      if (days.length < 3) return;                                        // データが3日以上ある連休だけ
      const mean = days.reduce((a, b) => a + b.sales, 0) / days.length;
      if (!(mean > 0)) return;
      days.forEach(x => {
        const relpos = x.len > 1 ? (x.day - 1) / (x.len - 1) : 0;
        chainShapeSamples.push({ relpos, shape: x.sales / mean });        // 連休平均=1に正規化した各日の倍率
      });
    });
  })();
  // 相対位置relposの形を距離加重で補間（サンプルが少なければ穏やかに。無ければ1.0＝影響なし）
  const chainPositionFactor_ = (chainDay, chainLen) => {
    if (chainShapeSamples.length < 3 || !(chainLen >= 3) || !chainDay) return 1.0;
    const relpos = chainLen > 1 ? (chainDay - 1) / (chainLen - 1) : 0;
    let ws = 0, wsum = 0;
    chainShapeSamples.forEach(s => { const w = 1 / (Math.abs(s.relpos - relpos) + 0.12); ws += w * s.shape; wsum += w; });
    return wsum > 0 ? Math.min(1.4, Math.max(0.6, Math.round(ws / wsum * 1000) / 1000)) : 1.0;
  };

  // 天候データ
  const weatherActualMap = {}, weatherForecastMap = {};
  const wxSheet = ss.getSheetByName('天候解析');
  if (wxSheet && wxSheet.getLastRow() > 1) {
    wxSheet.getRange(2, 1, wxSheet.getLastRow()-1, 11).getValues().forEach(row => {
      if (!row[0] || !(row[0] instanceof Date)) return;
      const dateStr = toDateStr_(row[0]);
      const type    = (row[4] || '').toString().trim();
      const wxCat   = (row[6] || '').toString().trim(); // G列：天候区分（row[5]はF列=気温差のため修正）
      const maxTemp = Number(row[2]) || 0;
      if (type === 'actual') weatherActualMap[dateStr] = { wxCat, maxTemp };
      else if (type === 'forecast') weatherForecastMap[dateStr] = { wxCat, maxTemp };
    });
  }

  // 降水量帯マップを事前構築（天候生データのH列から）
  const precipBandMap = {};
  if (wxSheetForPrecip && wxSheetForPrecip.getLastRow() > 1) {
    wxSheetForPrecip.getRange(2, 1, wxSheetForPrecip.getLastRow()-1, 8).getValues().forEach(row => {
      if (!row[0] || row[7] === '' || row[7] === null || row[7] === undefined) return;
      const ds = toDateStr_(row[0]);
      const mm = Number(row[7]);
      precipBandMap[ds] = mm === 0 ? '0' : mm <= 5 ? '1-5' : mm <= 20 ? '5-20' : '20+';
    });
  }
  console.log('降水量帯マップ件数: ' + Object.keys(precipBandMap).length + '件');

  const getWeatherAndTempRate = (dateStr) => {
    const data = weatherActualMap[dateStr] || weatherForecastMap[dateStr];
    const type = weatherActualMap[dateStr] ? 'actual' : (weatherForecastMap[dateStr] ? 'forecast' : 'none');
    if (!data) return { rate: 1.0, type: 'none' };

    // 気温係数
    const tmpRate = getTempRate_(data.maxTemp, tempCoeffMap);
    // 降水量帯係数（優先）
    const band = precipBandMap[dateStr] || null;
    if (band) {
      const precipRate = precipCoeffMap[band] || precipCoeffMap['default'] || 1.0;
      return { rate: Math.round(precipRate * tmpRate * 1000) / 1000, type };
    }
    // 降水量データなし → 天候区分係数にフォールバック
    const wxRate = weatherCoeffMap[data.wxCat] || weatherCoeffMap['default'] || 1.0;
    return { rate: Math.round(wxRate * tmpRate * 1000) / 1000, type };
  };

  // 長期の期間イベント用：期間開始前PRE日間の「曜日別平均」を基準にする（循環しないきれいな基準）。
  //   通常の曜日平均(weightedDowAvg)は分母に自期間が入って倍率が1前後に濁る（夏休みが0.98に見える等）。
  //   開始前の平均と比べれば、夏休みが売上を下げている(≈0.91)といった本当の効果が出る。
  const cleanPreBaseline_ = (startDate) => {
    const PRE = 28;
    const from = new Date(startDate); from.setDate(from.getDate() - PRE); from.setHours(0,0,0,0);
    const acc = {};
    for (const cur = new Date(from); cur < startDate; cur.setDate(cur.getDate() + 1)) {
      const ds = toDateStr_(cur);
      if (salesMap[ds] && !excludedDates.has(ds) && !realHolidayTaggedDates.has(ds) && salesMap[ds].sales > 0) {
        const k = salesMap[ds].dow;
        if (!acc[k]) acc[k] = { s: 0, n: 0 };
        acc[k].s += salesMap[ds].sales; acc[k].n++;
      }
    }
    const base = {};
    Object.keys(acc).forEach(k => { base[k] = acc[k].s / acc[k].n; });
    return base;
  };
  // IQRトリム平均（重なった祝日/CPなどの外れ値に強い。1.5×IQRの外を除外）
  const trimmedMean_ = (arr) => {
    if (arr.length < 4) return arr.reduce((a, b) => a + b, 0) / arr.length;
    const s = [...arr].sort((a, b) => a - b);
    const q1 = s[Math.floor(s.length * 0.25)], q3 = s[Math.floor(s.length * 0.75)];
    const iqr = q3 - q1;
    const f = arr.filter(r => r >= q1 - iqr * 1.5 && r <= q3 + iqr * 1.5);
    const use = f.length >= 2 ? f : arr;
    return use.reduce((a, b) => a + b, 0) / use.length;
  };

  // イベント実績倍率を計算
  const eventActualRates = {};
  eventList.forEach(ev => {
    if (newProductList.some(np => np.name === ev.name)) return;
    const { name, start, end, rate: setRate } = ev;
    const isLong = (end - start) / 86400000 >= LONG_CP_DAYS;
    const useClean = isLong && !ev.isCP;               // 長期の期間イベント（夏休み等）だけ循環しない基準を使う
    const cleanBase = useClean ? cleanPreBaseline_(start) : null;
    const actualRates = [], details = [];
    const cur = new Date(start);
    while (cur <= end) {
      const dateStr = toDateStr_(cur);
      const d = new Date(dateStr);
      // 祝日は祝日効果が混ざりイベント倍率が過大になる（例：宮若花火×海の日で2.6倍）ため学習から除外。
      // 祝日日は予測側で祝日基準値を使うので、イベント倍率は純粋な非祝日の効果だけを学習させる。
      if (d < today && salesMap[dateStr] && !excludedDates.has(dateStr) && !realHolidayTaggedDates.has(dateStr)) {
        const actualSales = salesMap[dateStr].sales, dow = salesMap[dateStr].dow;
        const baseAvg = useClean ? (cleanBase[dow] || 0) : (weightedDowAvg[dow] || 0);
        const wxData = weatherActualMap[dateStr];
        const wxRate = wxData ? (weatherCoeffMap[wxData.wxCat] || 1.0) : 1.0;
        if (baseAvg > 0 && actualSales > 0 && wxRate >= 0.9) {
          const actualRate = actualSales / baseAvg;
          actualRates.push(actualRate);
          details.push({ date: dateStr, dow, actualSales, baseAvg, actualRate, weather: wxRate });
        }
      }
      cur.setDate(cur.getDate() + 1);
    }
    if (actualRates.length > 0) {
      // 長期の期間イベントは外れ値(重なった祝日/CP)に強いIQRトリム平均。短期は従来どおり単純平均。
      const avgActualRate = useClean ? trimmedMean_(actualRates) : actualRates.reduce((a,b) => a+b, 0) / actualRates.length;
      // 実績1件でも採用（年1回イベントはサンプルが1件しか溜まらないため）
      // 1件の場合は実績80% + 設定20%で信頼性を担保しつつ実績を優先
      const adoptedRate = actualRates.length === 1
        ? avgActualRate * 0.8 + setRate * 0.2
        : avgActualRate;
      eventActualRates[name] = { actualRate: Math.round(avgActualRate*1000)/1000, count: actualRates.length, adoptedRate: Math.round(adoptedRate*1000)/1000, setRate, details };
    }
  });

  // 長期CPの経過週別倍率を集計（将来の予測反映用：10サンプル以上で運用開始）
  const cpProgressMap = {}; // { 'w1': [{cpName, date, dow, actualSales, baseAvg, rate, wxRate}, ...], 'w2': [...] }
  eventList.forEach(ev => {
    if (newProductList.some(np => np.name === ev.name)) return;
    const cpDays = (ev.end - ev.start) / (1000 * 60 * 60 * 24);
    if (cpDays < LONG_CP_DAYS) return;
    const cur = new Date(ev.start);
    while (cur <= ev.end) {
      const dateStr = toDateStr_(cur);
      const d = new Date(dateStr);
      // 祝日はCP進行倍率も水増しするため除外（イベント学習と同じ理由）
      if (d < today && salesMap[dateStr] && !excludedDates.has(dateStr) && !realHolidayTaggedDates.has(dateStr)) {
        const actualSales = salesMap[dateStr].sales, dow = salesMap[dateStr].dow;
        const baseAvg = weightedDowAvg[dow] || 0;
        const wxData = weatherActualMap[dateStr];
        const wxRate = wxData ? (weatherCoeffMap[wxData.wxCat] || 1.0) : 1.0;
        if (baseAvg > 0 && actualSales > 0 && wxRate >= 0.9) {
          const elapsedWeek = Math.floor((d - ev.start) / (1000 * 60 * 60 * 24 * 7)) + 1;
          const weekKey = 'w' + elapsedWeek;
          if (!cpProgressMap[weekKey]) cpProgressMap[weekKey] = [];
          // 同一日付は1件だけ登録（複数CP同時進行による重複カウント防止）
          if (!cpProgressMap[weekKey].some(r => r.date === dateStr)) {
            cpProgressMap[weekKey].push({ cpName: ev.name, date: dateStr, dow, actualSales, baseAvg: Math.round(baseAvg), rate: Math.round(actualSales / baseAvg * 1000) / 1000, wxRate });
          }
        }
      }
      cur.setDate(cur.getDate() + 1);
    }
  });

  // CP進行実績シートへ書き込み（蓄積データの可視化）
  let cpProgSheet = ss.getSheetByName('CP進行実績');
  if (!cpProgSheet) cpProgSheet = ss.insertSheet('CP進行実績');
  cpProgSheet.clearContents();
  const cpProgRows = [['経過週','サンプル数','平均倍率','状態','CP名','記録日','曜日','実績売上','曜日平均','実績倍率','天候係数']];
  Object.entries(cpProgressMap).sort((a, b) => Number(a[0].replace('w','')) - Number(b[0].replace('w',''))).forEach(([weekKey, records]) => {
    const avgRate = Math.round(records.reduce((s, r) => s + r.rate, 0) / records.length * 1000) / 1000;
    const status = records.length >= 10 ? '▶ 予測反映中' : '蓄積中(' + records.length + '/10)';
    records.forEach((r, i) => {
      cpProgRows.push([
        i === 0 ? Number(weekKey.replace('w', '')) : '',
        i === 0 ? records.length : '',
        i === 0 ? avgRate : '',
        i === 0 ? status : '',
        r.cpName, r.date, r.dow, r.actualSales, r.baseAvg, r.rate, r.wxRate
      ]);
    });
  });
  if (cpProgRows.length > 1) { cpProgSheet.getRange(1, 1, cpProgRows.length, 11).setValues(cpProgRows); cpProgSheet.setFrozenRows(1); }

  // CP別週別実績（分析・可視化用）- 各CPの経過週ごとの売上倍率推移を個別集計
  const cpByNameMap = {};
  eventList.forEach(ev => {
    if (newProductList.some(np => np.name === ev.name)) return;
    const cpDays = (ev.end - ev.start) / (1000 * 60 * 60 * 24);
    if (cpDays < LONG_CP_DAYS) return;
    if (!cpByNameMap[ev.name]) cpByNameMap[ev.name] = { start: ev.start, end: ev.end, weeks: {} };
    const cur = new Date(ev.start);
    while (cur <= ev.end) {
      const dateStr = toDateStr_(cur);
      const d = new Date(dateStr);
      if (d < today && salesMap[dateStr] && !excludedDates.has(dateStr)) {
        const actualSales = salesMap[dateStr].sales, dow = salesMap[dateStr].dow;
        const baseAvg = weightedDowAvg[dow] || 0;
        const wxData = weatherActualMap[dateStr];
        const wxRate = wxData ? (weatherCoeffMap[wxData.wxCat] || 1.0) : 1.0;
        if (baseAvg > 0 && actualSales > 0 && wxRate >= 0.9) {
          const wk = 'w' + (Math.floor((d - ev.start) / (1000 * 60 * 60 * 24 * 7)) + 1);
          if (!cpByNameMap[ev.name].weeks[wk]) cpByNameMap[ev.name].weeks[wk] = [];
          cpByNameMap[ev.name].weeks[wk].push({ date: dateStr, dow, actualSales, baseAvg: Math.round(baseAvg), rate: Math.round(actualSales / baseAvg * 1000) / 1000 });
        }
      }
      cur.setDate(cur.getDate() + 1);
    }
  });
  let cpByNameSheet = ss.getSheetByName('CP別週別実績');
  if (!cpByNameSheet) cpByNameSheet = ss.insertSheet('CP別週別実績');
  cpByNameSheet.clearContents();
  const cnRows = [['CP名', '経過週', 'サンプル数', '平均倍率', '前週比', '並行CP有']];
  Object.entries(cpByNameMap).sort((a, b) => a[1].start - b[1].start).forEach(([cpName, cpData]) => {
    const weekNums = Object.keys(cpData.weeks).map(w => Number(w.replace('w',''))).sort((a,b) => a-b);
    let prevRate = null;
    weekNums.forEach(wNum => {
      const recs = cpData.weeks['w' + wNum];
      const avgRate = Math.round(recs.reduce((s,r) => s+r.rate, 0) / recs.length * 1000) / 1000;
      const wow = prevRate !== null ? Math.round((avgRate / prevRate - 1) * 1000) / 10 + '%' : '-';
      // 同期間に他の長期CPが走っていたか確認
      const wkStart = new Date(cpData.start.getTime() + (wNum - 1) * 7 * 86400000);
      const wkEnd = new Date(wkStart.getTime() + 6 * 86400000);
      const concurrentCp = eventList.filter(ev2 => {
        if (ev2.name === cpName) return false;
        if (newProductList.some(np => np.name === ev2.name)) return false;
        const d2 = (ev2.end - ev2.start) / (1000 * 60 * 60 * 24);
        if (d2 < LONG_CP_DAYS) return false;
        return ev2.start <= wkEnd && ev2.end >= wkStart;
      }).map(ev2 => ev2.name);
      cnRows.push([cpName, wNum, recs.length, avgRate, wow, concurrentCp.join(' / ')]);
      prevRate = avgRate;
    });
  });
  if (cnRows.length > 1) { cpByNameSheet.getRange(1, 1, cnRows.length, 6).setValues(cnRows); cpByNameSheet.setFrozenRows(1); }

  // 10サンプル以上の経過週倍率のみ予測に反映
  const cpProgressRateMap = {};
  Object.entries(cpProgressMap).forEach(([weekKey, records]) => {
    if (records.length >= 10) {
      cpProgressRateMap[weekKey] = Math.round(records.reduce((s, r) => s + r.rate, 0) / records.length * 1000) / 1000;
    }
  });

  const custList = Object.entries(salesMap).filter(([ds, v]) => !excludedDates.has(ds) && v.customers > 0).map(([,v]) => v.customers);
  const avgCustomers = custList.length > 0 ? custList.reduce((a,b) => a+b, 0) / custList.length : 100;
  const estDailyDishes = Math.ceil(avgCustomers * NEW_PRODUCT_RATE);

  // 【統合】イベント実績シートは廃止。学習結果は「イベント登録」「CP設定」の結果列へ書き戻す。
  //   → 書き戻しは repeatEventCoeffMap / namedEventCoeffMap も揃った後（getEventRate定義の直前）で実行する。

  // 繰り返しイベント（毎月XX日 / 毎年X月Y日）の実績倍率を自動計算
  const repeatEventCoeffMap = {};
  repeatEventList.forEach(ev => {
    const mAnn = ev.repeat.match(/毎年(\d+)月(\d+)/);
    const mMon = !mAnn && ev.repeat.match(/毎月(\d+)/);
    if (!mAnn && !mMon) return;
    const actualRates = [];
    Object.keys(salesMap).sort().forEach(ds => {
      const d = new Date(ds);
      if (mAnn) { if (d.getMonth() !== Number(mAnn[1]) - 1 || d.getDate() !== Number(mAnn[2])) return; }
      else { if (d.getDate() !== Number(mMon[1])) return; }
      if (d >= today) return;
      if (excludedDates.has(ds)) return;
      const { sales, dow } = salesMap[ds];
      if (!sales || sales === 0) return;
      // 祝日と重なった場合：祝日の実績平均を基準に倍率計算（曜日平均でなく祝日水準が正しい比較対象）
      const hInfo = holidayInfoMap[ds];
      const useHolidayBase = hInfo && hInfo.isOff && hInfo.salesAvg > 0;
      const baseAvg = useHolidayBase ? hInfo.salesAvg : (weightedDowAvg[dow] || 0);
      if (baseAvg === 0) return;
      const wxData = weatherActualMap[ds];
      const wxRate = wxData ? (weatherCoeffMap[wxData.wxCat] || 1.0) : 1.0;
      if (wxRate < 0.88) return; // 荒天日は除外
      actualRates.push(sales / baseAvg);
    });
    let adoptedRate;
    if (actualRates.length >= 2) {
      const sorted = [...actualRates].sort((a, b) => a - b);
      const q1 = sorted[Math.floor(sorted.length * 0.25)], q3 = sorted[Math.floor(sorted.length * 0.75)];
      const iqr = q3 - q1;
      const filtered = actualRates.filter(r => r >= q1 - iqr * 1.5 && r <= q3 + iqr * 1.5);
      const use = filtered.length >= 2 ? filtered : actualRates;
      adoptedRate = Math.round(use.reduce((a, b) => a + b, 0) / use.length * 1000) / 1000;
    } else if (ev.rate > 0) {
      adoptedRate = ev.rate;
    } else {
      adoptedRate = 1.0;
    }
    repeatEventCoeffMap[ev.name] = { rate: adoptedRate, count: actualRates.length };
    console.log('繰り返しイベント [' + ev.name + '] 自動計算倍率: ' + adoptedRate + ' (実績' + actualRates.length + '件)');
  });

  // 繰り返しルール由来イベント（イベント登録シートのフラグ「自動」＝固定日・暫定日）の名前別実績倍率を自動計算
  //   同名の過去実施日の実績を集め、2件以上で実績から自動計算。少ないうちは設定倍率（例1.1）をフォールバック。
  //   → 「1.1倍」と手入力しても、実績が貯まれば自動で実績ベースの倍率に置き換わる。
  const namedEventCoeffMap = {};
  const namedEventDates = {};
  eventList.forEach(ev => {
    if (!ev.learnByName) return;
    (namedEventDates[ev.name] = namedEventDates[ev.name] || []).push(ev);
  });
  Object.keys(namedEventDates).forEach(name => {
    const evs = namedEventDates[name];
    const setRate = evs.map(e => Number(e.rate)).find(r => r > 0 && r !== 1.0) || 0; // 手入力の設定倍率（増1.1・減0.9など。1.0=基準は無視）
    const actualRates = [];
    evs.forEach(ev => {
      const d = new Date(ev.start); d.setHours(0,0,0,0);
      if (d >= today) return;                 // 未来（今年・来年の予定）は実績なし
      const ds = toDateStr_(d);
      if (excludedDates.has(ds)) return;
      const sm = salesMap[ds];
      if (!sm || !sm.sales) return;
      const hInfo = holidayInfoMap[ds];
      const useHolidayBase = hInfo && hInfo.isOff && hInfo.salesAvg > 0;
      const baseAvg = useHolidayBase ? hInfo.salesAvg : (weightedDowAvg[sm.dow] || 0);
      if (baseAvg === 0) return;
      const wxData = weatherActualMap[ds];
      const wxRate = wxData ? (weatherCoeffMap[wxData.wxCat] || 1.0) : 1.0;
      if (wxRate < 0.88) return;              // 荒天日は除外
      actualRates.push(sm.sales / baseAvg);
    });
    let adoptedRate;
    if (actualRates.length >= 2) {
      const sorted = [...actualRates].sort((a,b) => a - b);
      const q1 = sorted[Math.floor(sorted.length * 0.25)], q3 = sorted[Math.floor(sorted.length * 0.75)];
      const iqr = q3 - q1;
      const filtered = actualRates.filter(r => r >= q1 - iqr * 1.5 && r <= q3 + iqr * 1.5);
      const use = filtered.length >= 2 ? filtered : actualRates;
      adoptedRate = Math.round(use.reduce((a,b) => a + b, 0) / use.length * 1000) / 1000;
    } else if (actualRates.length === 1 && setRate > 0) {
      adoptedRate = Math.round((actualRates[0] + setRate) / 2 * 1000) / 1000; // 実績1件＋設定倍率 → 平均で少し実績を反映
    } else if (actualRates.length === 1) {
      adoptedRate = Math.round(actualRates[0] * 1000) / 1000;
    } else if (setRate > 0) {
      adoptedRate = setRate;                  // 実績なし → 設定倍率をそのまま使用
    } else {
      adoptedRate = 1.0;
    }
    namedEventCoeffMap[name] = { rate: adoptedRate, count: actualRates.length };
    console.log('繰り返しイベント[名前別] [' + name + '] 採用倍率: ' + adoptedRate + ' (実績' + actualRates.length + '件, 設定' + (setRate || '無') + ')');
  });

  // 【統合】学習結果を「イベント登録」「CP設定」の結果列へ書き戻す（イベント実績シートの後継）
  try { writeEventResultsBack_(ss, eventActualRates, repeatEventCoeffMap, namedEventCoeffMap); }
  catch (e) { console.log('結果の書き戻しでエラー: ' + e.message); }

  // イベント倍率取得
  const getEventRate = (d) => {
    const dt = new Date(d); dt.setHours(0,0,0,0);
    const dayOfMonth = dt.getDate();
    const monthOfYear = dt.getMonth();
    const hits = eventList.filter(ev => dt >= ev.start && dt <= ev.end);
    const repeatHits = repeatEventList.filter(ev => {
      const mAnn = ev.repeat.match(/毎年(\d+)月(\d+)/);
      if (mAnn) return monthOfYear === Number(mAnn[1]) - 1 && dayOfMonth === Number(mAnn[2]);
      const mMon = ev.repeat.match(/毎月(\d+)/);
      return mMon ? dayOfMonth === Number(mMon[1]) : false;
    });
    const allHits = [...hits, ...repeatHits];
    if (allHits.length === 0) return { rate: 1.0, name: '' };
    allHits.sort((a, b) => b.priority - a.priority);
    const top = allHits[0];
    // 長期CPかつ経過週別倍率が10サンプル以上蓄積されていれば予測に反映
    // ただし適用は平日のみ。土日・祝日には掛けない。
    //   理由：経過週別倍率は「過去のCP日の実績 ÷ 現在の指数平滑同曜日平均」で算出しており、
    //   売上が伸びている曜日ほど倍率が1未満に偏る（CPの効果でなく成長トレンドを拾ってしまう）。
    //   この偏りは伸び幅の大きい土日で特に強く出る。
    //   実績22日の検証：log(実績/基準)=λ*log(倍率) の最小二乗推定で
    //     平日 λ=0.81 (t=2.36 有意) → 倍率はほぼそのまま使うのが正しい
    //     週末 λ=-0.71 (t=-2.31)   → 倍率は逆向き。掛けると必ず外れる
    //   比重wのグリッドサーチ＋1件抜き交差検証でも、平日は w=1（現行維持）が最良、
    //   週末は w=0（不適用）で MAPE 17.95%→10.52% と大幅改善した。
    //   ※週末は w を負にすると更に下がるが、それはCPでなく成長トレンドの当てはめ（外挿で破綻）なので採らない。
    //   祝日は実績が無く検証不能だが、基準値が既に祝日水準の見積りであり、
    //   平日から学習した係数を重ねない既存方針（月内パターン・トレンド・biasも同様に不適用）に合わせる。
    const cpDays = top.end && top.start ? (top.end - top.start) / (1000 * 60 * 60 * 24) : 0;
    if (cpDays >= LONG_CP_DAYS) {
      const wd = dt.getDay();                       // 0=日, 6=土
      const hi = holidayInfoMap[toDateStr_(dt)];
      const hTagE = hi ? (hi.tag || '') : '';
      const isRealHolidayE = !!(hi && hi.isOff && hTagE && hTagE !== '土曜' && hTagE !== '日曜' && hTagE !== '通常営業');
      if (wd !== 0 && wd !== 6 && !isRealHolidayE) {
        const elapsedWeek = Math.floor((dt - top.start) / (1000 * 60 * 60 * 24 * 7)) + 1;
        const progressRate = cpProgressRateMap['w' + elapsedWeek];
        if (progressRate !== undefined) return { rate: progressRate, name: top.name };
      }
    }
    // 繰り返しイベントは実績ベースを継続（毎年・毎月で十分なサンプルが蓄積される）
    const repeatInfo = repeatEventCoeffMap[top.name];
    if (repeatInfo) return { rate: repeatInfo.rate, name: top.name };
    // 繰り返しルール由来イベント（固定日・暫定日）は名前別の実績学習倍率を使用（実績が少なければ設定倍率）
    const namedInfo = namedEventCoeffMap[top.name];
    if (namedInfo && top.learnByName) return { rate: namedInfo.rate, name: top.name };
    // 【自動収束】実績が貯まるほど設定倍率を実績へ寄せる（信頼度加重）。
    //   ・短期(<21日)イベント/CP：曜日平均基準で収束。
    //   ・長期の期間イベント(夏休み等・!isCP)：開始前基準で循環を除いた実績に収束（eventActualRatesで算出済み）。
    //   ・長期CP(isCP)だけは循環が残るため対象外＝設定倍率/進行倍率のまま。
    const setRate = top.rate || 1.0;
    const actInfo = eventActualRates[top.name];
    if ((cpDays < LONG_CP_DAYS || !top.isCP) && actInfo && actInfo.count > 0) {
      return { rate: blendEventRate_(setRate, actInfo.actualRate, actInfo.count), name: top.name };
    }
    return { rate: setRate, name: top.name };
  };

  // 既存の「実績確定」行を読み込む。あわせて未来だったときに「届けた予測(biasAdj込み)」も記録する。
  const lockedRowMap = {};
  const deliveredMap = {};   // 予報反映中/予測 の行の予測＝発注に使った値。過去確定時に生の再計算でなくこれを凍結する
  let fSheet = ss.getSheetByName('売上予測');
  if (fSheet && fSheet.getLastRow() > 1) {
    fSheet.getRange(2, 1, fSheet.getLastRow()-1, 12).getValues().forEach(row => {
      if (!row[0]) return;
      const ds = toDateStr_(row[0]);
      if ((row[9] || '').toString() === '実績確定' && Number(row[8]) > 0) lockedRowMap[ds] = row;
      else if (Number(row[8]) > 0) deliveredMap[ds] = Number(row[8]);
    });
  }

  // 天候係数の最終適用強度（1=そのまま / 0=無効）。“薄める”判断は各係数側で自動化した
  // （shrinkWeatherCoeff_：有意な分だけ効かせ、不確かな区分は1.0＝無効。データが貯まれば自動で効き始める）
  // ため、ここは固定で薄めず1.0にする。緊急時に一括で弱めたいときだけこの値を下げる安全弁。
  const WEATHER_WEIGHT = 1.0;
  // 全期間の予測行を生成
  const forecastRows = [];
  const cur2 = new Date(startDate);
  while (cur2 <= endDate) {
    const dateStr = toDateStr_(cur2);
    const dow = cur2.getDay() + 1;
    const dowName = getDowName_(dow);
    const isPast = cur2 < today;

    if (lockedRowMap[dateStr]) {
      forecastRows.push(lockedRowMap[dateStr].slice());
      cur2.setDate(cur2.getDate() + 1);
      continue;
    }
    if (excludedDates.has(dateStr)) { forecastRows.push([new Date(cur2), dow, dowName, '', '', '', '', '', '', '除外', '', '']); cur2.setDate(cur2.getDate()+1); continue; }

    const holidayInfo = holidayInfoMap[dateStr];
    const isOff = holidayInfo ? holidayInfo.isOff : false;
    const hTag = holidayInfo ? (holidayInfo.tag || '') : '';
    // 通常の土日（タグ=土曜/日曜）は祝日基準値ではなく曜日別平均（直近10件の指数平滑）を使う。
    // 真の祝日（GW・お盆・連休・○○の日等）のみ祝日基準値を使用。ただしその基準値が空白(0)の場合は
    // 曜日別平均に戻す（過去実績のない連休で予測=0→発注0になるのを防止）。
    const isRealHoliday = isOff && hTag && hTag !== '土曜' && hTag !== '日曜' && hTag !== '通常営業';
    const hTier = isRealHoliday ? getHolidayTier_(hTag) : '';
    // 連休（3連休等）は「タグに曜日が混在＋冬の少数データで固定」のため凍結した実績平均は使わず、
    // 常に同ティア倍率×直近の指数平滑（＝今の売上水準を反映）で見積もる
    const forceTierRatio = (hTier === '連休');
    let baseAvg = weightedDowAvg[dow] || 0, holidayBase = '';
    if (isRealHoliday) {
      if (!forceTierRatio && holidayInfo.hasRealData && holidayInfo.salesAvg > 0) {
        // 実績のある祝日：過去の同タグ実績の平均を使う（毎年データが貯まればその祝日自身の実績で精度が上がる）
        baseAvg = holidayInfo.salesAvg; holidayBase = Math.round(holidayInfo.salesAvg);
      } else if (weightedDowAvg[dow] > 0) {
        // 実績なし祝日（お盆・山の日等）＆連休：セルの固定値は使わず、同ティア×平日/週末の中央値倍率で見積もる。
        // さらに連休なら「連休何日目か」の形（過去の正月/GWから学習）を掛けて、最終日の落ち込み等を反映する。
        const cls = (dow === 1 || dow === 7) ? '週末' : '平日';
        const mult = holidayTierMult[hTier + '/' + cls];
        if (mult && mult > 0) {
          const posF = chainPositionFactor_(holidayInfo.chainDay, holidayInfo.chainLen);
          baseAvg = Math.round(weightedDowAvg[dow] * mult * posF); holidayBase = baseAvg;
        }
      }
    }

    // ③ 前週アンカー：未来日のみ前週実績を35%ブレンド（祝日・除外日・CP期間はスキップ）
    if (!isPast && !isOff && baseAvg > 0) {
      const lastWeekStr = toDateStr_(new Date(cur2.getTime() - 7 * 86400000));
      if (!avgExcludedDates.has(lastWeekStr) && salesMap[lastWeekStr] && salesMap[lastWeekStr].sales > 0) {
        baseAvg = Math.round(salesMap[lastWeekStr].sales * 0.35 + baseAvg * 0.65);
      }
    }

    const { rate: eventRate }                = getEventRate(cur2);
    const { rate: weatherRateRaw, type: wt } = getWeatherAndTempRate(dateStr);
    // 天候係数は WEATHER_WEIGHT で縮小（現状0＝中立）。1に戻せば従来どおり全適用。
    const weatherRate = 1 + (weatherRateRaw - 1) * WEATHER_WEIGHT;
    const dayOfMonth   = cur2.getDate();
    // 月内パターン係数は土日ではサンプル不足でノイズとなり精度を悪化させるため平日のみ適用（バックテストで確認）。
    // 祝日は基準値自体が既に祝日の水準なので平日向けの係数は掛けない
    const monthPattern = (dow === 1 || dow === 7 || isRealHoliday) ? 1.0 : getMonthPatternCoeff_(dayOfMonth, dow, monthPatternCoeff);
    const postEndCoeff = getCpPostEndCoeff_(dateStr, cpPostEndCoeff);

    const month = cur2.getMonth() + 1;
    const daysAhead = Math.round((cur2 - today) / 86400000);
    const monthCoeff = (!isPast && daysAhead > 70 && monthCoeffMap[month]) ? monthCoeffMap[month] : 1.0;
    // 祝日は基準値が既に祝日水準のためトレンド係数を掛けない（通常日の伸びを二重計上しない）
    const longTrend   = isRealHoliday ? 1.0 : (longTermTrend[dow] || 1.0);
    // 水準ズレ（売上の伸びに基準値が追いつかない遅れ）を補正する。
    //   全曜日まとめた「全体バイアス」(recentBiasCoeff.overall)を使う。
    //   ※曜日別は1曜日あたり件数が少なくブレて、日別MAPEを改善せず週合計でも劣る
    //     （因果検証：全体プール 週MAPE3.9% < 曜日別4.9%、生11.6%）。
    //   過去日（実績確定）→ 補正しない（生の予測。ただし下で“届けた予測”があればそれを凍結する）
    const biasAdj = isPast ? 1.0 : (recentBiasCoeff.overall || 1.0);
    const rawForecast = Math.round(baseAvg * eventRate * weatherRate * monthPattern * postEndCoeff * longTrend * biasAdj * monthCoeff);

    const actualData  = salesMap[dateStr];
    const actualSales = (actualData && actualData.sales > 0) ? actualData.sales : '';
    // 過去日を初めて確定するときは、生の再計算値でなく「その日までに届けていた予測(biasAdj込み)」を凍結する。
    // → 精度表示・週合計が“実際に発注で使った予測”を反映する（凍結後は lockedRowMap で保持され再計算されない）。
    const forecast    = (isPast && actualSales !== '' && deliveredMap[dateStr] > 0) ? deliveredMap[dateStr] : rawForecast;
    const errorRate   = (actualSales !== '' && forecast > 0) ? Math.round((forecast - actualSales) / actualSales * 1000) / 1000 : '';

    let status;
    if (isPast && actualSales !== '')  status = '実績確定';
    else if (wt === 'forecast')        status = '予報反映中';
    else if (!isPast && wt === 'none') status = '天候未定';
    else                               status = '予測';

    const combinedCoeff = Math.round(longTrend * monthPattern * postEndCoeff * 100) / 100;
    forecastRows.push([new Date(cur2), dow, dowName, Math.round(baseAvg), holidayBase, eventRate, weatherRate, combinedCoeff, forecast, status, actualSales, errorRate]);

    cur2.setDate(cur2.getDate() + 1);
  }

  // 売上予測シートに書き込み
  if (!fSheet) fSheet = ss.insertSheet('売上予測');
  fSheet.clearContents();
  fSheet.getRange(1, 1, 1, 12).setValues([['日付','曜日コード','曜日名','曜日別平均売上','祝日基準値','イベント倍率','天候影響度','トレンド係数','予測売上','ステータス','実績売上','誤差率']]);
  fSheet.getRange(2, 1, forecastRows.length, 12).setValues(forecastRows);
  const dr = forecastRows.length;
  fSheet.getRange(2,1,dr,1).setNumberFormat('yyyy/mm/dd');
  fSheet.getRange(2,4,dr,1).setNumberFormat('#,##0');
  fSheet.getRange(2,9,dr,1).setNumberFormat('#,##0');
  fSheet.getRange(2,11,dr,1).setNumberFormat('#,##0');
  fSheet.getRange(2,12,dr,1).setNumberFormat('0.0%');
  fSheet.setFrozenRows(1);

  const statusColorMap = { '実績確定':'#f0f0f0','除外':'#ffe0e0','予報反映中':'#e8f4fd','天候未定':'#fff9e6','予測':'#ffffff' };
  forecastRows.forEach((row, i) => { fSheet.getRange(i+2,1,1,12).setBackground(statusColorMap[row[9]] || '#ffffff'); });

  if (newProductList.length > 0) writeNewProductEstimates_(ss, newProductList, estDailyDishes, today);

  const countOf = s => forecastRows.filter(r => r[9] === s).length;
  console.log('売上予測シートを生成しました（' + forecastRows.length + '日分）\n'
    + '  実績確定: ' + countOf('実績確定') + '日\n  除外: ' + countOf('除外')
    + '日\n  予報反映中: ' + countOf('予報反映中') + '日\n  天候未定: ' + countOf('天候未定')
    + '日\n  予測: ' + countOf('予測') + '日');

  // 天候解析シートを降水量ベースで更新
  updateWeatherAnalysis_(ss, precipCoeffMap, wxSheetForPrecip);
}

// ─────────────────────────────────────────────────────────────────────
// 天候解析シートを降水量ベースで更新
// ─────────────────────────────────────────────────────────────────────

function updateWeatherAnalysis() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const salesMap      = buildSalesMap_(ss);
  const excludedDates = buildExcludedDates_(ss);
  const { coeffs: precipCoeffMap } = calcPrecipCoeff_(ss, salesMap, excludedDates);
  const wxSheetForPrecip  = ss.getSheetByName('天候生データ');
  console.log('降水量帯係数（実績）: ' + JSON.stringify(precipCoeffMap));
  updateWeatherAnalysis_(ss, precipCoeffMap, wxSheetForPrecip);
}

function updateWeatherAnalysis_(ss, precipCoeffMap, wxSheet) {
  const SHEET_NAME = '天候解析';
  const wxaSheet = ss.getSheetByName(SHEET_NAME);
  if (!wxaSheet) { console.log('天候解析シートが見つかりません'); return; }

  const tz = 'Asia/Tokyo';

  // 天候生データから日付→降水量(mm)のマップを作成
  const precipDataMap = {};
  if (wxSheet && wxSheet.getLastRow() > 1) {
    wxSheet.getRange(2, 1, wxSheet.getLastRow()-1, 8).getValues().forEach(row => {
      if (!row[0]) return;
      const ds = row[0] instanceof Date ? Utilities.formatDate(row[0], tz, 'yyyy/MM/dd') : row[0].toString();
      precipDataMap[ds] = (row[7] !== '' && row[7] !== null) ? Number(row[7]) : null;
    });
  }

  // ヘッダーにL列「降水量(mm)」を追加
  if (wxaSheet.getRange(1, 12).getValue() === '') {
    wxaSheet.getRange(1, 12).setValue('降水量(mm)');
  }

  const lastRow = wxaSheet.getLastRow();
  if (lastRow < 2) return;

  const data = wxaSheet.getRange(2, 1, lastRow-1, 11).getValues();

  const updatedK   = [];  // K列（発注影響度）
  const updatedL   = [];  // L列（降水量mm）

  data.forEach(row => {
    if (!row[0]) { updatedK.push([row[10] || '']); updatedL.push(['']); return; }

    const ds         = row[0] instanceof Date ? Utilities.formatDate(row[0], tz, 'yyyy/MM/dd') : row[0].toString();
    const rainFlag   = Number(row[7]) || 0;  // H列：降水フラグ
    const changeFlag = Number(row[8]) || 0;  // I列：急変フラグ
    const precipMm   = precipDataMap[ds];    // 降水量(mm) または null

    let influence;
    if (precipMm === null || precipMm === undefined) {
      // 降水量データなし → 降水フラグで従来通り計算
      if (rainFlag === 0 && changeFlag === 0)      influence = 1.0;
      else if (rainFlag === 0 && changeFlag === 1) influence = 0.9;
      else                                          influence = precipCoeffMap['1-5'] || precipCoeffMap['default'] || 0.9;
    } else if (precipMm === 0) {
      // 雨なし
      influence = changeFlag === 1 ? 0.9 : 1.0;
    } else {
      // 降水量帯で係数決定（実績ベース）
      const band = precipMm <= 5 ? '1-5' : precipMm <= 20 ? '5-20' : '20+';
      influence  = precipCoeffMap[band] || precipCoeffMap['default'] || 1.0;
    }

    updatedK.push([Math.round(influence * 1000) / 1000]);
    updatedL.push([precipMm !== null && precipMm !== undefined ? precipMm : '']);
  });

  wxaSheet.getRange(2, 11, updatedK.length, 1).setValues(updatedK);  // K列
  wxaSheet.getRange(2, 12, updatedL.length, 1).setValues(updatedL);  // L列
  wxaSheet.getRange(2, 11, updatedK.length, 1).setNumberFormat('0.000');
  wxaSheet.getRange(2, 12, updatedL.length, 1).setNumberFormat('0.0');
  // ヘッダーを値で上書き（数式参照ずれによる#REF!を防ぐ）
  wxaSheet.getRange(1, 11).setValue('発注影響度');
  wxaSheet.getRange(1, 12).setValue('降水量(mm)');

  console.log('天候解析シートを更新しました（' + updatedK.length + '行・降水量帯係数適用）');
}

// ─────────────────────────────────────────────────────────────────────
// buildSalesMap_ / buildExcludedDates_ （updateWeatherAnalysis 単独実行用）
// ─────────────────────────────────────────────────────────────────────

function buildSalesMap_(ss) {
  const salesSheet = ss.getSheetByName('日次売上実績');
  const result = {};
  if (!salesSheet || salesSheet.getLastRow() < 2) return result;
  salesSheet.getRange(2, 1, salesSheet.getLastRow()-1, 5).getValues().forEach(row => {
    if (!row[0]) return;
    const ds  = row[0] instanceof Date ? Utilities.formatDate(row[0], 'Asia/Tokyo', 'yyyy/MM/dd') : row[0].toString();
    const dow = Number(row[1]) || 0;
    const sales = Number(row[3]) || 0;
    result[ds] = { dow, sales };
  });
  return result;
}

function buildExcludedDates_(ss) {
  const result = new Set();
  const evSheet = ss.getSheetByName('イベント登録');
  if (!evSheet || evSheet.getLastRow() < 2) return result;
  evSheet.getRange(2, 1, evSheet.getLastRow()-1, 5).getValues().forEach(row => {
    if (!row[0] || (row[4] || '').toString().trim() !== '除外') return;
    const start = row[1] instanceof Date ? row[1] : new Date(row[1]);
    const end   = row[2] instanceof Date ? row[2] : new Date(row[2]);
    if (isNaN(start) || isNaN(end)) return;
    const cur = new Date(start); cur.setHours(0,0,0,0);
    const endD = new Date(end);  endD.setHours(0,0,0,0);
    while (cur <= endD) {
      result.add(Utilities.formatDate(cur, 'Asia/Tokyo', 'yyyy/MM/dd'));
      cur.setDate(cur.getDate() + 1);
    }
  });
  return result;
}

// ─────────────────────────────────────────────────────────────────────
// 曜日別加重平均・各種係数計算
// ─────────────────────────────────────────────────────────────────────

function calcWeightedDowAvg_(salesMap, excludedDates, today) {
  // ② 指数平滑法：曜日ごとに過去を遡り最大10件をα=0.35で指数減衰
  const ALPHA = 0.35;
  const MAX_WEEKS = 10;
  const result = {};
  [1,2,3,4,5,6,7].forEach(dow => {
    let weightedSum = 0, weightedTotal = 0, sampleCount = 0;
    // 最大83日遡って対象曜日のデータを最大10件収集
    for (let daysBack = 1; daysBack <= MAX_WEEKS * 7 + 13; daysBack++) {
      if (sampleCount >= MAX_WEEKS) break;
      const target = new Date(today);
      target.setDate(today.getDate() - daysBack);
      target.setHours(0, 0, 0, 0);
      if (target.getDay() + 1 !== dow) continue; // 対象曜日以外はスキップ
      const ds = toDateStr_(target);
      if (excludedDates.has(ds)) continue;
      if (!salesMap[ds] || salesMap[ds].sales <= 0) continue;
      const weight = ALPHA * Math.pow(1 - ALPHA, sampleCount);
      weightedSum   += salesMap[ds].sales * weight;
      weightedTotal += weight;
      sampleCount++;
    }
    result[dow] = weightedTotal > 0 ? weightedSum / weightedTotal : 0;
  });
  return result;
}

// 指定日(asOf)時点での通常同曜日の指数平滑平均（calcWeightedDowAvg_ と同じ式を単一曜日・任意基準日で）
function expSmoothDowAt_(salesMap, excludedDates, asOf, dow) {
  const ALPHA = 0.35, MAX_WEEKS = 10;
  let weightedSum = 0, weightedTotal = 0, sampleCount = 0;
  for (let daysBack = 1; daysBack <= MAX_WEEKS * 7 + 13; daysBack++) {
    if (sampleCount >= MAX_WEEKS) break;
    const target = new Date(asOf);
    target.setDate(asOf.getDate() - daysBack);
    target.setHours(0, 0, 0, 0);
    if (target.getDay() + 1 !== dow) continue;
    const ds = toDateStr_(target);
    if (excludedDates.has(ds)) continue;
    if (!salesMap[ds] || salesMap[ds].sales <= 0) continue;
    const weight = ALPHA * Math.pow(1 - ALPHA, sampleCount);
    weightedSum   += salesMap[ds].sales * weight;
    weightedTotal += weight;
    sampleCount++;
  }
  return weightedTotal > 0 ? weightedSum / weightedTotal : 0;
}

// 祝日タグをティアに分類（最大＝正月/GW/お盆、連休＝3連休等、単発＝その他の単発祝日）
function getHolidayTier_(tag) {
  const t = (tag || '').toString();
  if (/年始|9連休/.test(t)) return '最大';
  if (/GW|振替/.test(t))    return '最大';
  if (/お盆|4連休/.test(t)) return '最大';
  if (/連休/.test(t))       return '連休';
  return '単発';
}

// 実績のない祝日を見積もるための「ティア×平日/週末」別の中央値倍率を過去実績から算出
// 倍率 = 祝日の実績 ÷ その時点の通常同曜日の指数平滑平均。個々の祝日はブレるため中央値で外れ値に強くする
function calcHolidayTierMultipliers_(salesMap, dowAvgExcludedDates, holidayInfoMap) {
  const ratios = {};
  Object.keys(salesMap).forEach(ds => {
    const info = holidayInfoMap[ds];
    if (!info || !info.isOff) return;
    const tag = info.tag;
    if (!tag || tag === '通常営業' || tag === '土曜' || tag === '日曜') return; // 真の祝日のみ
    const sales = salesMap[ds] ? salesMap[ds].sales : 0;
    if (sales <= 0) return;
    const d = new Date(ds); d.setHours(0, 0, 0, 0);
    const dow = d.getDay() + 1;
    const base = expSmoothDowAt_(salesMap, dowAvgExcludedDates, d, dow);
    if (base <= 0) return;
    const cls = (dow === 1 || dow === 7) ? '週末' : '平日';
    const key = getHolidayTier_(tag) + '/' + cls;
    if (!ratios[key]) ratios[key] = [];
    ratios[key].push(sales / base);
  });
  const median_ = arr => {
    const a = arr.slice().sort((x, y) => x - y);
    const n = a.length;
    if (n === 0) return 0;
    return n % 2 ? a[(n - 1) / 2] : (a[n / 2 - 1] + a[n / 2]) / 2;
  };
  const result = {};
  Object.keys(ratios).forEach(k => { result[k] = median_(ratios[k]); });
  console.log('祝日ティア別倍率(中央値): ' + JSON.stringify(Object.fromEntries(Object.entries(result).map(([k, v]) => [k, Math.round(v * 1000) / 1000]))));
  return result;
}

// イベント登録シートの「繰り返し」列のパターンを、今年・来年分の日付範囲に展開する
// 対応書式：毎年M月D日 ／ 毎年M月D日〜E日 ／ 毎年M月D日から{N}日間（月またぎOK） ／
//           毎年M月第N{曜}（例:毎年7月第3月曜）＋任意で「から{N}日間」 ／「暫定」を含めば要確認フラグ
// 戻り値: [{ start, end, provisional }]
function expandEventRepeat_(pattern, today, endDate) {
  const result = [];
  const pat = (pattern || '').toString().trim();
  if (!pat) return result;
  const DOW = { '日':0,'月':1,'火':2,'水':3,'木':4,'金':5,'土':6 };
  const provisional = /暫定/.test(pat);
  const cutoff = new Date(); cutoff.setDate(cutoff.getDate() - 28); cutoff.setHours(0, 0, 0, 0);
  const mDur = pat.match(/から\s*(\d+)\s*日間/); // 「から43日間」等（月をまたぐ期間指定）
  const durDays = mDur ? Math.max(1, Number(mDur[1])) : null;
  [today.getFullYear(), today.getFullYear() + 1].forEach(year => {
    let sd = null, ed = null, m;
    if ((m = pat.match(/毎年(\d+)月第(\d+)([日月火水木金土])/))) {
      // 第N週の曜日（例：毎年7月第3月曜）
      const month = Number(m[1]) - 1, weekN = Number(m[2]), dow = DOW[m[3]];
      if (isNaN(month) || dow === undefined) return;
      const first = new Date(year, month, 1);
      const diff = (dow - first.getDay() + 7) % 7;
      const nthDay = 1 + diff + (weekN - 1) * 7;
      const test = new Date(year, month, nthDay);
      if (test.getMonth() !== month) return; // その月に存在しない週
      sd = new Date(year, month, nthDay);
    } else if ((m = pat.match(/毎年(\d+)月(\d+)日\s*[〜~]\s*(\d+)日/))) {
      // 同月内の期間（例：毎年8月13日〜16日）
      const month = Number(m[1]) - 1;
      sd = new Date(year, month, Number(m[2])); ed = new Date(year, month, Number(m[3]));
    } else if ((m = pat.match(/毎年(\d+)月(\d+)日/))) {
      // 単日（例：毎年7月15日）
      const month = Number(m[1]) - 1;
      sd = new Date(year, month, Number(m[2]));
    } else { return; }
    if (!sd || isNaN(sd.getTime())) return;
    if (durDays) { ed = new Date(sd); ed.setDate(sd.getDate() + durDays - 1); } // 「から{N}日間」優先（月またぎ対応）
    if (!ed) ed = new Date(sd);
    sd.setHours(0, 0, 0, 0); ed.setHours(0, 0, 0, 0);
    if (ed < cutoff) return;    // 28日以上前に終了 → スキップ
    if (sd > endDate) return;   // 予測期間より先 → スキップ
    result.push({ start: sd, end: ed, provisional });
  });
  return result;
}

// 天候係数を「実績から学習しつつ、統計的に確かな分だけ効かせる」共通処理。
//   ・overallMean（天候以外の全体水準）で割り、天候固有のズレだけを取り出す（季節や伸びの二重計上を防ぐ）
//   ・サンプルが少ない/効果がノイズと区別できない区分は自動で1.0（＝効かせない）
//   ・データが貯まって効果が有意になれば、その分だけ自動で効き始める（固定値と違い頭打ちにならない）
// ratios: その区分の (実績 ÷ 曜日平均) の配列 / overallMean: 全区分をまたいだ ratios の平均。
function shrinkWeatherCoeff_(ratios, overallMean) {
  const n = ratios ? ratios.length : 0;
  if (n < 8 || !(overallMean > 0)) return 1.0;      // データ不足は効かせない（貯まれば自動有効化）
  const rel = ratios.map(r => r / overallMean);      // 天候固有のズレに正規化
  const m   = rel.reduce((a, b) => a + b, 0) / n;
  const sd  = Math.sqrt(rel.reduce((a, b) => a + Math.pow(b - m, 2), 0) / n);
  const se  = sd / Math.sqrt(n);
  if (se <= 0) return Math.round(m * 1000) / 1000;
  const t = Math.abs(m - 1) / se;                    // 効果がノイズの何倍か（t値）
  if (t <= 1) return 1.0;                            // ノイズと区別できない → 効かせない
  const shrink  = 1 - 1 / (t * t);                   // 信頼度で縮小（確かなほど生の効果に近づく）
  const applied = 1 + (m - 1) * shrink;
  return Math.round(Math.min(1.3, Math.max(0.7, applied)) * 1000) / 1000; // 暴走防止クランプ
}

// 全区分をまたいだ「天候以外の水準」= (実績÷曜日平均) の全体平均。天候固有のズレを測る基準。
function overallDowRatio_(bySales, normalAvgByDow) {
  const all = [];
  Object.keys(bySales).forEach(k => bySales[k].forEach(it => {
    const base = normalAvgByDow[it.dow] || 0; if (base > 0) all.push(it.sales / base);
  }));
  return all.length ? all.reduce((a, b) => a + b, 0) / all.length : 1;
}

function calcWeatherCoeff_(ss, salesMap, excludedDates, cpDates) {
  const defaults = { '快晴':1.0,'晴れ':1.0,'曇り':1.0,'小雨':1.0,'雨':1.0,'豪雨':1.0,'雷雨':1.0,'雪':1.0,'大雪':1.0,'default':1.0 };
  const wxSheet = ss.getSheetByName('天候生データ');
  if (!wxSheet || wxSheet.getLastRow() < 2) return { coeffs: defaults, sampleInfo: {} };
  const wxMap = {};
  wxSheet.getRange(2, 1, wxSheet.getLastRow()-1, 7).getValues().forEach(row => {
    if (!row[0] || row[4] !== 'actual') return;
    const ds = row[0] instanceof Date ? Utilities.formatDate(row[0], 'Asia/Tokyo', 'yyyy/MM/dd') : row[0].toString();
    wxMap[ds] = (row[6] || '').toString().trim();
  });
  const today = new Date(); today.setHours(0,0,0,0);
  const normalByDow = {1:[],2:[],3:[],4:[],5:[],6:[],7:[]}, wxBySales = {};
  Object.keys(salesMap).forEach(dateStr => {
    if (excludedDates.has(dateStr)) return;
    if (cpDates && cpDates.has(dateStr)) return; // 長期CP期間を除外（CP効果と天候効果の混在を防ぐ）
    const d = new Date(dateStr); d.setHours(0,0,0,0); if (d >= today) return;
    const { dow, sales } = salesMap[dateStr]; if (sales <= 0) return;
    normalByDow[dow].push(sales);
    const wxCat = wxMap[dateStr];
    if (wxCat) { if (!wxBySales[wxCat]) wxBySales[wxCat] = []; wxBySales[wxCat].push({ dow, sales }); }
  });
  const normalAvgByDow = {};
  [1,2,3,4,5,6,7].forEach(dow => { const arr = normalByDow[dow]; normalAvgByDow[dow] = arr.length > 0 ? arr.reduce((a,b)=>a+b,0)/arr.length : 0; });
  const coeffs = { ...defaults };
  const sampleInfo = {};
  const overallMean = overallDowRatio_(wxBySales, normalAvgByDow);
  Object.keys(wxBySales).forEach(wxCat => {
    const items = wxBySales[wxCat]; if (items.length < 3) return;
    const ratios = items.map(item => { const base = normalAvgByDow[item.dow] || 0; return base > 0 ? item.sales / base : null; }).filter(r => r !== null);
    if (ratios.length >= 3) {
      const avg = ratios.reduce((a,b) => a+b, 0) / ratios.length;
      const stddev = Math.sqrt(ratios.reduce((a,b) => a + Math.pow(b - avg, 2), 0) / ratios.length);
      const applied = shrinkWeatherCoeff_(ratios, overallMean); // 有意な分だけ効かせる（不確かなら1.0）
      coeffs[wxCat] = applied;
      sampleInfo[wxCat] = { coeff: applied, rawCoeff: Math.round(avg*1000)/1000, n: ratios.length, stddev: Math.round(stddev*1000)/1000, min: Math.round(Math.min(...ratios)*1000)/1000, max: Math.round(Math.max(...ratios)*1000)/1000 };
    }
  });
  return { coeffs, sampleInfo };
}

// ─────────────────────────────────────────────────────────────────────
// 降水量帯別係数を実績から算出
// 帯区分: 0mm（晴曇）/ 0超〜5mm（小雨）/ 5〜20mm（雨）/ 20mm超（大雨・豪雨）
// ─────────────────────────────────────────────────────────────────────

function calcPrecipCoeff_(ss, salesMap, excludedDates, cpDates) {
  const defaults = { '0':1.0, '1-5':1.0, '5-20':1.0, '20+':1.0, 'default':1.0 };
  const wxSheet = ss.getSheetByName('天候生データ');
  if (!wxSheet || wxSheet.getLastRow() < 2) return { coeffs: defaults, sampleInfo: {} };

  const precipMap = {};
  wxSheet.getRange(2, 1, wxSheet.getLastRow()-1, 8).getValues().forEach(row => {
    if (!row[0] || row[4] !== 'actual') return;
    const precip = row[7];
    if (precip === '' || precip === null) return;
    const ds = row[0] instanceof Date ? Utilities.formatDate(row[0], 'Asia/Tokyo', 'yyyy/MM/dd') : row[0].toString();
    const mm  = Number(precip);
    const band = mm === 0 ? '0' : mm <= 5 ? '1-5' : mm <= 20 ? '5-20' : '20+';
    precipMap[ds] = band;
  });

  const today = new Date(); today.setHours(0,0,0,0);
  const normalByDow = {1:[],2:[],3:[],4:[],5:[],6:[],7:[]};
  const bandBySales = {};

  Object.keys(salesMap).forEach(dateStr => {
    if (excludedDates.has(dateStr)) return;
    if (cpDates && cpDates.has(dateStr)) return; // 長期CP期間を除外
    const d = new Date(dateStr); d.setHours(0,0,0,0); if (d >= today) return;
    const { dow, sales } = salesMap[dateStr]; if (sales <= 0) return;
    normalByDow[dow].push(sales);
    const band = precipMap[dateStr];
    if (band) { if (!bandBySales[band]) bandBySales[band] = []; bandBySales[band].push({ dow, sales }); }
  });

  const normalAvgByDow = {};
  [1,2,3,4,5,6,7].forEach(dow => {
    const arr = normalByDow[dow];
    normalAvgByDow[dow] = arr.length > 0 ? arr.reduce((a,b)=>a+b,0)/arr.length : 0;
  });

  const coeffs = { ...defaults };
  const sampleInfo = {};
  const overallMean = overallDowRatio_(bandBySales, normalAvgByDow);
  Object.keys(bandBySales).forEach(band => {
    const items = bandBySales[band]; if (items.length < 3) return;
    const ratios = items.map(item => {
      const base = normalAvgByDow[item.dow] || 0;
      return base > 0 ? item.sales / base : null;
    }).filter(r => r !== null);
    if (ratios.length >= 3) {
      const avg = ratios.reduce((a,b) => a+b, 0) / ratios.length;
      const stddev = Math.sqrt(ratios.reduce((a,b) => a + Math.pow(b - avg, 2), 0) / ratios.length);
      const applied = shrinkWeatherCoeff_(ratios, overallMean); // 有意な分だけ効かせる（不確かなら1.0）
      coeffs[band] = applied;
      sampleInfo[band] = { coeff: applied, rawCoeff: Math.round(avg*1000)/1000, n: ratios.length, stddev: Math.round(stddev*1000)/1000, min: Math.round(Math.min(...ratios)*1000)/1000, max: Math.round(Math.max(...ratios)*1000)/1000 };
    }
  });

  return { coeffs, sampleInfo };
}

function getPrecipBand_(dateStr, wxSheet) {
  if (!wxSheet || wxSheet.getLastRow() < 2) return null;
  const rows = wxSheet.getRange(2, 1, wxSheet.getLastRow()-1, 8).getValues();
  for (const row of rows) {
    if (!row[0]) continue;
    const ds = row[0] instanceof Date ? Utilities.formatDate(row[0], 'Asia/Tokyo', 'yyyy/MM/dd') : row[0].toString();
    if (ds !== dateStr) continue;
    const precip = row[7];
    if (precip === '' || precip === null) return null;
    const mm = Number(precip);
    return mm === 0 ? '0' : mm <= 5 ? '1-5' : mm <= 20 ? '5-20' : '20+';
  }
  return null;
}

function calcTempCoeff_(ss, salesMap, excludedDates, cpDates) {
  const defaults = { low:1.0, normal:1.0, high:1.0 };
  const wxSheet = ss.getSheetByName('天候生データ');
  if (!wxSheet || wxSheet.getLastRow() < 2) return { coeffs: defaults, sampleInfo: {} };
  const tempMap = {};
  wxSheet.getRange(2, 1, wxSheet.getLastRow()-1, 5).getValues().forEach(row => {
    if (!row[0] || row[4] !== 'actual') return;
    const ds = row[0] instanceof Date ? Utilities.formatDate(row[0], 'Asia/Tokyo', 'yyyy/MM/dd') : row[0].toString();
    tempMap[ds] = Number(row[2]) || 0;
  });
  const today = new Date(); today.setHours(0,0,0,0);
  const normalByDow = {1:[],2:[],3:[],4:[],5:[],6:[],7:[]}, tempBands = {low:[],normal:[],high:[]};
  Object.keys(salesMap).forEach(dateStr => {
    if (excludedDates.has(dateStr)) return;
    if (cpDates && cpDates.has(dateStr)) return; // 長期CP期間を除外
    const d = new Date(dateStr); d.setHours(0,0,0,0); if (d >= today) return;
    const { dow, sales } = salesMap[dateStr]; if (sales <= 0) return;
    normalByDow[dow].push(sales);
    const temp = tempMap[dateStr];
    if (temp !== undefined) { const band = temp <= 14 ? 'low' : temp <= 24 ? 'normal' : 'high'; tempBands[band].push({ dow, sales }); }
  });
  const normalAvgByDow = {};
  [1,2,3,4,5,6,7].forEach(dow => { const arr = normalByDow[dow]; normalAvgByDow[dow] = arr.reduce((a,b)=>a+b,0)/arr.length || 0; });
  const coeffs = { ...defaults };
  const sampleInfo = {};
  const overallMean = overallDowRatio_(tempBands, normalAvgByDow);
  Object.keys(tempBands).forEach(band => {
    const items = tempBands[band]; if (items.length < 5) return;
    const ratios = items.map(item => { const base = normalAvgByDow[item.dow] || 0; return base > 0 ? item.sales / base : null; }).filter(r => r !== null);
    if (ratios.length >= 5) {
      const avg = ratios.reduce((a,b) => a+b, 0) / ratios.length;
      const stddev = Math.sqrt(ratios.reduce((a,b) => a + Math.pow(b - avg, 2), 0) / ratios.length);
      const applied = shrinkWeatherCoeff_(ratios, overallMean); // 有意な分だけ効かせる（不確かなら1.0）
      coeffs[band] = applied;
      sampleInfo[band] = { coeff: applied, rawCoeff: Math.round(avg*1000)/1000, n: ratios.length, stddev: Math.round(stddev*1000)/1000, min: Math.round(Math.min(...ratios)*1000)/1000, max: Math.round(Math.max(...ratios)*1000)/1000 };
    }
  });
  return { coeffs, sampleInfo };
}

function getTempRate_(maxTemp, tempCoeffMap) {
  if (maxTemp <= 14) return tempCoeffMap.low || 1.0;
  if (maxTemp <= 24) return tempCoeffMap.normal || 1.0;
  return tempCoeffMap.high || 1.0;
}

// 天候係数逆算結果を「天候係数逆算」シートに出力
function writeWeatherCoeffAnalysis_(ss, wxSampleInfo, precipCoeffs, tempCoeffs) {
  let sheet = ss.getSheetByName('天候係数逆算');
  if (!sheet) sheet = ss.insertSheet('天候係数逆算');
  sheet.clearContents();

  const rows = [['天候係数逆算（CP期間除外・通常日のみの実績から算出）', '', '', '', '', '']];

  rows.push(['', '', '', '', '', '']);
  rows.push(['■ 天候区分別係数', '', '', '', '', '']);
  rows.push(['天候区分', '逆算係数', 'サンプル数', '標準偏差', '最小値', '最大値']);
  rows.push(['※ 1.0より大きい = 売上UP（モール内効果）', '1.0より小さい = 売上DOWN', '', '', '', '']);
  ['快晴','晴れ','曇り','小雨','雨','豪雨','雷雨','雪','大雪'].forEach(wxCat => {
    const info = (wxSampleInfo || {})[wxCat];
    if (info) rows.push([wxCat, info.coeff, info.n + '件', info.stddev, info.min, info.max]);
    else      rows.push([wxCat, '（データ不足）', '-', '-', '-', '-']);
  });

  rows.push(['', '', '', '', '', '']);
  rows.push(['■ 降水量帯別係数', '', '', '', '', '']);
  rows.push(['降水量帯', '逆算係数', '', '', '', '']);
  const precipLabels = {'0':'0mm（降雨なし）','1-5':'0〜5mm（小雨）','5-20':'5〜20mm（雨）','20+':'20mm超（大雨）'};
  ['0','1-5','5-20','20+'].forEach(band => {
    const coeff = (precipCoeffs || {})[band];
    rows.push([precipLabels[band], coeff != null ? coeff : '（データ不足）', '', '', '', '']);
  });

  rows.push(['', '', '', '', '', '']);
  rows.push(['■ 気温帯別係数', '', '', '', '', '']);
  rows.push(['気温帯', '逆算係数', '', '', '', '']);
  [['low','低温（〜14℃）'],['normal','通常（15〜24℃）'],['high','高温（25℃〜）']].forEach(([k, label]) => {
    const coeff = (tempCoeffs || {})[k];
    rows.push([label, coeff != null ? coeff : '（データ不足）', '', '', '', '']);
  });

  rows.push(['', '', '', '', '', '']);
  rows.push(['更新日時：' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm'), '', '', '', '', '']);

  sheet.getRange(1, 1, rows.length, 6).setValues(rows);
  sheet.setFrozenRows(1);
  console.log('天候係数逆算シート更新完了');
}

function median_(arr) {
  if (!arr || arr.length === 0) return 1.0;
  const sorted = [...arr].sort((a, b) => a - b);
  const mid = Math.floor(sorted.length / 2);
  return sorted.length % 2 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;
}

function calcRecentBiasCoeff_(salesMap, excludedDates, ss) {
  // ④ 曜日別誤差補正：28日（各曜日4件確保）＋EWMA重み付け（直近誤差を重視）
  const WINDOW = 28;
  const today = new Date(); today.setHours(0, 0, 0, 0);

  // 売上予測シートから既存の予測値を読み込む
  const forecastMap = {};
  const fcSheet = ss.getSheetByName('売上予測');
  if (fcSheet && fcSheet.getLastRow() > 1) {
    fcSheet.getRange(2, 1, fcSheet.getLastRow() - 1, 9).getValues().forEach(row => {
      if (!row[0] || !row[8]) return;
      const ds = toDateStr_(row[0]);
      forecastMap[ds] = Number(row[8]); // I列：予測売上
    });
  }

  // { ratio, weight } 形式で収集（i=1が最新、EWMA減衰率0.7）
  const dowItems = { 1:[], 2:[], 3:[], 4:[], 5:[], 6:[], 7:[] };
  const allItems = [];

  for (let i = 1; i <= WINDOW; i++) {
    const d = new Date(today); d.setDate(today.getDate() - i);
    const ds = toDateStr_(d);
    if (excludedDates.has(ds)) continue;
    const actual    = salesMap[ds] ? salesMap[ds].sales : 0;
    const predicted = forecastMap[ds] || 0;
    if (actual <= 0 || predicted <= 0) continue;
    const ratio  = actual / predicted;
    const weight = Math.pow(0.7, i - 1); // 直近ほど重い
    const dow = salesMap[ds].dow;
    dowItems[dow].push({ ratio, weight });
    allItems.push({ ratio, weight });
  }

  const ewmaMean_ = items => {
    if (!items || items.length === 0) return 1.0;
    const totalW = items.reduce((s, x) => s + x.weight, 0);
    return items.reduce((s, x) => s + x.ratio * x.weight, 0) / totalW;
  };

  const overallMean = ewmaMean_(allItems);
  const result = {};
  for (let dow = 1; dow <= 7; dow++) {
    const items = dowItems[dow];
    const raw = items.length >= 2 ? ewmaMean_(items) : overallMean;
    result[dow] = Math.min(1.25, Math.max(0.75, Math.round((raw || 1.0) * 1000) / 1000));
  }
  // 全体バイアス（土日・祝日用）。売上水準の伸び遅れを補正する。振れ幅は平日より狭く抑える。
  result.overall = Math.min(1.15, Math.max(0.90, Math.round((overallMean || 1.0) * 1000) / 1000));
  console.log('誤差率バイアス補正係数: ' + JSON.stringify(result));
  return result;
}

function calcMonthPatternCoeff_(salesMap, excludedDates) {
  const today = new Date(); today.setHours(0,0,0,0);
  const PERIODS = {
    early: d => d <= 7,
    mid:   d => d >= 8 && d <= 23,
    late:  d => d >= 24
  };
  const byPeriod = {};
  Object.keys(PERIODS).forEach(p => {
    byPeriod[p] = {};
    [1,2,3,4,5,6,7].forEach(dow => { byPeriod[p][dow] = []; });
  });
  Object.keys(salesMap).forEach(dateStr => {
    if (excludedDates.has(dateStr)) return;
    const d = new Date(dateStr); d.setHours(0,0,0,0); if (d >= today) return;
    const { dow, sales } = salesMap[dateStr]; if (sales <= 0) return;
    const day = d.getDate();
    const period = Object.keys(PERIODS).find(p => PERIODS[p](day));
    if (period) byPeriod[period][dow].push(sales);
  });
  const allByDow = {};
  [1,2,3,4,5,6,7].forEach(dow => {
    const all = Object.keys(PERIODS).flatMap(p => byPeriod[p][dow]);
    allByDow[dow] = all.length > 0 ? all.reduce((a,b) => a+b, 0) / all.length : 0;
  });
  const result = {};
  Object.keys(PERIODS).forEach(period => {
    result[period] = {};
    [1,2,3,4,5,6,7].forEach(dow => {
      const arr  = byPeriod[period][dow];
      const base = allByDow[dow];
      if (arr.length >= 3 && base > 0) {
        const coeff = arr.reduce((a,b) => a+b, 0) / arr.length / base;
        result[period][dow] = Math.min(1.2, Math.max(0.8, Math.round(coeff * 1000) / 1000));
      } else {
        result[period][dow] = 1.0;
      }
    });
  });
  return result;
}

function calcLongTermTrend_(salesMap, excludedDates, today) {
  const RECENT_WEEKS = 4, LONG_WEEKS = 12;
  const recentLimit = new Date(today); recentLimit.setDate(today.getDate() - RECENT_WEEKS * 7);
  const longLimit   = new Date(today); longLimit.setDate(today.getDate() - LONG_WEEKS * 7);
  const result = {};
  [1,2,3,4,5,6,7].forEach(dow => {
    const recentSales = [], longSales = [];
    Object.keys(salesMap).forEach(dateStr => {
      if (excludedDates.has(dateStr)) return;
      const d = new Date(dateStr); d.setHours(0,0,0,0);
      if (d >= today) return;
      if (salesMap[dateStr].dow !== dow || salesMap[dateStr].sales <= 0) return;
      if (d >= recentLimit) recentSales.push(salesMap[dateStr].sales);
      if (d >= longLimit)   longSales.push(salesMap[dateStr].sales);
    });
    if (recentSales.length < 2 || longSales.length < 4) {
      result[dow] = 1.0; return;
    }
    const recentAvg = recentSales.reduce((a,b) => a+b, 0) / recentSales.length;
    const longAvg   = longSales.reduce((a,b)   => a+b, 0) / longSales.length;
    result[dow] = Math.min(1.3, Math.max(0.7, Math.round(recentAvg / longAvg * 1000) / 1000));
  });
  console.log('長期トレンド係数（曜日別）: ' + JSON.stringify(result));
  return result;
}

function getMonthPatternCoeff_(dayOfMonth, dow, monthPatternCoeff) {
  const period = dayOfMonth <= 7 ? 'early' : dayOfMonth <= 23 ? 'mid' : 'late';
  return (monthPatternCoeff[period] && monthPatternCoeff[period][dow]) || 1.0;
}

// 月別補正係数（1〜12月それぞれの売上水準を全体平均と比較）
function calcMonthCoeff_(salesMap, excludedDates) {
  const today = new Date(); today.setHours(0,0,0,0);
  const monthSales = {};
  for (let m = 1; m <= 12; m++) monthSales[m] = [];

  Object.keys(salesMap).forEach(dateStr => {
    if (excludedDates.has(dateStr)) return;
    const d = new Date(dateStr); d.setHours(0,0,0,0);
    if (d >= today) return;
    const { sales } = salesMap[dateStr];
    if (sales <= 0) return;
    const month = d.getMonth() + 1;
    monthSales[month].push(sales);
  });

  const allSales = Object.values(monthSales).flat();
  if (allSales.length === 0) return {};
  const overallAvg = allSales.reduce((a,b) => a+b, 0) / allSales.length;

  const result = {};
  for (let m = 1; m <= 12; m++) {
    const arr = monthSales[m];
    if (arr.length >= 20 && overallAvg > 0) {
      const monthAvg = arr.reduce((a,b) => a+b, 0) / arr.length;
      result[m] = Math.min(1.3, Math.max(0.7, Math.round(monthAvg / overallAvg * 1000) / 1000));
    } else {
      result[m] = 1.0;
    }
  }
  console.log('月別補正係数: ' + JSON.stringify(result));
  return result;
}

// CP設定の補正倍率を実績から更新（手動実行用）
// ※buildSalesForecastからは呼ばない。単独でupdateCpRateFromActuals()を実行する。
function updateCpRateFromActuals() {
  updateCpRateFromActuals_(SpreadsheetApp.getActiveSpreadsheet());
}

function updateCpRateFromActuals_(ss) {
  const fcSheet = ss.getSheetByName('売上予測');
  const cpSheet = ss.getSheetByName('CP設定');
  if (!fcSheet || !cpSheet || fcSheet.getLastRow() < 2 || cpSheet.getLastRow() < 2) return;

  // 売上予測シートから実績確定データを収集
  // A:日付, D:baseAvg, F:eventRate(現在のCP倍率), G:weatherRate, H:combinedCoeff, K:実績売上
  const actualMap = {};
  fcSheet.getRange(2, 1, fcSheet.getLastRow()-1, 11).getValues().forEach(row => {
    if ((row[9] || '').toString() !== '実績確定') return;
    const baseAvg     = Number(row[3]);
    const eventRate   = Number(row[5]) || 1.0;
    const weatherRate = Number(row[6]) || 1.0;
    const combined    = Number(row[7]) || 1.0; // longTrend × monthPattern × postEndCoeff
    const actual      = Number(row[10]);
    if (baseAvg <= 0 || actual <= 0) return;
    // 他の係数を除いた純粋なCP効果: actual / (baseAvg × weatherRate × combined)
    const otherFactors = weatherRate * combined;
    const pureRate = otherFactors > 0 ? actual / (baseAvg * otherFactors) : actual / baseAvg;
    actualMap[toDateStr_(row[0])] = { pureRate, currentRate: eventRate };
  });

  const today = new Date(); today.setHours(0,0,0,0);
  const cpData = cpSheet.getRange(2, 1, cpSheet.getLastRow()-1, 8).getValues();
  let updated = 0;
  cpData.forEach((row, i) => {
    if (!row[0] || (row[7] || '').toString().trim() === '除外') return;
    const s = row[1] instanceof Date ? row[1] : new Date(row[1]);
    const e = row[2] instanceof Date ? row[2] : new Date(row[2]);
    if (!s || !e || isNaN(s.getTime())) return;
    s.setHours(0,0,0,0); e.setHours(0,0,0,0);
    // 終了済みのCPのみ更新（進行中は除外）
    if (e >= today) return;

    const rates = [];
    const cur = new Date(s);
    while (cur <= e) {
      const ds = toDateStr_(cur);
      if (actualMap[ds]) rates.push(actualMap[ds].pureRate);
      cur.setDate(cur.getDate() + 1);
    }
    if (rates.length < 3) return;

    const calcRate = rates.reduce((a,b) => a+b, 0) / rates.length;
    const currentRate = Number(row[5]) || 1.0;
    // 急激な変更を防ぐため新旧を50%ブレンドし1.0〜2.5の範囲に収める
    const blended = currentRate * 0.5 + calcRate * 0.5;
    const clamped = Math.min(2.5, Math.max(1.0, Math.round(blended * 100) / 100));
    cpSheet.getRange(i + 2, 6).setValue(clamped);
    console.log('CP倍率更新: ' + row[0] + ' → ' + clamped + '（実績' + rates.length + '日, 計算値' + Math.round(calcRate*100)/100 + '）');
    updated++;
  });

  console.log('CP倍率更新完了: ' + updated + '件');
  ss.toast('CP倍率を' + updated + '件更新しました', '更新完了', 5);
}

function calcCpPostEndCoeff_(eventList, salesMap, excludedDates, today) {
  const result = {};
  eventList.forEach(ev => {
    if (ev.end >= today || ev.rate <= 1.0) return;
    const prePeriodEnd = new Date(ev.start); prePeriodEnd.setDate(prePeriodEnd.getDate()-1);
    const prePeriodStart = new Date(prePeriodEnd); prePeriodStart.setDate(prePeriodEnd.getDate()-27);
    const preSalesByDow = {1:[],2:[],3:[],4:[],5:[],6:[],7:[]};
    Object.keys(salesMap).forEach(dateStr => {
      if (excludedDates.has(dateStr)) return;
      const d = new Date(dateStr); d.setHours(0,0,0,0);
      if (d < prePeriodStart || d > prePeriodEnd) return;
      const { dow, sales } = salesMap[dateStr]; if (sales > 0) preSalesByDow[dow].push(sales);
    });
    const preAvgByDow = {};
    [1,2,3,4,5,6,7].forEach(dow => { const arr = preSalesByDow[dow]; preAvgByDow[dow] = arr.length > 0 ? arr.reduce((a,b)=>a+b,0)/arr.length : 0; });
    const cpSalesByDow = {1:[],2:[],3:[],4:[],5:[],6:[],7:[]};
    Object.keys(salesMap).forEach(dateStr => {
      if (excludedDates.has(dateStr)) return;
      const d = new Date(dateStr); d.setHours(0,0,0,0);
      if (d < ev.start || d > ev.end) return;
      const { dow, sales } = salesMap[dateStr]; if (sales > 0) cpSalesByDow[dow].push(sales);
    });
    let totalLift = 0, liftCount = 0;
    [1,2,3,4,5,6,7].forEach(dow => {
      const cpArr = cpSalesByDow[dow], preAvg = preAvgByDow[dow];
      if (cpArr.length > 0 && preAvg > 0) { totalLift += (cpArr.reduce((a,b)=>a+b,0)/cpArr.length) / preAvg; liftCount++; }
    });
    if (liftCount === 0) return;
    const avgLift = totalLift / liftCount; if (avgLift < 1.05) return;
    const c1 = Math.round((1/avgLift)*1000)/1000, c2 = Math.round(Math.pow(1/avgLift,0.5)*1000)/1000;
    console.log('CP「'+ev.name+'」終了後補正: 上乗せ倍率='+Math.round(avgLift*100)/100+' → 1週目×'+c1+' 2週目×'+c2);
    for (let i = 1; i <= 14; i++) {
      const d = new Date(ev.end); d.setDate(d.getDate()+i); d.setHours(0,0,0,0);
      if (d > today) { const ds = Utilities.formatDate(d,'Asia/Tokyo','yyyy/MM/dd'); const coeff = i<=7?c1:c2; result[ds] = result[ds] ? Math.min(result[ds],coeff) : coeff; }
    }
  });
  return result;
}

function getCpPostEndCoeff_(dateStr, cpPostEndCoeff) { return cpPostEndCoeff[dateStr] || 1.0; }

function writeNewProductEstimates_(ss, newProductList, estDailyDishes, today) {
  // 食材マスタから仕入規格量・歩留率を取得（buildIngredientDailyUsage_と同じロジック）
  const imSheet = ss.getSheetByName('食材マスタ');
  const ingSpecMap = {};
  if (imSheet && imSheet.getLastRow() > 1) {
    imSheet.getRange(2, 1, imSheet.getLastRow()-1, 8).getValues().forEach(row => {
      if (!row[0]) return;
      ingSpecMap[String(Number(row[0]))] = {
        spec:  Number(row[5]) > 0 ? Number(row[5]) : 0,  // 0 = F列未入力マーカー
        yield: Number(row[7]) > 0 ? Number(row[7]) : 1
      };
    });
  }
  const pmSheet = ss.getSheetByName('商品マスタ');
  if (!pmSheet || pmSheet.getLastRow() < 2) { console.log('商品マスタが見つかりません'); return; }
  const pmData = pmSheet.getRange(2, 1, pmSheet.getLastRow()-1, 15).getValues().filter(r => r[0]);
  const productRecipeMap = {};
  const specSkipped = [];
  pmData.forEach(row => {
    const prodCode = Number(row[0]); if (!prodCode || !row[4]) return;
    const ingCode = String(Number(row[4]));
    const ingSpec = ingSpecMap[ingCode];
    if (!ingSpec || ingSpec.spec <= 0) {
      specSkipped.push('食材コード ' + ingCode + '（' + row[5] + '）: 食材マスタF列（仕入規格量）が未入力 → 推定行スキップ');
      return;
    }
    if (!productRecipeMap[prodCode]) productRecipeMap[prodCode] = [];
    productRecipeMap[prodCode].push({ ingCode:row[4], ingName:row[5], specQty:ingSpec.spec, yieldRate:ingSpec.yield, useQty:Number(row[12])||0 });
  });
  if (specSkipped.length > 0) {
    console.warn('【食材マスタF列 未入力のためスキップ】\n' + specSkipped.join('\n') + '\n→ 食材マスタのF列（仕入規格量）を入力してください');
  }
  const ingSheet = ss.getSheetByName('食材日時実績');
  if (!ingSheet) { console.log('食材日時実績シートが見つかりません'); return; }
  const existingData = ingSheet.getLastRow() > 1 ? ingSheet.getRange(2, 1, ingSheet.getLastRow()-1, 6).getValues() : [];
  const actualKeys = new Set();
  existingData.forEach(row => { if (row[0] && row[2] && (!row[5] || row[5].toString().trim() === '')) actualKeys.add(toDateStr_(row[0]) + '_' + row[2]); });
  const rowsToKeep = existingData.filter(row => row[0] && (row[5]||'').toString().trim() !== '推定');
  const newEstRows = [], warnings = [];
  newProductList.forEach(np => {
    const { name, start, end, productCode } = np;
    if (!productCode) { warnings.push('「'+name+'」: F列に商品コードが未入力'); return; }
    const recipe = productRecipeMap[productCode];
    if (!recipe) { warnings.push('「'+name+'」: 商品コード'+productCode+'が商品マスタに未登録'); return; }
    const writeStart = new Date(start); writeStart.setDate(writeStart.getDate() - WRITE_START_DAYS);
    const writeCur = new Date(writeStart);
    while (writeCur <= end) {
      const dateStr = toDateStr_(writeCur), dow = writeCur.getDay() + 1;
      recipe.forEach(ing => {
        const key = dateStr + '_' + ing.ingCode; if (actualKeys.has(key)) return;
        const estQty = estDailyDishes * ing.useQty / (ing.specQty || 1) / (ing.yieldRate || 1) * NEW_PRODUCT_SAFETY;
        newEstRows.push([new Date(writeCur), dow, ing.ingCode, ing.ingName, Math.round(estQty*10)/10, '推定']);
      });
      writeCur.setDate(writeCur.getDate() + 1);
    }
  });
  ingSheet.clearContents();
  ingSheet.getRange(1, 1, 1, 6).setValues([['日付','曜日','食材コード','食材名','使用量','フラグ']]);
  const allRows = [...rowsToKeep, ...newEstRows].filter(r => r[0]).sort((a, b) => {
    const da = a[0] instanceof Date ? a[0] : new Date(a[0]), db = b[0] instanceof Date ? b[0] : new Date(b[0]);
    return da - db || String(a[2]).localeCompare(String(b[2]));
  });
  if (allRows.length > 0) {
    ingSheet.getRange(2, 1, allRows.length, 6).setValues(allRows);
    ingSheet.getRange(2, 1, allRows.length, 1).setNumberFormat('yyyy/mm/dd');
    allRows.forEach((row, i) => { if ((row[5]||'').toString().trim() === '推定') ingSheet.getRange(i+2,1,1,6).setBackground('#fff9e6'); });
  }
  if (warnings.length > 0) { try { SpreadsheetApp.getUi().alert('新商品設定の警告', warnings.join('\n'), SpreadsheetApp.getUi().ButtonSet.OK); } catch(e) { console.log(warnings.join('\n')); } }
  console.log('新商品推定行を書き込みました（'+newEstRows.length+'行）');
}


// ─────────────────────────────────────────────────────────────────────
// 新商品推定行の強制再計算（14日制限なし・手動実行用）
// buildSalesForecastの14日制限で更新されない場合に使う
// ─────────────────────────────────────────────────────────────────────
function 新商品推定行再計算() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const today = new Date(); today.setHours(0,0,0,0);
  const newProductList = [];

  // イベント登録シートから「新商品」タイプを取得
  const evSheet = ss.getSheetByName('イベント登録');
  if (evSheet && evSheet.getLastRow() > 1) {
    evSheet.getRange(2, 1, evSheet.getLastRow()-1, 10).getValues().forEach(row => {
      if (!row[0] || (row[4]||'').toString().trim() !== '新商品') return;
      const start = toDate_(row[1]), end = toDate_(row[2]);
      if (!start || !end) return;
      const productCode = row[5] ? Number(row[5]) : null;
      newProductList.push({ name: row[0], start, end, productCode, priority: 0 });
    });
  }

  // CP設定から全CP（14日制限なし）の新商品を取得
  const cpSheet = ss.getSheetByName('CP設定');
  const pmSh = ss.getSheetByName('商品マスタ');
  if (cpSheet && cpSheet.getLastRow() > 1 && pmSh && pmSh.getLastRow() > 1) {
    const cpProdIngMap = {};
    pmSh.getRange(2, 1, pmSh.getLastRow()-1, 5).getValues()
      .filter(r => r[0] && r[4])
      .forEach(r => {
        const mc = String(Math.round(Number(r[0])));
        if (!cpProdIngMap[mc]) cpProdIngMap[mc] = [];
        cpProdIngMap[mc].push(String(r[4]));
      });
    const ingJissekiSh = ss.getSheetByName('食材日時実績');
    const existIngSet = new Set();
    if (ingJissekiSh && ingJissekiSh.getLastRow() > 1) {
      ingJissekiSh.getRange(2, 1, ingJissekiSh.getLastRow()-1, 6).getValues()
        .filter(r => r[2] && (!r[5] || r[5].toString().trim() !== '推定'))
        .forEach(r => existIngSet.add(String(r[2])));
    }
    cpSheet.getRange(2, 1, cpSheet.getLastRow()-1, 9).getValues().forEach(row => {
      const cpName = (row[0]||'').toString().trim(); if (!cpName) return;
      if ((row[8] || '').toString().trim() === '休止') return; // CP休止行は新商品推定の対象外
      const start = toDate_(row[1]), end = toDate_(row[2]);
      if (!start || !end) return;
      const rawVal = row[3];
      const menuCodes = typeof rawVal === 'number'
        ? [String(Math.round(rawVal))]
        : String(rawVal||'').split(',').map(c=>c.trim()).filter(c=>c);
      menuCodes.forEach(mc => {
        const ingCodes = cpProdIngMap[mc] || [];
        if (ingCodes.length === 0 || !ingCodes.some(ic => !existIngSet.has(ic))) return;
        const entryName = cpName + '（商品' + mc + '）';
        if (!newProductList.some(np => np.name === entryName)) {
          newProductList.push({ name: entryName, start, end, productCode: Number(mc), priority: 0 });
          console.log('対象CP新商品: ' + entryName);
        }
      });
    });
  }

  if (newProductList.length === 0) { console.log('新商品が見つかりません（CP設定・イベント登録を確認してください）'); return; }

  const salesMap = buildSalesMap_(ss);
  const excludedDates = buildExcludedDates_(ss);
  const custList = Object.entries(salesMap).filter(([ds, v]) => !excludedDates.has(ds) && v.customers > 0).map(([,v]) => v.customers);
  const avgCustomers = custList.length > 0 ? custList.reduce((a,b) => a+b, 0) / custList.length : 100;
  const estDailyDishes = Math.ceil(avgCustomers * NEW_PRODUCT_RATE);
  console.log('平均客数: ' + Math.round(avgCustomers) + '人, 推定皿数: ' + estDailyDishes + '皿');

  writeNewProductEstimates_(ss, newProductList, estDailyDishes, today);
  console.log('完了。次に「発注バッファ_ステップ2」を実行してください。');
}


// ═══════════════════════════════════════════════════════════════════════
// 【F】 発注処理セクション
// ═══════════════════════════════════════════════════════════════════════

function 発注計算バッファ_完全生成(startDateOverride) {
  const SHEET_NAME = '発注計算バッファ', START_ROW = 2;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) { console.log('「発注計算バッファ」シートが見つかりません'); return; }

  applyFoodCodeSwitches_();

  // キャッシュからcpSubtractMapを読む（なければ再計算）
  let cpSubtractMap = loadCpSubtractMap_();
  if (!cpSubtractMap) {
    console.log('cpSubtractMapキャッシュなし → buildIngredientDailyUsage() を実行します');
    cpSubtractMap = buildIngredientDailyUsage();
    saveCpSubtractMap_(cpSubtractMap);
  } else {
    console.log('cpSubtractMapをキャッシュから読み込みました');
  }

  // 曜日別平均を最新の食材日時実績から計算し直す（仕入規格量を変えても自動追従させるため）
  try { buildTrendHistoryAverages_(); } catch (e) { console.log('トレンド履歴の平均更新をスキップ: ' + e.message); }

  const masterSheet = ss.getSheetByName('食材マスタ');
  if (!masterSheet) { console.log('「食材マスタ」シートが見つかりません'); return; }
  const lastRowMaster = masterSheet.getLastRow();
  if (lastRowMaster < 2) { console.log('食材マスタにデータがありません'); return; }

  const masterData = masterSheet.getRange(2, 1, lastRowMaster-1, 3).getValues().filter(r => r[0] !== '' && r[2] !== '消耗品');
  const codeCount = {};
  masterData.forEach(r => { const code = String(r[0]); codeCount[code] = (codeCount[code]||0)+1; });
  const dupCodes = Object.entries(codeCount).filter(([,cnt]) => cnt > 1).map(([code,cnt]) => code+'（'+cnt+'件）');
  if (dupCodes.length > 0) { SpreadsheetApp.getUi().alert('【警告】食材マスタに重複コードがあります。処理を中止します。\n\n'+dupCodes.join('\n')); return; }

  const uniqueList = masterData.map(r => [r[0], r[1]]);
  // 起点日：既定は「次の火曜」。臨時で今週分を作り直したい時は startDateOverride(Date) を渡す。
//   （トリガーからはイベントオブジェクトが渡るが Date ではないので、必ず既定＝次週になる）
let startDate;
if (startDateOverride instanceof Date) {
  startDate = new Date(startDateOverride);
  startDate.setHours(0, 0, 0, 0);
} else {
  const today = new Date(), todayCode = today.getDay() + 1;
  let diff = (3 - todayCode + 7) % 7; if (diff === 0) diff = 7;
  startDate = new Date(today);
  startDate.setDate(today.getDate() + diff);
  startDate.setHours(0, 0, 0, 0);
}

const output = [];
uniqueList.forEach(([foodCode, foodName]) => {
  for (let i = 0; i < 9; i++) {
    const d = new Date(startDate);
    d.setDate(startDate.getDate() + i);
    d.setHours(0, 0, 0, 0);  // ★追加：各日付も時刻をクリア
    output.push([foodCode, foodName, d, d.getDay() + 1]);
  }
});

  // 基本データ（A〜D列）を書き込み、シートを常に適正サイズに保つ
  // （発注バッファ_ステップ1 と同じ方式。余分な行を毎回削除して肥大化・再計算増大を防ぐ）
  const MARGIN = 5;
  const dataRows   = output.length;
  const desiredMax = (START_ROW - 1) + dataRows + MARGIN;

  if (dataRows > 0) sheet.getRange(START_ROW, 1, dataRows, 4).setValues(output);

  let curMax = sheet.getMaxRows();
  const firstBlank = START_ROW + dataRows;
  if (curMax >= firstBlank) sheet.getRange(firstBlank, 1, curMax - firstBlank + 1, 4).clearContent();

  curMax = sheet.getMaxRows();
  if (curMax > desiredMax)      sheet.deleteRows(desiredMax + 1, curMax - desiredMax);
  else if (curMax < desiredMax) sheet.insertRowsAfter(curMax, desiredMax - curMax);

  SpreadsheetApp.flush();
  Utilities.sleep(700);

  buildOrderBufferF_(sheet);
  buildOrderBufferO_(sheet, cpSubtractMap);
  updateOrderSheetColor_();
  updateOrderSheetHeader_(startDate);
  trimOrderSheetRows_();

  console.log('発注計算バッファの生成が完了しました');
}

// 【臨時】今週分（直近の火曜起点）で発注ワークシートを作り直す。
//   通常は毎朝「次の火曜〜9日分」を作るが、予測を変えた・商品を追加した等で
//   “今動いている今週分”を今すぐ反映したい時に使う。翌朝の自動処理で通常（次週分）に戻る。
//   在庫の手入力は横型シート側にあるので消えない（横型の日付・基準はバッファを数式参照するため自動追従）。
function 発注ワークシート_今週分で作り直す() {
  const ui = SpreadsheetApp.getUi();
  const today = new Date();
  const todayCode = today.getDay() + 1;          // 日=1..土=7
  const back = (todayCode - 3 + 7) % 7;          // 直近の火曜まで戻る日数（今日が火曜なら0）
  const startDate = new Date(today);
  startDate.setDate(today.getDate() - back);
  startDate.setHours(0, 0, 0, 0);
  const label = Utilities.formatDate(startDate, 'Asia/Tokyo', 'M/d');
  const resp = ui.alert('今週分で作り直す（臨時）',
    '起点 ' + label + '（火）から9日分で発注ワークシートを作り直します。\n' +
    '予測の変更・商品の追加を“今週分”に反映する臨時処理です。\n' +
    '※在庫の手入力は残ります。翌朝の自動処理で通常の「次週分」に戻ります。\n\n実行しますか？',
    ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;
  発注計算バッファ_完全生成(startDate);
  ui.alert('今週分（' + label + '〜）で作り直しました。\n横型シートの日付・基準は自動で追従します。');
}

// 発注計算バッファ_完全生成 を2分割したもの（タイムアウト対策）
// トリガー: 06:50 → 発注バッファ_ステップ1 / 07:10 → 発注バッファ_ステップ2
function 発注バッファ_ステップ1() {
  const SHEET_NAME = '発注計算バッファ', START_ROW = 2;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) { console.log('「発注計算バッファ」シートが見つかりません'); return; }

  applyFoodCodeSwitches_();  // 食材コードの自動切替（対象がある日だけ商品マスタを更新。無ければ即終了で軽い）

  let cpSubtractMap = loadCpSubtractMap_();
  if (!cpSubtractMap) {
    console.log('cpSubtractMapキャッシュなし → buildIngredientDailyUsage() を実行します');
    cpSubtractMap = buildIngredientDailyUsage();
    saveCpSubtractMap_(cpSubtractMap);
  } else {
    console.log('cpSubtractMapをキャッシュから読み込みました');
  }

  // 曜日別平均を最新の食材日時実績から計算し直す（仕入規格量を変えても自動追従させるため）
  try { buildTrendHistoryAverages_(); } catch (e) { console.log('トレンド履歴の平均更新をスキップ: ' + e.message); }

  const masterSheet = ss.getSheetByName('食材マスタ');
  if (!masterSheet) { console.log('「食材マスタ」シートが見つかりません'); return; }
  const lastRowMaster = masterSheet.getLastRow();
  if (lastRowMaster < 2) { console.log('食材マスタにデータがありません'); return; }

  const masterData = masterSheet.getRange(2, 1, lastRowMaster-1, 3).getValues().filter(r => r[0] !== '' && r[2] !== '消耗品');
  const codeCount = {};
  masterData.forEach(r => { const code = String(r[0]); codeCount[code] = (codeCount[code]||0)+1; });
  const dupCodes = Object.entries(codeCount).filter(([,cnt]) => cnt > 1).map(([code,cnt]) => code+'（'+cnt+'件）');
  if (dupCodes.length > 0) { SpreadsheetApp.getUi().alert('【警告】食材マスタに重複コードがあります。処理を中止します。\n\n'+dupCodes.join('\n')); return; }

  const uniqueList = masterData.map(r => [r[0], r[1]]);
  const today = new Date(), todayCode = today.getDay() + 1;
  let diff = (3 - todayCode + 7) % 7; if (diff === 0) diff = 7;
  const startDate = new Date(today);
  startDate.setDate(today.getDate() + diff);
  startDate.setHours(0, 0, 0, 0);

  const output = [];
  uniqueList.forEach(([foodCode, foodName]) => {
    for (let i = 0; i < 9; i++) {
      const d = new Date(startDate);
      d.setDate(startDate.getDate() + i);
      d.setHours(0, 0, 0, 0);
      output.push([foodCode, foodName, d, d.getDay() + 1]);
    }
  });

  // ── 基本データ（A〜D列）を書き込み、シートを常に適正サイズに保つ ──
  //   E〜P列は全行を計算する重い数式（ARRAYFORMULA/MAP/LAMBDA等）が入っている。
  //   シートの行数が多いほど毎回その全行が再計算され、書込みがタイムアウトする。
  //   そこで「空欄にするだけ（＝行は残る）」をやめ、余分な行を毎回削除して
  //   常に『ヘッダー1行＋データ行＋わずかな余白』のサイズに保つ。
  //   これにより一度肥大化しても自動で適正サイズに戻り、二度と重くならない。
  const MARGIN = 5;                                   // 予備の空行（食材が数件増えても即対応できる余白）
  const dataRows   = output.length;
  const desiredMax = (START_ROW - 1) + dataRows + MARGIN;   // ヘッダー + データ + 余白

  if (dataRows > 0) sheet.getRange(START_ROW, 1, dataRows, 4).setValues(output);

  // データより下に残っている A〜D 列の古い値を消す（削除前の掃除）
  let curMax = sheet.getMaxRows();
  const firstBlank = START_ROW + dataRows;            // データ直後の行
  if (curMax >= firstBlank) sheet.getRange(firstBlank, 1, curMax - firstBlank + 1, 4).clearContent();

  // シートの行数を適正サイズへ。多ければ削除、足りなければ追加（食材が増えた場合）
  curMax = sheet.getMaxRows();
  if (curMax > desiredMax)      sheet.deleteRows(desiredMax + 1, curMax - desiredMax);
  else if (curMax < desiredMax) sheet.insertRowsAfter(curMax, desiredMax - curMax);

  // P2セル数式を最新化（lot = 仕入納品量/仕入規格量 で正しい発注単位を計算）
  setOrderBufferLotFormula_(sheet);

  SpreadsheetApp.flush();
  console.log('発注バッファ ステップ1（基本データ書込）完了');
}

// P2セルのlot計算式を食材マスタ参照方式に設定する
// （初回または数式が壊れた場合に備えてステップ1から毎回呼ぶ）
function setOrderBufferLotFormula_(sheet) {
  const formulaSheet = sheet || SpreadsheetApp.getActiveSpreadsheet().getSheetByName('発注計算バッファ');
  if (!formulaSheet) return;
  const p2 = formulaSheet.getRange('P2');
  const formula = `=LET(
  start, MIN(FILTER(C2:C, C2:C<>"")),
  MAP(A2:A, C2:C, O2:O,
    LAMBDA(code, dt, need,
      IF(need="", "",
        IF(dt > start + 6, "",
          IF(need=0, 0,
            LET(
              spec_g,  IFERROR(VLOOKUP(code, 食材マスタ!A:F, 6, FALSE), 1),
              deliv_g, IFERROR(VLOOKUP(code, 食材マスタ!A:G, 7, FALSE), 1),
              lot,     IF(spec_g > 0, MAX(1, ROUND(deliv_g / spec_g, 6)), 1),
              MAX(1, CEILING(need / lot - 0.25, 1))
            )
          )
        )
      )
    )
  )
)`;
  p2.setFormula(formula);
  console.log('発注計算バッファ P2セル: lot計算式を食材マスタ参照方式に更新しました');
}

// ── 手動1回実行：P2セル数式を即時修正する ──
function 発注バッファP列数式修正() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('発注計算バッファ');
  if (!sheet) { console.log('「発注計算バッファ」シートが見つかりません'); return; }
  setOrderBufferLotFormula_(sheet);
  console.log('P列数式を修正しました。次にステップ2（またはトリガー）でO列を再計算してください。');
}

// トレンド係数(K列)の更新 ── 週1回程度、手動実行
function 発注バッファ_トレンド更新() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('発注計算バッファ');
  if (!sheet) { console.log('「発注計算バッファ」シートが見つかりません'); return; }
  buildOrderBufferF_(sheet);
  console.log('トレンド係数(K列)の更新が完了しました');
}

function 発注バッファ_ステップ2() {
  const SHEET_NAME = '発注計算バッファ';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) { console.log('「発注計算バッファ」シートが見つかりません'); return; }

  let cpSubtractMap = loadCpSubtractMap_();
  if (!cpSubtractMap) {
    console.log('cpSubtractMapキャッシュなし → buildIngredientDailyUsage() を実行します');
    cpSubtractMap = buildIngredientDailyUsage();
    saveCpSubtractMap_(cpSubtractMap);
  }

  const today = new Date(), todayCode = today.getDay() + 1;
  let diff = (3 - todayCode + 7) % 7; if (diff === 0) diff = 7;
  const startDate = new Date(today);
  startDate.setDate(today.getDate() + diff);
  startDate.setHours(0, 0, 0, 0);

  buildOrderBufferO_(sheet, cpSubtractMap);
  updateOrderSheetColor_();
  updateOrderSheetHeader_(startDate);
  trimOrderSheetRows_();
  console.log('発注バッファ ステップ2（発注量計算・仕上げ）完了');
}

// ─────────────────────────────────────────────────────────────────────
// Step2.5: トレンド履歴K列
// ─────────────────────────────────────────────────────────────────────

/**
 * トレンド履歴の曜日別平均使用量（C〜I列）と算出週数（J列）を
 * 食材日時実績から自動再計算する。
 *
 * これまでC〜I列は固定値で、GASのどこからも更新されていなかった。
 * そのため仕入規格量（食材マスタF列）を直しても平均値が古い単位のまま残り、
 * 発注量が狂う危険があった。ここを毎回計算し直すことで単位変更に自動追従する。
 *
 * 計算方法：食材日時実績の「使用量」を 食材コード×曜日 で単純平均（行がある日だけ）。
 *   ※現在の固定値と完全一致することを実データで検証済み（比1.000）。
 * 使用量は 仕入規格量あたりの数量なので、規格量を変えれば平均も自動でその単位になる。
 * K列（店舗トレンド係数）は buildOrderBufferF_ が計算するのでここでは触らない。
 */
function buildTrendHistoryAverages_() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const src = ss.getSheetByName('食材日時実績');
  const dst = ss.getSheetByName('トレンド履歴');
  if (!src || !dst) { console.log('食材日時実績またはトレンド履歴が見つかりません'); return; }
  if (src.getLastRow() < 2) { console.log('食材日時実績にデータがありません'); return; }

  // ── 食材コード×曜日ごとに、直近重み(EWMA)で使用量を集計 ───────────────
  //   単純平均でなく「直近を重く見る指数平滑」。8週窓での交差検証で日次・週合計とも最良。
  //   古い日→新しい日の順に e = α×今回 + (1-α)×これまで。最終値＝直近重視の平均。
  const EWMA_ALPHA = 0.35;
  const seriesK = {}, dateSet = {}, nameMap = {};
  src.getRange(2, 1, src.getLastRow() - 1, 5).getValues().forEach(r => {
    if (!r[0] || !r[2]) return;
    const dow = Number(r[1]); if (!(dow >= 1 && dow <= 7)) return;
    const code = String(Number(r[2]));
    const key  = code + '_' + dow;
    const d = r[0] instanceof Date ? r[0] : new Date(r[0]);
    (seriesK[key] = seriesK[key] || []).push({ t: d.getTime(), v: Number(r[4]) || 0 });
    if (r[3]) nameMap[code] = String(r[3]);
    dateSet[Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy/MM/dd')] = true;
  });
  const weeks = Math.max(1, Math.round(Object.keys(dateSet).length / 7));

  const ewmaOf = key => {
    const arr = seriesK[key]; if (!arr || arr.length === 0) return 0;
    arr.sort((a, b) => a.t - b.t);           // 古い順
    let e = arr[0].v;
    for (let i = 1; i < arr.length; i++) e = EWMA_ALPHA * arr[i].v + (1 - EWMA_ALPHA) * e;
    return Math.round(e * 1e6) / 1e6;
  };

  const avgRow = code => {
    const a = [];
    for (let dow = 1; dow <= 7; dow++) a.push(ewmaOf(code + '_' + dow));
    a.push(weeks);
    return a;   // C〜J の8個
  };

  // ── 既存行は順序を保ったままC〜J列を更新 ───────────────────────────
  const last = dst.getLastRow();
  if (last < 2) { console.log('トレンド履歴に食材行がありません'); return; }
  const existing = dst.getRange(2, 1, last - 1, 2).getValues();
  const known = {};
  const out = existing.map(r => {
    if (r[0] === '' || r[0] === null) return ['', '', '', '', '', '', '', ''];
    const code = String(Number(r[0]));
    known[code] = true;
    return avgRow(code);
  });
  dst.getRange(2, 3, out.length, 8).setValues(out);

  // ── 食材マスタにあってトレンド履歴に無い食材を追加（新規食材への自動対応）──
  const im = ss.getSheetByName('食材マスタ');
  const added = [];
  if (im && im.getLastRow() > 1) {
    im.getRange(2, 1, im.getLastRow() - 1, 2).getValues().forEach(r => {
      if (!r[0]) return;
      const code = String(Number(r[0]));
      if (known[code]) return;
      known[code] = true;
      added.push([Number(r[0]), String(r[1] || nameMap[code] || '')].concat(avgRow(code)).concat([1]));
    });
  }
  if (added.length > 0) dst.getRange(dst.getLastRow() + 1, 1, added.length, 11).setValues(added);

  console.log('トレンド履歴の曜日別平均を更新しました（' + out.length + '品目'
            + (added.length > 0 ? ' ＋新規' + added.length + '品目' : '')
            + '／集計' + Object.keys(dateSet).length + '日＝' + weeks + '週）');
}

function buildTrendHistoryAverages() {
  buildTrendHistoryAverages_();
  SpreadsheetApp.getUi().alert('トレンド履歴の曜日別平均を再計算しました。\n\n続けて「発注バッファ_ステップ1」→「ステップ2」を実行すると発注数に反映されます。');
}


/**
 * イベント適用結果シートの掃除。
 * 発注のイベント倍率(G列)は、直近使用量ベースのトレンド履歴に対して二重計上になるため常に1でよい。
 * G列の依存を外して常に1にしてから、シートを「_削除待ち_」へ退避する（削除しない）。
 * 発注結果は従来と変わらない（イベント倍率は元々全て1だった）。
 */
function cleanupEventApplyResult() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const PREFIX = '_削除待ち_';
  const log = [];
  // ① 発注バッファ G列（イベント倍率）を「常に1」に。イベント適用結果への依存を外す。
  const buf = ss.getSheetByName('発注計算バッファ');
  if (buf) {
    buf.getRange('G2').setFormula('=ARRAYFORMULA(IF(A2:A="","",1))');
    log.push('発注バッファ G列（イベント倍率）を常に1に変更（イベント適用結果への依存を解消）');
  } else {
    log.push('発注計算バッファが見つかりませんでした');
  }
  // ② イベント適用結果シートを退避（削除せず _削除待ち_ に改名して非表示）
  const ev = ss.getSheetByName('イベント適用結果');
  if (ev) {
    if (!ss.getSheetByName(PREFIX + 'イベント適用結果')) {
      ev.setName(PREFIX + 'イベント適用結果'); ev.hideSheet();
      log.push('イベント適用結果 → ' + PREFIX + 'イベント適用結果（非表示・いつでも戻せます）');
    } else { log.push('イベント適用結果 はすでに退避済みです'); }
  } else { log.push('イベント適用結果 は見つかりませんでした（すでに整理済みの可能性）'); }
  // 注意: SpreadsheetApp.flush() は付けない。
  // G2変更で発注計算バッファの重い数式(E〜P列)が全再計算になり、それを同期で待つと6分制限を超える。
  // 変更は関数終了時に自動確定し、再計算はシート側で非同期に進むため flush は不要。
  const msg = 'イベント適用結果の掃除が完了しました。\n発注結果は今までと変わりません（イベント倍率は元々全て1でした）。\n\n' + log.map(s => '・' + s).join('\n');
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { console.log(msg); }
}

// 【自動収束】短期イベントの信頼度加重ブレンド。実績が貯まるほど設定倍率を実績へ寄せる。
//   採用 = (設定×K + 実績平均×n) / (K+n)。 K=設定の信頼度(実績K件相当・大きいほど収束が遅い)。
//   n=0 → 設定のまま。n≫K → ほぼ実績。安全域 0.5〜1.6 にクランプ。
//   ※長期(≥21日)CP/イベントは「実績÷曜日平均」が循環して濁るため、呼び出し側で対象外にする。
function blendEventRate_(setRate, actualRate, count) {
  const K = 6;
  const s = setRate > 0 ? setRate : 1.0;
  if (!(count > 0) || !(actualRate > 0)) return Math.round(s * 1000) / 1000;
  const blended = (s * K + actualRate * count) / (K + count);
  return Math.round(Math.min(1.6, Math.max(0.5, blended)) * 1000) / 1000;
}

// 【統合】イベント/CPの学習結果を「イベント登録」「CP設定」の結果列へ書き戻す。
//   ・各行の「名前」をキーに、実績倍率(参考)・サンプル数・次回予測倍率(実際に適用する値)・更新日を右側に記入。
//   ・入力列（左側）は一切触らない（clearContentsは使わない）。役目を終えた「イベント実績」は _削除待ち_ へ退避。
//   ・「次回予測倍率」は getEventRate と同じ解決順（繰り返し学習→名前別学習→設定倍率）に合わせる。
//     日付指定イベント/CPは方針どおり設定倍率を適用するため、実績倍率(参考)とは別に表示する。
function writeEventResultsBack_(ss, eventActualRates, repeatEventCoeffMap, namedEventCoeffMap) {
  const stamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');

  // 名前 → 実績倍率(参考)とサンプル数（優先順：日付実績→繰り返し→名前別）
  const pickActual = (name) => {
    if (eventActualRates[name])    return { rate: eventActualRates[name].actualRate, count: eventActualRates[name].count };
    if (repeatEventCoeffMap[name]) return { rate: repeatEventCoeffMap[name].rate,    count: repeatEventCoeffMap[name].count };
    if (namedEventCoeffMap[name])  return { rate: namedEventCoeffMap[name].rate,     count: namedEventCoeffMap[name].count };
    return null;
  };
  // 名前 → 実際に適用する倍率（getEventRateと同じ。短期＋長期の期間イベントは実績へ自動収束、長期CPのみ対象外）
  const pickApplied = (name, settingRate, learnByName, durationDays, isCPSheet) => {
    if (repeatEventCoeffMap[name]) return repeatEventCoeffMap[name].rate;
    if (learnByName && namedEventCoeffMap[name]) return namedEventCoeffMap[name].rate;
    const info = eventActualRates[name];
    if ((durationDays < 21 || !isCPSheet) && info && info.count > 0) return blendEventRate_(settingRate, info.actualRate, info.count);
    return settingRate > 0 ? settingRate : 1.0;
  };

  // 1シートに結果4列を書く。startCol=結果開始列(1始まり)、setRateIdx/flagIdx/repeatIdx=入力列のindex(0始まり、無ければ-1)
  const writeBack = (sheetName, startCol, setRateIdx, flagIdx, repeatIdx, isCPSheet) => {
    const sh = ss.getSheetByName(sheetName);
    if (!sh || sh.getLastRow() < 2) return;
    if (sh.getMaxColumns() < startCol + 3) sh.insertColumnsAfter(sh.getMaxColumns(), startCol + 3 - sh.getMaxColumns());
    sh.getRange(1, startCol, 1, 4).setValues([['実績倍率(参考)', 'サンプル数', '次回予測倍率(適用)', '更新日']]);
    const inputData = sh.getRange(2, 1, sh.getLastRow() - 1, startCol - 1).getValues();  // 入力列だけ読む
    const out = inputData.map(row => {
      const name = (row[0] || '').toString().trim();
      if (!name) return ['', '', '', ''];
      const settingRate = Number(row[setRateIdx]) || 0;
      const learnByName = flagIdx >= 0
        ? ((row[flagIdx] || '').toString().trim() === '自動' || (repeatIdx >= 0 && (row[repeatIdx] || '').toString().trim() !== ''))
        : false;
      const sD = toDate_(row[1]), eD = toDate_(row[2]);
      const durationDays = (sD && eD) ? (eD - sD) / 86400000 : 0;  // 期間(日)。長期(≥21)は自動収束の対象外
      const a = pickActual(name);
      const applied = pickApplied(name, settingRate, learnByName, durationDays, isCPSheet);
      return [a ? a.rate : '', a ? a.count : 0, applied, stamp];
    });
    sh.getRange(2, startCol, out.length, 4).setValues(out);
  };

  // イベント登録: 入力A〜K(11列)。結果L(12)〜。設定倍率G(idx6)・自動化区分I(idx8)・繰り返しK(idx10)。期間イベント＝isCPSheet:false
  writeBack('イベント登録', 12, 6, 8, 10, false);
  // CP設定: 入力A〜I(9列)。結果J(10)〜。補正倍率F(idx5)。自動化・繰り返し列なし(-1)。CP＝isCPSheet:true
  writeBack('CP設定', 10, 5, -1, -1, true);

  // 役目を終えた「イベント実績」シートを退避（1回だけ・削除はしない）
  const ev = ss.getSheetByName('イベント実績');
  if (ev && !ss.getSheetByName('_削除待ち_イベント実績')) { ev.setName('_削除待ち_イベント実績'); ev.hideSheet(); }
}

/**
 * 使われていないシートを「_削除待ち_」に改名して非表示にする（削除はしない）。
 * 中身はそのまま残るので、いつでも元に戻せる。
 *
 * 対象と、その役割が今どこで実現されているか：
 *  ・週次集計       … 週トレンド分析用。現在は calcLongTermTrend_（直近4週÷直近12週を
 *                     曜日別に算出・異常値除外・0.7〜1.3で上下限）が同じ役割をより高精度に担う。
 *  ・イベント適用結果 … イベントの予測と実績の比較用。現在は eventActualRates
 *                     （実績÷通常同曜日平均を学習し次回予測に自動反映）が担う＝イベント実績シート。
 *  ・日次集計       … 日次売上実績と重複。F列「使用判定」はコードのどこからも参照されておらず、
 *                     除外判定は異常値ログ（IQR自動）＋イベント登録の「除外」区分が担う。
 */
function archiveUnusedSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const PREFIX = '_削除待ち_';
  const targets = ['週次集計', '日次集計'];  // イベント適用結果は cleanupEventApplyResult()（G列依存を先に外す）で退避する
  const done = [], skipped = [];

  targets.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) { skipped.push(name + '（見つかりません。すでに整理済みの可能性）'); return; }
    sh.setName(PREFIX + name);
    sh.hideSheet();
    done.push(name);
  });

  const msg = (done.length > 0
      ? '次のシートを「' + PREFIX + '」に改名して非表示にしました。\n（削除していないので、いつでも戻せます）\n\n'
        + done.map(n => '・' + n + ' → ' + PREFIX + n).join('\n')
      : '対象のシートはありませんでした。')
    + (skipped.length > 0 ? '\n\n【スキップ】\n' + skipped.map(n => '・' + n).join('\n') : '')
    + '\n\n※元に戻すには、シート名の「' + PREFIX + '」を消して再表示してください。'
    + '\n※完全に削除してよいと判断できたら、手動で削除してください。';
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { console.log(msg); }
}

/** 「_削除待ち_」にしたシートを元の名前に戻して再表示する */
function restoreArchivedSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const PREFIX = '_削除待ち_';
  const done = [];
  ss.getSheets().forEach(sh => {
    const n = sh.getName();
    if (n.indexOf(PREFIX) !== 0) return;
    sh.setName(n.substring(PREFIX.length));
    sh.showSheet();
    done.push(n.substring(PREFIX.length));
  });
  const msg = done.length > 0
    ? '次のシートを元に戻しました。\n\n' + done.map(n => '・' + n).join('\n')
    : '「' + PREFIX + '」が付いたシートはありませんでした。';
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { console.log(msg); }
}


function buildOrderBufferF_(targetSheet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const today = new Date(); today.setHours(0,0,0,0);
  const past4wks = new Date(today); past4wks.setDate(today.getDate()-28);
  const past12wks = new Date(today); past12wks.setDate(today.getDate()-84);
  const codeMap = {};
  const switchSheet = ss.getSheetByName('食材コード切替履歴');
  if (switchSheet && switchSheet.getLastRow() > 1) {
    switchSheet.getRange(2,1,switchSheet.getLastRow()-1,3).getValues().filter(r=>r[0]&&r[1]).forEach(r=>{ codeMap[String(r[0])]=String(r[1]); });
  }
  const foodTrend = {};
  const jissekiSheet = ss.getSheetByName('食材日時実績');
  if (jissekiSheet && jissekiSheet.getLastRow() > 1) {
    const jData = jissekiSheet.getRange(2,1,jissekiSheet.getLastRow()-1,5).getValues().filter(r=>r[0]&&r[2]&&r[4]>0);
    const recent4 = {}, past12 = {};
    jData.forEach(row => {
      const d = new Date(row[0]); d.setHours(0,0,0,0); if (d >= today) return;
      const dow=Number(row[1]), rawCode=String(row[2]), code=codeMap[rawCode]||rawCode, qty=Number(row[4]), key=`${code}_${dow}`;
      if (d >= past4wks)  { if(!recent4[key]) recent4[key]=[]; recent4[key].push(qty); }
      if (d >= past12wks) { if(!past12[key])  past12[key] =[]; past12[key].push(qty);  }
    });
    Object.keys(past12).forEach(key => {
      const p12arr=past12[key], p4arr=recent4[key]||[]; if(!p12arr.length) return;
      const p12avg=p12arr.reduce((a,b)=>a+b,0)/p12arr.length, p4avg=p4arr.length?p4arr.reduce((a,b)=>a+b,0)/p4arr.length:p12avg;
      foodTrend[key] = p12avg > 0 ? p4avg/p12avg : 1.0;
    });
  }
  const salesTrend = {};
  const salesSheet = ss.getSheetByName('日次売上実績');
  if (salesSheet && salesSheet.getLastRow() > 1) {
    const sData = salesSheet.getRange(2,1,salesSheet.getLastRow()-1,4).getValues().filter(r=>r[0]&&r[3]>0);
    const sRecent4={1:[],2:[],3:[],4:[],5:[],6:[],7:[]}, sPast12={1:[],2:[],3:[],4:[],5:[],6:[],7:[]};
    sData.forEach(row => {
      const d=new Date(row[0]); d.setHours(0,0,0,0); if(d>=today) return;
      const dow=Number(row[1]), sales=Number(row[3]); if(dow<1||dow>7) return;
      if(d>=past4wks)  sRecent4[dow].push(sales);
      if(d>=past12wks) sPast12[dow].push(sales);
    });
    [1,2,3,4,5,6,7].forEach(dow => {
      const p12arr=sPast12[dow], p4arr=sRecent4[dow];
      if(!p12arr.length){ salesTrend[dow]=1.0; return; }
      const p12avg=p12arr.reduce((a,b)=>a+b,0)/p12arr.length, p4avg=p4arr.length?p4arr.reduce((a,b)=>a+b,0)/p4arr.length:p12avg;
      salesTrend[dow]=p12avg>0?p4avg/p12avg:1.0;
    });
  }
  const trendSheet = ss.getSheetByName('トレンド履歴');
  if (!trendSheet || trendSheet.getLastRow() < 2) { console.log('「トレンド履歴」シートが見つかりません'); return; }
  const trendData = trendSheet.getRange(2,1,trendSheet.getLastRow()-1,1).getValues();
  const targetDows = [3,4,6,1];
  const kOutput = trendData.map(row => {
    const code = row[0]; if (!code) return [''];
    let foodSum=0, foodCount=0;
    targetDows.forEach(dow => { const key=`${code}_${dow}`; if(foodTrend[key]!==undefined){ foodSum+=foodTrend[key]; foodCount++; } });
    const foodCoeff = foodCount>0 ? foodSum/foodCount : 1.0;
    const salesCoeff = targetDows.reduce((s,dow)=>s+(salesTrend[dow]||1.0),0)/targetDows.length;
    return [Math.round(foodCoeff*salesCoeff*100)/100];
  });
  trendSheet.getRange(2, 11, kOutput.length, 1).setValues(kOutput);
}

function buildOrderBufferF() { buildOrderBufferF_(null); console.log('トレンド履歴K列の再計算が完了しました'); }

// ─────────────────────────────────────────────────────────────────────
// CP休止（一定期間の販売停止）：CP設定のI列(通常商品CP)='休止'の行を読む。
//   期間中はその商品が売れないので、使う食材を「その商品ぶんだけ」発注・仕込みから差し引く。
//   共有食材（カルビ・米・野菜等）も対象。※専用食材だけ止める既存ingCpEndとは別で、按分して減らす。
// ─────────────────────────────────────────────────────────────────────
function getCpPauseRows_(ss) {
  const sh = ss.getSheetByName('CP設定');
  if (!sh || sh.getLastRow() < 2) return [];
  const out = [];
  sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues().forEach(row => {
    const flag = (row[8] || '').toString().trim();
    if (flag !== '休止') return;
    const parseD = v => v instanceof Date ? v : (v ? new Date(String(v).replace(/^(\d{4})(\d{2})(\d{2})$/, '$1/$2/$3')) : null);
    const s = parseD(row[1]), e = parseD(row[2]);
    if (!s || !e || isNaN(s.getTime()) || isNaN(e.getTime())) return;
    s.setHours(0, 0, 0, 0); e.setHours(0, 0, 0, 0);
    const codes = (typeof row[3] === 'number') ? [Math.round(row[3])]
      : String(row[3] || '').split(',').flatMap(c => { c = c.trim(); const m = c.match(/^(\d+)-(\d+)$/); if (m) { const a = []; for (let n = +m[1]; n <= +m[2]; n++) a.push(n); return a; } const n = +c; return n > 0 ? [n] : []; });
    if (!codes.length) return;
    out.push({ name: (row[0] || '').toString().trim(), start: s, end: e, codes: codes.map(String) });
  });
  return out;
}

// 休止商品の食材使用量を「食材_曜日 → 規格量単位/日」で見積る（トレンド履歴・発注ingDowAvgと同じ単位）。
//   単位使用量(商品マスタM) × 出食数平均(実績データ・曜日別) / 規格量(食材マスタ) / 歩留率
function buildCpPauseSubMap_(ss) {
  const windows = getCpPauseRows_(ss);
  if (!windows.length) return { windows: [], sub: {} };
  const pausedCodes = new Set();
  windows.forEach(w => w.codes.forEach(c => pausedCodes.add(String(c))));
  // 商品マスタ：休止商品コード → [{ing, qty}]／全食材 → その食材を使う全商品コード（専用食材の判定用）
  const recipe = {}, ingToProducts = {};
  const pm = ss.getSheetByName('商品マスタ');
  if (pm && pm.getLastRow() > 1) {
    pm.getRange(2, 1, pm.getLastRow() - 1, 13).getValues().forEach(r => {
      const code = String(r[0] || '').trim();
      const ing = String(r[4] || '').trim(), qty = Number(r[12]) || 0;
      if (code && ing) { if (!ingToProducts[ing]) ingToProducts[ing] = new Set(); ingToProducts[ing].add(code); }
      if (!pausedCodes.has(code)) return;
      if (!ing || qty <= 0) return;
      if (!recipe[code]) recipe[code] = [];
      recipe[code].push({ ing, qty });
    });
  }
  // 食材マスタ：規格量(col6)・歩留(col8)
  const spec = {}, yld = {};
  const im = ss.getSheetByName('食材マスタ');
  if (im && im.getLastRow() > 1) {
    im.getRange(2, 1, im.getLastRow() - 1, 8).getValues().forEach(r => {
      const c = String(r[0] || '').trim(); if (!c) return;
      spec[c] = Number(r[5]) || 0; yld[c] = Number(r[7]) || 1;
    });
  }
  // 実績データ：休止商品コード×曜日の平均出食数（直近90日・0除く）＋最近売れた全商品コード（専用食材判定用）
  const cntSum = {}, cntN = {}, soldCodes = new Set();
  const log = ss.getSheetByName('実績データ');
  if (log && log.getLastRow() > 1) {
    const lim = new Date(); lim.setDate(lim.getDate() - 90); lim.setHours(0, 0, 0, 0);
    log.getDataRange().getValues().forEach((r, i) => {
      if (i === 0) return;
      const d = r[0], mc = String(r[4] || '').trim(), q = Number(r[6]) || 0;
      if (!d || !mc || q <= 0) return;
      const dt = d instanceof Date ? d : new Date(d); if (isNaN(dt.getTime())) return;
      dt.setHours(0, 0, 0, 0); if (dt < lim) return;
      soldCodes.add(mc);
      if (!pausedCodes.has(mc)) return;
      const k = mc + '_' + (dt.getDay() + 1);
      cntSum[k] = (cntSum[k] || 0) + q; cntN[k] = (cntN[k] || 0) + 1;
    });
  }
  const allSum = {}, allN = {};
  Object.keys(cntSum).forEach(k => { const c = k.split('_')[0]; allSum[c] = (allSum[c] || 0) + cntSum[k]; allN[c] = (allN[c] || 0) + cntN[k]; });
  const avgCount = (code, dow) => { const k = code + '_' + dow; if (cntN[k]) return cntSum[k] / cntN[k]; if (allN[code]) return allSum[code] / allN[code]; return 0; };
  const sub = {};
  for (let dow = 1; dow <= 7; dow++) {
    pausedCodes.forEach(code => {
      const rc = recipe[code]; if (!rc) return;
      const ac = avgCount(code, dow); if (ac <= 0) return;
      rc.forEach(({ ing, qty }) => {
        const sp = spec[ing]; if (!sp) return;
        const u = ac * qty / sp / (yld[ing] || 1);
        const key = ing + '_' + dow;
        sub[key] = (sub[key] || 0) + u;
      });
    });
  }
  // 専用食材：その食材を「最近売れている商品の中で」使うのが休止商品だけ＝期間中は完全に0にする
  //   （古い休売コードは除外。休止商品が1つでも最近実績を持つ食材が対象）
  const exclusive = new Set();
  Object.keys(ingToProducts).forEach(ing => {
    const users = [...ingToProducts[ing]].filter(c => soldCodes.has(c));
    if (users.length > 0 && users.every(c => pausedCodes.has(c))) exclusive.add(ing);
  });
  if (exclusive.size) console.log('CP休止・専用食材(期間中0): ' + [...exclusive].join(','));
  return { windows, sub, exclusive };
}

// ─────────────────────────────────────────────────────────────────────
// Step3: O列計算
// ─────────────────────────────────────────────────────────────────────

function buildOrderBufferO_(targetSheet, cpSubtractMap) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = targetSheet || ss.getSheetByName('発注計算バッファ');
  if (!sheet) return;
  const lastRow = sheet.getLastRow(); if (lastRow < 2) return;
  const subtractMap = cpSubtractMap || {};

  // CP休止：休止期間中は、休止商品ぶんの食材使用量を差し引く（食材_曜日→規格量単位/日）。専用食材は0にする。
  const pauseData = buildCpPauseSubMap_(ss);
  const pauseWindows = pauseData.windows, pauseSub = pauseData.sub, pauseExcl = pauseData.exclusive || new Set();
  const inPauseWindow = (d) => pauseWindows.some(w => d >= w.start && d <= w.end);
  if (pauseWindows.length) console.log('CP休止: ' + pauseWindows.map(w => w.name + '(' + Utilities.formatDate(w.start, 'Asia/Tokyo', 'MM/dd') + '〜' + Utilities.formatDate(w.end, 'Asia/Tokyo', 'MM/dd') + ')').join(', '));

  // 祝日タグマップ
  const holidayTagMap = {};
  const hmSheet = ss.getSheetByName('祝日マスタ');
  if (hmSheet && hmSheet.getLastRow() > 1) {
    hmSheet.getRange(2,1,hmSheet.getLastRow()-1,7).getValues().forEach(row => {
      if (!row[0]||!row[6]) return;
      const tag=(row[6]||'').toString().trim(); if(tag==='通常営業') return;
      const ds=row[0] instanceof Date ? Utilities.formatDate(row[0],'Asia/Tokyo','yyyy/MM/dd') : row[0].toString();
      holidayTagMap[ds]=tag;
    });
  }

  // 祝日倍率マスタ
  const holidayRateMap = {};
  const hrSheet = ss.getSheetByName('祝日倍率マスタ');
  if (hrSheet && hrSheet.getLastRow() > 1) {
    hrSheet.getRange(2,1,hrSheet.getLastRow()-1,8).getValues().forEach(row => {
      if (!row[0]||!row[2]) return;
      const foodCode=String(row[0]), tag=(row[2]||'').toString().trim();
      const autoRate=Number(row[6])||1, manualRate=Number(row[7])||0;
      holidayRateMap[foodCode+'_'+tag] = manualRate>0 ? manualRate : autoRate;
    });
    console.log('祝日倍率マスタ読込: '+Object.keys(holidayRateMap).length+'件');
  }

  // 食材日時実績から集計（28日フィルター）
  const ingDowSum = {}, ingDowCount = {};
const todayO = new Date(); todayO.setHours(0, 0, 0, 0);
  const limitO = new Date(todayO); limitO.setDate(todayO.getDate() - 28);

  // ★切替CP：CP期間中のデータのみ除外し、CP前・CP後は両方使用する
  //   CP中    → 開始日〜今日を除外期間とする（CP前の通常実績を使用）
  //   CP終了後 → 開始日〜終了日を除外期間とする（CP前＋CP後の通常実績を両方使用）
  const ingExcludePeriodMap = {};  // { ingCode: [{start: Date, end: Date}] }
  const cpSheetO = ss.getSheetByName('CP設定');
  if (cpSheetO && cpSheetO.getLastRow() > 1) {
    // 商品マスタから メニューコード→食材コード のマップを作成
    const menuIngMapO = {};
    const menuMasterO = ss.getSheetByName('商品マスタ');
    if (menuMasterO && menuMasterO.getLastRow() > 1) {
      menuMasterO.getDataRange().getValues().forEach((row, i) => {
        if (i === 0 || !row[0] || !row[4]) return;
        const mc = String(row[0]);
        if (!menuIngMapO[mc]) menuIngMapO[mc] = [];
        menuIngMapO[mc].push(String(row[4]));
      });
    }
    cpSheetO.getRange(2, 1, cpSheetO.getLastRow()-1, 9).getValues().forEach(row => {
      const cpFlagO = (row[8] || '').toString().trim();
      const isSwitchCpO = cpFlagO === '切替';
      const isNormalCpO = cpFlagO !== '' && cpFlagO !== '切替';
      if (isNormalCpO) return; // 通常CP（実績除外型）は食材日時実績に入っていないのでスキップ
      const sRaw = row[1], eRaw = row[2];
      const sDate = sRaw instanceof Date ? sRaw : (sRaw ? new Date(String(sRaw).replace(/^(\d{4})(\d{2})(\d{2})$/, '$1/$2/$3')) : null);
      const eDate = eRaw instanceof Date ? eRaw : (eRaw ? new Date(String(eRaw).replace(/^(\d{4})(\d{2})(\d{2})$/, '$1/$2/$3')) : null);
      if (!sDate || isNaN(sDate.getTime())) return;
      sDate.setHours(0,0,0,0);
      const rawVal = row[3];
      const menuCodes = typeof rawVal === 'number' ? [Math.round(rawVal)]
        : String(rawVal||'').split(',').flatMap(c => { c = c.trim(); const m = c.match(/^(\d+)-(\d+)$/); if (m) { const arr = []; for (let n = Number(m[1]); n <= Number(m[2]); n++) arr.push(n); return arr; } const n = Number(c); return n > 0 ? [n] : []; });
      let excludeEnd;
      if (isSwitchCpO) {
        // 切替CP：CP期間中データを除外（CP前・CP後の通常実績を使う）
        if (!eDate || isNaN(eDate.getTime())) {
          excludeEnd = new Date(todayO);  // 終了日未定（CP継続中）：今日まで除外
        } else {
          eDate.setHours(0,0,0,0);
          // CP終了後7日間も除外（終了直後の在庫消化販売データを除く）
          const afterBuffer = new Date(eDate.getTime() + 7 * 24 * 60 * 60 * 1000);
          if (afterBuffer < limitO) return;  // 7日猶予後も28日以上前なら無視
          excludeEnd = afterBuffer;
        }
      } else {
        // I列=空のCP（新商品型）：終了済みの場合のみCP期間中データを除外
        // 進行中CPの実績はそのまま発注計算に使う
        if (!eDate || isNaN(eDate.getTime())) return;
        eDate.setHours(0,0,0,0);
        if (eDate >= todayO) return;  // 進行中CPは除外しない
        if (eDate < limitO) return;   // 28日以上前に終了したCPは無視
        excludeEnd = eDate;           // 終了済みCP：CP期間中データを除外
      }
      menuCodes.forEach(mc => {
        (menuIngMapO[String(mc)]||[]).forEach(ingCode => {
          if (!ingExcludePeriodMap[ingCode]) ingExcludePeriodMap[ingCode] = [];
          ingExcludePeriodMap[ingCode].push({ start: new Date(sDate), end: new Date(excludeEnd) });
        });
      });
    });
    if (Object.keys(ingExcludePeriodMap).length > 0) {
      console.log('CP除外期間フィルター: ' +
        Object.entries(ingExcludePeriodMap)
          .map(([k,v]) => k + '=' + v.map(p =>
            Utilities.formatDate(p.start,'Asia/Tokyo','MM/dd') + '〜' +
            Utilities.formatDate(p.end,'Asia/Tokyo','MM/dd') + '除外').join(','))
          .join(', '));
    }
  }

  const jissekiSheet = ss.getSheetByName('食材日時実績');
  if (jissekiSheet && jissekiSheet.getLastRow() > 1) {
    jissekiSheet.getRange(2,1,jissekiSheet.getLastRow()-1,6).getValues()
      .filter(r => {
        if (!r[0]||!r[2]||r[4]<=0||(r[5]||'')==='推定') return false;
        const d = r[0] instanceof Date ? r[0] : new Date(r[0]); d.setHours(0,0,0,0);
        if (d < limitO) return false;
        // ★切替CPのCP期間中データのみ除外（CP前・CP後は有効）
        const excludePeriods = ingExcludePeriodMap[String(r[2])] || [];
        if (excludePeriods.some(p => d >= p.start && d <= p.end)) return false;
        return true;
      })
      .forEach(row => {
        const key=row[2]+'_'+Number(row[1]);
        if (!ingDowSum[key])   ingDowSum[key]  =0;
        if (!ingDowCount[key]) ingDowCount[key]=0;
        ingDowSum[key]  +=Number(row[4]);
        ingDowCount[key]+=1;
      });
  }

  // CP除外量を差し引く
  let subtractCount = 0;
  Object.keys(subtractMap).forEach(key => {
    if (ingDowSum[key]!==undefined) { ingDowSum[key]=Math.max(0,ingDowSum[key]-subtractMap[key]); subtractCount++; }
  });
  if (subtractCount > 0) console.log('CP除外処理: '+subtractCount+'個の食材×曜日の合計を差し引きました');

  // 実績がない食材×曜日に「推定」行でフォールバック補完
  // （新食材は実績ゼロなので推定値を基準に使う。実績が1件でもあれば推定値は使わない）
  // ★ 複数商品が同じ食材を使う場合、同じ日付に複数の推定行が書かれる。
  //   先に「日付単位で合計」してから「曜日単位で平均」することで正しく集計する。
  if (jissekiSheet && jissekiSheet.getLastRow() > 1) {
    // Step1: 同日・同食材の推定行を日付単位で合計（複数商品の重複を正しく集計）
    const ingDateMap = {};  // key: ingCode+'@'+dateStr → {qty, dow}
    jissekiSheet.getRange(2,1,jissekiSheet.getLastRow()-1,6).getValues()
      .filter(r => {
        if (!r[0]||!r[2]||r[4]<=0||(r[5]||'').toString().trim()!=='推定') return false;
        const d=r[0] instanceof Date ? r[0] : new Date(r[0]); d.setHours(0,0,0,0);
        return d >= limitO;
      })
      .forEach(row => {
        const d=row[0] instanceof Date ? row[0] : new Date(row[0]); d.setHours(0,0,0,0);
        const ds=Utilities.formatDate(d,'Asia/Tokyo','yyyy/MM/dd');
        const dk=String(row[2])+'@'+ds;
        if (!ingDateMap[dk]) ingDateMap[dk]={qty:0, dow:Number(row[1])};
        ingDateMap[dk].qty+=Number(row[4]);
      });
    // Step2: 日付合計値を曜日単位で平均（典型的な1日使用量を計算）
    const ingEstSum = {}, ingEstCount = {};
    Object.entries(ingDateMap).forEach(([dk, v]) => {
      const ingCode=dk.split('@')[0];
      const key=ingCode+'_'+v.dow;
      if (!ingEstSum[key]) ingEstSum[key]=0;
      if (!ingEstCount[key]) ingEstCount[key]=0;
      ingEstSum[key]+=v.qty; ingEstCount[key]+=1;
    });
    // Step3: 実績が1件もない食材×曜日にのみフォールバック適用
    let estFallbackCount=0;
    Object.keys(ingEstSum).forEach(key => {
      if (!ingDowCount[key]||ingDowCount[key]===0) {
        ingDowSum[key]=ingEstSum[key]; ingDowCount[key]=ingEstCount[key]; estFallbackCount++;
      }
    });
    if (estFallbackCount>0) console.log('推定値フォールバック: '+estFallbackCount+'件（実績なし食材×曜日を推定値で補完）');
  }

  // サンプル2件未満の曜日は全曜日平均で補完
  const MIN_DOW_SAMPLES = 2;
  const ingAllSum = {}, ingAllCnt = {};
  Object.keys(ingDowSum).forEach(key => {
    const code=key.split('_')[0];
    if (!ingAllSum[code]){ ingAllSum[code]=0; ingAllCnt[code]=0; }
    ingAllSum[code]+=ingDowSum[key]; ingAllCnt[code]+=ingDowCount[key];
  });
  const ingDowAvg = {};
  Object.keys(ingDowSum).forEach(key => {
    const code=key.split('_')[0];
    if (ingDowCount[key]>=MIN_DOW_SAMPLES) {
      ingDowAvg[key]=ingDowSum[key]/ingDowCount[key];
    } else if (ingAllCnt[code]>=1) {
      ingDowAvg[key]=ingAllSum[code]/ingAllCnt[code];
      console.log('サンプル不足補完: '+key+'('+ingDowCount[key]+'件) → 全曜日平均'+(ingAllSum[code]/ingAllCnt[code]).toFixed(2));
    } else {
      ingDowAvg[key]=0;
    }
  });

  const data = sheet.getRange(2,1,lastRow-1,10).getValues();
  const toSerial = date => date instanceof Date ? Math.floor(date.getTime()/86400000)+25569 : Number(date);
  const rowMap = {};
  data.forEach((row,i) => { if(row[0]&&row[2]) rowMap[row[0]+'_'+toSerial(row[2])]=i; });

  // ── CP専用食材の「発注終了日」を算出（CP終了後は発注を止める）─────────────────
  //   例：鰻食材はうなぎ茶漬けPR（8/6終了）だけが使う→8/7以降は発注0。
  //   その食材を使う商品が「全て新商品型CP（通常商品CP列が空）の商品」の場合のみ対象。
  //   通常商品でも使う食材（カルビ等が通常メニューにもある場合）は止めない。
  const ingCpEnd = {};
  {
    const ingToProducts = {};
    const pmO = ss.getSheetByName('商品マスタ');
    if (pmO && pmO.getLastRow() > 1) {
      pmO.getRange(2,1,pmO.getLastRow()-1,5).getValues().forEach(r => {
        if (!r[0] || !r[4]) return;
        const ing = String(r[4]).trim(), prod = String(r[0]).trim();
        if (!ingToProducts[ing]) ingToProducts[ing] = {};
        ingToProducts[ing][prod] = true;
      });
    }
    const boundedProductEnd = {};   // 新商品型CP商品 → 最も遅いCP終了日
    const cpSheetE = ss.getSheetByName('CP設定');
    if (cpSheetE && cpSheetE.getLastRow() > 1) {
      cpSheetE.getRange(2,1,cpSheetE.getLastRow()-1,9).getValues().forEach(row => {
        const flag = (row[8] || '').toString().trim();
        if (flag !== '') return;                     // 通常商品CP・切替は対象外（商品は通常メニュー）
        const eRaw = row[2];
        const eDate = eRaw instanceof Date ? eRaw : (eRaw ? new Date(String(eRaw).replace(/^(\d{4})(\d{2})(\d{2})$/, '$1/$2/$3')) : null);
        if (!eDate || isNaN(eDate.getTime())) return; // 終了日未定（継続中）は止めない
        eDate.setHours(0,0,0,0);
        const codes = (typeof row[3] === 'number') ? [Math.round(row[3])]
          : String(row[3]||'').split(',').flatMap(c => { c = c.trim(); const m = c.match(/^(\d+)-(\d+)$/); if (m) { const a=[]; for (let n=+m[1]; n<=+m[2]; n++) a.push(n); return a; } const n=+c; return n>0?[n]:[]; });
        codes.forEach(c => { const k = String(c); if (!boundedProductEnd[k] || eDate > boundedProductEnd[k]) boundedProductEnd[k] = eDate; });
      });
    }
    Object.keys(ingToProducts).forEach(ing => {
      const prods = Object.keys(ingToProducts[ing]);
      if (prods.length === 0) return;
      let maxEnd = null, allBounded = true;
      prods.forEach(p => {
        if (boundedProductEnd[p]) { if (!maxEnd || boundedProductEnd[p] > maxEnd) maxEnd = boundedProductEnd[p]; }
        else allBounded = false;                     // 通常商品でも使う食材は止めない
      });
      if (allBounded && maxEnd) ingCpEnd[ing] = maxEnd;
    });
    if (Object.keys(ingCpEnd).length > 0) {
      console.log('CP専用食材の発注終了日: ' + Object.entries(ingCpEnd).map(([k,v]) => k + '=' + Utilities.formatDate(v,'Asia/Tokyo','MM/dd')).join(', '));
    }
  }

  const coverDowsMap = { 3:[3,4,5], 4:[4,5,6,7], 6:[6,7,1,2], 1:[1,2,3,4] };

  // ── CP専用食材：終了日をカバーできる「最も早い納品日」を求める ─────────────
  //   その納品日に終了日まで（＝終了日の出食数分まで）まとめて発注し、
  //   それ以降の納品日は0にする。→ 終了日ちょうどで在庫が尽き、余りも重複発注も出ない。
  const reachDateByIng = {};
  Object.keys(ingCpEnd).forEach(code => {
    const E = ingCpEnd[code];
    let reach = null;
    data.forEach(row => {
      if (String(row[0]) !== code || !row[2]) return;
      const cd = coverDowsMap[row[3]]; if (!cd) return;           // 納品日のみ
      const dt = row[2] instanceof Date ? new Date(row[2]) : new Date((Number(row[2]) - 25569) * 86400000);
      dt.setHours(0, 0, 0, 0);
      if (dt > E) return;
      const normalEnd = new Date(dt); normalEnd.setDate(dt.getDate() + cd.length - 1);
      if (normalEnd >= E && (!reach || dt < reach)) reach = dt;     // 終了日まで届く最も早い納品日
    });
    if (reach) reachDateByIng[code] = reach;
  });

  const oOutput = data.map(row => {
    const code=row[0], dt=row[2], dow=row[3], trd=row[5], sft=row[9];
    if (!code) return [''];
    const coverDows=coverDowsMap[dow]; if(!coverDows) return [''];
    const startDateObj=dt instanceof Date ? new Date(dt) : new Date((Number(dt)-25569)*86400000);
    startDateObj.setHours(0,0,0,0);
    // CP専用食材：終了日をカバーする納品日より後は発注しない（手前でまとめて発注済み）
    const reachDate = reachDateByIng[String(code)];
    if (reachDate && startDateObj > reachDate) return [0];

    // CP休止・専用食材：納品日(発注日+2)が休止期間に丸ごと入る発注は0にする。
    //   本部ルール＝納品は発注日の2日後(火→木/水→金/金→日/日→火)。納品分が全部休止期間で使われないため。
    //   ※末日納品(再開日前日)は翌日に再開分を使うので0にしない → deliv < 終了日 で判定
    if (pauseWindows.length && pauseExcl.has(String(code))) {
      const deliv = new Date(startDateObj); deliv.setDate(startDateObj.getDate() + 2); deliv.setHours(0, 0, 0, 0);
      if (pauseWindows.some(w => deliv >= w.start && deliv < w.end)) return [0];
    }

let baseTotal = 0;
  coverDows.forEach((d, idx) => {
    const key = code + '_' + d;
    let dayAvg;
    if (ingDowAvg[key] !== undefined) {
      dayAvg = ingDowAvg[key];
    } else if (ingAllCnt[code] >= 1) {
      dayAvg = ingAllSum[code] / ingAllCnt[code];
    } else {
      dayAvg = 0;
    }
    const coverDate = new Date(startDateObj);
    coverDate.setDate(startDateObj.getDate() + idx);
    // CP専用食材：CP終了日を過ぎた日は使用量0（発注を止める）
    const cpEnd = ingCpEnd[String(code)];
    if (cpEnd && coverDate > cpEnd) return;
    // CP休止：休止期間に入る日は減算。専用食材(その2商品しか使わない)は0、共有食材は休止商品ぶんだけ差し引く（0未満は0）
    if (pauseWindows.length && inPauseWindow(coverDate)) {
      if (pauseExcl.has(String(code))) dayAvg = 0;
      else { const ps = pauseSub[String(code) + '_' + d]; if (ps) dayAvg = Math.max(0, dayAvg - ps); }
    }
    const coverDateStr = Utilities.formatDate(coverDate, 'Asia/Tokyo', 'yyyy/MM/dd');
    const eventTag = holidayTagMap[coverDateStr];
    let holidayRate = 1.0;
    if (eventTag) {
      const hrKey = String(code) + '_' + eventTag;
      if (holidayRateMap[hrKey] !== undefined) holidayRate = holidayRateMap[hrKey];
    }
    baseTotal += dayAvg * holidayRate;
  });
    if (baseTotal===0) return [0];

    const base=toSerial(dt), days=coverDows.length;
    let evtSum=0, wtrSum=0, hldSum=0;
    for (let d=0; d<days; d++) {
      const idx=rowMap[code+'_'+(base+d)];
      const coverDate=new Date(startDateObj); coverDate.setDate(startDateObj.getDate()+d);
      const coverDateStr=Utilities.formatDate(coverDate,'Asia/Tokyo','yyyy/MM/dd');
      const eventTag=holidayTagMap[coverDateStr];
      if (idx!==undefined) {
        evtSum+=Number(data[idx][6])||1; wtrSum+=Number(data[idx][7])||1;
        const hrKey=eventTag ? String(code)+'_'+eventTag : null;
        if (hrKey&&holidayRateMap[hrKey]!==undefined) { hldSum+=1.0; }
        else { const hldVal=Number(data[idx][8])||1; hldSum+=(hldVal!==1&&(coverDows[d]===7||coverDows[d]===1))?1.0:hldVal; }
      } else { evtSum+=1; wtrSum+=1; hldSum+=1; }
    }
    const result=baseTotal*(Number(trd)||1)*(evtSum/days)*(wtrSum/days)*(hldSum/days)*(Number(sft)||1);
    return [result];
  });

  sheet.getRange(2,15,oOutput.length,1).setValues(oOutput);
}

function buildOrderBufferO() {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  let cpSubtractMap=loadCpSubtractMap_();
  if (!cpSubtractMap) { cpSubtractMap=buildIngredientDailyUsage(); saveCpSubtractMap_(cpSubtractMap); }
  buildOrderBufferO_(null, cpSubtractMap);
  console.log('O列の再計算が完了しました');
}

// ─────────────────────────────────────────────────────────────────────
// Step4: グレーアウト
// ─────────────────────────────────────────────────────────────────────

function updateOrderSheetColor_() {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const orderSheet=ss.getSheetByName('発注ワークシート'), masterSheet=ss.getSheetByName('食材マスタ');
  if (!orderSheet||!masterSheet) return;
  const categoryMap={};
  if (masterSheet.getLastRow()>1) masterSheet.getRange(2,1,masterSheet.getLastRow()-1,3).getValues().forEach(row=>{ if(row[0]) categoryMap[String(row[0])]=(row[2]||'').toString().trim(); });
  const GRAY='#b7b7b7', WHITE='#ffffff';
  const DOW_CONFIG=[{dow:3,codeCol:1,startCol:1,colCount:4},{dow:4,codeCol:6,startCol:6,colCount:4},{dow:6,codeCol:11,startCol:11,colCount:4},{dow:1,codeCol:16,startCol:16,colCount:4}];
  const orderLastRow=orderSheet.getLastRow(); if(orderLastRow<3) return;
  const dataRows=orderLastRow-2;
  DOW_CONFIG.forEach(cfg => {
    const codes=orderSheet.getRange(3,cfg.codeCol,dataRows,1).getValues().map(r=>String(r[0]));
    const bgColors=codes.map(code => {
      if (!code||code==='') return Array(cfg.colCount).fill(WHITE);
      const category=categoryMap[code]||'';
      if (category==='発注終了'||category==='消耗品') return Array(cfg.colCount).fill(GRAY);
      if (category==='ドリンク') return (cfg.dow===6||cfg.dow===1) ? Array(cfg.colCount).fill(GRAY) : Array(cfg.colCount).fill(WHITE);
      if (category==='酒') return (cfg.dow===3||cfg.dow===6||cfg.dow===1) ? Array(cfg.colCount).fill(GRAY) : Array(cfg.colCount).fill(WHITE);
      return Array(cfg.colCount).fill(WHITE);
    });
    orderSheet.getRange(3,cfg.startCol,dataRows,cfg.colCount).setBackgrounds(bgColors);
    // 罫線：まずブロック全体の罫線を消し、実データ(コード)がある最終行までだけ格子罫線を引く。
    // これで食材が増減しても最終行まで追従し、下の空行には罫線が残らない。
    let lastIdx = -1;
    for (let i = 0; i < codes.length; i++) { if (codes[i] !== '' && codes[i] !== 'null') lastIdx = i; }
    orderSheet.getRange(1,cfg.startCol,orderLastRow,cfg.colCount)
              .setBorder(false,false,false,false,false,false);
    if (lastIdx >= 0) {
      orderSheet.getRange(1,cfg.startCol,3+lastIdx,cfg.colCount)
                .setBorder(true,true,true,true,true,true,'#000000',SpreadsheetApp.BorderStyle.SOLID);
    }
  });
  console.log('発注ワークシートのグレーアウトと罫線を更新しました');
  try { refreshOrderSheetWide_(); } catch(e) { console.log('横型の自動更新をスキップ: ' + e.message); }  // 横型も毎朝ここで自動連動
  SpreadsheetApp.flush();
}

function updateOrderSheetColor() { updateOrderSheetColor_(); }

// ─────────────────────────────────────────────────────────────────────
// Step5: ヘッダー
// ─────────────────────────────────────────────────────────────────────

function updateOrderSheetHeader_(startDate) {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const orderSheet=ss.getSheetByName('発注ワークシート'); if(!orderSheet) return;
  const tueDt=new Date(startDate);
  const wedDt=new Date(startDate); wedDt.setDate(startDate.getDate()+1);
  const friDt=new Date(startDate); friDt.setDate(startDate.getDate()+3);
  const sunDt=new Date(startDate); sunDt.setDate(startDate.getDate()+5);
  const makeHeader=dt=>formatDate_(dt)+' '+getDowName_(dt.getDay()+1)+'曜日発注';
  orderSheet.getRange('A1').setValue(makeHeader(tueDt));
  orderSheet.getRange('F1').setValue(makeHeader(wedDt));
  orderSheet.getRange('K1').setValue(makeHeader(friDt));
  orderSheet.getRange('P1').setValue(makeHeader(sunDt));
}

function updateOrderSheetHeader() {
  const today=new Date(), todayCode=today.getDay()+1;
  let diff=(3-todayCode+7)%7; if(diff===0) diff=7;
  const startDate=new Date(today); startDate.setDate(today.getDate()+diff);
  updateOrderSheetHeader_(startDate); console.log('ヘッダーを更新しました');
}

// 発注ワークシートの余分な空行を削除して印刷を1ページに収まりやすくする。
// 食材行は行3のARRAYFORMULA（食材マスタから消耗品を除いて自動抽出）でスピルするため、
// データ末尾＋余白ぶんだけ残して以降を削除する。食材が増えても余白内なら自動対応。
// ※GASには印刷設定（1ページに収める）を変えるAPIが無いので、行の掃除まで行う。
//   最後の「1ページに収める」だけは印刷画面で設定する（一度設定すれば記憶される）。
function trimOrderSheetRows_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('発注ワークシート');
  if (!sh) return;
  const MARGIN = 20;                                   // 食材が増えてもスピルできる余白
  const lastData = Math.max(sh.getLastRow(), 3);       // ヘッダー2行＋食材行の末尾
  const desiredMax = lastData + MARGIN;
  const curMax = sh.getMaxRows();
  if (curMax > desiredMax) sh.deleteRows(desiredMax + 1, curMax - desiredMax);
  console.log('発注ワークシートの空行を整理しました（' + curMax + '行 → ' + Math.min(curMax, desiredMax) + '行）');
}

function trimOrderSheetRows() { trimOrderSheetRows_(); }

// 発注ワークシートの各ブロックで「在庫」列と「基準」列を入れ替える。
//  ・4ブロックの (在庫列, 基準列) = (C,D)(H,I)(M,N)(R,S)。基準は行3の1つのスピル数式、在庫は手入力の空列。
//  ・基準の数式は自分のブロックのコード列(A/F/K/P)を参照しているので、列を移動しても壊れない。
//  ・GASの色付け(updateOrderSheetColor_)はコード列基準で4列一括なので位置に依存しない → 影響なし。
//  ・純粋な入れ替えなので、もう一度実行すると元に戻る（可逆）。※在庫を手入力する前に実行すること。
function swapOrderSheetStockPar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('発注ワークシート');
  if (!sh) { try { SpreadsheetApp.getUi().alert('発注ワークシートが見つかりません'); } catch (e) {} return; }
  const pairs = [[3, 4], [8, 9], [13, 14], [18, 19]];  // [在庫列, 基準列]（1=A,2=B,...）
  pairs.forEach(([c1, c2]) => {
    // 行3の数式を入れ替え（先に両方読んでから書く＝上書き事故を防ぐ）
    const f1 = sh.getRange(3, c1).getFormula();
    const f2 = sh.getRange(3, c2).getFormula();
    if (f2) sh.getRange(3, c1).setFormula(f2); else sh.getRange(3, c1).clearContent();
    if (f1) sh.getRange(3, c2).setFormula(f1); else sh.getRange(3, c2).clearContent();
    // 行2のヘッダー（在庫/基準）も入れ替え
    const h1 = sh.getRange(2, c1).getValue();
    const h2 = sh.getRange(2, c2).getValue();
    sh.getRange(2, c1).setValue(h2);
    sh.getRange(2, c2).setValue(h1);
  });
  try { SpreadsheetApp.getUi().alert('各ブロックの「在庫」と「基準」を入れ替えました。\nもう一度実行すると元に戻ります。'); } catch (e) { console.log('在庫と基準を入れ替えました'); }
}
function swapOrderSheetStockPar_() { swapOrderSheetStockPar(); }

// 発注ワークシートの「基準」列を、切り上げ前の必要量で表示する数式に置き換える。
//  ・値＝発注計算バッファO列(カバー日数の必要量) ÷ lot(納品量G÷規格量F) ＝切り上げ前の必要発注量。
//  ・1以上は整数(四捨五入)、1未満は0.1単位で表示（少量品ほど「実はほとんど要らない」が数字で分かる）。
//  ・基準列は各ブロック内で見出しが「基準」の列を自動検出（在庫と入れ替えていても対応）。
function setOrderSheetParFormula_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('発注ワークシート');
  if (!sh) { try { SpreadsheetApp.getUi().alert('発注ワークシートが見つかりません'); } catch(e){} return; }
  const blocks = [
    { codeCol: 'A', start: 1,  dow: 3 },   // 火曜発注
    { codeCol: 'F', start: 6,  dow: 4 },   // 水曜発注
    { codeCol: 'K', start: 11, dow: 6 },   // 金曜発注
    { codeCol: 'P', start: 16, dow: 1 }    // 日曜発注
  ];
  const maxR = sh.getMaxRows();
  let done = 0;
  blocks.forEach(b => {
    let parCol = -1;
    for (let c = b.start; c <= b.start + 3; c++) {
      if (String(sh.getRange(2, c).getValue()).trim() === '基準') { parCol = c; break; }
    }
    if (parCol < 0) return;
    const A = b.codeCol;
    const f =
      '=ARRAYFORMULA(IF(' + A + '3:' + A + '="","",IFERROR(LET(v,'
      + 'VLOOKUP(' + A + '3:' + A + "&MINIFS('発注計算バッファ'!C:C,'発注計算バッファ'!D:D," + b.dow + '),'
      + "{'発注計算バッファ'!A:A&'発注計算バッファ'!C:C,'発注計算バッファ'!O:O},2,FALSE)"
      + '/(VLOOKUP(' + A + '3:' + A + ",'食材マスタ'!A:G,7,FALSE)/VLOOKUP(" + A + '3:' + A + ",'食材マスタ'!A:F,6,FALSE)),"
      + 'IF(v>=1,ROUND(v,0),ROUND(v,1))),"")))';
    sh.getRange(3, parCol).setFormula(f);
    if (maxR >= 3) sh.getRange(3, parCol, maxR - 2, 1).setNumberFormat('#,##0.#');  // 1以上=整数, 1未満=0.1単位
    done++;
  });
  try { ss.toast('基準を「1以上は整数・1未満は0.1単位」の表示にしました（' + done + 'ブロック）', '基準の表示を変更', 6); } catch(e){}
  console.log('setOrderSheetParFormula_: ' + done + 'ブロックの基準列を更新');
}
function setOrderSheetParFormula() { setOrderSheetParFormula_(); }

// 仕込み・解凍表（週間・A4横1枚）。食材×曜日で、その曜日の必要量を発注の使用量データ(トレンド履歴)から自動表示。
// 全食材とも冷蔵・前日夜解凍。必要量＝曜日別平均使用量×規格量÷小分け単位。
function buildPrepThawSheet_(weekOffset) {
  if (weekOffset === undefined) weekOffset = 7;   // 既定＝来週分（自動作成・フォーム操作時もこれ）。0＝今週分
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = '仕込み・解凍表';
  let sh = ss.getSheetByName(name);
  ensurePrepMaster_();   // 品目の保存場所（無ければ作成＆初期投入）
  ensurePrepAdjust_();   // 数量調整の保存場所
  if (sh) {
    sh.clear();
    sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
    sh.showRows(1, sh.getMaxRows());          // 前回の版で隠れた行（例：ディナー行）を必ず表示に戻す
    sh.showColumns(1, sh.getMaxColumns());     // clearだけでは非表示は解除されないため明示的に戻す
  } else {
    sh = ss.insertSheet(name);
  }
  sh.setFrozenRows(0); sh.setFrozenColumns(0);   // 前回の固定を解除してから結合（固定線と結合の衝突を回避）
  const items = readPrepItems_();
  const dows   = ['月', '火', '水', '木', '金', '土', '日'];
  const trCols = [4, 5, 6, 7, 8, 9, 3];   // トレンド履歴の列(月=D..土=I,日=C)
  const dowNumByDi = [2, 3, 4, 5, 6, 7, 1];  // di(月..日)→曜日番号(日=1..土=7・buildCpPauseSubMap_と一致)
  // CP休止：休止期間に入る日は、休止商品ぶんの食材使用量をトレンド値から差し引く（規格量単位）。専用食材は0にする。
  const pauseData = buildCpPauseSubMap_(ss);
  const pauseWindowsPrep = pauseData.windows, pauseSubPrep = pauseData.sub, pauseExclPrep = pauseData.exclusive || new Set();
  // 日付セル(A1)が休止期間内かの判定式。休止が無ければnull（数式に何も足さない）。
  const pauseCondExpr = function(cellA1) {
    if (!pauseWindowsPrep.length) return null;
    const parts = pauseWindowsPrep.map(function(w) {
      const s = w.start, e = w.end;
      return 'AND(' + cellA1 + '>=DATE(' + s.getFullYear() + ',' + (s.getMonth() + 1) + ',' + s.getDate() + '),'
        + cellA1 + '<=DATE(' + e.getFullYear() + ',' + (e.getMonth() + 1) + ',' + e.getDate() + '))';
    });
    return parts.length === 1 ? parts[0] : 'OR(' + parts.join(',') + ')';
  };
  // トレンド項：VLOOKUP(code,トレンド,col)。休止対象食材なら「MAX(0, 値 - IF(休止日,減算,0))」に置換。
  const vTerm = function(codeExpr, code, col, dow, dateCellA1) {
    const v = "VLOOKUP(" + codeExpr + ",'トレンド履歴'!$A:$I," + col + ",FALSE)";
    const cond = pauseCondExpr(dateCellA1);
    if (!cond) return v;
    if (pauseExclPrep.has(String(code))) return "IF(" + cond + ",0," + v + ")";  // 専用食材：休止日は0
    const ps = pauseSubPrep[String(code) + '_' + dow];
    if (ps && ps > 0) return "MAX(0," + v + "-IF(" + cond + "," + ps + ",0))";   // 共有食材：休止商品ぶん減算
    return v;
  };
  // 列: A=コード(隠) B=食材名 C=期限 D=単位(ルール) E=単位g(隠) F=単位(g/袋・右寄せ) G=ランチ/ディナー H..N=曜日
  const dayStart = 8;
  const totalCols = dayStart - 1 + dows.length;   // 7 + 7 = 14
  // 行: 1=表題 2=説明 3=日付 4=曜日 5=予測 6..=データ（1食材2段：上=ランチ 下=ディナー）
  const rTitle = 1, rDesc = 2, rDate = 3, rDow = 4, rFcst = 5, rData = 6;
  const lastRow = rData - 1 + items.length * 2;

  // 表題
  sh.getRange(rTitle, 1, 1, totalCols).merge();
  sh.getRange(rTitle, 1).setValue('仕込み・解凍表')
    .setFontWeight('bold').setFontSize(16).setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(rTitle, 32);
  // 説明
  sh.getRange(rDesc, 1, 1, totalCols).merge();
  const weekLabel = (weekOffset === 0) ? '今週' : '来週';
  sh.getRange(rDesc, 1).setValue(
    '★この表は「' + weekLabel + '（月〜日）」の予定です。自動作成では来週分になります（今週分はメニューから作成）。　'
    + '■ランチ（上段）＝その日＋翌日の半分（1.5日分）の仕込み目安。フォームの「数量調整」で手入力補正できます。　'
    + '■ディナー（下段）＝その日の半分（夜の分）の目安（自動計算）。　'
    + '※前日夜に翌日分を冷蔵解凍。数字は袋・バット等の単位です。')
    .setFontSize(10).setWrap(true).setVerticalAlignment('middle');
  sh.setRowHeight(rDesc, 42);

  // 左見出し（日付〜予測の3行を縦結合）。F列(単位g/袋)も「単位」ヘッダーで結合。
  const leftHeads = { 2: '食材', 3: '期限', 4: '単位', 6: '単位' };
  Object.keys(leftHeads).forEach(function(col) {
    const c = Number(col);
    sh.getRange(rDate, c, 3, 1).merge();
    sh.getRange(rDate, c).setValue(leftHeads[col]).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
  });
  // G列（ランチ/ディナーの上）に「日付」「曜日」「予測」の行ラベル
  sh.getRange(rDate, 7).setValue('日付').setFontSize(8).setFontColor('#888888').setHorizontalAlignment('right').setVerticalAlignment('middle');
  sh.getRange(rDow, 7).setValue('曜日').setFontSize(8).setFontColor('#888888').setHorizontalAlignment('right').setVerticalAlignment('middle');
  sh.getRange(rFcst, 7).setValue('予測').setFontSize(8).setFontColor('#888888').setHorizontalAlignment('right').setVerticalAlignment('middle');
  // 曜日行
  sh.getRange(rDow, dayStart, 1, dows.length).setValues([dows]).setFontWeight('bold').setHorizontalAlignment('center');
  // 日付行・売上予測行（数式）
  dows.forEach(function(d, di) {
    const col = dayStart + di;
    sh.getRange(rDate, col).setFormula('=TODAY()-WEEKDAY(TODAY(),3)+' + (weekOffset + di))   // 月曜+weekOffset+di（既定=来週7・今週0）
      .setNumberFormat('M/d').setHorizontalAlignment('center').setFontColor('#666666');
    const dc = sh.getRange(rDate, col).getA1Notation();
    sh.getRange(rFcst, col).setFormula("=IFERROR(VLOOKUP(" + dc + ",'売上予測'!$A:$I,9,FALSE),\"\")")
      .setNumberFormat('#,##0').setHorizontalAlignment('center').setFontSize(9).setFontColor('#666666');
  });

  // データ（1食材2段：上=ランチ＝その日＋翌日の半分／下=ディナー＝その日の半分）
  items.forEach(function(it, i) {
    const top = rData + i * 2, bot = top + 1;
    const code = it.partsCodes ? it.partsCodes[0] : it.code;
    const unit = it.partsCodes ? it.perUnitG : it.unitG;
    sh.getRange(top, 1).setValue(code); sh.getRange(bot, 1).setValue(code);   // コードは両段に
    sh.getRange(top, 2, 1, 5).setValues([[it.name, '', it.rule, unit, it.uname]]);  // B食材 C期限 D単位 E単位g F単位名
    sh.getRange(top, 3).setFormula("=IFERROR(VLOOKUP($A" + top + ",'仕込みマスタ'!$A:$I,9,FALSE),\"\")");  // 期限（テキスト）
    [2, 3, 4, 6].forEach(function(c) { sh.getRange(top, c, 2, 1).merge(); });  // 食材名・期限・単位・単位名は2段結合
    sh.getRange(top, 7).setValue('ランチ').setHorizontalAlignment('center').setVerticalAlignment('middle').setFontSize(9).setFontWeight('bold').setFontColor('#1a56db');
    sh.getRange(bot, 7).setValue('ディナー').setHorizontalAlignment('center').setVerticalAlignment('middle').setFontSize(9).setFontWeight('bold').setFontColor('#c0392b');
    trCols.forEach(function(trCol, di) {
      const nextCol = trCols[(di + 1) % 7];  // 翌日の列（土→日、日→月で1周）
      const step = it.step || 1;
      const col = dayStart + di;
      const dcAbs = sh.getRange(rDate, col).getA1Notation().replace(/(\d+)/, '$$$1');  // 例 H$3
      const dcTom = "(" + dcAbs + "+1)";                       // 翌日の日付（休止判定用）
      const dowToday = dowNumByDi[di], dowTom = dowNumByDi[(di + 1) % 7];
      const adj = "+SUMIFS('仕込み調整'!$C:$C,'仕込み調整'!$B:$B,$A" + top + ",'仕込み調整'!$A:$A," + dcAbs + ")";
      let lunch, dinner;
      if (it.partsCodes) {
        const lp = it.partsCodes.map(function(c) {
          return "(" + vTerm(c, c, trCol, dowToday, dcAbs) + "+" + vTerm(c, c, nextCol, dowTom, dcTom) + "/2)*VLOOKUP(" + c + ",'食材マスタ'!$A:$F,6,FALSE)";
        }).join('+');
        const dp = it.partsCodes.map(function(c) {
          return vTerm(c, c, trCol, dowToday, dcAbs) + "*0.5*VLOOKUP(" + c + ",'食材マスタ'!$A:$F,6,FALSE)";
        }).join('+');
        lunch = "CEILING((" + lp + ")/" + it.perUnitG + "," + step + ")";
        dinner = "CEILING((" + dp + ")/" + it.perUnitG + "," + step + ")";
      } else {
        lunch = "CEILING((" + vTerm("$A" + top, it.code, trCol, dowToday, dcAbs) + "+" + vTerm("$A" + top, it.code, nextCol, dowTom, dcTom) + "/2)*VLOOKUP($A" + top + ",'食材マスタ'!$A:$F,6,FALSE)/$E" + top + "," + step + ")";
        dinner = "CEILING(" + vTerm("$A" + bot, it.code, trCol, dowToday, dcAbs) + "*0.5*VLOOKUP($A" + bot + ",'食材マスタ'!$A:$F,6,FALSE)/$E" + top + "," + step + ")";
      }
      sh.getRange(top, col).setFormula("=IFERROR(" + lunch + ",0)" + adj);   // ランチ（数量調整もここに反映）
      sh.getRange(bot, col).setFormula("=IFERROR(" + dinner + ",0)");         // ディナー（自動計算のみ）
    });
  });
  // 体裁
  const nData = items.length * 2;
  sh.getRange(rData, 2, nData, 1).setFontSize(12).setFontWeight('bold').setVerticalAlignment('middle');  // 食材名
  sh.getRange(rData, 3, nData, 1).setHorizontalAlignment('center').setVerticalAlignment('middle').setNumberFormat('@');  // 期限
  sh.getRange(rData, 4, nData, 1).setWrap(true).setVerticalAlignment('middle');                          // 単位(ルール)
  sh.getRange(rData, 6, nData, 1).setHorizontalAlignment('right').setVerticalAlignment('middle');        // F列 単位(g/袋)は右寄せ
  sh.getRange(rData, dayStart, nData, dows.length).setHorizontalAlignment('center').setFontWeight('bold').setVerticalAlignment('middle');
  sh.getRange(rDate, 1, lastRow - rDate + 1, totalCols).setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  // 食材ごとの2段（B..N）を太枠で囲んで区切る
  items.forEach(function(it, i) {
    const top = rData + i * 2;
    sh.getRange(top, 2, 2, totalCols - 1).setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_THICK);
  });
  sh.setColumnWidth(2, 190); sh.setColumnWidth(3, 72); sh.setColumnWidth(4, 200); sh.setColumnWidth(6, 48); sh.setColumnWidth(7, 58);
  dows.forEach(function(d, di) { sh.setColumnWidth(dayStart + di, 46); });
  sh.setRowHeights(rDate, 3, 17);                // 日付・曜日・予測の3行は半分の高さ
  sh.setRowHeights(rData, nData, 24);            // 各段24（1食材48）
  // 土=薄い青・日=薄い赤（列は固定なので常に正しい）＋曜日文字も色分け
  const satCol = dayStart + 5, sunCol = dayStart + 6, nRows = lastRow - rDate + 1;
  sh.getRange(rDate, satCol, nRows, 1).setBackground('#e8f0fe');
  sh.getRange(rDate, sunCol, nRows, 1).setBackground('#fce8e6');
  sh.getRange(rDow, satCol).setFontColor('#1a56db');  // 土曜=青
  sh.getRange(rDow, sunCol).setFontColor('#c0392b');  // 日曜=赤
  // 祝日（祝日マスタでisOff=TRUE）の列を薄い赤に。日付が変わっても自動で追従。
  // ※条件付き書式は別シート参照が禁止なので、隠し行に祝日フラグ(1/0)を計算し、その同一シートのセルを見る。
  const rHelp = lastRow + 1;
  dows.forEach(function(d, di) {
    const col = dayStart + di;
    const dcAbs = sh.getRange(rDate, col).getA1Notation();
    sh.getRange(rHelp, col).setFormula("=IF(SUMPRODUCT(('祝日マスタ'!$A$2:$A$800=" + dcAbs + ")*'祝日マスタ'!$C$2:$C$800)>0,1,0)");
  });
  const helpAnchor = sh.getRange(rHelp, dayStart).getA1Notation().replace(/(\d+)/, '$$$1');
  // 赤くするのは月〜金の5列だけ（土=青・日=赤は固定色で残す。祝日マスタは土日も休み扱いなので範囲を絞る）
  const holRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=" + helpAnchor + "=1")
    .setBackground('#fce8e6').setFontColor('#c0392b')
    .setRanges([sh.getRange(rDate, dayStart, nRows, 5)])
    .build();
  sh.setConditionalFormatRules([holRule]);
  sh.hideColumns(1); sh.hideColumns(5);          // コード・単位gは計算用なので隠す
  sh.hideRows(rHelp);                            // 祝日フラグの計算行は隠す
  sh.setFrozenRows(rFcst);   // 見出し（表題〜予測の5行）を固定
  const maxR = sh.getMaxRows(); if (maxR > rHelp) sh.deleteRows(rHelp + 1, maxR - rHelp);
  const maxC = sh.getMaxColumns(); if (maxC > totalCols) sh.deleteColumns(totalCols + 1, maxC - totalCols);
  try { ss.toast('仕込み・解凍表（' + weekLabel + '分）を作り直しました', '完成', 8); } catch(e){}
  console.log('buildPrepThawSheet_: 2段作成完了 weekOffset=' + weekOffset);
}
function buildPrepThawSheet() { buildPrepThawSheet_(7); }            // メニュー：来週分（自動と同じ）
function buildPrepThawSheetThisWeek() { buildPrepThawSheet_(0); }    // メニュー：今週分を作り直す

// 主要9品目の安全係数を最適値に設定（食材マスタI列）。
//   実データの在庫シミュレーション(補充方式・賞味期限で持ち越し)で、欠品36%→12%・合計ロス12.2→9.2%が最良だったSL85%の値。
//   全て日持ち食材なので上げても廃棄はほぼ増えない。生もの・当日消費品には使わないこと。
function setPrepSafetyFactors_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const m = ss.getSheetByName('食材マスタ');
  if (!m || m.getLastRow() < 2) return { n: 0, lines: [] };
  const rec = { 612087: 1.4, 612089: 1.4, 612047: 1.4, 612072: 1.5, 612098: 1.6, 612103: 1.5, 613065: 1.6, 620099: 1.5, 614047: 1.5, 613062: 1.6, 620081: 1.5 };
  const lr = m.getLastRow();
  const codes = m.getRange(2, 1, lr - 1, 1).getValues();
  const lines = [];
  let n = 0;
  for (let i = 0; i < codes.length; i++) {
    const c = Number(codes[i][0]);
    if (rec[c] !== undefined) {
      const cell = m.getRange(i + 2, 9);   // I列＝安全係数
      const old = cell.getValue();
      cell.setValue(rec[c]);
      lines.push(c + '（' + m.getRange(i + 2, 2).getValue() + '）：' + old + ' → ' + rec[c]);
      n++;
    }
  }
  try { ss.toast(n + '品目の安全係数を更新しました', '安全係数', 6); } catch (e) {}
  return { n: n, lines: lines };
}
function setPrepSafetyFactors() {
  const r = setPrepSafetyFactors_();
  SpreadsheetApp.getUi().alert('主要9品目の安全係数を最適値に更新しました（' + r.n + '品目）\n\n' + r.lines.join('\n') + '\n\n翌朝の発注計算で自動反映されます。今すぐ反映するには「食材マスタを区分順・コード順に並べ替える」を実行してください。');
}

// コーン(615004)を「缶」単位に直す。1缶=1810g（1ケース6缶=10860g＝発注1単位）。仕込みは缶数で表示。
function fixCornCanUnit_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const m = ss.getSheetByName('仕込みマスタ');
  if (!m) throw new Error('仕込みマスタが見つかりません');
  const code = 615004;
  const canG = 1810;   // 1缶のg
  const mv = m.getRange(2, 1, m.getLastRow() - 1, 1).getValues();
  let done = false;
  for (let i = 0; i < mv.length; i++) {
    if (Number(mv[i][0]) === code) {
      m.getRange(i + 2, 3).setValue('1缶=1810g（1ケース6缶）');  // C 単位(ルール)
      m.getRange(i + 2, 4).setValue(canG);                       // D 単位g ＝ 1缶のg
      m.getRange(i + 2, 5).setValue('缶');                        // E 単位名
      m.getRange(i + 2, 6).setValue(1);                          // F 切上単位
      done = true; break;
    }
  }
  if (!done) throw new Error('仕込みマスタにコーン(615004)が見つかりません');
  buildPrepThawSheet_();
  return canG;
}
function fixCornCanUnit() {
  fixCornCanUnit_();
  SpreadsheetApp.getUi().alert('コーンを「缶」単位にしました（1缶＝1810g）。\n仕込み・解凍表は缶数で表示されます（発注は1ケース6缶のまま）。');
}

// 仕込みマスタの重複（同コードが複数「有効」）を整理。各コードは先頭1行だけ残す。
//   ついでにブロッコリー(616661)を「300g単位で切り上げ」に統一（過去に二重設定されたため）。
function fixPrepMasterDedupe_() {
  const m = ensurePrepMaster_();
  const lr = m.getLastRow();
  if (lr < 2) return 0;
  const vals = m.getRange(2, 1, lr - 1, 12).getValues();
  const seen = {};
  var removed = 0;
  for (var i = 0; i < vals.length; i++) {
    const code = String(vals[i][0]); if (!code) continue;
    if (!(vals[i][11] === true || vals[i][11] === 'TRUE')) continue;
    if (seen[code]) { m.getRange(i + 2, 12).setValue(false); removed++; }
    else seen[code] = i + 2;
  }
  if (seen['616661']) {   // ブロッコリーを300g単位に統一
    const r = seen['616661'];
    m.getRange(r, 3).setValue('300g単位で切り上げ');
    m.getRange(r, 4).setValue(1);    // 単位g
    m.getRange(r, 5).setValue('g');  // 単位名
    m.getRange(r, 6).setValue(300);  // 切上単位
  }
  buildPrepThawSheet_();
  return removed;
}
function fixPrepMasterDedupe() {
  const n = fixPrepMasterDedupe_();
  SpreadsheetApp.getUi().alert('仕込みマスタの重複を整理しました（重複 ' + n + ' 行を外しました）。\nブロッコリーは「300g単位で切り上げ」に直しました。');
}

// ─────────────────────────────────────────────────────────────────────
// 仕込み・解凍表の 保存場所（仕込みマスタ／仕込み調整）と 入力フォーム
// ─────────────────────────────────────────────────────────────────────

// 初期品目 [コード, 食材名, 単位(ルール文), 単位g, 単位名, 切上単位, 合算コード, 1単位g]
function PREP_ITEMS_DEFAULT_() {
  return [
    ['612087', 'クロ・バラ',              '【A】【B】US産バラ各1袋＋クロッド1kg＝2バット', '',    'バット', 1,   '612087+612089+612047', 1000],
    ['612072', 'US産ラウンドオイスター',  '1.5kg＝1袋',        1500, '袋', 1,   '', ''],
    ['612098', 'US産4mmバラスライス',     '500g単位で切り上げ', 1,    'g',  500, '', ''],
    ['612103', '牛カルビ(ブリスケ)',      '300g単位で切り上げ', 1,    'g',  300, '', ''],
    ['613065', '焼成ハンバーグ',          '枚数で解凍',         1,    '枚', 1,   '', ''],
    ['620099', '冷凍ほうれん草IQF',       '50g＝1で小分け',     50,   '個', 1,   '', ''],
    ['614047', 'イタリアン野菜ミックス',  '20g＝1で小分け',     20,   '個', 1,   '', ''],
    ['613062', 'チキンポップコーン',      '90g＝1で小分け',     90,   '個', 1,   '', ''],
    ['620081', 'さつまいもスティック',    '100g単位で切り上げ', 1,    'g',  100, '', '']
  ];
}

// 仕込みマスタ（品目の保存場所）。無ければ作成し初期品目を入れる。隠しシート。
function ensurePrepMaster_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const head = ['コード', '食材名', '単位(ルール)', '単位g', '単位名', '切上単位', '合算コード', '1単位g', '賞味期限', '在庫', '表示順', '有効'];
  let m = ss.getSheetByName('仕込みマスタ');
  const seed = function() {
    const def = PREP_ITEMS_DEFAULT_();
    const rows = def.map(function(r, i) { return r.concat(['', '', (i + 1) * 10, true]); });
    m.getRange(2, 1, rows.length, head.length).setValues(rows);
  };
  if (!m) {
    m = ss.insertSheet('仕込みマスタ');
    m.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    seed();
    m.setFrozenRows(1);
    try { m.hideSheet(); } catch (e) {}
  } else if (m.getLastRow() < 2) {
    seed();
  }
  return m;
}

// 仕込み調整（日ごとの手直し）の保存場所。隠しシート。
function ensurePrepAdjust_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let a = ss.getSheetByName('仕込み調整');
  if (!a) {
    a = ss.insertSheet('仕込み調整');
    a.getRange(1, 1, 1, 6).setValues([['日付', 'コード', '増減量', '食材名', 'メモ', '登録日時']]).setFontWeight('bold');
    a.setFrozenRows(1);
    try { a.hideSheet(); } catch (e) {}
  }
  return a;
}

// 仕込みマスタから有効な品目を表示順で読む
function readPrepItems_() {
  const m = ensurePrepMaster_();
  const lr = m.getLastRow();
  if (lr < 2) return [];
  const vals = m.getRange(2, 1, lr - 1, 12).getValues();
  const items = vals.filter(function(r) { return r[11] === true || r[11] === 'TRUE'; }).map(function(r) {
    return {
      code: r[0], name: r[1], rule: r[2], unitG: r[3], uname: r[4],
      step: Number(r[5]) || 1,
      partsCodes: r[6] ? String(r[6]).split('+').map(function(x) { return Number(String(x).trim()); }) : null,
      perUnitG: Number(r[7]) || 0,
      order: Number(r[10]) || 0
    };
  });
  items.sort(function(a, b) { return a.order - b.order; });
  return items;
}

function prepFindRow_(m, code) {
  const lr = m.getLastRow();
  if (lr < 2) return -1;
  const col = m.getRange(2, 1, lr - 1, 1).getValues();
  for (var i = 0; i < col.length; i++) { if (String(col[i][0]) === String(code)) return i + 2; }
  return -1;
}
function prepNameOf_(m, code) { const r = prepFindRow_(m, code); return r < 0 ? '' : m.getRange(r, 2).getValue(); }
function prepFmtDate_(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy/MM/dd');
  return String(v);
}

// ===== 入力フォーム =====
function showPrepForm() {
  const html = HtmlService.createHtmlOutput(PREP_FORM_HTML_()).setWidth(560).setHeight(660);
  SpreadsheetApp.getUi().showModalDialog(html, '仕込み・解凍表の入力');
}

// フォーム初期データ（品目リスト＋今週の日付）
function prepForm_getData() {
  const m = ensurePrepMaster_();
  const lr = m.getLastRow();
  const items = [];
  if (lr >= 2) {
    const vals = m.getRange(2, 1, lr - 1, 12).getValues();
    vals.forEach(function(r) {
      if (r[11] === true || r[11] === 'TRUE') {
        items.push({ code: String(r[0]), name: r[1], expiry: (r[8] === '' || r[8] == null ? '' : String(r[8])), order: Number(r[10]) || 0 });
      }
    });
    items.sort(function(a, b) { return a.order - b.order; });
  }
  const days = [];
  const names = ['月', '火', '水', '木', '金', '土', '日'];
  const base = new Date(); base.setHours(0, 0, 0, 0);
  const wd = (base.getDay() + 6) % 7;  // 月=0
  const mon = new Date(base); mon.setDate(base.getDate() - wd);
  for (var i = 0; i < 7; i++) {
    var d = new Date(mon); d.setDate(mon.getDate() + i);
    days.push({ value: Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy/MM/dd'), label: (d.getMonth() + 1) + '/' + d.getDate() + '(' + names[i] + ')' });
  }
  // 食材マスタの一覧（追加フォームのプルダウン用）
  const fs = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('食材マスタ');
  const foods = [];
  if (fs) {
    const flr = fs.getLastRow();
    if (flr >= 2) {
      fs.getRange(2, 1, flr - 1, 2).getValues().forEach(function(r) {
        if (r[0] !== '' && r[1] !== '') foods.push({ code: String(r[0]), name: String(r[1]) });
      });
    }
  }
  return { items: items, days: days, foods: foods };
}

// 期限を保存（当日/24時間/48時間/60時間/3日 のテキスト）
function prepForm_saveExpiry(code, kigen) {
  const m = ensurePrepMaster_();
  const r = prepFindRow_(m, code);
  if (r < 0) throw new Error('食材が見つかりません');
  m.getRange(r, 9).setValue(kigen || '');
  return '期限を保存しました';
}

// 数量調整を保存（同じ日付＋食材は上書き。0で取り消し）
function prepForm_saveAdjust(code, dateStr, delta, memo) {
  const a = ensurePrepAdjust_();
  const m = ensurePrepMaster_();
  const nm = prepNameOf_(m, code);
  const d = new Date(String(dateStr).replace(/-/g, '/')); d.setHours(0, 0, 0, 0);
  const lr = a.getLastRow();
  if (lr >= 2) {
    const vals = a.getRange(2, 1, lr - 1, 2).getValues();
    for (var i = vals.length - 1; i >= 0; i--) {
      var rd = vals[i][0], rc = String(vals[i][1]);
      if (rc === String(code) && rd instanceof Date && rd.getTime() === d.getTime()) a.deleteRow(i + 2);
    }
  }
  const val = Number(delta) || 0;
  if (val !== 0) a.appendRow([d, String(code), val, nm, memo || '', new Date()]);
  return '数量調整を保存しました（' + nm + ' ' + (val >= 0 ? '+' : '') + val + '）';
}

// 品目を追加（同コードが既にあれば更新＝重複を作らない）
function prepForm_addItem(data) {
  if (!data.code || !data.name) throw new Error('食材名とコードは必須です');
  const m = ensurePrepMaster_();
  const lr = m.getLastRow();
  const parts = (data.parts || '').trim();
  const row = [
    data.code, data.name, data.rule || '',
    parts ? '' : (data.unitG === '' || data.unitG == null ? 1 : Number(data.unitG)),
    data.uname || '', Number(data.step) || 1,
    parts, parts ? (Number(data.perUnitG) || 1) : '',
    data.expiry || '', '', 0, true
  ];
  var maxOrder = 0, target = -1;
  if (lr >= 2) {
    const cur = m.getRange(2, 1, lr - 1, 11).getValues();
    for (var i = 0; i < cur.length; i++) {
      maxOrder = Math.max(maxOrder, Number(cur[i][10]) || 0);
      if (String(cur[i][0]) === String(data.code) && target < 0) target = i + 2;
    }
  }
  if (target > 0) {
    // 既存を更新するとき、フォームで空欄だった項目は「今の値を残す」（重さだけ変えても期限・在庫・並び順が消えない）
    const ex = m.getRange(target, 1, 1, 12).getValues()[0];
    if (row[8] === '' || row[8] == null) row[8] = ex[8];   // 期限：空なら既存を維持
    if (row[9] === '' || row[9] == null) row[9] = ex[9];   // 在庫：空なら既存を維持
    row[10] = Number(ex[10]) || (maxOrder + 10);           // 表示順：既存を維持
    m.getRange(target, 1, 1, 12).setValues([row]);
    const cur2 = m.getRange(2, 1, lr - 1, 1).getValues();
    for (var j = 0; j < cur2.length; j++) { if (String(cur2[j][0]) === String(data.code) && (j + 2) !== target) m.getRange(j + 2, 12).setValue(false); }
    buildPrepThawSheet_();
    return '「' + data.name + '」を更新しました';
  }
  row[10] = maxOrder + 10;
  m.appendRow(row);
  buildPrepThawSheet_();
  return '「' + data.name + '」を追加しました';
}

// 品目を上下に並べ替え（dir=-1 上へ / +1 下へ）。表示順を連番で振り直す。
function prepForm_moveItem(code, dir) {
  const m = ensurePrepMaster_();
  const lr = m.getLastRow();
  if (lr < 2) throw new Error('品目がありません');
  const data = m.getRange(2, 1, lr - 1, 12).getValues();
  const active = [];
  for (var i = 0; i < data.length; i++) {
    if (data[i][11] === true || data[i][11] === 'TRUE') active.push({ row: i + 2, code: String(data[i][0]), order: Number(data[i][10]) || 0 });
  }
  active.sort(function(a, b) { return a.order - b.order; });
  var idx = -1;
  for (var j = 0; j < active.length; j++) { if (active[j].code === String(code)) idx = j; }
  if (idx < 0) throw new Error('食材が見つかりません');
  var swap = idx + dir;
  if (swap < 0 || swap >= active.length) return '端です';
  var tmp = active[idx]; active[idx] = active[swap]; active[swap] = tmp;
  for (var k = 0; k < active.length; k++) { m.getRange(active[k].row, 11).setValue((k + 1) * 10); }
  buildPrepThawSheet_();
  return 'moved';
}

// 品目を表から外す（有効=FALSE。同コードの全行を外す＝重複も一括で消える）
function prepForm_removeItem(code) {
  const m = ensurePrepMaster_();
  const lr = m.getLastRow();
  if (lr < 2) throw new Error('品目がありません');
  const col = m.getRange(2, 1, lr - 1, 1).getValues();
  var nm = '', cnt = 0;
  for (var i = 0; i < col.length; i++) {
    if (String(col[i][0]) === String(code)) {
      m.getRange(i + 2, 12).setValue(false);
      if (!nm) nm = m.getRange(i + 2, 2).getValue();
      cnt++;
    }
  }
  if (cnt === 0) throw new Error('食材が見つかりません');
  buildPrepThawSheet_();
  return '「' + nm + '」を表から外しました（データは残ります）';
}

function PREP_FORM_HTML_() {
  return `<!DOCTYPE html><html><head><meta charset="utf-8">
<style>
body{font-family:Arial,sans-serif;font-size:13px;margin:12px}
label{display:block;margin-top:10px;font-weight:bold}
input,select{width:100%;padding:6px;margin-top:3px;box-sizing:border-box;border:1px solid #ccc;border-radius:3px}
.tabs{display:flex;gap:4px;margin-bottom:12px}
.tab{padding:7px 12px;cursor:pointer;border:1px solid #ccc;background:#f5f5f5;border-radius:3px;font-size:12px}
.tab.active{background:#4285f4;color:white;border-color:#4285f4}
.panel{display:none}.panel.active{display:block}
button{margin-top:16px;padding:9px 22px;background:#4285f4;color:white;border:none;cursor:pointer;border-radius:3px;font-size:13px}
button.sub{background:#c0392b}
.row{display:flex;gap:8px}.row>div{flex:1}
.note{color:#888;font-size:11px;margin-top:3px}
.cur{color:#333;font-size:12px;margin-top:6px;background:#f0f6ff;padding:6px;border-radius:3px}
#stat{min-height:18px;margin:4px 0 2px;font-weight:bold;font-size:12px}
</style></head><body>
<div class="tabs">
  <div class="tab active" onclick="sw('stock',this)">期限</div>
  <div class="tab" onclick="sw('adj',this)">数量の調整</div>
  <div class="tab" onclick="sw('item',this)">品目の追加・削除</div>
  <div class="tab" onclick="sw('sort',this)">並び順</div>
</div>
<div id="stat"></div>

<div id="tab-stock" class="panel active">
  <label>食材</label><select id="s-food" onchange="showCur()"></select>
  <div class="cur" id="s-cur"></div>
  <label>期限</label>
  <select id="s-exp">
    <option value="当日">当日</option>
    <option value="24時間">24時間</option>
    <option value="48時間">48時間</option>
    <option value="60時間">60時間</option>
    <option value="3日">3日</option>
  </select>
  <button onclick="saveExp()">保存する</button>
</div>

<div id="tab-adj" class="panel">
  <label>食材</label><select id="a-food"></select>
  <label>日付</label><select id="a-day"></select>
  <label>増減量（＋で増やす／−で減らす）</label>
  <input type="number" id="a-delta" step="1" placeholder="例: 2 や -3">
  <div class="note">その日の仕込み量に足し引きします。0を保存すると調整を取り消します。単位は表の単位（バット・袋・個・g等）</div>
  <label>メモ</label><input type="text" id="a-memo" placeholder="例: 予約20名">
  <button onclick="saveAdj()">保存する</button>
</div>

<div id="tab-item" class="panel">
  <b>■ 追加</b>
  <label>食材名 *（食材マスタから選ぶ）</label>
  <select id="i-food" onchange="onPickFood()"></select>
  <label>食材コード（自動）</label>
  <input type="text" id="i-code" readonly style="background:#eee">
  <label>表示名 *（表の「食材」列に出る名前）</label>
  <input type="text" id="i-name" placeholder="ふつうは食材名のまま。合算はクロ・バラ等に">
  <label>期限</label>
  <select id="i-exp">
    <option value="">（なし）</option>
    <option value="当日">当日</option>
    <option value="24時間">24時間</option>
    <option value="48時間">48時間</option>
    <option value="60時間">60時間</option>
    <option value="3日">3日</option>
  </select>
  <label>表示のしかた *</label>
  <select id="i-mode" onchange="modeChange()">
    <option value="gram">グラムで表示（◯g単位で切り上げ）</option>
    <option value="unit">個数で表示（缶・袋・個など。◯gで1つ）</option>
    <option value="combo">合算（クロ・バラのように複数食材を1つにまとめる）</option>
  </select>
  <div id="m-gram">
    <label>何g単位で切り上げ</label><input type="number" id="i-step" placeholder="例: 300">
    <div class="note">例：300 と入れると 300g単位で切り上げ（0→300→600…）。1なら切り上げなし。</div>
  </div>
  <div id="m-unit" style="display:none">
    <label>単位の名前</label><input type="text" id="i-uname" placeholder="缶・袋・個・枚">
    <label>1つ＝何g</label><input type="number" id="i-unitg" placeholder="例: 1810（1缶=1810g）">
    <div class="note">「1つが何gか」を入れます。例：コーンは1缶=1810g、ほうれん草は1個=50g。</div>
  </div>
  <div id="m-combo" style="display:none">
    <label>合算コード（＋でつなぐ）</label><input type="text" id="i-parts" placeholder="612087+612089+612047">
    <label>1つ＝何g</label><input type="number" id="i-perunit" placeholder="例: 1000（1バット=1000g）">
    <label>単位の名前</label><input type="text" id="i-uname2" placeholder="バット">
    <div class="note">複数の食材を合計して1つ（バット等）にまとめる特殊用。</div>
  </div>
  <button onclick="addItem()">追加する</button>
  <hr style="margin:18px 0">
  <b>■ 削除（表から外す）</b>
  <label>食材</label><select id="d-food"></select>
  <button class="sub" onclick="removeItem()">表から外す</button>
  <div class="note">データは残るので、あとで戻せます。</div>
</div>

<div id="tab-sort" class="panel">
  <div class="note">▲▼で、表に出る順番を入れ替えます。</div>
  <div id="sortlist"></div>
</div>

<script>
var DATA={items:[],days:[]};
function msg(t,err){var el=document.getElementById('stat');el.textContent=t;el.style.color=err?'#c0392b':'#188038';clearTimeout(window._mt);window._mt=setTimeout(function(){el.textContent='';},4000);}
function sw(t,el){document.querySelectorAll('.tab').forEach(function(x){x.classList.remove('active');});document.querySelectorAll('.panel').forEach(function(x){x.classList.remove('active');});document.getElementById('tab-'+t).classList.add('active');el.classList.add('active');}
function load(){google.script.run.withSuccessHandler(function(d){DATA=d;fill();}).withFailureHandler(function(e){msg('読込エラー: '+e,true);}).prepForm_getData();}
function fill(){
  var opt=DATA.items.map(function(it){return '<option value="'+it.code+'">'+it.name+'</option>';}).join('');
  ['s-food','a-food','d-food'].forEach(function(id){document.getElementById(id).innerHTML=opt;});
  document.getElementById('a-day').innerHTML=DATA.days.map(function(x){return '<option value="'+x.value+'">'+x.label+'</option>';}).join('');
  document.getElementById('i-food').innerHTML=DATA.foods.map(function(f){return '<option value="'+f.code+'" data-name="'+f.name+'">'+f.name+'</option>';}).join('');
  onPickFood();
  document.getElementById('sortlist').innerHTML=DATA.items.map(function(it,ix){return '<div style="display:flex;align-items:center;gap:6px;padding:5px 0;border-bottom:1px solid #eee"><span style="flex:1">'+(ix+1)+'. '+it.name+'</span><button style="margin:0;padding:4px 10px" onclick="mv('+it.code+',-1)">▲</button><button style="margin:0;padding:4px 10px" onclick="mv('+it.code+',1)">▼</button></div>';}).join('');
  showCur();
}
function mv(code,dir){google.script.run.withSuccessHandler(function(r){if(r==='端です'){msg('これ以上動かせません',true);return;}msg('並び替えました');load();}).withFailureHandler(function(e){msg('エラー: '+e,true);}).prepForm_moveItem(code,dir);}
function onPickFood(){var sel=document.getElementById('i-food');var o=sel.options[sel.selectedIndex];if(!o){return;}document.getElementById('i-code').value=sel.value;document.getElementById('i-name').value=o.getAttribute('data-name');}
function curItem(code){for(var i=0;i<DATA.items.length;i++){if(DATA.items[i].code===code)return DATA.items[i];}return null;}
function showCur(){var it=curItem(document.getElementById('s-food').value);var el=document.getElementById('s-cur');if(!it){el.textContent='';return;}el.textContent='今の期限: '+(it.expiry||'—');if(it.expiry){document.getElementById('s-exp').value=it.expiry;}}
function saveExp(){google.script.run.withSuccessHandler(function(m){msg(m);load();}).withFailureHandler(function(e){msg('エラー: '+e,true);}).prepForm_saveExpiry(document.getElementById('s-food').value,document.getElementById('s-exp').value);}
function saveAdj(){var v=document.getElementById('a-delta').value;if(v===''){msg('増減量を入れてください',true);return;}google.script.run.withSuccessHandler(function(m){msg(m);}).withFailureHandler(function(e){msg('エラー: '+e,true);}).prepForm_saveAdjust(document.getElementById('a-food').value,document.getElementById('a-day').value,v,document.getElementById('a-memo').value);}
function modeChange(){var mo=document.getElementById('i-mode').value;document.getElementById('m-gram').style.display=mo==='gram'?'':'none';document.getElementById('m-unit').style.display=mo==='unit'?'':'none';document.getElementById('m-combo').style.display=mo==='combo'?'':'none';}
function addItem(){
  var n=document.getElementById('i-name').value.trim();var c=document.getElementById('i-code').value.trim();
  if(!n||!c){msg('表示名と食材コードは必須です',true);return;}
  var mo=document.getElementById('i-mode').value;var d={name:n,code:c,expiry:document.getElementById('i-exp').value,parts:'',perUnitG:''};
  if(mo==='gram'){var st=document.getElementById('i-step').value||'1';d.uname='g';d.unitG='1';d.step=st;d.rule=(st==='1'?'グラムで解凍':st+'g単位で切り上げ');}
  else if(mo==='unit'){var un=document.getElementById('i-uname').value.trim();var ug=document.getElementById('i-unitg').value;if(!un||!ug){msg('単位の名前と「1つ＝何g」を入れてください',true);return;}d.uname=un;d.unitG=ug;d.step='1';d.rule='1'+un+'='+ug+'g';}
  else{var pt=document.getElementById('i-parts').value.trim();var pu=document.getElementById('i-perunit').value;if(!pt||!pu){msg('合算コードと「1つ＝何g」を入れてください',true);return;}d.parts=pt;d.perUnitG=pu;d.uname=document.getElementById('i-uname2').value.trim()||'個';d.step='1';d.rule='合算（1'+d.uname+'='+pu+'g）';}
  google.script.run.withSuccessHandler(function(m){msg(m);load();}).withFailureHandler(function(e){msg('エラー: '+e,true);}).prepForm_addItem(d);
}
function removeItem(){var c=document.getElementById('d-food').value;var it=curItem(c);if(!confirm((it?it.name:'')+' を表から外しますか？')){return;}google.script.run.withSuccessHandler(function(m){msg(m);load();}).withFailureHandler(function(e){msg('エラー: '+e,true);}).prepForm_removeItem(c);}
load();
</script></body></html>`;
}

// 【お試し】横型の発注ワークシートを別シートに作る。既存の発注ワークシートは触らない。
//  レイアウト：A=コード B=食材名／各曜日=基準・在庫の2列＋区切り空列。
//  行1=日付, 行2=曜日(2列結合), 行3=基準/在庫の小見出し。発注不可はグレー、行はしま模様、左2列固定。
function buildOrderSheetWide_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = '発注ワークシート_横型';
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  else { sh.clear(); sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart(); }

  // 火 C(3)D(4) [E空] / 水 F(6)G(7) [H空] / 金 I(9)J(10) [K空] / 日 L(12)M(13)
  const days = [{label:'火',dow:3},{label:'水',dow:4},{label:'金',dow:6},{label:'日',dow:1}];
  const parCols = [3, 6, 9, 12];   // 各曜日の基準列（在庫はその隣、区切りは更に隣）
  const gapCols = [5, 8, 11];      // 区切り空列
  const totalCols = 13;

  // 見出し：行1=日付(2列結合・数式で自動更新)、行2=曜日(2列結合)、行3=小見出し
  sh.getRange(3, 1).setValue('コード');
  sh.getRange(3, 2).setValue('食材名');
  days.forEach((dy, i) => {
    const c = parCols[i];
    // 日付は数式にして毎日の発注計算に自動追従させる（作り直さなくても最新になる）
    sh.getRange(1, c, 1, 2).merge();
    sh.getRange(1, c).setFormula("=IFERROR(TEXT(MINIFS('発注計算バッファ'!C:C,'発注計算バッファ'!D:D," + dy.dow + "),\"M/d\"),\"\")");
    sh.getRange(2, c, 1, 2).merge().setValue(dy.label + '曜発注');
    sh.getRange(3, c).setValue('基準');
    sh.getRange(3, c + 1).setValue('在庫');
  });

  // コード・食材名（消耗品を除き食材マスタ順で自動抽出）／各曜日の基準（1以上=整数, 1未満=0.1単位）
  sh.getRange(4, 1).setFormula("=ARRAYFORMULA(IFERROR(FILTER('食材マスタ'!A2:A,'食材マスタ'!C2:C<>\"消耗品\",'食材マスタ'!A2:A<>\"\"),\"\"))");
  sh.getRange(4, 2).setFormula("=ARRAYFORMULA(IFERROR(FILTER('食材マスタ'!B2:B,'食材マスタ'!C2:C<>\"消耗品\",'食材マスタ'!A2:A<>\"\"),\"\"))");
  days.forEach((dy, i) => {
    const c = parCols[i];
    const f = "=ARRAYFORMULA(IF(A4:A=\"\",\"\",IFERROR(LET(v,"
      + "VLOOKUP(A4:A&MINIFS('発注計算バッファ'!C:C,'発注計算バッファ'!D:D," + dy.dow + "),"
      + "{'発注計算バッファ'!A:A&'発注計算バッファ'!C:C,'発注計算バッファ'!O:O},2,FALSE)"
      + "/(VLOOKUP(A4:A,'食材マスタ'!A:G,7,FALSE)/VLOOKUP(A4:A,'食材マスタ'!A:F,6,FALSE)),"
      + "IF(v>=1,ROUND(v,0),ROUND(v,1))),\"\")))";
    sh.getRange(4, c).setFormula(f);
  });
  SpreadsheetApp.flush();

  // データ行数は食材マスタの「消耗品以外」の件数で確定する。
  //   ※基準列の ARRAYFORMULA(IF(A4:A="",...)) は空行にも "" を書き込むため getLastRow は1000を返す。
  //     それを信じると空行が大量に残り印刷が何十ページにもなる → 件数で正しい最終行を決める。
  const master = ss.getSheetByName('食材マスタ');
  const catMap = {}, origCatOfOld = {};
  let dataRows = 0;
  master.getRange(2, 1, master.getLastRow() - 1, 14).getValues().forEach(r => {
    if (r[0]) catMap[String(r[0])] = (r[2] || '').toString().trim();
    const oc = String(r[11] || '').trim();
    if (oc) origCatOfOld[String(Number(oc))] = (r[2] || '').toString().trim();
    if (r[0] !== '' && String(r[2]).trim() !== '消耗品') dataRows++;
  });
  const lastRow = 3 + dataRows;

  // 見出しの体裁
  sh.getRange(1, 1, 3, totalCols).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
  parCols.forEach(c => { sh.getRange(4, c, dataRows, 1).setNumberFormat('#,##0.#').setFontWeight('bold').setHorizontalAlignment('center'); });  // 基準は太字・中央寄せ
  const trueCat = code => { let c = catMap[code] || ''; if (c === '発注終了') { const o = origCatOfOld[code]; if (o && o !== '発注終了') c = o; } return c; };
  const GRAY = '#b7b7b7', WHITE = '#ffffff', ZEBRA = '#eaf1fb';
  const codes = sh.getRange(4, 1, dataRows, 1).getValues().map(r => String(r[0]));
  const grayFor = (code, dow) => {
    const isEnded = (catMap[code] === '発注終了');
    const cat = trueCat(code);
    if (isEnded || cat === '消耗品') return true;
    if (cat === 'ドリンク' && (dow === 6 || dow === 1)) return true;
    if (cat === '酒' && (dow === 3 || dow === 6 || dow === 1)) return true;
    return false;
  };
  // 左2列（コード・食材名）
  sh.getRange(4, 1, dataRows, 2).setBackgrounds(codes.map((code, ri) => {
    if (code === '') return [WHITE, WHITE];
    const b = (catMap[code] === '発注終了') ? GRAY : (ri % 2 === 1 ? ZEBRA : WHITE);
    return [b, b];
  }));
  // 各曜日ブロック
  days.forEach((dy, i) => {
    const c = parCols[i];
    sh.getRange(4, c, dataRows, 2).setBackgrounds(codes.map((code, ri) => {
      if (code === '') return [WHITE, WHITE];
      const b = grayFor(code, dy.dow) ? GRAY : (ri % 2 === 1 ? ZEBRA : WHITE);
      return [b, b];
    }));
  });

  // 罫線（区切り空列は引かない）
  sh.getRange(1, 1, lastRow, 2).setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  parCols.forEach(c => { sh.getRange(1, c, lastRow, 2).setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID); });

  // 列幅・固定
  sh.setColumnWidth(1, 55); sh.setColumnWidth(2, 260);   // 食材名は広く＝全部見える
  parCols.forEach(c => { sh.setColumnWidth(c, 40); sh.setColumnWidth(c + 1, 40); });
  gapCols.forEach(c => sh.setColumnWidth(c, 16));         // 区切り空列は細く
  sh.setFrozenRows(3); sh.setFrozenColumns(2);            // 見出し3行＋左2列を固定＝日曜まで目で追える

  // 余分な行・列を削除して、印刷範囲をコードのある最終行までにする
  const extraR = sh.getMaxRows() - lastRow;
  if (extraR > 0) sh.deleteRows(lastRow + 1, extraR);
  const extraC = sh.getMaxColumns() - totalCols;
  if (extraC > 0) sh.deleteColumns(totalCols + 1, extraC);

  try { ss.toast('横型（改良版）を作りました。シート「発注ワークシート_横型」を見てください', '完成', 8); } catch(e){}
  console.log('buildOrderSheetWide_: ' + dataRows + '品目で作成');
}
function buildOrderSheetWide() { buildOrderSheetWide_(); }

// 【毎朝自動】横型の発注書を最新に保つ。グレーアウト・行数・罫線だけ更新し、在庫(手入力)・基準(数式)・日付(数式)は触らない。
// これで食材を追加・切替しても翌朝の発注計算で自動連動し、「作り直す」は基本不要になる（縦型と同じ扱い）。
function refreshOrderSheetWide_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('発注ワークシート_横型');
  if (!sh) return;  // まだ作られていなければ何もしない
  const master = ss.getSheetByName('食材マスタ');
  if (!master || master.getLastRow() < 2) return;
  const catMap = {}, origCatOfOld = {};
  let dataRows = 0;
  master.getRange(2, 1, master.getLastRow() - 1, 14).getValues().forEach(r => {
    if (r[0]) catMap[String(r[0])] = (r[2] || '').toString().trim();
    const oc = String(r[11] || '').trim();
    if (oc) origCatOfOld[String(Number(oc))] = (r[2] || '').toString().trim();
    if (r[0] !== '' && String(r[2]).trim() !== '消耗品') dataRows++;
  });
  if (dataRows < 1) return;
  const lastRow = 3 + dataRows, totalCols = 13;
  const parCols = [3, 6, 9, 12];
  // 行数を食材件数に合わせる（増えた分は行を足し、余分な空行は消す＝印刷が伸びない）
  const maxR = sh.getMaxRows();
  if (maxR < lastRow) sh.insertRowsAfter(maxR, lastRow - maxR);
  else if (maxR > lastRow) sh.deleteRows(lastRow + 1, maxR - lastRow);
  SpreadsheetApp.flush();
  // グレーアウト＋しま模様を再適用（発注終了＝元区分に戻して判定）
  const days = [3, 4, 6, 1];  // 火水金日のdow
  const trueCat = code => { let c = catMap[code] || ''; if (c === '発注終了') { const o = origCatOfOld[code]; if (o && o !== '発注終了') c = o; } return c; };
  const GRAY = '#b7b7b7', WHITE = '#ffffff', ZEBRA = '#eaf1fb';
  const codes = sh.getRange(4, 1, dataRows, 1).getValues().map(r => String(r[0]));
  const grayFor = (code, dow) => {
    if (catMap[code] === '発注終了') return true;
    const cat = trueCat(code);
    if (cat === '消耗品') return true;
    if (cat === 'ドリンク' && (dow === 6 || dow === 1)) return true;
    if (cat === '酒' && (dow === 3 || dow === 6 || dow === 1)) return true;
    return false;
  };
  sh.getRange(4, 1, dataRows, 2).setBackgrounds(codes.map((code, ri) => {
    if (code === '') return [WHITE, WHITE];
    const b = (catMap[code] === '発注終了') ? GRAY : (ri % 2 === 1 ? ZEBRA : WHITE);
    return [b, b];
  }));
  days.forEach((dow, i) => {
    const c = parCols[i];
    sh.getRange(4, c, dataRows, 2).setBackgrounds(codes.map((code, ri) => {
      if (code === '') return [WHITE, WHITE];
      const b = grayFor(code, dow) ? GRAY : (ri % 2 === 1 ? ZEBRA : WHITE);
      return [b, b];
    }));
  });
  // 罫線を最終行まで引き直す（区切り空列は引かない）
  sh.getRange(1, 1, lastRow, 2).setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  parCols.forEach(c => { sh.getRange(1, c, lastRow, 2).setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID); });
}

// ─────────────────────────────────────────────────────────────────────
// デバッグ
// ─────────────────────────────────────────────────────────────────────

function debugMissingFoods() {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet=ss.getSheetByName('食材マスタ');
  const masterData=masterSheet.getRange(2,1,masterSheet.getLastRow()-1,3).getValues().filter(r=>r[0]&&r[2]!=='消耗品');
  const bufSheet=ss.getSheetByName('発注計算バッファ');
  const bufSet=new Set(bufSheet.getRange(2,1,bufSheet.getLastRow()-1,1).getValues().map(r=>String(r[0])));
  let missing='';
  masterData.forEach(r=>{ const code=String(r[0]); if(!bufSet.has(code)) missing+=code+' '+r[1]+' ['+r[2]+']\n'; });
  console.log('食材マスタ件数: '+masterData.length+'\n発注バッファ食材数: '+new Set([...bufSet].filter(c=>c)).size+'\n\n【欠落】\n'+(missing||'なし'));
}

// ─────────────────────────────────────────────────────────────────────
// 食材マスタ重複チェック
// ─────────────────────────────────────────────────────────────────────

function 食材マスタ重複チェック(e) {
  if (!e||!e.source) return;
  if (e.source.getActiveSheet().getName()!=='食材マスタ') return;
  const masterSheet=e.source.getSheetByName('食材マスタ');
  const lastRow=masterSheet.getLastRow(); if(lastRow<2) return;
  const masterData=masterSheet.getRange(2,1,lastRow-1,3).getValues().filter(r=>r[0]!==''&&r[2]!=='消耗品');
  const codeCount={};
  masterData.forEach(r=>{ const code=String(r[0]); codeCount[code]=(codeCount[code]||0)+1; });
  const dupCodes=Object.entries(codeCount).filter(([,cnt])=>cnt>1).map(([code,cnt])=>code+'（'+cnt+'件）');
  if (dupCodes.length>0) SpreadsheetApp.getUi().alert('【警告】食材マスタに重複コードがあります\n\n'+dupCodes.join('\n'));
}

// ─────────────────────────────────────────────────────────────────────
// イベント登録・CP設定の日付入力バリデーション
// ─────────────────────────────────────────────────────────────────────

function onEdit(e) {
  if (!e || !e.range) return;
  const sheetName = e.range.getSheet().getName();

  // 食材マスタ重複チェック（既存）
  if (sheetName === '食材マスタ') {
    食材マスタ重複チェック(e);
    return;
  }

  // イベント登録・CP設定の日付バリデーション
  if (sheetName === 'イベント登録' || sheetName === 'CP設定') {
    validateDateInput_(e, sheetName);
  }
}

function validateDateInput_(e, sheetName) {
  const col = e.range.getColumn();
  const row = e.range.getRow();
  if (row < 2) return; // ヘッダー行は除外

  // イベント登録：B列(2)・C列(3)が日付列
  // CP設定：B列(2)・C列(3)が日付列
  if (col !== 2 && col !== 3) return;

  const val = e.range.getValue();
  if (val === '' || val === null) return; // 空欄はOK

  // Date型ならOK
  if (val instanceof Date && !isNaN(val.getTime())) {
    // 年が明らかにおかしい場合は警告
    const year = val.getFullYear();
    const thisYear = new Date().getFullYear();
    if (year < thisYear - 1 || year > thisYear + 2) {
      SpreadsheetApp.getUi().alert(
        '⚠️ 日付の確認',
        `${row}行目の日付（${year}年）が想定範囲外です。\n正しい日付か確認してください。`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    return;
  }

  // 数値型チェック（float型で入力された場合）
  if (typeof val === 'number') {
    const str = String(Math.round(val));
    if (/^\d{8}$/.test(str)) {
      // 8桁数値（例:20260206）→ 自動的にDate型に変換して書き戻す
      const year  = Number(str.substring(0, 4));
      const month = Number(str.substring(4, 6)) - 1;
      const day   = Number(str.substring(6, 8));
      const fixed = new Date(year, month, day);
      if (!isNaN(fixed.getTime())) {
        e.range.setValue(fixed);
        console.log(`${sheetName} ${row}行${col}列: ${val} → Date型に自動変換`);
        return;
      }
    }
    // 8桁でない数値は警告
    SpreadsheetApp.getUi().alert(
      '⚠️ 日付の入力形式が正しくありません',
      `${row}行目に「${val}」と入力されましたが、日付として認識できません。\n\n正しい入力形式：\n・2026/05/12\n・2026-05-12\n・セルの書式を「日付」にして入力`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    e.range.setValue(''); // 不正値をクリア
    return;
  }

  // 文字列チェック
  if (typeof val === 'string') {
    const str = val.trim().replace(/\.0+$/, '');
    // 8桁数字文字列
    if (/^\d{8}$/.test(str)) {
      const year  = Number(str.substring(0, 4));
      const month = Number(str.substring(4, 6)) - 1;
      const day   = Number(str.substring(6, 8));
      const fixed = new Date(year, month, day);
      if (!isNaN(fixed.getTime())) {
        e.range.setValue(fixed);
        return;
      }
    }
    // 一般的な日付文字列を試みる
    const parsed = new Date(str);
    if (!isNaN(parsed.getTime())) {
      e.range.setValue(parsed);
      return;
    }
    // 解析できない場合は警告
    SpreadsheetApp.getUi().alert(
      '⚠️ 日付の入力形式が正しくありません',
      `${row}行目に「${val}」と入力されましたが、日付として認識できません。\n\n正しい入力形式：\n・2026/05/12\n・2026-05-12`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    e.range.setValue('');
  }
}
// ─────────────────────────────────────────────────────────────────────
// 食材コード自動切替
// ─────────────────────────────────────────────────────────────────────

function applyFoodCodeSwitches_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName('食材マスタ');
  if (!masterSheet || masterSheet.getLastRow() < 2) return;
  const today = new Date(); today.setHours(0, 0, 0, 0);

  // 食材マスタ A〜N（14列）を読む
  // idx: 0=A食材コード 1=B食材名 2=C区分 3=D払出単価 4=E仕入単位
  //      5=F仕入規格量 6=G仕入納品量 7=H歩留率 8=I安全係数
  //      11=L旧食材コード 12=M切替日 13=N移行完了日
  const mVals = masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, 14).getValues();

  // 対象＝L(旧コード)あり & N(移行完了日)空 & today>=M(切替日)
  const targets = [];
  mVals.forEach((row, i) => {
    const oldCode = String(row[11]).trim();
    if (!oldCode || oldCode === 'undefined') return;
    if (row[13]) return;                        // 既に切替済み（移行完了日あり）
    const sw = parseFlexibleDate_(row[12]);     // 切替日
    if (!sw || sw > today) return;              // 切替日がまだ来ていない
    targets.push({
      mRow: i + 2,
      newCode: Number(row[0]), oldCode: Number(oldCode),
      name: row[1], payout: row[3], unit: row[4],
      spec: row[5], deliv: row[6], yieldR: row[7]
    });
  });
  if (targets.length === 0) return 0;           // 対象なし＝一瞬で終了（毎朝呼んでも軽い）

  const pSheet = ss.getSheetByName('商品マスタ');
  if (!pSheet || pSheet.getLastRow() < 2) return;
  const rng   = pSheet.getRange(2, 1, pSheet.getLastRow() - 1, 15); // A〜O
  const pVals = rng.getValues();
  const pForm = rng.getFormulas();
  // idx: 4=E食材コード 5=F食材名 6=G払出単価 7=H仕入単位 8=I仕入規格量
  //      9=J仕入納品量 10=K歩留率 12=M単位使用量 13=N金額(原価)

  const reportLines = [];
  targets.forEach(t => {
    let count = 0, autoCount = 0;
    for (let r = 0; r < pVals.length; r++) {
      if (Number(pVals[r][4]) !== t.oldCode) continue;
      const sheetRow = r + 2;
      if (pForm[r][4]) { autoCount++; continue; }   // E列が数式なら自動更新に任せる
      pSheet.getRange(sheetRow, 5).setValue(t.newCode);   // E食材コードだけ書き換える
      // F〜K(食材名・払出単価・規格量など)とN(金額)は商品マスタ行3のARRAYFORMULAが自動計算する列。
      // スピル結果セルに値を書くとスピル全体が#REF!になるので、逆に空にして数式のスピルを効かせる
      // （＝新コードの値が自動で入る）。
      pSheet.getRange(sheetRow, 6, 1, 6).clearContent();  // F〜K
      pSheet.getRange(sheetRow, 14).clearContent();       // N金額
      count++;
    }
    // 旧コードの食材マスタ行を「発注終了」にして新規発注を止める
    for (let i = 0; i < mVals.length; i++) {
      if (Number(mVals[i][0]) === t.oldCode && String(mVals[i][2]).trim() !== '発注終了') {
        masterSheet.getRange(i + 2, 3).setValue('発注終了');
      }
    }
    // 食材マスタN列に移行完了日を記録＝以後スキップ
    masterSheet.getRange(t.mRow, 14).setValue(today);
    let line = t.oldCode + ' → ' + t.newCode + '（' + t.name + '）：商品マスタ ' + count + '件を更新';
    if (autoCount > 0) line += '／数式の' + autoCount + '件は自動更新';
    reportLines.push(line);
  });

  // 規格量が変わったので使用量・トレンドを自動再計算
  try { const cp = buildIngredientDailyUsage(); saveCpSubtractMap_(cp); } catch(e) { console.log('使用量再計算スキップ: ' + e.message); }
  try { buildTrendHistoryAverages_(); } catch(e) { console.log('トレンド再計算スキップ: ' + e.message); }

  const msg = '食材コードの自動切替を実行しました。\n' + reportLines.join('\n');
  console.log(msg);
  try { ss.toast(reportLines.join('\n'), '食材コード切替', 8); } catch(e){}
  try { MailApp.sendEmail(Session.getEffectiveUser().getEmail(), '【発注書】食材コードを自動切替しました', msg); } catch(e){}
  return targets.length;
}

// 過去の切替で商品マスタの自動計算列(F〜K・N)にベタ値が入りスピルが#REF!になったのを直す。
// 切替が完了した新コード(食材マスタでL旧コードありN移行完了日あり→A列が新コード)の商品マスタ行を
// 空にして、行3のARRAYFORMULAスピルを復活させる。
function repairProductMasterSpill_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const m = ss.getSheetByName('食材マスタ'), p = ss.getSheetByName('商品マスタ');
  if (!m || !p || m.getLastRow() < 2 || p.getLastRow() < 2) return;
  const switched = {};
  m.getRange(2, 1, m.getLastRow() - 1, 14).getValues().forEach(r => {
    if (String(r[11] || '').trim() && r[13]) switched[String(Number(r[0]))] = true;  // L旧ありN完了あり→A新コード
  });
  const eVals = p.getRange(2, 5, p.getLastRow() - 1, 1).getValues();  // E列(食材コード)
  let n = 0;
  eVals.forEach((r, i) => {
    if (switched[String(Number(r[0]))]) {
      const row = i + 2;
      p.getRange(row, 6, 1, 6).clearContent();  // F〜K
      p.getRange(row, 14).clearContent();        // N金額
      n++;
    }
  });
  try { ss.toast('商品マスタの自動計算(スピル)を復活しました（' + n + '行）。#REF!が消えます', '修正完了', 6); } catch(e){}
  console.log('repairProductMasterSpill_: ' + n + '行を空にしてスピル復活');
}
function repairProductMasterSpill() { repairProductMasterSpill_(); }

function applyFoodCodeSwitches() {
  const n = applyFoodCodeSwitches_();
  if (n > 0) {
    try { sortIngredientMaster_(); } catch(e){ console.log('並べ替えスキップ: ' + e.message); }  // 新食材を番号順の正しい位置へ
    // 切替があったので発注ワークシートを新しい規格で作り直す
    try { 発注バッファ_ステップ1(); } catch(e){ console.log('ステップ1スキップ: ' + e.message); }
    try { 発注バッファ_ステップ2(); } catch(e){ console.log('ステップ2スキップ: ' + e.message); }
    SpreadsheetApp.getActiveSpreadsheet().toast('食材コード切替＋発注ワークシートの作り直しが完了しました', '完了', 6);
    console.log('食材コード切替＋発注ワークシート再生成が完了しました');
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast('切替対象はありませんでした（切替日が来た未処理の食材なし）', '確認', 6);
    console.log('切替対象はありませんでした');
  }
}

// 【1回だけ実行】ブロッコリー切替(旧614044→新616661・切替日2026/8/4)を食材マスタに登録。
// 実行すると食材マスタ末尾に新616661行を追加。既に登録済みなら何もしない（冪等）。
function registerBroccoliSwitch_202608() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const m = ss.getSheetByName('食材マスタ');
  const last = m.getLastRow();
  const codes = m.getRange(2, 1, last - 1, 1).getValues().map(r => Number(r[0]));
  const switchDate = new Date(2026, 7, 3);  // 2026/8/3（日本時間）
  const idx = codes.indexOf(616661);
  if (idx >= 0) {
    // 既に616661あり → 切替日を8/3に直し、完了日をクリアして再切替できるようにする
    const rowNum = idx + 2;
    m.getRange(rowNum, 12).setValue(614044);      // L 旧食材コード
    m.getRange(rowNum, 13).setValue(switchDate);  // M 切替日 = 8/3
    m.getRange(rowNum, 14).setValue('');          // N 移行完了日クリア
    try { ss.toast('616661の切替日を2026/8/3に直しました。続けて「食材コードの切替を今すぐ実行」を押してください。', '切替日を更新', 8); } catch(e){}
    console.log('registerBroccoliSwitch_202608: 616661の切替日を8/3に更新');
    return;
  }
  // A〜N: 食材コード,食材名,区分,払出単価,仕入単位,仕入規格量,仕入納品量,歩留率,安全係数,連動元コピー,連動倍率,旧食材コード,切替日,移行完了日
  const row = [616661, 'ｴｸｱﾄﾞﾙ産ﾌﾞﾛｯｺﾘｰ500g', '食材', 290, 'ｇ', 500, 500, 1, 1.2, '', '', 614044, switchDate, ''];
  m.getRange(last + 1, 1, 1, row.length).setValues([row]);
  try { ss.toast('新616661を登録しました（切替日2026/8/3）。続けて「食材コードの切替を今すぐ実行」を押してください。', 'ブロッコリー切替登録', 8); } catch(e){}
  console.log('registerBroccoliSwitch_202608: 616661を食材マスタに追加しました');
}

// 食材マスタをA列(食材コード)昇順で並べ替える。人間が番号順に探せるように。
// 発注ワークシート・発注計算バッファは食材マスタの並び順で生成されるので、これで両方コード順になる。
function sortIngredientMaster_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('食材マスタ');
  if (!sheet || sheet.getLastRow() < 3) return 0;
  const n = sheet.getLastRow() - 1;
  const rng = sheet.getRange(2, 1, n, sheet.getLastColumn());
  const vals = rng.getValues();
  // 新食材のL列(旧食材コード)から「旧コード→その新食材の区分」を作り、発注終了品の“元の区分”を復元する。
  // 例：616661(食材)のL列に614044→614044は元「食材」。区分列は発注終了のまま残すが、並びは元の食材の位置に。
  const origCatOfOld = {};
  vals.forEach(r => {
    const oldCode = String(r[11] || '').trim();  // L列 旧食材コード
    if (oldCode) origCatOfOld[String(Number(oldCode))] = String(r[2] || '').trim();  // その新行のC列＝元区分
  });
  // 区分の並び順：酒→ドリンク→食材→消耗品（同じ区分の中は食材コード昇順）
  const CAT_ORDER = { '酒': 1, 'ドリンク': 2, '食材': 3, '消耗品': 4 };
  const catOf = r => {
    let cat = String(r[2] || '').trim();
    if (cat === '発注終了') {                       // 発注終了品は元の区分に戻して並べる（いつもの場所にグレーで残す）
      const oc = origCatOfOld[String(Number(r[0]))];
      if (oc && oc !== '発注終了') cat = oc;
    }
    return cat;
  };
  const rank = r => CAT_ORDER[catOf(r)] || 8;  // 未知の区分は末尾側
  vals.sort((a, b) => {
    const ra = rank(a), rb = rank(b);
    if (ra !== rb) return ra - rb;
    return (Number(a[0]) || 0) - (Number(b[0]) || 0);  // A列＝食材コード
  });
  rng.setValues(vals);
  return n;
}
function sortIngredientMaster() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const n = sortIngredientMaster_();
  ss.toast('食材マスタを並べ替えました。発注ワークシートも作り直します（2〜3分お待ちください）', '並べ替え中', 8);
  try { 発注バッファ_ステップ1(); } catch(e){ console.log('ステップ1スキップ: ' + e.message); }
  try { 発注バッファ_ステップ2(); } catch(e){ console.log('ステップ2スキップ: ' + e.message); }
  ss.toast('区分順（酒→ドリンク→食材→消耗品）＋コード順に並べ替えました。発注終了品は元の位置にグレーのまま。発注ワークシートも作り直しました（' + n + '品目）', '完了', 6);
}

function parseFlexibleDate_(val) {
  if (!val) return null;
  if (val instanceof Date) { const d=new Date(val); d.setHours(0,0,0,0); return d; }
  const num=Number(val);
  if (!isNaN(num)&&num>40000&&num<60000) { const d=new Date(Date.UTC(1899,11,30)+num*86400000); d.setHours(0,0,0,0); return d; }
  const str=String(val).trim();
  const m8=str.match(/^(\d{4})(\d{2})(\d{2})$/);
  if (m8) { const d=new Date(Number(m8[1]),Number(m8[2])-1,Number(m8[3])); d.setHours(0,0,0,0); return d; }
  const d=new Date(str);
  return isNaN(d.getTime()) ? null : (d.setHours(0,0,0,0), d);
}


// ═══════════════════════════════════════════════════════════════════════
// 【G】 異常値処理セクション
// ═══════════════════════════════════════════════════════════════════════

/**
 * 外れ値を検出してログシートを更新する。
 * ★祝日マスタに登録されている日はIQR判定から除外する。
 *   （祝日・連休の売上変動は正常のため異常値扱いしない）
 *   ただし曜日別平均への影響は getHolidayTaggedDates_ で別途制御する。
 */
function buildOutlierLog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const salesSheet = ss.getSheetByName('日次売上実績');
  if (!salesSheet||salesSheet.getLastRow()<2) { console.log('日次売上実績にデータがありません'); return; }

  const seen=new Set(), salesList=[];
  salesSheet.getRange(2,1,salesSheet.getLastRow()-1,5).getValues().forEach(row=>{
    if (!row[0]||!row[3]||row[3]<=0) return;
    const dateStr=toDateStr_(row[0]); if(seen.has(dateStr)) return; seen.add(dateStr);
    salesList.push({ dateStr, date:row[0] instanceof Date ? row[0] : new Date(row[0]), dowCode:Number(row[1]), dow:row[2], sales:Number(row[3]), customers:Number(row[4]) });
  });

  const dowGroups={'日':[],'月':[],'火':[],'水':[],'木':[],'金':[],'土':[]};
  salesList.forEach(s=>{ if(dowGroups[s.dow]) dowGroups[s.dow].push(s); });

  // ★祝日マスタの日付を取得（IQR判定から除外するため）
  const holidayTaggedDates = getHolidayTaggedDates_();

  // ★CP設定・イベント登録の期間もIQR判定から除外する
  //   （特売・CP期間は売上が通常より高く、外れ値として誤検知されるため）
  const cpEventTaggedDates = new Set();
  const cpSheet2 = ss.getSheetByName('CP設定');
  if (cpSheet2 && cpSheet2.getLastRow() > 1) {
    cpSheet2.getRange(2, 1, cpSheet2.getLastRow()-1, 8).getValues().forEach(row => {
      const note = (row[7] || '').toString().trim();
      if (note === '除外') return;
      const startRaw = row[1], endRaw = row[2];
      const start = startRaw instanceof Date ? startRaw : (startRaw ? new Date(String(startRaw).replace(/^(\d{4})(\d{2})(\d{2})$/, '$1/$2/$3')) : null);
      const end   = endRaw   instanceof Date ? endRaw   : (endRaw   ? new Date(String(endRaw  ).replace(/^(\d{4})(\d{2})(\d{2})$/, '$1/$2/$3')) : null);
      if (!start || !end || isNaN(start.getTime()) || isNaN(end.getTime())) return;
      start.setHours(0,0,0,0); end.setHours(0,0,0,0);
      const cur = new Date(start);
      while (cur <= end) { cpEventTaggedDates.add(toDateStr_(cur)); cur.setDate(cur.getDate() + 1); }
    });
  }
  const evSheet3 = ss.getSheetByName('イベント登録');
  if (evSheet3 && evSheet3.getLastRow() > 1) {
    evSheet3.getRange(2, 1, evSheet3.getLastRow()-1, 5).getValues().forEach(row => {
      if (!row[0] || (row[4] || '').toString().trim() === '除外') return;
      const start = row[1] instanceof Date ? row[1] : (row[1] ? new Date(String(row[1]).replace(/^(\d{4})(\d{2})(\d{2})$/, '$1/$2/$3')) : null);
      const end   = row[2] instanceof Date ? row[2] : (row[2] ? new Date(String(row[2]).replace(/^(\d{4})(\d{2})(\d{2})$/, '$1/$2/$3')) : null);
      if (!start || !end || isNaN(start.getTime())) return;
      start.setHours(0,0,0,0); end.setHours(0,0,0,0);
      const cur = new Date(start);
      while (cur <= end) { cpEventTaggedDates.add(toDateStr_(cur)); cur.setDate(cur.getDate() + 1); }
    });
  }

  const outliers=[], normalMap={};

  Object.entries(dowGroups).forEach(([dow, items]) => {
    if (items.length<OUTLIER_MIN_SAMPLES) { items.forEach(s=>{ normalMap[s.dateStr]=true; }); return; }

    // 祝日・CP/イベント期間を判定対象から分離
    const normalItems  = items.filter(s => !holidayTaggedDates.has(s.dateStr) && !cpEventTaggedDates.has(s.dateStr));
    const holidayItems = items.filter(s =>  holidayTaggedDates.has(s.dateStr) || cpEventTaggedDates.has(s.dateStr));

    // 祝日・CP期間は無条件でスキップ（異常値ログに載せない）
    holidayItems.forEach(s => { normalMap[s.dateStr]=true; });

    // 通常営業日だけでIQR判定
    if (normalItems.length<OUTLIER_MIN_SAMPLES) { normalItems.forEach(s=>{ normalMap[s.dateStr]=true; }); return; }

    const sorted=[...normalItems].sort((a,b)=>a.sales-b.sales);
    const n=sorted.length;
    const median=calcMedian_(sorted.map(s=>s.sales));
    const q1=calcMedian_(sorted.slice(0,Math.floor(n/2)).map(s=>s.sales));
    const q3=calcMedian_(sorted.slice(Math.ceil(n/2)).map(s=>s.sales));
    const iqr=q3-q1, lower=q1-OUTLIER_IQR_MULTIPLIER*iqr, upper=q3+OUTLIER_IQR_MULTIPLIER*iqr;

    normalItems.forEach(s=>{
      if (s.sales<lower||s.sales>upper) {
        const ratio=median>0?s.sales/median:1;
        outliers.push({ dateStr:s.dateStr, date:s.date, dow, sales:s.sales, median:Math.round(median), ratio:Math.round(ratio*100)/100, judgment:'外れ値', lower:Math.round(lower), upper:Math.round(upper) });
      } else { normalMap[s.dateStr]=true; }
    });
  });

  // 祝日マスタ・イベント登録から理由の自動候補を作成
  const holidayTagMap={};
  const holidaySheet=ss.getSheetByName('祝日マスタ');
  if (holidaySheet&&holidaySheet.getLastRow()>1) {
    holidaySheet.getRange(2,1,holidaySheet.getLastRow()-1,7).getValues().forEach(row=>{
      if (!row[0]||!row[6]) return;
      const tag=(row[6]||'').toString().trim(); if(tag==='土曜'||tag==='日曜') return;
      holidayTagMap[toDateStr_(row[0])]=tag;
    });
  }
  const eventNameMap={};
  const evSheet2=ss.getSheetByName('イベント登録');
  if (evSheet2&&evSheet2.getLastRow()>1) {
    evSheet2.getRange(2,1,evSheet2.getLastRow()-1,3).getValues().forEach(row=>{
      if (!row[0]||!row[1]||!row[2]) return;
      const name=(row[0]||'').toString().trim();
      const start=row[1] instanceof Date ? row[1] : new Date(String(row[1]).replace(/^(\d{4})(\d{2})(\d{2})$/,'$1/$2/$3'));
      const end=row[2] instanceof Date ? row[2] : new Date(String(row[2]).replace(/^(\d{4})(\d{2})(\d{2})$/,'$1/$2/$3'));
      if (!start||!end||isNaN(start.getTime())) return;
      const cur=new Date(start); cur.setHours(0,0,0,0);
      const endD=new Date(end); endD.setHours(0,0,0,0);
      while (cur<=endD) { const ds=toDateStr_(cur); if(!eventNameMap[ds]) eventNameMap[ds]=name; cur.setDate(cur.getDate()+1); }
    });
  }

  // 既存の理由欄を保持
  let existingReasons={};
  let sheet=ss.getSheetByName('異常値ログ');
  if (sheet&&sheet.getLastRow()>1) {
    sheet.getRange(2,1,sheet.getLastRow()-1,7).getValues().forEach(row=>{ if(row[0]&&row[6]) existingReasons[toDateStr_(row[0])]=row[6].toString().trim(); });
  }
  if (!sheet) sheet=ss.insertSheet('異常値ログ');
  sheet.clearContents();
  sheet.getRange(1,1,1,8).setValues([['日付','曜日','実績売上','曜日中央値','乖離率','判定','理由（手入力可）','処理状態']]);
  sheet.getRange(1,1,1,8).setBackground('#e8f4fd');
  sheet.setFrozenRows(1);

  if (outliers.length===0) { console.log('外れ値は検出されませんでした'); return; }

  outliers.sort((a,b)=>a.date-b.date);
  const dataRows=outliers.map(o=>{
    const manualReason=existingReasons[o.dateStr]||'';
    const autoReason=eventNameMap[o.dateStr]||holidayTagMap[o.dateStr]||'';
    const reason=manualReason||autoReason;
    const ratioStr=o.ratio>=1?'+'+Math.round((o.ratio-1)*100)+'%':'-'+Math.round((1-o.ratio)*100)+'%';
    return [o.date, o.dow, o.sales, o.median, ratioStr, o.judgment, reason, '曜日平均から除外'];
  });

  sheet.getRange(2,1,dataRows.length,8).setValues(dataRows);
  sheet.getRange(2,1,dataRows.length,1).setNumberFormat('yyyy/mm/dd');
  sheet.getRange(2,3,dataRows.length,2).setNumberFormat('#,##0');
  dataRows.forEach((row,i)=>{ sheet.getRange(i+2,1,1,8).setBackground('#ffe0e0'); });

  updateHolidayMasterFromOutliers_(ss, outliers);
  console.log('異常値ログを更新しました\n  外れ値: '+outliers.length+'件\n  正常値: '+(salesList.length-outliers.length)+'件');
}

function updateHolidayMasterFromOutliers_(ss, outliers) {
  const hmSheet=ss.getSheetByName('祝日マスタ'); if(!hmSheet||hmSheet.getLastRow()<2) return;
  const hmData=hmSheet.getRange(2,1,hmSheet.getLastRow()-1,9).getValues();
  const outlierMap={};
  outliers.forEach(o=>{ outlierMap[o.dateStr]=o.sales; });
  let updateCount=0;
  hmData.forEach((row,i)=>{
    if (!row[0]) return;
    const dateStr=toDateStr_(row[0]);
    if (outlierMap[dateStr]!==undefined&&!row[8]) { hmSheet.getRange(i+2,9).setValue(outlierMap[dateStr]); updateCount++; }
  });
  if (updateCount>0) console.log('祝日マスタI列を'+updateCount+'件更新しました');
}

function getOutlierExcludedDates_() {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const sheet=ss.getSheetByName('異常値ログ');
  const result=new Set();
  if (!sheet||sheet.getLastRow()<2) return result;
  sheet.getRange(2,1,sheet.getLastRow()-1,6).getValues().forEach(row=>{
    if (!row[0]) return;
    if ((row[5]||'').toString().trim()==='外れ値') result.add(toDateStr_(row[0]));
  });
  return result;
}

/**
 * 祝日マスタから「通常営業」以外の日付セットを返す
 * 曜日別平均の計算から除外するために使う（異常値ログとは独立）
 */
function getHolidayTaggedDates_(includeWeekendTags) {
  if (includeWeekendTags === undefined) includeWeekendTags = true;
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const sheet=ss.getSheetByName('祝日マスタ');
  const result=new Set();
  if (!sheet||sheet.getLastRow()<2) return result;
  sheet.getRange(2,1,sheet.getLastRow()-1,7).getValues().forEach(row=>{
    if (!row[0]||!row[6]) return;
    const tag=(row[6]||'').toString().trim();
    if (tag==='通常営業') return;
    // 通常の土日（タグ=土曜/日曜）は「真の祝日」ではない。曜日別平均の計算では除外せず実績として使う
    if (!includeWeekendTags && (tag==='土曜' || tag==='日曜')) return;
    result.add(toDateStr_(row[0]));
  });
  return result;
}


// ═══════════════════════════════════════════════════════════════════════
// 【H】 祝日マスタセクション
// ═══════════════════════════════════════════════════════════════════════

function updateSmartHolidayMaster() {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  let sheet=ss.getSheetByName('祝日マスタ'); if(!sheet) sheet=ss.insertSheet('祝日マスタ');
  sheet.clear();
  sheet.getRange(1,1,1,9).setValues([['日付','曜日コード','営業区分(isOff)','連休ID','連休日数','連休何日目','イベントタグ','備考','売上実績平均']]).setBackground('#444444').setFontColor('#ffffff').setFontWeight('bold');
  const tz='Asia/Tokyo', now=new Date(), year=now.getFullYear();
  const scanStart=new Date(year-1,11,20), outputStart=new Date(year,0,1), endTime=new Date(year+1,11,31);
  const calendar=CalendarApp.getCalendarById('ja.japanese#holiday@group.v.calendar.google.com');
  const holidayEvents=calendar.getEvents(scanStart,endTime);
  const LEGAL_HOLIDAYS=new Set(['元日','成人の日','建国記念の日','天皇誕生日','春分の日','昭和の日','憲法記念日','みどりの日','こどもの日','海の日','山の日','敬老の日','秋分の日','スポーツの日','文化の日','勤労感謝の日','振替休日','国民の休日']);
  const holidayMap={};
  holidayEvents.forEach(e=>{
    const title=e.getTitle();
    const isLegal = LEGAL_HOLIDAYS.has(title)
                 || title.includes('振替休日')
                 || title.includes('国民の休日');
    if (!isLegal) return;
    const key = Utilities.formatDate(e.getStartTime(), tz, 'yyyy-MM-dd');
    holidayMap[key] = title.includes('振替休日') ? '振替休日'
                    : title.includes('国民の休日') ? '国民の休日'
                    : title;
  });
  const getPeriodName=date=>{ const m=date.getMonth()+1, d=date.getDate(); if(m===12&&d>=29) return '年末'; if(m===1&&d<=3) return '年始'; if(m===8&&d>=13&&d<=15) return 'お盆'; if(m===5&&d>=3&&d<=5) return 'GW'; return null; };
  const isOff=date=>{ const key=Utilities.formatDate(date,tz,'yyyy-MM-dd'), dow=date.getDay(); return (dow===0||dow===6||!!holidayMap[key]||!!getPeriodName(date)); };
  const tempTagMap=buildTempTagMap_(scanStart,endTime,holidayMap,getPeriodName,isOff);
  const salesByDowTag={};
  const salesSheet=ss.getSheetByName('日次売上実績');
  if (salesSheet&&salesSheet.getLastRow()>1) {
    const seen=new Set();
    salesSheet.getRange(2,1,salesSheet.getLastRow()-1,4).getValues().filter(r=>r[0]&&r[3]>0).forEach(row=>{
      const dateStr=toDateStr_(row[0]); if(seen.has(dateStr)) return; seen.add(dateStr);
      const dow=Number(row[1]), sales=Number(row[3]), tag=tempTagMap[dateStr];
      if (!tag||tag==='通常営業') return;
      const key=`${dow}_${tag}`; if(!salesByDowTag[key]) salesByDowTag[key]=[]; salesByDowTag[key].push(sales);
      const chainMatch=tag.match(/^(\d+)連休_(\d+)日目$/);
      if (chainMatch) {
        const chainKey=`${dow}_${chainMatch[1]}連休`, chainDayKey=`${dow}_${chainMatch[1]}連休_${chainMatch[2]}日目`;
        if(!salesByDowTag[chainKey]) salesByDowTag[chainKey]=[];
        if(!salesByDowTag[chainDayKey]) salesByDowTag[chainDayKey]=[];
        salesByDowTag[chainKey].push(sales); salesByDowTag[chainDayKey].push(sales);
      }
      const isGoogleHoliday=!!holidayMap[dateStr.replace(/\//g,'-')];
      if (isGoogleHoliday) {
        if(!salesByDowTag['祝日_全般']) salesByDowTag['祝日_全般']=[];
        salesByDowTag['祝日_全般'].push(sales);
        const holidayKey=`${dow}_祝日`; if(!salesByDowTag[holidayKey]) salesByDowTag[holidayKey]=[]; salesByDowTag[holidayKey].push(sales);
      }
    });
  }
  const avgByDowTag={};
  Object.keys(salesByDowTag).forEach(key=>{ const arr=salesByDowTag[key]; avgByDowTag[key]=Math.round(arr.reduce((a,b)=>a+b,0)/arr.length); });
  const normalSalesByDow={1:[],2:[],3:[],4:[],5:[],6:[],7:[]};
  if (salesSheet&&salesSheet.getLastRow()>1) {
    const seen2=new Set();
    salesSheet.getRange(2,1,salesSheet.getLastRow()-1,4).getValues().filter(r=>r[0]&&r[3]>0).forEach(row=>{
      const dateStr=toDateStr_(row[0]); if(seen2.has(dateStr)) return; seen2.add(dateStr);
      const dow=Number(row[1]), sales=Number(row[3]), tag=tempTagMap[dateStr];
      if (tag==='通常営業'&&dow>=1&&dow<=7) normalSalesByDow[dow].push(sales);
    });
  }
  const normalAvgByDow={};
  [1,2,3,4,5,6,7].forEach(dow=>{ const arr=normalSalesByDow[dow]; normalAvgByDow[dow]=arr.length>0?Math.round(arr.reduce((a,b)=>a+b,0)/arr.length):0; });
  const DEFAULT_RATE={'GW':{1:1.1,2:1.1,3:1.1,4:1.1,5:1.1,6:1.1,7:1.1},'お盆':{1:1.05,2:1.05,3:1.05,4:1.05,5:1.05,6:1.05,7:1.05},'年末':{1:1.1,2:1.1,3:1.1,4:1.1,5:1.1,6:1.1,7:1.1},'年始':{1:1.05,2:1.05,3:1.05,4:1.05,5:1.05,6:1.05,7:1.05},'3連休_1日目':{1:1.05,2:1.3,3:1.3,4:1.3,5:1.3,6:1.1,7:1.05},'3連休_2日目':{1:1.1,2:1.4,3:1.4,4:1.4,5:1.4,6:1.1,7:1.1},'3連休_3日目':{1:1.05,2:1.2,3:1.2,4:1.2,5:1.2,6:1.05,7:1.05},'祝日':{1:1.05,2:1.3,3:1.3,4:1.3,5:1.3,6:1.1,7:1.05},'土曜':{1:1.0,2:1.0,3:1.0,4:1.0,5:1.0,6:1.0,7:1.0},'日曜':{1:1.0,2:1.0,3:1.0,4:1.0,5:1.0,6:1.0,7:1.0},'飛び石平日':{1:1.0,2:1.1,3:1.1,4:1.1,5:1.1,6:1.0,7:1.0}};
  const results=[];
  let chainIdCounter=0, d=new Date(scanStart);
  while (d<=endTime) {
    const current=new Date(d), yesterday=new Date(current); yesterday.setDate(current.getDate()-1);
    if (isOff(current)&&!isOff(yesterday)) {
      let tempChain=[], temp=new Date(current);
      while (isOff(temp)) { tempChain.push(new Date(temp)); temp.setDate(temp.getDate()+1); }
      const chainLength=tempChain.length, structured=chainLength>=3;
      const chainId=structured?'CHAIN_'+Utilities.formatString('%03d',++chainIdCounter):'';
      tempChain.forEach((target,i)=>{
        if (target>endTime||target<outputStart) return;
        const key=Utilities.formatDate(target,tz,'yyyy-MM-dd'), tDow=target.getDay(), tCode=(tDow===0)?1:tDow+1;
        let tag=getPeriodName(target); if(!tag) tag=holidayMap[key]; if(!tag&&structured) tag=chainLength+'連休_'+(i+1)+'日目'; if(!tag) tag=(tDow===0?'日曜':'土曜');
        const salesAvg=getSalesAvg_(tCode,tag,chainLength,structured,holidayMap[key],avgByDowTag,normalAvgByDow,DEFAULT_RATE,salesByDowTag);
        results.push([new Date(target),tCode,true,chainId,structured?chainLength:'',structured?i+1:'',tag,'',salesAvg]);
      });
      d.setDate(d.getDate()+chainLength); continue;
    } else {
      if (current>=outputStart) {
        const dow=current.getDay(), tCode=(dow===0)?1:dow+1;
        const key=Utilities.formatDate(current,tz,'yyyy-MM-dd');
        let tag=getPeriodName(current); if(!tag) tag=holidayMap[key];
        if (!tag) { const tomorrow=new Date(current); tomorrow.setDate(current.getDate()+1); tag=(isOff(yesterday)&&isOff(tomorrow))?'飛び石平日':'通常営業'; }
        results.push([new Date(current),tCode,isOff(current),'','','',tag,'','']);
      }
    }
    d.setDate(d.getDate()+1);
  }
  if (results.length>0) sheet.getRange(2,1,results.length,9).setValues(results);
  sheet.getRange('A:A').setNumberFormat('yyyy/MM/dd');
  sheet.getRange(2,9,results.length,1).setNumberFormat('#,##0');
  sheet.setFrozenRows(1); sheet.autoResizeColumns(1,9);
  const iColRange=sheet.getRange(2,9,results.length,1), iValues=iColRange.getValues(), tagValues=sheet.getRange(2,7,results.length,1).getValues();
  iColRange.setBackgrounds(iValues.map((r,i)=>{ if(!r[0]) return['#ffffff']; const dow=results[i][1], tag=tagValues[i][0]; const hasReal=!!avgByDowTag[`${dow}_${tag}`]||!!avgByDowTag[`${dow}_祝日`]; return[hasReal?'#e8f5e9':'#fff3e0']; }));
  console.log('祝日マスタを更新しました（'+results.length+'日分）');
}

function getSalesAvg_(dow,tag,chainLength,structured,isHoliday,avgByDowTag,normalAvgByDow,DEFAULT_RATE,salesByDowTag) {
  const key1=`${dow}_${tag}`;
  if (salesByDowTag[key1]&&salesByDowTag[key1].length>=2) return avgByDowTag[key1];
  if (structured&&chainLength>=3) { const dayNum=tag.match(/_(\d+)日目$/)?.[1]; if(dayNum){ const key2=`${dow}_${chainLength}連休_${dayNum}日目`; if(avgByDowTag[key2]) return avgByDowTag[key2]; } }
  if (isHoliday&&avgByDowTag['祝日_全般']) return avgByDowTag['祝日_全般'];
  if (structured&&chainLength>=3) { const key4=`${dow}_${chainLength}連休`; if(avgByDowTag[key4]) return avgByDowTag[key4]; }
  if (isHoliday) { const key5=`${dow}_祝日`; if(avgByDowTag[key5]) return avgByDowTag[key5]; }
  if (avgByDowTag[key1]) return avgByDowTag[key1];
  const baseAvg=normalAvgByDow[dow]||0; if(baseAvg===0) return '';
  let rate=1.0;
  if (DEFAULT_RATE[tag]?.[dow]) { rate=DEFAULT_RATE[tag][dow]; }
  else if (structured&&chainLength>=3) { const dayNum=tag.match(/_(\d+)日目$/)?.[1]||1; const chainTag=`${chainLength}連休_${dayNum}日目`; if(DEFAULT_RATE[chainTag]?.[dow]) rate=DEFAULT_RATE[chainTag][dow]; }
  else if (isHoliday&&DEFAULT_RATE['祝日']?.[dow]) { rate=DEFAULT_RATE['祝日'][dow]; }
  return Math.round(baseAvg*rate);
}

function buildTempTagMap_(startTime,endTime,holidayMap,getPeriodName,isOff) {
  const tz='Asia/Tokyo', tagMap={};
  let d=new Date(startTime);
  while (d<=endTime) {
    const current=new Date(d), yesterday=new Date(current); yesterday.setDate(current.getDate()-1);
    const dateStr=Utilities.formatDate(current,tz,'yyyy/MM/dd');
    if (isOff(current)&&!isOff(yesterday)) {
      let tempChain=[], temp=new Date(current);
      while (isOff(temp)) { tempChain.push(new Date(temp)); temp.setDate(temp.getDate()+1); }
      const chainLength=tempChain.length, structured=chainLength>=3;
      tempChain.forEach((target,i)=>{
        const key=Utilities.formatDate(target,tz,'yyyy-MM-dd'), dStr=Utilities.formatDate(target,tz,'yyyy/MM/dd');
        let tag=getPeriodName(target); if(!tag) tag=holidayMap[key]; if(!tag&&structured) tag=chainLength+'連休_'+(i+1)+'日目';
        const tDow=target.getDay(); if(!tag) tag=(tDow===0?'日曜':'土曜');
        tagMap[dStr]=tag;
      });
      d.setDate(d.getDate()+chainLength); continue;
    } else {
      const key=Utilities.formatDate(current,tz,'yyyy-MM-dd');
      let tag=getPeriodName(current); if(!tag) tag=holidayMap[key];
      if (!tag) { const tomorrow=new Date(current); tomorrow.setDate(current.getDate()+1); tag=(isOff(yesterday)&&isOff(tomorrow))?'飛び石平日':'通常営業'; }
      tagMap[dateStr]=tag;
    }
    d.setDate(d.getDate()+1);
  }
  return tagMap;
}


// ═══════════════════════════════════════════════════════════════════════
// 【I】 祝日倍率マスタセクション
// ═══════════════════════════════════════════════════════════════════════

function buildHolidayRateMaster() {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const ingSheet=ss.getSheetByName('食材マスタ'), holidaySheet=ss.getSheetByName('祝日マスタ');
  const eventSheet=ss.getSheetByName('イベント登録'), logSheet=ss.getSheetByName('食材日時実績');
  let outSheet=ss.getSheetByName('祝日倍率マスタ'); if(!outSheet) outSheet=ss.insertSheet('祝日倍率マスタ');
  const tz='Asia/Tokyo', fmt=d=>Utilities.formatDate(d instanceof Date?d:new Date(d),tz,'yyyy-MM-dd');
  const ingLastRow=ingSheet.getLastRow(); if(ingLastRow<2){ console.log('食材マスタにデータがありません'); return; }
  const ingredients=ingSheet.getRange(2,1,ingLastRow-1,2).getValues().filter(r=>r[0]).map(r=>({code:String(r[0]),name:r[1]}));
  const excludeDates=new Set();
  const toDateHRM=(val)=>{
    if (!val&&val!==0) return null;
    if (val instanceof Date){ const d=new Date(val); d.setHours(0,0,0,0); return d; }
    const str=String(val).trim().replace(/\s/g,'');
    if (/^\d{8}$/.test(str)){ const d=new Date(parseInt(str.substring(0,4),10),parseInt(str.substring(4,6),10)-1,parseInt(str.substring(6,8),10)); d.setHours(0,0,0,0); return d; }
    const d=new Date(str); d.setHours(0,0,0,0); return d;
  };
  if (eventSheet) {
    const evLastRow=eventSheet.getLastRow();
    if (evLastRow>=2) {
      eventSheet.getRange(2,1,evLastRow-1,9).getValues().forEach(r=>{
        if (String(r[4]).trim()!=='除外') return;
        const start=toDateHRM(r[1]), end=toDateHRM(r[2]);
        if (!start||!end||isNaN(start.getTime())||isNaN(end.getTime())) return;
        if (start.getFullYear()<1900||start>end) return;
        for (let d=new Date(start);d<=end;d.setDate(d.getDate()+1)) excludeDates.add(fmt(new Date(d)));
      });
    }
  }
  console.log('除外日数: '+excludeDates.size+' 日');
  const normalDates=new Set(), tagDateMap={};
  holidaySheet.getDataRange().getValues().slice(1).forEach(r=>{
    if (!r[0]||!r[6]) return;
    const key=fmt(r[0]), tag=String(r[6]);
    if (excludeDates.has(key)) return;
    if (tag==='通常営業'){ normalDates.add(key); } else { if(!tagDateMap[tag]) tagDateMap[tag]=new Set(); tagDateMap[tag].add(key); }
  });
  const eventTags=Object.keys(tagDateMap);
  if (eventTags.length===0){ console.log('祝日マスタに「通常営業」以外のタグが見つかりません。先に updateSmartHolidayMaster() を実行してください。'); return; }
  const usageMap={};
  logSheet.getDataRange().getValues().slice(1).forEach(r=>{
    const code=r[2], qty=Number(r[4]);
    if (!r[0]||!code||isNaN(qty)||qty<=0) return;
    const key=fmt(r[0]); if(excludeDates.has(key)) return;
    const mk=`${code}_${key}`; usageMap[mk]=(usageMap[mk]||0)+qty;
  });
  const calcAvg=(code,dateSet)=>{ let sum=0,cnt=0; dateSet.forEach(d=>{ const v=usageMap[`${code}_${d}`]; if(v!==undefined){sum+=v;cnt++;} }); return cnt>0?sum/cnt:null; };
  const toNum=v=>(v!==null?Math.round(v*100)/100:'');
  const rows=[['食材コード','食材名','イベントタグ','該当日数','イベント平均','通常平均','自動倍率','手動倍率','実績フラグ']];
  ingredients.forEach(ing=>{
    const normalAvg=calcAvg(ing.code,normalDates);
    eventTags.forEach(tag=>{
      const dates=tagDateMap[tag], cnt=dates.size, eventAvg=calcAvg(ing.code,dates);
      let autoRate=1, flag;
      if (normalAvg===null){ flag='通常実績なし'; }
      else if (eventAvg===null){ flag='イベント実績なし'; }
      else if (cnt<3){ flag=`サンプル不足(${cnt}日)`; }
      else if (normalAvg===0){ flag='通常平均ゼロ'; }
      else { autoRate=Math.round((eventAvg/normalAvg)*100)/100; flag='実績あり'; }
      rows.push([ing.code,ing.name,tag,cnt,toNum(eventAvg),toNum(normalAvg),autoRate,'',flag]);
    });
  });
  outSheet.clear();
  outSheet.getRange(1,1,rows.length,9).setValues(rows);
  outSheet.getRange(1,1,1,9).setBackground('#444444').setFontColor('#ffffff').setFontWeight('bold');
  if (rows.length>1) outSheet.getRange(2,4,rows.length-1,4).setNumberFormat('0.00');
  if (rows.length>1) {
    const flagValues=outSheet.getRange(2,9,rows.length-1,1).getValues();
    flagValues.forEach((r,i)=>{ const cell=outSheet.getRange(i+2,9), f=String(r[0]); if(f==='実績あり') cell.setBackground('#d9ead3'); else if(f.startsWith('サンプル不足')) cell.setBackground('#fff2cc'); else cell.setBackground('#f4cccc'); });
  }
  outSheet.setFrozenRows(1); outSheet.autoResizeColumns(1,9);
  console.log('完了：'+(rows.length-1)+' 行を生成しました\n除外日数: '+excludeDates.size+' 日');
}


// ═══════════════════════════════════════════════════════════════════════
// 【J】 共通ユーティリティ
// ═══════════════════════════════════════════════════════════════════════

function getDayOfWeek_(dateVal) {
  const d = dateVal instanceof Date ? dateVal : new Date(dateVal);
  return ['日','月','火','水','木','金','土'][d.getDay()];
}
function getDayOfWeek(dateStr) { return getDayOfWeek_(dateStr); }

function getDowName_(dow) { return ['日','月','火','水','木','金','土'][dow-1]||''; }

function toDateStr_(val) {
  if (!val) return '';
  const d = (val instanceof Date) ? val : new Date(val);
  return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy/MM/dd');
}

function toDate_(val) {
  if (!val) return null;
  if (val instanceof Date) { const d = new Date(val); d.setHours(0,0,0,0); return d; }
  // ★float型（20260206.0）対応：小数点以下を除去してから処理
  const str = String(val).trim().replace(/\.0+$/, '');
  if (/^\d{8}$/.test(str)) {
    const d = new Date(Number(str.substring(0,4)), Number(str.substring(4,6))-1, Number(str.substring(6,8)));
    d.setHours(0,0,0,0); return d;
  }
  const d = new Date(str);
  return isNaN(d.getTime()) ? null : (d.setHours(0,0,0,0), d);
}

function formatDate_(date) {
  const y=date.getFullYear(), m=String(date.getMonth()+1).padStart(2,'0'), d=String(date.getDate()).padStart(2,'0');
  return `${y}/${m}/${d}`;
}

function calcMedian_(sortedArr) {
  const n=sortedArr.length; if(n===0) return 0;
  const mid=Math.floor(n/2);
  return n%2!==0 ? sortedArr[mid] : (sortedArr[mid-1]+sortedArr[mid])/2;
}

function showMessage_(msg) {
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { console.log(msg); }
}

function clearIngredientLinkSettings() {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName('食材マスタ');
  if (!master || master.getLastRow() < 2) return;

  const lastRow = master.getLastRow();
  const headers = master.getRange(1, 1, 1, master.getLastColumn()).getValues()[0];

  const clearCols = [];
  headers.forEach((h, i) => {
    const name = (h || '').toString();
    if (name.includes('連動元') || name.includes('連動倍率') || name.includes('連動')) {
      clearCols.push(i + 1);
    }
  });

  if (clearCols.length === 0) {
    console.log('「連動元」「連動倍率」列が見つかりませんでした');
    return;
  }

  const data = master.getRange(2, 1, lastRow - 1, master.getLastColumn()).getValues();
  const TARGET_CODES = [612087, 612089];
  let count = 0;
  data.forEach((row, i) => {
    if (!TARGET_CODES.includes(Number(row[0]))) return;
    clearCols.forEach(col => {
      if (master.getRange(i + 2, col).getValue() !== '') {
        master.getRange(i + 2, col).clearContent();
        count++;
      }
    });
  });

  console.log('連動元コピー列クリア完了: ' + count + 'セル');
}

// ============================================================
// 予測精度分析（analyzeForecastAccuracy）
// 売上予測シートの実績確定行を集計し「予測精度分析」シートに出力
// ============================================================
function analyzeForecastAccuracy() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const fSheet = ss.getSheetByName('売上予測');
  if (!fSheet || fSheet.getLastRow() < 2) {
    console.log('売上予測シートにデータがありません');
    return;
  }

  const rows = fSheet.getRange(2, 1, fSheet.getLastRow() - 1, 12).getValues();

  // CP設定から全CP期間の日付セットを作成
  const cpDateSet = new Set();
  const cpSheet = ss.getSheetByName('CP設定');
  if (cpSheet && cpSheet.getLastRow() > 1) {
    cpSheet.getRange(2, 1, cpSheet.getLastRow() - 1, 8).getValues().forEach(row => {
      if (!row[0] || (row[7] || '').toString().trim() === '除外') return;
      const s = row[1] instanceof Date ? row[1] : new Date(row[1]);
      const e = row[2] instanceof Date ? row[2] : new Date(row[2]);
      if (!s || !e || isNaN(s.getTime())) return;
      s.setHours(0,0,0,0); e.setHours(0,0,0,0);
      const cur = new Date(s);
      while (cur <= e) {
        cpDateSet.add(Utilities.formatDate(cur, 'Asia/Tokyo', 'yyyy/MM/dd'));
        cur.setDate(cur.getDate() + 1);
      }
    });
  }
  const evSheet = ss.getSheetByName('イベント登録');
  if (evSheet && evSheet.getLastRow() > 1) {
    evSheet.getRange(2, 1, evSheet.getLastRow() - 1, 5).getValues().forEach(row => {
      if (!row[0] || (row[4] || '').toString().trim() === '除外') return;
      const s = row[1] instanceof Date ? row[1] : new Date(row[1]);
      const e = row[2] instanceof Date ? row[2] : new Date(row[2]);
      if (!s || !e || isNaN(s.getTime())) return;
      s.setHours(0,0,0,0); e.setHours(0,0,0,0);
      const cur = new Date(s);
      while (cur <= e) {
        cpDateSet.add(Utilities.formatDate(cur, 'Asia/Tokyo', 'yyyy/MM/dd'));
        cur.setDate(cur.getDate() + 1);
      }
    });
  }

  // 実績確定行のみ集計
  const DOW_NAMES = ['', '月', '火', '水', '木', '金', '土', '日'];
  const dowStats    = {};  // 曜日別
  const typeStats   = { CP: { sum: 0, absSum: 0, count: 0, over: 0, under: 0 },
                        通常: { sum: 0, absSum: 0, count: 0, over: 0, under: 0 } };
  const monthStats  = {};  // 月別
  const weekStats   = {};  // 週別（月曜始まり）：予測合計・実績合計

  rows.forEach(row => {
    const dateVal = row[0], status = (row[9] || '').toString();
    if (status !== '実績確定') return;
    const forecast   = Number(row[8]);
    const actual     = Number(row[10]);
    if (!forecast || !actual) return;

    const dateObj = dateVal instanceof Date ? dateVal : new Date(dateVal);
    const dateStr = Utilities.formatDate(dateObj, 'Asia/Tokyo', 'yyyy/MM/dd');
    const dow     = Number(row[1]);
    const month   = dateObj.getMonth() + 1;
    const err     = forecast - actual;
    const absErr  = Math.abs(err);
    const errRate = err / actual;
    const isCP    = cpDateSet.has(dateStr);
    const typeKey = isCP ? 'CP' : '通常';

    // 曜日別
    if (!dowStats[dow]) dowStats[dow] = { sum: 0, absSum: 0, count: 0, name: DOW_NAMES[dow] || dow };
    dowStats[dow].sum    += errRate;
    dowStats[dow].absSum += absErr / actual;
    dowStats[dow].count  += 1;

    // CP/通常別
    typeStats[typeKey].sum    += errRate;
    typeStats[typeKey].absSum += absErr / actual;
    typeStats[typeKey].count  += 1;
    if (err > 0) typeStats[typeKey].over++;
    else         typeStats[typeKey].under++;

    // 月別
    if (!monthStats[month]) monthStats[month] = { sum: 0, absSum: 0, count: 0 };
    monthStats[month].sum    += errRate;
    monthStats[month].absSum += absErr / actual;
    monthStats[month].count  += 1;

    // 週別（月曜始まり）：予測合計・実績合計を積み上げ
    const wkMon = new Date(dateObj);
    wkMon.setDate(wkMon.getDate() - ((dateObj.getDay() + 6) % 7)); // その週の月曜へ
    wkMon.setHours(0, 0, 0, 0);
    const wkKey = Utilities.formatDate(wkMon, 'Asia/Tokyo', 'yyyy/MM/dd');
    if (!weekStats[wkKey]) weekStats[wkKey] = { fSum: 0, aSum: 0, count: 0 };
    weekStats[wkKey].fSum += forecast;
    weekStats[wkKey].aSum += actual;
    weekStats[wkKey].count += 1;
  });

  // 出力シート準備
  let outSheet = ss.getSheetByName('予測精度分析');
  if (!outSheet) outSheet = ss.insertSheet('予測精度分析');
  outSheet.clearContents();
  outSheet.clearFormats();  // 前回の「#,##0」書式が残ると①②③の%表示(例"14.3%")が0に潰れるため書式もリセット

  const pct = v => Math.round(v * 1000) / 10 + '%';

  const output = [];
  const headerRows = [];
  const H = (title) => { headerRows.push(output.length + 1); output.push([title, '', '', '', '']); };

  H('【予測精度分析】');
  output.push(['生成日時', new Date().toLocaleString('ja-JP'), '', '', '']);
  output.push(['', '', '', '', '']);

  // ① CP/通常別
  H('① CP期間 vs 通常日（日別）');
  output.push(['区分', '件数', '平均誤差率(+過大/-過小)', '絶対誤差率(MAPE)', '過大/過小 件数']);
  ['CP', '通常'].forEach(k => {
    const s = typeStats[k];
    if (!s.count) { output.push([k, 0, '-', '-', '-']); return; }
    output.push([k, s.count, pct(s.sum / s.count), pct(s.absSum / s.count), '過大' + s.over + '件 / 過小' + s.under + '件']);
  });
  output.push(['', '', '', '', '']);

  // ② 曜日別
  H('② 曜日別（日別MAPE）');
  output.push(['曜日', '件数', '平均誤差率', '絶対誤差率(MAPE)', '']);
  [1,2,3,4,5,6,7].forEach(d => {
    const s = dowStats[d];
    if (!s || !s.count) return;
    output.push([s.name + '曜日', s.count, pct(s.sum / s.count), pct(s.absSum / s.count), '']);
  });
  output.push(['', '', '', '', '']);

  // ③ 月別
  H('③ 月別（日別MAPE）');
  output.push(['月', '件数', '平均誤差率', '絶対誤差率(MAPE)', '']);
  for (let m = 1; m <= 12; m++) {
    const s = monthStats[m];
    if (!s || !s.count) continue;
    output.push([m + '月', s.count, pct(s.sum / s.count), pct(s.absSum / s.count), '']);
  }
  output.push(['', '', '', '', '']);

  // ④ 週別 合計（発注・シフトの実務指標。日別の凹凸は週合計で打ち消し合い、誤差が小さくなる）
  H('④ 週別 合計（★実務で見る指標）');
  output.push(['週(月曜〜日曜)', '日数', '予測合計', '実績合計', '誤差率(+過大/-過小)']);
  const wkStartRow = output.length + 1;                    // 週別の金額行が始まる行（千区切りはここだけに適用）
  const recentWk = Object.keys(weekStats).sort().slice(-13);
  let wkAbsSum = 0, wkFull = 0;
  recentWk.forEach(wk => {
    const s = weekStats[wk];
    const er = s.aSum > 0 ? (s.fSum - s.aSum) / s.aSum : 0;
    output.push([wk + '〜', s.count + '日', Math.round(s.fSum), Math.round(s.aSum), pct(er)]);
    if (s.count >= 7) { wkAbsSum += Math.abs(er); wkFull++; }   // 7日そろった週だけ平均に使う
  });
  output.push(['', '', '', '', '']);
  output.push(['→ 週合計の平均誤差（7日そろった週のみ）', wkFull + '週', '', '', wkFull > 0 ? pct(wkAbsSum / wkFull) : '-']);

  outSheet.getRange(1, 1, output.length, 5).setValues(output);
  // 千区切りは④の予測合計・実績合計(C:D)の“金額行だけ”に適用する。
  // ※シート全体のC:Dに掛けると、①②③のパーセント文字列(例"14.3%")が0扱いに潰れるため範囲を限定。
  if (recentWk.length > 0) outSheet.getRange(wkStartRow, 3, recentWk.length, 2).setNumberFormat('#,##0');
  outSheet.autoResizeColumns(1, 5);

  // セクション見出しを色付け（行位置は動的に記録済み）
  headerRows.forEach(r => outSheet.getRange(r, 1, 1, 5).setBackground('#c9daf8').setFontWeight('bold'));

  ss.toast('「予測精度分析」シートを確認してください', '分析完了', 5);
  console.log('分析完了：「予測精度分析」シートを確認してください');
}


// ═══════════════════════════════════════════════════════════════════════
// 【構造書生成】buildStructureDoc  ─── Ver10 全シート構造書をDriveに書き出す
// ═══════════════════════════════════════════════════════════════════════
/**
 * 全シートの名前・gid・サイズ・先頭5行を
 * Driveの「ペッパーランチ発注書Ver10_構造書.txt」に書き出す。
 *
 * 実行後にログに表示されるファイルIDをClaudeに伝えると、
 * Claude側でそのファイルを読み込んでシートの中身を確認できる。
 *
 * 【実行タイミング】
 *   ・初回セットアップ時に手動実行
 *   ・シート構成を変更したあとに再実行
 *   ・「Claudeにデータを見せたい」ときにも再実行
 */
function buildStructureDoc() {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const ssId   = ss.getId();
  const sheets = ss.getSheets();
  const lines  = [];
  const now    = new Date().toLocaleString('ja-JP');

  lines.push('# ペッパーランチ発注書 Ver10 — 全シート構造書');
  lines.push('生成日時: ' + now);
  lines.push('スプレッドシートID: ' + ssId);
  lines.push('');

  sheets.forEach(function(sheet) {
    try {
      const name    = sheet.getName();
      const gid     = sheet.getSheetId();
      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();

      lines.push('## シート: ' + name);
      lines.push('gid: ' + gid);
      lines.push('行数: ' + lastRow + ' / 列数: ' + lastCol);

      // 先頭5行をサンプルとして記録
      if (lastRow > 0 && lastCol > 0) {
        const sampleRows = Math.min(lastRow, 5);
        try {
          const data = sheet.getRange(1, 1, sampleRows, lastCol).getValues();
          lines.push('（先頭' + sampleRows + '行サンプル）');
          data.forEach(function(row) {
            lines.push(row.join('\t'));
          });
        } catch (e2) {
          lines.push('  [サンプル取得エラー: ' + e2.message + ']');
        }
      }
      lines.push('');
    } catch (e) {
      lines.push('  [シート読み込みエラー: ' + e.message + ']');
      lines.push('');
    }
  });

  const content  = lines.join('\n');
  const fileName = 'ペッパーランチ発注書Ver10_構造書.txt';

  // 既存ファイルがあれば上書き、なければ新規作成
  let file;
  const files = DriveApp.getFilesByName(fileName);
  if (files.hasNext()) {
    file = files.next();
    file.setContent(content);
    console.log('[buildStructureDoc] 更新完了 fileId=' + file.getId());
  } else {
    file = DriveApp.createFile(fileName, content, MimeType.PLAIN_TEXT);
    console.log('[buildStructureDoc] 新規作成 fileId=' + file.getId());
  }

  ss.toast('Drive に ' + fileName + ' を保存しました\nfileId: ' + file.getId(), '構造書生成完了', 8);
  return file.getId();
}


// ═══════════════════════════════════════════════════════════════════════
// 【診断】diagnoseSalesForecast  ─── 水曜予測の異常値を診断する
// ═══════════════════════════════════════════════════════════════════════
/**
 * 売上予測が異常に低い（または高い）ときに実行する診断関数。
 * GASエディタのログ（表示 > ログ）に診断結果が出る。
 *
 * 確認内容：
 *   1. 日次集計シートの水曜実績（直近10件）
 *   2. 売上予測シートの水曜予測と実績（直近10件）
 *   3. biasAdj（誤差補正係数）の計算結果
 *   4. 売上予測シートに残っている異常に低い予測値（5万円未満）
 */
function diagnoseSalesForecast() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const today = new Date(); today.setHours(0, 0, 0, 0);

  console.log('=== diagnoseSalesForecast 開始 ===');
  console.log('診断日時: ' + new Date().toLocaleString('ja-JP'));

  // ① 日次売上実績シートの水曜実績（直近10件）
  //    列＝日付/曜日コード/曜日名/総売上/総客数（曜日コードは 日=1…土=7 なので水曜=4）
  const dailySheet = ss.getSheetByName('日次売上実績');
  if (dailySheet && dailySheet.getLastRow() > 1) {
    const rows = dailySheet.getRange(2, 1, dailySheet.getLastRow() - 1, 5).getValues();
    const wedRows = rows.filter(r => Number(r[1]) === 4).slice(-10);
    console.log('--- ① 日次売上実績：水曜実績 直近10件 ---');
    wedRows.forEach(r => {
      const d = r[0] instanceof Date ? r[0].toLocaleDateString('ja-JP') : r[0];
      console.log(d + ' 水 売上=' + r[3] + ' 客数=' + r[4]);
    });
  } else {
    console.log('日次売上実績シート：データなし');
  }

  // ② 売上予測シートの水曜行（直近10件）
  const fcSheet = ss.getSheetByName('売上予測');
  if (fcSheet && fcSheet.getLastRow() > 1) {
    const rows = fcSheet.getRange(2, 1, fcSheet.getLastRow() - 1, 12).getValues();
    // 曜日コード4=水曜
    const wedRows = rows.filter(r => r[1] === 4).slice(-10);
    console.log('--- ② 売上予測シート：水曜 直近10件 ---');
    console.log('日付\t曜日別平均(D)\tイベント率(F)\t天候(G)\tトレンド(H)\t予測(I)\tステータス\t実績(K)\t誤差(L)');
    wedRows.forEach(r => {
      const d = r[0] instanceof Date ? r[0].toLocaleDateString('ja-JP') : r[0];
      console.log([d, Math.round(r[3]), r[5], r[6], r[7], Math.round(r[8]), r[9], r[10], r[11]].join('\t'));
    });

    // 異常に低い予測値を検出（未来日で5万円未満）
    const anomalies = rows.filter(r => {
      if (!r[0] || r[9] === '実績確定') return false;
      const d = new Date(r[0]); d.setHours(0,0,0,0);
      return d >= today && r[8] > 0 && r[8] < 50000;
    });
    if (anomalies.length > 0) {
      console.log('--- ⚠️ 異常低予測（5万円未満の未来日） ---');
      anomalies.forEach(r => {
        const d = r[0] instanceof Date ? r[0].toLocaleDateString('ja-JP') : r[0];
        const dow = ['','日','月','火','水','木','金','土'][r[1]] || r[1];
        console.log(d + ' ' + dow + ' 予測=' + Math.round(r[8]) + ' 曜日別平均=' + Math.round(r[3]));
      });
    } else {
      console.log('--- 異常低予測なし（5万円未満の未来日はゼロ）---');
    }
  }

  // ③ biasAdjの元になるI列の値を確認（直近28日のI列）
  if (fcSheet && fcSheet.getLastRow() > 1) {
    const rows = fcSheet.getRange(2, 1, fcSheet.getLastRow() - 1, 12).getValues();
    console.log('--- ③ 直近28日のI列（予測売上）と実績 ---');
    const cutoff = new Date(today); cutoff.setDate(today.getDate() - 28);
    const recent = rows.filter(r => {
      if (!r[0]) return false;
      const d = new Date(r[0]); d.setHours(0,0,0,0);
      return d >= cutoff && d < today;
    }).sort((a,b) => new Date(a[0]) - new Date(b[0]));
    recent.forEach(r => {
      const d = r[0] instanceof Date ? r[0].toLocaleDateString('ja-JP') : r[0];
      const dow = ['','日','月','火','水','木','金','土'][r[1]] || r[1];
      const ratio = (r[10] > 0 && r[8] > 0) ? (r[10] / r[8]).toFixed(3) : '-';
      console.log(d + ' ' + dow + ' 予測(I)=' + Math.round(r[8]) + ' 実績(K)=' + r[10] + ' ratio=' + ratio);
    });
  }

  console.log('=== 診断完了 ===');
  SpreadsheetApp.getActiveSpreadsheet().toast('診断完了。GASエディタのログを確認してください。', '診断', 5);
}


// ─────────────────────────────────────────────────────────────────────
// 【診断2】debugWeightedDowAvg  ─── 曜日別加重平均の計算根拠を確認する
// ─────────────────────────────────────────────────────────────────────
/**
 * 「なぜ水曜の曜日別平均が60,516円になったのか」を診断する。
 * avgExcludedDates に含まれる水曜日（直近83日）と、
 * 使われた水曜日実績の一覧をログに出力する。
 */
function debugWeightedDowAvg() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const today = new Date(); today.setHours(0, 0, 0, 0);

  console.log('=== debugWeightedDowAvg 開始 ===');

  // salesMapを再現（日次売上実績シートから読む＝buildSalesForecastと同じ元データ）
  //    列＝日付/曜日コード/曜日名/総売上/総客数
  const salesMap = {};
  const dailySheet = ss.getSheetByName('日次売上実績');
  if (dailySheet && dailySheet.getLastRow() > 1) {
    dailySheet.getRange(2, 1, dailySheet.getLastRow() - 1, 5).getValues().forEach(row => {
      if (!row[0]) return;
      const ds = toDateStr_(row[0]);
      const dow = (row[0] instanceof Date ? row[0] : new Date(row[0])).getDay() + 1;
      salesMap[ds] = { sales: Number(row[3]), dow };
    });
  }

  // excludedDatesを再現（イベント登録で「除外」の日）
  const excludedDates = new Set();
  const evSheet = ss.getSheetByName('イベント登録');
  if (evSheet && evSheet.getLastRow() > 1) {
    evSheet.getRange(2, 1, evSheet.getLastRow() - 1, 6).getValues().forEach(row => {
      if (!row[0] || (row[4] || '').toString() !== '除外') return;
      if (!row[1] || !row[2]) return;
      const s = new Date(row[1]); const e = new Date(row[2]);
      const cur = new Date(s);
      while (cur <= e) { excludedDates.add(toDateStr_(cur)); cur.setDate(cur.getDate() + 1); }
    });
  }

  // holidayTaggedDates
  const holidayTaggedDates = getHolidayTaggedDates_();

  // cpEventDates
  const cpEventDates = new Set();
  const cpSheet = ss.getSheetByName('CP設定');
  if (cpSheet && cpSheet.getLastRow() > 1) {
    cpSheet.getRange(2, 1, cpSheet.getLastRow() - 1, 3).getValues().forEach(row => {
      if (!row[1] || !row[2]) return;
      const s = new Date(row[1]); const e = new Date(row[2]);
      if (isNaN(s.getTime()) || isNaN(e.getTime())) return;
      const cur = new Date(s);
      while (cur <= e) { cpEventDates.add(toDateStr_(cur)); cur.setDate(cur.getDate() + 1); }
    });
  }
  // イベント登録（除外以外のイベント）もcpEventDatesに含める
  if (evSheet && evSheet.getLastRow() > 1) {
    evSheet.getRange(2, 1, evSheet.getLastRow() - 1, 6).getValues().forEach(row => {
      if (!row[0] || !row[1] || !row[2]) return;
      const s = new Date(row[1]); const e = new Date(row[2]);
      if (isNaN(s.getTime()) || isNaN(e.getTime())) return;
      const cur = new Date(s);
      while (cur <= e) { cpEventDates.add(toDateStr_(cur)); cur.setDate(cur.getDate() + 1); }
    });
  }

  const avgExcludedDates = new Set([...excludedDates, ...holidayTaggedDates, ...cpEventDates]);

  // 水曜日（dow=4）で直近83日分を調べる
  console.log('--- 水曜日（直近83日）の除外状況 ---');
  console.log('日付\t実績売上\t除外理由');
  const DOW_TARGET = 4; // 水曜
  let usedCount = 0;
  let usedWeightedSum = 0, usedWeightedTotal = 0;
  const ALPHA = 0.35;
  let sampleCount = 0;

  for (let daysBack = 1; daysBack <= 83; daysBack++) {
    const target = new Date(today);
    target.setDate(today.getDate() - daysBack);
    target.setHours(0, 0, 0, 0);
    if (target.getDay() + 1 !== DOW_TARGET) continue;

    const ds = toDateStr_(target);
    const sales = salesMap[ds] ? salesMap[ds].sales : 0;

    let reason = '';
    if (excludedDates.has(ds)) reason += '除外日 ';
    if (holidayTaggedDates.has(ds)) reason += '祝日マスタ ';
    if (cpEventDates.has(ds)) reason += 'CP/イベント ';

    const isExcluded = avgExcludedDates.has(ds);
    const hasData = sales > 0;

    if (!isExcluded && hasData && sampleCount < 10) {
      const weight = ALPHA * Math.pow(1 - ALPHA, sampleCount);
      usedWeightedSum   += sales * weight;
      usedWeightedTotal += weight;
      sampleCount++;
      console.log(ds + '\t' + sales + '\t✅使用(サンプル' + sampleCount + '件目) weight=' + weight.toFixed(4));
    } else {
      const skipReason = isExcluded ? (reason || '除外') : (!hasData ? 'データなし' : '');
      console.log(ds + '\t' + sales + '\t❌スキップ: ' + skipReason);
    }
  }

  const calcResult = usedWeightedTotal > 0 ? Math.round(usedWeightedSum / usedWeightedTotal) : 0;
  console.log('--- 計算結果 ---');
  console.log('使用サンプル数: ' + sampleCount + ' 件');
  console.log('計算された曜日別平均（水曜）: ' + calcResult + ' 円');
  console.log('売上予測シートのD列と一致しているか確認してください（シートの値: 60,516円）');

  console.log('=== 終了 ===');
  ss.toast('debugWeightedDowAvg完了。ログを確認してください。', '診断2', 5);
}


// ═══════════════════════════════════════════════════════════════════════
// 【ウェブアプリ】Claude がシートデータ・数式を読み取るための API
//   デプロイ設定: ウェブアプリ → アクセスできるユーザー「全員（匿名を含む）」
//   パラメータ:
//     sheet    = シート名（デフォルト: 発注計算バッファ）
//     mode     = 'values'（値のみ）| 'formulas'（数式のみ）| 'both'（デフォルト: 両方）
//     rows     = 取得行数 最大200（デフォルト: 50）
//     startRow = 開始行番号（デフォルト: 1）
// ═══════════════════════════════════════════════════════════════════════

function doGet(e) {
  try {
    const p = (e && e.parameter) ? e.parameter : {};
    const sheetName = p.sheet    || '発注計算バッファ';
    const mode      = p.mode     || 'both';
    const maxRows   = Math.min(parseInt(p.rows     || '50',  10), 200);
    const startRow  = Math.max(parseInt(p.startRow || '1',   10), 1);

    // このスクリプトが紐づいているスプレッドシート自身を読む。
    // IDを直書きするとブックをコピーしたとき（Ver10→Ver11など）コピー元を見続けてしまうため、
    // getActiveSpreadsheet() を使ってコピー先でもそのまま動くようにしている。
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return ok_({ error: 'スプレッドシートに紐づいていません' });
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      return ok_({ error: 'シートが見つかりません: ' + sheetName });
    }

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    // どのブックを読んでいるかを必ず返す（Ver10とVer11を取り違えないため）
    const book = { book: ss.getName(), bookId: ss.getId() };

    if (lastRow < startRow || lastRow === 0) {
      return ok_(Object.assign({}, book, { sheetName, totalRows: lastRow, totalCols: lastCol, startRow, returnedRows: 0, data: [] }));
    }

    const numRows = Math.min(maxRows, lastRow - startRow + 1);
    const range   = sheet.getRange(startRow, 1, numRows, lastCol);

    let data;
    if (mode === 'values') {
      data = range.getValues();
    } else if (mode === 'formulas') {
      data = range.getFormulas();
    } else {
      const values   = range.getValues();
      const formulas = range.getFormulas();
      data = values.map((row, r) =>
        row.map((val, c) => {
          const f = formulas[r][c];
          return f ? { v: val, f } : val;
        })
      );
    }

    return ok_(Object.assign({}, book, { sheetName, totalRows: lastRow, totalCols: lastCol, startRow, returnedRows: numRows, data }));

  } catch (err) {
    return ok_({ error: err.toString() });
  }
}

function ok_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─────────────────────────────────────────────────────────────────────
// イベント入力フォーム
// ─────────────────────────────────────────────────────────────────────

function showEventForm_() {
  const html = HtmlService.createHtmlOutput(EVENT_FORM_HTML_())
    .setWidth(520).setHeight(560);
  SpreadsheetApp.getUi().showModalDialog(html, 'イベント追加');
}

function EVENT_FORM_HTML_() {
  return `<!DOCTYPE html><html><head><meta charset="utf-8">
<style>
body{font-family:Arial,sans-serif;font-size:13px;margin:12px}
label{display:block;margin-top:10px;font-weight:bold}
input,select{width:100%;padding:6px;margin-top:3px;box-sizing:border-box;border:1px solid #ccc;border-radius:3px}
.tabs{display:flex;gap:4px;margin-bottom:12px}
.tab{padding:7px 14px;cursor:pointer;border:1px solid #ccc;background:#f5f5f5;border-radius:3px}
.tab.active{background:#4285f4;color:white;border-color:#4285f4}
.panel{display:none}.panel.active{display:block}
button{margin-top:16px;padding:9px 22px;background:#4285f4;color:white;border:none;cursor:pointer;border-radius:3px;font-size:13px}
.note{color:#888;font-size:11px;margin-top:3px}
</style></head><body>
<div class="tabs">
  <div class="tab active" onclick="switchTab('direct',this)">日付直接入力</div>
  <div class="tab" onclick="switchTab('rule',this)">繰り返しルール登録</div>
</div>

<div id="tab-direct" class="panel active">
  <label>イベント名 *</label>
  <input type="text" id="d-name" placeholder="例: 卒業式、体育祭">
  <label>開始日 *</label><input type="date" id="d-start">
  <label>終了日 *</label><input type="date" id="d-end">
  <label>対象区分</label>
  <select id="d-type">
    <option value="客数">客数</option>
    <option value="全品">全品</option>
    <option value="除外">除外（集計から除く）</option>
  </select>
  <label>補正倍率</label>
  <input type="number" id="d-rate" step="0.1" min="0.1" value="1" placeholder="空欄＝実績から自動計算">
  <div class="note">1.0が基準。▲で増（1.1,1.2…）、▼で減（0.9,0.8…）。他地域のイベントで客が流れて減るなら0.9など。自動計算にしたいときは欄を空にする</div>
  <label>備考</label><input type="text" id="d-note">
  <button onclick="submitDirect()">登録する</button>
</div>

<div id="tab-rule" class="panel">
  <label>イベント名 * <span style="font-weight:normal;color:#555">（毎年同じ名前にすること）</span></label>
  <input type="text" id="r-name" placeholder="例: 卒業式">
  <label>繰り返し種別 *</label>
  <select id="r-kind" onchange="updateRuleFields()">
    <option value="第N週曜日">第N週◯曜日（例: 3月第3土曜）</option>
    <option value="固定日">毎年同じ日付（例: 毎年7月20日）</option>
    <option value="暫定日">毎年ほぼ同じ時期（祭り等・日がずれる／後で修正可）</option>
  </select>
  <label>月 *</label>
  <select id="r-month">
    ${[1,2,3,4,5,6,7,8,9,10,11,12].map(m=>`<option value="${m}">${m}月</option>`).join('')}
  </select>
  <div id="r-week-row">
    <label>第何週目</label>
    <select id="r-week">
      ${[1,2,3,4,5].map(n=>`<option value="${n}">第${n}週</option>`).join('')}
    </select>
    <label>曜日</label>
    <select id="r-dow">
      ${['日','月','火','水','木','金','土'].map(d=>`<option value="${d}">${d}曜日</option>`).join('')}
    </select>
  </div>
  <div id="r-fixday-row" style="display:none">
    <label>何日 *</label>
    <input type="number" id="r-fixday" min="1" max="31" placeholder="例: 20">
    <div class="note" id="r-fixday-note"></div>
  </div>
  <label>期間（日数）</label>
  <input type="number" id="r-duration" value="1" min="1">
  <div class="note">1日イベントなら1。夏休み7/20〜8/31なら43</div>
  <label>補正倍率</label>
  <input type="number" id="r-rate" step="0.1" min="0.1" value="1" placeholder="空欄＝実績から自動計算">
  <div class="note">1.0が基準。▲で増（祭り等は1.1〜1.2）、▼で減（他地域イベントで客が流れて減るなら0.9等）。自動計算にしたいときは欄を空にする</div>
  <label>対象区分</label>
  <select id="r-type">
    <option value="客数">客数</option>
    <option value="全品">全品</option>
    <option value="除外">除外（集計から除く）</option>
  </select>
  <label>備考</label><input type="text" id="r-note">
  <button onclick="submitRule()">登録する</button>
</div>

<script>
function switchTab(tab,el){
  document.querySelectorAll('.tab').forEach(t=>t.classList.remove('active'));
  document.querySelectorAll('.panel').forEach(p=>p.classList.remove('active'));
  document.getElementById('tab-'+tab).classList.add('active');
  el.classList.add('active');
}
function updateRuleFields(){
  const k=document.getElementById('r-kind').value;
  document.getElementById('r-week-row').style.display=k==='第N週曜日'?'':'none';
  document.getElementById('r-fixday-row').style.display=(k==='固定日'||k==='暫定日')?'':'none';
  const note=document.getElementById('r-fixday-note');
  if(k==='暫定日'){note.textContent='去年と同じ日付（分からなければおおよその日）を入れてください。毎年その日付で仮登録されます。正確な日が分かったら、イベント登録シートの「繰り返し」列の日付を直接直せばOKです';}
  else{note.textContent='';}
}
function submitDirect(){
  const name=document.getElementById('d-name').value.trim();
  const start=document.getElementById('d-start').value;
  const end=document.getElementById('d-end').value;
  if(!name||!start||!end){alert('イベント名・開始日・終了日は必須です');return;}
  google.script.run
    .withSuccessHandler(msg=>{alert(msg);google.script.host.close();})
    .withFailureHandler(err=>alert('エラー: '+err))
    .registerDirectEvent({name,start:start.replace(/-/g,'/'),end:end.replace(/-/g,'/'),
      type:document.getElementById('d-type').value,
      rate:document.getElementById('d-rate').value,
      note:document.getElementById('d-note').value});
}
function submitRule(){
  const name=document.getElementById('r-name').value.trim();
  if(!name){alert('イベント名は必須です');return;}
  const kind=document.getElementById('r-kind').value;
  if((kind==='固定日'||kind==='暫定日')&&!document.getElementById('r-fixday').value){alert('「何日」を入力してください（暫定日は去年と同じ日でOK）');return;}
  google.script.run
    .withSuccessHandler(msg=>{alert(msg);google.script.host.close();})
    .withFailureHandler(err=>alert('エラー: '+err))
    .registerRuleEvent({name,
      kind,
      month:document.getElementById('r-month').value,
      week:document.getElementById('r-week').value,
      dow:document.getElementById('r-dow').value,
      fixday:document.getElementById('r-fixday').value||'',
      duration:document.getElementById('r-duration').value||'1',
      rate:document.getElementById('r-rate').value,
      type:document.getElementById('r-type').value,
      note:document.getElementById('r-note').value});
}
</script></body></html>`;
}

function registerDirectEvent(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const evSheet = ss.getSheetByName('イベント登録');
  if (!evSheet) throw new Error('イベント登録シートが見つかりません');
  const rate = data.rate ? Number(data.rate) : '';
  const startFmt = data.start.replace(/\//g, '');
  const endFmt   = data.end.replace(/\//g, '');
  evSheet.appendRow([data.name, startFmt, endFmt, '', data.type, '', rate, 10, '手動', data.note]);
  return '「' + data.name + '」を登録しました。\n\nbuildSalesForecast を実行すると予測に反映されます。';
}

// イベント登録シートに「繰り返し」列(K列)の見出しが無ければ用意する
function ensureEventRepeatColumn_(evSheet) {
  const head = evSheet.getLastColumn() >= 11 ? (evSheet.getRange(1, 11).getValue() || '').toString().trim() : '';
  if (head === '') evSheet.getRange(1, 11).setValue('繰り返し');
}

function registerRuleEvent(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const evSheet = ss.getSheetByName('イベント登録');
  if (!evSheet) throw new Error('イベント登録シートが見つかりません');

  // フォーム入力 → 「繰り返し」列のパターン文字列に変換
  const month = Number(data.month);
  let pattern;
  if (data.kind === '第N週曜日') {
    pattern = '毎年' + month + '月第' + Number(data.week) + (data.dow || '');
  } else { // 固定日 / 暫定日
    pattern = '毎年' + month + '月' + Number(data.fixday) + '日';
  }
  const duration = Number(data.duration) || 1;
  if (duration > 1) pattern += 'から' + duration + '日間';
  if (data.kind === '暫定日') pattern += ' 暫定';

  ensureEventRepeatColumn_(evSheet);
  const rate = data.rate ? Number(data.rate) : '';
  // 日付は空欄。K列(繰り返し)にパターンを入れると予測時に今年・来年へ自動展開される
  evSheet.appendRow([data.name, '', '', '', data.type, '', rate, 10, '', data.note, pattern]);
  return '「' + data.name + '」を毎年くり返しイベントとして登録しました。\n（繰り返し設定：' + pattern + '）\n\nbuildSalesForecast を実行すると予測に反映されます。';
}

// （旧「繰り返しルール→今年分に生成」機能 generateAnnualEvents_ は廃止。
//   繰り返しイベントはイベント登録シートの「繰り返し」列にパターンを書くだけで
//   予測時に自動展開されるため、手動生成は不要になりました）

// ─────────────────────────────────────────────────────────────────────
// キャンペーン商品（CP）入力フォーム
//   「CP設定」シートに直接入力していたものをフォーム化。
//   新商品CP … 実績が無いので「予測出食率(客数の何%が頼むか)」で発注量を見積もる（I列=空）。
//   通常商品CP … 既存商品がCPで伸びる。「補正倍率」で増やす（I列=通常）。空欄なら実績から自動学習。
// ─────────────────────────────────────────────────────────────────────
function showCpForm_() {
  const html = HtmlService.createHtmlOutput(CP_FORM_HTML_())
    .setWidth(540).setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'キャンペーン商品の登録');
}

// フォーム初期データ：商品マスタから 商品コード×商品名 の重複なし一覧を返す
function getCpFormData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('商品マスタ');
  const products = [];
  if (sh && sh.getLastRow() > 1) {
    const vals = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
    const seen = {};
    vals.forEach(r => {
      const code = Number(r[0]);
      const name = (r[1] || '').toString().trim();
      if (!code || !name || seen[code]) return;
      seen[code] = true;
      products.push({ code: code, name: name });
    });
    products.sort((a, b) => a.code - b.code);
  }
  return { products: products };
}

// フォーム登録：CP設定シートに1行追加する
function registerCpProduct(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('CP設定');
  if (!sh) throw new Error('CP設定シートが見つかりません');
  const name = (data.name || '').toString().trim();
  if (!name) throw new Error('CP名は必須です');
  const codes = (data.codes || '').toString().trim();
  if (!codes) throw new Error('対象商品を1つ以上選んでください');
  const start = (data.start || '').replace(/[-\/]/g, '');   // yyyymmdd
  const end   = (data.end   || '').replace(/[-\/]/g, '');
  if (!start || !end) throw new Error('開始日・終了日は必須です');
  // 列: A CP名 / B 開始 / C 終了 / D 商品コード / E 予測出食率 / F 補正倍率 / G 繰り返し / H 備考 / I 通常商品CP
  let servingRate = '', hosenRate = '', normalFlag = '';
  if (data.kind === '休止') {
    // 休止（一定期間の販売停止）：I列='休止'。出食率・倍率は不要。期間中は使う食材を発注・仕込みから差し引く。
    normalFlag = '休止';
  } else if (data.kind === '通常') {
    normalFlag = '通常';
    hosenRate = data.rate ? Number(data.rate) : '';   // 空欄＝実績から自動学習
  } else {
    // 新商品CP：予測出食率（％入力→小数）。空欄なら空（読込側が0.10でフォールバック）
    servingRate = data.rate ? (Number(data.rate) / 100) : '';
  }
  const note = (data.note || '').toString().trim();
  sh.appendRow([name, start, end, '', servingRate, hosenRate, '', note, normalFlag]);
  // ★商品コードD列は必ず文字列で入れる（"4986,4987"を桁区切り数値49864987と誤解される事故を防止）
  const newRow = sh.getLastRow();
  sh.getRange(newRow, 4).setNumberFormat('@').setValue(codes);
  if (data.kind === '休止') {
    return '「' + name + '」を休止（販売停止）として登録しました。\n\n対象商品コード：' + codes +
           '\n期間：' + start + '〜' + end +
           '\n\n次の朝の発注計算、または「仕込み・解凍表を作る」で、この期間の食材が自動で差し引かれます。';
  }
  const kindLabel = data.kind === '通常' ? '通常商品CP' : '新商品CP';
  return '「' + name + '」を' + kindLabel + 'として登録しました。\n\n対象商品コード：' + codes +
         '\n\n売上予測を更新すると発注・予測に反映されます。';
}

function CP_FORM_HTML_() {
  return `<!DOCTYPE html><html><head><base target="_top"><meta charset="utf-8">
<style>
body{font-family:'Meiryo',sans-serif;font-size:13px;padding:14px;color:#222}
h3{margin:0 0 4px 0}
label{display:block;margin-top:11px;font-weight:bold}
input,select{width:100%;padding:6px;margin-top:3px;box-sizing:border-box;border:1px solid #ccc;border-radius:3px}
.row2{display:flex;gap:8px}.row2>div{flex:1}
.note{color:#888;font-size:11px;margin-top:3px}
.pickrow{display:flex;gap:6px;margin-top:3px}
.pickrow input{flex:1}
.pickrow button{margin:0;padding:6px 14px;flex:0 0 auto}
button{margin-top:16px;padding:9px 22px;background:#4285f4;color:white;border:none;cursor:pointer;border-radius:3px;font-size:13px}
button.sub{background:#34a853}
#chips{margin-top:8px;display:flex;flex-wrap:wrap;gap:6px}
.chip{background:#e8f0fe;border:1px solid #a8c7fa;border-radius:14px;padding:3px 10px;font-size:12px;display:flex;align-items:center;gap:6px}
.chip span.x{cursor:pointer;color:#c0392b;font-weight:bold}
#stat{margin-top:12px;padding:9px;border-radius:4px;display:none;font-size:12px}
#stat.ok{background:#e6f4ea;color:#137333;display:block}
#stat.ng{background:#fce8e6;color:#c5221f;display:block}
</style></head><body>
<h3>キャンペーン商品の登録</h3>
<div class="note">CP設定シートに1行追加します。対象商品は検索して複数追加できます。</div>

<label>CP名 *</label>
<input type="text" id="c-name" placeholder="例: ねぎ塩PRキャンペーン">

<label>種類 *</label>
<select id="c-kind" onchange="kindChange()">
  <option value="新商品">新商品CP（新メニュー登場・実績なし）</option>
  <option value="通常">通常商品CP（既存商品がCPで伸びる）</option>
  <option value="休止">休止（一定期間の販売停止・食材を減らす）</option>
</select>

<div class="row2">
  <div><label>開始日 *</label><input type="date" id="c-start"></div>
  <div><label>終了日 *</label><input type="date" id="c-end"></div>
</div>

<label>対象商品 *（名前で検索して「追加」）</label>
<div class="pickrow">
  <input type="text" id="c-pick" list="prodlist" placeholder="読み込み中…" disabled>
  <button type="button" class="sub" onclick="addProd()">追加</button>
</div>
<datalist id="prodlist"></datalist>
<div id="chips"></div>

<div id="m-new">
  <label>予測出食率（客数の何％が頼むか）</label>
  <input type="number" id="c-serving" step="1" min="0" placeholder="空欄＝10％で見積り">
  <div class="note">例：客数の5％が頼むなら「5」。新商品の発注量の見積りに使います。実績が2日ぶん出たら自動で実測値に切り替わります。</div>
</div>

<div id="m-normal" style="display:none">
  <label>補正倍率（1.0が基準）</label>
  <input type="number" id="c-hosen" step="0.1" min="0.1" placeholder="空欄＝実績から自動計算">
  <div class="note">CPで普段の1.2倍出るなら「1.2」。空欄にすると実績から自動で倍率を学習します。</div>
</div>

<label>備考</label>
<input type="text" id="c-note" placeholder="例: 客数の10%を想定 / 60食">

<button onclick="submit()">登録する</button>
<div id="stat"></div>

<script>
var selected = [];   // {code,name}
function msg(text, ok){ var s=document.getElementById('stat'); s.textContent=text; s.className=ok?'ok':'ng'; }
function kindChange(){
  var k=document.getElementById('c-kind').value;
  // 休止は出食率・補正倍率とも不要（両方隠す）。新商品=出食率、通常=補正倍率。
  document.getElementById('m-new').style.display   = (k==='新商品') ? '' : 'none';
  document.getElementById('m-normal').style.display = (k==='通常') ? '' : 'none';
}
function renderChips(){
  var box=document.getElementById('chips'); box.innerHTML='';
  selected.forEach(function(p,i){
    var d=document.createElement('div'); d.className='chip';
    d.innerHTML='<span>'+p.code+'：'+p.name+'</span><span class="x" onclick="removeProd('+i+')">×</span>';
    box.appendChild(d);
  });
}
function removeProd(i){ selected.splice(i,1); renderChips(); }
function addProd(){
  var v=document.getElementById('c-pick').value.trim();
  var m=v.match(/^(\\d+)/);
  if(!m){ msg('商品を一覧から選んでください（先頭が商品コードの候補）', false); return; }
  var code=Number(m[1]);
  if(selected.some(function(p){return p.code===code;})){ msg('その商品はすでに追加済みです', false); return; }
  var hit=window._prodMap[code];
  selected.push({code:code, name:hit||''});
  document.getElementById('c-pick').value='';
  document.getElementById('stat').style.display='none';
  renderChips();
}
function submit(){
  var name=document.getElementById('c-name').value.trim();
  var start=document.getElementById('c-start').value;
  var end=document.getElementById('c-end').value;
  if(!name){ msg('CP名を入れてください', false); return; }
  if(!start||!end){ msg('開始日・終了日を入れてください', false); return; }
  if(selected.length===0){ msg('対象商品を1つ以上追加してください', false); return; }
  var kind=document.getElementById('c-kind').value;
  var rate = kind==='通常' ? document.getElementById('c-hosen').value : document.getElementById('c-serving').value;
  var codes=selected.map(function(p){return p.code;}).join(',');
  msg('登録中…', true);
  google.script.run
    .withSuccessHandler(function(r){ msg(r, true); setTimeout(function(){google.script.host.close();}, 1800); })
    .withFailureHandler(function(e){ msg('エラー: '+e, false); })
    .registerCpProduct({name:name, kind:kind, start:start, end:end, codes:codes, rate:rate,
      note:document.getElementById('c-note').value});
}
// 商品一覧を読み込んで datalist を作る
window._prodMap={};
google.script.run.withSuccessHandler(function(d){
  var dl=document.getElementById('prodlist');
  var frag=document.createDocumentFragment();
  d.products.forEach(function(p){
    window._prodMap[p.code]=p.name;
    var o=document.createElement('option'); o.value=p.code+'：'+p.name; frag.appendChild(o);
  });
  dl.appendChild(frag);
  var pick=document.getElementById('c-pick');
  pick.disabled=false; pick.placeholder='商品名を入力して候補から選ぶ（'+d.products.length+'件）';
}).withFailureHandler(function(e){
  document.getElementById('c-pick').placeholder='商品一覧の読込に失敗: '+e;
}).getCpFormData();
</script></body></html>`;
}


// ─────────────────────────────────────────────────────────────────────
// コードの稼働チェック（旧コード・再利用・切替の検出）
//   実績データ（POSのメニューCD×メニュー名×日付×出食数）と商品マスタ（レシピ）から、
//   「最終販売日」を軸に、使用中／休止候補／旧コード削除候補／再利用の疑い／マスタ未登録を判定。
//   ・名前そのまま切替 → 同名の別コードで検出
//   ・名前が変わった切替（Pﾗｲｽ→ﾍﾟｯﾊﾟｰﾗｲｽ）→ レシピ一致度（食材コード＋使用量／量は±許容）で検出
//   ・別商品への再利用 → 同一コードで過去に別のメニュー名
//   自動では消さない・書き換えない。判断材料として「稼働チェック」シートに出すだけ。
function buildCodeUsageCheck() { buildCodeUsageCheck_(); }
function buildCodeUsageCheck_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = 'Asia/Tokyo';
  const sheetName = '稼働チェック';
  let sh = ss.getSheetByName(sheetName);

  // 設定（既存シートがあれば店長が変えた値を引き継ぐ）
  let dormantDays = 90, simThresh = 90, amtTol = 20, reuseGapDays = 60;
  if (sh) {
    const gv = r => { const v = sh.getRange(r, 2).getValue(); return (typeof v === 'number' && v > 0) ? v : null; };
    dormantDays = gv(1) || dormantDays; simThresh = gv(2) || simThresh; amtTol = gv(3) || amtTol; reuseGapDays = gv(4) || reuseGapDays;
  }
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const dormantCutoff = new Date(today); dormantCutoff.setDate(today.getDate() - dormantDays);
  // 名前の正規化：空白・記号（()（）/／・,.。、-_）を落として「表記ゆれ」を同じ名前とみなす。数字(150/200)は残す
  const norm = s => String(s == null ? '' : s).replace(/[\s　()（）\/／・,，.。、\-ー_]/g, '').trim();
  const fmt = d => d ? Utilities.formatDate(d, tz, 'yyyy/MM/dd') : '';

  // ── 実績データ集計：コードごとに 初回/最終販売日・名前の履歴 ──
  const js = ss.getSheetByName('実績データ');
  const codeAgg = {};   // code -> {first,last,total,lastName, names:{normName:{min,max,cnt,raw}}}
  if (js && js.getLastRow() > 1) {
    js.getRange(2, 1, js.getLastRow() - 1, 7).getValues().forEach(r => {
      const cnt = Number(r[6]) || 0; if (cnt === 0) return;
      const code = Number(r[4]); if (!code) return;
      let d = r[0] instanceof Date ? new Date(r[0]) : new Date(String(r[0])); if (isNaN(d.getTime())) return; d.setHours(0, 0, 0, 0);
      const raw = String(r[5] || '').trim(), nm = norm(r[5]);
      let a = codeAgg[code]; if (!a) a = codeAgg[code] = { first: d, last: d, total: 0, lastName: raw, names: {} };
      if (d < a.first) a.first = d;
      if (d > a.last) { a.last = d; a.lastName = raw; }
      a.total += cnt;
      let nn = a.names[nm]; if (!nn) nn = a.names[nm] = { min: d, max: d, cnt: 0, raw: raw };
      if (d < nn.min) nn.min = d; if (d > nn.max) nn.max = d; nn.cnt += cnt;
    });
  }

  // ── 商品マスタ：レシピ（食材コード→使用量）・商品名・価格 ──
  const pm = ss.getSheetByName('商品マスタ');
  const recipe = {};   // code -> {ing:{食材コード:使用量}, name, price}
  if (pm && pm.getLastRow() > 1) {
    pm.getRange(2, 1, pm.getLastRow() - 1, 15).getValues().forEach(r => {
      const code = Number(r[0]); if (!code) return;
      let rc = recipe[code]; if (!rc) rc = recipe[code] = { ing: {}, name: String(r[1] || '').trim(), price: Number(r[2]) || 0 };
      const ing = Number(r[4]); const use = Number(r[12]) || 0;
      if (ing) rc.ing[ing] = (rc.ing[ing] || 0) + use;
    });
  }

  // ── CP設定：進行中CPの商品コード（実績が無くても使用中扱い）──
  const activeCpCodes = {};
  const cp = ss.getSheetByName('CP設定');
  if (cp && cp.getLastRow() > 1) {
    cp.getRange(2, 1, cp.getLastRow() - 1, 9).getValues().forEach(row => {
      const eRaw = row[2];
      const eDate = eRaw instanceof Date ? eRaw : (eRaw ? new Date(String(eRaw).replace(/^(\d{4})(\d{2})(\d{2})$/, '$1/$2/$3')) : null);
      const ongoing = !eDate || isNaN(eDate.getTime()) || eDate >= today;   // 終了日未定 or 未来＝進行中
      if (!ongoing) return;
      const codes = (typeof row[3] === 'number') ? [Math.round(row[3])]
        : String(row[3] || '').split(',').flatMap(c => { c = c.trim(); const n = Number(c); return n > 0 ? [n] : []; });
      codes.forEach(c => activeCpCodes[c] = true);
    });
  }

  const isActive = code => activeCpCodes[code] ? true : (codeAgg[code] && codeAgg[code].last >= dormantCutoff);

  // 同名（正規化した最新POS名）の別コード
  const nameToCodes = {};
  Object.keys(codeAgg).forEach(c => { const nm = norm(codeAgg[c].lastName); if (nm) (nameToCodes[nm] = nameToCodes[nm] || []).push(Number(c)); });

  // 食材コード → 現行コード の索引（レシピ照合を軽くするため）
  const ingToActive = {};
  Object.keys(recipe).forEach(c => { c = Number(c); if (!isActive(c)) return; Object.keys(recipe[c].ing).forEach(ing => (ingToActive[ing] = ingToActive[ing] || []).push(c)); });

  // レシピ一致度（%）＝ 量も近い共通食材 ÷ 和集合。量は±amtTol%まで同じとみなす
  const similarity = (ra, rb) => {
    const keys = {}; Object.keys(ra).forEach(k => keys[k] = 1); Object.keys(rb).forEach(k => keys[k] = 1);
    const ks = Object.keys(keys); if (ks.length === 0) return 0;
    let m = 0;
    ks.forEach(k => { const a = ra[k], b = rb[k]; if (a != null && b != null) { const hi = Math.max(a, b), lo = Math.min(a, b); if (hi === 0 || (hi - lo) / hi <= amtTol / 100) m++; } });
    return m / ks.length * 100;
  };
  // ある古いコードに最も近い「現行コード」を探す（1つでも共通食材がある現行だけ比較）
  const bestRecipeMatch = code => {
    const rc = recipe[code]; if (!rc || Object.keys(rc.ing).length === 0) return null;
    const cand = {};
    Object.keys(rc.ing).forEach(ing => (ingToActive[ing] || []).forEach(ac => { if (ac !== code) cand[ac] = 1; }));
    let best = null, bestScore = 0;
    Object.keys(cand).forEach(ac => { ac = Number(ac); const s = similarity(rc.ing, recipe[ac].ing); if (s > bestScore) { bestScore = s; best = ac; } });
    return (best && bestScore >= simThresh) ? { code: best, score: Math.round(bestScore) } : null;
  };

  // ── 表A：商品コードの稼働＆再利用チェック ──
  //   ★「今、対応が必要なものだけ」を出す。通常の使用中・相方のいない古いコードは出さない。
  const rowsA = [];
  const DAY = 1000 * 60 * 60 * 24;
  Object.keys(codeAgg).map(Number).forEach(code => {
    const a = codeAgg[code], rc = recipe[code];
    const posName = a.lastName;
    const masterName = rc ? rc.name : '';
    const active = isActive(code);
    const distinctNames = Object.keys(a.names).filter(n => n);
    let status = '', evidence = '';

    // ① 再利用の疑い：表記ゆれを除いても名前が2つ以上あり、かつ期間の間に空白があるものだけ（連続した改名は除外）
    if (distinctNames.length >= 2) {
      const ns = distinctNames.map(n => a.names[n]).sort((p, q) => p.max - q.max);
      let maxGap = 0, pair = null;
      for (let i = 1; i < ns.length; i++) {
        const gap = (ns[i].min - ns[i - 1].max) / DAY;
        if (gap > maxGap) { maxGap = gap; pair = [ns[i - 1], ns[i]]; }
      }
      if (maxGap >= reuseGapDays) {
        status = '⚠️再利用の疑い';
        evidence = '空白' + Math.round(maxGap) + '日あけて名前が変化：' + pair[0].raw + '→' + pair[1].raw;
      }
    }
    // ② マスタ未登録：今も売れている（最近）ものだけ＝要登録。古い未登録は履歴なので出さない
    if (!status && !rc) {
      if (active) { status = '⚠️マスタ未登録'; evidence = '実績にあるが商品マスタに無い→登録が必要（本部の新コードの可能性）'; }
      else return;
    }
    // ③ 旧コード削除候補／新コード（同名 or レシピ一致）。削除候補は「まだ商品マスタに残っている」ものだけ
    if (!status) {
      const nm = norm(posName);
      const nameSibs = (nameToCodes[nm] || []).filter(c => c !== code);
      if (active) {
        const older = nameSibs.find(c => !isActive(c) && recipe[c]);   // まだ残っている旧コードがある
        if (older) { status = '🆕新コード'; evidence = '旧コード=' + older + '（同名）→旧コードが削除候補'; }
        else return;   // 通常の使用中は出さない
      } else {
        if (!rc) return;   // すでにマスタに無い古いコードは対象外（掃除済み）
        const activeSib = nameSibs.find(c => isActive(c));
        if (activeSib) { status = '🔁旧コード削除候補'; evidence = '新コード=' + activeSib + '（同名：' + (codeAgg[activeSib] ? codeAgg[activeSib].lastName : '') + '）'; }
        else {
          const rm = bestRecipeMatch(code);
          if (rm) { status = '🔁旧コード削除候補'; evidence = '新コード=' + rm.code + '（レシピ' + rm.score + '%一致：' + (recipe[rm.code] ? recipe[rm.code].name : '') + '）'; }
          else return;   // 相方のいない古いコードは出さない（ただの古いコード＝ノイズ）
        }
      }
    }
    rowsA.push([code, posName, masterName, fmt(a.last), fmt(a.first), status, evidence]);
    rowsA[rowsA.length - 1]._st = status; rowsA[rowsA.length - 1]._last = a.last;
  });
  const prio = st => st.indexOf('削除候補') >= 0 ? 0 : st.indexOf('再利用') >= 0 ? 1 : st.indexOf('マスタ未登録') >= 0 ? 2 : st.indexOf('新コード') >= 0 ? 3 : st.indexOf('休止') >= 0 ? 4 : 5;
  rowsA.sort((p, q) => { const d = prio(p._st) - prio(q._st); return d ? d : (q._last ? q._last.getTime() : 0) - (p._last ? p._last.getTime() : 0); });

  // ── 表B：食材の稼働チェック ──
  const ingToCodes = {};
  Object.keys(recipe).forEach(c => { c = Number(c); Object.keys(recipe[c].ing).forEach(ing => (ingToCodes[ing] = ingToCodes[ing] || []).push(c)); });
  const fm = ss.getSheetByName('食材マスタ');
  const rowsB = [];
  if (fm && fm.getLastRow() > 1) {
    fm.getRange(2, 1, fm.getLastRow() - 1, 3).getValues().forEach(r => {
      const ing = Number(r[0]); if (!ing) return;
      const fname = String(r[1] || '').trim(), cat = String(r[2] || '').trim();
      const prods = ingToCodes[ing] || [];
      let lastD = null; const currentCodes = [];
      prods.forEach(c => { const a = codeAgg[c]; if (a) { if (!lastD || a.last > lastD) lastD = a.last; if (isActive(c)) currentCodes.push(c); } });
      const cpActive = prods.some(c => activeCpCodes[c]);
      let judge;
      if (cat === '発注終了') judge = '発注終了';
      else if (cpActive) judge = '使用中(CP)';
      else if (lastD && lastD >= dormantCutoff) judge = '使用中';
      else judge = '休止候補';
      rowsB.push([ing, fname, cat, fmt(lastD), currentCodes.slice(0, 12).join(','), judge]);
      rowsB[rowsB.length - 1]._j = judge; rowsB[rowsB.length - 1]._last = lastD;
    });
  }
  rowsB.sort((p, q) => { const jp = j => j.indexOf('休止') >= 0 ? 0 : 1; const d = jp(p._j) - jp(q._j); return d ? d : (q._last ? q._last.getTime() : 0) - (p._last ? p._last.getTime() : 0); });

  // ── シート出力 ──
  if (!sh) sh = ss.insertSheet(sheetName); else { sh.clear(); sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearDataValidations(); }
  const bgOf = st => st.indexOf('削除候補') >= 0 || st.indexOf('再利用') >= 0 || st.indexOf('未登録') >= 0 ? '#fce8e6'
    : st.indexOf('新コード') >= 0 ? '#fef7e0' : (st.indexOf('休止') >= 0 || st === '発注終了') ? '#f0f0f0' : '#ffffff';

  // 設定ブロック
  sh.getRange(1, 1, 4, 2).setValues([['休止しきい値(日)', dormantDays], ['レシピ一致しきい値(%)', simThresh], ['量許容(±%)', amtTol], ['再利用とみなす空白(日)', reuseGapDays]]);
  sh.getRange(1, 1, 4, 1).setFontWeight('bold');
  sh.getRange(1, 4).setValue('※数値を変えて「コードの稼働チェック（旧コード・再利用・切替の検出）」を再実行すると反映されます。自動では何も消しません。').setFontColor('#888888');
  sh.getRange(2, 4).setValue('※この表は「今、対応が必要なもの」だけを表示します（通常の使用中・ただの古いコードは省略）。').setFontColor('#888888');
  sh.getRange(1, 6).setValue('作成: ' + Utilities.formatDate(new Date(), tz, 'yyyy/MM/dd HH:mm')).setFontColor('#888888');

  let r = 6;
  // 表A
  sh.getRange(r, 1).setValue('■ 商品コードの要対応リスト（' + rowsA.length + '件）　→ 消したい行の右端「削除する」にチェックし、メニュー「チェックした商品コードを削除待ちに移す」を実行').setFontWeight('bold').setFontSize(12); r++;
  const headA = ['商品コード', '最新POS名', '商品マスタ名', '最終販売日', '初回販売日', '状態', '根拠・関連コード', '削除する'];
  sh.getRange(r, 1, 1, headA.length).setValues([headA]).setFontWeight('bold').setBackground('#e8eaed'); r++;
  if (rowsA.length) {
    const firstDataRowA = r;
    const body = rowsA.map(x => x.slice(0, 7));
    sh.getRange(r, 1, body.length, 7).setValues(body);
    sh.getRange(r, 1, body.length, 7).setBackgrounds(rowsA.map(x => Array(7).fill(bgOf(x._st))));
    sh.getRange(firstDataRowA, 8, body.length, 1).insertCheckboxes();   // 「削除する」チェックボックス
    r += body.length;
  }
  r += 2;
  // 表B
  sh.getRange(r, 1).setValue('■ 食材の稼働チェック（' + rowsB.length + '件）').setFontWeight('bold').setFontSize(12); r++;
  const headB = ['食材コード', '食材名', '区分', '最終使用日', '使っている現行商品コード', '判定'];
  sh.getRange(r, 1, 1, headB.length).setValues([headB]).setFontWeight('bold').setBackground('#e8eaed'); r++;
  if (rowsB.length) {
    const body = rowsB.map(x => x.slice(0, 6));
    sh.getRange(r, 1, body.length, 6).setValues(body);
    sh.getRange(r, 1, body.length, 6).setBackgrounds(rowsB.map(x => Array(6).fill(x._j.indexOf('休止') >= 0 || x._j === '発注終了' ? '#f0f0f0' : '#ffffff')));
    r += body.length;
  }

  sh.setColumnWidth(1, 90); sh.setColumnWidth(2, 200); sh.setColumnWidth(3, 200);
  sh.setColumnWidth(4, 90); sh.setColumnWidth(5, 160); sh.setColumnWidth(6, 150); sh.setColumnWidth(7, 300); sh.setColumnWidth(8, 80);
  sh.setFrozenRows(0);
  const maxC = sh.getMaxColumns(); if (maxC > 8) sh.deleteColumns(9, maxC - 8);

  // サマリー件数
  const cntA = {};
  rowsA.forEach(x => cntA[x._st] = (cntA[x._st] || 0) + 1);
  const summary = ['🔁旧コード削除候補', '⚠️再利用の疑い', '⚠️マスタ未登録', '🆕新コード', '休止候補']
    .filter(k => cntA[k]).map(k => k + ':' + cntA[k]).join(' / ');
  try { ss.toast('稼働チェックを作成しました。要チェック → ' + (summary || 'なし'), '完了', 10); } catch (e) {}
  console.log('buildCodeUsageCheck_: A=' + rowsA.length + ' B=' + rowsB.length + ' / ' + summary);
}

// 稼働チェック表Aで「削除する」にチェックした商品コードを、商品マスタから _削除待ち_商品マスタ へ移す。
//   ・物理削除ではなく移動なので、いつでも元に戻せる（restoreArchivedProductCodes）
//   ・商品マスタのrow3（スピルの土台）は絶対に消さない（対象でもスキップ）
//   ・商品マスタの入力列だけ記録・復元し、F〜N（原価等のスピル）は触らない
const PM_ARCHIVE_SHEET_ = '_削除待ち_商品マスタ';
const PM_INPUT_COLS_ = [1, 2, 3, 4, 5, 12, 13, 15];   // A商品コード,B商品名,C販売価格,D税抜,E食材コード,Lレシピ単位,M単位使用量,O提供量（F〜K・N金額はスピル）
function archiveCheckedProductCodes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const chk = ss.getSheetByName('稼働チェック');
  if (!chk || chk.getLastRow() < 2) { ui.alert('先にメニュー「コードの稼働チェック」を実行してください。'); return; }
  // チェック（H列=8がTRUE）が入った行の商品コード（A列）を集める
  const cv = chk.getRange(1, 1, chk.getLastRow(), 8).getValues();
  const codeSet = {}, codes = [];
  cv.forEach(row => { if (row[7] === true) { const c = Number(row[0]); if (c && !codeSet[c]) { codeSet[c] = true; codes.push(c); } } });
  if (codes.length === 0) { ui.alert('「削除する」にチェックが入っていません。\n消したい行の右端のチェックボックスを入れてから実行してください。'); return; }

  const pm = ss.getSheetByName('商品マスタ');
  if (!pm) { ui.alert('商品マスタが見つかりません。'); return; }
  const pmLast = pm.getLastRow(), pmCols = pm.getLastColumn();
  const pmVals = pm.getRange(1, 1, pmLast, pmCols).getValues();
  // 対象行（A列=商品コードがcodeSet）。row1(見出し)・row3(スピルの土台)は除外
  const targetRows = [];
  let skippedFoundation = false;
  for (let i = 1; i < pmVals.length; i++) {
    const rowNum = i + 1;
    if (Number(pmVals[i][0]) && codeSet[Number(pmVals[i][0])]) {
      if (rowNum === 3) { skippedFoundation = true; continue; }   // スピルの土台は消さない
      targetRows.push(rowNum);
    }
  }
  if (targetRows.length === 0) { ui.alert('対象の行が商品マスタに見つかりませんでした。' + (skippedFoundation ? '\n（土台の行だったためスキップしました）' : '')); return; }

  const resp = ui.alert('確認',
    codes.length + '個の商品コード（合計' + targetRows.length + '行）を「' + PM_ARCHIVE_SHEET_ + '」に移します。\n' +
    '物理削除ではなく移動なので、あとで元に戻せます。\n\n実行しますか？', ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;

  // 移動先シート（無ければ作成）
  let arc = ss.getSheetByName(PM_ARCHIVE_SHEET_);
  if (!arc) {
    arc = ss.insertSheet(PM_ARCHIVE_SHEET_); arc.hideSheet();
    const hdr = pmVals[0].slice(); hdr.push('移動日時');
    arc.getRange(1, 1, 1, hdr.length).setValues([hdr]).setFontWeight('bold');
  }
  const now = new Date();
  const arcRows = targetRows.map(rn => { const v = pmVals[rn - 1].slice(); v.push(now); return v; });
  arc.getRange(arc.getLastRow() + 1, 1, arcRows.length, arcRows[0].length).setValues(arcRows);

  // 商品マスタから削除（下の行から消す＝行番号がずれない。row3は対象外なので安全）
  targetRows.sort((a, b) => b - a).forEach(rn => pm.deleteRow(rn));

  ui.alert('完了',
    codes.length + '個の商品コード（' + targetRows.length + '行）を「' + PM_ARCHIVE_SHEET_ + '」に移しました。\n' +
    (skippedFoundation ? '※一部は表の土台の行だったため残しました。\n' : '') +
    '元に戻すには メニュー「削除待ちから商品コードを元に戻す」。\n\n' +
    '「コードの稼働チェック」を再実行すると一覧が最新になります。', ui.ButtonSet.OK);
  console.log('archiveCheckedProductCodes: codes=' + codes.join(',') + ' rows=' + targetRows.length);
}

// _削除待ち_商品マスタ の中身を商品マスタへ戻す（入力列だけ書き戻し、原価等はスピルが再計算）
function restoreArchivedProductCodes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const arc = ss.getSheetByName(PM_ARCHIVE_SHEET_);
  if (!arc || arc.getLastRow() < 2) { ui.alert('削除待ちに商品コードはありません。'); return; }
  const n = arc.getLastRow() - 1;
  const resp = ui.alert('確認', '「' + PM_ARCHIVE_SHEET_ + '」の全' + n + '行を商品マスタに戻しますか？', ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;

  const pm = ss.getSheetByName('商品マスタ');
  if (!pm) { ui.alert('商品マスタが見つかりません。'); return; }
  const aVals = arc.getRange(2, 1, n, arc.getLastColumn()).getValues();
  const startRow = pm.getLastRow() + 1;
  // 入力列だけを書き戻す（F〜K・N のスピル結果セルには一切書かない＝#REF!事故を防ぐ）
  PM_INPUT_COLS_.forEach(ci => {
    pm.getRange(startRow, ci, n, 1).setValues(aVals.map(v => [v[ci - 1]]));
  });
  arc.deleteRows(2, n);   // 削除待ちを空にする
  ui.alert('完了', n + '行を商品マスタに戻しました。\n原価・食材名などは自動計算（スピル）で復活します。\n\n※並びは末尾に戻ります。必要なら並べ替えてください。', ui.ButtonSet.OK);
  console.log('restoreArchivedProductCodes: rows=' + n);
}


// ═══════════════════════════════════════════════════════════════════════
// 【F】 理論原価率セクション
// ═══════════════════════════════════════════════════════════════════════
/**
 * 理論原価率を算出して「原価分析」シートに出力する。
 *
 * 【計算のしくみ】
 *   理論原価  = Σ（出食数 × その商品の1食あたり原価）
 *               1食あたり原価は商品マスタの「金額」列(N)を商品コードで合計したもの
 *               （1商品が複数の食材行に分かれているため）
 *   税抜売上  = Σ（出食数 × その商品の税抜き価格）
 *               ★POS総売上ではなく明細から積み上げる。理論原価（分子）は明細に載った
 *               商品ぶんしか出せないため、分母だけPOS総売上にすると範囲が食い違い、
 *               原価率が低く出てしまう（実測で約5ポイント低かった）。
 *   理論原価率 = 理論原価 ÷ 税抜売上
 *               原価に含めるのは 食材・ドリンク・酒 のみ。消耗品は本部の定義に合わせて含めない
 *               （除外した消耗品費はサマリーに参考表示する）
 *
 * 【期間をまとめた原価率】
 *   月別・直近N日の原価率は「その期間の原価合計 ÷ 売上合計」で出す。
 *   日ごとの率を単純平均すると、売上の小さい日が売上の大きい日と同じ重みで
 *   効いてしまい実態からずれるため。
 *   月別は全期間、日別は直近 DAILY_DAYS 日ぶんを表示する。
 *
 * 【マスタ価格一致率について】
 *   Σ(出食数×販売価格) ÷ POS総売上。100%に近いほど商品マスタの価格が実態と合っている。
 *   継続的に低い場合は値上げ後の価格がマスタに未反映の可能性があり、原価率が高めに出る。
 *
 * 【注意】ここで出るのは理論原価率のみ。棚卸ベースの実際原価率は納品実績シートが
 *   空のため算出できない。仕入データが入れば「理論との差＝ロス」が見えるようになる。
 */
// ── 原価調整設定（本部の原価戻し・商品コード読み替え）───────────────
//   1枚のシートで2つのズレを吸収する。
//   ・読み替え先 … POSのメニューCDを商品マスタの商品コードに変換する。
//       例1) ワールドペッパーデーはPOSが「4995 2皿目」で計上するので 4960 へ読み替える。
//       例2) 商品コードが4200番台→5000番台に切り替わった旧コードを新コードへ読み替える。
//   ・店舗負担率 … レシピの原価のうち店舗が負担する割合（本部から原価戻しがある商品用）。
//       例) 4960 はレシピが2食分で1食分は本部が原価戻し → 0.5。空欄なら1.0（全額店舗負担）。
const COST_ADJUST_SHEET  = '原価調整設定';
const COST_ADJUST_HEADER = ['メニューCD', '商品名（メモ）', '読み替え先', '店舗負担率', '備考'];

function ensureCostAdjustSheet_(ss) {
  let sh = ss.getSheetByName(COST_ADJUST_SHEET);
  if (sh) return sh;
  sh = ss.insertSheet(COST_ADJUST_SHEET);
  sh.getRange(1, 1, 1, COST_ADJUST_HEADER.length).setValues([COST_ADJUST_HEADER])
    .setFontWeight('bold').setBackground('#e8eaed');
  sh.setFrozenRows(1);
  [100, 240, 100, 100, 460].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.appendRow(['4995', '2皿目', '4960', '', 'ワールドペッパーデー：POSが「2皿目」で計上するため4960へ読み替え']);
  sh.appendRow(['4960', 'ワールドペッパーライス', '', 0.5, 'レシピが2食分。1食分は本部から原価戻しがあるため店舗負担は半分']);
  return sh;
}

function readCostAdjust_(ss) {
  const remap = {}, burden = {};
  const sh = ss.getSheetByName(COST_ADJUST_SHEET);
  if (!sh || sh.getLastRow() < 2) return { remap: remap, burden: burden };
  sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues().forEach(r => {
    const code = (r[0] === '' || r[0] === null) ? '' : String(r[0]).trim();
    if (!code) return;
    const to = (r[2] === '' || r[2] === null) ? '' : String(r[2]).trim();
    if (to) remap[code] = to;
    const b = Number(r[3]);
    if (r[3] !== '' && r[3] !== null && !isNaN(b) && b >= 0) burden[code] = b;
  });
  return { remap: remap, burden: burden };
}

/**
 * 原価調整設定シートを作り、商品マスタ未登録のメニューCDのうち
 * 「商品名が完全に一致する商品」が1つだけあるものを読み替え候補として自動追記する。
 * 商品コードが切り替わった旧コード（4200番台→5000番台など）を拾うのが目的。
 * 名前が同じ商品が複数ある場合は誤対応の危険があるので候補にしない。
 */
function setupCostAdjustSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ensureCostAdjustSheet_(ss);
  const pmSheet = ss.getSheetByName('商品マスタ');
  const sdSheet = ss.getSheetByName('実績データ');
  if (!pmSheet || !sdSheet) { SpreadsheetApp.getUi().alert('商品マスタまたは実績データが見つかりません'); return; }

  const norm = s => String(s === null || s === undefined ? '' : s).replace(/[\s　]/g, '').toUpperCase();

  // 商品マスタ：商品名 → 商品コード（同名が複数ある名前は候補から外す）
  const registered = {}, nameToCode = {}, nameDup = {};
  pmSheet.getRange(2, 1, pmSheet.getLastRow() - 1, 2).getValues().forEach(r => {
    if (!r[0]) return;
    const code = String(r[0]).trim(), key = norm(r[1]);
    registered[code] = true;
    if (!key) return;
    if (nameToCode[key] && nameToCode[key] !== code) nameDup[key] = true;
    else nameToCode[key] = code;
  });

  // 実績データ：商品マスタに無いメニューCDを出食数つきで集める
  const unknown = {};
  sdSheet.getRange(2, 1, sdSheet.getLastRow() - 1, 7).getValues().forEach(r => {
    if (!r[0] || !r[4]) return;
    const code = String(r[4]).trim();
    if (registered[code]) return;
    if (!unknown[code]) unknown[code] = { name: String(r[5] || ''), qty: 0 };
    unknown[code].qty += Number(r[6]) || 0;
  });

  // すでに設定済みのメニューCDは触らない（手で直した内容を上書きしないため）
  const already = {};
  if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().forEach(r => {
      if (r[0] !== '' && r[0] !== null) already[String(r[0]).trim()] = true;
    });
  }

  const add = [], skipped = [];
  Object.keys(unknown).sort((a, b) => unknown[b].qty - unknown[a].qty).forEach(code => {
    if (already[code]) return;
    const u = unknown[code], key = norm(u.name);
    if (!key || nameDup[key] || !nameToCode[key]) { if (u.qty > 0) skipped.push(code + ' ' + u.name + '（' + Math.round(u.qty) + '食）'); return; }
    add.push([code, u.name, nameToCode[key], '', '自動候補：商品名が完全一致（' + Math.round(u.qty) + '食）']);
  });
  if (add.length > 0) sh.getRange(sh.getLastRow() + 1, 1, add.length, 5).setValues(add);

  const msg = '原価調整設定を更新しました。\n\n'
    + '読み替え候補を ' + add.length + '件 追加しました（商品名が完全に一致したもの）。\n'
    + '内容を確認して、違うものがあれば行ごと削除してください。\n\n'
    + (skipped.length > 0
        ? '名前が一致せず候補にできなかったメニュー（' + skipped.length + '件）:\n' + skipped.slice(0, 15).join('\n')
        : '')
    + '\n\nそのあと「理論原価率を計算」を実行してください。';
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { console.log(msg); }
}


// ─────────────────────────────────────────────────────────────────────
// 原価マップ（商品コード別の1食あたり原価）— buildCostAnalysis と月次保存で共通利用
//   itemMap[商品コード] = { name, gross(販売価格), net(税抜), cost(原価), consumable(消耗品費), badIng }
//   原価は「食材・ドリンク・酒」のみ。消耗品は cost に入れず consumable に別持ち。本部戻し(店舗負担率)も反映。
// ─────────────────────────────────────────────────────────────────────
function computeCostMaps_(ss) {
  const pmSheet = ss.getSheetByName('商品マスタ');
  ensureCostAdjustSheet_(ss);
  const { remap: codeRemap, burden: costBurden } = readCostAdjust_(ss);

  const EXCLUDE_CATEGORIES = ['消耗品'];
  const ingCatMap = {}, ingRows = [];
  const imSheet = ss.getSheetByName('食材マスタ');
  if (imSheet && imSheet.getLastRow() > 1) {
    imSheet.getRange(2, 1, imSheet.getLastRow() - 1, 7).getValues().forEach(r => {
      if (!r[0]) return;
      ingCatMap[String(Number(r[0]))] = String(r[2] || '').trim();   // C列 区分(発注グループ)
      ingRows.push({ code: String(Number(r[0])), name: String(r[1] || ''), cat: String(r[2] || '').trim(),
                     price: Number(r[3]) || 0, spec: Number(r[5]) || 0, deliv: Number(r[6]) || 0 });
    });
  }

  const itemMap = {}, recipeRows = [];
  if (pmSheet && pmSheet.getLastRow() > 1) {
    pmSheet.getRange(2, 1, pmSheet.getLastRow() - 1, 15).getValues().forEach(r => {
      const code = (r[0] === '' || r[0] === null) ? '' : String(r[0]).trim();
      if (!code) return;
      if (!itemMap[code]) itemMap[code] = { name: String(r[1] || ''), gross: Number(r[2]) || 0, net: Number(r[3]) || 0, cost: 0, consumable: 0, badIng: [] };
      const ingCode = (r[4] === '' || r[4] === null) ? '' : String(Number(r[4]));
      const cat = ingCatMap[ingCode] || '';
      const amt = Number(r[13]);
      if (r[13] === '' || r[13] === null || isNaN(amt)) itemMap[code].badIng.push(String(r[5] || r[4] || ''));
      else if (EXCLUDE_CATEGORIES.indexOf(cat) >= 0) itemMap[code].consumable += amt;
      else itemMap[code].cost += amt;
      const useQty = Number(r[12]) || 0, servQty = Number(r[14]) || 0;
      if (ingCode && useQty > 0 && servQty > 0) {
        recipeRows.push({ prod: code, prodName: String(r[1] || ''), ing: ingCode, ingName: String(r[5] || ''),
                          use: useQty, serv: servQty, ratio: Math.round(servQty / useQty * 100) / 100 });
      }
    });
  }
  // 本部から原価戻しがある商品は、店舗が負担する分だけを原価として扱う
  Object.keys(itemMap).forEach(c => {
    const b = costBurden[c];
    if (b === undefined || b === 1) return;
    const it = itemMap[c];
    it.cost *= b; it.consumable *= b; it.burden = b;
    it.name = it.name + '（本部戻し後 ' + Math.round(b * 100) + '%）';
  });
  return { codeRemap, costBurden, ingCatMap, ingRows, itemMap, recipeRows };
}

// ─────────────────────────────────────────────────────────────────────
// 月次の理論原価スナップショット保存
//   月末に「値」で保存しておくと、あとで古い商品コード・食材コードを消しても
//   その月の理論原価はそのまま残る（数式で都度計算しないため）。
//   実原価・実現率の列も用意（納品実績が稼働したら埋める。今は空）。
// ─────────────────────────────────────────────────────────────────────
const COST_HISTORY_SHEET_ = '原価月次履歴';
// 年月キーの正規化：文字列'2026/07'も 月末日付(15:00Z)も 'yyyy/MM' に揃える（日付はUTCで月を取る＝末日15:00Z対策）
function ymKey_(v) {
  if (v instanceof Date) return v.getUTCFullYear() + '/' + ('0' + (v.getUTCMonth() + 1)).slice(-2);
  const m = String(v == null ? '' : v).trim().match(/(\d{4})[\/\-](\d{1,2})/);
  return m ? m[1] + '/' + ('0' + Number(m[2])).slice(-2) : String(v == null ? '' : v).trim();
}
function ensureCostHistorySheet_(ss) {
  // I実原価 の隣に「原価戻し(手入力)」列を追加（実際原価率＝(実原価−原価戻し)÷税抜売上）
  const hdr = ['年月', '理論原価', '税抜売上(明細)', '理論原価率', 'POS総売上', 'マスタ価格一致率', 'カバー率', '消耗品費(参考)', '実原価', '原価戻し(手入力)', '実際原価率', '実現率', '計算日時', '備考'];
  let sh = ss.getSheetByName(COST_HISTORY_SHEET_);
  if (!sh) {
    sh = ss.insertSheet(COST_HISTORY_SHEET_);
    sh.getRange(1, 1, 1, hdr.length).setValues([hdr]).setFontWeight('bold').setBackground('#e8eaed');
    sh.setFrozenRows(1);
  } else {
    // 旧レイアウト（原価戻し列が無い13列）を移行：実原価(I列=9)の後ろに空列を1本挿入し、実際原価率以降を右へずらす
    const lc = Math.max(sh.getLastColumn(), 1);
    const h1 = sh.getRange(1, 1, 1, lc).getValues()[0];
    if (h1.indexOf('原価戻し(手入力)') < 0 && lc >= 13) sh.insertColumnAfter(9);
    sh.getRange(1, 1, 1, hdr.length).setValues([hdr]).setFontWeight('bold').setBackground('#e8eaed');
    sh.setFrozenRows(1);
  }
  sh.getRange('A:A').setNumberFormat('@');   // 年月は文字列（日付に化けさせない）
  ['D', 'F', 'G', 'K', 'L'].forEach(col => sh.getRange(col + ':' + col).setNumberFormat('0.0%'));
  ['B', 'C', 'E', 'H', 'I', 'J'].forEach(col => sh.getRange(col + ':' + col).setNumberFormat('#,##0'));
  [70, 90, 110, 80, 90, 110, 70, 100, 90, 100, 90, 80, 130, 260].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  // 既存データの自己修復：年月を'yyyy/MM'文字列に正規化＋重複月を除去（先の行を残す）
  if (sh.getLastRow() > 1) {
    const n = sh.getLastRow() - 1;
    const norm = sh.getRange(2, 1, n, 1).getValues().map(r => [ymKey_(r[0])]);
    sh.getRange(2, 1, n, 1).setValues(norm);
    const seen = {}, dup = [];
    for (let i = 0; i < n; i++) { const k = norm[i][0]; if (seen[k]) dup.push(i + 2); else seen[k] = true; }
    dup.sort((x, y) => y - x).forEach(rn => sh.deleteRow(rn));
  }
  return sh;
}
function costHistoryMonths_(sh) {
  const s = {};
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().forEach(r => { const k = ymKey_(r[0]); if (k) s[k] = true; });
  return s;
}
// 実績データを1回なめて、年月ごとの原価・売上を集計
function aggregateMonthlyCosts_(ss, maps) {
  const tz = 'Asia/Tokyo';
  const sd = ss.getSheetByName('実績データ');
  const out = {};
  if (!sd || sd.getLastRow() < 2) return out;
  sd.getRange(2, 1, sd.getLastRow() - 1, 7).getValues().forEach(r => {
    if (!r[0]) return;
    const d = new Date(r[0]); if (isNaN(d.getTime())) return;
    const ym = Utilities.formatDate(d, tz, 'yyyy/MM');
    const ds = Utilities.formatDate(d, tz, 'yyyy/MM/dd');
    let o = out[ym]; if (!o) o = out[ym] = { cost: 0, net: 0, gross: 0, consumable: 0, qty: 0, unknownQty: 0, posTotal: 0, _pos: {} };
    const q = Number(r[6]) || 0, tot = Number(r[2]) || 0;
    if (!(ds in o._pos) || tot > o._pos[ds]) o._pos[ds] = tot;   // 総売上は同日同値なので日ごとに1回
    o.qty += q;
    const code0 = (r[4] === '' || r[4] === null) ? '' : String(r[4]).trim();
    const code = maps.codeRemap[code0] || code0;
    const it = maps.itemMap[code];
    if (!it) { o.unknownQty += q; return; }
    o.cost += it.cost * q; o.net += it.net * q; o.gross += it.gross * q; o.consumable += it.consumable * q;
  });
  Object.keys(out).forEach(ym => { const o = out[ym]; o.posTotal = Object.keys(o._pos).reduce((s, k) => s + o._pos[k], 0); delete o._pos; });
  return out;
}
// 指定した年月(ym='yyyy/MM')を履歴に1行追加（既に保存済みなら追加しない＝過去は上書きしない）
function appendCostSnapshot_(ss, ym, agg) {
  const sh = ensureCostHistorySheet_(ss);
  if (costHistoryMonths_(sh)[ym]) return { saved: false, reason: '既に保存済み' };
  const o = agg[ym];
  if (!o || o.qty === 0) return { saved: false, reason: 'データなし' };
  const rate = o.net > 0 ? o.cost / o.net : '';
  const fit = o.posTotal > 0 ? o.gross / o.posTotal : '';
  const cov = o.qty > 0 ? (o.qty - o.unknownQty) / o.qty : '';
  const note = [];
  if (cov !== '' && cov < 0.95) note.push('カバー率低：一部商品がマスタ未登録で原価が抜けている可能性');
  note.push('実原価は納品実績の稼働後に記録');
  sh.appendRow([ym, Math.round(o.cost), Math.round(o.net), rate, Math.round(o.posTotal), fit, cov, Math.round(o.consumable), '', '', '', '', new Date(), note.join(' / ')]);
  return { saved: true, rate: rate };
}

// メニュー：先月分を今すぐ保存
function saveLastMonthCostSnapshotNow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); const ui = SpreadsheetApp.getUi();
  const now = new Date(); const p = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const ym = Utilities.formatDate(p, 'Asia/Tokyo', 'yyyy/MM');
  const maps = computeCostMaps_(ss); const agg = aggregateMonthlyCosts_(ss, maps);
  const r = appendCostSnapshot_(ss, ym, agg);
  ui.alert(r.saved
    ? ('先月（' + ym + '）の理論原価を「' + COST_HISTORY_SHEET_ + '」に保存しました。' + (r.rate !== '' ? '\n理論原価率：' + (r.rate * 100).toFixed(1) + '%' : ''))
    : ('保存しませんでした：' + r.reason + '（' + ym + '）'));
}

// メニュー：過去の完了月（今月以外）をまとめて保存。コード削除の前に実行しておくと過去が守られる
function saveAllMonthlyCostSnapshots() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); const ui = SpreadsheetApp.getUi();
  const maps = computeCostMaps_(ss); const agg = aggregateMonthlyCosts_(ss, maps);
  const curYm = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM');
  const yms = Object.keys(agg).filter(m => m < curYm).sort();
  let n = 0, skip = 0; const done = [];
  yms.forEach(ym => { const r = appendCostSnapshot_(ss, ym, agg); if (r.saved) { n++; done.push(ym); } else skip++; });
  ui.alert('完了',
    '過去の完了月を保存しました。\n新規保存：' + n + '件' + (done.length ? '\n（' + done.join(', ') + '）' : '') +
    '\nスキップ（保存済み等）：' + skip + '件\n\n※今月（' + curYm + '）は未完了のため対象外です。', ui.ButtonSet.OK);
}

// メニュー：既存の原価月次履歴の「理論原価」側を、今の商品マスタ（食材復元後）で再計算して上書きする。
//   実原価・実際原価率・実現率（納品実績由来）は保持する。
//   食材マスタを消したまま保存してしまい原価が抜けて低く出た月を、あとから正しい値に直す用。
function refreshTheoreticalCostHistory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); const ui = SpreadsheetApp.getUi();
  const sh = ensureCostHistorySheet_(ss);
  if (sh.getLastRow() < 2) { ui.alert('原価月次履歴に月データがありません。先に「過去の全月を今すぐ保存」を実行してください。'); return; }
  const maps = computeCostMaps_(ss); const agg = aggregateMonthlyCosts_(ss, maps);
  const n = sh.getLastRow() - 1;
  const rng = sh.getRange(2, 1, n, 14);
  const vals = rng.getValues();
  const changed = []; let touched = 0;
  for (let i = 0; i < n; i++) {
    const ym = ymKey_(vals[i][0]);
    const o = agg[ym];
    if (!o || o.qty === 0) continue;   // 実績が無い月はそのまま
    const cost = Math.round(o.cost);
    const net = Math.round(o.net);
    const rate = o.net > 0 ? o.cost / o.net : '';
    const fit = o.posTotal > 0 ? o.gross / o.posTotal : '';
    const cov = o.qty > 0 ? (o.qty - o.unknownQty) / o.qty : '';
    const oldRate = Number(vals[i][3]) || 0;
    const actNum = Number(vals[i][8]) || 0;   // I 実原価（保持する）
    const rebate = Number(vals[i][9]) || 0;   // J 原価戻し（手入力・保持する）
    const netAct = actNum - rebate;           // 本部負担を引いた実質の実原価
    vals[i][0] = ym;                       // A 年月（文字列に正規化）
    vals[i][1] = cost;                     // B 理論原価
    vals[i][2] = net;                      // C 税抜売上(明細)
    vals[i][3] = rate;                     // D 理論原価率
    vals[i][4] = Math.round(o.posTotal);   // E POS総売上
    vals[i][5] = fit;                      // F マスタ価格一致率
    vals[i][6] = cov;                      // G カバー率
    vals[i][7] = Math.round(o.consumable); // H 消耗品費(参考)
    // I実原価・J原価戻し は触らない（棚卸実額と手入力を保持）
    if (actNum > 0) {                      // 実原価がある月は率も整合させる
      vals[i][10] = net > 0 ? netAct / net : '';       // K 実際原価率＝(実原価−原価戻し)÷税抜売上
      vals[i][11] = netAct > 0 ? cost / netAct : '';   // L 実現率＝理論÷(実原価−原価戻し)
    }
    vals[i][12] = new Date();              // M 計算日時
    const note = [];
    if (cov !== '' && cov < 0.95) note.push('カバー率低：一部商品がマスタ未登録で原価が抜けている可能性');
    note.push('食材復元後に理論原価を再計算');
    if (actNum <= 0) note.push('実原価は納品実績の稼働後に記録');
    else if (rebate > 0) note.push('原価戻し' + Math.round(rebate).toLocaleString() + '円を差引済');
    vals[i][13] = note.join(' / ');        // N 備考
    touched++;
    if (Math.abs((rate || 0) - oldRate) >= 0.0005) {
      changed.push(ym + '：' + (oldRate * 100).toFixed(1) + '% → ' + ((rate || 0) * 100).toFixed(1) + '%');
    }
  }
  const resp = ui.alert('理論原価の再計算',
    touched + '月分の理論原価を、今の商品マスタ（食材の復元後）で再計算して上書きします。\n' +
    '実原価・実際原価率・実現率（納品実績から入れた列）はそのまま残します。\n\n実行しますか？', ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) { ui.alert('中止しました。'); return; }
  rng.setValues(vals);
  ui.alert('完了', '理論原価を再計算しました（' + touched + '月）。\n\n変化した月：\n' +
    (changed.length ? changed.join('\n') : '（変化なし）'), ui.ButtonSet.OK);
  console.log('refreshTheoreticalCostHistory: touched=' + touched + ' changed=' + changed.length);
}

// トリガー本体：毎月1日に前月分を保存
function monthlyCostSnapshot() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const now = new Date(); const p = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const ym = Utilities.formatDate(p, 'Asia/Tokyo', 'yyyy/MM');
  const maps = computeCostMaps_(ss); const agg = aggregateMonthlyCosts_(ss, maps);
  const r = appendCostSnapshot_(ss, ym, agg);
  console.log('monthlyCostSnapshot ' + ym + ': ' + JSON.stringify(r));
  try {
    if (r.saved) {
      const o = agg[ym]; const rate = o.net > 0 ? (o.cost / o.net * 100).toFixed(1) : '-';
      MailApp.sendEmail(Session.getEffectiveUser().getEmail(), '【発注システム】' + ym + ' 理論原価を保存しました',
        ym + 'の理論原価率：' + rate + '%\n理論原価：' + Math.round(o.cost).toLocaleString() + '円\n税抜売上：' + Math.round(o.net).toLocaleString() + '円\n\n「' + COST_HISTORY_SHEET_ + '」シートに保存しました。');
    }
  } catch (e) {}
}

// メニュー：毎月1日の自動保存をON
function installMonthlyCostSnapshotTrigger() {
  const ui = SpreadsheetApp.getUi();
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'monthlyCostSnapshot') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('monthlyCostSnapshot').timeBased().onMonthDay(1).atHour(5).create();
  ui.alert('設定完了', '毎月1日の朝5時台に、前月分の理論原価を自動で「' + COST_HISTORY_SHEET_ + '」に保存します。\n（月初に前月が確定してから保存する安全な方式です）', ui.ButtonSet.OK);
}


// ─────────────────────────────────────────────────────────────────────
// 納品データCSV（商品別棚卸高）の取込 → 「納品実績」シート
//   本部の月次CSV（Shift-JIS）を マイドライブ「納品データ」フォルダに置いておくと、
//   取込済み以外の月を「納品実績」に追加する。列は検算で確定：
//     1:食材コード 2:商品名 3:仕入先 / 9:仕入数量 10:単位 11:仕入金額 /
//     13:単価 14:期末数量 16:期末金額 / 18:使用数量 20:使用金額(=実原価)
//   （0基インデックスでは code=0,name=1,仕入数量=8,単位=9,仕入金額=10,単価=12,期末数量=13,期末金額=15,使用数量=17,使用金額=19）
//   3月は期首在庫0のため「使用金額＝仕入−期末＝実際に使った原価」。実際原価率の元データになる。
// ─────────────────────────────────────────────────────────────────────
const DELIVERY_SHEET_ = '納品実績';
const DELIVERY_FOLDER_ = '納品データ';
function ensureDeliverySheet_(ss) {
  const hdr = ['年月', '食材コード', '商品名', '区分', '単価', '仕入数量', '単位', '仕入金額', '期末数量', '期末金額', '使用数量', '使用金額(実原価)', '取込日時'];
  let sh = ss.getSheetByName(DELIVERY_SHEET_);
  if (!sh) sh = ss.insertSheet(DELIVERY_SHEET_);
  // 1行目が見出しでなければ（空シート／見出し無しで素データが入っている等）先頭に見出し行を用意する
  const firstA = sh.getLastRow() >= 1 ? String(sh.getRange(1, 1).getValue()).trim() : '';
  if (firstA !== '年月') {
    sh.insertRowsBefore(1, 1);
    sh.getRange(1, 1, 1, hdr.length).setValues([hdr]).setFontWeight('bold').setBackground('#e8eaed');
    sh.setFrozenRows(1);
    [64, 90, 240, 90, 80, 80, 60, 100, 80, 100, 80, 120, 140].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  }
  sh.getRange('A:A').setNumberFormat('@');                                  // 年月は文字列（日付に化けさせない）
  ['E', 'H', 'J', 'L'].forEach(c => sh.getRange(c + ':' + c).setNumberFormat('#,##0'));
  // 年月列が日付で入っている行を 'yyyy/MM' 文字列に正規化（重複判定を安定させる）
  if (sh.getLastRow() > 1) {
    const rng = sh.getRange(2, 1, sh.getLastRow() - 1, 1);
    const vals = rng.getValues(); let changed = false;
    const out = vals.map(r => { let v = r[0]; if (v instanceof Date) { v = Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy/MM'); changed = true; } return [v]; });
    if (changed) rng.setValues(out);
  }
  return sh;
}

function importDeliveryInventoryCsv() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); const ui = SpreadsheetApp.getUi();
  const folders = DriveApp.getFoldersByName(DELIVERY_FOLDER_);
  if (!folders.hasNext()) { ui.alert('マイドライブに「' + DELIVERY_FOLDER_ + '」フォルダが見つかりません。'); return; }
  const folder = folders.next();

  // 食材マスタ → 区分
  const catMap = {};
  const im = ss.getSheetByName('食材マスタ');
  if (im && im.getLastRow() > 1) im.getRange(2, 1, im.getLastRow() - 1, 3).getValues().forEach(r => { if (r[0]) catMap[String(Number(r[0]))] = String(r[2] || '').trim(); });

  const sh = ensureDeliverySheet_(ss);
  const seen = {};   // 年月|コード
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(r => { seen[String(r[0]).trim() + '|' + Number(r[1])] = true; });

  const num = x => { const v = Number(String(x == null ? '' : x).replace(/,/g, '').trim()); return isNaN(v) ? '' : v; };
  const now = new Date();
  const toAppend = []; const filesDone = []; const filesSkip = [];
  const files = folder.getFiles();
  while (files.hasNext()) {
    const f = files.next();
    const fname = f.getName();
    if (!/\.csv$/i.test(fname)) continue;
    const mm = fname.match(/(\d{4})(\d{2})/);
    if (!mm) { filesSkip.push(fname + '（年月がファイル名から読めない）'); continue; }
    const ym = mm[1] + '/' + mm[2];
    let text;
    try { text = f.getBlob().getDataAsString('Shift_JIS'); }
    catch (e) { text = f.getBlob().getDataAsString('UTF-8'); }
    const rows = Utilities.parseCsv(text);
    let added = 0;
    rows.forEach(c => {
      const code = String(c[0] == null ? '' : c[0]).trim();
      if (!/^\d+$/.test(code)) return;                 // 見出し・合計など数字でない行は除外
      const key = ym + '|' + Number(code);
      if (seen[key]) return;                            // 取込済みはスキップ
      seen[key] = true; added++;
      toAppend.push([ym, Number(code), String(c[1] || '').trim(), catMap[String(Number(code))] || '',
        num(c[12]), num(c[8]), String(c[9] || '').trim(), num(c[10]),
        num(c[13]), num(c[15]), num(c[17]), num(c[19]), now]);
    });
    if (added > 0) filesDone.push(fname + '（' + added + '件）'); else filesSkip.push(fname + '（取込済み/対象なし）');
  }

  if (toAppend.length === 0) { ui.alert('取り込む新しいデータはありませんでした。', (filesSkip.join('\n') || 'CSVが見つかりません'), ui.ButtonSet.OK); return; }
  toAppend.sort((a, b) => a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : a[1] - b[1]);
  sh.getRange(sh.getLastRow() + 1, 1, toAppend.length, 13).setValues(toAppend);
  ui.alert('納品実績に取り込みました',
    toAppend.length + '件を追加しました。\n\n【取込】\n' + filesDone.join('\n') +
    (filesSkip.length ? '\n\n【スキップ】\n' + filesSkip.join('\n') : ''), ui.ButtonSet.OK);
  console.log('importDeliveryInventoryCsv: +' + toAppend.length + ' rows');
}

// 納品実績（使用金額＝実原価）を月ごとに集計し、「原価月次履歴」の 実原価/実際原価率/実現率 を埋める。
//   実原価＝Σ使用金額（食材・ドリンク・酒のみ。消耗品・未分類は除く＝理論原価と同じ範囲）
//   実際原価率＝(実原価 − 原価戻し[J列手入力]) ÷ その月の税抜売上（履歴C列） / 実現率＝理論原価 ÷ (実原価−原価戻し)（100%＝理論通り、低いほどロス大）
function updateActualCostFromDelivery() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); const ui = SpreadsheetApp.getUi();
  const del = ss.getSheetByName(DELIVERY_SHEET_);
  if (!del || del.getLastRow() < 2) { ui.alert('「' + DELIVERY_SHEET_ + '」にデータがありません。先に納品CSVを取り込んでください。'); return; }
  if (!ss.getSheetByName(COST_HISTORY_SHEET_)) { ui.alert('「' + COST_HISTORY_SHEET_ + '」がありません。先に「過去の全月を今すぐ保存」を実行してください。'); return; }
  const hist = ensureCostHistorySheet_(ss);   // 列レイアウトを最新化（原価戻し列が無ければ挿入してから処理）
  if (hist.getLastRow() < 2) { ui.alert('「' + COST_HISTORY_SHEET_ + '」に月データがありません。先に「過去の全月を今すぐ保存」を実行してください。'); return; }

  // 年月キーの正規化（文字列'2026/03' も 月末日付 も 'yyyy/MM' に揃える。日付はUTCで月を取る＝末日15:00Z対策）
  const ymKey = v => {
    if (v instanceof Date) return v.getUTCFullYear() + '/' + ('0' + (v.getUTCMonth() + 1)).slice(-2);
    const m = String(v == null ? '' : v).trim().match(/(\d{4})[\/\-](\d{1,2})/);
    return m ? m[1] + '/' + ('0' + Number(m[2])).slice(-2) : String(v == null ? '' : v).trim();
  };
  const COST_CATS = { '食材': 1, 'ドリンク': 1, '酒': 1 };

  // 納品実績を年月別に集計
  const agg = {};   // ym -> {actual, other}
  del.getRange(2, 1, del.getLastRow() - 1, 12).getValues().forEach(r => {
    const ym = ymKey(r[0]); if (!/^\d{4}\/\d{2}$/.test(ym)) return;
    const use = Number(r[11]) || 0;
    let a = agg[ym]; if (!a) a = agg[ym] = { actual: 0, other: 0 };
    if (COST_CATS[String(r[3] || '').trim()]) a.actual += use; else a.other += use;
  });

  // 履歴の行マップ（年月→行番号）。J列(10)＝原価戻し(手入力)も読む
  const hv = hist.getRange(2, 1, hist.getLastRow() - 1, 10).getValues();
  const rowOf = {};
  hv.forEach((r, i) => { rowOf[ymKey(r[0])] = { row: i + 2, theo: Number(r[1]) || 0, net: Number(r[2]) || 0, rebate: Number(r[9]) || 0 }; });

  const done = [], missing = [];
  Object.keys(agg).sort().forEach(ym => {
    const h = rowOf[ym];
    if (!h) { missing.push(ym); return; }
    const actual = agg[ym].actual;                                     // 棚卸ベースの実原価
    const netAct = actual - h.rebate;                                  // 本部原価戻しを引いた実質
    const rawRate = h.net > 0 ? actual / h.net : '';                   // 戻し前の実際原価率
    const actualRate = h.net > 0 ? netAct / h.net : '';                // 戻し後の実際原価率
    const achieve = netAct > 0 ? h.theo / netAct : '';                 // 実現率＝理論÷(実原価−原価戻し)
    hist.getRange(h.row, 9).setValue(Math.round(actual));             // I 実原価（棚卸実額）
    hist.getRange(h.row, 11, 1, 2).setValues([[actualRate, achieve]]); // K実際率 L実現率（J原価戻しは手入力なので触らない）
    const impact = (rawRate !== '' && actualRate !== '') ? (rawRate - actualRate) * 100 : 0;
    done.push(ym + '：実原価' + Math.round(actual).toLocaleString() + '円' +
      (h.rebate > 0 ? ' −原価戻し' + Math.round(h.rebate).toLocaleString() + '円' : '') +
      ' / 実際率' + (rawRate !== '' ? (rawRate * 100).toFixed(1) + '%' : '-') +
      (h.rebate > 0 ? '→' + (actualRate * 100).toFixed(1) + '%（−' + impact.toFixed(1) + 'pt）' : '') +
      ' / 実現率' + (achieve !== '' ? (achieve * 100).toFixed(0) + '%' : '-') +
      (agg[ym].other ? '（未計上' + Math.round(agg[ym].other).toLocaleString() + '円）' : ''));
  });

  ui.alert('実際原価率を反映しました',
    (done.length ? done.join('\n') : '更新対象なし') +
    (missing.length ? '\n\n【履歴に無い月＝先に理論原価を保存してください】\n' + missing.join(', ') : '') +
    '\n\n※実際原価率＝(実原価−原価戻し)÷税抜売上。原価戻し(J列)は本部請求の実額を手入力してから再実行すると反映されます。\n※実現率100%＝理論通り。下回るほどロス（廃棄・盛りすぎ・まかない・棚卸差異）が大きい。', ui.ButtonSet.OK);
  console.log('updateActualCostFromDelivery: ' + done.length + '月更新');
}

// ─────────────────────────────────────────────────────────────────────
// 実損分析：食材ごとに「実際の使用金額(納品実績) vs 理論の使用金額(出食数×レシピ)」を並べ、
//   差＝ロス（廃棄・盛りすぎ・まかない・棚卸差異）を月別・食材別に洗い出す。
//   ※肉の増量対象(クロッド・バラ等)の差には本部原価戻し分が含まれるため、実損はやや小さい。
// ─────────────────────────────────────────────────────────────────────
function buildLossAnalysis() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); const ui = SpreadsheetApp.getUi();
  const del = ss.getSheetByName(DELIVERY_SHEET_);
  if (!del || del.getLastRow() < 2) { ui.alert('「' + DELIVERY_SHEET_ + '」に納品データがありません。先に納品CSVを取り込んでください。'); return; }
  const maps = computeCostMaps_(ss);
  const tz = 'Asia/Tokyo';
  const ymKey = v => { if (v instanceof Date) return v.getUTCFullYear() + '/' + ('0' + (v.getUTCMonth() + 1)).slice(-2); const m = String(v == null ? '' : v).trim().match(/(\d{4})[\/\-](\d{1,2})/); return m ? m[1] + '/' + ('0' + Number(m[2])).slice(-2) : String(v == null ? '' : v).trim(); };

  // 食材コード → 名前（納品・食材マスタから）
  const ingName = {};
  maps.ingRows.forEach(r => { ingName[r.code] = r.name; });

  // 商品コード → その商品の食材別1食原価（消耗品除く・本部戻し前のフル原価。N列=r[13]）
  const prodIng = {};
  const pm = ss.getSheetByName('商品マスタ');
  if (pm && pm.getLastRow() > 1) {
    pm.getRange(2, 1, pm.getLastRow() - 1, 15).getValues().forEach(r => {
      const code = (r[0] === '' || r[0] === null) ? '' : String(r[0]).trim(); if (!code) return;
      const ing = (r[4] === '' || r[4] === null) ? '' : String(Number(r[4])); if (!ing) return;
      if ((maps.ingCatMap[ing] || '') === '消耗品') return;
      const amt = Number(r[13]); if (isNaN(amt)) return;
      if (!prodIng[code]) prodIng[code] = {};
      prodIng[code][ing] = (prodIng[code][ing] || 0) + amt;
    });
  }

  // 実績データ → 月×食材コード 理論使用金額（出食数×レシピ）
  const theoUse = {};
  const sd = ss.getSheetByName('実績データ');
  if (sd && sd.getLastRow() > 1) {
    sd.getRange(2, 1, sd.getLastRow() - 1, 7).getValues().forEach(r => {
      if (!r[0]) return; const d = new Date(r[0]); if (isNaN(d.getTime())) return;
      const ym = Utilities.formatDate(d, tz, 'yyyy/MM');
      const q = Number(r[6]) || 0; if (!q) return;
      const cd0 = (r[4] === '' || r[4] === null) ? '' : String(r[4]).trim();
      const cd = maps.codeRemap[cd0] || cd0;
      const pi = prodIng[cd]; if (!pi) return;
      let t = theoUse[ym]; if (!t) t = theoUse[ym] = {};
      for (const ing in pi) t[ing] = (t[ing] || 0) + q * pi[ing];
    });
  }

  // 納品実績 → 月×食材コード 実際使用金額（食材・ドリンク・酒のみ）
  const COST_CATS = { '食材': 1, 'ドリンク': 1, '酒': 1 };
  const actUse = {};
  del.getRange(2, 1, del.getLastRow() - 1, 12).getValues().forEach(r => {
    const ym = ymKey(r[0]); if (!/^\d{4}\/\d{2}$/.test(ym)) return;
    const cat = String(r[3] || '').trim(); if (!COST_CATS[cat]) return;
    const ing = String(Number(r[1])); const use = Number(r[11]) || 0;
    if (r[2] && !ingName[ing]) ingName[ing] = String(r[2] || '');
    let a = actUse[ym]; if (!a) a = actUse[ym] = {};
    if (!a[ing]) a[ing] = { cat: cat, actual: 0 };
    a[ing].actual += use;
  });

  // 履歴から月別 原価戻し
  const rebate = {};
  const hist = ss.getSheetByName(COST_HISTORY_SHEET_);
  if (hist && hist.getLastRow() > 1) hist.getRange(2, 1, hist.getLastRow() - 1, 10).getValues().forEach(r => { rebate[ymKey(r[0])] = Number(r[9]) || 0; });

  const months = Object.keys(actUse).sort();
  if (!months.length) { ui.alert('納品実績に集計できる月がありません。'); return; }

  // シート出力
  const name = '実損分析';
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name); else sh.clear();
  const out = [];
  out.push(['■ 実損分析：実際の使用金額(納品) − 理論の使用金額(出食数×レシピ) ＝ ロス', '', '', '', '', '', '', '']);
  out.push(['※差(ロス)＝廃棄・盛りすぎ・まかない・棚卸差異。肉の増量対象(クロッド・バラ等)の差には本部原価戻し分が含まれます。', '', '', '', '', '', '', '']);
  out.push(['', '', '', '', '', '', '', '']);

  // 月別サマリー
  out.push(['【月別サマリー】', '', '', '', '', '', '', '']);
  out.push(['年月', '実際使用計', '理論使用計', '差(ロス)', '原価戻し', '実質ロス', '実現率', '']);
  const detailBlocks = [];
  months.forEach(ym => {
    const a = actUse[ym] || {}, t = theoUse[ym] || {};
    const keys = {}; Object.keys(a).forEach(k => keys[k] = 1); Object.keys(t).forEach(k => keys[k] = 1);
    let actSum = 0, theoSum = 0; const rows = [];
    Object.keys(keys).forEach(ing => {
      const av = a[ing] ? a[ing].actual : 0;
      const tv = t[ing] || 0;
      actSum += av; theoSum += tv;
      const diff = av - tv;
      rows.push({ ing: ing, name: ingName[ing] || '', cat: (a[ing] && a[ing].cat) || maps.ingCatMap[ing] || '', act: av, theo: tv, diff: diff });
    });
    const reb = rebate[ym] || 0;
    const realLoss = (actSum - theoSum) - reb;
    const achieve = actSum > 0 ? theoSum / actSum : '';
    out.push([ym, Math.round(actSum), Math.round(theoSum), Math.round(actSum - theoSum), Math.round(reb), Math.round(realLoss), achieve, '']);
    rows.sort((x, y) => y.diff - x.diff);
    detailBlocks.push({ ym: ym, rows: rows });
  });
  out.push(['', '', '', '', '', '', '', '']);

  // 月別の食材ロス明細（差の大きい順・上位30）
  detailBlocks.slice().reverse().forEach(b => {   // 新しい月から
    out.push(['【' + b.ym + ' 食材別ロス（差の大きい順）】', '', '', '', '', '', '', '']);
    out.push(['食材コード', '食材名', '区分', '実際使用金額', '理論使用金額', '差(ロス)', '差率(実/理)', 'メモ']);
    const top = b.rows.filter(r => Math.abs(r.diff) >= 500).slice(0, 30);
    top.forEach(r => {
      const ratio = r.theo > 0 ? r.act / r.theo : (r.act > 0 ? '' : '');
      const memo = (r.theo === 0 && r.act > 0) ? '理論0＝レシピ/マスタ未登録の可能性' : (r.act === 0 && r.theo > 0) ? '納品実績に無し' : '';
      out.push([Number(r.ing), r.name, r.cat, Math.round(r.act), Math.round(r.theo), Math.round(r.diff), (r.theo > 0 ? ratio : ''), memo]);
    });
    const shown = top.reduce((s, r) => s + r.diff, 0);
    const totalDiff = b.rows.reduce((s, r) => s + r.diff, 0);
    out.push(['', '（上位30以外の差 合計）', '', '', '', Math.round(totalDiff - shown), '', '']);
    out.push(['', '', '', '', '', '', '', '']);
  });

  sh.getRange(1, 1, out.length, 8).setValues(out);
  sh.getRange(1, 1).setFontWeight('bold').setFontSize(11);
  sh.getRange(2, 1).setFontColor('#c00000');
  // 書式：金額列・率列
  sh.getRange(1, 4).setNumberFormat('#,##0');
  [70, 240, 70, 120, 120, 110, 90, 260].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  // 率列(G)を%表示・数値列を#,##0（ヘッダ行は除くのでざっくり列全体に適用）
  sh.getRange(1, 2, out.length, 5).setNumberFormat('#,##0');
  sh.getRange(1, 7, out.length, 1).setNumberFormat('0.0%');
  sh.setFrozenRows(0);
  const maxC = sh.getMaxColumns(); if (maxC > 8) sh.deleteColumns(9, maxC - 8);

  // 総括メッセージ
  const latest = months[months.length - 1];
  const la = actUse[latest], lt = theoUse[latest] || {};
  let laSum = 0, ltSum = 0; Object.keys(la).forEach(k => laSum += la[k].actual); Object.keys(lt).forEach(k => ltSum += lt[k]);
  ui.alert('実損分析を作成しました',
    'シート「' + name + '」に、月別サマリーと食材別ロス（差の大きい順）を出しました。\n\n' +
    '最新月（' + latest + '）：実際' + Math.round(laSum).toLocaleString() + '円 − 理論' + Math.round(ltSum).toLocaleString() +
    '円 ＝ 差' + Math.round(laSum - ltSum).toLocaleString() + '円' +
    ((rebate[latest] || 0) ? '（うち原価戻し' + Math.round(rebate[latest]).toLocaleString() + '円）' : '') +
    '\n\n差の大きい食材から、盛りすぎ・歩留・廃棄を見直すのが効果的です。', ui.ButtonSet.OK);
  console.log('buildLossAnalysis: months=' + months.join(','));
}

// ─────────────────────────────────────────────────────────────────────
// ドリンク実測分析：セルフ注ぎ(飲み放題)で1杯量が読めないドリンクを、
//   「実際のドリンク使用金額(納品) ÷ ドリンク出食数」で"実測の1杯単価"にする。
//   月が増えるほど累計平均が安定＝店長の当初プラン(出食数で割る)をそのまま実装。
// ─────────────────────────────────────────────────────────────────────
function analyzeDrinkUsage() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); const ui = SpreadsheetApp.getUi();
  const del = ss.getSheetByName(DELIVERY_SHEET_);
  if (!del || del.getLastRow() < 2) { ui.alert('「' + DELIVERY_SHEET_ + '」に納品データがありません。'); return; }
  const maps = computeCostMaps_(ss);
  const tz = 'Asia/Tokyo';
  const ymKey = v => { if (v instanceof Date) return v.getUTCFullYear() + '/' + ('0' + (v.getUTCMonth() + 1)).slice(-2); const m = String(v == null ? '' : v).trim().match(/(\d{4})[\/\-](\d{1,2})/); return m ? m[1] + '/' + ('0' + Number(m[2])).slice(-2) : String(v == null ? '' : v).trim(); };

  // ドリンクの食材コード集合（食材マスタ 区分=ドリンク）
  const drinkIng = {};
  Object.keys(maps.ingCatMap).forEach(c => { if (maps.ingCatMap[c] === 'ドリンク') drinkIng[c] = 1; });

  // ドリンクを使う商品（＝ドリンク提供商品）と、その商品の理論ドリンク原価(1食)・現状レシピ
  const drinkProd = {}, prodDrinkCost = {}, prodRecipe = {};
  const pm = ss.getSheetByName('商品マスタ');
  if (pm && pm.getLastRow() > 1) {
    pm.getRange(2, 1, pm.getLastRow() - 1, 15).getValues().forEach(r => {
      const code = (r[0] === '' || r[0] === null) ? '' : String(r[0]).trim(); if (!code) return;
      const ing = (r[4] === '' || r[4] === null) ? '' : String(Number(r[4])); if (!ing) return;
      if (!drinkIng[ing]) return;
      drinkProd[code] = String(r[1] || '');
      prodDrinkCost[code] = (prodDrinkCost[code] || 0) + (Number(r[13]) || 0);
      let rec = prodRecipe[code]; if (!rec) rec = prodRecipe[code] = [];
      rec.push({ ing: ing, ingName: String(r[5] || ''), use: Number(r[12]) || 0, unit: String(r[11] || ''), n: Number(r[13]) || 0 });
    });
  }

  // 実績データ：月別ドリンク出食数・理論ドリンク原価、商品別出食数
  const sd = ss.getSheetByName('実績データ');
  const qtyM = {}, theoM = {}, prodQty = {};
  if (sd && sd.getLastRow() > 1) {
    sd.getRange(2, 1, sd.getLastRow() - 1, 7).getValues().forEach(r => {
      if (!r[0]) return; const d = new Date(r[0]); if (isNaN(d.getTime())) return;
      const ym = Utilities.formatDate(d, tz, 'yyyy/MM');
      const q = Number(r[6]) || 0; if (!q) return;
      const cd0 = (r[4] === '' || r[4] === null) ? '' : String(r[4]).trim();
      const cd = maps.codeRemap[cd0] || cd0;
      if (!drinkProd[cd]) return;
      qtyM[ym] = (qtyM[ym] || 0) + q;
      theoM[ym] = (theoM[ym] || 0) + q * (prodDrinkCost[cd] || 0);
      let p = prodQty[cd]; if (!p) p = prodQty[cd] = { name: drinkProd[cd], total: 0 };
      p.total += q;
    });
  }

  // 納品実績：月別ドリンク実際使用金額＋ドリンク食材別
  const actM = {}, ingAct = {};
  del.getRange(2, 1, del.getLastRow() - 1, 12).getValues().forEach(r => {
    const ym = ymKey(r[0]); if (!/^\d{4}\/\d{2}$/.test(ym)) return;
    if (String(r[3] || '').trim() !== 'ドリンク') return;
    const use = Number(r[11]) || 0;
    actM[ym] = (actM[ym] || 0) + use;
    const ing = String(Number(r[1])); if (!ingAct[ing]) ingAct[ing] = { name: String(r[2] || ''), act: 0 }; ingAct[ing].act += use;
  });

  const months = Object.keys(actM).sort();
  if (!months.length) { ui.alert('納品実績にドリンクの月データがありません。'); return; }

  // 出力
  const name = 'ドリンク実測分析';
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name); else sh.clear();
  const out = [];
  out.push(['■ ドリンク実測分析：セルフ注ぎのため「実際使用金額 ÷ 出食数」で1杯単価を実測（月が増えるほど累計が安定）', '', '', '', '', '']);
  out.push(['', '', '', '', '', '']);
  out.push(['【月別】', '', '', '', '', '']);
  out.push(['年月', 'ドリンク実際使用金額', 'ドリンク出食数', '実測1杯単価', '理論1杯単価(現状)', '差(実測−理論)']);
  let sumAct = 0, sumQty = 0;
  months.forEach(ym => {
    const a = actM[ym] || 0, q = qtyM[ym] || 0, t = theoM[ym] || 0;
    sumAct += a; sumQty += q;
    const realU = q > 0 ? a / q : '';
    const theoU = q > 0 ? t / q : '';
    out.push([ym, Math.round(a), q, (realU === '' ? '' : Math.round(realU * 10) / 10), (theoU === '' ? '' : Math.round(theoU * 10) / 10), (realU === '' || theoU === '' ? '' : Math.round((realU - theoU) * 10) / 10)]);
  });
  const cumU = sumQty > 0 ? sumAct / sumQty : 0;
  out.push(['累計（全月）', Math.round(sumAct), sumQty, Math.round(cumU * 10) / 10, '', '']);
  out.push(['', '', '', '', '', '']);

  // ドリンク商品別の出食数（M/L等の内訳）
  out.push(['【ドリンク提供商品の出食数（全期間）】', '', '', '', '', '']);
  out.push(['商品コード', '商品名', '出食数', '', '', '']);
  Object.keys(prodQty).map(c => ({ c: c, n: prodQty[c].name, t: prodQty[c].total })).sort((a, b) => b.t - a.t)
    .forEach(o => out.push([Number(o.c), o.n, o.t, '', '', '']));
  out.push(['', '', '', '', '', '']);

  // ドリンク食材別の実際使用金額（何が高いか）
  out.push(['【ドリンク食材別の実際使用金額（全期間・高い順）】', '', '', '', '', '']);
  out.push(['食材コード', '食材名', '実際使用金額', '', '', '']);
  Object.keys(ingAct).map(c => ({ c: c, n: ingAct[c].name, a: ingAct[c].act })).sort((a, b) => b.a - a.a)
    .forEach(o => out.push([Number(o.c), o.n, Math.round(o.a), '', '', '']));
  out.push(['', '', '', '', '', '']);

  // 現状レシピ：ドリンク商品別の分量(単位使用量)とドリンク別割合（＝今の見積り。M/Lの分量・割合の記録）
  out.push(['【現状レシピ：ドリンク商品別の分量(単位使用量)と割合】', '', '', '', '', '']);
  out.push(['商品コード', '商品名', 'ドリンク食材', '単位使用量', '単位', '1食金額']);
  Object.keys(prodRecipe).map(c => ({ c: c, name: drinkProd[c], q: (prodQty[c] ? prodQty[c].total : 0), rec: prodRecipe[c] }))
    .sort((a, b) => b.q - a.q)
    .forEach(o => {
      o.rec.sort((x, y) => y.use - x.use);
      o.rec.forEach((r, i) => out.push([i === 0 ? Number(o.c) : '', i === 0 ? o.name : '', r.ingName, r.use, r.unit, Math.round(r.n * 10) / 10]));
      out.push(['', '（' + o.name + ' 合計）', '', '', '', Math.round(o.rec.reduce((s, r) => s + r.n, 0) * 10) / 10]);
    });

  sh.getRange(1, 1, out.length, 6).setValues(out);
  sh.getRange(1, 1).setFontWeight('bold').setFontSize(11);
  [70, 200, 150, 130, 90, 110].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  // 大きい金額も小さい単位使用量(0.002等)も両方見えるよう、小数を保持する書式にする
  sh.getRange(1, 2, out.length, 5).setNumberFormat('#,##0.#####');
  const maxC = sh.getMaxColumns(); if (maxC > 6) sh.deleteColumns(7, maxC - 6);

  ui.alert('ドリンク実測分析',
    '実測の1杯あたりドリンク原価（累計）：約' + (Math.round(cumU * 10) / 10).toLocaleString() + '円\n' +
    '（全期間：実際' + Math.round(sumAct).toLocaleString() + '円 ÷ 出食数' + sumQty.toLocaleString() + '杯）\n\n' +
    'シート「' + name + '」に月別・商品別・食材別を出しました。月が増えるほど累計単価が安定します。', ui.ButtonSet.OK);
  console.log('analyzeDrinkUsage: cum=' + cumU + ' months=' + months.join(','));
}

// ─────────────────────────────────────────────────────────────────────
// ドリンクの単位使用量を実測に合わせて更新。
//   コーラ等の割合・M:L比は維持したまま、全体の単位使用量を一律倍率(実測÷理論)で調整。
//   初月など在庫積上げで極端に低い月は実測から除外（中央値の60%未満）。
//   ★上書き前に現行レシピを「ドリンク単位使用量_履歴」に必ず保存（前回値の参照・復元用）。
// ─────────────────────────────────────────────────────────────────────
function adjustDrinkUsageToActual() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); const ui = SpreadsheetApp.getUi();
  const del = ss.getSheetByName(DELIVERY_SHEET_);
  if (!del || del.getLastRow() < 2) { ui.alert('「' + DELIVERY_SHEET_ + '」に納品データがありません。'); return; }
  const pm = ss.getSheetByName('商品マスタ');
  if (!pm || pm.getLastRow() < 2) { ui.alert('商品マスタがありません。'); return; }
  const maps = computeCostMaps_(ss);
  const tz = 'Asia/Tokyo';
  const ymKey = v => { if (v instanceof Date) return v.getUTCFullYear() + '/' + ('0' + (v.getUTCMonth() + 1)).slice(-2); const m = String(v == null ? '' : v).trim().match(/(\d{4})[\/\-](\d{1,2})/); return m ? m[1] + '/' + ('0' + Number(m[2])).slice(-2) : String(v == null ? '' : v).trim(); };

  const drinkIng = {};
  Object.keys(maps.ingCatMap).forEach(c => { if (maps.ingCatMap[c] === 'ドリンク') drinkIng[c] = 1; });

  // 商品マスタ ドリンクレシピ行（行番号・数式つき）
  const pmVals = pm.getRange(2, 1, pm.getLastRow() - 1, 15).getValues();
  const mForm = pm.getRange(2, 13, pmVals.length, 1).getFormulas();
  const lines = []; const drinkProd = {}, prodDrinkCost = {};
  pmVals.forEach((r, i) => {
    const code = (r[0] === '' || r[0] === null) ? '' : String(r[0]).trim(); if (!code) return;
    const ing = (r[4] === '' || r[4] === null) ? '' : String(Number(r[4])); if (!ing) return;
    if (!drinkIng[ing]) return;
    drinkProd[code] = String(r[1] || '');
    prodDrinkCost[code] = (prodDrinkCost[code] || 0) + (Number(r[13]) || 0);
    lines.push({ row: i + 2, prod: code, prodName: String(r[1] || ''), ing: ing, ingName: String(r[5] || ''), use: Number(r[12]) || 0, formula: mForm[i][0] || '', n: Number(r[13]) || 0 });
  });
  if (!lines.length) { ui.alert('ドリンクのレシピ行が見つかりません（食材マスタで区分＝ドリンクの食材を使う商品が無い）。'); return; }

  // 実績データ 月別 出食数・理論ドリンク原価
  const sd = ss.getSheetByName('実績データ');
  const qtyM = {}, theoM = {};
  if (sd && sd.getLastRow() > 1) sd.getRange(2, 1, sd.getLastRow() - 1, 7).getValues().forEach(r => {
    if (!r[0]) return; const d = new Date(r[0]); if (isNaN(d.getTime())) return;
    const ym = Utilities.formatDate(d, tz, 'yyyy/MM'); const q = Number(r[6]) || 0; if (!q) return;
    const cd0 = (r[4] === '' || r[4] === null) ? '' : String(r[4]).trim(); const cd = maps.codeRemap[cd0] || cd0;
    if (!drinkProd[cd]) return;
    qtyM[ym] = (qtyM[ym] || 0) + q; theoM[ym] = (theoM[ym] || 0) + q * (prodDrinkCost[cd] || 0);
  });

  // 納品 月別 実測ドリンク使用金額
  const actM = {};
  del.getRange(2, 1, del.getLastRow() - 1, 12).getValues().forEach(r => {
    const ym = ymKey(r[0]); if (!/^\d{4}\/\d{2}$/.test(ym)) return;
    if (String(r[3] || '').trim() !== 'ドリンク') return;
    actM[ym] = (actM[ym] || 0) + (Number(r[11]) || 0);
  });

  const months = Object.keys(actM).filter(m => qtyM[m] > 0).sort();
  if (!months.length) { ui.alert('実測できる月（納品＋出食数）がありません。'); return; }

  // 初月など極端に低い月（在庫積上げ）を実測から除外：1杯単価が中央値の60%未満を除外
  const perc = months.map(m => actM[m] / qtyM[m]).sort((a, b) => a - b);
  const med = perc[Math.floor(perc.length / 2)];
  const used = [], excluded = [];
  months.forEach(m => { (actM[m] / qtyM[m] < med * 0.6 ? excluded : used).push(m); });
  if (!used.length) { used.push.apply(used, months); excluded.length = 0; }

  let sumAct = 0, sumTheo = 0, sumQty = 0;
  used.forEach(m => { sumAct += actM[m]; sumTheo += theoM[m]; sumQty += qtyM[m]; });
  if (sumTheo <= 0) { ui.alert('理論ドリンク原価が0のため倍率を計算できません。'); return; }
  const k = sumAct / sumTheo;
  const realU = sumQty > 0 ? sumAct / sumQty : 0;
  const mOld = prodDrinkCost['3417'] || 0, lOld = prodDrinkCost['3418'] || 0;
  const r1 = x => Math.round(x * 10) / 10, r2 = x => Math.round(x * 100) / 100;

  const resp = ui.alert('ドリンクの単位使用量を実測に合わせる',
    '実測1杯単価：約' + r1(realU) + '円\n' +
    '採用月：' + used.join(', ') + (excluded.length ? '\n除外（初月等で低すぎ）：' + excluded.join(', ') : '') + '\n\n' +
    '倍率：現状の単位使用量を ×' + r2(k) + '（コーラ等の割合・M:L比は維持）\n' +
    (mOld ? 'ドリンクM：' + r1(mOld) + '円 → ' + r1(mOld * k) + '円\n' : '') +
    (lOld ? 'ドリンクL：' + r1(lOld) + '円 → ' + r1(lOld * k) + '円\n' : '') +
    '\n実行前に現行レシピを「ドリンク単位使用量_履歴」に保存します。実行しますか？', ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) { ui.alert('中止しました。'); return; }

  // ★バックアップ（履歴に追記・前回値を残す）
  const bkName = 'ドリンク単位使用量_履歴';
  let bk = ss.getSheetByName(bkName);
  if (!bk) { bk = ss.insertSheet(bkName); bk.getRange(1, 1, 1, 9).setValues([['保存日時', '商品コード', '商品名', '食材コード', '食材名', '旧単位使用量', '旧単位使用量(数式)', '旧金額', '適用倍率']]).setFontWeight('bold').setBackground('#e8eaed'); bk.setFrozenRows(1); }
  const now = new Date();
  const bkRows = lines.map(l => [now, Number(l.prod), l.prodName, Number(l.ing), l.ingName, l.use, l.formula, l.n, Math.round(k * 10000) / 10000]);
  bk.getRange(bk.getLastRow() + 1, 1, bkRows.length, 9).setValues(bkRows);

  // 単位使用量を ×k で書き換え（M列=13）。金額(N列)は自動再計算
  lines.forEach(l => { pm.getRange(l.row, 13).setValue(l.use * k); });

  ui.alert('完了',
    'ドリンクの単位使用量を ×' + r2(k) + ' に更新しました（' + lines.length + '行）。金額(N列)は自動再計算されます。\n\n' +
    '旧レシピは「' + bkName + '」に保存しました（前回値の参照・復元用）。\n\n' +
    'この後「理論原価率を計算」「実損分析」を実行すると、ドリンクが実態に合い、本当のロスがくっきりします。', ui.ButtonSet.OK);
  console.log('adjustDrinkUsageToActual: k=' + k + ' lines=' + lines.length + ' used=' + used.join(','));
}

// ─────────────────────────────────────────────────────────────────────
// 米使用分析：1食あたりの生米グラムを実データで出し、理論(レシピ)と実際(納品)を突き合わせる。
//   盛りすぎ・廃棄・おかわり・まかないを消した後の「レシピ側/棚卸側のズレ」を特定するための道具。
// ─────────────────────────────────────────────────────────────────────
function analyzeRiceUsage() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); const ui = SpreadsheetApp.getUi();
  const del = ss.getSheetByName(DELIVERY_SHEET_);
  if (!del || del.getLastRow() < 2) { ui.alert('「' + DELIVERY_SHEET_ + '」に納品データがありません。'); return; }
  const maps = computeCostMaps_(ss);
  const tz = 'Asia/Tokyo';
  const ymKey = v => { if (v instanceof Date) return v.getUTCFullYear() + '/' + ('0' + (v.getUTCMonth() + 1)).slice(-2); const m = String(v == null ? '' : v).trim().match(/(\d{4})[\/\-](\d{1,2})/); return m ? m[1] + '/' + ('0' + Number(m[2])).slice(-2) : String(v == null ? '' : v).trim(); };

  // 米の食材コード（名前に「米」を含む食材・消耗品除く）＋ 円/g
  const riceCodes = {}, yenPerG = {};
  maps.ingRows.forEach(r => { if (/米/.test(r.name) && r.cat !== '消耗品') { riceCodes[r.code] = r.name; yenPerG[r.code] = (r.spec > 0 ? r.price / r.spec : 0); } });
  if (!Object.keys(riceCodes).length) { ui.alert('「米」を含む食材が見つかりません。'); return; }

  // 商品別：米の生米g(単位使用量M)・盛りg(提供量O)・金額N
  const prodRice = {};
  const pm = ss.getSheetByName('商品マスタ');
  pm.getRange(2, 1, pm.getLastRow() - 1, 15).getValues().forEach(r => {
    const code = (r[0] === '' || r[0] === null) ? '' : String(r[0]).trim(); if (!code) return;
    const ing = (r[4] === '' || r[4] === null) ? '' : String(Number(r[4])); if (!riceCodes[ing]) return;
    if (!prodRice[code]) prodRice[code] = { name: String(r[1] || ''), raw: 0, serv: 0, yen: 0 };
    prodRice[code].raw += Number(r[12]) || 0;   // 生米g
    prodRice[code].serv += Number(r[14]) || 0;  // 盛りg
    prodRice[code].yen += Number(r[13]) || 0;
  });

  // 実績データ 月別：米提供の出食数・理論生米g・盛りg
  const sd = ss.getSheetByName('実績データ');
  const qtyM = {}, theoRawM = {}, servGM = {}, prodQty = {};
  if (sd && sd.getLastRow() > 1) sd.getRange(2, 1, sd.getLastRow() - 1, 7).getValues().forEach(r => {
    if (!r[0]) return; const d = new Date(r[0]); if (isNaN(d.getTime())) return;
    const ym = Utilities.formatDate(d, tz, 'yyyy/MM'); const q = Number(r[6]) || 0; if (!q) return;
    const cd0 = (r[4] === '' || r[4] === null) ? '' : String(r[4]).trim(); const cd = maps.codeRemap[cd0] || cd0;
    const pr = prodRice[cd]; if (!pr) return;
    qtyM[ym] = (qtyM[ym] || 0) + q; theoRawM[ym] = (theoRawM[ym] || 0) + q * pr.raw; servGM[ym] = (servGM[ym] || 0) + q * pr.serv;
    let p = prodQty[cd]; if (!p) p = prodQty[cd] = { name: pr.name, q: 0 }; p.q += q;
  });

  // 納品 月別：米の実際使用金額 → 実際生米g（金額 ÷ 円/g）
  const actRawM = {}, actYenM = {};
  del.getRange(2, 1, del.getLastRow() - 1, 12).getValues().forEach(r => {
    const ym = ymKey(r[0]); if (!/^\d{4}\/\d{2}$/.test(ym)) return;
    const ing = String(Number(r[1])); if (!riceCodes[ing]) return;
    const yen = Number(r[11]) || 0; const g = yenPerG[ing] > 0 ? yen / yenPerG[ing] : 0;
    actYenM[ym] = (actYenM[ym] || 0) + yen; actRawM[ym] = (actRawM[ym] || 0) + g;
  });

  const months = Object.keys(actRawM).filter(m => qtyM[m] > 0).sort();
  if (!months.length) { ui.alert('米の実測できる月がありません。'); return; }

  const name = '米使用分析';
  let sh = ss.getSheetByName(name); if (!sh) sh = ss.insertSheet(name); else sh.clear();
  const out = [];
  out.push(['■ 米使用分析：1食あたりの生米g（理論=レシピ／実際=納品）。盛り・廃棄・おかわり・まかない無しでの差＝レシピか棚卸のズレ', '', '', '', '', '', '', '']);
  out.push(['対象食材：' + Object.keys(riceCodes).map(c => c + ':' + riceCodes[c]).join(' / '), '', '', '', '', '', '', '']);
  out.push(['', '', '', '', '', '', '', '']);
  out.push(['【月別】', '', '', '', '', '', '', '']);
  out.push(['年月', '米提供の出食数', '理論生米g計', '実際生米g計', '実際/理論', '理論 生米g/食', '実際 生米g/食', '盛りg/食(理論)']);
  months.forEach(ym => {
    const q = qtyM[ym] || 0, tR = theoRawM[ym] || 0, aR = actRawM[ym] || 0, sG = servGM[ym] || 0;
    out.push([ym, Math.round(q), Math.round(tR), Math.round(aR), tR > 0 ? Math.round(aR / tR * 100) / 100 : '',
      q > 0 ? Math.round(tR / q) : '', q > 0 ? Math.round(aR / q) : '', q > 0 ? Math.round(sG / q) : '']);
  });
  out.push(['', '', '', '', '', '', '', '']);
  out.push(['【米を使う商品別レシピ（生米g＝単位使用量／盛りg＝提供量）と出食数（全期間）】', '', '', '', '', '', '', '']);
  out.push(['商品コード', '商品名', '生米g/食', '盛りg/食', '炊上り倍率(盛り/生米)', '出食数', '', '']);
  Object.keys(prodRice).map(c => ({ c: c, r: prodRice[c], q: prodQty[c] ? prodQty[c].q : 0 })).sort((a, b) => b.q - a.q)
    .forEach(o => out.push([Number(o.c), o.r.name, Math.round(o.r.raw * 10) / 10, Math.round(o.r.serv * 10) / 10,
      o.r.raw > 0 ? Math.round(o.r.serv / o.r.raw * 100) / 100 : '', o.q, '', '']));

  sh.getRange(1, 1, out.length, 8).setValues(out);
  sh.getRange(1, 1).setFontWeight('bold').setFontSize(11);
  [70, 240, 120, 120, 100, 120, 120, 120].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.getRange(1, 2, out.length, 7).setNumberFormat('#,##0.##');
  const maxC = sh.getMaxColumns(); if (maxC > 8) sh.deleteColumns(9, maxC - 8);

  const lm = months[months.length - 1];
  ui.alert('米使用分析',
    '最新月（' + lm + '）：理論 生米' + Math.round((theoRawM[lm] || 0) / (qtyM[lm] || 1)) + 'g/食 vs 実際 生米' + Math.round((actRawM[lm] || 0) / (qtyM[lm] || 1)) + 'g/食\n' +
    '（実際/理論＝' + (theoRawM[lm] > 0 ? Math.round(actRawM[lm] / theoRawM[lm] * 100) / 100 : '-') + '倍）\n\n' +
    'シート「' + name + '」に月別と商品別レシピを出しました。「炊上り倍率(盛り/生米)」が実態とズレていないかも確認できます。', ui.ButtonSet.OK);
  console.log('analyzeRiceUsage: months=' + months.join(','));
}

// 商品名から仕入規格量(g)を推定。「1kg」→1000、「500g」→500、「700ｇ」→700 等。無ければ null
function parseSpecGrams_(name) {
  let s = String(name == null ? '' : name).replace(/[０-９]/g, d => String.fromCharCode(d.charCodeAt(0) - 0xFEE0)).replace(/．/g, '.').replace(/　/g, ' ');
  const re = /(\d+(?:\.\d+)?)\s*(㎏|kg|ｋｇ|Ｋｇ|㎖|ml|ｍｌ|g|ｇ|Ｇ|Ｌ|L)/g;
  let m, last = null;
  while ((m = re.exec(s)) !== null) {
    const q = parseFloat(m[1]);
    const isBig = /㎏|kg|ｋｇ|Ｋｇ|Ｌ|L/.test(m[2]);
    last = isBig ? q * 1000 : q;
  }
  return last;
}

// 区分を食材コードの頭で推定（310/311=ドリンク, 34=酒, 636/638=消耗品, それ以外=食材）
function guessIngredientCat_(code) {
  const c = String(code);
  if (c.indexOf('310') === 0 || c.indexOf('311') === 0) return 'ドリンク';
  if (c.indexOf('34') === 0) return '酒';
  if (c.indexOf('636') === 0 || c.indexOf('638') === 0) return '消耗品';
  return '食材';
}

// 原価抜けチェック：商品マスタが使う食材コードのうち、食材マスタに無いもの（＝#N/Aで原価が抜ける）を一覧化。
//   納品実績に名前・単価があるものは「復元候補」として規格量つきで出す（そのまま食材マスタへ戻せる）。
function checkMissingIngredients() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); const ui = SpreadsheetApp.getUi();
  const im = ss.getSheetByName('食材マスタ');
  const imCodes = {};
  if (im && im.getLastRow() > 1) im.getRange(2, 1, im.getLastRow() - 1, 1).getValues().forEach(r => { if (r[0] !== '' && r[0] != null) imCodes[String(Number(r[0]))] = true; });

  const pm = ss.getSheetByName('商品マスタ');
  const miss = {};   // code -> {prods:{}, count}
  if (pm && pm.getLastRow() > 1) {
    pm.getRange(2, 1, pm.getLastRow() - 1, 5).getValues().forEach(r => {
      if (r[4] === '' || r[4] == null) return;
      const code = String(Number(r[4])); if (!/^\d+$/.test(code)) return;
      if (imCodes[code]) return;                       // 食材マスタにある＝OK
      let m = miss[code]; if (!m) m = miss[code] = { prods: {}, count: 0 };
      const pn = String(r[1] || '').trim(); if (pn) m.prods[pn] = true;
      m.count++;
    });
  }

  // 納品実績から 名前・単価（単価>0を優先）
  const del = ss.getSheetByName(DELIVERY_SHEET_);
  const info = {};
  if (del && del.getLastRow() > 1) del.getRange(2, 1, del.getLastRow() - 1, 5).getValues().forEach(r => {
    const code = String(Number(r[1])); if (!/^\d+$/.test(code)) return;
    const price = Number(r[4]) || 0;
    if (!info[code] || (price > 0 && !info[code].price)) info[code] = { name: String(r[2] || '').trim(), price: price };
  });

  const codes = Object.keys(miss);
  const rows = codes.map(code => {
    const it = info[code] || { name: '', price: '' };
    const prodNames = Object.keys(miss[code].prods);
    const restorable = it.name ? '可' : '要手動（納品実績に無い）';
    return { code: Number(code), name: it.name, cat: guessIngredientCat_(code), price: it.price,
      spec: parseSpecGrams_(it.name) || '', pcount: prodNames.length, pex: prodNames.slice(0, 4).join('、'),
      restorable: restorable, check: (it.name && guessIngredientCat_(code) !== '消耗品') };
  });
  rows.sort((a, b) => (b.pcount - a.pcount) || (String(a.restorable).localeCompare(String(b.restorable))));

  // シート出力
  const name = '原価抜けチェック';
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name); else { sh.clear(); sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearDataValidations(); }
  sh.getRange(1, 1).setValue('■ 商品マスタが使うのに食材マスタに無い食材（＝原価が#N/Aで抜けている）　→ 「復元する」にチェックし、メニュー「チェックした食材を食材マスタに復元」')
    .setFontWeight('bold').setFontSize(11);
  sh.getRange(2, 4).setValue('※単価・規格量は必要なら直してから復元してください（原価に直結）。規格量＝その払出単価で買えるg数。').setFontColor('#888888');
  const head = ['食材コード', '食材名', '区分', '払出単価', '仕入規格量(g)', '使う商品数', '使う商品(例)', '復元可否', '復元する'];
  sh.getRange(3, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#e8eaed');
  if (rows.length) {
    const body = rows.map(r => [r.code, r.name, r.cat, r.price, r.spec, r.pcount, r.pex, r.restorable]);
    sh.getRange(4, 1, body.length, 8).setValues(body);
    sh.getRange(4, 9, body.length, 1).insertCheckboxes();
    sh.getRange(4, 9, body.length, 1).setValues(rows.map(r => [r.check]));
    sh.getRange(4, 1, body.length, 9).setBackgrounds(rows.map(r => Array(9).fill(r.restorable === '可' ? '#fff8e1' : '#fce8e6')));
  }
  sh.setFrozenRows(3);
  [80, 260, 70, 80, 100, 80, 300, 130, 70].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  const maxC = sh.getMaxColumns(); if (maxC > 9) sh.deleteColumns(10, maxC - 9);
  const maxR = sh.getMaxRows(), last = 3 + rows.length; if (maxR > last + 2) sh.deleteRows(last + 3, maxR - last - 2);
  ui.alert('原価抜けチェック', rows.length + '件の食材が「商品マスタで使用中なのに食材マスタに無い」状態です。\n' +
    'シート「' + name + '」で確認し、戻したい食材の「復元する」にチェック → メニュー「チェックした食材を食材マスタに復元」。', ui.ButtonSet.OK);
  console.log('checkMissingIngredients: ' + rows.length + '件');
}

// 「原価抜けチェック」でチェックした食材を、食材マスタに追加（復元）する。
function applyIngredientRestore() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); const ui = SpreadsheetApp.getUi();
  const sh = ss.getSheetByName('原価抜けチェック');
  if (!sh || sh.getLastRow() < 4) { ui.alert('先にメニュー「原価抜けチェック」を実行してください。'); return; }
  const im = ss.getSheetByName('食材マスタ');
  if (!im) { ui.alert('食材マスタが見つかりません。'); return; }
  const imCodes = {};
  if (im.getLastRow() > 1) im.getRange(2, 1, im.getLastRow() - 1, 1).getValues().forEach(r => { if (r[0] !== '' && r[0] != null) imCodes[String(Number(r[0]))] = true; });

  const vals = sh.getRange(4, 1, sh.getLastRow() - 3, 9).getValues();
  const toAdd = []; const problems = [];
  vals.forEach(r => {
    if (r[8] !== true) return;                         // 復元するチェックが入っていない
    const code = Number(r[0]); if (!code) return;
    if (imCodes[String(code)]) return;                 // 既に食材マスタにある
    const name = String(r[1] || '').trim();
    const cat = String(r[2] || '食材').trim();
    const price = Number(r[3]) || 0;
    const spec = Number(r[4]) || 0;
    if (!name || price <= 0 || spec <= 0) { problems.push(code + ' ' + name + '（単価/規格量が未入力）'); return; }
    // 食材マスタ：A食材コード B名 C区分 D払出単価 E仕入単位 F仕入規格量 G仕入納品量 H歩留率 I安全係数
    toAdd.push([code, name, cat, price, 'ｇ', spec, spec, 1, 1.2, '', '', '', '', '']);
    imCodes[String(code)] = true;
  });

  if (toAdd.length === 0) {
    ui.alert('復元対象がありません', '「復元する」にチェックが入り、単価・規格量が入っている行がありませんでした。' +
      (problems.length ? '\n\n【単価/規格量が未入力】\n' + problems.join('\n') : ''), ui.ButtonSet.OK);
    return;
  }
  const resp = ui.alert('食材マスタに復元', toAdd.length + '件の食材を食材マスタに追加します。\n（区分・払出単価・規格量はチェックシートの値を使います）\n\n実行しますか？', ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;

  im.getRange(im.getLastRow() + 1, 1, toAdd.length, 14).setValues(toAdd);
  ui.alert('復元しました', toAdd.length + '件を食材マスタに追加しました。\n\n' +
    '・商品マスタの#N/Aが解消され、原価が計上されます\n' +
    '・「理論原価率を計算」で改善を確認できます\n' +
    '・次回の発注計算からその食材も発注対象になります' +
    (problems.length ? '\n\n【スキップ（要手入力）】\n' + problems.join('\n') : ''), ui.ButtonSet.OK);
  console.log('applyIngredientRestore: +' + toAdd.length);
}


function buildCostAnalysis() {
  const DAILY_DAYS = 120;   // 日別に表示する日数
  const ITEM_DAYS  = 30;    // 商品別集計の対象期間
  const ITEM_TOP   = 40;    // 商品別に表示する上位件数

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pmSheet = ss.getSheetByName('商品マスタ');
  const sdSheet = ss.getSheetByName('実績データ');
  if (!pmSheet || pmSheet.getLastRow() < 2) { console.log('商品マスタにデータがありません'); return; }
  if (!sdSheet || sdSheet.getLastRow() < 2) { console.log('実績データにデータがありません'); return; }

  // ── 原価マップ（商品コード別1食原価・読み替え等）を共通関数から取得 ──
  //    月次スナップショット保存(saveCostSnapshot系)と同じ計算を使うため共通化した。
  const { codeRemap, itemMap, ingRows, recipeRows } = computeCostMaps_(ss);

  // ── 実績データの最終日を基準に集計期間を決める ─────────────────────
  const sdRows = sdSheet.getRange(2, 1, sdSheet.getLastRow() - 1, 7).getValues();
  let maxDate = null;
  sdRows.forEach(r => {
    if (!r[0]) return;
    const d = new Date(r[0]); if (isNaN(d.getTime())) return;
    d.setHours(0, 0, 0, 0);
    if (!maxDate || d > maxDate) maxDate = d;
  });
  if (!maxDate) { console.log('実績データに有効な日付がありません'); return; }
  const dailyFrom = new Date(maxDate); dailyFrom.setDate(dailyFrom.getDate() - DAILY_DAYS + 1);
  const itemFrom  = new Date(maxDate); itemFrom.setDate(itemFrom.getDate()  - ITEM_DAYS  + 1);

  // ── 日別・商品別に集計 ─────────────────────────────────────────────
  //    月別は全期間を見たいので、ここでは日付で絞らず全行を日別に積み上げる
  //    （日別表示だけを後段で直近 DAILY_DAYS 日に絞る）
  const dayMap = {}, itemAgg = {}, unknownMap = {};
  sdRows.forEach(r => {
    if (!r[0]) return;
    const d = new Date(r[0]); if (isNaN(d.getTime())) return;
    d.setHours(0, 0, 0, 0);
    const ds    = toDateStr_(d);
    const code0 = (r[4] === '' || r[4] === null) ? '' : String(r[4]).trim();
    const code  = codeRemap[code0] || code0;      // POSのコードを商品マスタのコードへ読み替え
    const qty   = Number(r[6]) || 0;
    const tot  = Number(r[2]) || 0;
    if (!dayMap[ds]) dayMap[ds] = { date: d, dow: String(r[1] || ''), total: 0, cost: 0, consumable: 0, gross: 0, net: 0, qty: 0, unknownQty: 0 };
    const day = dayMap[ds];
    if (tot > day.total) day.total = tot;   // 総売上は同じ日の全行に同じ値が入っている
    day.qty += qty;
    const it = itemMap[code];
    if (!it) {
      day.unknownQty += qty;
      if (qty > 0) unknownMap[code0] = { name: String(r[5] || ''), qty: ((unknownMap[code0] && unknownMap[code0].qty) || 0) + qty };
      return;
    }
    day.cost       += it.cost       * qty;
    day.consumable += it.consumable * qty;
    day.gross      += it.gross      * qty;
    day.net        += it.net        * qty;
    if (d >= itemFrom) {
      if (!itemAgg[code]) itemAgg[code] = { name: it.name, net: it.net, cost: it.cost, qty: 0 };
      itemAgg[code].qty += qty;
    }
  });

  // ── 日別の指標を確定（古い順に並べて7日移動平均を計算）─────────────
  const allDays = Object.keys(dayMap).sort().map(ds => {
    const v = dayMap[ds];
    // 【重要】分母は「POS総売上」ではなく「明細から積み上げた税抜売上」を使う。
    //   理論原価（分子）は明細に載った商品ぶんしか計算できないため、分母にPOS総売上
    //   （明細に載らない売上を含む）を使うと範囲が食い違い、原価率が低く出てしまう。
    //   実測：明細がPOS総売上の約87%しかないため、旧方式は原価率が約5ポイント低かった
    //   （旧33.6% / 新38.8%）。明細が100%揃っている日は新旧どちらも約39%で一致し、
    //   本部の39〜40%とも合致することを実績30日で確認済み。
    const netSales = v.net;
    return {
      ds: ds, date: v.date, dow: v.dow, total: v.total, netSales: netSales, cost: v.cost,
      gross: v.gross,
      consumable: v.consumable,
      rate: netSales > 0 ? v.cost / netSales : 0,                      // 食材・ドリンク・酒のみ（消耗品は含めない）
      fit:  v.total > 0 ? v.gross / v.total : 0,
      cov:  v.qty > 0 ? (v.qty - v.unknownQty) / v.qty : 0,
      unknownQty: v.unknownQty
    };
  }).filter(x => x.total > 0 && x.netSales > 0);

  // ── その日の数字が信用できるか自動判定 ─────────────────────────────
  //    原価率がおかしく見える日は、たいてい原価ではなくデータ側の問題なので、
  //    機械的に切り分けて集計から外す（手で直さなくても数字が壊れないようにする）。
  //  ・カバー率（出食数のうち商品マスタにある割合）が低い
  //      → 当時のメニューが商品マスタに無い。原価が計上されず原価率が低く出る。
  //  ・マスタ価格一致率（Σ出食数×販売価格 ÷ POS総売上）が極端に低い
  //      → 売上に対して出食数データが足りない＝取込の欠損。原価も同じだけ抜ける。
  const MIN_COV = 0.95, MIN_FIT = 0.70;
  allDays.forEach(x => {
    x.ng = '';
    if (x.cov < MIN_COV)      x.ng = '要確認(マスタ未整備)';
    else if (x.fit < MIN_FIT) x.ng = '要確認(データ不備)';
  });
  const validDays = allDays.filter(x => !x.ng && x.rate > 0);
  validDays.forEach((x, i) => {
    const win = validDays.slice(Math.max(0, i - 6), i + 1);
    x.ma7 = win.reduce((s, y) => s + y.rate, 0) / win.length;   // 移動平均も有効日だけで計算
  });
  const days = allDays.slice(-DAILY_DAYS);   // 日別に表示する分だけ切り出す
  // 期間全体の原価率は「原価合計 ÷ 売上合計」（日ごとの率の単純平均ではない）
  const avgOf = (n, withCons) => {
    const win = validDays.slice(-n);
    if (win.length === 0) return 0;
    const c  = win.reduce((s, x) => s + x.cost + (withCons ? x.consumable : 0), 0);
    const s2 = win.reduce((s, x) => s + x.netSales, 0);
    return s2 > 0 ? c / s2 : 0;
  };

  // ── 月別に集計（全期間）─────────────────────────────────────────────
  //    月の原価率も「その月の原価合計 ÷ 売上合計」。日ごとの率を平均すると
  //    売上の小さい日が同じ重みで効いてしまい実態とずれるため。
  //    「要確認」の日は原価率の計算から外す。営業日数・POS総売上は実績どおり全日で出し、
  //    有効日数を並べて「何日ぶんで出した原価率か」が一目で分かるようにする。
  const monthMap = {};
  allDays.forEach(x => {
    const ym = x.ds.substring(0, 7);   // toDateStr_ は 'yyyy/MM/dd' なので先頭7文字が 'yyyy/MM'
    if (!monthMap[ym]) monthMap[ym] = { ym: ym, days: 0, validDays: 0, total: 0, vTotal: 0, vNetSales: 0, vCost: 0, vCons: 0, vGross: 0 };
    const m = monthMap[ym];
    m.days++; m.total += x.total;
    if (x.ng || x.rate <= 0) return;
    m.validDays++; m.vTotal += x.total; m.vNetSales += x.netSales; m.vCost += x.cost; m.vCons += x.consumable; m.vGross += x.gross;
  });
  const months = Object.keys(monthMap).sort().map(ym => {
    const m = monthMap[ym];
    m.rate    = m.vNetSales > 0 ? m.vCost / m.vNetSales : '';
    m.fit     = m.vTotal > 0 ? m.vGross / m.vTotal : '';
    m.dayAvg  = m.days > 0 ? m.total / m.days : 0;
    return m;
  });
  // 前月差は、前後とも原価率が出せている月どうしでのみ計算する
  months.forEach((m, i) => {
    const prev = i > 0 ? months[i - 1] : null;
    m.diff = (prev && m.rate !== '' && prev.rate !== '') ? m.rate - prev.rate : '';
  });

  // ── シートへ出力 ───────────────────────────────────────────────────
  let sh = ss.getSheetByName('原価分析');
  if (!sh) sh = ss.insertSheet('原価分析');
  sh.clear();
  const W = 9;
  const pad = a => { const b = a.slice(); while (b.length < W) b.push(''); return b; };
  const rows = [];
  rows.push(pad(['理論原価率 分析']));
  rows.push(pad(['更新日時', Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm'), '', '対象最終日', Utilities.formatDate(maxDate, 'Asia/Tokyo', 'yyyy/MM/dd')]));
  rows.push(pad(['直近7日 原価率', avgOf(7, false), '直近30日 原価率', avgOf(30, false), '直近90日 原価率', avgOf(90, false)]));
  const cons30 = validDays.slice(-30).reduce((s, x) => s + x.consumable, 0);
  rows.push(pad(['※原価は 食材・ドリンク・酒 のみ（消耗品は含めない）。分母は「明細から積み上げた税抜売上」で、'
    + '理論原価（分子）と範囲を揃えている（分母をPOS総売上にすると原価率が低く出る）。'
    + '「要確認」の日は計算から自動除外。参考：除外した消耗品費(直近30日) ' + Math.round(cons30).toLocaleString() + '円']));
  rows.push(pad(['']));

  rows.push(pad(['■ 月別（新しい順・全期間）']));
  const monthHeadRow = rows.length + 1;
  rows.push(pad(['年月', '営業日数', '有効日数', 'POS総売上', '税抜売上(明細)', '理論原価', '理論原価率', '前月差', '日平均売上']));
  const monthStartRow = rows.length + 1;
  const monthsDesc = months.slice().reverse();
  monthsDesc.forEach(m => {
    // 'yyyy/MM' のまま書くとスプレッドシートが日付と解釈して「2026/07/01」になってしまう
    rows.push(pad([m.ym.replace('/', '年') + '月', m.days, m.validDays, Math.round(m.total), Math.round(m.vNetSales),
                   Math.round(m.vCost), m.rate, m.diff, Math.round(m.dayAvg)]));
  });
  rows.push(pad(['']));

  rows.push(pad(['■ 日別（新しい順・直近' + DAILY_DAYS + '日）']));
  const dailyHeadRow = rows.length + 1;
  rows.push(pad(['日付', '曜日', 'POS総売上', '税抜売上(明細)', '理論原価', '理論原価率', '7日移動平均', 'マスタ価格一致率', '判定']));
  const dailyStartRow = rows.length + 1;
  days.slice().reverse().forEach(x => {
    rows.push(pad([x.ds, x.dow, Math.round(x.total), Math.round(x.netSales), Math.round(x.cost),
                   x.ng ? '' : x.rate, x.ng ? '' : (x.ma7 || ''), x.fit, x.ng || 'OK']));
  });
  const dailyCount = days.length;

  rows.push(pad(['']));
  rows.push(pad(['■ 商品別（直近' + ITEM_DAYS + '日／原価インパクトの大きい順 上位' + ITEM_TOP + '）']));
  const itemHeadRow = rows.length + 1;
  rows.push(pad(['商品コード', '商品名', '出食数', '税抜価格', '1食あたり原価', '原価率', '原価合計', '税抜売上合計', '売上構成比']));
  const itemStartRow = rows.length + 1;
  const itemList = Object.keys(itemAgg).map(code => {
    const v = itemAgg[code];
    return { code: code, name: v.name, qty: v.qty, net: v.net, cost: v.cost,
             costSum: v.cost * v.qty, saleSum: v.net * v.qty,
             rate: v.net > 0 ? v.cost / v.net : 0 };
  }).filter(x => x.qty > 0);
  const saleTotal = itemList.reduce((s, x) => s + x.saleSum, 0);
  itemList.sort((a, b) => b.costSum - a.costSum);
  const itemShown = itemList.slice(0, ITEM_TOP);
  itemShown.forEach(x => {
    rows.push(pad([x.code, x.name, x.qty, Math.round(x.net), Math.round(x.cost * 10) / 10, x.rate,
                   Math.round(x.costSum), Math.round(x.saleSum), saleTotal > 0 ? x.saleSum / saleTotal : 0]));
  });

  // ── 食材マスタの登録ミス検出 ───────────────────────────────────────
  //   払出単価は「仕入規格量(I列)あたりの価格」。仕入納品量(J列)は1回の納品量で、
  //   発注ロット＝J÷I として使う。ここが崩れていると原価も発注量も狂う。
  //   例）2kg袋5,598円の食材で規格量を1,000gにすると単価が2倍になる（＝原価率が跳ね上がる）
  const ingAlerts = [];
  ingRows.forEach(g => {
    if (g.cat === '発注終了') return;
    const msg = [];
    if (g.spec <= 0) msg.push('仕入規格量が未入力');
    if (g.deliv <= 0) msg.push('仕入納品量が未入力');
    if (g.spec > 0 && g.deliv > 0) {
      const lot = g.deliv / g.spec;
      // 不定貫（箱ごとに重量が変わる食材）は端数が出て当たり前なので、
      // ロットが1未満のときだけ異常とみなす（納品量が規格量を下回るのは筋が通らない）
      if (lot < 1) msg.push('仕入納品量(' + g.deliv + ')が仕入規格量(' + g.spec + ')より小さい');
      else if (Math.abs(lot - Math.round(lot)) > 0.01 && lot < 5) {
        msg.push('発注ロットが ' + (Math.round(lot * 100) / 100) + ' と半端（納品量÷規格量が整数でない）');
      }
    }
    if (msg.length > 0) ingAlerts.push([g.code, g.name, g.price, g.spec, g.deliv, msg.join(' / ')]);
  });
  if (ingAlerts.length > 0) {
    rows.push(pad(['']));
    rows.push(pad(['■ 食材マスタの登録ミスの疑い']));
    rows.push(pad(['食材コード', '食材名', '払出単価', '仕入規格量', '仕入納品量', '内容']));
    ingAlerts.slice(0, 20).forEach(a => rows.push(pad(a)));
  }

  // ── レシピ入力ミスの自動検出 ───────────────────────────────────────
  //    同じ食材なのに「提供量 ÷ 単位使用量」の比率が他の商品と違う行を洗い出す。
  //    例：米は 炊き上がり = 生米×2 が正しいが、炊き上がりの量を単位使用量の欄に
  //    入れてしまうと比率が1になり、原価が倍に膨らむ。
  const byIng = {};
  recipeRows.forEach(r => {
    if (!byIng[r.ing]) byIng[r.ing] = {};
    if (!byIng[r.ing][r.ratio]) byIng[r.ing][r.ratio] = [];
    byIng[r.ing][r.ratio].push(r);
  });
  const recipeAlerts = [];
  Object.keys(byIng).forEach(ing => {
    const groups = byIng[ing], ratios = Object.keys(groups);
    if (ratios.length < 2) return;
    const total = ratios.reduce((s, k) => s + groups[k].length, 0);
    if (total < 5) return;                                  // サンプルが少ない食材は判定しない
    let mainRatio = ratios[0];
    ratios.forEach(k => { if (groups[k].length > groups[mainRatio].length) mainRatio = k; });
    ratios.forEach(k => {
      if (k === mainRatio) return;
      if (groups[k].length / total >= 0.1) return;           // 少数派（1割未満）だけを異常候補にする
      // 比率の差が小さいものは正常なばらつき（例：110g→100gの焼き縮み）なので除く。
      // 入力ミスは「生と炊き上がりを取り違える」等で2倍・半分になるため差は大きい。
      if (Math.abs(Number(k) / Number(mainRatio) - 1) < 0.25) return;
      groups[k].forEach(r => recipeAlerts.push([
        r.prod, r.prodName, r.ingName, r.use, r.serv,
        Number(k), Number(mainRatio), Math.round(r.serv / Number(mainRatio) * 10) / 10
      ]));
    });
  });
  if (recipeAlerts.length > 0) {
    rows.push(pad(['']));
    rows.push(pad(['■ レシピ入力ミスの疑い（同じ食材なのに他商品と使用量の比率が違う）']));
    rows.push(pad(['商品コード', '商品名', '食材', '単位使用量', '提供量', 'この行の比率', '他商品の比率', '想定される正しい使用量']));
    recipeAlerts.slice(0, 30).forEach(a => rows.push(pad(a)));
  }

  // ── 原価率が異常な商品（レシピ入力ミス・価格未更新の疑い）─────────
  //    判定は主力商品（税抜500円以上）だけに絞る。ドリンク・トッピングなどの
  //    低単価商品はもともと原価率が偏り、拾うと警告だらけになって役に立たないため。
  //    上限50%：主力商品は通常30〜45%。50%超は単価やレシピの登録ミスを疑う水準
  //    （実例：カルビの払出単価が2kg価格で入っていて原価率62%になっていた）。
  const RATE_LOW = 0.15, RATE_HIGH = 0.50, MIN_PRICE = 500;
  const rateAlerts = itemList.filter(x => x.net >= MIN_PRICE && (x.rate < RATE_LOW || x.rate > RATE_HIGH))
                             .sort((a, b) => b.rate - a.rate);
  if (rateAlerts.length > 0) {
    rows.push(pad(['']));
    rows.push(pad(['■ 原価率が異常な商品（直近' + ITEM_DAYS + '日に売れた税抜' + MIN_PRICE + '円以上／'
      + (RATE_LOW * 100) + '%未満 または ' + (RATE_HIGH * 100) + '%超）']));
    rows.push(pad(['商品コード', '商品名', '出食数', '税抜価格', '1食あたり原価', '原価率', '疑われる原因']));
    rateAlerts.slice(0, 30).forEach(x => rows.push(pad([
      x.code, x.name, x.qty, Math.round(x.net), Math.round(x.cost * 10) / 10, x.rate,
      x.rate < RATE_LOW ? '原価が少なすぎる（レシピの使用量もれ／食材未登録）'
                        : '原価が高すぎる（払出単価が納品1回ぶんの価格になっていないか／使用量の入れ間違い）'
    ])));
  }

  // 原価が一部欠けている商品（食材マスタ未登録などで金額が #N/A になっている）
  const badItems = Object.keys(itemMap).filter(c => itemMap[c].badIng.length > 0);
  if (badItems.length > 0) {
    rows.push(pad(['']));
    rows.push(pad(['■ 原価が一部計上できていない商品（食材マスタに無い食材があります）']));
    rows.push(pad(['商品コード', '商品名', '原価が取れない食材']));
    badItems.slice(0, 20).forEach(c => rows.push(pad([c, itemMap[c].name, itemMap[c].badIng.join(' / ')])));
  }

  // 商品マスタに無いメニュー（原価が計上できていない分）
  const unknownList = Object.keys(unknownMap).map(c => ({ code: c, name: unknownMap[c].name, qty: unknownMap[c].qty }))
                        .sort((a, b) => b.qty - a.qty);
  if (unknownList.length > 0) {
    rows.push(pad(['']));
    rows.push(pad(['■ 商品マスタ未登録のメニュー（原価が計上されていません）']));
    rows.push(pad(['メニューCD', 'メニュー名', '出食数']));
    unknownList.slice(0, 20).forEach(x => rows.push(pad([x.code, x.name, x.qty])));
  }

  sh.getRange(1, 1, rows.length, W).setValues(rows);

  // ── 書式 ───────────────────────────────────────────────────────────
  sh.getRange(1, 1).setFontSize(14).setFontWeight('bold');
  sh.getRange(3, 1, 1, 6).setFontWeight('bold');
  [2, 4, 6].forEach(c => sh.getRange(3, c).setNumberFormat('0.0%'));
  sh.getRange(monthHeadRow, 1, 1, W).setFontWeight('bold').setBackground('#d9ead3');
  sh.getRange(dailyHeadRow, 1, 1, W).setFontWeight('bold').setBackground('#e8eaed');
  sh.getRange(itemHeadRow,  1, 1, W).setFontWeight('bold').setBackground('#e8eaed');
  if (monthsDesc.length > 0) {
    sh.getRange(monthStartRow, 2, monthsDesc.length, 5).setNumberFormat('#,##0');
    sh.getRange(monthStartRow, 7, monthsDesc.length, 1).setNumberFormat('0.0%');
    sh.getRange(monthStartRow, 8, monthsDesc.length, 1).setNumberFormat('+0.0%;-0.0%;0.0%');
    sh.getRange(monthStartRow, 9, monthsDesc.length, 1).setNumberFormat('#,##0');
    sh.getRange(monthStartRow, 1, monthsDesc.length, W).setFontWeight('bold');
  }
  if (dailyCount > 0) {
    sh.getRange(dailyStartRow, 3, dailyCount, 3).setNumberFormat('#,##0');
    sh.getRange(dailyStartRow, 6, dailyCount, 3).setNumberFormat('0.0%');
  }
  if (itemShown.length > 0) {
    sh.getRange(itemStartRow, 3, itemShown.length, 3).setNumberFormat('#,##0.0');
    sh.getRange(itemStartRow, 6, itemShown.length, 1).setNumberFormat('0.0%');
    sh.getRange(itemStartRow, 7, itemShown.length, 2).setNumberFormat('#,##0');
    sh.getRange(itemStartRow, 9, itemShown.length, 1).setNumberFormat('0.0%');
  }
  sh.setFrozenRows(3);   // サマリー行だけ固定（月別・日別の表はスクロールで見る）
  for (let c = 1; c <= W; c++) sh.setColumnWidth(c, c === 2 ? 160 : 110);

  console.log('原価分析を更新しました（月別' + months.length + 'か月／日別' + dailyCount + '日／商品' + itemList.length + '種）'
            + ' 直近7日原価率=' + Math.round(avgOf(7, false) * 1000) / 10 + '%'
            + ' 直近30日原価率=' + Math.round(avgOf(30, false) * 1000) / 10 + '%');
}