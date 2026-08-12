$srcPAPath      = "C:\Users\matsu\OneDrive\Desktop\AI作業場\時給改定資料\ペッパーランチPA評価表  (1).xlsx"
$srcCollectPath = "C:\Users\matsu\OneDrive\Desktop\AI作業場\時給改定資料\時給変更集計原本 (1).xlsx"
$destPath       = "C:\Users\matsu\OneDrive\Desktop\AI作業場\時給改定資料\PA評価・時給変更管理.xlsm"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$srcPA      = $excel.Workbooks.Open($srcPAPath)
$srcCollect = $excel.Workbooks.Open($srcCollectPath)

$newWb = $excel.Workbooks.Add()

# デフォルトシートを全削除（後で追加するため一時的に1枚残す）
while ($newWb.Worksheets.Count -gt 1) {
    $newWb.Worksheets.Item($newWb.Worksheets.Count).Delete()
}
$newWb.Worksheets.Item(1).Name = "_tmp"

# ① PA評価表シートをコピー → 「評価入力」
$srcPA.Worksheets.Item("チェックシート").Copy($newWb.Worksheets.Item("_tmp"))
$wsEval = $newWb.Worksheets.Item(1)
$wsEval.Name = "評価入力"

# 行65以降（集計表）を別シートとして残す前に確認
# 行64まで（評価項目）+ 行65以降（クラスター集計表）を分離せずそのままコピー

# ② 時給変更集計シートをコピー
$srcCollect.Worksheets.Item("入力フォーム").Copy([System.Reflection.Missing]::Value, $newWb.Worksheets.Item($newWb.Worksheets.Count))
$wsCollect = $newWb.Worksheets.Item($newWb.Worksheets.Count)
$wsCollect.Name = "時給変更集計"

# ③ 評価履歴シートを追加
$wsHistory = $newWb.Worksheets.Add([System.Reflection.Missing]::Value, $wsCollect)
$wsHistory.Name = "評価履歴"

# ④ 評価詳細履歴シートを追加
$wsDetail = $newWb.Worksheets.Add([System.Reflection.Missing]::Value, $wsHistory)
$wsDetail.Name = "評価詳細履歴"

# ⑤ クラスター基準シートを追加
$wsCluster = $newWb.Worksheets.Add([System.Reflection.Missing]::Value, $wsDetail)
$wsCluster.Name = "クラスター基準"

# _tmpシートを削除
$newWb.Worksheets.Item("_tmp").Delete()

# シート順を確認
foreach ($ws in $newWb.Worksheets) {
    Write-Host "シート: $($ws.Name)"
}

# xlsm形式で保存 (52 = xlOpenXMLWorkbookMacroEnabled)
if (Test-Path $destPath) { Remove-Item $destPath -Force }
$newWb.SaveAs($destPath, 52)
Write-Host "保存完了: $destPath"

$srcPA.Close($false)
$srcCollect.Close($false)
$newWb.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
