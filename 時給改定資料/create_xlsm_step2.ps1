$destPath = "C:\Users\matsu\OneDrive\Desktop\AI作業場\時給改定資料\PA評価・時給変更管理.xlsm"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Open($destPath)

$wsEval    = $wb.Worksheets.Item("評価入力")
$wsCluster = $wb.Worksheets.Item("クラスター基準")
$wsHist    = $wb.Worksheets.Item("評価履歴")
$wsDetail  = $wb.Worksheets.Item("評価詳細履歴")

$cBlue   = 4485828
$cWhite  = 16777215
$cGreen  = 13033652
$cGray   = 14277081
$cGray2  = 15921906
$cYellow = 16777113
$cLBlue  = 13428223
$cGreen2 = 9621584
$cAmber  = 16760832
$cDGray  = 5855577

# A. クラスター基準シート
$wsCluster.Cells(1,1).Value = "点数下限"
$wsCluster.Cells(1,2).Value = "クラスター名"
$wsCluster.Cells(1,3).Value = "時給"
$wsCluster.Cells(1,4).Value = "説明"
for ($c=1; $c -le 4; $c++) {
    $wsCluster.Cells(1,$c).Font.Bold = $true
    $wsCluster.Cells(1,$c).Interior.Color = $cBlue
    $wsCluster.Cells(1,$c).Font.Color = $cWhite
}
$r = 2
foreach ($d in @("0|ルーキー|1100|入社3ヶ月以下","15|1クルー|1110|15〜24点","25|2クルー|1120|25〜34点","35|3クルー|1130|35〜49点","50|1トレーナー|1140|50〜59点","60|2トレーナー|1150|60〜69点","70|3トレーナー|1160|70〜89点","90|マネージャー|1200|90点以上")) {
    $p = $d.Split("|")
    $wsCluster.Cells($r,1).Value = [int]$p[0]
    $wsCluster.Cells($r,2).Value = [string]$p[1]
    $wsCluster.Cells($r,3).Value = [int]$p[2]
    $wsCluster.Cells($r,4).Value = [string]$p[3]
    $r++
}
$wsCluster.Columns(1).ColumnWidth = 10
$wsCluster.Columns(2).ColumnWidth = 14
$wsCluster.Columns(3).ColumnWidth = 8
$wsCluster.Columns(4).ColumnWidth = 18
Write-Host "A: クラスター基準完了"

# B. 評価入力シート改修
$wsEval.Rows(4).Insert() | Out-Null
$wsEval.Cells(4,1).Value = "直近評価：（氏名選択時に自動表示）"
$wsEval.Cells(4,1).Font.Italic = $true
$wsEval.Cells(4,1).Font.Color = $cDGray

$wsEval.Cells(5,13).Value = "今"
$wsEval.Cells(5,13).Font.Bold = $true
$wsEval.Cells(5,13).Interior.Color = $cGreen
$wsEval.Cells(5,14).Value = "前"
$wsEval.Cells(5,14).Font.Bold = $true
$wsEval.Cells(5,14).Interior.Color = $cGray
$wsEval.Cells(5,27).Value = "今"
$wsEval.Cells(5,27).Font.Bold = $true
$wsEval.Cells(5,27).Interior.Color = $cGreen
$wsEval.Cells(5,28).Value = "前"
$wsEval.Cells(5,28).Font.Bold = $true
$wsEval.Cells(5,28).Interior.Color = $cGray

$wsEval.Range("N6:N70").Interior.Color   = $cGray2
$wsEval.Range("AB6:AB50").Interior.Color = $cGray2
$wsEval.Columns(14).ColumnWidth = 3
$wsEval.Columns(28).ColumnWidth = 3

$wsEval.Cells(3,18).Formula = '=IFERROR(VLOOKUP(SUM(M6:M66)+SUM(AA6:AA45),クラスター基準!$A:$C,2,TRUE),"")'
$wsEval.Cells(3,25).Formula = '=IFERROR(VLOOKUP(SUM(M6:M66)+SUM(AA6:AA45),クラスター基準!$A:$C,3,TRUE),"")'

$wsEval.Cells(3,1).Interior.Color  = $cYellow
$wsEval.Cells(3,7).Interior.Color  = $cLBlue
$wsEval.Cells(3,9).Interior.Color  = $cLBlue
$wsEval.Cells(3,11).Interior.Color = $cLBlue
$wsEval.Cells(3,14).Interior.Color = $cLBlue
$wsEval.Cells(3,18).Interior.Color = $cLBlue
$wsEval.Cells(3,22).Interior.Color = $cLBlue
$wsEval.Cells(3,25).Interior.Color = $cLBlue
Write-Host "B: 評価入力シート完了"

# C. 評価履歴シートヘッダー
$hh = @("スタッフ名","評価年月","規律点(/17)","マインド点(/14)","共通スキル点(/29)","ポジション点(/40)","合計点(/100)","旧クラスター","新クラスター","旧時給","新時給","承認状況","保存日時")
for ($c=0; $c -lt $hh.Count; $c++) {
    $wsHist.Cells(1,$c+1).Value = [string]$hh[$c]
    $wsHist.Cells(1,$c+1).Font.Bold = $true
    $wsHist.Cells(1,$c+1).Interior.Color = $cBlue
    $wsHist.Cells(1,$c+1).Font.Color = $cWhite
}
$wsHist.Rows(1).RowHeight = 20
Write-Host "C: 評価履歴シート完了"

# D. 評価詳細履歴シートヘッダー
$bh = @("スタッフ名","評価年月","合計点")
for ($c=0; $c -lt $bh.Count; $c++) {
    $wsDetail.Cells(1,$c+1).Value = [string]$bh[$c]
    $wsDetail.Cells(1,$c+1).Font.Bold = $true
    $wsDetail.Cells(1,$c+1).Interior.Color = $cBlue
    $wsDetail.Cells(1,$c+1).Font.Color = $cWhite
}
$lL = @("規律_挨拶1","規律_挨拶2","規律_出退勤1","規律_出退勤2","規律_出退勤3","規律_出退勤4","規律_外見1","規律_外見2","規律_シフト1","規律_シフト2","規律_シフト3","規律_シフト4","規律_シフト5","規律_環境1","規律_環境2","規律_理念1","規律_理念2","マ_協調1","マ_協調2","マ_協調3","マ_協調4","マ_感受1","マ_感受2","マ_感受3","マ_感受4","マ_使命1","マ_使命2","マ_積極1","マ_積極2","マ_積極3","マ_責任1","SK_笑顔1","SK_笑顔2","SK_笑顔3","SK_声1","SK_声2","SK_声3","SK_声4","SK_掃除1","SK_掃除2","SK_掃除3","SK_掃除4","SK_商品知識","SK_報告1","SK_報告2","SK_OP1","SK_OP2","SK_OP3","SK_OP4","SK_OP5","SK_OP6","SK_CL1","SK_CL2","SK_CL3","SK_CL4","SK_CL5","SK_CL6","SK_クレーム1","SK_クレーム2","SK_クレーム3")
$rL = @("DS_接客1","DS_接客2","DS_接客3","DS_接客4","DS_レジ1","DS_レジ2","DS_レジ3","DS_レジ4","DS_レジ5","DS_DG1","DS_DG2","DS_DG3","DS_DG4","DS_DG5","DS_DG6","WS_洗1","WS_洗2","WS_洗3","WS_洗4","WS_洗5","WS_洗6","KT_基本1","KT_基本2","KT_基本3","KT_食材1","KT_食材2","KT_MP1","KT_MP2","KT_MP3","KT_MP4","KT_MP5","KT_MP6","KT_MP7","KT_MP8","KT_MP9","KT_仕込み1","KT_仕込み2","KT_仕込み3","KT_仕込み4","KT_仕込み5")
for ($c=0; $c -lt $lL.Count; $c++) {
    $wsDetail.Cells(1,4+$c).Value = [string]$lL[$c]
    $wsDetail.Cells(1,4+$c).Font.Bold = $true
    $wsDetail.Cells(1,4+$c).Interior.Color = $cGreen2
}
for ($c=0; $c -lt $rL.Count; $c++) {
    $wsDetail.Cells(1,4+$lL.Count+$c).Value = [string]$rL[$c]
    $wsDetail.Cells(1,4+$lL.Count+$c).Font.Bold = $true
    $wsDetail.Cells(1,4+$lL.Count+$c).Interior.Color = $cAmber
}
Write-Host "D: 評価詳細履歴シート完了"

$wb.Save()
Write-Host "Step2保存完了"
$wb.Close($true)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
