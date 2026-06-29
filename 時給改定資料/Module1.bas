Attribute VB_Name = "Module1"
Option Explicit

Function GetRosterPath() As String
    GetRosterPath = ThisWorkbook.Path & "\アルバイト名簿 (2).xlsx"
End Function

Sub LoadStaffInfo(staffName As String)
    Dim rosterPath As String
    rosterPath = GetRosterPath()
    If Dir(rosterPath) = "" Then
        MsgBox "アルバイト名簿ファイルが見つかりません。" & Chr(10) & rosterPath
        Exit Sub
    End If
    Dim rosterWb As Workbook
    Dim wasOpen As Boolean
    wasOpen = False
    Dim wbk As Workbook
    For Each wbk In Workbooks
        If wbk.FullName = rosterPath Then
            wasOpen = True
            Set rosterWb = wbk
            Exit For
        End If
    Next wbk
    If Not wasOpen Then Set rosterWb = Workbooks.Open(rosterPath, ReadOnly:=True)
    Dim masterWs As Worksheet
    Set masterWs = rosterWb.Worksheets("スタッフマスタ")
    Dim found As Range
    Set found = masterWs.Columns(4).Find(staffName, LookIn:=xlValues, LookAt:=xlWhole)
    Dim evalWs As Worksheet
    Set evalWs = ThisWorkbook.Worksheets("評価入力")
    Application.EnableEvents = False
    If Not found Is Nothing Then
        evalWs.Cells(3, 7).Value  = found.Offset(0, -1).Value
        evalWs.Cells(3, 9).Value  = found.Offset(0, 23).Value
        evalWs.Cells(3, 11).Value = found.Offset(0, 22).Value
        evalWs.Cells(3, 14).Value = found.Offset(0, 8).Value
        evalWs.Cells(3, 22).Value = found.Offset(0, 12).Value
    Else
        evalWs.Cells(3, 7).ClearContents
        evalWs.Cells(3, 9).ClearContents
        evalWs.Cells(3, 11).ClearContents
        evalWs.Cells(3, 14).ClearContents
        evalWs.Cells(3, 22).ClearContents
    End If
    Application.EnableEvents = True
    If Not wasOpen Then rosterWb.Close SaveChanges:=False
End Sub

Sub LoadPreviousScores(staffName As String)
    Dim detailWs As Worksheet
    Set detailWs = ThisWorkbook.Worksheets("評価詳細履歴")
    Dim evalWs As Worksheet
    Set evalWs = ThisWorkbook.Worksheets("評価入力")
    Dim lastRow As Long
    lastRow = detailWs.Cells(detailWs.Rows.Count, "A").End(xlUp).Row
    evalWs.Range("N6:N65").ClearContents
    evalWs.Range("AB6:AB45").ClearContents
    evalWs.Cells(4, 1).Value = "直近評価：なし"
    If lastRow < 2 Then Exit Sub
    Dim foundRow As Long
    foundRow = 0
    Dim i As Long
    For i = lastRow To 2 Step -1
        If detailWs.Cells(i, 1).Value = staffName Then
            foundRow = i
            Exit For
        End If
    Next i
    If foundRow = 0 Then Exit Sub
    evalWs.Cells(4, 1).Value = "直近評価：" & detailWs.Cells(foundRow, 2).Value & _
        "  合計" & detailWs.Cells(foundRow, 3).Value & "点"
    For i = 0 To 59
        evalWs.Cells(6 + i, 14).Value = detailWs.Cells(foundRow, 4 + i).Value
    Next i
    For i = 0 To 39
        evalWs.Cells(6 + i, 28).Value = detailWs.Cells(foundRow, 64 + i).Value
    Next i
End Sub

Sub BuildNameDropdown()
    Dim rosterPath As String
    rosterPath = GetRosterPath()
    If Dir(rosterPath) = "" Then Exit Sub
    Dim rosterWb As Workbook
    Dim wasOpen As Boolean
    wasOpen = False
    Dim wbk As Workbook
    For Each wbk In Workbooks
        If wbk.FullName = rosterPath Then
            wasOpen = True
            Set rosterWb = wbk
            Exit For
        End If
    Next wbk
    If Not wasOpen Then Set rosterWb = Workbooks.Open(rosterPath, ReadOnly:=True)
    Dim nameWs As Worksheet
    Set nameWs = rosterWb.Worksheets("名簿")
    Dim names() As String
    Dim cnt As Long
    cnt = 0
    Dim r As Long
    Dim lastR As Long
    lastR = nameWs.Cells(nameWs.Rows.Count, 4).End(xlUp).Row
    For r = 3 To lastR
        If nameWs.Cells(r, 4).Value <> "" Then
            cnt = cnt + 1
            ReDim Preserve names(0 To cnt - 1)
            names(cnt - 1) = nameWs.Cells(r, 4).Value
        End If
    Next r
    If Not wasOpen Then rosterWb.Close SaveChanges:=False
    If cnt = 0 Then Exit Sub
    Dim evalWs As Worksheet
    Set evalWs = ThisWorkbook.Worksheets("評価入力")
    With evalWs.Cells(3, 1).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:=Join(names, ",")
        .InCellDropdown = True
        .ShowError = False
    End With
End Sub