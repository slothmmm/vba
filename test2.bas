Sub アイテムコピー改()
'」」」」」」」」」」」」」」」」」」」」」」」」」」」」」」」」」」」」」」」」」
'」」」」」」」」」」」」」」」」」」」」」」」」」」」」」」」」」」」」」」」」」
'
'2020.12.04
'Call 保護.全保護解除

   
'宣言～設定 」」」」」
    Dim d As Date, n As Integer, クリア As Integer, 翌月 As String, _
    ws As Worksheet, flag As Boolean, _
    入力用日付 As Date, コピー先 As String
   
    If Worksheets("受注入力").Range("O1") = "☆" Then
      GoTo NN1
        End If
    
'原料展開対応件数確認
    If Worksheets("受注入力").Range("K1") = "" Then
    Else
    MsgBox "システム部担当者へ連絡してください。"
    Exit Sub
    End If
    
    '2021.02.26 フィルター解除
    Worksheets("出荷数アイテム").Activate
        With ActiveSheet
            .Range("A3").Select
            If .FilterMode Then .ShowAllData
        End With
    
NN1:
'データコピー 」」」」」
    Sheets("ロゴ").Visible = True
    Sheets("ロゴ").Select
    Application.Wait [Now() + "0:00:01"]
    Application.ScreenUpdating = False
    Sheets("ロゴ").Visible = False
    Sheets("出荷数アイテム").Select
    n = WorksheetFunction.CountA(Range("A:A")) - 3
    d = Worksheets("受注入力").Cells(1, 1)
    翌月 = Format(d + 50, "m月")
    コピー先 = Format(d, "m月")
    入力用日付 = DateSerial(Year(d + 50), Month(d + 50), 1)
    
    Application.DisplayAlerts = False
    Range(Cells(4, 4), Cells(4 + n, 5)).Select
    Selection.Copy
    Workbooks.Open filename:="\\Afnewt320-kyoyu\社内共有\AFSKS\生産管理\" & _
                   Format(d, "yyyy年") & "アイテム別出荷数量.xlsm", UpdateLinks:=3, IgnoreReadOnlyRecommended:=True
    クリア = WorksheetFunction.CountA(Range("A:A")) + 1
    For Each ws In Worksheets
    If ws.Name = 翌月 Then flag = True
    Next ws
    If flag = True Then
    Else
    Sheets(コピー先).Copy Before:=Sheets(コピー先)
    ActiveSheet.Name = 翌月
    Range(Cells(4, 5), Cells(クリア, 66)).Select
    Selection.ClearContents
    Range("A1").Select
    ActiveCell.FormulaR1C1 = 入力用日付
    End If
    Workbooks("生産表.xlsm").Activate

 ActiveWorkbook.UpdateLink Name:="\\Afnewt320-kyoyu\社内共有\AFSKS\生産管理\販売商品マスター.xlsm", _
                              Type:=xlExcelLinks

    Range(Cells(4, 4), Cells(4 + n, 5)).Select
    Selection.Copy
    Workbooks(Format(d, "yyyy年") & "アイテム別出荷数量.xlsm").Activate
    Sheets(コピー先).Select
    Range(Cells(4, Day(d) * 2 + 3), Cells(4 + n, Day(d) * 2 + 4)).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    ActiveWorkbook.Close SaveChanges:=True
    Windows("生産表.xlsm").Activate
    Sheets("受注入力").Select
'自動処理時メッセージSKIP」」」」」」」」」」」」」」」」」」」」」」」」」」」」」
    If Range("N1") = 1 Or Range("O1") = "☆" Then
    GoTo skip
    End If
    Application.ScreenUpdating = True
    Application.Wait [Now() + "0:00:01"]
    
    '2020.12.04
    'Call 保護.複数保護
    
    MsgBox "原料データ更新完了"
skip:
    Application.DisplayAlerts = True
'」」」」」」」」」」」」」」」」」」」」」」」」」」」」」」」」」」」」」」」」」
'」」」」」」」」」」」」」」」」」」」」」」」」」」」」」」」」」」」」」」」」」
End Sub