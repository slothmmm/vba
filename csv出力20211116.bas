Sub csv出力_main()
Dim sheet1 As Worksheet
    Worksheets("Sheet1").Calculate    '再計算
    Application.Calculation = xlCalculationManual     '手動計算
    Application.ScreenUpdating = False    '画面更新停止
    Set sheet1 = Worksheets("小分け品")
    Call csv出力処理(sheet1, Range("A1"))
    Application.Calculation = xlCalculationAutomatic  '自動計算
    Application.ScreenUpdating = True      '画面更新再開
End Sub

Sub csv出力処理(ByVal sht As Worksheet, Optional ByVal rngStart As Range = Nothing)
    Dim varFile As Variant
    Dim SaveDir As String
    
    SaveDir = ThisWorkbook.Path & "\csv"
    
    ' フォルダがなければ作成する
    If Dir(SaveDir, vbDirectory) = "" Then
        MkDir SaveDir
    End If

    varFile = SaveDir & "\" & Left(ThisWorkbook.Name, Len(ThisWorkbook.Name) - 5) & ".csv"
    
    If varFile = False Then
        Exit Sub
    End If
    
    sht.Select
    sht.Copy



    '不要な先頭の行列を削除します。
    If Not rngStart Is Nothing Then
        If rngStart.Row > 1 Then
            Range(Rows(1), Rows(rngStart.Row - 1)).Delete
        End If
        If rngStart.Column > 1 Then
            Range(Columns(1), Columns(rngStart.Column - 1)).Delete
        End If
    End If
  
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs Filename:=varFile, FileFormat:=xlCSV, CreateBackup:=False
    ActiveWindow.Close
    
    Application.DisplayAlerts = True
    MsgBox ("CSV出力しました。" & vbLf & vbLf & varFile)
End Sub