Sub csv出力_1シート_main()
    'csv出力する該当のシート名をcsv_sh_nameに入力
    '例 csv_sh_name = "商品マスター"
    csv_sh_name = ""
    
    call Microsoft_Scripting_Runtime    '参照設定の追加
    
    '再計算
    Worksheets(csv_sh_name).Calculate  
    
    Application.Calculation = xlCalculationManual     '手動計算
    Application.ScreenUpdating = False    '画面更新停止

    Dim sheet1 As Worksheet
    Set sheet1 = Worksheets(csv_sh_name)
    Call csv出力処理(sheet1)

    Application.Calculation = xlCalculationAutomatic  '自動計算
    Application.ScreenUpdating = True      '画面更新再開
End Sub

Sub csv出力処理(ByVal sht As Worksheet)
    Dim varFile As Variant
    Dim SaveDir As String
    
    SaveDir = ThisWorkbook.Path & "\csv"
    
    ' フォルダがなければ作成する
    If Dir(SaveDir, vbDirectory) = "" Then
        MkDir SaveDir
    End If

    varFile = SaveDir & "\" & Left(ThisWorkbook.Name, Len(ThisWorkbook.Name) - 5) & ".csv"
    
    '↓これ見直し必要かも。emptyとかエラーの場合とか。そもそも要らないかも。
    If varFile = False Then
        Exit Sub
    End If
    
    sht.Select
    sht.Copy
  
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs Filename:=varFile, FileFormat:=xlCSV, CreateBackup:=False
    ActiveWindow.Close
    
    Application.DisplayAlerts = True
    'MsgBox ("CSV出力しました。" & vbLf & vbLf & varFile)
    Debug.Print Now & " CSV出力しました。" & varFile
End Sub

Sub Microsoft_Scripting_Runtime()

    On Error GoTo Err
    
    'Microsoft Scripting RuntimeのGUID
    Const MSR_GUID = "{420B2830-E718-11CF-893D-00A0C9054228}"
    '参照設定を追加
    Application.VBE.ActiveVBProject.References.AddFromGuid MSR_GUID, 1, 0
    
'    MsgBox "参照設定を追加しました！"
    
    Exit Sub
    
Err:
'    MsgBox "エラーが発生しました！" & vbCrLf & Err.Description
 
End Sub