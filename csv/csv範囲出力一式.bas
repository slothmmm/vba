
Sub csvメイン()
    'csvファイル名を決める
    Dim csvFileName As String
    csvFileName = csvファイル名

    '直下csvディレクトリの作成
    ディレクトリ作成1

    '書き出す範囲指定
    csv範囲パス_指定

    '↑の選択範囲を書き出す
    csv書き出し (csvFileName)

End Sub

Function csvファイル名() As Variant
    'csvファイル名_現在時刻
    hiduke = Format(Year(Now), "0000") & "年" & Format(Month(Now), "00") & "月" & Format(Day(Now), "00") & "日" & Format(Hour(Now), "00") & "時" & Format(Minute(Now), "00") & "分" & Format(Second(Now), "00") & "秒"

    'csvファイル名
    Dim csvFileName As String
    csvFileName = ActiveWorkbook.Path & "\csv\" & ThisWorkbook.Name & hiduke & ".csv"

    csvファイル名 = csvFileName
End Function

Sub ディレクトリ作成1()
    Dim root As String
'    Dim yyyy As String
'    Dim mm As String
'    Dim dd As String
'    'F1出荷日
'    Dim ship_date As Date
'    ship_date = Range("F1")

    ' root = ActiveWorkbook.Path & "\csv"
    root = ActiveWorkbook.Path
'    yyyy = Format(Year(ship_date), "0000年")
'    mm = Format(Month(ship_date), "00月")
    ' dd = Format(Day(ship_date), "00日")

'    'F2 出荷先
'    Dim Customer_name As String
'    Customer_name = Range("F2")

    Dim rtn As Long
    '下位ディレクトリも出来る
    rtn = ディレクトリ作成2(root, "csv")
    'rtn = ディレクトリ作成2(root, "csv", Customer_name, yyyy, mm)
'    Select Case rtn
'        Case 0
'            MsgBox "フォルダを作成しました。"
'        Case 1
'            MsgBox "フォルダは存在します。"
'        Case Else
'            MsgBox "フォルダの作成に失敗しました。"
'    End Select
End Sub

Function ディレクトリ作成2(ParamArray arg()) As Long
    On Error GoTo ErrExit
    If Dir(Join(arg, "\"), vbDirectory) <> "" Then
        CreateDirectory = 1
        Exit Function
    End If

    Dim ary As Variant
    Dim i As Long
    For i = LBound(arg) To UBound(arg)
        ary = arg
        ReDim Preserve ary(i)
        If Dir(Join(ary, "\"), vbDirectory) = "" Then
            'ディレクトリ作成mkdir
            MkDir Join(ary, "\")
        End If
    Next

    CreateDirectory = 0
    Exit Function

ErrExit:
    CreateDirectory = 9
End Function

Sub csv範囲パス_指定()
    'この関数で範囲指定、パスを設定
    Dim range_select As Range
    Set select_rng = Range("F7:J11")
    select_rng.Select
End Sub

Function csv書き出し(csvFileName As Variant)

    Application.DisplayAlerts = False

    '■現在選択しているセル情報をrngに格納
    'select_rng.Select
    Set select_rng = Selection
    'Set select_range = Range("C8:F11")

    '■新規ブック作成→select_rngをA1にコピー→CSV保存→CSV閉じる
    Workbooks.Add
    select_rng.Copy
    ActiveSheet.Range("A1").PasteSpecial _
                                 Paste:=xlPasteValues, _
                                 Operation:=xlNone, _
                                 SkipBlanks:=False, _
                                 Transpose:=False

    'select_rng.Copy ActiveSheet.Range("A1")
    'ActiveSheet.Range("A1") = select_rng.Value
    ActiveWorkbook.SaveAs Filename:=csvFileName, FileFormat:=xlCSV
    ActiveWindow.Close

    Application.DisplayAlerts = True

End Function

Sub testt()
    If InStr(ThisWorkbook.Name, ".xlsm") > 0 Then
         a = Left(ThisWorkbook.Name, Len(ThisWorkbook.Name) - 5)
    End If
End Sub