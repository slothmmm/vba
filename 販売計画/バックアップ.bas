Sub バックアップ_main()
    '非表示
'    Application.ScreenUpdating = False
'    Application.DisplayAlerts = False

    FileName = "C:\Work\Sub\Book1.xls"
    a = Dir(FileName)
    'F2 出荷先
    Dim Customer_name As String
    'Customer_name = "東武"
    Customer_name = Mid(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, "\") + 1, InStrRev(ThisWorkbook.Name, ".") - InStrRev(ThisWorkbook.Name, "\") - 1)
    
    'F1出荷日
    Dim ship_date As Date
    ship_date = Now

    'ファイル名_出荷日名
    ship_filename = "【" & Format(Year(ship_date), "0000") & "年" & Format(Month(ship_date), "00") & "月" & Format(Day(ship_date), "00") & "日" & "】"
    'ファイル名_現在時刻
    'hiduke = Format(Year(Now), "0000") & "_" & Format(Month(Now), "00") & "_" & Format(Day(Now), "00") & "_" & Format(Hour(Now), "00") & "_" & Format(Minute(Now), "00") & "_" & Format(Second(Now), "00")
    
    'ファイル名
    Dim csvFileName As String
    yyyy = Format(Year(ship_date), "0000年")
    mm = Format(Month(ship_date), "00月")
    'saveFileName = ActiveWorkbook.Path & "\データ保管\" & Customer_name & "\" & yyyy & "\" & mm & "\" & ship_filename & Customer_name & hiduke & ".xlsm"
    saveFileName = ActiveWorkbook.Path & "\データ保管\" & Customer_name & "\" & yyyy & "\" & mm & "\" & ship_filename & Customer_name & ".xlsm"
    
    'ディレクトリ作成
    Call ディレクトリ作成1(Customer_name, ship_date)
    
     ThisWorkbook.SaveCopyAs saveFileName  'Backup保存するコード
    
    '非表示
'    Application.ScreenUpdating = True
'    Application.DisplayAlerts = True

End Sub

Sub ディレクトリ作成1(Customer_name As Variant, ship_date As Variant)
    Dim root As String
    Dim yyyy As String
    Dim mm As String
    Dim dd As String
    'F1出荷日
'    Dim ship_date As Date
'    ship_date = Now

    ' root = ActiveWorkbook.Path & "\csv"
    root = ActiveWorkbook.Path
    yyyy = Format(Year(ship_date), "0000年")
    mm = Format(Month(ship_date), "00月")
    ' dd = Format(Day(ship_date), "00日")

    'F2 出荷先
'    Dim Customer_name As String
'    Customer_name = Range("F2")

    Dim rtn As Long
    rtn = ディレクトリ作成2(root, "データ保管", Customer_name, yyyy, mm)
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
            MkDir Join(ary, "\")
        End If
    Next
  
    CreateDirectory = 0
    Exit Function
  
ErrExit:
    CreateDirectory = 9
End Function

