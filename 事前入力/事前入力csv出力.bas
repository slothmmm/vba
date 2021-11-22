
Sub 事前入力csv_main()
    'A1セルから空白セルまでdowhileで回し、その領域内をcsv出力する
    
    '非表示
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    'アクティブ
    Worksheets("出力用").Activate
    Worksheets("出力用").Select
    Range("A1").Select
    
    'ワークシート
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("出力用")

    'F2 出荷先
    Dim Customer_name As String
    Customer_name = Range("F2")

    'F1出荷日
    Dim ship_date As Date
    ship_date = Range("F1")

    'csvファイル名_出荷日名
    ship_filename = "出荷日【" & Format(Year(ship_date), "0000") & "年" & Format(Month(ship_date), "00") & "月" & Format(Day(ship_date), "00") & "日" & "】"
    'csvファイル名_現在時刻
    hiduke = Format(Year(Now), "0000") & "_" & Format(Month(Now), "00") & "_" & Format(Day(Now), "00") & "_" & Format(Hour(Now), "00") & "_" & Format(Minute(Now), "00") & "_" & Format(Second(Now), "00")
    
    
    'csvファイル名
    Dim csvFileName As String
    yyyy = Format(Year(ship_date), "0000年")
    mm = Format(Month(ship_date), "00月")
    csvFileName = ActiveWorkbook.Path & "\コープ事前入力csv\" & Customer_name & "\" & yyyy & "\" & mm & "\" & ship_filename & Customer_name & hiduke & ".csv"
    
    'ディレクトリ作成
    Call 事前入力ディレクトリ作成1
    'CSV Open >> Close
    Open csvFileName For Output As #1
    
    Dim i As Long, j As Long
    i = 1
    
    Do While ws.Cells(i, 1).Value <> ""
    
        j = 1
        Do While ws.Cells(i, j + 1).Value <> ""
    
            Print #1, ws.Cells(i, j).Value & ",";
            j = j + 1
    
        Loop
    
        Print #1, ws.Cells(i, j).Value & vbCr;
        i = i + 1
    
    Loop
    
    Close #1

    
    'アクティブ
    Worksheets("ピッキング表").Activate
    Worksheets("ピッキング表").Select
    Range("A1").Select
    
    '非表示
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

Sub 事前入力ディレクトリ作成1()
    Dim root As String
    Dim yyyy As String
    Dim mm As String
    Dim dd As String
    'F1出荷日
    Dim ship_date As Date
    ship_date = Range("F1")

    ' root = ActiveWorkbook.Path & "\csv"
    root = ActiveWorkbook.Path
    yyyy = Format(Year(ship_date), "0000年")
    mm = Format(Month(ship_date), "00月")
    ' dd = Format(Day(ship_date), "00日")

    'F2 出荷先
    Dim Customer_name As String
    Customer_name = Range("F2")

    Dim rtn As Long
    rtn = 事前入力ディレクトリ作成2(root, "コープ事前入力csv", Customer_name, yyyy, mm)
'    Select Case rtn
'        Case 0
'            MsgBox "フォルダを作成しました。"
'        Case 1
'            MsgBox "フォルダは存在します。"
'        Case Else
'            MsgBox "フォルダの作成に失敗しました。"
'    End Select
End Sub

Function 事前入力ディレクトリ作成2(ParamArray arg()) As Long
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



