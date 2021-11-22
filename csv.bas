Sub writeCSV()
    'A1セルから空白セルまでdowhileで回し、その領域内をcsv出力する
    Worksheets("出力用").Activate
    Worksheets("出力用").Select
    Range("A1").Select
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("出力用")

    'F2 出荷先
    Dim Customer_name As String
    Customer_name = Range("F2")

    'F1出荷日
    Dim ship_date As Date
    ship_date = Range("F1")
    'csv出力名
    ship_filename = "出荷日【" & Format(Year(ship_date), "0000") & "年" & Format(Month(ship_date), "00") & "月" & Format(Day(ship_date), "00") & "日" & "】"
    hiduke = Format(Year(Now), "0000") & "_" & Format(Month(Now), "00") & "_" & Format(Day(Now), "00") & "_" & Format(Hour(Now), "00") & "_" & Format(Minute(Now), "00") & "_" & Format(Second(Now), "00")
    Dim csvFile As String
    csvFile = ActiveWorkbook.Path & "\" & ship_filename & Customer_name & hiduke & ".csv"
    'CSV Open >> Close
    Open csvFile For Output As #1
    
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




End Sub


Sub sample()
    'A1セルから空白セルまでdowhileで回し、その領域内をcsv出力する
    Dim Customer_name As String
    Customer_name = "コープデリ"
    
    Dim ship_date As Date
    ship_date = Range("F1")
    Debug.Print ship_date

    Dim SaveDir As String

    SaveDir = ActiveWorkbook.Path & "\" & Customer_name & "\" & Format(Year(ship_date), "0000") & "年" & Format(Month(ship_date), "00") & "月"
    
    ' "C:\Data"の下に今日の日付のフォルダがなければ作成する
    If Dir(SaveDir, vbDirectory) = "" Then
        MkDir SaveDir
    End If
    

End Sub

Sub ディレクトリ作成1()
    Dim root As String
    Dim yyyy As String
    Dim mm As String
    Dim dd As String
  
    root = ActiveWorkbook.Path & "\csv"
    yyyy = Format(Year(Date), "0000年")
    mm = Format(Month(Date), "00月")
    ' dd = Format(Day(Date), "00日")

    Dim rtn As Long
    rtn = ディレクトリ作成2(root,"csv", yyyy, mm)
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
