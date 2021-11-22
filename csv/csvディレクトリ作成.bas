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