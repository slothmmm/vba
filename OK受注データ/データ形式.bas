Sub 形成main()
'    Application.ScreenUpdating = False
'    Application.Calculation = xlCalculationManual     '手動計算

    Worksheets("csv").Activate
    B_LastRow = Cells(Rows.Count, 2).End(xlUp).Row  ''B列の最終行取得
    read_Col = 80                                   'CB列まで
    
    csvtData = Worksheets("csv").Range(Cells(1, 1), Cells(B_LastRow, read_Col))   'シートデータ取得
    
    Dim paste_data() As Variant
    ReDim paste_data(200, 8)
    
    paste_data(0, 0) = "No"
    paste_data(0, 1) = "商品コード"
    paste_data(0, 2) = "商品名"
    paste_data(0, 3) = "センターコード"
    paste_data(0, 4) = "センター名"
    paste_data(0, 5) = "数量"
    paste_data(0, 6) = "バーコード"
    paste_data(0, 7) = "センター納品日"
    
    ' *****************************   重複なしリスト作成 ***********************************************
         Worksheets("csv").Activate
    B_LastRow = Cells(Rows.Count, 2).End(xlUp).Row  ''B列の最終行取得
    
    ' ******センターコードリスト
    serch_row = 3
    serch_data = Worksheets("csv").Range(Cells(2, serch_row), Cells(B_LastRow, serch_row))   'シートデータ取得
    center_list = 重複なしリスト(serch_data)

    Dim syoCD_list() As Variant
    ReDim syoCD_list(300, 1)
    add_no = 0
    ' ******商品コードリスト
    For i = 0 To 9
        serch_row = 9 + (i * 3)
        serch_data = Worksheets("csv").Range(Cells(2, serch_row), Cells(B_LastRow, serch_row))   'シートデータ取得
        syo_tmp_list = 重複なしリスト(serch_data)
        For s = 0 To UBound(syo_tmp_list)
            syoCD_list(add_no, 1) = syo_tmp_list(s)
            add_no = add_no + 1
        Next s
    Next i
    
    syoCD_list = 重複なしリスト(syoCD_list)

    ' *****************************   集計 ***********************************************
    '********paste_dataの土台作成
    add_p_no = 1
    For b = 0 To UBound(center_list)
      For h = 0 To UBound(syoCD_list)
        paste_data(add_p_no, 0) = add_p_no
        paste_data(add_p_no, 1) = syoCD_list(h)
        paste_data(add_p_no, 3) = center_list(b)
        add_p_no = add_p_no + 1
      Next h
    Next b

    '*********csvDataからpaste_dataへ集計
    For s = 1 To UBound(csvtData)
      For Z = 0 To 9
        serch_row = 11 + (Z * 3)
        If IsNumeric(csvtData(s, serch_row)) And Not IsEmpty(csvtData(s, serch_row)) And csvtData(s, serch_row) <> 0 Then
          'センターコードcsvtData(s, 3))
          '商品コードcsvtData(s, serch_row - 2))
          For p = 1 To UBound(paste_data)
            If paste_data(p, 1) = csvtData(s, serch_row - 2) And paste_data(p, 3) = csvtData(s, 3) Then
              paste_data(p, 5) = paste_data(p, 5) + csvtData(s, serch_row)
            End If
          Next p
        End If
      Next Z
    Next s
    
    '*********paste_dataへ納品日
    For p = 1 To UBound(paste_data)
        If Not IsEmpty(paste_data(p, 1)) Then
            paste_data(p, 7) = csvtData(2, 40)
        End If
    Next p
    
    'シートクリア
    Worksheets("形成").Activate
    Worksheets("形成").Cells.ClearContents
    
     Worksheets("形成").Range(Cells(1, 1), Cells(200, 8)) = paste_data
     
         Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
End Sub

Sub test()
     Worksheets("csv").Activate
    B_LastRow = Cells(Rows.Count, 2).End(xlUp).Row  ''B列の最終行取得
    
    'センターコードリスト
    serch_row = 3
    serch_data = Worksheets("csv").Range(Cells(2, serch_row), Cells(B_LastRow, serch_row))   'シートデータ取得
    center_list = 重複なしリスト(serch_data)

    Dim syoCD_list() As Variant
    ReDim syoCD_list(300, 1)
    add_no = 0
    '商品コードリスト
    For i = 0 To 9
        serch_row = 9 + (i * 3)
        serch_data = Worksheets("csv").Range(Cells(2, serch_row), Cells(B_LastRow, serch_row))   'シートデータ取得
        syo_tmp_list = 重複なしリスト(serch_data)
        For s = 0 To UBound(syo_tmp_list)
            syoCD_list(add_no, 1) = syo_tmp_list(s)
            add_no = add_no + 1
        Next s
    Next i
    
    syoCD_list = 重複なしリスト(syoCD_list)
    
End Sub

Function 重複なしリスト(dataList As Variant) As Variant
    

    Dim 辞書 As Object
    Set 辞書 = CreateObject("Scripting.Dictionary")

    For i = 1 To UBound(dataList)
        If dataList(i, UBound(dataList, 2)) <> Empty Then
            If Not 辞書.Exists(dataList(i, UBound(dataList, 2))) Then
                辞書(dataList(i, UBound(dataList, 2))) = Empty
            End If
        End If
    Next i
    
    重複なしリスト = 辞書.keys
    
End Function


