Option Base 1

Sub 更新_形成用小分けボタン()
    Call 更新_形成用小分け(1)
End Sub
 
Function 更新_形成用小分け(button_bool As Variant)
'    Dim rc As VbMsgBoxResult
'    rc = MsgBox("データ更新を行いますか？「**」シートが更新されます。", vbYesNo + vbQuestion)
'    If rc = vbNo Then
'        MsgBox "データ更新を中止します", vbCritical
'        Exit Sub
'    End If
    Application.ScreenUpdating = True
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual     '手動計算

    '１足に何コンテナか
    ashi = 6

    Worksheets("小分払出").Activate
    A_LastRow = Cells(Rows.Count, 1).End(xlUp).Row  ''B列の最終行取得
    read_Col = 40                                   'CB列まで
    
    sheetData = Worksheets("小分払出").Range(Cells(1, 1), Cells(A_LastRow, read_Col))   'シートデータ取得
    A_MAX_NUM_1_499 = WorksheetFunction.Max(Range("A1:A440"))                   '煮物を除く
    
    Dim paste_data() As Variant         '貼付け
    ReDim paste_data(A_LastRow, read_Col)
    dataNo = 1      '貼付けNo A列
    
    For i = 1 To A_MAX_NUM_1_499
        If sheetData((i * 5) + 1, 12) > 0 Then  '指示数があるか
            kon_num = WorksheetFunction.RoundUp(sheetData((i * 5) + 1, 12) / sheetData((i * 5) + 1, 27), 0) 'コンテナ数
            asi_num = WorksheetFunction.RoundUp(kon_num / ashi, 0)                                          '足の数
            For m = 1 To asi_num
                For a = (i * 5) + 1 To (i * 5) + 1 + 5
                    For b = 1 To read_Col
                        If b = 1 Then
                            paste_data((dataNo - 1) * 5 + (a - (i * 5) + 1), b) = dataNo
                        ElseIf b = 2 Then
                            paste_data((dataNo - 1) * 5 + (a - (i * 5) + 1), b) = dataNo + (((((i * 5) + 1) - a) / 10) * -1)
                        Else
                            paste_data((dataNo - 1) * 5 + (a - (i * 5) + 1), b) = sheetData(a, b)
                        End If
                        
                        If b = 30 Then
                            If m = asi_num Then
                                If sheetData((i * 5) + 1, 12) Mod sheetData((i * 5) + 1, 27) * ashi = 0 Then
                                    paste_data((dataNo - 1) * 5 + (a - (i * 5) + 1), b) = sheetData((i * 5) + 1, 27) * ashi
                                Else
                                    paste_data((dataNo - 1) * 5 + (a - (i * 5) + 1), b) = sheetData((i * 5) + 1, 12) Mod sheetData((i * 5) + 1, 27) * ashi
                                End If
                            Else
                                paste_data((dataNo - 1) * 5 + (a - (i * 5) + 1), b) = sheetData((i * 5) + 1, 27) * ashi
                            End If
                        End If
                    Next b
                Next a
                dataNo = dataNo + 1
            Next m
        End If
    Next i
    
    Worksheets("形成").Activate
    Worksheets("形成").Cells.ClearContents
    Worksheets("形成").Range(Cells(1, 1), Cells(A_LastRow, read_Col)) = paste_data
'    Worksheets("形成").Range("A1").AutoFilter
    Worksheets("小分保管表").Activate
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic  '自動計算
    
    If button_bool = 1 Then
        MsgBox ("更新完了しました。")
    End If
End Function



