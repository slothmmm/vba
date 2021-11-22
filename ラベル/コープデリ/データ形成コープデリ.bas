Sub 形成用()
    Dim rc As VbMsgBoxResult
    rc = MsgBox("データ更新を行いますか？「形成」シートが更新されます。", vbYesNo + vbQuestion)
    If rc = vbNo Then
        MsgBox "データ更新を中止します", vbCritical
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual     '手動計算

    Worksheets("ピッキング表").Activate
    B_LastRow = Cells(Rows.Count, 2).End(xlUp).Row  ''B列の最終行取得
    read_Col = 80                                   'CB列まで
    
    sheetData = Worksheets("ピッキング表").Range(Cells(1, 1), Cells(B_LastRow, read_Col))   'シートデータ取得
    isOne = False                               'B列は1から始まる
    dataNo = 1                                  'ペーストするデータ行数(カラム行は0、データ行は1から)
    cenDataNo = 1                               'センターごとのデータ行数
    B_MAX_NUM = WorksheetFunction.Max(Range("B:B"))
    Dim paste_data() As Variant
    ReDim paste_data(5000, 12)
    
    paste_data(0, 0) = "No"
    paste_data(0, 1) = "商品コード"
    paste_data(0, 2) = "商品名"
    paste_data(0, 3) = "入数"
    paste_data(0, 4) = "センターコード"
    paste_data(0, 5) = "センター名"
    paste_data(0, 6) = "数量"
    paste_data(0, 7) = "バーコード"
    paste_data(0, 8) = "センター納品日"
    paste_data(0, 9) = "センター毎データNo"
    paste_data(0, 10) = "センターNo_データNo"
    
    
'センターコードを判定。６行目を６列目～15列目まで回す。
    For k = 6 To 15
        cenDataNo = 1                               'センターごとのデータ行数リセット
        If IsNumeric(sheetData(6, k)) And Len(sheetData(6, k)) = 4 And sheetData(7, k) <> "" And sheetData(7, k) <> 0 Then
                        
        For i = 1 To B_LastRow
            'B列が「1」から開始するようにする
            If Not isOne Then
                If sheetData(i, 2) = 1 Then
                    isOne = True
                Else
                    GoTo Continue ' Continue: の行へ処理を飛ばす
                End If
            End If
            
            '↑のisOneがTrueでB列が1があった行数からのi
            If IsNumeric(sheetData(i, 2)) And sheetData(i, 2) <> "" And sheetData(i, 3) <> 0 And Len(sheetData(i, 3)) = 4 And sheetData(i, 3) <> "" And sheetData(i, k) <> 0 And sheetData(i, k) <> "" Then
                Debug.Print "商品コードB列は " + Str(sheetData(i, 2))
                
                If (sheetData(i, k) / sheetData(i, 5)) <= 1 Then
                        paste_data(dataNo, 0) = dataNo               'No
                        paste_data(dataNo, 1) = sheetData(i, 3)      '商品コード
                        paste_data(dataNo, 2) = sheetData(i, 4)      '商品名
                        paste_data(dataNo, 3) = sheetData(i, 5)     '入数
                        paste_data(dataNo, 4) = sheetData(6, k)     'センターコード
                        paste_data(dataNo, 5) = sheetData(3, k)    'センター名
                        paste_data(dataNo, 6) = sheetData(i, k)     '数量
                        paste_data(dataNo, 7) = "" 'Str(sheetData(i, 66))    'JAN
                        paste_data(dataNo, 8) = sheetData(6, 4) + 1  'センター納品日
                        paste_data(dataNo, 9) = cenDataNo            'センターごとデータ数
                        paste_data(dataNo, 10) = Int(sheetData(6, k) + Format(cenDataNo, "0000"))  'センターコード+センターごとデータ数
                        dataNo = dataNo + 1                         'データ数
                        cenDataNo = cenDataNo + 1                   'センターごとデータ数
                Else
                    kitsum = sheetData(i, k)
                    For n = 1 To WorksheetFunction.RoundUp((sheetData(i, k) / sheetData(i, 5)), 0)
                        paste_data(dataNo, 0) = dataNo               'No
                        paste_data(dataNo, 1) = sheetData(i, 3)      '商品コード
                        paste_data(dataNo, 2) = sheetData(i, 4)      '商品名
                        paste_data(dataNo, 3) = sheetData(i, 5)     '入数
                        paste_data(dataNo, 4) = sheetData(6, k)     'センターコード
                        paste_data(dataNo, 5) = sheetData(3, k)    'センター名
                        
                        If n = WorksheetFunction.RoundUp((sheetData(i, k) / sheetData(i, 5)), 0) Then   'for最後かどうか
                            If sheetData(i, k) Mod sheetData(i, 5) = 0 Then
                                paste_data(dataNo, 6) = sheetData(i, 5)                         'コンテナフル数量
                            Else
                                paste_data(dataNo, 6) = sheetData(i, k) Mod sheetData(i, 5)     '余り数量
                            End If
                        Else
                            paste_data(dataNo, 6) = sheetData(i, 5)                         'コンテナフル数量
                        End If
                        
                        paste_data(dataNo, 7) = "" 'Str(sheetData(i, 66))    'JAN
                        paste_data(dataNo, 8) = sheetData(6, 4) + 1    'センター納品日
                        paste_data(dataNo, 9) = cenDataNo            'センターごとデータ数
                        paste_data(dataNo, 10) = Int(sheetData(6, k) + Format(cenDataNo, "0000"))  'センターコード+センターごとデータ数
                        dataNo = dataNo + 1
                        cenDataNo = cenDataNo + 1                   'センターごとデータ数
                    Next n
                End If
            End If
       
Continue:             ' GoTo Continue の後はここから処理が行われる
        Next i
        
        End If
    Next k
    
    Worksheets("形成").Activate
    Worksheets("形成").Cells.ClearContents
    Worksheets("形成").Range(Cells(1, 1), Cells(5000, 12)) = paste_data
    Worksheets("形成").Range("A1").AutoFilter
    Worksheets("ラベル").Activate
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic  '自動計算

End Sub

