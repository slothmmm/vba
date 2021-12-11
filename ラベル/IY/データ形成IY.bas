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
    paste_data(0, 7) = "IY商品コード"
    paste_data(0, 8) = "出荷日"
    paste_data(0, 9) = "センター毎データNo"
    paste_data(0, 10) = "センターNo_データNo"
    
    conNo = 1   'コンテナ数
    
'センターコードを判定。６行目を６列目～15列目まで回す。
    For c = 5 To 15
        cenDataNo = 1                               'センターごとのデータ行数リセット
        If IsNumeric(sheetData(5, c)) And Len(sheetData(5, c)) = 5 And sheetData(4, c) <> "" And sheetData(4, c) <> 0 Then
                        
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
            If IsNumeric(sheetData(i, 2)) And sheetData(i, 2) <> "" And sheetData(i, 3) <> 0 And Len(sheetData(i, 3)) = 4 And sheetData(i, 3) <> "" And sheetData(i, c) <> 0 And sheetData(i, c) <> "" Then
                '単品商品計算
                vol = 0
                itemNo = 1
                for s = 1 To sheetData(i, c)
                    vol = vol + 1/sheetData(i, 36) 
                    if vol < 1 and timeNo <=7 Then
                        paste_data(dataNo, 0) = dataNo               'No
                        paste_data(dataNo, 1) = sheetData(i, 3)      '商品コード
                        paste_data(dataNo, 2) = sheetData(i, 4)      '商品名
                        paste_data(dataNo, 3) = sheetData(i, 36)     '入数
                        paste_data(dataNo, 4) = sheetData(5, c)     'センターコード
                        paste_data(dataNo, 5) = sheetData(4, c)    'センター名
                        paste_data(dataNo, 6) = sheetData(i, c)     '数量
                        paste_data(dataNo, 7) = sheetData(i, 21)    'IY商品コード
                        paste_data(dataNo, 8) = sheetData(1, 3)  '出荷日
                        paste_data(dataNo, 9) = cenDataNo            'センターごとデータ数
                        paste_data(dataNo, 10) = Int(sheetData(5, c) + Format(cenDataNo, "0000"))  'センターコード+センターごとデータ数
                        dataNo = dataNo + 1                         'データ数
                        cenDataNo = cenDataNo + 1                   'センターごとデータ数
                    Else

                        vol =  1/sheetData(i, 36)   '体積リセット
                        itemNo = 1                  '商品Noリセット
                        conNo = conNo +1            '混んて
                    end if
                next s
            End If

Continue:             ' GoTo Continue の後はここから処理が行われる
        Next i
        
        End If
    Next c
    
    Worksheets("形成").Activate
    Worksheets("形成").Cells.ClearContents
    Worksheets("形成").Range(Cells(1, 1), Cells(5000, 12)) = paste_data
    Worksheets("形成").Range("A1").AutoFilter
    Worksheets("ラベル").Activate
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic  '自動計算

End Sub



