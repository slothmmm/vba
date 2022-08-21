Sub 形成用()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual     '手動計算

    
    Worksheets("センター仕分け一覧").Activate
    Application.Calculate                             '再計算
    A_LastRow = Cells(Rows.Count, 1).End(xlUp).Row    'A列の最終行取得
    read_col = 26                                     'Z列まで
    
    sheetData = Worksheets("センター仕分け一覧").Range(Cells(1, 1), Cells(A_LastRow, read_col))    'シートデータ取得
    isOne = False                               'A列は1から始まる
    dataNo = 1                                  'ペーストするデータ行数(カラム行は0、データ行は1から)
    A_MAX_NUM = WorksheetFunction.Max(Range("A:A"))
    Dim paste_data() As Variant
    ReDim paste_data(5000, 5)
    
    paste_data(0, 0) = "No"
    paste_data(0, 1) = "自社商品コード"
    paste_data(0, 2) = "自社商品名"
    paste_data(0, 3) = "伝票商品名"
    paste_data(0, 4) = "入数"
    
    ashi = 10                                 '1足 equal 10コンテナ
    For i = 1 To A_LastRow
        'A列が「1」から開始するようにする
        If Not isOne Then
            If sheetData(i, 1) = 1 Then
                isOne = True
            Else
                GoTo Continue ' Continue: の行へ処理を飛ばす
            End If
        End If
        
        '↑のisOneがTrueでA列が1があった行数からのi
        If IsNumeric(sheetData(i, 1)) And sheetData(i, 1) <> "" And sheetData(i, 2) <> 0 And Len(sheetData(i, 2)) = 4 And sheetData(i, 2) <> "" Then
            For n = 1 To WorksheetFunction.RoundUp(((sheetData(i, 6) / sheetData(i, 5)) / ashi), 0)
                paste_data(dataNo, 0) = dataNo               'No
                paste_data(dataNo, 1) = sheetData(i, 2)      '自社商品コード
                paste_data(dataNo, 2) = sheetData(i, 3)      '自社商品名
                paste_data(dataNo, 3) = sheetData(i, 4)     '伝票商品名
                paste_data(dataNo, 4) = sheetData(i, 5)     '入数
                dataNo = dataNo + 1
            Next n
           
        End If

Continue:             ' GoTo Continue の後はここから処理が行われる
    Next i

    Worksheets("センター仕分け形成").Activate
    Worksheets("センター仕分け形成").Cells.ClearContents
    Worksheets("センター仕分け形成").Range(Cells(1, 1), Cells(5000, 5)) = paste_data
    Worksheets("センター仕分け形成").Range("A1").AutoFilter
    Worksheets("センター仕分け看板").Activate
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic  '自動計算

End Sub
