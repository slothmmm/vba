Attribute VB_Name = "ソート印刷他"
Sub ソート仕入先名()
    Call 保護.全保護解除
    Worksheets("印刷他").Activate
    
    ActiveWorkbook.Worksheets("印刷他").ListObjects("テーブル2").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("印刷他").ListObjects("テーブル2").Sort.SortFields.Add Key _
        :=Range("テーブル2[[#All],[仕入先名]]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("印刷他").ListObjects("テーブル2").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Call 保護.複数保護
End Sub

Sub ソート商品コード()
    Call 保護.全保護解除
    Worksheets("印刷他").Activate
    
    ActiveWorkbook.Worksheets("印刷他").ListObjects("テーブル2").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("印刷他").ListObjects("テーブル2").Sort.SortFields.Add Key _
        :=Range("テーブル2[[#All],[商品コード]]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal       'xlAscending 昇順     xlDescending 降順
    With ActiveWorkbook.Worksheets("印刷他").ListObjects("テーブル2").Sort
        .Header = xlYes                          '先頭行を見出しとして使用
        .MatchCase = False                      '大文字小文字を区別しない
        .Orientation = xlTopToBottom            '行単位で並べ替え
        .SortMethod = xlPinYin                  'ふりがなを使わない
        .Apply                                  '並べ替えを実行
    End With
    Call 保護.複数保護
End Sub

Sub ソート座標()
    Call 保護.全保護解除
    Worksheets("印刷他").Activate
    
    ActiveWorkbook.Worksheets("印刷他").ListObjects("テーブル2").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("印刷他").ListObjects("テーブル2").Sort.SortFields.Add Key _
        :=Range("テーブル2[[#All],[座標]]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal       'xlAscending 昇順     xlDescending 降順
    With ActiveWorkbook.Worksheets("印刷他").ListObjects("テーブル2").Sort
        .Header = xlYes                          '先頭行を見出しとして使用
        .MatchCase = False                      '大文字小文字を区別しない
        .Orientation = xlTopToBottom            '行単位で並べ替え
        .SortMethod = xlPinYin                  'ふりがなを使わない
        .Apply                                  '並べ替えを実行
    End With
    Call 保護.複数保護
End Sub

Sub ソートクリア()
    Call 保護.全保護解除
    ActiveWorkbook.Worksheets("印刷他").ListObjects("テーブル2").Sort.SortFields.Clear
    Call 保護.複数保護
End Sub

Sub ソートリセット()
    Call ソート商品コード
    Call ソートクリア
    Worksheets("印刷他").Activate
    MsgBox "ソートリセット完了"
End Sub
