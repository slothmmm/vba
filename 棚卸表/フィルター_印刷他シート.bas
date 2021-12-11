Attribute VB_Name = "フィルター_印刷他シート"
'''''''''''''''''''''''''''''''      フィルター関連                   ''''''''''''''''''''''''''''
Sub フィルター全部クリア_印刷他シート()
    Call 保護.全保護解除
    Worksheets("印刷他").Activate
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
        
        .Range("K9012").Select
        If .FilterMode Then .ShowAllData
    End With
    Call 保護.複数保護
    Worksheets("印刷他").Activate
    MsgBox "フィルタークリア完了(印刷他シート)"
End Sub

Sub フィルター日曜日()
    Call 保護.全保護解除
    Worksheets("印刷他").Activate
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
        .ListObjects("テーブル2").Range.AutoFilter Field:=39, Criteria1:="<>"
    End With
    Call 保護.複数保護
    Worksheets("印刷他").Activate
    MsgBox "日曜日フィルター完了"
End Sub

Sub フィルター月曜日()
    Call 保護.全保護解除
    Worksheets("印刷他").Activate
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
        .ListObjects("テーブル2").Range.AutoFilter Field:=40, Criteria1:="<>"
    End With
    
    Call ソート印刷他.ソート仕入先名
    Call 保護.複数保護
    Worksheets("印刷他").Activate
    MsgBox "月曜日フィルター完了"
End Sub

Sub フィルター火曜日()
    Call 保護.全保護解除
    Worksheets("印刷他").Activate
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
        .ListObjects("テーブル2").Range.AutoFilter Field:=41, Criteria1:="<>"
    End With
    
    Call ソート印刷他.ソート仕入先名
    Call 保護.複数保護
    Worksheets("印刷他").Activate
    MsgBox "火曜日フィルター完了"
End Sub

Sub フィルター水曜日()
    Call 保護.全保護解除
    Worksheets("印刷他").Activate
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
        .ListObjects("テーブル2").Range.AutoFilter Field:=42, Criteria1:="<>"
    End With
    
    Call ソート印刷他.ソート仕入先名
    Call 保護.複数保護
    Worksheets("印刷他").Activate
    MsgBox "水曜日フィルター完了"
End Sub

Sub フィルター木曜日()
    Call 保護.全保護解除
    Worksheets("印刷他").Activate
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
        .ListObjects("テーブル2").Range.AutoFilter Field:=43, Criteria1:="<>"
    End With
    Call 保護.複数保護
    Worksheets("印刷他").Activate
    MsgBox "木曜日フィルター完了"
End Sub

Sub フィルター金曜日()
    Call 保護.全保護解除
    Worksheets("印刷他").Activate
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
        .ListObjects("テーブル2").Range.AutoFilter Field:=44, Criteria1:="<>"
    End With
    Call 保護.複数保護
    Worksheets("印刷他").Activate
    MsgBox "金曜日フィルター完了"
End Sub

Sub フィルター土曜日()
    Call 保護.全保護解除
    Worksheets("印刷他").Activate
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
        .ListObjects("テーブル2").Range.AutoFilter Field:=45, Criteria1:="<>"
    End With
    Call 保護.複数保護
    Worksheets("印刷他").Activate
    MsgBox "土曜日フィルター完了"
End Sub

Sub フィルター副原()
    Call 保護.全保護解除
    Worksheets("印刷他").Activate
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
        .ListObjects("テーブル2").Range.AutoFilter Field:=26, Criteria1:="<>"
    End With
    Call 保護.複数保護
    Worksheets("印刷他").Activate
    MsgBox "副原フィルター完了"
End Sub

Sub フィルターCN_印刷他シート()
    Call 保護.全保護解除
    Worksheets("印刷他").Activate
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
        .ListObjects("テーブル2").Range.AutoFilter Field:=27, Criteria1:="<>"
    End With
    Call 保護.複数保護
    Worksheets("印刷他").Activate
    MsgBox "CNフィルター完了"
End Sub
Sub フィルターCN_印刷他シート_賞味期限用()
    Call 保護.全保護解除
    Worksheets("印刷他").Activate
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
        .ListObjects("テーブル2").Range.AutoFilter Field:=52, Criteria1:="<>"
    End With
    Call ソート印刷他.ソート座標
    Call 保護.複数保護
    Worksheets("印刷他").Activate
    MsgBox "CNフィルター完了"
End Sub

Sub フィルターIY()
    Call 保護.全保護解除
    Worksheets("印刷他").Activate
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
        .ListObjects("テーブル2").Range.AutoFilter Field:=31, Criteria1:="<>"
    End With
    Call 保護.複数保護
    Worksheets("印刷他").Activate
    MsgBox "IYフィルター完了"
End Sub
Sub フィルターIY重複()
    Call 保護.全保護解除
    Worksheets("印刷他").Activate
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
        .ListObjects("テーブル2").Range.AutoFilter Field:=30, Criteria1:="<>"
    End With
    Call 保護.複数保護
    Worksheets("印刷他").Activate
    MsgBox "IYフィルター完了"
End Sub
Sub フィルターシール()
    Call 保護.全保護解除
    Worksheets("印刷他").Activate
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
        .ListObjects("テーブル2").Range.AutoFilter Field:=38, Criteria1:="シール"
        .ListObjects("テーブル2").Range.AutoFilter Field:=6, Criteria1:="<>-"
    End With
    
    Call ソート印刷他.ソート仕入先名
    Call 保護.複数保護
    Worksheets("印刷他").Activate
    MsgBox "シールフィルター完了"
End Sub

Sub フィルター漏れ()
    Call 保護.全保護解除
    Worksheets("印刷他").Activate
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
        .ListObjects("テーブル2").Range.AutoFilter Field:=36, Criteria1:="<>"
    End With
    Call 保護.複数保護
    Worksheets("印刷他").Activate
    MsgBox "漏れフィルター完了"
End Sub

Sub フィルター小分け()
    '期末在庫専用のボタン 2020.09.13コメント
    'BJ列に追加（62列目）
    '数式は以下
    '=IF(IFERROR(MATCH([@商品コード],'\\Afnewt320-kyoyu\社内共有\AFSKS\在庫表\小分け品在庫表\[在庫『小分け品』_2020.08.xlsm]仕掛かり【中継】'!$D:$D,0),"")<>"","☆","")
    
    Call 保護.全保護解除
    Worksheets("印刷他").Activate
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
        .ListObjects("テーブル2").Range.AutoFilter Field:=11, Criteria1:="<>"
        .ListObjects("テーブル2").Range.AutoFilter Field:=20, Criteria1:=""
        .ListObjects("テーブル2").Range.AutoFilter Field:=27, Criteria1:=""
        .ListObjects("テーブル2").Range.AutoFilter Field:=28, Criteria1:=""
        .ListObjects("テーブル2").Range.AutoFilter Field:=62, Criteria1:="☆"
        .ListObjects("テーブル5").Range.AutoFilter Field:=1, Criteria1:="諸口"
    End With
    
    Call ソート印刷他.ソート仕入先名
    Call 保護.複数保護
    Worksheets("印刷他").Activate
    MsgBox "小分けフィルター完了"
End Sub
