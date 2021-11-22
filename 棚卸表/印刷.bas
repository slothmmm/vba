Attribute VB_Name = "印刷"
'''''''''''''''''''''''''''''''          印刷                     '''''''''''''''''''''''''''''''''''
Sub 印刷_副原材料()
    Call 保護.全保護解除
    Worksheets("印刷他").Activate
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
        .ListObjects("テーブル2").Range.AutoFilter Field:=26, Criteria1:="<>"
        .ListObjects("テーブル5").Range.AutoFilter Field:=1, Criteria1:="副原材料"
    End With
    
    Call ソート印刷他.ソート仕入先名
    
    '// プリンタとの接続を切断
    Application.PrintCommunication = False
    '// 印刷設定
    With ActiveSheet.PageSetup
        .PrintArea = Sheets("印刷他").Range("A8:J9016").Address
        .Orientation = xlPortrait '印刷向きを縦方向に設定
        .Zoom = False '拡大縮小を設定（しない）
        .FitToPagesWide = 1 'すべての列を1ページに印刷
        .FitToPagesTall = False 'シートを1ページに印刷
        
        .LeftMargin = Application.CentimetersToPoints(0.5)
        .RightMargin = Application.CentimetersToPoints(0.5)
        .TopMargin = Application.CentimetersToPoints(0.5)
        .BottomMargin = Application.CentimetersToPoints(0.5)
        .HeaderMargin = Application.CentimetersToPoints(0)
        .FooterMargin = Application.CentimetersToPoints(0)
    End With
    '// プリンタと接続する
    Application.PrintCommunication = True
    ActiveSheet.PrintOut
    
    Call 印刷_賞味期限
    Call 保護.複数保護
    Worksheets("印刷他").Activate
    ActiveSheet.Range("A11").Select
End Sub

Sub 印刷_日曜日()
    Call 保護.全保護解除
    Worksheets("印刷他").Activate
    
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
        .ListObjects("テーブル2").Range.AutoFilter Field:=39, Criteria1:="<>"
        .ListObjects("テーブル5").Range.AutoFilter Field:=1, Criteria1:="諸口"
    End With
    
    Call ソート印刷他.ソート仕入先名
    
    '// プリンタとの接続を切断
    Application.PrintCommunication = False
    '// 印刷設定
    With ActiveSheet.PageSetup
        .PrintArea = Sheets("印刷他").Range("A8:J9019").Address
        .Orientation = xlPortrait '印刷向きを縦方向に設定
        .Zoom = False '拡大縮小を設定（しない）
        .FitToPagesWide = 1 'すべての列を1ページに印刷
        .FitToPagesTall = False 'シートを1ページに印刷
        
        .LeftMargin = Application.CentimetersToPoints(0.5)
        .RightMargin = Application.CentimetersToPoints(0.5)
        .TopMargin = Application.CentimetersToPoints(0.5)
        .BottomMargin = Application.CentimetersToPoints(0.5)
        .HeaderMargin = Application.CentimetersToPoints(0)
        .FooterMargin = Application.CentimetersToPoints(0)
    End With
    '// プリンタと接続する
    Application.PrintCommunication = True
    ActiveSheet.PrintOut
    Call 保護.複数保護
    ActiveSheet.Range("A11").Select
End Sub

Sub 印刷_月曜日()
    Call 保護.全保護解除
    Worksheets("印刷他").Activate
    
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
        .ListObjects("テーブル2").Range.AutoFilter Field:=40, Criteria1:="<>"
        .ListObjects("テーブル5").Range.AutoFilter Field:=1, Criteria1:="諸口"
    End With
    
    Call ソート印刷他.ソート仕入先名
    
    '// プリンタとの接続を切断
    Application.PrintCommunication = False
    '// 印刷設定
    With ActiveSheet.PageSetup
        .PrintArea = Sheets("印刷他").Range("A8:J9019").Address
        .Orientation = xlPortrait '印刷向きを縦方向に設定
        .Zoom = False '拡大縮小を設定（しない）
        .FitToPagesWide = 1 'すべての列を1ページに印刷
        .FitToPagesTall = False 'シートを1ページに印刷
        
        .LeftMargin = Application.CentimetersToPoints(0.5)
        .RightMargin = Application.CentimetersToPoints(0.5)
        .TopMargin = Application.CentimetersToPoints(0.5)
        .BottomMargin = Application.CentimetersToPoints(0.5)
        .HeaderMargin = Application.CentimetersToPoints(0)
        .FooterMargin = Application.CentimetersToPoints(0)
    End With
    '// プリンタと接続する
    Application.PrintCommunication = True
    ActiveSheet.PrintOut
    Call 保護.複数保護
    ActiveSheet.Range("A11").Select
End Sub

Sub 印刷_火曜日()
    Call 保護.全保護解除
    Worksheets("印刷他").Activate
    
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
        .ListObjects("テーブル2").Range.AutoFilter Field:=41, Criteria1:="<>"
        .ListObjects("テーブル5").Range.AutoFilter Field:=1, Criteria1:="諸口"
    End With
    
    Call ソート印刷他.ソート仕入先名
    
    '// プリンタとの接続を切断
    Application.PrintCommunication = False
    '// 印刷設定
    With ActiveSheet.PageSetup
        .PrintArea = Sheets("印刷他").Range("A8:J9019").Address
        .Orientation = xlPortrait '印刷向きを縦方向に設定
        .Zoom = False '拡大縮小を設定（しない）
        .FitToPagesWide = 1 'すべての列を1ページに印刷
        .FitToPagesTall = False 'シートを1ページに印刷
        
        .LeftMargin = Application.CentimetersToPoints(0.5)
        .RightMargin = Application.CentimetersToPoints(0.5)
        .TopMargin = Application.CentimetersToPoints(0.5)
        .BottomMargin = Application.CentimetersToPoints(0.5)
        .HeaderMargin = Application.CentimetersToPoints(0)
        .FooterMargin = Application.CentimetersToPoints(0)
    End With
    '// プリンタと接続する
    Application.PrintCommunication = True
    ActiveSheet.PrintOut
    Call 保護.複数保護
    ActiveSheet.Range("A11").Select
End Sub

Sub 印刷_水曜日()
    Call 保護.全保護解除
    Worksheets("印刷他").Activate
    
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
        .ListObjects("テーブル2").Range.AutoFilter Field:=42, Criteria1:="<>"
        .ListObjects("テーブル5").Range.AutoFilter Field:=1, Criteria1:="諸口"
    End With
    
    Call ソート印刷他.ソート仕入先名
    
    '// プリンタとの接続を切断
    Application.PrintCommunication = False
    '// 印刷設定
    With ActiveSheet.PageSetup
        .PrintArea = Sheets("印刷他").Range("A8:J9019").Address
        .Orientation = xlPortrait '印刷向きを縦方向に設定
        .Zoom = False '拡大縮小を設定（しない）
        .FitToPagesWide = 1 'すべての列を1ページに印刷
        .FitToPagesTall = False 'シートを1ページに印刷
        
        .LeftMargin = Application.CentimetersToPoints(0.5)
        .RightMargin = Application.CentimetersToPoints(0.5)
        .TopMargin = Application.CentimetersToPoints(0.5)
        .BottomMargin = Application.CentimetersToPoints(0.5)
        .HeaderMargin = Application.CentimetersToPoints(0)
        .FooterMargin = Application.CentimetersToPoints(0)
    End With
    '// プリンタと接続する
    Application.PrintCommunication = True
    ActiveSheet.PrintOut
    Call 保護.複数保護
    ActiveSheet.Range("A11").Select
End Sub

Sub 印刷_木曜日()
    Call 保護.全保護解除
    Worksheets("印刷他").Activate
    
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
        .ListObjects("テーブル2").Range.AutoFilter Field:=43, Criteria1:="<>"
        .ListObjects("テーブル5").Range.AutoFilter Field:=1, Criteria1:="諸口"
    End With
    
    Call ソート印刷他.ソート仕入先名
    
    '// プリンタとの接続を切断
    Application.PrintCommunication = False
    '// 印刷設定
    With ActiveSheet.PageSetup
        .PrintArea = Sheets("印刷他").Range("A8:J9019").Address
        .Orientation = xlPortrait '印刷向きを縦方向に設定
        .Zoom = False '拡大縮小を設定（しない）
        .FitToPagesWide = 1 'すべての列を1ページに印刷
        .FitToPagesTall = False 'シートを1ページに印刷
        
        .LeftMargin = Application.CentimetersToPoints(0.5)
        .RightMargin = Application.CentimetersToPoints(0.5)
        .TopMargin = Application.CentimetersToPoints(0.5)
        .BottomMargin = Application.CentimetersToPoints(0.5)
        .HeaderMargin = Application.CentimetersToPoints(0)
        .FooterMargin = Application.CentimetersToPoints(0)
    End With
    '// プリンタと接続する
    Application.PrintCommunication = True
    ActiveSheet.PrintOut
    Call 保護.複数保護
    ActiveSheet.Range("A11").Select
End Sub

Sub 印刷_金曜日()
    Call 保護.全保護解除
    Worksheets("印刷他").Activate
    
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
        .ListObjects("テーブル2").Range.AutoFilter Field:=44, Criteria1:="<>"
        .ListObjects("テーブル5").Range.AutoFilter Field:=1, Criteria1:="諸口"
    End With
    
    Call ソート印刷他.ソート仕入先名
    
    '// プリンタとの接続を切断
    Application.PrintCommunication = False
    '// 印刷設定
    With ActiveSheet.PageSetup
        .PrintArea = Sheets("印刷他").Range("A8:J9019").Address
        .Orientation = xlPortrait '印刷向きを縦方向に設定
        .Zoom = False '拡大縮小を設定（しない）
        .FitToPagesWide = 1 'すべての列を1ページに印刷
        .FitToPagesTall = False 'シートを1ページに印刷
        
        .LeftMargin = Application.CentimetersToPoints(0.5)
        .RightMargin = Application.CentimetersToPoints(0.5)
        .TopMargin = Application.CentimetersToPoints(0.5)
        .BottomMargin = Application.CentimetersToPoints(0.5)
        .HeaderMargin = Application.CentimetersToPoints(0)
        .FooterMargin = Application.CentimetersToPoints(0)
    End With
    '// プリンタと接続する
    Application.PrintCommunication = True
    ActiveSheet.PrintOut
    Call 保護.複数保護
    ActiveSheet.Range("A11").Select
End Sub

Sub 印刷_土曜日()
    Call 保護.全保護解除
    Worksheets("印刷他").Activate
    
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
        .ListObjects("テーブル2").Range.AutoFilter Field:=45, Criteria1:="<>"
        .ListObjects("テーブル5").Range.AutoFilter Field:=1, Criteria1:="諸口"
    End With
    
    Call ソート印刷他.ソート仕入先名
    
    '// プリンタとの接続を切断
    Application.PrintCommunication = False
    '// 印刷設定
    With ActiveSheet.PageSetup
        .PrintArea = Sheets("印刷他").Range("A8:J9019").Address
        .Orientation = xlPortrait '印刷向きを縦方向に設定
        .Zoom = False '拡大縮小を設定（しない）
        .FitToPagesWide = 1 'すべての列を1ページに印刷
        .FitToPagesTall = False 'シートを1ページに印刷
        
        .LeftMargin = Application.CentimetersToPoints(0.5)
        .RightMargin = Application.CentimetersToPoints(0.5)
        .TopMargin = Application.CentimetersToPoints(0.5)
        .BottomMargin = Application.CentimetersToPoints(0.5)
        .HeaderMargin = Application.CentimetersToPoints(0)
        .FooterMargin = Application.CentimetersToPoints(0)
    End With
    '// プリンタと接続する
    Application.PrintCommunication = True
    ActiveSheet.PrintOut
    Call 保護.複数保護
    ActiveSheet.Range("A11").Select
End Sub

Sub 印刷_IY()
    Call 保護.全保護解除
    Worksheets("印刷他").Activate
    
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
        .ListObjects("テーブル2").Range.AutoFilter Field:=31, Criteria1:="<>"
        .ListObjects("テーブル5").Range.AutoFilter Field:=1, Criteria1:="IY"
    End With
    
    Call ソート印刷他.ソート仕入先名
    
    '// プリンタとの接続を切断
    Application.PrintCommunication = False
    '// 印刷設定
    With ActiveSheet.PageSetup
        .PrintArea = Sheets("印刷他").Range("A8:J9023").Address
        .Orientation = xlPortrait '印刷向きを縦方向に設定
        .Zoom = False '拡大縮小を設定（しない）
        .FitToPagesWide = 1 'すべての列を1ページに印刷
        .FitToPagesTall = False 'シートを1ページに印刷

        .LeftMargin = Application.CentimetersToPoints(0.5)
        .RightMargin = Application.CentimetersToPoints(0.5)
        .TopMargin = Application.CentimetersToPoints(0.5)
        .BottomMargin = Application.CentimetersToPoints(0.5)
        .HeaderMargin = Application.CentimetersToPoints(0)
        .FooterMargin = Application.CentimetersToPoints(0)
    End With
    '// プリンタと接続する
    Application.PrintCommunication = True
    ActiveSheet.PrintOut
    Call 保護.複数保護
    ActiveSheet.Range("A11").Select
End Sub

Sub 印刷_CNシート()
    Call 保護.全保護解除
    Worksheets("印刷CN").Activate
    
    With ActiveSheet
        .Range("B4").Select
        If .FilterMode Then .ShowAllData
        .Range("B4").AutoFilter Field:=26, Criteria1:="<>"
    End With
    
    '// プリンタとの接続を切断
    Application.PrintCommunication = False
    '// 印刷設定
    With ActiveSheet.PageSetup
        .PrintArea = Sheets("印刷CN").Range("D4:V227").Address
        .Orientation = xlLandscape                                  '印刷向きを横方向に設定
        .Zoom = False                                               '拡大縮小を設定（しない）
        .FitToPagesWide = 1                                         'すべての列を1ページに印刷
        .FitToPagesTall = False                                     'シートを1ページに印刷
        
        .LeftMargin = Application.CentimetersToPoints(0.5)
        .RightMargin = Application.CentimetersToPoints(0.5)
        .TopMargin = Application.CentimetersToPoints(0.5)
        .BottomMargin = Application.CentimetersToPoints(1)
        .HeaderMargin = Application.CentimetersToPoints(0)
        .FooterMargin = Application.CentimetersToPoints(0)
    End With
    '// プリンタと接続する
    Application.PrintCommunication = True
    ActiveSheet.PrintOut
    Call 保護.複数保護
End Sub

Sub 印刷_形成1()
    Call 保護.全保護解除
    Worksheets("形成1").Activate
    
    Call ソート形成1.ソート形成1_ALL
    
    With ActiveSheet
        .ListObjects("新館").Range.AutoFilter Field:=25, Criteria1:="<>"
        .ListObjects("商品管理").Range.AutoFilter Field:=25, Criteria1:="<>"
        .ListObjects("冷蔵庫").Range.AutoFilter Field:=25, Criteria1:="<>"
        .ListObjects("冷凍庫").Range.AutoFilter Field:=25, Criteria1:="<>"
        .ListObjects("その他").Range.AutoFilter Field:=25, Criteria1:="<>"
    End With
    
    '// プリンタとの接続を切断
    Application.PrintCommunication = False
    '// 印刷設定
    With ActiveSheet.PageSetup
        .PrintArea = Sheets("形成1").Range("D4:V216").Address
        .Orientation = xlLandscape                          '印刷向きを横方向に設定
        .Zoom = False                                       '拡大縮小を設定（しない）
        .FitToPagesWide = 1                                 'すべての列を1ページに印刷
        .FitToPagesTall = False                             'シートを1ページに印刷

        .LeftMargin = Application.CentimetersToPoints(0.5)
        .RightMargin = Application.CentimetersToPoints(0.5)
        .TopMargin = Application.CentimetersToPoints(0.5)
        .BottomMargin = Application.CentimetersToPoints(1)
        .HeaderMargin = Application.CentimetersToPoints(0)
        .FooterMargin = Application.CentimetersToPoints(0)
    End With
    '// プリンタと接続する
    Application.PrintCommunication = True
    ActiveSheet.PrintOut
    Call 保護.複数保護
    Worksheets("印刷他").Activate
    ActiveSheet.Range("A11").Select
End Sub

Sub 印刷_形成2()
    Call 保護.全保護解除
    Worksheets("形成2").Activate
    
    Call ソート形成2.ソート形成2_レシピ

    With ActiveSheet
        .Range("C5").Select
        .Range("C5").AutoFilter Field:=25, Criteria1:="<>"
    End With

    '// プリンタとの接続を切断
    Application.PrintCommunication = False
    '// 印刷設定
    With ActiveSheet.PageSetup
        .PrintArea = Sheets("形成2").Range("D4:V73").Address
        .Orientation = xlLandscape                          '印刷向きを横方向に設定
        .Zoom = False                                       '拡大縮小を設定（しない）
        .FitToPagesWide = 1                                 'すべての列を1ページに印刷
        .FitToPagesTall = False                             'シートを1ページに印刷

        .LeftMargin = Application.CentimetersToPoints(0.5)
        .RightMargin = Application.CentimetersToPoints(0.5)
        .TopMargin = Application.CentimetersToPoints(0.5)
        .BottomMargin = Application.CentimetersToPoints(1)
        .HeaderMargin = Application.CentimetersToPoints(0)
        .FooterMargin = Application.CentimetersToPoints(0)
    End With
    '// プリンタと接続する
    Application.PrintCommunication = True
    ActiveSheet.PrintOut
    Call 保護.複数保護
    Worksheets("印刷他").Activate
    ActiveSheet.Range("A11").Select
End Sub

Sub 印刷_CNまとめ()
    Call 保護.全保護解除
    Call 印刷_形成1
    Call 印刷_形成2
    Call 保護.複数保護
End Sub

Sub 印刷_賞味期限()
    Call 保護.全保護解除
    Worksheets("賞味期限").Activate
    
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
        .Range("A10").AutoFilter Field:=1, Criteria1:="<>"
    End With
    
    '// プリンタとの接続を切断
    Application.PrintCommunication = False
    '// 印刷設定
    With ActiveSheet.PageSetup
        .PrintArea = Sheets("賞味期限").Range("A8:J55").Address
        .Orientation = xlPortrait '印刷向きを縦方向に設定
        .Zoom = False '拡大縮小を設定（しない）
        .FitToPagesWide = 1 'すべての列を1ページに印刷
        .FitToPagesTall = False 'シートを1ページに印刷
        
        .LeftMargin = Application.CentimetersToPoints(0.5)
        .RightMargin = Application.CentimetersToPoints(0.5)
        .TopMargin = Application.CentimetersToPoints(0.5)
        .BottomMargin = Application.CentimetersToPoints(0.5)
        .HeaderMargin = Application.CentimetersToPoints(0)
        .FooterMargin = Application.CentimetersToPoints(0)
    End With
    '// プリンタと接続する
    Application.PrintCommunication = True
    ActiveSheet.PrintOut
    Call 保護.複数保護
End Sub

Sub 印刷_シール()
    Call 保護.全保護解除
    Worksheets("印刷他").Activate
    
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
        .ListObjects("テーブル2").Range.AutoFilter Field:=38, Criteria1:="シール"
        .ListObjects("テーブル2").Range.AutoFilter Field:=6, Criteria1:="<>-"
        .ListObjects("テーブル5").Range.AutoFilter Field:=1, Criteria1:="シール"
    End With
    
    Call ソート印刷他.ソート仕入先名
    
    '// プリンタとの接続を切断
    Application.PrintCommunication = False
    '// 印刷設定
    With ActiveSheet.PageSetup
        .PrintArea = Sheets("印刷他").Range("A8:J9026").Address
        .Orientation = xlPortrait '印刷向きを縦方向に設定
        .Zoom = False '拡大縮小を設定（しない）
        .FitToPagesWide = 1 'すべての列を1ページに印刷
        .FitToPagesTall = False 'シートを1ページに印刷
        
        .LeftMargin = Application.CentimetersToPoints(0.5)
        .RightMargin = Application.CentimetersToPoints(0.5)
        .TopMargin = Application.CentimetersToPoints(0.5)
        .BottomMargin = Application.CentimetersToPoints(0.5)
        .HeaderMargin = Application.CentimetersToPoints(0)
        .FooterMargin = Application.CentimetersToPoints(0)
    End With
    '// プリンタと接続する
    Application.PrintCommunication = True
    ActiveSheet.PrintOut
    Call 保護.複数保護
    ActiveSheet.Range("A11").Select
End Sub

Sub 印刷_トレー()
    Call 保護.全保護解除
    Worksheets("トレー").Activate

    '// プリンタとの接続を切断
    Application.PrintCommunication = False
    '// 印刷設定
    With ActiveSheet.PageSetup
        '.PrintArea = Sheets("印刷CN").Range("D4:V183").Address
        .Orientation = xlPortrait                                  '印刷向きを縦方向に設定
        .Zoom = False                                               '拡大縮小を設定（しない）
        .FitToPagesWide = 1                                         'すべての列を1ページに印刷
        .FitToPagesTall = False                                     'シートを1ページに印刷
        
        .LeftMargin = Application.CentimetersToPoints(1)
        .RightMargin = Application.CentimetersToPoints(1)
        .TopMargin = Application.CentimetersToPoints(1)
        .BottomMargin = Application.CentimetersToPoints(1)
        .HeaderMargin = Application.CentimetersToPoints(1)
        .FooterMargin = Application.CentimetersToPoints(0)
    End With
    '// プリンタと接続する
    Application.PrintCommunication = True
    ActiveSheet.PrintOut
    Call 保護.複数保護
End Sub

Sub 印刷_小分け()
    '期末棚卸専用のボタン 2020.09.13コメント
    '詳細は「フィルター_印刷他シート」の「フィルター小分け」関数にコメントで書いてある

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
    
    '// プリンタとの接続を切断
    Application.PrintCommunication = False
    '// 印刷設定
    With ActiveSheet.PageSetup
        .PrintArea = Sheets("印刷他").Range("A8:J9016").Address
        .Orientation = xlPortrait '印刷向きを縦方向に設定
        .Zoom = False '拡大縮小を設定（しない）
        .FitToPagesWide = 1 'すべての列を1ページに印刷
        .FitToPagesTall = False 'シートを1ページに印刷
        
        .LeftMargin = Application.CentimetersToPoints(0.5)
        .RightMargin = Application.CentimetersToPoints(0.5)
        .TopMargin = Application.CentimetersToPoints(0.5)
        .BottomMargin = Application.CentimetersToPoints(0.5)
        .HeaderMargin = Application.CentimetersToPoints(0)
        .FooterMargin = Application.CentimetersToPoints(0)
    End With
    '// プリンタと接続する
    Application.PrintCommunication = True
    ActiveSheet.PrintOut
    
    'Call 印刷_賞味期限
    Call 保護.複数保護
    Worksheets("印刷他").Activate
    ActiveSheet.Range("A11").Select
End Sub
