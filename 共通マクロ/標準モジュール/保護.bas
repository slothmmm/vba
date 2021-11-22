Attribute VB_Name = "保護"

Sub 保護()
    ActiveSheet.Protect Contents:=True, DrawingObjects:=False, UserInterfaceOnly:=True, AllowFormattingCells:=True
End Sub

Sub 保護解除()
    ActiveSheet.Unprotect
End Sub

Sub 保護_複数()
    For ws_num = 1 To (Sheets("合計金額").Index - 1)
        Worksheets(ws_num).Protect Contents:=True, DrawingObjects:=False, UserInterfaceOnly:=True, AllowFormattingCells:=True
    Next ws_num
    
End Sub

Sub 保護_全解除()
    Dim sh As Object
    On Error Resume Next
    For Each sh In Sheets
    sh.Unprotect
    Next sh
    
    Worksheets("合計金額").Activate
End Sub

Sub シート複数選択解除()
    ActiveWindow.SelectedSheets(1).Select
End Sub

Sub 保護ロック全シート循環()
    Call 在庫表作成_エクセル.画面とアラート非表示
    'シート循環　合計金額シートよりひだり
    For ws_num = 1 To (Sheets("合計金額").Index - 1)

        '隠れシートは飛ばす
        If Not Worksheets(ws_num).Visible Then
            GoTo CONTINUE
        End If

        '「棚卸表」「原料展開」シート飛ばす
        If Worksheets(ws_num).Name = "棚卸表" Or Worksheets(ws_num).Name = "原料展開" Then
            GoTo CONTINUE
        End If

        '表示シートアクティブ
        Worksheets(ws_num).Activate
        
        '実行
        Call 保護ロック解除obj
CONTINUE:
    Next ws_num
    Call 在庫表作成_エクセル.画面とアラート表示
    
    MsgBox "保護ロック全シート循環 終了しました。"
End Sub

Sub teeee()
    Debug.Print ActiveSheet.Name
End Sub


Sub 保護ロック解除obj()
    '全セル保護
     Cells.Locked = True

    '最終行取得
    Dim LastRow As Long
    LastRow = Cells(Rows.Count, 5).End(xlUp).Row

    'E列検索用
    serch_cell = Range(Cells(1, 5), Cells(LastRow, 9)).Formula
    

    
    '１～３１日
    Dim tikan_cell As Range
    Set tikan_cell = Range(Cells(1, 8), Cells(LastRow, 38))
    
    'サイキ用
    If ActiveSheet.Name = "サイキ食品㈱" Then
        'A列検索用
        Debug.Print "サイキ"
        A_serch_cell = Range(Cells(1, 1), Cells(LastRow, 9))
        For sss = 1 To LastRow
            If A_serch_cell(sss, 1) = "2557" Then
                Debug.Print "サイキ下ロック解除"
                For nnn = 1 To 31
                    tikan_cell(sss, nnn).Locked = False
                Next nnn
            End If
        Next sss
    End If
    
    '該当セルをロック解除
    For s = 1 To LastRow
        If serch_cell(s, 1) = "入荷数" Or _
            serch_cell(s, 1) = "合計入荷数" Or _
            serch_cell(s, 1) = "出荷数(手入力)" Or _
            serch_cell(s, 1) = "服部コーヒー" Or _
            serch_cell(s, 1) = "サポート" Or _
            serch_cell(s, 1) = "ヨネヤマ" Or _
            serch_cell(s, 1) = "返品等" Or _
            serch_cell(s, 1) = "預け" Or _
            serch_cell(s, 1) = "戻し" _
            Then
            
            Debug.Print ActiveSheet.Name & "___保護ロック解除するセル選択__" & s & "行"
            
            For n = 1 To 31
                tikan_cell(s, n).Locked = False
            Next n
            
        End If
        
        If serch_cell(s, 1) = "調整" Or _
            serch_cell(s, 1) = "調整1" Or _
            serch_cell(s, 1) = "調整2" _
            Then
                
            '数式が入っているかどうか
            If Mid(serch_cell(s, 4), 1, 1) = "=" Then
                Debug.Print ActiveSheet.Name & "___保護ロック解除するセル選択しない__" & s & "行"
            Else
                Debug.Print ActiveSheet.Name & "___保護ロック解除するセル選択__調整__" & s & "行"
                
                For n = 1 To 31
                    tikan_cell(s, n).Locked = False
                Next n
            End If
        End If
    Next s
    
End Sub
