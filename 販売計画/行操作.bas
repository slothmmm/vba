
Sub 最終行追加()
    'Application.Calculation = xlCalculationManual     '手動計算

    Cells(ActiveCell.Row, ActiveCell.Column).Select '単体選択

'    Dim rc As VbMsgBoxResult
'    rc = MsgBox("行追加しますか？", vbYesNo + vbQuestion)
'    If rc = vbNo Then
'        MsgBox "中止します", vbCritical
'        Exit Sub
'    End If
    
    Dim Last_Row As Long
    Last_Row = Cells(Rows.Count, 68).End(xlUp).Row
    Last_Row = Last_Row + 1
    
    Call 保護.保護解除
    '２行コピー
     Rows(Last_Row - 2 & ":" & Last_Row - 1).Select
    Rows(Last_Row - 2 & ":" & Last_Row - 1).Copy
    
'    Application.ScreenUpdating = False                  '画面停止

    'Rows("Last_Row - 2:Last_Row - 2").Select
    '２行挿入
    'Selection.Insert Shift:=xlDown

    Cells(Last_Row, 1).EntireRow.Insert
'    Application.ScreenUpdating = True                   '画面稼働
    Call 保護.保護

    '注釈、商品名、１～３１日、AL～ANの単価など削除
    'Range(Cells(Last_Row, 2), Cells(Last_Row, 34)).ClearContents
    Range(Cells(Last_Row, 1), Cells(Last_Row, 3)).ClearContents
    Range(Cells(Last_Row, 5), Cells(Last_Row, 35)).ClearContents
    Range(Cells(Last_Row, 37), Cells(Last_Row, 37)).ClearContents
    Range(Cells(Last_Row, 39), Cells(Last_Row, 39)).ClearContents
    Range(Cells(Last_Row, 47), Cells(Last_Row, 47)).ClearContents
    
    '新コード入力
    'Cells(Last_Row, 1).Value = new_code
    
    '商品名入力
    'Cells(Last_Row, 3).Value = product_name
    
    'Application.Calculation = xlCalculationAutomatic  '自動計算

End Sub

Sub 行削除()

    Cells(ActiveCell.Row, ActiveCell.Column).Select '単体選択
    
    Dim Last_Row As Long
    Last_Row = Cells(Rows.Count, 68).End(xlUp).Row
    Last_Row = Last_Row + 1
    
    BPcells = Range(Cells(1, 68), Cells(Last_Row, 68)).Formula
    BP_start = 0
    
    For i = 1 To Last_Row
        If BPcells(i, 1) Like "*エラー*" Then
            BP_start = i
            Exit For
        End If
        If i = Last_Row Then
            MsgBox ("エラーです。BP列がおかしいです。")
            Exit Sub
        End If
    Next i
    
    act_row = ActiveCell.Row
    act_col = ActiveCell.Column

    If act_row < BP_start Or act_row > Last_Row - 1 Then
        MsgBox (Str(BP_start) + "行から" + Str(Last_Row) + "行の間のセルを選択して下さい。")
        Exit Sub
    End If
    
    '該当のシートがアクティブになることが前提
    'Application.Calculation = xlCalculationManual     '手動計算
'    'Application.Calculation = xlCalculationAutomatic  '自動計算
'    Application.ScreenUpdating = False                  '画面停止

    If Cells(act_row, 68).Formula Like "*エラー*" Then
        Debug.Print "OK"
    ElseIf Cells(act_row - 1, 68).Formula Like "*エラー*" Then
        act_row = act_row - 1
        Debug.Print "OK2"
    Else
        Debug.Print "NG"
        MsgBox ("エラーです。BP列がおかしいです。")
        Exit Sub
    End If
    
    '２行選択
    Rows(act_row & ":" & act_row + 1).Select
    
    Dim rc As VbMsgBoxResult
    rc = MsgBox("選択された行を削除しますか？", vbYesNo + vbQuestion)
    If rc = vbNo Then
        MsgBox "中止します", vbCritical
        Exit Sub
    Else
        Call 保護.保護解除
        '削除
        Rows(act_row & ":" & act_row + 1).Delete
         Call 保護.保護
    End If
    
    'Application.Calculation = xlCalculationAutomatic  '自動計算
    Application.ScreenUpdating = True                   '画面稼働
   
End Sub
Sub 行追加()

    Cells(ActiveCell.Row, ActiveCell.Column).Select '単体選択
    
    Dim Last_Row As Long
    Last_Row = Cells(Rows.Count, 68).End(xlUp).Row
    Last_Row = Last_Row + 1
    
    BPcells = Range(Cells(1, 68), Cells(Last_Row, 68)).Formula
    BP_start = 0
    
    For i = 1 To Last_Row
        If BPcells(i, 1) Like "*エラー*" Then
            BP_start = i
            Exit For
        End If
        If i = Last_Row Then
            MsgBox ("エラーです。BP列がおかしいです。")
            Exit Sub
        End If
    Next i
    
    act_row = ActiveCell.Row
    act_col = ActiveCell.Column

    If act_row < BP_start Or act_row > Last_Row - 1 Then
        MsgBox (Str(BP_start) + "行から" + Str(Last_Row) + "行の間のセルを選択して下さい。")
        Exit Sub
    End If
    
    '該当のシートがアクティブになることが前提
    'Application.Calculation = xlCalculationManual     '手動計算
'    'Application.Calculation = xlCalculationAutomatic  '自動計算
'    Application.ScreenUpdating = False                  '画面停止

    If Cells(act_row, 68).Formula Like "*エラー*" Then
        Debug.Print "OK"
    ElseIf Cells(act_row - 1, 68).Formula Like "*エラー*" Then
        act_row = act_row - 1
        Debug.Print "OK2"
    Else
        Debug.Print "NG"
        MsgBox ("エラーです。BP列がおかしいです。")
        Exit Sub
    End If
    
    '２行選択
    Rows(act_row & ":" & act_row + 1).Select
    
    Dim rc As VbMsgBoxResult
    rc = MsgBox("選択された行へ行追加しますか？", vbYesNo + vbQuestion)
    If rc = vbNo Then
        MsgBox "中止します", vbCritical
        Exit Sub
    Else
        Call 保護.保護解除
        '追加
        Rows(act_row & ":" & act_row + 1).Copy
        Cells(act_row + 2, 1).EntireRow.Insert
        
        '注釈、商品名、１～３１日、AL～ANの単価など削除
        'Range(Cells(Last_Row, 2), Cells(Last_Row, 34)).ClearContents
        Range(Cells(act_row, 1), Cells(act_row, 3)).ClearContents
        Range(Cells(act_row, 5), Cells(act_row, 35)).ClearContents
        Cells(act_row, 37) = ""
        Cells(act_row, 39) = ""
        Cells(act_row, 47) = ""
          
        
         Call 保護.保護
    End If
    
    'Application.Calculation = xlCalculationAutomatic  '自動計算
    Application.ScreenUpdating = True                   '画面稼働
   
End Sub
