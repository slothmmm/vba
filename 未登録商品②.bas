Private Sub Workbook_Open()
    Worksheets("未登録商品一覧").Visible = True
    first_sheet = ActiveSheet.Name
    Worksheets("未登録商品一覧").Activate

'画面操作無効
    'Application.Interactive = False
　　'Application.EnableCancelKey = xlDisabled

'If Worksheets("未登録商品一覧").Range("F12").Value = "エラー" Then
        'MsgBox  "開いた時に再計算がされていないから" & vbCrLf & "未登録商品の取得ができません" & vbCrLf & vbCrLf & "もう一度再計算しますか？", vbYesNo + vbQuestion
            'Dim rc As VbMsgBoxResult
            'rc = MsgBox("開いた時に再計算がされていないから" & vbCrLf & "未登録商品の取得ができません" & vbCrLf & vbCrLf & "もう一度再計算しますか？", vbYesNo + vbQuestion)
            'If rc = vbYes Then
                'Worksheets("未登録商品一覧")を再計算
                'Worksheets("未登録商品一覧").Calculate
            'Else
                'MsgBox "処理を中止します", vbCritical
            'End If

'Debug.Print Worksheets("未登録商品一覧").Range("R53").Value
    If Worksheets("未登録商品一覧").Range("R53").Value <> "" Then
        unregistered_list = Worksheets("未登録商品一覧").Range(Cells(12, 2), Cells(41, 6))
        Dim msg1 As String
        
        For i = 1 To 30
            
            If unregistered_list(i, 2) <> "" Then
                msg1 = msg1 & unregistered_list(i, 2) & " " & unregistered_list(i, 4) & " 残り" & Str(unregistered_list(i, 5)) & "日" & vbCrLf
                
            End If
        Next
        MsgBox "出荷開始日が間近の未登録商品" & vbCrLf & vbCrLf & msg1, vbExclamation
     End If

     
     
     
     Worksheets(first_sheet).Activate
     Worksheets("未登録商品一覧").Visible = False
End Sub