Private Sub Workbook_Open()
    Worksheets("未登録商品一覧").Visible = True
    first_sheet = ActiveSheet.Name
    Worksheets("未登録商品一覧").Activate
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



