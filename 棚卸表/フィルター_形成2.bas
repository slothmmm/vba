Attribute VB_Name = "フィルター_形成2"
Sub フィルター全部クリア_形成2シート()
    Call 保護.全保護解除
    Worksheets("形成2").Activate
    With ActiveSheet
        .Range("C5").Select
        If .FilterMode Then .ShowAllData
        
    End With
    Call 保護.複数保護
    MsgBox "フィルタークリア完了(形成2シート)"
End Sub

Sub フィルター_形成2()
    Call 保護.全保護解除
    Worksheets("形成2").Activate
    
    With ActiveSheet
        .Range("C5").Select
        .Range("C5").AutoFilter Field:=25, Criteria1:="<>"
    End With
    Call 保護.複数保護
    MsgBox "フィルタークリア完了(形成2シート)"
End Sub
