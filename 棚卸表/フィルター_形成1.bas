Attribute VB_Name = "フィルター_形成1"
Sub フィルター全部クリア_形成1シート()
    Call 保護.全保護解除
    Worksheets("形成1").Activate
    With ActiveSheet
        .Range("C5").Select
        If .FilterMode Then .ShowAllData
        
        .Range("C37").Select
        If .FilterMode Then .ShowAllData
        
        .Range("C69").Select
        If .FilterMode Then .ShowAllData
        
        .Range("C101").Select
        If .FilterMode Then .ShowAllData
        
        .Range("C163").Select
        If .FilterMode Then .ShowAllData
        
        .Range("C5").Select
    End With
    Call 保護.複数保護
    MsgBox "フィルタークリア完了(形成1シート)"
End Sub

Sub フィルター_形成1()
    Call 保護.全保護解除
    Worksheets("形成1").Activate
    
    With ActiveSheet
        .ListObjects("新館").Range.AutoFilter Field:=25, Criteria1:="<>"
        .ListObjects("商品管理").Range.AutoFilter Field:=25, Criteria1:="<>"
        .ListObjects("冷蔵庫").Range.AutoFilter Field:=25, Criteria1:="<>"
        .ListObjects("冷凍庫").Range.AutoFilter Field:=25, Criteria1:="<>"
        .ListObjects("その他").Range.AutoFilter Field:=25, Criteria1:="<>"
    End With
    Call 保護.複数保護
    MsgBox "フィルタークリア完了(形成1シート)"
End Sub
    
