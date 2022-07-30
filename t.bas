Sub フィルター_一括()
    Call フィルター_D週間
    Call フィルター_Dカウンター
    Call フィルター_U週間
    Call フィルター_Uカウンター
End Sub

Sub フィルター_D週間()
    Call 保護.全保護解除
    Worksheets("D週間").Activate
    
    With ActiveSheet
        .Range("G9").Select
        If .FilterMode Then .ShowAllData
        .Range("G9").AutoFilter Field:=28, Criteria1:="<>"
    End With
    
    Call 保護.複数保護
End Sub

Sub フィルター_U週間()
    Call 保護.全保護解除
    Worksheets("U週間").Activate
    
    With ActiveSheet
        .Range("G9").Select
        If .FilterMode Then .ShowAllData
        .Range("G9").AutoFilter Field:=28, Criteria1:="<>"
    End With
    
    Call 保護.複数保護
End Sub

Sub フィルター_Dカウンター()
    Call 保護.全保護解除
    Worksheets("Dカウンター").Activate
    
    With ActiveSheet
        .Range("G9").Select
        If .FilterMode Then .ShowAllData
        .Range("G9").AutoFilter Field:=26, Criteria1:="<>"
    End With
    
    Call 保護.複数保護
End Sub

Sub フィルター_Uカウンター()
    Call 保護.全保護解除
    Worksheets("Uカウンター").Activate
    
    With ActiveSheet
        .Range("G9").Select
        If .FilterMode Then .ShowAllData
        .Range("G9").AutoFilter Field:=26, Criteria1:="<>"
    End With
    Call 保護.複数保護
End Sub

