Attribute VB_Name = "フィルター_印刷CN"
'''''''''''''''''''''''''''''''      フィルター関連                   ''''''''''''''''''''''''''''

Sub フィルター全部クリア_印刷CNシート()
    Call 保護.全保護解除
    Worksheets("印刷CN").Activate
    With ActiveSheet
        .Range("B4").Select
        If .FilterMode Then .ShowAllData
    End With
    Call 保護.複数保護
    MsgBox "フィルタークリア完了(CN判定)"
End Sub

Sub フィルターCN_印刷CNシート()
    Call 保護.全保護解除
    Worksheets("印刷CN").Activate
    With ActiveSheet
        .Range("B4").Select
        If .FilterMode Then .ShowAllData
        .Range("B4").AutoFilter Field:=26, Criteria1:="<>"
    End With
    Call 保護.複数保護
    MsgBox "フィルター完了(CN判定)"
End Sub

