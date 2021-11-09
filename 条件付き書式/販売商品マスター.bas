Sub 条件付き書式_main()
    Application.Calculation = xlCalculationManual     '手動計算
    Call 条件付き書式_商品マスター
    Call 条件付き書式_部品マスター
    Application.Calculation = xlCalculationAutomatic  '自動計算
End Sub

Sub 条件付き書式_商品マスター()
    Dim ws As Worksheet
    If ws.Name = "商品マスター" Then
        ws.Cells.FormatConditions.Delete    '条件付き書式の削除
        Call 条件付き書式_01_重複か
        Call 条件付き書式_02_終売か
    End If
End Sub

Sub 条件付き書式_部品マスター()
    
End Sub

Sub 条件付き書式_01_重複か()
    Dim fc As FormatCondition
       Set fc = Range("$F:$F").FormatConditions.Add(Type:=xlExpression, Formula1:="=$GI5=""重複""")
       fc.Interior.Color = RGB(255, 0, 0)
End Sub

Sub 条件付き書式_02_終売か()
    Dim fc As FormatCondition
        Set fc = Range("$A:$GS").FormatConditions.Add(Type:=xlExpression, Formula1:="=$FA5=""×""")
        fc.Interior.Color = RGB(89, 89, 89)
        fc.Font.Color = RGB(191, 191, 191)
        fc.Interior.Pattern = xlPatternUp
        fc.Interior.Pattern.Color = RGB(38, 38, 38)
End Sub
