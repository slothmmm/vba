Sub 条件付き書式()
    ThisWorkbook.Activate
    Worksheets("sheetname").Activate
    
    Dim fc As FormatCondition
    Set fc = Range("$G:$L").FormatConditions.Add(Type:=xlExpression, Formula1:="=$X1=""白塗り""")
    fc.Font.Color = RGB(255, 255, 255)　'白
    fc.Interior.Color = RGB(0, 0, 0)    '黒
End Sub