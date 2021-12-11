Sub 条件付き書式_main()
    Call 条件付き書式の削除と追加
End Sub

Sub 条件付き書式_1シート追加()
    Call 条件付き書式_01_日曜日か
    Call 条件付き書式_02_数式あるか
    Call 条件付き書式_03_空白か
    Call 条件付き書式_04_該当商品無しか
    Call 条件付き書式_05_その月に出荷あるか
    Call 条件付き書式_06_終売か
End Sub

Sub 条件付き書式の削除と追加()
    'これ実行する前に全シート保護解除a
    Call 保護_全解除
  Dim ws As Worksheet
  For Each ws In Worksheets
    For i = 1 To 12
        aa = Trim(Str(i) + "月")    '何故か空白が入るので
        If ws.Name = aa Then
            ws.Cells.FormatConditions.Delete  '条件付き書式の削除
            ws.Activate
            Call 条件付き書式_1シート追加
        End If
    Next i
  Next ws
  Call 保護_複数
End Sub

Sub 条件付き書式_01_日曜日か()
    Dim fc As FormatCondition
      Set fc = Range("$G$1:$AK$2").FormatConditions.Add(Type:=xlExpression, Formula1:="=TEXT(G$2,""aaa"")=""日""")
      fc.Interior.Color = RGB(0, 0, 255)
      'fc.Font.Color = RGB(255, 255, 255)
End Sub

Sub 条件付き書式_02_数式あるか()
    Dim fc As FormatCondition
      Set fc = Range("$G4:$AK9999").FormatConditions.Add(Type:=xlExpression, Formula1:="=ISFORMULA(INDIRECT(""RC"", false))")
      fc.Interior.Color = RGB(255, 204, 153)
End Sub

Sub 条件付き書式_03_空白か()
    Dim fc As FormatCondition
      Set fc = Range("$G4:$AK9999").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""""")
      fc.Interior.Color = RGB(255, 255, 255)
End Sub

Sub 条件付き書式_04_該当商品無しか()
    Dim fc As FormatCondition
      Set fc = Range("$D:$D").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""該当商品無し""")
      fc.Interior.Color = RGB(255, 255, 0)
End Sub

Sub 条件付き書式_05_その月に出荷あるか()
    Dim fc As FormatCondition
      Set fc = Range("$E:$F").FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0")
      fc.Interior.Color = RGB(0, 128, 128)
End Sub

Sub 条件付き書式_06_終売か()
    Dim fc As FormatCondition
      Set fc = Range("$C:$D").FormatConditions.Add(Type:=xlExpression, Formula1:="=$A1=""×""")
      fc.Interior.Color = RGB(51, 51, 51)
      fc.Font.Color = RGB(255, 255, 255)
End Sub

