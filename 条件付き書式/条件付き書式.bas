Sub 条件付き書式_main()
    Application.Calculation = xlCalculationManual     '手動計算
    Call 条件付き書式の削除と追加
    Application.Calculation = xlCalculationAutomatic  '自動計算
End Sub

Sub 条件付き書式_1シート追加()
    Call 条件付き書式_10_新規か
    Call 条件付き書式_11_終売か
    Call 条件付き書式_01_未登録か
    Call 条件付き書式_02_仮か
    Call 条件付き書式_03_本か
    Call 条件付き書式_04_空か
    Call 条件付き書式_05_エラーか
    Call 条件付き書式_06_○か
    Call 条件付き書式_07_Xか
    Call 条件付き書式_08_0以上か
    Call 条件付き書式_09_日付

End Sub

Sub 条件付き書式の削除と追加()
    Call 保護_全解除     'これ実行する前に全シート保護解除a
    Dim ws As Worksheet
    For Each ws In Worksheets
        For i = 1 To 12
            aa = Trim(Str(i) + "月")     '何故か空白が入るので
            If ws.Name = aa Then
                ws.Cells.FormatConditions.Delete     '条件付き書式の削除
                ws.Activate
                Call 条件付き書式_1シート追加
            End If
        Next i
    Next ws
    Call 保護_複数
End Sub

Sub 条件付き書式_01_未登録か()
    Dim fc As FormatCondition
       Set fc = Range("$A:$D").FormatConditions.Add(Type:=xlExpression, Formula1:="=INDIRECT(""RC"",FALSE)=""未登録""")
       fc.Interior.Color = RGB(255, 255, 0)
'       fc.Font.Color = RGB(0, 0, 0)
End Sub

Sub 条件付き書式_02_仮か()
    Dim fc As FormatCondition
       Set fc = Range("$A:$D").FormatConditions.Add(Type:=xlExpression, Formula1:="=$BP1=""仮""")
       fc.Interior.Color = RGB(13, 13, 13)
       fc.Font.Color = RGB(255, 255, 255)
End Sub

Sub 条件付き書式_03_本か()
    Dim fc As FormatCondition
       Set fc = Range("$A:$D").FormatConditions.Add(Type:=xlExpression, Formula1:="=$BP1=""本""")
       fc.Interior.Color = RGB(252, 213, 180)
       fc.Font.Color = RGB(0, 0, 0)
End Sub

Sub 条件付き書式_04_空か()
    Dim fc As FormatCondition
       Set fc = Range("$A:$D").FormatConditions.Add(Type:=xlExpression, Formula1:="=$BP1=""空""")
       fc.Interior.Color = RGB(255, 255, 255)
       fc.Font.Color = RGB(0, 0, 0)
End Sub

Sub 条件付き書式_05_エラーか()
    Dim fc As FormatCondition
       Set fc = Range("$A:$D").FormatConditions.Add(Type:=xlExpression, Formula1:="=$BP1=""エラー""")
       fc.Interior.Color = RGB(255, 0, 0)
       fc.Font.Color = RGB(255, 255, 255)
End Sub

Sub 条件付き書式_06_○か()
    Dim fc As FormatCondition
      Set fc = Range("$BD:$BH").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""○""")
      fc.Interior.Color = RGB(0, 0, 255)
      fc.Font.Color = RGB(255, 255, 255)
End Sub

Sub 条件付き書式_07_Xか()
    Dim fc As FormatCondition
      Set fc = Range("$BD:$BH").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""X""")
      fc.Interior.Color = RGB(255, 0, 0)
      fc.Font.Color = RGB(255, 255, 255)
End Sub

Sub 条件付き書式_08_0以上か()
    Dim fc As FormatCondition
       Set fc = Range("AY1:AY4").FormatConditions.Add(Type:=xlExpression, Formula1:="=AY1>=1")
       fc.Interior.Color = RGB(255, 0, 0)
       fc.Font.Color = RGB(255, 255, 255)
End Sub

Sub 条件付き書式_09_日付()
    Dim fc As FormatCondition
       Set fc = Range("$AG$8:$AI$8").FormatConditions.Add(Type:=xlExpression, Formula1:="=AG8=""""")
       fc.Interior.Color = RGB(0, 0, 0)
End Sub

Sub 条件付き書式_10_新規か()
    Dim fc As FormatCondition
       Set fc = Range("$A:$D").FormatConditions.Add(Type:=xlExpression, Formula1:="=$BQ1=""新規""")
       fc.Interior.Color = RGB(255, 0, 0)
End Sub

Sub 条件付き書式_11_終売か()
    Dim fc As FormatCondition
       Set fc = Range("$A:$D").FormatConditions.Add(Type:=xlExpression, Formula1:="=$BQ1=""終売""")
       fc.Interior.Color = RGB(146, 208, 80)
End Sub



