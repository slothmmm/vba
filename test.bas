Sub データコピー()

  On Error Resume Next
  
    Dim このブック, 対象ブック, MYキー As String, セル位置 As Range, 使用数 As Variant
    
    このブック = ActiveWorkbook.Name
      対象ブック = "生産表.xlsm"
        MYキー = Sheets("抽出2").Range("X3").Value
          使用数 = Sheets("抽出2").Range("Y3").Value
        
    Workbooks(対象ブック).Activate
      Sheets("手入力").Select
        Set セル位置 = Range("A:A").Find(What:=MYキー, LookIn:=xlValues, _
          LookAt:=xlWhole, SearchOrder:=xlByRows)
            If (セル位置 Is Nothing) Then
              If Len(MYキー) = 4 Then
                Workbooks(このブック).Worksheets("抽出2").Range("P22"). _
                  Formula = Workbooks(このブック).Worksheets("抽出2").Range("P22") & _
                    "/" & Date & ">" & MYキー
                      End If
                        GoTo 最終
                          End If
                
    Cells(セル位置.Row, セル位置.Column + 3).Formula = Cells(セル位置.Row, セル位置.Column + 3) + 使用数

最終:
    Sheets("受注入力").Select
      Workbooks(このブック).Activate
      
  On Error GoTo 0
    
End Sub