Sub フォント調整()
    Worksheets("チェックシート").Activate
    Worksheets("チェックシート").Select
    Range("A1").Select
    ActiveSheet.Unprotect
    
    B_LastRow = Cells(Rows.Count, 1).End(xlUp).Row  ''B列の最終行取得
    read_Col = 11                                   'L列まで
    
    Dim sheetData As Variant
     
    sheetData = Worksheets("チェックシート").Range(Cells(1, 1), Cells(B_LastRow, read_Col))   'シートデータ取得
    
    
    For i = 1 To UBound(sheetData)
        If Cells(i, 2).Value = "原材料名" Then
            g = LenB(StrConv(Cells(i + 1, 2).Value, vbFromUnicode))
                If g > 348 Then
                    Cells(i + 1, 2).Font.Size = 6.5
                Else
                    Cells(i + 1, 2).Font.Size = 9
                End If
         End If
    Next i
End Sub
