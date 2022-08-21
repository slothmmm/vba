Sub 一覧入力()
    Worksheets("csv形成").Activate
    
    Dim datalist_1() As Variant
    ReDim datalist_1(1000, 0)
    
    Dim datalist_2() As Variant
    ReDim datalist_2(1000, 0)
    
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    k = 0
    

    
    
    
    For r = 2 To LastRow
        If Cells(r, 12) = Worksheets("一覧").Range("H1") Then
            If Cells(r, 7) = "通常" Then
                If Right(Cells(r, 1), 2) = ".1" Then
                    search_data = Trim(Cells(r, 1))
                    Worksheets("小分け品").Activate
                    kowakehin_LastRow = Cells(Rows.Count, 1).End(xlUp).Row
                    For i = 1 To kowakehin_LastRow
                        If search_data = Cells(i, 1) Then
                            datalist_1(k, 0) = Cells(i, 2)
                            datalist_1(k + 1, 0) = Cells(i + 1, 2)
                            datalist_1(k + 2, 0) = Cells(i + 2, 2)
                            datalist_1(k + 3, 0) = Cells(i + 3, 2)
                            datalist_1(k + 4, 0) = Cells(i + 4, 2)
                            Worksheets("csv形成").Activate
                            k = k + 5
                        End If
                    Next i
                End If
            ElseIf Cells(r, 7) = "煮物" Then
                If Right(Cells(r, 1), 2) = ".1" Then
                    search_data = Trim(Cells(r, 1))
                    Worksheets("小分け品").Activate
                    kowakehin_LastRow = Cells(Rows.Count, 1).End(xlUp).Row
                    For i = 1 To kowakehin_LastRow
                        If search_data = Cells(i, 1) Then
                            datalist_2(k, 0) = Cells(i, 2)
                            datalist_2(k + 1, 0) = Cells(i + 1, 2)
                            datalist_2(k + 2, 0) = Cells(i + 2, 2)
                            datalist_2(k + 3, 0) = Cells(i + 3, 2)
                            datalist_2(k + 4, 0) = Cells(i + 4, 2)
                            Worksheets("csv形成").Activate
                            k = k + 5
                        End If
                    Next i
                End If
            End If
        End If
    Next r
    
    Worksheets("Sheet3").Activate
    
    Worksheets("Sheet3").Range(Cells(1, 1), Cells(1000, 1)) = datalist_1
    Worksheets("Sheet3").Range(Cells(1, 2), Cells(1000, 2)) = datalist_2
    
    
End Sub
