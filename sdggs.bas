
Sub 一覧データ()
    Dim datalist_3() As Variant
    ReDim datalist_3(1000, 2)
    
    Worksheets("一覧").Activate

    A_LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    m = 0

    For s = 6 To A_LastRow
        If Cells(s, 6) <> "" And Cells(s, 2) = Int(Cells(s, 2)) Then
            For u = 1 To 5
                search_data_1 = Trim(Cells(s, 6) & u)
                Worksheets("csv形成").Activate
                N_LastRow = Cells(Rows.Count, 14).End(xlUp).Row
                For t = 1 To N_LastRow
                    If IsNumeric(Cells(t, 14)) Then
                        b_data = Trim(Str(Cells(t, 14)))
                    Else
                        b_data = Cells(t, 14)
                    End If
                    
                    If search_data_1 = b_data Then
                        if left(search_data_1,1)=1 Then
                            datalist_3(m, 0) = Cells(t, 15)
                            datalist_3(m, 1) = Cells(t, 16)
                            datalist_3(m, 2) = Cells(t, 14)
                            m = m + 5
                            Worksheets("一覧").Activate
                            Exit For
                        ElseIf left(search_data_1,1)=2 Then
                            datalist_3(m+1, 0) = Cells(t, 15)
                            datalist_3(m+1, 1) = Cells(t, 16)
                            datalist_3(m+1, 2) = Cells(t, 14)
                            m = m + 4
                            Worksheets("一覧").Activate
                            Exit For
                        ElseIf left(search_data_1,1)= 3 Then
                            datalist_3(m+2, 0) = Cells(t, 15)
                            datalist_3(m+2, 1) = Cells(t, 16)
                            datalist_3(m+2, 2) = Cells(t, 14)
                            m = m + 3
                            Worksheets("一覧").Activate
                            Exit For
                        ElseIf left(search_data_1,1)=4 Then
                            datalist_3(m+3, 0) = Cells(t, 15)
                            datalist_3(m+3, 1) = Cells(t, 16)
                            datalist_3(m+3, 2) = Cells(t, 14)
                            m = m + 2
                            Worksheets("一覧").Activate
                            Exit For
                        ElseIf left(search_data_1,1)=5 Then
                        datalist_3(m+4, 0) = Cells(t, 15)
                            datalist_3(m+4, 1) = Cells(t, 16)
                            datalist_3(m+4, 2) = Cells(t, 14)
                            m = m + 1
                            Worksheets("一覧").Activate
                            Exit For
                        End If
                    End If                            
                Next t
            Next u
        End If
    Next s

    
    Worksheets("Sheet3").Activate

    Worksheets("Sheet3").Range(Cells(1, 1), Cells(1000, 3)) = datalist_3
    

End Sub