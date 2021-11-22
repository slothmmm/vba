Sub 重複なしリスト()
    Dim C As New Collection, i As Long
    
    On Error Resume Next
    For i = 31 To Cells(Rows.Count, 3).End(xlUp).Row
        C.Add Cells(i, 3), Cells(i, 3)
    Next i
    
    On Error GoTo 0
    For i = 1 To C.Count
        Debug.Print C(i)
    Next i
End Sub