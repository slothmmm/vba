Sub óÒïù()
    Application.ScreenUpdating = False
    For i = 2 To 97
        Columns(i).ColumnWidth = Columns(i).ColumnWidth * 1.05
    Next
    Application.ScreenUpdating = True
End Sub

Sub çsïù()
    Application.ScreenUpdating = False
    For i = 2 To 97
        Rows(i).RowHeight = Rows(i).RowHeight * 1.05
    Next
    Application.ScreenUpdating = True
End Sub
