Attribute VB_Name = "•ÛŒì"
Sub •ÛŒì()
    ActiveSheet.Protect AllowFiltering:=True
End Sub

Sub •ÛŒì‰ğœ()
    ActiveSheet.Unprotect
End Sub

Sub •¡”•ÛŒì()
    Worksheets("İ’è").Protect AllowFiltering:=True
    Worksheets("ˆóü‘¼").Protect AllowFiltering:=True
    Worksheets("Ü–¡ŠúŒÀ").Protect AllowFiltering:=True
    Worksheets("ˆóüCN").Protect AllowFiltering:=True
    Worksheets("Œ`¬1").Protect AllowFiltering:=True
    Worksheets("Œ`¬2").Protect AllowFiltering:=True
End Sub

Sub ‘S•ÛŒì‰ğœ()
    Dim sh As Object
    On Error Resume Next
    For Each sh In Sheets
    sh.Unprotect
    Next sh
End Sub
