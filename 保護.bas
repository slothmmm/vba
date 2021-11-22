Sub •ÛŒì()
    ActiveSheet.Protect Contents:=True, DrawingObjects:=False, UserInterfaceOnly:=True, AllowFormattingCells:=True
End Sub

Sub •ÛŒì‰ğœ()
    ActiveSheet.Unprotect
End Sub

Sub •ÛŒì_•¡”()
    For ws_num = 1 To Worksheets.Count
        Worksheets(ws_num).Protect Contents:=True, DrawingObjects:=False, UserInterfaceOnly:=True, AllowFormattingCells:=True
    Next ws_num
    
End Sub

Sub •ÛŒì_‘S‰ğœ()
    Dim sh As Object
    On Error Resume Next
    For Each sh In Sheets
    sh.Unprotect
    Next sh
    
End Sub