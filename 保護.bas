Sub 保護()
    ActiveSheet.Protect Contents:=True, DrawingObjects:=False, UserInterfaceOnly:=True, AllowFormattingCells:=True
End Sub

Sub 保護解除()
    ActiveSheet.Unprotect
End Sub

Sub 保護_複数()
    For ws_num = 1 To Worksheets.Count
        Worksheets(ws_num).Protect Contents:=True, DrawingObjects:=False, UserInterfaceOnly:=True, AllowFormattingCells:=True
    Next ws_num
    
End Sub

Sub 保護_全解除()
    Dim sh As Object
    On Error Resume Next
    For Each sh In Sheets
    sh.Unprotect
    Next sh
    
End Sub