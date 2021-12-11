
Sub 保護()
    ActiveSheet.Protect Contents:=True, DrawingObjects:=False, UserInterfaceOnly:=True, AllowFormattingCells:=True
End Sub

Sub 保護解除()
    ActiveSheet.Unprotect
End Sub

Sub 保護_複数()
    For ws_num = 1 To ThisWorkbook.Worksheets.Count
        Worksheets(ws_num).Protect Contents:=True, DrawingObjects:=False, UserInterfaceOnly:=True, AllowFormattingCells:=True
    Next ws_num
End Sub

Sub 保護_全解除()
    For ws_num = 1 To ThisWorkbook.Worksheets.Count
        Worksheets(ws_num).Unprotect
    Next ws_num
End Sub

Sub シート複数選択解除()
    ActiveWindow.SelectedSheets(1).Select
End Sub
