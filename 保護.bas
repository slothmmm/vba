Sub 保護()
    ActiveSheet.Protect Contents:=True, DrawingObjects:=False, UserInterfaceOnly:=True, AllowFormattingCells:=True
End Sub


Sub 保護_複数()
    first_sheet = ActiveSheet.Name
    For ws_num = 1 To Worksheets.Count
        Worksheets(ws_num).Protect Contents:=True, DrawingObjects:=False, UserInterfaceOnly:=True, AllowFormattingCells:=True
    Next ws_num
    Worksheets(first_sheet).Activate
End Sub