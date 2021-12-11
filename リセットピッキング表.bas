sub ƒŠƒZƒbƒg()
    ActiveSheet.Unprotect
    Application.ScreenUpdating = False

    Range("D1").Formula = Date + 1
    Worksheets("csv").Cells.Clear
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    Application.ScreenUpdating = True
End Sub