'帳票印刷処理2
    Sheets("チェックシート").Select
    ActiveSheet.AutoFilterMode = False
    Range("A1").AutoFilter Field:=1, Criteria1:="<>"
    ActiveWindow.SelectedSheets.PrintOut Copies:=p1, Collate:=True, _
        IgnorePrintAreas:=False
    Range("A1").AutoFilter Field:=1