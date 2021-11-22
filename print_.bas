'印刷処理2
    Sheets("作業順番表").Select
    ActiveSheet.Range("I6").AutoFilter Field:=9, Criteria1:="<>"
    ActiveWindow.SelectedSheets.PrintOut Copies:=p2, Collate:=True, _
        IgnorePrintAreas:=False
    ActiveSheet.Range("I6").AutoFilter Field:=9

'印刷処理3
    Sheets("チェックシート").Select
    ActiveSheet.Range("A1").AutoFilter Field:=1, Criteria1:="<>"
    ActiveWindow.SelectedSheets.PrintOut Copies:=p3, Collate:=True, _
        IgnorePrintAreas:=False
    ActiveSheet.Range("A1").AutoFilter Field:=1

'看板
    Sheets("看板").Select
    ActiveSheet.Range("A1").AutoFilter Field:=1, Criteria1:="<>"
    ActiveWindow.SelectedSheets.PrintOut Copies:=2, Collate:=True, _
        IgnorePrintAreas:=False
    ActiveSheet.Range("A1").AutoFilter Field:=1

'    2021.02.09
    Sheets("事務用シール").Select
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False