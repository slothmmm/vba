Sub フィルタークリア()
    ActiveSheet.Range("").Select
    If ActiveSheet.FilterMode Then ActiveSheet.ShowAllData
End Sub


Sub フィルター()
    ActiveSheet.Range("").Select
    If ActiveSheet.FilterMode Then ActiveSheet.ShowAllData
End Sub
