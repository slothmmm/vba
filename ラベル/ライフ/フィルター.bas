
Sub フィルタークリア()
    With ActiveSheet
        .Range("B10").Select
        If .FilterMode Then .ShowAllData
    End With
End Sub


