Sub フィルタークリア()
    With ActiveSheet
        .Range("A1").Select
        If .FilterMode Then .ShowAllData
    End With
End Sub
