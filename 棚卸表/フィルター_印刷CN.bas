Attribute VB_Name = "tB^[_óüCN"
'''''''''''''''''''''''''''''''      tB^[ÖA                   ''''''''''''''''''''''''''''

Sub tB^[SNA_óüCNV[g()
    Call Ûì.SÛìð
    Worksheets("óüCN").Activate
    With ActiveSheet
        .Range("B4").Select
        If .FilterMode Then .ShowAllData
    End With
    Call Ûì.¡Ûì
    MsgBox "tB^[NA®¹(CN»è)"
End Sub

Sub tB^[CN_óüCNV[g()
    Call Ûì.SÛìð
    Worksheets("óüCN").Activate
    With ActiveSheet
        .Range("B4").Select
        If .FilterMode Then .ShowAllData
        .Range("B4").AutoFilter Field:=26, Criteria1:="<>"
    End With
    Call Ûì.¡Ûì
    MsgBox "tB^[®¹(CN»è)"
End Sub

