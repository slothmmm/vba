Attribute VB_Name = "í"

''''''''''''''''''''''''''''''''               íÖA         ''''''''''''''''''''''''''''''''''''''
Sub í_d|è()
     Worksheets("d|").Activate
     Worksheets("d|").Range(Cells(11, 1), Cells(9010, 34)).Clear
End Sub

Sub í_ü()
     Worksheets("ü").Activate
     Worksheets("ü").Range(Cells(11, 1), Cells(9010, 34)).Clear
End Sub

Sub í_ÝÉ()
     Worksheets("ÝÉ").Activate
     Worksheets("ÝÉ").Range(Cells(11, 1), Cells(9010, 47)).Clear
End Sub

Sub í_CN»è()
     Worksheets("CN»è").Activate
     Worksheets("CN»è").Range(Cells(11, 15), Cells(9010, 47)).Clear
     Worksheets("CN»è").Range(Cells(11, 2), Cells(9010, 2)).Clear
                   '´¿CD1001-9999
              Dim i As Long, B As Variant
              ReDim B(9010, 0)
              For i = 0 To 8998
                B(i, 0) = i + 1001
              Next i
              Range("M11:M9999") = B
End Sub

Sub í_IY»è()
     Worksheets("IY»è").Activate
     Worksheets("IY»è").Range(Cells(11, 3), Cells(9010, 34)).Clear
              '´¿CD1001-9999
              Dim i As Long, B As Variant
              ReDim B(9010, 0)
              For i = 0 To 8998
                B(i, 0) = i + 1001
              Next i
              Range("A11:A9999") = B
End Sub

Sub í_ALL()
    Call í_d|è
    Call í_ü
    Call í_ÝÉ
    Call í_CN»è
    Call í_IY»è
End Sub
