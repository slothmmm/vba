Attribute VB_Name = "íœ"

''''''''''''''''''''''''''''''''               íœŠÖ˜A         ''''''''''''''''''''''''''''''''''''''
Sub íœ_dŠ|‚è()
     Worksheets("dŠ|").Activate
     Worksheets("dŠ|").Range(Cells(11, 1), Cells(9010, 34)).Clear
End Sub

Sub íœ_“ü”()
     Worksheets("“ü”").Activate
     Worksheets("“ü”").Range(Cells(11, 1), Cells(9010, 34)).Clear
End Sub

Sub íœ_İŒÉ”()
     Worksheets("İŒÉ”").Activate
     Worksheets("İŒÉ”").Range(Cells(11, 1), Cells(9010, 47)).Clear
End Sub

Sub íœ_CN”»’è()
     Worksheets("CN”»’è").Activate
     Worksheets("CN”»’è").Range(Cells(11, 15), Cells(9010, 47)).Clear
     Worksheets("CN”»’è").Range(Cells(11, 2), Cells(9010, 2)).Clear
                   'Œ´—¿CD1001-9999
              Dim i As Long, B As Variant
              ReDim B(9010, 0)
              For i = 0 To 8998
                B(i, 0) = i + 1001
              Next i
              Range("M11:M9999") = B
End Sub

Sub íœ_IY”»’è()
     Worksheets("IY”»’è").Activate
     Worksheets("IY”»’è").Range(Cells(11, 3), Cells(9010, 34)).Clear
              'Œ´—¿CD1001-9999
              Dim i As Long, B As Variant
              ReDim B(9010, 0)
              For i = 0 To 8998
                B(i, 0) = i + 1001
              Next i
              Range("A11:A9999") = B
End Sub

Sub íœ_ALL()
    Call íœ_dŠ|‚è
    Call íœ_“ü”
    Call íœ_İŒÉ”
    Call íœ_CN”»’è
    Call íœ_IY”»’è
End Sub
