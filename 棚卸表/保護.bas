Attribute VB_Name = "Ûì"
Sub Ûì()
    ActiveSheet.Protect AllowFiltering:=True
End Sub

Sub Ûìð()
    ActiveSheet.Unprotect
End Sub

Sub ¡Ûì()
    Worksheets("Ýè").Protect AllowFiltering:=True
    Worksheets("óü¼").Protect AllowFiltering:=True
    Worksheets("Ü¡úÀ").Protect AllowFiltering:=True
    Worksheets("óüCN").Protect AllowFiltering:=True
    Worksheets("`¬1").Protect AllowFiltering:=True
    Worksheets("`¬2").Protect AllowFiltering:=True
End Sub

Sub SÛìð()
    Dim sh As Object
    On Error Resume Next
    For Each sh In Sheets
    sh.Unprotect
    Next sh
End Sub
