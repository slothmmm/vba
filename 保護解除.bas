Sub 保護解除()
    ActiveSheet.Unprotect
End Sub


Sub 保護_全解除()
    first_sheet = ActiveSheet.Name
    Dim sh As Object
    On Error Resume Next
    For Each sh In Sheets
    sh.Unprotect
    Next sh
    Worksheets(first_sheet).Activate
End Sub