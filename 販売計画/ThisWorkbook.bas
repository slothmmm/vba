Private Sub Workbook_Open()
    Call 保護.シート複数選択解除
    Call 保護.保護_複数
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Call バックアップ_main
End Sub
