Sub 保護()
    ActiveSheet.Protect
End Sub

Sub 保護解除()
    ActiveSheet.UnProtect
End Sub

Sub 複数保護()
  '  Worksheets("出荷数アイテム").Protect
 '   Worksheets("出荷数キット").Protect
    Worksheets("受注入力").Protect
    'Worksheets("手入力").Protect
End Sub

Sub 全保護解除()
    'ActiveSheet.UnProtect Password:="tetsuya0001"
    Dim sh As Object
    On Error Resume Next
    For Each sh In Sheets
    'sh.UnProtect Password:="tetsuya0001"
    sh.UnProtect
    Next sh
End Sub
