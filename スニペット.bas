'エクセルopen時に計算自動へ設定(ThisWorkbookへ貼り付け)
Private Sub Workbook_Open()
    Application.Calculation = xlAutomatic
End Sub


'部分一致検索
Sub test()
    Dim a As Variant
    a = "abcdefg"
    If a Like "**" Then
        Debug.Print "あ"
    End If
End Sub