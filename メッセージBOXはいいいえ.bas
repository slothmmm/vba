    Dim rc As VbMsgBoxResult
    rc = MsgBox("リセットしますか？" & vbCrLf & vbCrLf & "�@出荷日を翌日へ" & vbCrLf & "�Acsvシートのリセット", vbYesNo + vbQuestion)
    If rc = vbNo Then
        MsgBox "処理を中止します", vbCritical
        Exit Sub
    End If
