Sub 印刷ラベル()
    'プリンター選択確認
    Dim myPrinter As String
    myPrinter = Application.ActivePrinter

    If myPrinter Like "*○○○○*" Then
        Application.ActivePrinter = myPrinter
    Else
        MsgBox myPrinter & "が選択されています。" & vbCrLf & "プリンターの設定を複合機5573へ変更して下さい。"
        End
    End If
    
    印刷
    Call フィルター
    ActiveSheet.PrintOut
End Sub


Sub フィルター()
    Application.Calculate
    ActiveSheet
        ActiveSheet.Range("A1").Select
        If ActiveSheet.FilterMode Then ActiveSheet.ShowAllData
        ActiveSheet.Range("A1").AutoFilter Field:=20, Criteria1:="印刷", Operator:=xlFilterValues
End Sub