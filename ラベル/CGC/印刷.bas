Sub ラベル印刷と計算()
    Dim rc As VbMsgBoxResult
    rc = MsgBox("データ更新後、印刷を行います。IP190で印刷設定していますか？", vbYesNo + vbQuestion)
    If rc = vbNo Then
        MsgBox "印刷を中止します", vbCritical
        Exit Sub
    End If
    
    Call 形成用
    Call ラベル印刷

End Sub

Sub ラベル印刷()
    Dim rc As VbMsgBoxResult
    rc = MsgBox("データ更新を行わず、印刷を行います。IP190で印刷設定していますか？", vbYesNo + vbQuestion)
    If rc = vbNo Then
        MsgBox "印刷を中止します", vbCritical
        Exit Sub
    End If

    'プリンター選択確認
    Dim myPrinter As String
    myPrinter = Application.ActivePrinter

    If myPrinter Like "*190*" Then
    Application.ActivePrinter = myPrinter
    Else
        MsgBox myPrinter & "が選択されています。" & vbCrLf & "プリンターの設定をIP190へ変更して下さい。"
        Exit Sub
    End If

    Worksheets("センターマスター").Activate
    K_LastRow = Cells(Rows.Count, 11).End(xlUp).Row  'K列の最終行取得
    centerList = Worksheets("センターマスター").Range(Cells(4, 11), Cells(K_LastRow, 11))   'シートデータ取得

    Dim ListArray(2) As String
    ListArray(0) = "START"
    ListArray(1) = ""
    ListArray(2) = "END"
        'ラベル
    Worksheets("ラベル").Activate
    
    For i = 1 To UBound(centerList)
        
        If centerList(i, 1) <> "" Then
            ListArray(1) = centerList(i, 1)
        Else
             GoTo Continue ' Continue: の行へ処理を飛ばす
        End If
    
        With ActiveSheet
            .Range("A1").Select
            If .FilterMode Then .ShowAllData
            .Range("A1").AutoFilter Field:=2, Criteria1:=ListArray, Operator:=xlFilterValues
        End With
        ActiveSheet.PrintOut
Continue:             ' GoTo Continue の後はここから処理が行われる
    Next i
    
    'ActiveSheet.PrintOut
End Sub
