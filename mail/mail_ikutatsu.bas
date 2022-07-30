Sub ラベル印刷と計算()
    Dim rc As VbMsgBoxResult
    rc = MsgBox("①データ更新後、②印刷を行います。よろしいですか？", vbYesNo + vbQuestion)
    If rc = vbNo Then
        MsgBox "印刷を中止します", vbCritical
        Exit Sub
    End If

    Call 形成用
    Call ラベル印刷

End Sub

Sub ラベル印刷()

    Dim rc As VbMsgBoxResult
    rc = MsgBox("データ更新を行わず、印刷を行います。IP67で印刷設定していますか？", vbYesNo + vbQuestion)
    If rc = vbNo Then
        MsgBox "印刷を中止します", vbCritical
        Exit Sub
    End If
    
    this_sheet_name = ActiveSheet.Name

    'プリンター選択確認
    Dim myPrinter As String
    myPrinter = Application.ActivePrinter

    If myPrinter Like "*67*" Then
        Application.ActivePrinter = myPrinter
    Else
        MsgBox myPrinter & "が選択されています。" & vbCrLf & "プリンターの設定をIP190へ変更して下さい。"
        Exit Sub
    End If

    Worksheets("センターマスター").Activate
    E_LastRow = Cells(Rows.Count, 5).End(xlUp).Row  'E列の最終行取得
    centerList = Worksheets("センターマスター").Range(Cells(4, 5), Cells(E_LastRow, 5))   'シートデータ取得

    Dim ListArray(1) As String
    ListArray(0) = "START"
    ListArray(1) = ""

    Worksheets(this_sheet_name).Activate
    Call フィルタークリア
    
    For cen = 1 To UBound(centerList)
        Worksheets(this_sheet_name).Range("C7") = cen
        
        dup = 重複なしリスト(cen)
        
        Worksheets(this_sheet_name).Activate
        
        If cen = 1 Then
            MsgBox "※沼津の印刷を開始します。"
        End If
        
        If cen = 2 Then
            MsgBox "※森の里の印刷を開始します。"
        End If
        
        For i = 0 To UBound(dup)
            
            '１番目の商品のみ「センター区切り」を印刷する
            If i = 0 Then
                'ListArray(0) = "START"
                 ListArray(0) = "これはフィルターにかからない文字列"    'STARTいらなそうなので
            Else
                ListArray(0) = "これはフィルターにかからない文字列"
            End If
        
            If dup(i) <> "" Then
                ListArray(1) = dup(i)
            Else
                 GoTo Continue ' Continue: の行へ処理を飛ばす
            End If
        
            With ActiveSheet
                .Range("B10").Select
                'If .FilterMode Then .ShowAllData
                .Range("B10").AutoFilter Field:=3, Criteria1:=ListArray, Operator:=xlFilterValues
                'Application.CalculateFull
                Application.Calculate   '再計算
            End With
            ActiveSheet.PrintOut
Continue:                 ' GoTo Continue の後はここから処理が行われる
        Next i
    Next cen
    
    Call フィルタークリア
    MsgBox "ラベル印刷が完了しました"
End Sub

Function 重複なしリスト(centerNo As Variant) As Variant
    
    Worksheets("形成").Activate
    
    Dim 辞書 As Object
    Set 辞書 = CreateObject("Scripting.Dictionary")
    A_LastRow = Cells(Rows.Count, 1).End(xlUp).Row  'A列の最終行取得
    
    dataList = Worksheets("形成").Range(Cells(1, 1), Cells(A_LastRow, 12))
    
    For i = 1 To A_LastRow
        If IsNumeric(dataList(i, 11)) Then
            If centerNo = Int(Left(dataList(i, 11), Len(dataList(i, 11)) - 4)) Then
                '辞書に登録されていない時は
                 If Not 辞書.Exists(dataList(i, 3)) Then
                     '辞書に登録する。
                     辞書(dataList(i, 3)) = Empty
                 End If
            End If
        End If
    Next i
    
    重複なしリスト = 辞書.keys
    
End Function

