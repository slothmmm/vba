Sub 受注更新()
    Dim rc As VbMsgBoxResult
    rc = MsgBox("BLACKSHIPからの混載ラベルcsvの読み込みを開始します。" & vbCrLf & "BLACKSHIPからラベルcsvをダウンロードしましたか？", vbYesNo + vbQuestion)
    If rc = vbNo Then
        MsgBox "読み込みを中止します", vbCritical
        End
    End If
    
    Call csv読み込みmain
    Call 形成main
    Call 受注データ反映main
End Sub

Sub 受注データ反映main()

    p_shipdate = Worksheets("ピッキング表").Range("D6").Value
    k_shipdate = Worksheets("形成").Range("H2").Value
    k_shipdate = k_shipdate - 1
    
    If p_shipdate <> k_shipdate Then
        MsgBox ("ピッキング表の出荷日と、形成シートの出荷日(納品日-1)が一致しません。")
        End
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual     '手動計算

    Worksheets("ピッキング表").Activate
    pData = Worksheets("ピッキング表").Range(Cells(3, 6), Cells(38, 12))   'F3からL38

    'クリア
    For r = 1 To 7  'F3 から L3　　列
        If Not IsEmpty(pData(1, r)) Then
            For s = 6 To 36 'F8からF38   行
                pData(s, r) = Empty
            Next s
        End If
    Next r
    
    Worksheets("形成").Activate
    B_LastRow = Cells(Rows.Count, 2).End(xlUp).Row  ''B列の最終行取得
    
    Dim kData As Variant
    
    kData = Worksheets("形成").Range(Cells(1, 1), Cells(B_LastRow, 8))
    
    Worksheets("ピッキング表").Activate
    '数量反映
    For p = 1 To UBound(kData)
        For r = 1 To 7  'F3 から L3　　列
            If pData(1, r) = kData(p, 4) Then
                For s = 6 To 36 'F8からF38   行
                    If Not IsEmpty(pData(s, 7)) Then
                        If pData(s, 7) <> "" Then
                            If str(pData(s, 7)) = str(kData(p, 2)) Then
                                pData(s, r) = kData(p, 6)
                            End If
                        End if
                    End If
                Next s
            End If
        Next r
    Next p
    
    '******************ピッキング表へデータ反映
    Worksheets("ピッキング表").Activate
    Dim paste_data As Variant
    
    '******************中之島
    ReDim paste_data(30, 0)
    For i = 0 To UBound(paste_data)
        paste_data(i, 0) = pData(i + 6, 1)
    Next i
    
    Worksheets("ピッキング表").Range(Cells(8, 6), Cells(38, 6)) = paste_data
    
    '******************上越
    ReDim paste_data(30, 0)
    For i = 0 To UBound(paste_data)
        paste_data(i, 0) = pData(i + 6, 3)
    Next i
    
    '貼り付け
    Worksheets("ピッキング表").Unprotect
    Worksheets("ピッキング表").Range(Cells(8, 8), Cells(38, 8)) = paste_data
    Worksheets("ピッキング表").Protect
    
    Application.ScreenUpdating = False
    Application.Calculation = xlAutomatic
End Sub


