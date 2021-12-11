Attribute VB_Name = "更新日"
Sub 賞味期限_COPY_OPEN()
    Call 保護.全保護解除

    Worksheets("印刷他").Activate
    
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
    End With
    
    Dim y As Long, i As Long, j As Long, v As Long, syomi As Variant
    syomi = Range(Cells(11, 13), Cells(9010, 18))           ''賞味期限更新前
    Range(Cells(11, 46), Cells(9010, 51)) = syomi           ''ペースト
End Sub

Sub 更新日_判定()
    Call 保護.全保護解除

    Worksheets("印刷他").Activate
    
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
    End With

    Dim koushin As Variant, syo_mi_mae As Variant, syo_mi_ato As Variant
    Dim i As Long, s As Long, d As Date
    d = Date
    syo_mi_mae = Range(Cells(11, 46), Cells(9010, 51))     ''賞味期限変更前
    syo_mi_ato = Range(Cells(11, 13), Cells(9010, 18))     ''賞味期限変更後
    koushin = Range(Cells(11, 12), Cells(9010, 12))        ''更新日
    
    For i = 1 To 8999                                      ''比較判定
        For s = 1 To 6
            If syo_mi_mae(i, s) <> syo_mi_ato(i, s) Then
                koushin(i, 1) = d                          ''更新日へ現在の日付
            End If
        Next s
    Next i
    
    Range(Cells(11, 12), Cells(9010, 12)) = koushin       ''更新日ペースト
    
End Sub

Sub 更新日_コープ()
    Call 保護.全保護解除
    
    Worksheets("印刷他").Activate
    
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
    End With
    
    Dim cn_bool As Variant, koushin As Variant
    Dim i As Long, s As Long, d As Date
    d = Date
    cn_bool = Range(Cells(11, 27), Cells(9010, 27))     ''賞味期限変更前
    koushin = Range(Cells(11, 12), Cells(9010, 12))        ''更新日
    
    For i = 1 To 8999                                      ''比較判定
            If cn_bool(i, 1) = 1 Or cn_bool(i, 1) = 2 Then
                koushin(i, 1) = d                          ''更新日へ現在の日付
            End If
    Next i
    Range(Cells(11, 12), Cells(9010, 12)) = koushin       ''更新日ペースト
    MsgBox "更新日の更新完了 コープ"
End Sub
Sub 更新日_IY()
    Call 保護.全保護解除
    
    Worksheets("印刷他").Activate
    
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
    End With
    
    Dim cn_bool As Variant, koushin As Variant
    Dim i As Long, s As Long, d As Date
    d = Date
    cn_bool = Range(Cells(11, 31), Cells(9010, 31))     ''賞味期限変更前
    koushin = Range(Cells(11, 12), Cells(9010, 12))        ''更新日
    
    For i = 1 To 8999                                      ''比較判定
            If cn_bool(i, 1) = 1 Or cn_bool(i, 1) = 2 Then
                koushin(i, 1) = d                          ''更新日へ現在の日付
            End If
    Next i
    Range(Cells(11, 12), Cells(9010, 12)) = koushin       ''更新日ペースト
    MsgBox "更新日の更新完了 IY"
End Sub

Sub 更新日_副原材料()
    Call 保護.全保護解除
    
    Worksheets("印刷他").Activate
    
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
    End With
    
    Dim cn_bool As Variant, koushin As Variant
    Dim i As Long, s As Long, d As Date
    d = Date
    cn_bool = Range(Cells(11, 26), Cells(9010, 26))     ''賞味期限変更前
    koushin = Range(Cells(11, 12), Cells(9010, 12))        ''更新日
    
    For i = 1 To 8999                                      ''比較判定
            If cn_bool(i, 1) = 1 Or cn_bool(i, 1) = 2 Then
                koushin(i, 1) = d                          ''更新日へ現在の日付
            End If
    Next i
    Range(Cells(11, 12), Cells(9010, 12)) = koushin       ''更新日ペースト
    MsgBox "更新日の更新完了 副原材料"
End Sub
Sub 更新日_諸口_土()
    Call 保護.全保護解除
    
    Worksheets("印刷他").Activate
    
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
    End With
    
    Dim cn_bool As Variant, koushin As Variant
    Dim i As Long, s As Long, d As Date
    d = Date
    cn_bool = Range(Cells(11, 45), Cells(9010, 45))     ''賞味期限変更前
    koushin = Range(Cells(11, 12), Cells(9010, 12))        ''更新日
    
    For i = 1 To 8999                                      ''比較判定
            If cn_bool(i, 1) = 1 Or cn_bool(i, 1) = 2 Then
                koushin(i, 1) = d                          ''更新日へ現在の日付
            End If
    Next i
    Range(Cells(11, 12), Cells(9010, 12)) = koushin       ''更新日ペースト
    MsgBox "更新日の更新完了 諸口_土"
End Sub
Sub 更新日_諸口_金()
    Call 保護.全保護解除
    
    Worksheets("印刷他").Activate
    
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
    End With
    
    Dim cn_bool As Variant, koushin As Variant
    Dim i As Long, s As Long, d As Date
    d = Date
    cn_bool = Range(Cells(11, 44), Cells(9010, 44))     ''賞味期限変更前
    koushin = Range(Cells(11, 12), Cells(9010, 12))        ''更新日
    
    For i = 1 To 8999                                      ''比較判定
            If cn_bool(i, 1) = 1 Or cn_bool(i, 1) = 2 Then
                koushin(i, 1) = d                          ''更新日へ現在の日付
            End If
    Next i
    Range(Cells(11, 12), Cells(9010, 12)) = koushin       ''更新日ペースト
    MsgBox "更新日の更新完了 諸口_金"
End Sub

Sub 更新日_諸口_木()
    Call 保護.全保護解除
    
    Worksheets("印刷他").Activate
    
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
    End With
    
    Dim cn_bool As Variant, koushin As Variant
    Dim i As Long, s As Long, d As Date
    d = Date
    cn_bool = Range(Cells(11, 43), Cells(9010, 43))     ''賞味期限変更前
    koushin = Range(Cells(11, 12), Cells(9010, 12))        ''更新日
    
    For i = 1 To 8999                                      ''比較判定
            If cn_bool(i, 1) = 1 Or cn_bool(i, 1) = 2 Then
                koushin(i, 1) = d                          ''更新日へ現在の日付
            End If
    Next i
    Range(Cells(11, 12), Cells(9010, 12)) = koushin       ''更新日ペースト
    MsgBox "更新日の更新完了 諸口_木"
End Sub

Sub 更新日_諸口_水()
    Call 保護.全保護解除
    
    Worksheets("印刷他").Activate
    
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
    End With
    
    Dim cn_bool As Variant, koushin As Variant
    Dim i As Long, s As Long, d As Date
    d = Date
    cn_bool = Range(Cells(11, 42), Cells(9010, 42))     ''賞味期限変更前
    koushin = Range(Cells(11, 12), Cells(9010, 12))        ''更新日
    
    For i = 1 To 8999                                      ''比較判定
            If cn_bool(i, 1) = 1 Or cn_bool(i, 1) = 2 Then
                koushin(i, 1) = d                          ''更新日へ現在の日付
            End If
    Next i
    Range(Cells(11, 12), Cells(9010, 12)) = koushin       ''更新日ペースト
    MsgBox "更新日の更新完了 諸口_水"
End Sub

Sub 更新日_諸口_火()
    Call 保護.全保護解除
    
    Worksheets("印刷他").Activate
    
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
    End With
    
    Dim cn_bool As Variant, koushin As Variant
    Dim i As Long, s As Long, d As Date
    d = Date
    cn_bool = Range(Cells(11, 41), Cells(9010, 41))     ''賞味期限変更前
    koushin = Range(Cells(11, 12), Cells(9010, 12))        ''更新日
    
    For i = 1 To 8999                                      ''比較判定
            If cn_bool(i, 1) = 1 Or cn_bool(i, 1) = 2 Then
                koushin(i, 1) = d                          ''更新日へ現在の日付
            End If
    Next i
    Range(Cells(11, 12), Cells(9010, 12)) = koushin       ''更新日ペースト
    MsgBox "更新日の更新完了 諸口_火"
End Sub

Sub 更新日_諸口_月()
    Call 保護.全保護解除
    
    Worksheets("印刷他").Activate
    
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
    End With
    
    Dim cn_bool As Variant, koushin As Variant
    Dim i As Long, s As Long, d As Date
    d = Date
    cn_bool = Range(Cells(11, 40), Cells(9010, 40))     ''賞味期限変更前
    koushin = Range(Cells(11, 12), Cells(9010, 12))        ''更新日
    
    For i = 1 To 8999                                      ''比較判定
            If cn_bool(i, 1) = 1 Or cn_bool(i, 1) = 2 Then
                koushin(i, 1) = d                          ''更新日へ現在の日付
            End If
    Next i
    Range(Cells(11, 12), Cells(9010, 12)) = koushin       ''更新日ペースト
    MsgBox "更新日の更新完了 諸口_月"
End Sub

Sub 更新日_諸口_日()
    Call 保護.全保護解除
    
    Worksheets("印刷他").Activate
    
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
    End With
    
    Dim cn_bool As Variant, koushin As Variant
    Dim i As Long, s As Long, d As Date
    d = Date
    cn_bool = Range(Cells(11, 39), Cells(9010, 39))     ''賞味期限変更前
    koushin = Range(Cells(11, 12), Cells(9010, 12))        ''更新日
    
    For i = 1 To 8999                                      ''比較判定
            If cn_bool(i, 1) = 1 Or cn_bool(i, 1) = 2 Then
                koushin(i, 1) = d                          ''更新日へ現在の日付
            End If
    Next i
    Range(Cells(11, 12), Cells(9010, 12)) = koushin       ''更新日ペースト
    MsgBox "更新日の更新完了 諸口_日"
End Sub

