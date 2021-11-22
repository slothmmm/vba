Attribute VB_Name = "賞味期限並び替え"
'''''''''''''''''''''''''''''''        「印刷他」関連         ''''''''''''''''''
Sub 賞味期限並び替え()
    Call 保護.全保護解除

    Worksheets("印刷他").Activate
    
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
    End With
    
    'ソート
    Dim y As Long, i As Long, j As Long, v As Long, swap As Variant, B As Variant, C As Variant
    B = Range(Cells(11, 13), Cells(9010, 18))

    For y = 1 To 8999 ''ソート開始
        For i = 1 To 6
            For j = 1 To 6
                If B(y, i) < B(y, j) Then
                    swap = B(y, i)
                    B(y, i) = B(y, j)
                    B(y, j) = swap
                End If
            Next j
        Next i
    Next y

    For y = 1 To 8999 ''左詰め
        For v = 1 To 5
         For i = 1 To 5
             If B(y, i) = "" Then
                 For j = i To 5
                     B(y, j) = B(y, j + 1)
                     B(y, j + 1) = ""
                 Next j
             End If
        Next i
       Next v
   Next y

    Range(Cells(11, 13), Cells(9010, 18)) = B
    
    'Call 更新日.更新日_判定
    Call 保護.複数保護
    MsgBox "賞味期限並び替え完了"
End Sub


