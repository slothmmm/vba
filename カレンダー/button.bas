
Sub 今週の計画読み込み()
    'アクティブ
    Worksheets("ピッキング表").Activate
    Worksheets("ピッキング表").Select
    
    Dim rc As VbMsgBoxResult
    rc = MsgBox("事前入力の準備をします。" & vbCrLf & "販売計画集計表で「今週」に設定した商品をC列に貼り付けますか？", vbYesNo + vbQuestion)
    If rc = vbNo Then
        MsgBox "事前入力の準備を中止します", vbCritical
        Exit Sub
    End If
    Worksheets("ピッキング表").Unprotect   '保護解除
    
    modosu_ship = Worksheets("ピッキング表").Range("D6")   '戻す用
    Application.ScreenUpdating = False                 '画面停止
    Application.Calculation = xlCalculationManual     '手動計算
    
    CalenderForm.Show   'カレンダー
    'Application.Calculate   '再計算
    kakunin_ship = Worksheets("ピッキング表").Range("D6")  '出荷日取得
    
    rc = MsgBox(Str(kakunin_ship) + "の商品一覧を読み込み、数量をクリアしますか？", vbYesNo + vbQuestion)
    If rc = vbNo Then
        Worksheets("ピッキング表").Range("D6") = modosu_ship    '戻す
    Application.Calculation = xlCalculationAutomatic    '自動計算
    Application.ScreenUpdating = True                 '画面
    Worksheets("ピッキング表").Protect   '保護
        
        MsgBox "事前入力の準備をを中止します", vbCritical
        Exit Sub
    End If
    
    'アクティブ
    Worksheets("計画").Activate
    Worksheets("計画").Select
    
    'AU列を取得して商品の数を取得
    au_column = Range(Cells(3, 47), Cells(100, 47))
    max_I = 1
    For i = 1 To UBound(au_column)
        'If Left(AU_column(i, 1), 1) <> "Ｉ" Then
        If IsError(au_column(i, 1)) Then
            max_I = i - 1
            Exit For
        End If
    Next i
    
    'アクティブ
    Worksheets("ピッキング表").Activate
    Worksheets("ピッキング表").Select
    
    
    '数量削除とC列削除
    Range("F8:I38,K8:L38").ClearContents
    Range(Cells(8, 3), Cells(38, 3)).ClearContents
    'C列貼付け
    Range(Cells(8, 3), Cells(8 + max_I - 1, 3)) = au_column
    
    
    Application.Calculation = xlCalculationAutomatic    '自動計算
    Application.ScreenUpdating = True                 '画面
    Worksheets("ピッキング表").Protect   '保護
    
    MsgBox ("事前入力の準備が完了しました。" & vbCrLf & "数量入力後、②事前入力の保存(出力)ボタンを押してください。")
End Sub

