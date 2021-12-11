Attribute VB_Name = "在庫表作成"
Sub 翌月在庫表作成()
    Call 翌月エクセル作成
    Call 翌月置換.置換main
    Call ファイル保存(dateList, fileList)
    Call 保護.保護ロック全シート循環
    Call 画面とアラート表示
    Call 終了
End Sub

Sub 翌月エクセル作成()
'セッティング 」」」」」
    Set dateList = 在庫表作成_置換.日付クラス()
    Set fileList = 在庫表作成_置換.ファイルクラス()
    Call 作成確認
    Call 前処理(dateList, fileList)
    Call ファイル存在確認(dateList, fileList)
    Call ファイル作成(dateList, fileList)
    Call 合計金額シートB2翌月へ変更(dateList, fileList)
'    Call 在庫表作成_置換.main
'    Call ファイル保存(dateList, fileList)
'    Call 保護.保護ロック全シート循環
'    Call 画面とアラート表示
'    Call 終了
End Sub



Sub 作成確認()
    boool = MsgBox("翌月在庫表の作成　を行いますが、よろしいですか？", vbYesNo + vbQuestion)
    If boool = vbYes Then
        Exit Sub
    Else
        End
    End If
End Sub

Sub 前処理(dateList As Variant, fileList As Variant)
    Call 画面とアラート非表示
    Call データチェック(dateList, fileList)
End Sub

Sub 画面とアラート非表示()
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
End Sub

Sub 画面とアラート表示()
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

Sub データチェック(dateList As Variant, fileList As Variant)
   Debug.Print "「dateList.date_now」 " & dateList.date_now
   Debug.Print "「dateList.date_next」 " & dateList.date_next
   Debug.Print "「dateList.date_last」 " & dateList.date_last
   Debug.Print "「dateList.H_AL_nowMonth」 " & dateList.H_AL_nowMonth
   Debug.Print "「dateList.H_AL_nextMonth」 " & dateList.H_AL_nextMonth
   Debug.Print "「dateList.G_now_date」 " & dateList.G_now_date
   Debug.Print "「dateList.G_last_date」 " & dateList.G_last_date
   Debug.Print "「dateList.G_next_date」 " & dateList.G_next_date
   Debug.Print "「dateList.kowake_now_date」 " & dateList.kowake_now_date
   Debug.Print "「dateList.kowake_last_date」 " & dateList.kowake_last_date
   Debug.Print "「dateList.kowake_next_date」 " & dateList.kowake_next_date

   Debug.Print "「fileList.this_filename」 " & fileList.this_filename
   Debug.Print "「fileList.mypath」 " & fileList.mypath
   Debug.Print "「fileList.mybook」 " & fileList.mybook
   Debug.Print "「fileList.mybook_month」 " & fileList.mybook_month
   Debug.Print "「fileList.mmfn」 " & fileList.mmfn
   Debug.Print "「fileList.fn」 " & fileList.fn
End Sub

Sub 合計金額シートB2翌月へ変更(dateList As Variant, fileList As Variant)
    ActiveWorkbook.Sheets("合計金額").Activate
    Range("B2").Formula = dateList.date_next '合計金額シートB2翌月へ変更
End Sub


Sub ファイル存在確認(dateList As Variant, fileList As Variant)
   If Dir(fileList.next_mybook) = "" Then
        'MsgBox "翌月在庫表を作成します"
    Else
        MsgBox "翌月分の在庫表は" & vbNewLine & _
               "既に存在しています。" & vbNewLine & _
               "  " & vbNewLine & _
               "処理を中止します。"
        End
    End If
End Sub

Sub ファイル作成(dateList As Variant, fileList As Variant)

    Dim mybk As Workbook
    Set mybk = ThisWorkbook
    mybk.SaveAs (fileList.next_mybook)

    MsgBox fileList.next_mybook & "作成しました"
End Sub

Sub ファイル保存(dateList As Variant, fileList As Variant)
        ActiveWorkbook.Save
End Sub


Sub 終了()
    MsgBox "終了しました。"
End Sub
