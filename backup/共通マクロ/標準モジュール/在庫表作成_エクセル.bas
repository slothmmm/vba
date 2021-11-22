Attribute VB_Name = "在庫表作成_エクセル"
Sub 作成main()
'セッティング 」」」」」
    Set dateList = 在庫表作成_置換.日付クラス()
    Set fileList = 在庫表作成_置換.ファイルクラス()

    Call 画面とアラート非表示
    Call 在庫表作成_置換.データチェック(dateList, fileList)
    Call ファイル存在確認(dateList, fileList)
    Call ファイル作成(dateList, fileList)
    Call 在庫表作成_置換.main
    Call ファイル保存(dateList, fileList)
    Call 画面とアラート表示
    Call 終了
End Sub

Sub ファイル存在確認(dateList As Variant, fileList As Variant)
   If Dir(fileList.next_mybook) = "" Then
        MsgBox "翌月在庫表を作成します"
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
'    boool = MsgBox("ファイルを保存しますか？", vbYesNo + vbQuestion)
'    If boool = vbYes Then
'        ActiveWorkbook.Save
'    Else
'        End
'    End If
        ActiveWorkbook.Save
End Sub

Sub 画面とアラート非表示()
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
End Sub

Sub 画面とアラート表示()
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

Sub 終了()
    MsgBox "終了しました。"
End Sub
