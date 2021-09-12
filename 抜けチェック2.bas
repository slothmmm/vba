Sub Error_check()
    '保存したときのシート
    first_name = ActiveSheet.Name
    
    ActiveWorkbook.Sheets("合計金額").Activate
 '最終行取得
        Dim LastRow As Long
        LastRow = Cells(Rows.Count, 9).End(xlUp).Row
        Debug.Print LastRow
        '数式を配列へ格納     F列からAJ列
        search_cell = Range(Cells(1, 9), Cells(LastRow, 9))

        '置換
        Dim nuke As Integer
        nuke = 0
        For s = 1 To LastRow
               
                    If search_cell(s, 1) = "エラー" Then
                          nuke = 1
                    End If
              
        Next s

        'アクティブシートをもとに戻す
        ActiveWorkbook.Sheets(first_name).Activate

        If nuke = 1 Then
            MsgBox "数式が抜けています。確認してください。"
        End If
End Sub
