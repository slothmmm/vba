Attribute VB_Name = "在庫表作成_置換"
Sub main()
    'セッティング 」」」」」
    Set dateList = 日付クラス()
    Set fileList = ファイルクラス()

    Call 実行前チェック(dateList, fileList)
    Call 全シート循環(dateList, fileList)
    Call 原料展開の置換(dateList, fileList) '重い　ポリ在庫表専用
    Call 合計金額シートB2翌月へ変更(dateList, fileList)
End Sub

Sub 実行前チェック(dateList As Variant, fileList As Variant)
    'ファイル名チェック
'    If fileList.bool_filename Then
'        Debug.Print "OK"
'    Else
'        Debug.Print fileList.bool_filename
'
'        MsgBox ("ファイル名orVBAのファイル名チェックがおかしいです。" & vbCrLf & fileList.this_filename & vbCrLf & fileList.checkFilename)
'        End
'    End If

    '合計金額シートB2セルチェック
    Dim boool As Long
    If fileList.mybook_month = (Month(dateList.date_now) + 1) Then
        boool = MsgBox("翌月在庫表作成のため「置換、削除」　を行いますが、よろしいですか？", vbYesNo + vbQuestion)
        If boool = vbYes Then
            Exit Sub
        Else
            End
        End If
    Else
        MsgBox ("�@ファイル名の月" & vbCrLf & "�A「合計金額」シート「B2」セル" & vbCrLf & "の関係がおかしいです。" & vbCrLf & vbCrLf & "正しい例" & vbCrLf & "�@在庫『包材』_2020.8.xlsm" & vbCrLf & "�A2020/7/1")
        End
    End If
End Sub

Sub 全シート循環(dateList As Variant, fileList As Variant)
    'H-AN列再表示 〜入荷数消去 」」」」」
    'シート循環　合計金額シートよりひだり
    For ws_num = 1 To (Sheets("合計金額").Index - 1)
        '隠れシートは飛ばす
        If Not Worksheets(ws_num).Visible Then
            GoTo CONTINUE
        End If

        '「棚卸表」「原料展開」シート飛ばす
        If Worksheets(ws_num).Name = "棚卸表" Or Worksheets(ws_num).Name = "原料展開" Then
            GoTo CONTINUE
        End If

        '表示シートアクティブ
        Worksheets(ws_num).Activate

        '実行
        Call H_AL列置換削除(dateList, fileList)
        Call G列置換(dateList.G_last_date, dateList.G_now_date)
        Call 入荷数の色(dateList, fileList)
        Call 預けの色(dateList, fileList)
        Call 戻しの色(dateList, fileList)
        
        If ActiveSheet.Name = "コープ_パック" Then
            Call サポートの色(dateList, fileList)
            Call ヨネヤマの色(dateList, fileList)
        End If
        
        If ActiveSheet.Name = "神田物産��" Then
            Call 服部コーヒーの色(dateList, fileList)
        End If
        
CONTINUE:
    Next ws_num
End Sub

Sub 原料展開の置換(dateList As Variant, fileList As Variant)
        If fileList.this_filename <> "在庫『ポリ』_" Then
            Debug.Print "原料展開の置換はやらない"
            Exit Sub
        End If
        Debug.Print "原料展開の置換開始"

        ActiveWorkbook.Sheets("原料展開").Activate
        '最終行取得
        Dim LastRow As Long
        LastRow = Cells(Rows.Count, 2).End(xlUp).Row
        Debug.Print LastRow
        '数式を配列へ格納     F列からAJ列
        tikan_cell = Range(Cells(2, 6), Cells(LastRow, 36)).Formula

        '置換
        For s = 1 To (LastRow - 1)
                For n = 1 To 31
                    tikan_cell(s, n) = Replace(tikan_cell(s, n), dateList.H_AL_nowMonth, dateList.H_AL_nextMonth)
                Next n
        Next s

        '配列をシートへ ペースト
        Range(Cells(2, 6), Cells(LastRow, 36)).Formula = tikan_cell
End Sub

Sub H_AL列置換削除(dateList As Variant, fileList As Variant)
        '最終行取得
        Dim LastRow1, LastRow2, LastRow3, LastRow4, LastRow5, LastRow As Long
        LastRow1 = Cells(Rows.Count, 5).End(xlUp).Row
        LastRow2 = Cells(Rows.Count, 6).End(xlUp).Row
        LastRow3 = Cells(Rows.Count, 7).End(xlUp).Row
        LastRow4 = Cells(Rows.Count, 8).End(xlUp).Row
        LastRow5 = Cells(Rows.Count, 38).End(xlUp).Row

        Dim arr As Variant
        arr = Array(LastRow1, LastRow2, LastRow3, LastRow4, LastRow5)
        LastRow = WorksheetFunction.Max(arr)

        '数式を配列へ格納     E列からAL列
        tikan_cell = Range(Cells(1, 5), Cells(LastRow, 38)).Formula

        '数式を配列へ格納     AO列からBS列
        If fileList.this_filename = "在庫『包材』_" Then
            tikan_222 = Range(Cells(1, 41), Cells(LastRow, 71)).Formula
        End If

        'E列を検索
        Debug.Print ActiveSheet.Name & " E列最終行は" & LastRow
        For s = 1 To LastRow

            '入荷数の削除
            If tikan_cell(s, 1) = "入荷数" Or tikan_cell(s, 1) = "合計入荷数" Or tikan_cell(s, 1) = "服部コーヒー" Then
                Debug.Print ActiveSheet.Name & s & "行 " & "入荷数の削除"
                For n = 4 To 34
                    tikan_cell(s, n) = ""
                Next n
            End If


            '「コープ_パック」シートの入荷数削除
            If ActiveSheet.Name = "コープ_パック" Then
                'サポートの削除
                If tikan_cell(s, 1) = "サポート" Then
                    Debug.Print "サポート入荷数"
                    For n = 4 To 34
                        tikan_cell(s, n) = ""
                    Next n
                End If

                'ヨネヤマの削除
                If tikan_cell(s, 1) = "ヨネヤマ" Then
                    Debug.Print "ヨネヤマ入荷数"
                    For n = 4 To 34
                        tikan_cell(s, n) = ""
                    Next n
                End If
            End If

            '調整の削除
            If tikan_cell(s, 1) = "調整" Then
                '数式が入っているかどうか
                If Mid(tikan_cell(s, 4), 1, 1) = "=" Then
                    Debug.Print ActiveSheet.Name & s & "行 " & "調整に数式あるので削除しない"
                Else
                    Debug.Print ActiveSheet.Name & s & "行 " & "調整削除"
                    For n = 4 To 34
                        tikan_cell(s, n) = ""
                    Next n
                End If
            End If

            '返品等の削除
            If tikan_cell(s, 1) = "返品等" Then
                Debug.Print ActiveSheet.Name & s & "行 " & "返品等"
                For n = 4 To 34
                    tikan_cell(s, n) = ""
                Next n
            End If


            '預けの削除
            If tikan_cell(s, 1) = "預け" Then
                Debug.Print ActiveSheet.Name & s & "行 " & "預け"
                For n = 4 To 34
                    tikan_cell(s, n) = ""
                Next n
            End If


            '戻しの削除
            If tikan_cell(s, 1) = "戻し" Then
                Debug.Print ActiveSheet.Name & s & "行 " & "戻し"
                For n = 4 To 34
                    tikan_cell(s, n) = ""
                Next n
            End If

            '出荷数の置換
'            If tikan_cell(s, 1) = "出荷数" Then
'                Debug.Print ActiveSheet.Name & s &  "出荷数"
'                For n = 4 To 34
'                    tikan_cell(s, n) = Replace(tikan_cell(s, n), dateList.H_AL_nowMonth, dateList.H_AL_nextMonth)
'                Next n
'            End If
            '計画数の置換
'            If tikan_cell(s, 1) = "計画数" Then
'                Debug.Print ActiveSheet.Name & s &  "計画数"
'                For n = 4 To 34
'                    tikan_cell(s, n) = Replace(tikan_cell(s, n), dateList.H_AL_nowMonth, dateList.H_AL_nextMonth)
'                Next n
'            End If
            '福島の置換
'            If tikan_cell(s, 1) = "福島" Then
'                Debug.Print ActiveSheet.Name & s &  "福島"
'                For n = 4 To 34
'                    tikan_cell(s, n) = Replace(tikan_cell(s, n), dateList.H_AL_nowMonth, dateList.H_AL_nextMonth)
'                Next n
'            End If
            'サンプルの置換
'            If tikan_cell(s, 1) = "サンプル" Then
'                Debug.Print ActiveSheet.Name & s &  "サンプル"
'                For n = 4 To 34
'                    tikan_cell(s, n) = Replace(tikan_cell(s, n), dateList.H_AL_nowMonth, dateList.H_AL_nextMonth)
'                Next n
'            End If
'            '仕掛量の置換
'            If tikan_cell(s, 1) = "仕掛量" Then
'                Debug.Print ActiveSheet.Name & s & "行 " & "仕掛量"
'                For n = 4 To 34
'                    tikan_cell(s, n) = Replace(tikan_cell(s, n), dateList.G_now_date, dateList.G_next_date)
'                Next n
'            End If

            '**************下の置換**************
'            If tikan_cell(s, 1) = "比率&必要数" Or tikan_cell(s, 1) = "比率" Then
'                Debug.Print ActiveSheet.Name & s & "行 " &  "比率&必要数"
'                For k = s To LastRow
'                    For n = 4 To 34
'                        tikan_cell(k, n) = Replace(tikan_cell(k, n), dateList.H_AL_nowMonth, dateList.H_AL_nextMonth)
'                    Next n
'                Next k
'                s = LastRow
'            End If
            '**************下の置換**************
            '小分けの置換
'            If tikan_cell(s, 3) = "小分け" Then
'                Debug.Print ActiveSheet.Name & s & "行 " &  "小分け"
'                For n = 4 To 34
'                    tikan_cell(s, n) = Replace(tikan_cell(s, n), dateList.kowake_now_date, dateList.kowake_next_date)
'                Next n
'            End If

        Next s

        '**************H〜ALの置換**************
        Debug.Print ActiveSheet.Name & "H〜ALの置換"
        For k = 1 To LastRow
            For n = 4 To 34
                tikan_cell(k, n) = Replace(tikan_cell(k, n), dateList.H_AL_nowMonth, dateList.H_AL_nextMonth)
                tikan_cell(k, n) = Replace(tikan_cell(k, n), dateList.kowake_now_date, dateList.kowake_next_date)
                tikan_cell(k, n) = Replace(tikan_cell(k, n), dateList.G_now_date, dateList.G_next_date)
            Next n
        Next k

        '**************AO〜BSの置換**************
        If fileList.this_filename = "在庫『包材』_" Then
            Debug.Print ActiveSheet.Name & "AO〜BSの置換"
            For k = 1 To LastRow
                For n = 1 To 31
                    tikan_222(k, n) = Replace(tikan_222(k, n), dateList.H_AL_nextMonth, dateList.H_AL_Month_after_next)
                Next n
            Next k
        End If

        '配列をシートへ ペースト
        On Error Resume Next
        Range(Cells(1, 5), Cells(LastRow, 38)).Formula = tikan_cell

        '包材AO〜BS配列をシートへ ペースト
        If fileList.this_filename = "在庫『包材』_" Then
            On Error Resume Next
            Range(Cells(1, 41), Cells(LastRow, 71)).Formula = tikan_222
        End If

        'コメント削除
        Range("H:AL").ClearComments

        '包材AO〜BSコメント削除
        If fileList.this_filename = "在庫『包材』_" Then
            Range("AO:BS").ClearComments
        End If
End Sub

Sub G列置換(ByVal now_tsuki As String, ByVal next_tsuki As String)
    '最終行取得
    Dim LastRow As Long
    LastRow = Cells(Rows.Count, 7).End(xlUp).Row
    'G列 tikan_cell
    tikan_cell = Range(Cells(1, 7), Cells(LastRow, 7)).Formula
    '置換
    For i = 1 To LastRow
        tikan_cell(i, 1) = Replace(tikan_cell(i, 1), now_tsuki, next_tsuki)
    Next i

    Range(Cells(1, 7), Cells(LastRow, 7)).Formula = tikan_cell
End Sub

Sub 合計金額シートB2翌月へ変更(dateList As Variant, fileList As Variant)
    ActiveWorkbook.Sheets("合計金額").Activate
    Range("B2").Formula = dateList.date_next '合計金額シートB2翌月へ変更
End Sub

Sub 入荷数の色(dateList As Variant, fileList As Variant)
    '最終行取得
    Dim LastRow As Long
    LastRow = Cells(Rows.Count, 5).End(xlUp).Row

    '数式を配列へ格納     E列からAL列 「入荷数」検索用
    tikan_cell = Range(Cells(1, 5), Cells(LastRow, 5)).Formula

    'セルのselect＆コピー
    For s = 1 To LastRow
        If tikan_cell(s, 1) = "入荷数" Then
            Cells(s, 39).Select 'selectセルのリセット
            Cells(s, 39).Copy   '入荷数の合計セル(AM4)コピー (理由はReFAX等色変更しないセルだから）
            Exit For
        End If
    Next s

    '入荷数select
    For s = 1 To LastRow
        '入荷数の削除
        If tikan_cell(s, 1) = "入荷数" Then
            Debug.Print ActiveSheet.Name & "入荷数_色をデフォルトへ"
            For n = 4 To 34
                Union(Selection, Range(Cells(s, 8), Cells(s, 39))).Select
            Next n
        End If
    Next s

    'ペースト
    'ActiveSheet.Paste
        Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
End Sub

Sub 服部コーヒーの色(dateList As Variant, fileList As Variant)
    '最終行取得
    Dim LastRow As Long
    LastRow = Cells(Rows.Count, 5).End(xlUp).Row

    '数式を配列へ格納     E列からAL列 「服部コーヒー」検索用
    tikan_cell = Range(Cells(1, 5), Cells(LastRow, 5)).Formula

    'セルのselect＆コピー
    For s = 1 To LastRow
        If tikan_cell(s, 1) = "服部コーヒー" Then
            Cells(s, 39).Select 'selectセルのリセット
            Cells(s, 39).Copy   '服部コーヒーの合計セル(AM4)コピー (理由はReFAX等色変更しないセルだから）
            Exit For
        End If
    Next s

    '服部コーヒーselect
    For s = 1 To LastRow
        '服部コーヒーの削除
        If tikan_cell(s, 1) = "服部コーヒー" Then
            Debug.Print ActiveSheet.Name & "服部コーヒー_色をデフォルトへ"
            For n = 4 To 34
                Union(Selection, Range(Cells(s, 8), Cells(s, 39))).Select
            Next n
        End If
    Next s

    'ペースト
    'ActiveSheet.Paste
        Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
End Sub

Sub サポートの色(dateList As Variant, fileList As Variant)
    '最終行取得
    Dim LastRow As Long
    LastRow = Cells(Rows.Count, 5).End(xlUp).Row

    '数式を配列へ格納     E列からAL列 「入荷数」検索用
    tikan_cell = Range(Cells(1, 5), Cells(LastRow, 5)).Formula

    'セルのselect＆コピー
    For s = 1 To LastRow
        If tikan_cell(s, 1) = "サポート" Then
            Cells(s, 39).Select 'selectセルのリセット
            Cells(s, 39).Copy   '入荷数の合計セル(AM4)コピー (理由はReFAX等色変更しないセルだから）
            Exit For
        End If
    Next s

    '入荷数select
    For s = 1 To LastRow
        '入荷数の削除
        If tikan_cell(s, 1) = "サポート" Then
            Debug.Print ActiveSheet.Name & "入荷数_色をデフォルトへ"
            For n = 4 To 34
                Union(Selection, Range(Cells(s, 8), Cells(s, 39))).Select
            Next n
        End If
    Next s

    'ペースト
    'ActiveSheet.Paste
        Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
End Sub

Sub ヨネヤマの色(dateList As Variant, fileList As Variant)
    '最終行取得
    Dim LastRow As Long
    LastRow = Cells(Rows.Count, 5).End(xlUp).Row

    '数式を配列へ格納     E列からAL列 「入荷数」検索用
    tikan_cell = Range(Cells(1, 5), Cells(LastRow, 5)).Formula

    'セルのselect＆コピー
    For s = 1 To LastRow
        If tikan_cell(s, 1) = "ヨネヤマ" Then
            Cells(s - 1, 39).Select '(s -1)はサポートの色 セル[AM5]
            Cells(s - 1, 39).Copy '(AM5)コピー (理由はReFAX等色変更しないセルだから）
            Exit For
        End If
    Next s

    '入荷数select
    For s = 1 To LastRow
        '入荷数の削除
        If tikan_cell(s, 1) = "ヨネヤマ" Then
            Debug.Print ActiveSheet.Name & "入荷数_色をデフォルトへ"
            For n = 4 To 33
                Union(Selection, Range(Cells(s, 8), Cells(s, 38))).Select
            Next n
        End If
    Next s

    'ペースト
    'ActiveSheet.Paste
        Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
End Sub


Sub 預けの色(dateList As Variant, fileList As Variant)
    '最終行取得
    Dim LastRow As Long
    LastRow = Cells(Rows.Count, 5).End(xlUp).Row

    '数式を配列へ格納     E列からAL列 「預け」検索用
    tikan_cell = Range(Cells(1, 5), Cells(LastRow, 5)).Formula

    'セルのselect＆コピー
    For s = 1 To LastRow
        If tikan_cell(s, 1) = "預け" Then
            Cells(s, 39).Select 'selectセルのリセット
            Cells(s, 39).Copy   '戻しセル(AM12付近)コピー (理由は色変更しないセルだから）
            Exit For
        End If
    Next s

    '入荷数select
    For s = 1 To LastRow
        '入荷数の削除
        If tikan_cell(s, 1) = "預け" Then
            Debug.Print ActiveSheet.Name & "預け_色をデフォルトへ"
            For n = 4 To 34
                Union(Selection, Range(Cells(s, 8), Cells(s, 39))).Select
            Next n
        End If
    Next s

    'ペースト
    'ActiveSheet.Paste
        Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
End Sub

Sub 戻しの色(dateList As Variant, fileList As Variant)
    '最終行取得
    Dim LastRow As Long
    LastRow = Cells(Rows.Count, 5).End(xlUp).Row

    '数式を配列へ格納     E列からAL列 「戻し」検索用
    tikan_cell = Range(Cells(1, 5), Cells(LastRow, 5)).Formula

    'セルのselect＆コピー
    For s = 1 To LastRow
        If tikan_cell(s, 1) = "戻し" Then
            Cells(s, 39).Select 'selectセルのリセット
            Cells(s, 39).Copy   '戻しセル(AM12付近)コピー (理由は色変更しないセルだから）
            Exit For
        End If
    Next s

    '入荷数select
    For s = 1 To LastRow
        '入荷数の削除
        If tikan_cell(s, 1) = "戻し" Then
            Debug.Print ActiveSheet.Name & "戻し_色をデフォルトへ"
            For n = 4 To 34
                Union(Selection, Range(Cells(s, 8), Cells(s, 39))).Select
            Next n
        End If
    Next s

    'ペースト
    'ActiveSheet.Paste
        Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
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

Public Function 日付クラス() As dateClass
    '宣言
    Dim dateList  As dateClass
    Set dateList = New dateClass

    'セット
    ActiveWorkbook.Sheets("合計金額").Activate
    dateList.date_now = Range("B2")    '合計金額シートのB2日付
    dateList.date_after_next = DateAdd("m", 2, DateSerial(Year(dateList.date_now), Month(dateList.date_now), 1)) '翌々月date
    dateList.date_next = DateAdd("m", 1, DateSerial(Year(dateList.date_now), Month(dateList.date_now), 1))       '翌月date
    dateList.date_last = DateAdd("m", -1, DateSerial(Year(dateList.date_now), Month(dateList.date_now), 1))       '先月date
    dateList.H_AL_nowMonth = Format(dateList.date_now, "m月")    '今月「x月」
    dateList.H_AL_nextMonth = Format(dateList.date_next, "m月") '翌月「x月」
    dateList.H_AL_Month_after_next = Format(dateList.date_after_next, "m月") '翌々月「x月」
    dateList.G_now_date = Year(dateList.date_now) & "." & Month(dateList.date_now) '例「2020.7」
    dateList.G_last_date = Year(dateList.date_last) & "." & Month(dateList.date_last) '例「2020.6」
    dateList.G_next_date = Year(dateList.date_next) & "." & Month(dateList.date_next) '例「2020.8」
    dateList.kowake_now_date = Year(dateList.date_now) & "." & Right("0" & Month(dateList.date_now), 2) '例「2020.07」
    dateList.kowake_last_date = Year(dateList.date_last) & "." & Right("0" & Month(dateList.date_last), 2) '例「2020.06」
    dateList.kowake_next_date = Year(dateList.date_next) & "." & Right("0" & Month(dateList.date_next), 2) '例「2020.08」

    Set 日付クラス = dateList
End Function

Public Function ファイルクラス() As fileClass
    '宣言
    Dim fileList  As fileClass
    Set fileList = New fileClass
    'セット
    fileList.mypath = ThisWorkbook.Path
    fileList.mybook = ThisWorkbook.Name
    fileList.mmfn = fileList.mypath & "\" & fileList.mybook
    'fileList.checkFilename = "在庫『副原材料』_"

    'ファイル名の月を抽出
    ActiveWorkbook.Sheets("合計金額").Activate
    date_now = Range("B2")    '合計金額シートのB2日付
    s = InStr(fileList.mybook, Year(date_now) & ".") + 5  '「2020.」は何文字目から始まるか
    l = Len(fileList.mybook)
    str2 = Mid(fileList.mybook, s, l)   '例 「8.xlsm」
    s2 = InStr(str2, ".")
    l2 = Len(str2)

    Fname = InStr(fileList.mybook, Year(date_now) & ".") '「2020.」は何文字目から始まるか
    fileList.this_filename = Mid(fileList.mybook, 1, s - 6) '例   在庫『副原材料』_

'    If fileList.this_filename = fileList.checkFilename Then
'        fileList.bool_filename = True
'    Else
'        fileList.bool_filename = False
'    End If

    fileList.mybook_month = Int(Mid(str2, 1, s2 - 1)) '8
    '例[   \\Afnewt320-kyoyu\社内共有\個人フォルダ\笠間\在庫表作成マクロ\実験\在庫『ポリ』_2020.9.xlsm   ]
    fileList.next_mybook = fileList.mypath & "\" & fileList.this_filename & Year(date_now) & "." & (Month(date_now) + 1) & ".xlsm"

    Set ファイルクラス = fileList
End Function

